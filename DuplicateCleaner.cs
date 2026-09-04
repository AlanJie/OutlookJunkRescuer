using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkRescuer
{
    public enum DuplicateRetentionPolicy
    {
        /// <summary>
        /// 保留最早接收或创建的归档副本（推荐）
        /// </summary>
        KeepEarliest,

        /// <summary>
        /// 保留最新接收或创建的归档副本
        /// </summary>
        KeepLatest
    }

    public sealed class OwnedCopyInfo
    {
        public string EntryId { get; set; }
        public string StoreId { get; set; }
        public string ArchiveKey { get; set; }
        public string CopyId { get; set; }
        public string ReplicaId { get; set; }
        public string Subject { get; set; }
        public DateTime ReceivedTime { get; set; }
    }

    public sealed class DuplicateGroup
    {
        public string ArchiveKey { get; }
        public List<OwnedCopyInfo> Copies { get; }

        public DuplicateGroup(string archiveKey, List<OwnedCopyInfo> copies)
        {
            ArchiveKey = archiveKey;
            Copies = copies ?? new List<OwnedCopyInfo>();
        }

        public string Subject => Copies.Count > 0 ? Copies[0].Subject : "(无主题)";
        public int TotalCopies => Copies.Count;
        public int RedundantCopies => Math.Max(0, Copies.Count - 1);
    }

    public sealed class CleanupResult
    {
        public int TotalGroups { get; set; }
        public int TotalCopiesScanned { get; set; }
        public int MovedToTrash { get; set; }
        public int Skipped { get; set; }
        public int Failed { get; set; }
    }

    internal sealed class DuplicateCleaner
    {
        private const string PR_SEARCH_KEY = "http://schemas.microsoft.com/mapi/proptag/0x300B0102";
        private const string ReplicaIdDasl = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/OJRReplicaId";
        private const string CopyIdDasl = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/OJRCopyId";

        private readonly Outlook.NameSpace _session;
        private readonly ArchiveWriter _archiveWriter;
        private readonly OwnedCopyLocator _ownedCopies;

        public DuplicateCleaner(Outlook.NameSpace session)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
            _archiveWriter = new ArchiveWriter();
            _ownedCopies = new OwnedCopyLocator(session);
        }

        /// <summary>
        /// 扫描 Junk Archive 文件夹中的所有副本，按逻辑邮件标识（SearchKey）分组，筛选出存在多个副本的重复组。
        /// 扫描基于 Outlook Table 进行只读轻量遍历，不修改任何邮件。
        /// </summary>
        public List<DuplicateGroup> ScanDuplicates(
            Outlook.MAPIFolder archiveFolder,
            string storeId,
            out int totalScannedCount)
        {
            totalScannedCount = 0;
            if (archiveFolder == null)
                return new List<DuplicateGroup>();

            Outlook.Table table = null;
            Outlook.Columns columns = null;

            var allCopies = new List<OwnedCopyInfo>();

            try
            {
                table = archiveFolder.GetTable(Type.Missing, Outlook.OlTableContents.olUserItems);
                columns = table.Columns;
                columns.RemoveAll();

                columns.Add("EntryID");
                columns.Add("Subject");
                columns.Add("ReceivedTime");
                columns.Add("MessageClass");
                columns.Add(PR_SEARCH_KEY);

                bool hasReplicaCol = false;
                bool hasCopyCol = false;

                try
                {
                    columns.Add(ReplicaIdDasl);
                    hasReplicaCol = true;
                }
                catch
                {
                    // DASL custom column may fail on some store configurations; handled gracefully
                }

                try
                {
                    columns.Add(CopyIdDasl);
                    hasCopyCol = true;
                }
                catch
                {
                    // Handled gracefully
                }

                while (!table.EndOfTable)
                {
                    Outlook.Row row = null;

                    try
                    {
                        row = table.GetNextRow();
                        if (row == null)
                            break;

                        totalScannedCount++;

                        string msgClass = Convert.ToString(row["MessageClass"]);
                        if (string.IsNullOrEmpty(msgClass) ||
                            !msgClass.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        string entryId = Convert.ToString(row["EntryID"]);
                        if (string.IsNullOrEmpty(entryId))
                            continue;

                        byte[] searchKeyBytes = row[PR_SEARCH_KEY] as byte[];
                        if (searchKeyBytes == null || searchKeyBytes.Length == 0)
                            continue;

                        string searchKeyHex = MapiIdentity.ToHex(searchKeyBytes);
                        if (string.IsNullOrEmpty(searchKeyHex))
                            continue;

                        string subject = Convert.ToString(row["Subject"]);
                        DateTime receivedTime = DateTime.MinValue;
                        try
                        {
                            object rtVal = row["ReceivedTime"];
                            if (rtVal is DateTime dt)
                                receivedTime = dt;
                        }
                        catch
                        {
                            // Ignore date parse errors
                        }

                        string replicaId = null;
                        if (hasReplicaCol)
                        {
                            try
                            {
                                replicaId = Convert.ToString(row[ReplicaIdDasl]);
                            }
                            catch
                            {
                                replicaId = null;
                            }
                        }

                        string copyId = null;
                        if (hasCopyCol)
                        {
                            try
                            {
                                copyId = Convert.ToString(row[CopyIdDasl]);
                            }
                            catch
                            {
                                copyId = null;
                            }
                        }

                        allCopies.Add(new OwnedCopyInfo
                        {
                            EntryId = entryId,
                            StoreId = storeId,
                            ArchiveKey = searchKeyHex,
                            CopyId = copyId,
                            ReplicaId = replicaId,
                            Subject = subject,
                            ReceivedTime = receivedTime
                        });
                    }
                    finally
                    {
                        ComUtil.Release(row);
                    }
                }
            }
            finally
            {
                ComUtil.Release(columns);
                ComUtil.Release(table);
            }

            // 按 ArchiveKey 分组并筛选重复数 > 1 的组
            var duplicateGroups = allCopies
                .GroupBy(c => c.ArchiveKey, StringComparer.Ordinal)
                .Where(g => g.Count() > 1)
                .Select(g => new DuplicateGroup(g.Key, g.ToList()))
                .ToList();

            return duplicateGroups;
        }

        /// <summary>
        /// 保守重验证清理：对指定的重复组执行清理，将多余的副本安全移动至 Duplicate Trash 软隔离目录。
        /// 遵循绝对铁律：Never reduce 1 -> 0（绝不导致该邮件在归档中清零）。
        /// </summary>
        public CleanupResult CleanDuplicates(
            Outlook.MAPIFolder archiveFolder,
            List<DuplicateGroup> groups,
            DuplicateRetentionPolicy policy,
            Action<int, int> progressCallback)
        {
            var result = new CleanupResult
            {
                TotalGroups = groups != null ? groups.Count : 0
            };

            if (archiveFolder == null || groups == null || groups.Count == 0)
                return result;

            Outlook.MAPIFolder duplicateTrash = null;
            try
            {
                duplicateTrash = _archiveWriter.GetOrCreateDuplicateTrashFolder(archiveFolder);
                if (duplicateTrash == null)
                    throw new InvalidOperationException("无法解析或创建 Duplicate Trash 目录。");

                int currentGroupIndex = 0;

                foreach (var group in groups)
                {
                    currentGroupIndex++;
                    progressCallback?.Invoke(currentGroupIndex, groups.Count);

                    if (group.Copies.Count <= 1)
                        continue;

                    // 根据策略排序副本
                    List<OwnedCopyInfo> orderedCopies;
                    if (policy == DuplicateRetentionPolicy.KeepLatest)
                    {
                        orderedCopies = group.Copies.OrderByDescending(c => c.ReceivedTime).ToList();
                    }
                    else
                    {
                        // 默认：保留最早的
                        orderedCopies = group.Copies.OrderBy(c => c.ReceivedTime).ToList();
                    }

                    // 1. 验证并寻找 Winner（留在 Junk Archive 中的法定副本）
                    OwnedCopyInfo winnerCopy = null;
                    Outlook.MailItem winnerItem = null;

                    for (int i = 0; i < orderedCopies.Count; i++)
                    {
                        var candidate = orderedCopies[i];
                        Outlook.MailItem item = null;
                        try
                        {
                            item = _session.GetItemFromID(candidate.EntryId, candidate.StoreId) as Outlook.MailItem;
                            if (item != null && _ownedCopies.IsInFolder(item, archiveFolder))
                            {
                                winnerCopy = candidate;
                                winnerItem = item;
                                item = null;
                                break;
                            }
                        }
                        catch
                        {
                            // Winner candidate not accessible
                        }
                        finally
                        {
                            ComUtil.Release(item);
                        }
                    }

                    // 如果整组连 1 份能够被证实存在于 Junk Archive 的有效副本都没有，绝对不动任何邮件
                    if (winnerCopy == null || winnerItem == null)
                    {
                        Logger.Write($"[DuplicateCleaner] 组 {group.ArchiveKey} 未找到有效的 Junk Archive 留存副本，跳过以防误删。");
                        result.Skipped += group.Copies.Count;
                        continue;
                    }

                    try
                    {
                        // 2. 遍历其余候选 Loser 副本进行保守移动
                        foreach (var loserCopy in orderedCopies)
                        {
                            if (string.Equals(loserCopy.EntryId, winnerCopy.EntryId, StringComparison.Ordinal))
                                continue;

                            Outlook.MailItem loserItem = null;
                            try
                            {
                                loserItem = _session.GetItemFromID(loserCopy.EntryId, loserCopy.StoreId) as Outlook.MailItem;
                                if (loserItem == null)
                                {
                                    result.Skipped++;
                                    continue;
                                }

                                // 验证 loser 是否仍在 Junk Archive 文件夹中
                                if (!_ownedCopies.IsInFolder(loserItem, archiveFolder))
                                {
                                    result.Skipped++;
                                    continue;
                                }

                                // 验证 MAPI 逻辑标识是否匹配
                                string actualSearchKey = MapiIdentity.GetSearchKeyHex(loserItem);
                                if (!string.Equals(actualSearchKey, group.ArchiveKey, StringComparison.Ordinal))
                                {
                                    Logger.Write($"[DuplicateCleaner] 副本 {loserCopy.EntryId} 的 SearchKey 不匹配，跳过。");
                                    result.Skipped++;
                                    continue;
                                }

                                // 关键铁律再次核实：Winner 仍然在 Junk Archive 中
                                if (!_ownedCopies.IsInFolder(winnerItem, archiveFolder))
                                {
                                    Logger.Write($"[DuplicateCleaner] Winner 副本意外不再存在于 Junk Archive，立即终止对组 {group.ArchiveKey} 的后续移动以坚守 Never-reduce-1->0。");
                                    result.Skipped++;
                                    break;
                                }

                                // 安全移动到 Duplicate Trash 软隔离目录
                                loserItem.Move(duplicateTrash);
                                result.MovedToTrash++;
                            }
                            catch (Exception ex)
                            {
                                result.Failed++;
                                Logger.Write($"[DuplicateCleaner] 移动副本 {loserCopy.EntryId} 至 Duplicate Trash 失败: {ex}");
                            }
                            finally
                            {
                                ComUtil.Release(loserItem);
                            }
                        }
                    }
                    finally
                    {
                        ComUtil.Release(winnerItem);
                    }
                }
            }
            finally
            {
                ComUtil.Release(duplicateTrash);
            }

            return result;
        }
    }
}
