using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkRescuer
{
    public enum DuplicateRetentionPolicy
    {
        /// <summary>
        /// 保留最早创建的归档副本（推荐）
        /// </summary>
        KeepEarliest,

        /// <summary>
        /// 保留最新创建的归档副本
        /// </summary>
        KeepLatest
    }

    public enum DuplicateDestination
    {
        /// <summary>
        /// 移至专用的 Duplicate Trash 软隔离目录（Junk Archive\OutlookJunkRescuer Duplicate Trash）
        /// </summary>
        DuplicateTrash,

        /// <summary>
        /// 移至 Outlook 系统自带的「已删除邮件」废件箱 (Deleted Items)
        /// </summary>
        DeletedItems
    }

    public sealed class OwnedCopyInfo
    {
        public string EntryId { get; set; }
        public string StoreId { get; set; }
        public string ActualSearchKey { get; set; }
        public string ArchiveKey { get; set; }
        public string CopyId { get; set; }
        public string ReplicaId { get; set; }
        public string PluginId { get; set; }
        public DateTime CreatedUtc { get; set; }
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
        private const string PluginIdDasl = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/" + ArchiveWriter.PluginIdProperty;
        private const string ArchiveKeyDasl = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/" + ArchiveWriter.ArchiveKeyProperty;
        private const string CopyIdDasl = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/" + ArchiveWriter.CopyIdProperty;
        private const string ReplicaIdDasl = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/" + ArchiveWriter.ReplicaIdProperty;
        private const string CreatedUtcDasl = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/" + ArchiveWriter.CreatedUtcProperty;

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
        /// 严格验证所有权标记（OJRPluginId, OJRArchiveKey, OJRCopyId, OJRReplicaId, PR_SEARCH_KEY）。
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

                bool hasPluginCol = false;
                bool hasArchiveKeyCol = false;
                bool hasCopyCol = false;
                bool hasReplicaCol = false;
                bool hasCreatedUtcCol = false;

                try { columns.Add(PluginIdDasl); hasPluginCol = true; } catch { }
                try { columns.Add(ArchiveKeyDasl); hasArchiveKeyCol = true; } catch { }
                try { columns.Add(CopyIdDasl); hasCopyCol = true; } catch { }
                try { columns.Add(ReplicaIdDasl); hasReplicaCol = true; } catch { }
                try { columns.Add(CreatedUtcDasl); hasCreatedUtcCol = true; } catch { }

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

                        string actualSearchKeyHex = MapiIdentity.ToHex(searchKeyBytes);
                        if (string.IsNullOrEmpty(actualSearchKeyHex))
                            continue;

                        string pluginId = null;
                        string archiveKey = null;
                        string copyId = null;
                        string replicaId = null;
                        string createdUtcStr = null;

                        if (hasPluginCol && hasArchiveKeyCol && hasCopyCol && hasReplicaCol)
                        {
                            try { pluginId = Convert.ToString(row[PluginIdDasl]); } catch { }
                            try { archiveKey = Convert.ToString(row[ArchiveKeyDasl]); } catch { }
                            try { copyId = Convert.ToString(row[CopyIdDasl]); } catch { }
                            try { replicaId = Convert.ToString(row[ReplicaIdDasl]); } catch { }
                            if (hasCreatedUtcCol)
                            {
                                try { createdUtcStr = Convert.ToString(row[CreatedUtcDasl]); } catch { }
                            }
                        }
                        else
                        {
                            // Table custom columns not supported by the store; fallback to reading item properties
                            Outlook.MailItem mail = null;
                            try
                            {
                                mail = _session.GetItemFromID(entryId, storeId) as Outlook.MailItem;
                                if (mail != null)
                                {
                                    pluginId = GetTextProperty(mail, ArchiveWriter.PluginIdProperty);
                                    archiveKey = GetTextProperty(mail, ArchiveWriter.ArchiveKeyProperty);
                                    copyId = GetTextProperty(mail, ArchiveWriter.CopyIdProperty);
                                    replicaId = GetTextProperty(mail, ArchiveWriter.ReplicaIdProperty);
                                    createdUtcStr = GetTextProperty(mail, ArchiveWriter.CreatedUtcProperty);
                                }
                            }
                            catch
                            {
                                // Ignore
                            }
                            finally
                            {
                                ComUtil.Release(mail);
                            }
                        }

                        // Strict OJR-owned validation: all markers must be present and valid
                        if (!string.Equals(pluginId, ArchiveWriter.PluginIdValue, StringComparison.Ordinal))
                            continue;

                        if (string.IsNullOrEmpty(archiveKey) ||
                            string.IsNullOrEmpty(copyId) ||
                            string.IsNullOrEmpty(replicaId))
                            continue;

                        // Actual search key must match stamped archive key
                        if (!string.Equals(actualSearchKeyHex, archiveKey, StringComparison.OrdinalIgnoreCase))
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

                        DateTime createdUtc = receivedTime;
                        if (!string.IsNullOrEmpty(createdUtcStr) &&
                            DateTime.TryParse(createdUtcStr, null, System.Globalization.DateTimeStyles.RoundtripKind, out DateTime dtParsed))
                        {
                            createdUtc = dtParsed.ToUniversalTime();
                        }

                        allCopies.Add(new OwnedCopyInfo
                        {
                            EntryId = entryId,
                            StoreId = storeId,
                            ActualSearchKey = actualSearchKeyHex,
                            ArchiveKey = archiveKey,
                            CopyId = copyId,
                            ReplicaId = replicaId,
                            PluginId = pluginId,
                            CreatedUtc = createdUtc,
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
                .GroupBy(c => c.ArchiveKey, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .Select(g => new DuplicateGroup(g.Key, g.ToList()))
                .ToList();

            return duplicateGroups;
        }

        /// <summary>
        /// 保留重验证清理：对指定的重复组执行清理，将多余的副本安全移动至指定目标文件夹（Duplicate Trash 软隔离目录或废件箱）。
        /// 遵循绝对铁律：Never reduce 1 -> 0（绝不导致该邮件在归档中清零）。
        /// 每次移动前重新对具体项目执行完整的 OJR 所有权正向验证。
        /// </summary>
        public CleanupResult CleanDuplicates(
            Outlook.MAPIFolder archiveFolder,
            string storeId,
            List<DuplicateGroup> groups,
            DuplicateRetentionPolicy policy,
            DuplicateDestination destination,
            Action<int, int> progressCallback)
        {
            var result = new CleanupResult
            {
                TotalGroups = groups != null ? groups.Count : 0
            };

            if (archiveFolder == null || groups == null || groups.Count == 0)
                return result;

            Outlook.MAPIFolder targetFolder = null;
            Outlook.Store store = null;
            try
            {
                if (destination == DuplicateDestination.DeletedItems)
                {
                    if (string.IsNullOrEmpty(storeId))
                        throw new InvalidOperationException("未提供 StoreId，无法解析当前账户的「已删除邮件」废件箱。");

                    try
                    {
                        store = _session.GetStoreFromID(storeId);
                        targetFolder = store?.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
                    }
                    catch (Exception ex)
                    {
                        Logger.Write($"[DuplicateCleaner] 获取 StoreId={storeId} 的 Deleted Items 失败: {ex}");
                        targetFolder = null;
                    }

                    if (targetFolder == null)
                        throw new InvalidOperationException("无法获取当前账户的「已删除邮件」废件箱。为防止跨账户误投递，操作已中止。");
                }
                else
                {
                    targetFolder = _archiveWriter.GetOrCreateDuplicateTrashFolder(archiveFolder);
                }

                if (targetFolder == null)
                    throw new InvalidOperationException("无法解析或获取目标隔离文件夹。");

                int currentGroupIndex = 0;

                foreach (var group in groups)
                {
                    currentGroupIndex++;
                    progressCallback?.Invoke(currentGroupIndex, groups.Count);

                    if (group.Copies.Count <= 1)
                        continue;

                    // 模糊性检查：如果同一 OJRCopyId 出现多次，判定为状态模糊（可能是用户手工复制或其他干扰），整组跳过自动清理
                    bool hasDuplicateCopyId = group.Copies
                        .GroupBy(c => c.CopyId, StringComparer.OrdinalIgnoreCase)
                        .Any(g => g.Count() > 1);

                    if (hasDuplicateCopyId)
                    {
                        Logger.Write($"[DuplicateCleaner] 组 {group.ArchiveKey} 存在重复的 OJRCopyId，状态模糊 (Ambiguous)，整组跳过清理。");
                        result.Skipped += group.Copies.Count;
                        continue;
                    }

                    // 根据策略排序副本：按 OJRCreatedUtc 排序，并以 OJRCopyId 作为稳定 tie-break
                    List<OwnedCopyInfo> orderedCopies;
                    if (policy == DuplicateRetentionPolicy.KeepLatest)
                    {
                        orderedCopies = group.Copies
                            .OrderByDescending(c => c.CreatedUtc)
                            .ThenByDescending(c => c.CopyId, StringComparer.Ordinal)
                            .ToList();
                    }
                    else
                    {
                        // 默认：保留最早创建的副本
                        orderedCopies = group.Copies
                            .OrderBy(c => c.CreatedUtc)
                            .ThenBy(c => c.CopyId, StringComparer.Ordinal)
                            .ToList();
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
                            if (ValidateOwnedCopy(item, archiveFolder, candidate, group.ArchiveKey))
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

                    // 如果整组连 1 份能够被正面证明为 OJR 所有且存在于 Junk Archive 的有效副本都没有，绝不移动任何邮件
                    if (winnerCopy == null || winnerItem == null)
                    {
                        Logger.Write($"[DuplicateCleaner] 组 {group.ArchiveKey} 未找到通过完整所有权验证的 Junk Archive 留存副本，跳过以防误删。");
                        result.Skipped += group.Copies.Count;
                        continue;
                    }

                    try
                    {
                        // 2. 遍历其余候选 Loser 副本进行保守移动
                        foreach (var loserCopy in orderedCopies)
                        {
                            if (ReferenceEquals(loserCopy, winnerCopy))
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

                                // 严格验证 loser 的 OJR 所有权与标记
                                if (!ValidateOwnedCopy(loserItem, archiveFolder, loserCopy, group.ArchiveKey))
                                {
                                    Logger.Write($"[DuplicateCleaner] 副本 {loserCopy.EntryId} 未能通过完整所有权验证，跳过。");
                                    result.Skipped++;
                                    continue;
                                }

                                // 关键铁律再次核实：Winner 仍然有效且留在 Junk Archive 中
                                if (!ValidateOwnedCopy(winnerItem, archiveFolder, winnerCopy, group.ArchiveKey))
                                {
                                    Logger.Write($"[DuplicateCleaner] Winner 副本状态改变，立即终止对组 {group.ArchiveKey} 的后续移动以坚守 Never-reduce-1->0。");
                                    result.Skipped++;
                                    break;
                                }

                                // 安全移动到目标隔离目录（Duplicate Trash 或废件箱）
                                Outlook.MailItem moved = null;
                                try
                                {
                                    moved = loserItem.Move(targetFolder) as Outlook.MailItem;
                                    result.MovedToTrash++;
                                }
                                finally
                                {
                                    ComUtil.Release(moved);
                                }
                            }
                            catch (Exception ex)
                            {
                                result.Failed++;
                                Logger.Write($"[DuplicateCleaner] 移动副本 {loserCopy.EntryId} 至目标目录失败: {ex}");
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
                ComUtil.Release(targetFolder);
                ComUtil.Release(store);
            }

            return result;
        }

        /// <summary>
        /// 安全清空 Duplicate Trash 目录中的冗余副本（将其移入系统「已删除邮件」）。
        /// 核心安全原则：逐项执行 OJR 所有权正向验证，任何未打上完整 OJR 标记的项目绝不触碰并予以保留。
        /// </summary>
        public void EmptyTrash(
            Outlook.MAPIFolder trashFolder,
            out int deletedCount,
            out int skippedUnknownCount)
        {
            deletedCount = 0;
            skippedUnknownCount = 0;

            if (trashFolder == null)
                return;

            Outlook.Items items = null;
            try
            {
                items = trashFolder.Items;
                int count = items.Count;
                for (int i = count; i >= 1; i--)
                {
                    object itemObj = null;
                    Outlook.MailItem mail = null;
                    try
                    {
                        itemObj = items[i];
                        mail = itemObj as Outlook.MailItem;
                        if (mail == null)
                        {
                            skippedUnknownCount++;
                            continue;
                        }

                        // Verify OJR ownership
                        string pluginId = GetTextProperty(mail, ArchiveWriter.PluginIdProperty);
                        string archiveKey = GetTextProperty(mail, ArchiveWriter.ArchiveKeyProperty);
                        string copyId = GetTextProperty(mail, ArchiveWriter.CopyIdProperty);

                        if (!string.Equals(pluginId, ArchiveWriter.PluginIdValue, StringComparison.Ordinal) ||
                            string.IsNullOrEmpty(archiveKey) ||
                            string.IsNullOrEmpty(copyId))
                        {
                            Logger.Write($"[DuplicateCleaner] 隔离目录中发现非 OJR 项目或标记不完整 (Subject={mail.Subject})，保留不予清理。");
                            skippedUnknownCount++;
                            continue;
                        }

                        // Delete() moves mail from regular folder to Deleted Items
                        mail.Delete();
                        deletedCount++;
                    }
                    catch (Exception ex)
                    {
                        Logger.Write($"[DuplicateCleaner] 清理隔离副本失败: {ex}");
                        skippedUnknownCount++;
                    }
                    finally
                    {
                        ComUtil.Release(mail);
                        ComUtil.Release(itemObj);
                    }
                }
            }
            finally
            {
                ComUtil.Release(items);
            }
        }

        private bool ValidateOwnedCopy(
            Outlook.MailItem item,
            Outlook.MAPIFolder archiveFolder,
            OwnedCopyInfo copyInfo,
            string expectedArchiveKey)
        {
            if (item == null || archiveFolder == null || copyInfo == null)
                return false;

            // 1. item still in Archive folder
            if (!_ownedCopies.IsInFolder(item, archiveFolder))
                return false;

            // 2. OJRPluginId matches
            string pluginId = GetTextProperty(item, ArchiveWriter.PluginIdProperty);
            if (!string.Equals(pluginId, ArchiveWriter.PluginIdValue, StringComparison.Ordinal))
                return false;

            // 3. OJRArchiveKey == expected ArchiveKey
            string archiveKey = GetTextProperty(item, ArchiveWriter.ArchiveKeyProperty);
            if (!string.Equals(archiveKey, expectedArchiveKey, StringComparison.OrdinalIgnoreCase))
                return false;

            // 4. OJRCopyId == scanned CopyId
            string copyId = GetTextProperty(item, ArchiveWriter.CopyIdProperty);
            if (!string.Equals(copyId, copyInfo.CopyId, StringComparison.OrdinalIgnoreCase))
                return false;

            // 5. PR_SEARCH_KEY == expected ArchiveKey
            string actualSearchKey = MapiIdentity.GetSearchKeyHex(item);
            if (!string.Equals(actualSearchKey, expectedArchiveKey, StringComparison.OrdinalIgnoreCase))
                return false;

            return true;
        }

        private static string GetTextProperty(Outlook.MailItem mail, string name)
        {
            if (mail == null || string.IsNullOrEmpty(name))
                return null;

            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = mail.UserProperties;
                property = properties.Find(name, true);
                return property?.Value == null ? null : Convert.ToString(property.Value);
            }
            catch
            {
                return null;
            }
            finally
            {
                ComUtil.Release(property);
                ComUtil.Release(properties);
            }
        }
    }
}
