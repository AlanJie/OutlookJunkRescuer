using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkRescuer
{
    /// <summary>
    /// Narrow capability boundary for user-owned Junk mail.
    /// Allowed: read identifiers, enumerate, Copy().
    /// Forbidden by design: Save, Move, Delete, property writes, flag/category
    /// changes, or any other mutation of the source item.
    /// </summary>
    internal sealed class OutlookSourceReader
    {
        private readonly Outlook.NameSpace _session;

        public OutlookSourceReader(Outlook.NameSpace session)
        {
            _session = session;
        }

        public List<SourceMessageDescriptor> ReadJunk(
            string accountSmtp,
            string storeId,
            Outlook.MAPIFolder junkFolder)
        {
            try
            {
                return ReadJunkViaTable(accountSmtp, storeId, junkFolder);
            }
            catch (Exception ex)
            {
                Logger.Write($"[{accountSmtp}] Table-based ReadJunk failed, falling back to Items enumeration: {ex.Message}");
                return ReadJunkViaItems(accountSmtp, storeId, junkFolder);
            }
        }

        private List<SourceMessageDescriptor> ReadJunkViaTable(
            string accountSmtp,
            string storeId,
            Outlook.MAPIFolder junkFolder)
        {
            var result = new List<SourceMessageDescriptor>();
            Outlook.Table table = null;
            Outlook.Columns columns = null;

            try
            {
                table = junkFolder.GetTable();
                columns = table.Columns;

                try { columns.RemoveAll(); } catch { }
                try { columns.Add("EntryID"); } catch { }
                try { columns.Add("MessageClass"); } catch { }
                try { columns.Add(MapiIdentity.SearchKeySchema); } catch { }
                try { columns.Add(MapiIdentity.RecordKeySchema); } catch { }

                while (!table.EndOfTable)
                {
                    Outlook.Row row = null;
                    try
                    {
                        row = table.GetNextRow();

                        string messageClass = row["MessageClass"] as string;
                        if (string.IsNullOrEmpty(messageClass) ||
                            !messageClass.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        string entryId = row["EntryID"] as string;
                        object rawSearch = row[MapiIdentity.SearchKeySchema];
                        object rawRecord = row[MapiIdentity.RecordKeySchema];

                        string searchKeyHex = rawSearch is byte[] sBytes ? MapiIdentity.ToHex(sBytes) : null;
                        string recordKeyHex = rawRecord is byte[] rBytes ? MapiIdentity.ToHex(rBytes) : null;

                        if (!string.IsNullOrEmpty(entryId) &&
                            !string.IsNullOrEmpty(searchKeyHex) &&
                            !string.IsNullOrEmpty(recordKeyHex))
                        {
                            result.Add(new SourceMessageDescriptor(
                                accountSmtp,
                                storeId,
                                entryId,
                                searchKeyHex,
                                recordKeyHex));
                        }
                    }
                    finally
                    {
                        ComUtil.Release(row);
                    }
                }

                return result;
            }
            finally
            {
                ComUtil.Release(columns);
                ComUtil.Release(table);
            }
        }

        private List<SourceMessageDescriptor> ReadJunkViaItems(
            string accountSmtp,
            string storeId,
            Outlook.MAPIFolder junkFolder)
        {
            var result = new List<SourceMessageDescriptor>();
            Outlook.Items items = null;

            try
            {
                items = junkFolder.Items;
                int count = items.Count;

                for (int i = 1; i <= count; i++)
                {
                    object raw = null;

                    try
                    {
                        raw = items[i];
                        var descriptor = TryReadDescriptor(accountSmtp, storeId, raw);
                        if (descriptor != null)
                        {
                            result.Add(descriptor);
                        }
                    }
                    catch (COMException ex)
                    {
                        Logger.Write(
                            $"[{accountSmtp}] Could not read Junk item #{i}: " +
                            $"0x{ex.HResult:X8} {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        Logger.Write(
                            $"[{accountSmtp}] Could not read Junk item #{i}: {ex}");
                    }
                    finally
                    {
                        ComUtil.Release(raw);
                    }
                }
            }
            finally
            {
                ComUtil.Release(items);
            }

            return result;
        }

        public Outlook.MailItem CreateCopy(SourceMessageDescriptor source)
        {
            object raw = null;

            try
            {
                raw = _session.GetItemFromID(
                    source.EntryId,
                    source.StoreId);

                var original = raw as Outlook.MailItem;

                if (original == null)
                    throw new InvalidOperationException(
                        "The source EntryID no longer resolves to a MailItem.");

                // EntryID is only a locator. Validate the resolved object using
                // persistent MAPI identity before exercising the Copy capability.
                string searchKey = MapiIdentity.GetSearchKeyHex(original);
                string recordKey = MapiIdentity.GetRecordKeyHex(original);

                if (!string.Equals(
                        searchKey,
                        source.SearchKeyHex,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        recordKey,
                        source.RecordKeyHex,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The source EntryID resolved to a different MAPI record.");
                }

                // Verify the source item is still in the Junk folder.
                // If the user marked it as "Not Junk" or moved it, do not copy.
                Outlook.Store store = null;
                Outlook.MAPIFolder junkFolder = null;
                object parentObj = null;

                try
                {
                    store = _session.GetStoreFromID(source.StoreId);
                    if (store == null)
                        throw new InvalidOperationException($"Unable to resolve store {source.StoreId} to verify folder location.");

                    junkFolder = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);
                    if (junkFolder == null)
                        throw new InvalidOperationException($"Unable to resolve Junk folder for store {source.StoreId}.");

                    parentObj = original.Parent;
                    var parentFolder = parentObj as Outlook.MAPIFolder;
                    if (parentFolder == null)
                        throw new InvalidOperationException("Unable to resolve parent folder of the source message.");

                    string parentEntryId = parentFolder.EntryID;
                    string junkEntryId = junkFolder.EntryID;

                    if (string.IsNullOrEmpty(parentEntryId) || string.IsNullOrEmpty(junkEntryId))
                        throw new InvalidOperationException("Unable to read EntryID for parent or Junk folder.");

                    // Positive verification: return null only when parent is provably not Junk
                    if (!_session.CompareEntryIDs(parentEntryId, junkEntryId))
                    {
                        // Source item has provably moved out of Junk
                        return null;
                    }
                }
                finally
                {
                    ComUtil.Release(parentObj);
                    ComUtil.Release(junkFolder);
                    ComUtil.Release(store);
                }

                // The only operation this adapter performs on the source object.
                return (Outlook.MailItem)original.Copy();
            }
            finally
            {
                ComUtil.Release(raw);
            }
        }

        public SourceMessageDescriptor TryReadDescriptor(
            string accountSmtp,
            string storeId,
            object rawItem)
        {
            var mail = rawItem as Outlook.MailItem;
            if (mail == null)
                return null;

            try
            {
                string messageClass = mail.MessageClass;
                if (string.IsNullOrEmpty(messageClass) ||
                    !messageClass.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                string entryId = mail.EntryID;
                string searchKey = MapiIdentity.GetSearchKeyHex(mail);
                string recordKey = MapiIdentity.GetRecordKeyHex(mail);

                if (string.IsNullOrEmpty(entryId) ||
                    string.IsNullOrEmpty(searchKey) ||
                    string.IsNullOrEmpty(recordKey))
                {
                    return null;
                }

                return new SourceMessageDescriptor(
                    accountSmtp,
                    storeId,
                    entryId,
                    searchKey,
                    recordKey);
            }
            catch (Exception ex)
            {
                Logger.Write($"[{accountSmtp}] TryReadDescriptor failed: {ex.Message}");
                return null;
            }
        }
    }
}
