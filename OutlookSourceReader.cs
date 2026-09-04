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
                        var mail = raw as Outlook.MailItem;

                        if (mail == null)
                            continue;

                        string entryId = mail.EntryID;
                        string searchKey = MapiIdentity.GetSearchKeyHex(mail);
                        string recordKey = MapiIdentity.GetRecordKeyHex(mail);

                        if (string.IsNullOrEmpty(entryId) ||
                            string.IsNullOrEmpty(searchKey) ||
                            string.IsNullOrEmpty(recordKey))
                        {
                            continue;
                        }

                        result.Add(new SourceMessageDescriptor(
                            accountSmtp,
                            storeId,
                            entryId,
                            searchKey,
                            recordKey));
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

                // The only operation this adapter performs on the source object.
                return (Outlook.MailItem)original.Copy();
            }
            finally
            {
                ComUtil.Release(raw);
            }
        }
    }
}
