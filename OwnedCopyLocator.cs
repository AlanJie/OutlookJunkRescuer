using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkRescuer
{
    /// <summary>
    /// Resolves objects that are already proven or potentially provable as
    /// plugin-owned copies. This class only reads; ArchiveWriter performs writes.
    /// </summary>
    internal sealed class OwnedCopyLocator
    {
        private readonly Outlook.NameSpace _session;

        public OwnedCopyLocator(Outlook.NameSpace session)
        {
            _session = session;
        }

        public WorkingCopyDescriptor Describe(Outlook.MailItem copy)
        {
            if (copy == null)
                throw new ArgumentNullException(nameof(copy));

            string entryId = copy.EntryID;
            string recordKey = MapiIdentity.GetRecordKeyHex(copy);

            if (string.IsNullOrEmpty(entryId) ||
                string.IsNullOrEmpty(recordKey))
            {
                throw new InvalidOperationException(
                    "The plugin-owned copy has no stable locator.");
            }

            return new WorkingCopyDescriptor(entryId, recordKey);
        }

        public bool IsInFolder(
            Outlook.MailItem item,
            Outlook.MAPIFolder folder)
        {
            object parent = null;

            try
            {
                parent = item.Parent;
                var parentFolder = parent as Outlook.MAPIFolder;

                if (parentFolder == null)
                    return false;

                string parentEntryId = parentFolder.EntryID;
                string targetEntryId = folder.EntryID;

                if (string.IsNullOrEmpty(parentEntryId) ||
                    string.IsNullOrEmpty(targetEntryId))
                {
                    return false;
                }

                // EntryIDs are opaque. Equality must be provider-aware.
                // A comparison failure is Unknown and intentionally propagates.
                return _session.CompareEntryIDs(
                    parentEntryId,
                    targetEntryId);
            }
            finally
            {
                ComUtil.Release(parent);
            }
        }

        public Outlook.MailItem ResolveJournaledCopy(MessageState state)
        {
            if (state == null ||
                string.IsNullOrEmpty(state.WorkingCopyEntryId) ||
                string.IsNullOrEmpty(state.WorkingCopyRecordKeyHex))
            {
                return null;
            }

            object raw = null;

            try
            {
                try
                {
                    raw = _session.GetItemFromID(
                        state.WorkingCopyEntryId,
                        state.StoreId);
                }
                catch (COMException)
                {
                    return null;
                }

                var candidate = raw as Outlook.MailItem;
                if (candidate == null)
                    return null;

                string recordKey = MapiIdentity.GetRecordKeyHex(candidate);
                string searchKey = MapiIdentity.GetSearchKeyHex(candidate);
                string marker = GetTextUserProperty(
                    candidate,
                    ArchiveWriter.ArchiveIdProperty);

                if (!string.Equals(
                        recordKey,
                        state.WorkingCopyRecordKeyHex,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        searchKey,
                        state.SearchKeyHex,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        marker,
                        state.OperationId,
                        StringComparison.Ordinal))
                {
                    return null;
                }

                Outlook.MailItem result = candidate;
                raw = null;
                return result;
            }
            finally
            {
                ComUtil.Release(raw);
            }
        }

        public Outlook.MailItem FindMarkedOwnedCopy(
            IEnumerable<SourceMessageDescriptor> candidates,
            string operationId,
            string searchKeyHex)
        {
            foreach (SourceMessageDescriptor candidate in candidates)
            {
                object raw = null;

                try
                {
                    try
                    {
                        raw = _session.GetItemFromID(
                            candidate.EntryId,
                            candidate.StoreId);
                    }
                    catch (COMException)
                    {
                        continue;
                    }

                    var mail = raw as Outlook.MailItem;
                    if (mail == null)
                        continue;

                    string marker = GetTextUserProperty(
                        mail,
                        ArchiveWriter.ArchiveIdProperty);

                    if (!string.Equals(
                            marker,
                            operationId,
                            StringComparison.Ordinal) ||
                        !string.Equals(
                            MapiIdentity.GetSearchKeyHex(mail),
                            searchKeyHex,
                            StringComparison.Ordinal))
                    {
                        continue;
                    }

                    Outlook.MailItem result = mail;
                    raw = null;
                    return result;
                }
                finally
                {
                    ComUtil.Release(raw);
                }
            }

            return null;
        }

        private static string GetTextUserProperty(
            Outlook.MailItem mail,
            string name)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;

            try
            {
                properties = mail.UserProperties;
                property = properties.Find(name, true);

                return property?.Value == null
                    ? null
                    : Convert.ToString(property.Value);
            }
            finally
            {
                ComUtil.Release(property);
                ComUtil.Release(properties);
            }
        }
    }
}
