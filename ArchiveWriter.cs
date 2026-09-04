using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkRescuer
{
    /// <summary>
    /// Owns mutations of plugin-created objects only: the Junk Archive folder
    /// and archive copies. It never receives or mutates an original Junk item.
    /// </summary>
    internal sealed class ArchiveWriter
    {
        public const string ArchiveFolderName = "Junk Archive";
        public const string DuplicateTrashFolderName = "OutlookJunkRescuer Duplicate Trash";
        public const string LegacyDuplicateTrashFolderName = "Duplicate Trash";

        public const string PluginIdProperty = "OJRPluginId";
        public const string ArchiveKeyProperty = "OJRArchiveKey";
        public const string CopyIdProperty = "OJRCopyId";
        public const string ReplicaIdProperty = "OJRReplicaId";

        public const string ArchiveIdProperty = "OJRArchiveId";
        public const string SearchKeyProperty = "OJRSearchKey";

        public const string PluginIdValue = "OutlookJunkRescuer";

        public Outlook.MAPIFolder GetOrCreateArchiveFolder(
            Outlook.MAPIFolder inbox)
        {
            Outlook.Folders folders = null;

            try
            {
                folders = inbox.Folders;

                for (int i = 1; i <= folders.Count; i++)
                {
                    Outlook.MAPIFolder candidate = null;

                    try
                    {
                        candidate = folders[i];

                        if (string.Equals(
                            candidate.Name,
                            ArchiveFolderName,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            EnsureQueryableFields(candidate);
                            Outlook.MAPIFolder result = candidate;
                            candidate = null;
                            return result;
                        }
                    }
                    finally
                    {
                        ComUtil.Release(candidate);
                    }
                }

                Outlook.MAPIFolder created =
                    folders.Add(ArchiveFolderName, Type.Missing);

                EnsureQueryableFields(created);
                return created;
            }
            finally
            {
                ComUtil.Release(folders);
            }
        }

        public Outlook.MAPIFolder FindDuplicateTrashFolder(
            Outlook.MAPIFolder archive)
        {
            if (archive == null)
                return null;

            Outlook.Folders folders = null;

            try
            {
                folders = archive.Folders;

                // Prefer namespaced folder name
                for (int i = 1; i <= folders.Count; i++)
                {
                    Outlook.MAPIFolder candidate = null;

                    try
                    {
                        candidate = folders[i];

                        if (string.Equals(
                            candidate.Name,
                            DuplicateTrashFolderName,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            Outlook.MAPIFolder result = candidate;
                            candidate = null;
                            return result;
                        }
                    }
                    finally
                    {
                        ComUtil.Release(candidate);
                    }
                }

                // Fallback to legacy folder name if it already exists
                for (int i = 1; i <= folders.Count; i++)
                {
                    Outlook.MAPIFolder candidate = null;

                    try
                    {
                        candidate = folders[i];

                        if (string.Equals(
                            candidate.Name,
                            LegacyDuplicateTrashFolderName,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            Outlook.MAPIFolder result = candidate;
                            candidate = null;
                            return result;
                        }
                    }
                    finally
                    {
                        ComUtil.Release(candidate);
                    }
                }

                return null;
            }
            finally
            {
                ComUtil.Release(folders);
            }
        }

        public Outlook.MAPIFolder GetOrCreateDuplicateTrashFolder(
            Outlook.MAPIFolder archive)
        {
            if (archive == null)
                return null;

            Outlook.MAPIFolder existing = FindDuplicateTrashFolder(archive);
            if (existing != null)
                return existing;

            Outlook.Folders folders = null;
            try
            {
                folders = archive.Folders;
                Outlook.MAPIFolder created =
                    folders.Add(DuplicateTrashFolderName, Type.Missing);

                return created;
            }
            finally
            {
                ComUtil.Release(folders);
            }
        }

        public ArchiveMatch FindByOperationId(
            Outlook.MAPIFolder archive,
            string operationId)
        {
            return FindByTextProperty(
                archive,
                ArchiveIdProperty,
                operationId);
        }

        public ArchiveMatch FindBySearchKey(
            Outlook.MAPIFolder archive,
            string searchKeyHex)
        {
            return FindByTextProperty(
                archive,
                SearchKeyProperty,
                searchKeyHex);
        }

        public void StampOwnedCopy(
            Outlook.MailItem copy,
            string operationId,
            string searchKeyHex,
            string replicaId = null)
        {
            SetTextProperty(copy, PluginIdProperty, PluginIdValue);
            SetTextProperty(copy, ArchiveKeyProperty, searchKeyHex);
            SetTextProperty(copy, CopyIdProperty, operationId);
            if (!string.IsNullOrEmpty(replicaId))
            {
                SetTextProperty(copy, ReplicaIdProperty, replicaId);
            }

            // Maintain backward compatibility properties for v1.0.0
            SetTextProperty(copy, ArchiveIdProperty, operationId);
            SetTextProperty(copy, SearchKeyProperty, searchKeyHex);
            copy.Save();
        }

        public ArchiveMatch DescribeOwnedCopy(Outlook.MailItem copy)
        {
            if (copy == null)
                throw new ArgumentNullException(nameof(copy));

            return DescribeArchiveItem(copy);
        }

        public ArchiveMatch MoveOwnedCopy(
            Outlook.MailItem copy,
            Outlook.MAPIFolder archive)
        {
            Outlook.MailItem moved = null;

            try
            {
                moved = (Outlook.MailItem)copy.Move(archive);

                if (moved == null || string.IsNullOrEmpty(moved.EntryID))
                    throw new InvalidOperationException(
                        "Outlook moved the archive copy but returned no EntryID.");

                return DescribeArchiveItem(moved);
            }
            finally
            {
                ComUtil.Release(moved);
            }
        }

        private static ArchiveMatch FindByTextProperty(
            Outlook.MAPIFolder folder,
            string propertyName,
            string value)
        {
            Outlook.Items items = null;
            object found = null;

            try
            {
                items = folder.Items;

                // Values are generated operation IDs or hex search keys, never
                // user-controlled strings. If the folder is empty, Find returns
                // null. If the folder-level field is unavailable/query fails, the
                // COM exception intentionally propagates and callers fail closed.
                string filter =
                    "[" + propertyName + "] = '" + value + "'";

                found = items.Find(filter);

                var mail = found as Outlook.MailItem;
                if (mail == null)
                    return null;

                return DescribeArchiveItem(mail);
            }
            finally
            {
                ComUtil.Release(found);
                ComUtil.Release(items);
            }
        }

        private static ArchiveMatch DescribeArchiveItem(Outlook.MailItem mail)
        {
            string entryId = mail.EntryID;
            string recordKey = MapiIdentity.GetRecordKeyHex(mail);
            string actualSearchKey = MapiIdentity.GetSearchKeyHex(mail);
            string operationId = GetTextProperty(mail, ArchiveIdProperty);
            string stampedSearchKey = GetTextProperty(mail, SearchKeyProperty);

            if (string.IsNullOrEmpty(entryId) ||
                string.IsNullOrEmpty(recordKey) ||
                string.IsNullOrEmpty(actualSearchKey))
            {
                throw new InvalidOperationException(
                    "Archive item is missing required MAPI identity properties.");
            }

            if (!string.IsNullOrEmpty(stampedSearchKey) &&
                !string.Equals(
                    stampedSearchKey,
                    actualSearchKey,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Archive item search-key marker does not match PR_SEARCH_KEY.");
            }

            return new ArchiveMatch(
                entryId,
                recordKey,
                operationId,
                actualSearchKey);
        }

        private static void EnsureQueryableFields(
            Outlook.MAPIFolder archive)
        {
            // This only makes the custom fields legal for folder-level Find/
            // Restrict. It does NOT prove that a Cached Exchange/Outlook.com
            // folder is fully synchronized. Therefore a null query result is
            // observational, not globally authoritative during recovery.
            EnsureFolderField(archive, PluginIdProperty);
            EnsureFolderField(archive, ArchiveKeyProperty);
            EnsureFolderField(archive, CopyIdProperty);
            EnsureFolderField(archive, ReplicaIdProperty);
            EnsureFolderField(archive, ArchiveIdProperty);
            EnsureFolderField(archive, SearchKeyProperty);
        }

        private static void EnsureFolderField(
            Outlook.MAPIFolder folder,
            string name)
        {
            Outlook.UserDefinedProperties properties = null;
            Outlook.UserDefinedProperty existing = null;
            Outlook.UserDefinedProperty created = null;

            try
            {
                properties = folder.UserDefinedProperties;
                existing = properties.Find(name);

                if (existing != null)
                    return;

                created = properties.Add(
                    name,
                    Outlook.OlUserPropertyType.olText,
                    Type.Missing,
                    Type.Missing);
            }
            finally
            {
                ComUtil.Release(created);
                ComUtil.Release(existing);
                ComUtil.Release(properties);
            }
        }

        private static void SetTextProperty(
            Outlook.MailItem mail,
            string name,
            string value)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;

            try
            {
                properties = mail.UserProperties;
                property = properties.Find(name, true);

                if (property == null)
                {
                    property = properties.Add(
                        name,
                        Outlook.OlUserPropertyType.olText,
                        false,
                        Type.Missing);
                }

                property.Value = value;
            }
            finally
            {
                ComUtil.Release(property);
                ComUtil.Release(properties);
            }
        }

        private static string GetTextProperty(
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
