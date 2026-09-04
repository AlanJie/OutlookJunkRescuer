namespace OutlookJunkRescuer
{
    // Numeric values intentionally preserve the v3 database meaning:
    // 0 = old Pending, 1 = old Archived. Migration rewrites legacy 0 rows
    // to Uncertain before the new state machine starts using Pending=0.
    internal enum ArchiveState
    {
        Pending = 0,
        Archived = 1,
        CopyCreated = 2,
        Moving = 3,
        Uncertain = 4
    }

    internal sealed class SourceMessageDescriptor
    {
        public string AccountSmtp { get; }
        public string StoreId { get; }
        public string EntryId { get; }
        public string SearchKeyHex { get; }
        public string RecordKeyHex { get; }

        public SourceMessageDescriptor(
            string accountSmtp,
            string storeId,
            string entryId,
            string searchKeyHex,
            string recordKeyHex)
        {
            AccountSmtp = accountSmtp;
            StoreId = storeId;
            EntryId = entryId;
            SearchKeyHex = searchKeyHex;
            RecordKeyHex = recordKeyHex;
        }
    }

    internal sealed class WorkingCopyDescriptor
    {
        public string EntryId { get; }
        public string RecordKeyHex { get; }

        public WorkingCopyDescriptor(string entryId, string recordKeyHex)
        {
            EntryId = entryId;
            RecordKeyHex = recordKeyHex;
        }
    }

    internal sealed class MessageState
    {
        public string StoreId { get; set; }
        public string SearchKeyHex { get; set; }
        public string AccountSmtp { get; set; }
        public string SourceEntryId { get; set; }
        public string SourceRecordKeyHex { get; set; }
        public string OperationId { get; set; }
        public ArchiveState State { get; set; }
        public string WorkingCopyEntryId { get; set; }
        public string WorkingCopyRecordKeyHex { get; set; }
        public string ArchiveEntryId { get; set; }
        public string ArchiveRecordKeyHex { get; set; }
    }

    internal sealed class ArchiveMatch
    {
        public string EntryId { get; }
        public string RecordKeyHex { get; }
        public string OperationId { get; }
        public string SearchKeyHex { get; }

        public ArchiveMatch(
            string entryId,
            string recordKeyHex,
            string operationId,
            string searchKeyHex)
        {
            EntryId = entryId;
            RecordKeyHex = recordKeyHex;
            OperationId = operationId;
            SearchKeyHex = searchKeyHex;
        }
    }

    internal sealed class SweepStatistics
    {
        public System.DateTime LastSweepTime { get; set; } = System.DateTime.MinValue;
        public long LastSweepDurationMs { get; set; }
        public int LastVisibleCount { get; set; }
        public int LastArchivedCount { get; set; }
        public int LastRecoveredCount { get; set; }
        public int LastUncertainCount { get; set; }
        public int LastSkippedCount { get; set; }
        public int LastFailedCount { get; set; }
        public int TotalArchivedSession { get; set; }
        public int TotalRealtimeIntercepted { get; set; }
    }
}
