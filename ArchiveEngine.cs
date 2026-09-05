using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkRescuer
{
    internal sealed class ArchiveEngine
    {
        private static readonly HashSet<string> AllowedConsumerDomains =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "outlook.com",
                "hotmail.com",
                "live.com",
                "live.cn",
                "msn.com"
            };

        public static SweepStatistics Statistics { get; } = new SweepStatistics();

        private readonly HashSet<string> _activeOperations =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private readonly Outlook.NameSpace _session;
        private readonly OutlookSourceReader _source;
        private readonly OwnedCopyLocator _ownedCopies;
        private readonly ArchiveWriter _archiveWriter;
        private readonly SqliteStateStore _state;

        public ArchiveEngine(
            Outlook.NameSpace session,
            SqliteStateStore state)
        {
            _session = session;
            _state = state;
            _source = new OutlookSourceReader(session);
            _ownedCopies = new OwnedCopyLocator(session);
            _archiveWriter = new ArchiveWriter();
        }

        public Outlook.MAPIFolder ResolveArchiveFolder(Outlook.Store store)
        {
            Outlook.MAPIFolder inbox = null;
            try
            {
                inbox = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                if (inbox == null)
                    return null;

                return _archiveWriter.GetOrCreateArchiveFolder(inbox);
            }
            finally
            {
                ComUtil.Release(inbox);
            }
        }

        public void RunStartupSweep()
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            Outlook.Accounts accounts = null;
            int totalVisible = 0;
            int totalArchived = 0;
            int totalRecovered = 0;
            int totalUncertain = 0;
            int totalSkipped = 0;
            int totalFailed = 0;

            try
            {
                accounts = _session.Accounts;

                for (int i = 1; i <= accounts.Count; i++)
                {
                    Outlook.Account account = null;
                    Outlook.Store store = null;

                    try
                    {
                        account = accounts[i];
                        string smtp = TryGetSmtpAddress(account);

                        if (!IsSupportedAddress(smtp))
                            continue;

                        smtp = smtp.Trim().ToLowerInvariant();

                        store = account.DeliveryStore;
                        if (store == null)
                        {
                            Logger.Write($"[{smtp}] No DeliveryStore; skipped.");
                            continue;
                        }

                        string storeId = store.StoreID;
                        if (string.IsNullOrEmpty(storeId))
                        {
                            Logger.Write($"[{smtp}] Empty StoreID; skipped.");
                            continue;
                        }

                        int storeArchived = 0;
                        int storeRecovered = 0;
                        int storeUncertain = 0;
                        int storeSkipped = 0;
                        int storeFailed = 0;
                        int storeVisible = 0;

                        ProcessStore(
                            store,
                            smtp,
                            storeId,
                            out storeVisible,
                            out storeArchived,
                            out storeRecovered,
                            out storeUncertain,
                            out storeSkipped,
                            out storeFailed);

                        totalVisible += storeVisible;
                        totalArchived += storeArchived;
                        totalRecovered += storeRecovered;
                        totalUncertain += storeUncertain;
                        totalSkipped += storeSkipped;
                        totalFailed += storeFailed;
                    }
                    catch (Exception ex)
                    {
                        Logger.Write($"Account #{i} failed: {ex}");
                        totalFailed++;
                    }
                    finally
                    {
                        ComUtil.Release(store);
                        ComUtil.Release(account);
                    }
                }

                sw.Stop();
                lock (Statistics)
                {
                    Statistics.LastSweepTime = DateTime.Now;
                    Statistics.LastSweepDurationMs = sw.ElapsedMilliseconds;
                    Statistics.LastVisibleCount = totalVisible;
                    Statistics.LastArchivedCount = totalArchived;
                    Statistics.LastRecoveredCount = totalRecovered;
                    Statistics.LastUncertainCount = totalUncertain;
                    Statistics.LastSkippedCount = totalSkipped;
                    Statistics.LastFailedCount = totalFailed;
                    Statistics.TotalArchivedSession += totalArchived;
                }
            }
            finally
            {
                ComUtil.Release(accounts);
            }
        }

        private void ProcessStore(
            Outlook.Store store,
            string smtp,
            string storeId,
            out int visible,
            out int archived,
            out int recovered,
            out int uncertain,
            out int skipped,
            out int failed)
        {
            Outlook.MAPIFolder junk = null;
            Outlook.MAPIFolder inbox = null;
            Outlook.MAPIFolder archive = null;

            visible = 0;
            archived = 0;
            recovered = 0;
            uncertain = 0;
            skipped = 0;
            failed = 0;

            try
            {
                junk = store.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderJunk);
                inbox = store.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox);

                if (junk == null || inbox == null)
                {
                    Logger.Write($"[{smtp}] Missing Inbox/Junk; skipped.");
                    return;
                }

                archive = _archiveWriter.GetOrCreateArchiveFolder(inbox);

                List<SourceMessageDescriptor> items =
                    _source.ReadJunk(smtp, storeId, junk);

                visible = items.Count;

                var itemsBySearchKey = items
                    .GroupBy(x => x.SearchKeyHex, StringComparer.Ordinal)
                    .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.Ordinal);

                // Ensure startup and reconciliation sweeps are driven by durable journal state:
                // Reconcile union of currently visible Junk SearchKeys and any non-terminal states
                // (Pending, CopyCreated, Moving, Uncertain) recorded in SQLite for this store.
                List<MessageState> nonTerminalStates = _state.GetNonTerminalStates(smtp, storeId);

                var allSearchKeys = new HashSet<string>(StringComparer.Ordinal);
                foreach (var k in itemsBySearchKey.Keys)
                {
                    allSearchKeys.Add(k);
                }
                foreach (var s in nonTerminalStates)
                {
                    allSearchKeys.Add(s.SearchKeyHex);
                }

                foreach (var searchKey in allSearchKeys)
                {
                    try
                    {
                        List<SourceMessageDescriptor> candidates;
                        if (!itemsBySearchKey.TryGetValue(searchKey, out candidates))
                        {
                            candidates = new List<SourceMessageDescriptor>();
                        }

                        ProcessLogicalMessage(
                            candidates,
                            archive,
                            ref archived,
                            ref recovered,
                            ref uncertain,
                            ref skipped,
                            smtp,
                            searchKey);
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        Logger.Write(
                            $"[{smtp}] SearchKey {searchKey}: {ex}");
                    }
                }

                Logger.Write(
                    $"[{smtp}] sweep: visible={items.Count}, " +
                    $"reconciledKeys={allSearchKeys.Count}, archived={archived}, " +
                    $"recovered={recovered}, uncertain={uncertain}, " +
                    $"skipped={skipped}, failed={failed}");
            }
            finally
            {
                ComUtil.Release(archive);
                ComUtil.Release(inbox);
                ComUtil.Release(junk);
            }
        }

        public bool ProcessSingleItem(
            object rawItem,
            string accountSmtp,
            string storeId,
            Outlook.MAPIFolder archiveFolder)
        {
            var descriptor = _source.TryReadDescriptor(accountSmtp, storeId, rawItem);
            if (descriptor == null)
                return false;

            int archived = 0;
            int recovered = 0;
            int uncertain = 0;
            int skipped = 0;

            try
            {
                ProcessLogicalMessage(
                    new List<SourceMessageDescriptor> { descriptor },
                    archiveFolder,
                    ref archived,
                    ref recovered,
                    ref uncertain,
                    ref skipped,
                    accountSmtp,
                    descriptor.SearchKeyHex);

                if (archived > 0)
                {
                    lock (Statistics)
                    {
                        Statistics.TotalArchivedSession++;
                        Statistics.TotalRealtimeIntercepted++;
                    }
                    Logger.Write($"[{accountSmtp}] Real-time intercepted and archived junk item: SearchKey={descriptor.SearchKeyHex}");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Logger.Write($"[{accountSmtp}] Error in ProcessSingleItem for {descriptor.SearchKeyHex}: {ex}");
                return false;
            }
        }

        private void ProcessLogicalMessage(
            List<SourceMessageDescriptor> candidates,
            Outlook.MAPIFolder archive,
            ref int archived,
            ref int recovered,
            ref int uncertain,
            ref int skipped,
            string accountSmtp,
            string searchKeyHex)
        {
            if (string.IsNullOrEmpty(accountSmtp) || string.IsNullOrEmpty(searchKeyHex))
                return;

            string activeKey = accountSmtp + ":" + searchKeyHex;
            lock (_activeOperations)
            {
                if (!_activeOperations.Add(activeKey))
                {
                    Logger.Write($"[{accountSmtp}] Operation already in progress for {searchKeyHex}; skipped re-entrant call.");
                    skipped++;
                    return;
                }
            }

            try
            {
                MessageState state = _state.Get(accountSmtp, searchKeyHex);

                if (state == null)
                {
                    if (candidates == null || candidates.Count == 0)
                        return;

                    SourceMessageDescriptor source = candidates[0];
                    state = _state.BeginOrGet(source);

                    ProcessPending(
                        candidates,
                        source,
                        state,
                        archive,
                        ref archived,
                        ref recovered);
                    return;
                }

                switch (state.State)
                {
                    case ArchiveState.Pending:
                    {
                        if (candidates == null || candidates.Count == 0)
                        {
                            skipped++;
                            return;
                        }

                        SourceMessageDescriptor source =
                            ChooseReadOnlySource(candidates, state);

                        state = _state.BeginOrGet(source);

                        ProcessPending(
                            candidates,
                            source,
                            state,
                            archive,
                            ref archived,
                            ref recovered);
                        return;
                    }

                    case ArchiveState.CopyCreated:
                        if (!RecoverOwnedCopyOrWait(
                            candidates ?? new List<SourceMessageDescriptor>(),
                            state,
                            archive,
                            true,
                            ref recovered))
                        {
                            uncertain++;
                            LogUncertain(state, "copy-created object is not currently provable/visible");
                        }
                        return;

                    case ArchiveState.Moving:
                        if (!RecoverOwnedCopyOrWait(
                            candidates ?? new List<SourceMessageDescriptor>(),
                            state,
                            archive,
                            true,
                            ref recovered))
                        {
                            uncertain++;
                            LogUncertain(state, "Move outcome is not currently observable");
                        }
                        return;

                    case ArchiveState.Uncertain:
                        if (!RecoverLegacyUncertain(
                            candidates ?? new List<SourceMessageDescriptor>(),
                            state,
                            archive,
                            ref recovered))
                        {
                            uncertain++;
                        }
                        return;

                    case ArchiveState.Archived:
                        skipped++;
                        return;

                    case ArchiveState.SourceGone:
                        if (candidates == null || candidates.Count == 0)
                        {
                            skipped++;
                            return;
                        }

                        // Reappeared in Junk: revive from SourceGone back to Pending and process
                        Logger.Write($"[{accountSmtp}] SearchKey {searchKeyHex} previously marked SourceGone has reappeared in Junk; reviving to Pending.");
                        SourceMessageDescriptor reviveSource = ChooseReadOnlySource(candidates, state);
                        state = _state.ReviveSourceGone(reviveSource);

                        ProcessPending(
                            candidates,
                            reviveSource,
                            state,
                            archive,
                            ref archived,
                            ref recovered);
                        return;

                    default:
                        throw new InvalidOperationException(
                            "Unknown archive state: " + state.State);
                }
            }
            finally
            {
                lock (_activeOperations)
                {
                    _activeOperations.Remove(activeKey);
                }
            }
        }

        private void ProcessPending(
            List<SourceMessageDescriptor> candidates,
            SourceMessageDescriptor source,
            MessageState state,
            Outlook.MAPIFolder archive,
            ref int archived,
            ref int recovered)
        {
            Outlook.MailItem ownedCopy = null;

            try
            {
                // Pending is replayable in v4: Move has definitely not started.
                // We still recover a copy that was stamped just before a crash,
                // avoiding an unnecessary duplicate when it is locally visible.
                ownedCopy = _ownedCopies.FindMarkedOwnedCopy(
                    candidates,
                    state.OperationId,
                    state.SearchKeyHex);

                if (ownedCopy != null)
                {
                    WorkingCopyDescriptor descriptor =
                        _ownedCopies.Describe(ownedCopy);

                    _state.MarkCopyCreated(
                        state.AccountSmtp,
                        state.SearchKeyHex,
                        descriptor);

                    MoveExistingOwnedCopy(
                        state,
                        ownedCopy,
                        archive);

                    recovered++;
                    return;
                }

                ownedCopy = _source.CreateCopy(source);
                if (ownedCopy == null)
                {
                    _state.MarkSourceGone(source.AccountSmtp, source.SearchKeyHex);
                    Logger.Write($"[{source.AccountSmtp}] Source item is no longer in Junk; marked SourceGone for SearchKey {source.SearchKeyHex}.");
                    return;
                }

                // Copy() -> marker Save() is the only unavoidable ownership gap.
                // If killed there, the unmarked orphan is never mutated later.
                _archiveWriter.StampOwnedCopy(
                    ownedCopy,
                    state.OperationId,
                    state.SearchKeyHex,
                    _state.GetOrCreateReplicaId());

                WorkingCopyDescriptor working =
                    _ownedCopies.Describe(ownedCopy);

                // This durable transition means Copy exists and Move has not yet
                // been invoked. Missing evidence after this point is fail-closed.
                _state.MarkCopyCreated(
                    source.AccountSmtp,
                    source.SearchKeyHex,
                    working);

                MoveExistingOwnedCopy(
                    state,
                    ownedCopy,
                    archive);

                archived++;
            }
            finally
            {
                ComUtil.Release(ownedCopy);
            }
        }

        private bool RecoverOwnedCopyOrWait(
            List<SourceMessageDescriptor> candidates,
            MessageState state,
            Outlook.MAPIFolder archive,
            bool allowMoveExistingCopy,
            ref int recovered)
        {
            Outlook.MailItem ownedCopy = null;

            try
            {
                ownedCopy = _ownedCopies.ResolveJournaledCopy(state);

                if (ownedCopy == null)
                {
                    ownedCopy = _ownedCopies.FindMarkedOwnedCopy(
                        candidates,
                        state.OperationId,
                        state.SearchKeyHex);
                }

                if (ownedCopy != null)
                {
                    if (_ownedCopies.IsInFolder(ownedCopy, archive))
                    {
                        ArchiveMatch committed =
                            _archiveWriter.DescribeOwnedCopy(ownedCopy);

                        ValidateCommittedMatch(state, committed);
                        _state.MarkArchived(
                            state.AccountSmtp,
                            state.SearchKeyHex,
                            committed);

                        recovered++;
                        return true;
                    }

                    if (!allowMoveExistingCopy)
                        return false;

                    WorkingCopyDescriptor descriptor =
                        _ownedCopies.Describe(ownedCopy);

                    // Refresh only the locator. Do not downgrade Moving/Uncertain
                    // back to CopyCreated during recovery.
                    _state.RefreshWorkingCopyLocator(
                        state.AccountSmtp,
                        state.SearchKeyHex,
                        descriptor);

                    MoveExistingOwnedCopy(
                        state,
                        ownedCopy,
                        archive);

                    recovered++;
                    return true;
                }

                // A null query result is NOT proof of absence: Cached Outlook may
                // not yet expose the server-side archive item. It can only prove
                // success when an item is found; otherwise caller waits/retries.
                ArchiveMatch archiveMatch =
                    _archiveWriter.FindByOperationId(
                        archive,
                        state.OperationId);

                if (archiveMatch == null)
                    return false;

                ValidateCommittedMatch(state, archiveMatch);
                _state.MarkArchived(
                    state.AccountSmtp,
                    state.SearchKeyHex,
                    archiveMatch);

                recovered++;
                return true;
            }
            finally
            {
                ComUtil.Release(ownedCopy);
            }
        }

        private bool RecoverLegacyUncertain(
            List<SourceMessageDescriptor> candidates,
            MessageState state,
            Outlook.MAPIFolder archive,
            ref int recovered)
        {
            // v3 Pending may already have crossed Move(). Never create a new copy.
            // First seek positive archive evidence; null remains Unknown.
            ArchiveMatch archiveMatch =
                _archiveWriter.FindByOperationId(
                    archive,
                    state.OperationId);

            if (archiveMatch == null)
            {
                archiveMatch = _archiveWriter.FindBySearchKey(
                    archive,
                    state.SearchKeyHex);
            }

            if (archiveMatch != null)
            {
                ValidateCommittedMatch(state, archiveMatch);
                _state.MarkArchived(
                    state.AccountSmtp,
                    state.SearchKeyHex,
                    archiveMatch);

                recovered++;
                return true;
            }

            // If a provably plugin-owned copy is still visible outside Archive,
            // moving that exact object is safe; no new copy is created.
            return RecoverOwnedCopyOrWait(
                candidates,
                state,
                archive,
                true,
                ref recovered);
        }

        private void MoveExistingOwnedCopy(
            MessageState state,
            Outlook.MailItem ownedCopy,
            Outlook.MAPIFolder archive)
        {
            // Write-ahead barrier: once Moving is durable, recovery must never
            // create another copy based on a negative/empty archive query.
            _state.MarkMoving(
                state.AccountSmtp,
                state.SearchKeyHex);

            ArchiveMatch moved =
                _archiveWriter.MoveOwnedCopy(
                    ownedCopy,
                    archive);

            ValidateCommittedMatch(state, moved);

            _state.MarkArchived(
                state.AccountSmtp,
                state.SearchKeyHex,
                moved);
        }

        private static void ValidateCommittedMatch(
            MessageState state,
            ArchiveMatch match)
        {
            if (match == null)
                throw new ArgumentNullException(nameof(match));

            if (!string.Equals(
                    match.SearchKeyHex,
                    state.SearchKeyHex,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Archive copy PR_SEARCH_KEY does not match journal state.");
            }

            if (!string.IsNullOrEmpty(match.OperationId) &&
                !string.Equals(
                    match.OperationId,
                    state.OperationId,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Archive copy operation marker does not match journal state.");
            }
        }

        private SourceMessageDescriptor ChooseReadOnlySource(
            List<SourceMessageDescriptor> candidates,
            MessageState state)
        {
            if (state == null)
                return candidates[0];

            SourceMessageDescriptor exactRecord = candidates.FirstOrDefault(
                x => string.Equals(
                    x.RecordKeyHex,
                    state.SourceRecordKeyHex,
                    StringComparison.Ordinal));

            if (exactRecord != null)
                return exactRecord;

            if (!string.IsNullOrEmpty(state.SourceEntryId))
            {
                foreach (SourceMessageDescriptor candidate in candidates)
                {
                    if (string.IsNullOrEmpty(candidate.EntryId))
                        continue;

                    // EntryID is opaque. Never infer inequality from string
                    // inequality; ask the provider whether the IDs are equivalent.
                    if (_session.CompareEntryIDs(
                        candidate.EntryId,
                        state.SourceEntryId))
                    {
                        return candidate;
                    }
                }
            }

            throw new InvalidOperationException(
                "The journaled source object is not currently visible in Junk; " +
                "refusing to substitute another same-SearchKey object.");
        }

        private static void LogUncertain(MessageState state, string reason)
        {
            Logger.Write(
                $"[{state.AccountSmtp}] SearchKey {state.SearchKeyHex} " +
                $"state={state.State}: {reason}; no Outlook object modified.");
        }

        public static bool IsSupportedAddress(string smtp)
        {
            if (string.IsNullOrWhiteSpace(smtp))
                return false;

            int at = smtp.LastIndexOf('@');
            if (at <= 0 || at == smtp.Length - 1)
                return false;

            string domain = smtp.Substring(at + 1).Trim();
            return AllowedConsumerDomains.Contains(domain);
        }

        public static string TryGetSmtpAddress(Outlook.Account account)
        {
            try
            {
                return account.SmtpAddress;
            }
            catch (COMException)
            {
                return null;
            }
        }
    }
}
