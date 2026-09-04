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
                "hotmail.com"
            };

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

        public void RunStartupSweep()
        {
            Outlook.Accounts accounts = null;

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

                        ProcessStore(store, smtp, storeId);
                    }
                    catch (Exception ex)
                    {
                        Logger.Write($"Account #{i} failed: {ex}");
                    }
                    finally
                    {
                        ComUtil.Release(store);
                        ComUtil.Release(account);
                    }
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
            string storeId)
        {
            Outlook.MAPIFolder junk = null;
            Outlook.MAPIFolder inbox = null;
            Outlook.MAPIFolder archive = null;

            int archived = 0;
            int recovered = 0;
            int uncertain = 0;
            int skipped = 0;
            int failed = 0;

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

                var groups = items
                    .GroupBy(x => x.SearchKeyHex, StringComparer.Ordinal)
                    .ToList();

                foreach (var group in groups)
                {
                    try
                    {
                        ProcessLogicalMessage(
                            group.ToList(),
                            archive,
                            ref archived,
                            ref recovered,
                            ref uncertain,
                            ref skipped);
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        Logger.Write(
                            $"[{smtp}] SearchKey {group.Key}: {ex}");
                    }
                }

                Logger.Write(
                    $"[{smtp}] sweep: visible={items.Count}, " +
                    $"logical={groups.Count}, archived={archived}, " +
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

        private void ProcessLogicalMessage(
            List<SourceMessageDescriptor> candidates,
            Outlook.MAPIFolder archive,
            ref int archived,
            ref int recovered,
            ref int uncertain,
            ref int skipped)
        {
            if (candidates == null || candidates.Count == 0)
                return;

            string accountSmtp = candidates[0].AccountSmtp;
            string searchKey = candidates[0].SearchKeyHex;

            MessageState state = _state.Get(accountSmtp, searchKey);

            if (state == null)
            {
                // No durable state exists, so this is a fresh replayable operation.
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
                    // Pending is the only existing state that may create a new
                    // copy. Re-identify the exact source before using Copy().
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
                        candidates,
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
                        candidates,
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
                        candidates,
                        state,
                        archive,
                        ref recovered))
                    {
                        uncertain++;
                        LogUncertain(state, "legacy v3 Pending outcome cannot be proven");
                    }
                    return;

                case ArchiveState.Archived:
                    skipped++;
                    return;

                default:
                    throw new InvalidOperationException(
                        "Unknown archive state: " + state.State);
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

                // Copy() -> marker Save() is the only unavoidable ownership gap.
                // If killed there, the unmarked orphan is never mutated later.
                _archiveWriter.StampOwnedCopy(
                    ownedCopy,
                    state.OperationId,
                    state.SearchKeyHex);

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

        private static bool IsSupportedAddress(string smtp)
        {
            if (string.IsNullOrWhiteSpace(smtp))
                return false;

            int at = smtp.LastIndexOf('@');
            if (at <= 0 || at == smtp.Length - 1)
                return false;

            string domain = smtp.Substring(at + 1).Trim();
            return AllowedConsumerDomains.Contains(domain);
        }

        private static string TryGetSmtpAddress(Outlook.Account account)
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
