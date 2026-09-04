using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

namespace OutlookJunkRescuer
{
    internal sealed class SqliteStateStore : IDisposable
    {
        private const int CurrentSchemaVersion = 4;

        private readonly SQLiteConnection _connection;
        private readonly object _gate = new object();
        private bool _disposed;

        public SqliteStateStore(string databasePath)
        {
            string directory = Path.GetDirectoryName(databasePath);
            if (string.IsNullOrEmpty(directory))
                throw new ArgumentException("Database path must be absolute.", nameof(databasePath));

            Directory.CreateDirectory(directory);

            var builder = new SQLiteConnectionStringBuilder
            {
                DataSource = databasePath,
                Version = 3,
                Pooling = false,
                FailIfMissing = false
            };

            _connection = new SQLiteConnection(builder.ConnectionString);
            _connection.Open();

            ExecuteNonQuery("PRAGMA journal_mode=WAL;");
            ExecuteNonQuery("PRAGMA synchronous=FULL;");
            ExecuteNonQuery("PRAGMA foreign_keys=ON;");
            ExecuteNonQuery("PRAGMA busy_timeout=5000;");

            EnsureSchemaAndMigrate();
        }

        private void EnsureSchemaAndMigrate()
        {
            lock (_gate)
            {
                ExecuteNonQueryUnlocked(@"
CREATE TABLE IF NOT EXISTS message_state (
    store_id                    TEXT NOT NULL,
    search_key_hex              TEXT NOT NULL,
    account_smtp                TEXT NOT NULL,
    source_entry_id             TEXT NOT NULL,
    source_record_key_hex       TEXT NOT NULL,
    operation_id                TEXT NOT NULL,
    state                       INTEGER NOT NULL,
    working_copy_entry_id       TEXT NULL,
    working_copy_record_key_hex TEXT NULL,
    archive_entry_id            TEXT NULL,
    archive_record_key_hex      TEXT NULL,
    created_utc                 TEXT NOT NULL,
    updated_utc                 TEXT NOT NULL,
    PRIMARY KEY (account_smtp, search_key_hex),
    UNIQUE (operation_id)
);

CREATE INDEX IF NOT EXISTS idx_message_state_state
    ON message_state(state);

CREATE INDEX IF NOT EXISTS idx_message_state_account_state
    ON message_state(account_smtp, state);
");

                if (!ColumnExistsUnlocked("message_state", "archive_record_key_hex"))
                {
                    ExecuteNonQueryUnlocked(
                        "ALTER TABLE message_state ADD COLUMN archive_record_key_hex TEXT NULL;");
                }

                int version = GetUserVersionUnlocked();

                if (version == 0)
                {
                    // v3 had only Pending(0)/Archived(1), and Pending did not have
                    // a durable pre-Move barrier. Such rows may already have crossed
                    // Move(), so replaying Copy would be unsafe. Convert only legacy
                    // Pending rows to Uncertain. Archived=1 remains valid.
                    ExecuteNonQueryUnlocked(
                        "UPDATE message_state SET state = 4 WHERE state = 0;");

                    SetUserVersionUnlocked(CurrentSchemaVersion);
                }
                else if (version < CurrentSchemaVersion)
                {
                    throw new InvalidOperationException(
                        "Unsupported intermediate state database schema version: " + version);
                }
                else if (version > CurrentSchemaVersion)
                {
                    throw new InvalidOperationException(
                        "State database was created by a newer add-in version: " + version);
                }
            }
        }

        public MessageState Get(string accountSmtp, string searchKeyHex)
        {
            lock (_gate)
            {
                ThrowIfDisposed();

                using (var cmd = _connection.CreateCommand())
                {
                    cmd.CommandText = SelectSql + @"
WHERE account_smtp = @account_smtp AND search_key_hex = @search_key_hex;";
                    cmd.Parameters.AddWithValue("@account_smtp", accountSmtp);
                    cmd.Parameters.AddWithValue("@search_key_hex", searchKeyHex);

                    using (var reader = cmd.ExecuteReader())
                    {
                        return reader.Read() ? ReadState(reader) : null;
                    }
                }
            }
        }

        public List<MessageState> GetNonTerminalStates(string accountSmtp)
        {
            lock (_gate)
            {
                ThrowIfDisposed();

                var list = new List<MessageState>();
                using (var cmd = _connection.CreateCommand())
                {
                    cmd.CommandText = SelectSql + @"
WHERE account_smtp = @account_smtp 
  AND state IN (@pending, @copy_created, @moving, @uncertain);";
                    cmd.Parameters.AddWithValue("@account_smtp", accountSmtp);
                    cmd.Parameters.AddWithValue("@pending", (int)ArchiveState.Pending);
                    cmd.Parameters.AddWithValue("@copy_created", (int)ArchiveState.CopyCreated);
                    cmd.Parameters.AddWithValue("@moving", (int)ArchiveState.Moving);
                    cmd.Parameters.AddWithValue("@uncertain", (int)ArchiveState.Uncertain);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(ReadState(reader));
                        }
                    }
                }

                return list;
            }
        }

        public MessageState BeginOrGet(SourceMessageDescriptor source)
        {
            lock (_gate)
            {
                ThrowIfDisposed();

                using (var tx = _connection.BeginTransaction())
                {
                    MessageState existing = GetWithinTransaction(
                        tx,
                        source.AccountSmtp,
                        source.SearchKeyHex);

                    if (existing != null)
                    {
                        UpdateSourceLocatorWithinTransaction(tx, source);
                        tx.Commit();

                        existing.StoreId = source.StoreId;
                        existing.SourceEntryId = source.EntryId;
                        existing.SourceRecordKeyHex = source.RecordKeyHex;
                        existing.AccountSmtp = source.AccountSmtp;
                        return existing;
                    }

                    string now = DateTime.UtcNow.ToString("o");
                    string operationId = Guid.NewGuid().ToString("N");

                    using (var cmd = _connection.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"
INSERT INTO message_state (
    store_id, search_key_hex, account_smtp,
    source_entry_id, source_record_key_hex,
    operation_id, state,
    working_copy_entry_id, working_copy_record_key_hex,
    archive_entry_id, archive_record_key_hex,
    created_utc, updated_utc
) VALUES (
    @store_id, @search_key_hex, @account_smtp,
    @source_entry_id, @source_record_key_hex,
    @operation_id, @state,
    NULL, NULL,
    NULL, NULL,
    @created_utc, @updated_utc
);";

                        cmd.Parameters.AddWithValue("@store_id", source.StoreId);
                        cmd.Parameters.AddWithValue("@search_key_hex", source.SearchKeyHex);
                        cmd.Parameters.AddWithValue("@account_smtp", source.AccountSmtp);
                        cmd.Parameters.AddWithValue("@source_entry_id", source.EntryId);
                        cmd.Parameters.AddWithValue("@source_record_key_hex", source.RecordKeyHex);
                        cmd.Parameters.AddWithValue("@operation_id", operationId);
                        cmd.Parameters.AddWithValue("@state", (int)ArchiveState.Pending);
                        cmd.Parameters.AddWithValue("@created_utc", now);
                        cmd.Parameters.AddWithValue("@updated_utc", now);
                        cmd.ExecuteNonQuery();
                    }

                    tx.Commit();

                    return new MessageState
                    {
                        StoreId = source.StoreId,
                        SearchKeyHex = source.SearchKeyHex,
                        AccountSmtp = source.AccountSmtp,
                        SourceEntryId = source.EntryId,
                        SourceRecordKeyHex = source.RecordKeyHex,
                        OperationId = operationId,
                        State = ArchiveState.Pending
                    };
                }
            }
        }

        public void MarkCopyCreated(
            string accountSmtp,
            string searchKeyHex,
            WorkingCopyDescriptor copy)
        {
            lock (_gate)
            {
                ThrowIfDisposed();

                using (var cmd = _connection.CreateCommand())
                {
                    // CAS: Only allow transition from Pending or re-asserting CopyCreated.
                    // If already advanced to Moving or Archived, do not regress.
                    cmd.CommandText = @"
UPDATE message_state
SET state = @state,
    working_copy_entry_id = @entry_id,
    working_copy_record_key_hex = @record_key,
    updated_utc = @updated_utc
WHERE account_smtp = @account_smtp 
  AND search_key_hex = @search_key_hex
  AND state IN (@pending, @copy_created);";

                    cmd.Parameters.AddWithValue("@state", (int)ArchiveState.CopyCreated);
                    cmd.Parameters.AddWithValue("@entry_id", copy.EntryId);
                    cmd.Parameters.AddWithValue("@record_key", copy.RecordKeyHex);
                    cmd.Parameters.AddWithValue("@updated_utc", DateTime.UtcNow.ToString("o"));
                    cmd.Parameters.AddWithValue("@account_smtp", accountSmtp);
                    cmd.Parameters.AddWithValue("@search_key_hex", searchKeyHex);
                    cmd.Parameters.AddWithValue("@pending", (int)ArchiveState.Pending);
                    cmd.Parameters.AddWithValue("@copy_created", (int)ArchiveState.CopyCreated);

                    int rows = cmd.ExecuteNonQuery();
                    if (rows == 0)
                    {
                        var current = Get(accountSmtp, searchKeyHex);
                        if (current == null)
                            throw new InvalidOperationException($"MarkCopyCreated failed: row not found for {searchKeyHex}");

                        if (current.State == ArchiveState.Moving || current.State == ArchiveState.Archived)
                        {
                            return; // Already advanced; ignore CAS regression safely
                        }

                        throw new InvalidOperationException(
                            $"MarkCopyCreated CAS failed for {searchKeyHex}: current state is {current.State}");
                    }
                }
            }
        }

        public void RefreshWorkingCopyLocator(
            string accountSmtp,
            string searchKeyHex,
            WorkingCopyDescriptor copy)
        {
            lock (_gate)
            {
                ThrowIfDisposed();

                using (var cmd = _connection.CreateCommand())
                {
                    cmd.CommandText = @"
UPDATE message_state
SET working_copy_entry_id = @entry_id,
    working_copy_record_key_hex = @record_key,
    updated_utc = @updated_utc
WHERE account_smtp = @account_smtp AND search_key_hex = @search_key_hex;";

                    cmd.Parameters.AddWithValue("@entry_id", copy.EntryId);
                    cmd.Parameters.AddWithValue("@record_key", copy.RecordKeyHex);
                    cmd.Parameters.AddWithValue("@updated_utc", DateTime.UtcNow.ToString("o"));
                    cmd.Parameters.AddWithValue("@account_smtp", accountSmtp);
                    cmd.Parameters.AddWithValue("@search_key_hex", searchKeyHex);
                    RequireSingleRow(cmd.ExecuteNonQuery(), "RefreshWorkingCopyLocator");
                }
            }
        }

        public void MarkMoving(string accountSmtp, string searchKeyHex)
        {
            lock (_gate)
            {
                ThrowIfDisposed();

                using (var cmd = _connection.CreateCommand())
                {
                    // CAS: Only allow transition from CopyCreated, Moving, or Uncertain.
                    // Never allow downgrading terminal Archived back to Moving.
                    cmd.CommandText = @"
UPDATE message_state
SET state = @state,
    updated_utc = @updated_utc
WHERE account_smtp = @account_smtp 
  AND search_key_hex = @search_key_hex
  AND state IN (@copy_created, @moving, @uncertain);";

                    cmd.Parameters.AddWithValue("@state", (int)ArchiveState.Moving);
                    cmd.Parameters.AddWithValue("@updated_utc", DateTime.UtcNow.ToString("o"));
                    cmd.Parameters.AddWithValue("@account_smtp", accountSmtp);
                    cmd.Parameters.AddWithValue("@search_key_hex", searchKeyHex);
                    cmd.Parameters.AddWithValue("@copy_created", (int)ArchiveState.CopyCreated);
                    cmd.Parameters.AddWithValue("@moving", (int)ArchiveState.Moving);
                    cmd.Parameters.AddWithValue("@uncertain", (int)ArchiveState.Uncertain);

                    int rows = cmd.ExecuteNonQuery();
                    if (rows == 0)
                    {
                        var current = Get(accountSmtp, searchKeyHex);
                        if (current == null)
                            throw new InvalidOperationException($"MarkMoving failed: row not found for {searchKeyHex}");

                        if (current.State == ArchiveState.Archived)
                        {
                            return; // Already terminal Archived; do not regress to Moving
                        }

                        throw new InvalidOperationException(
                            $"MarkMoving CAS failed for {searchKeyHex}: current state is {current.State}");
                    }
                }
            }
        }

        public void MarkArchived(
            string accountSmtp,
            string searchKeyHex,
            ArchiveMatch archive)
        {
            if (archive == null)
                throw new ArgumentNullException(nameof(archive));

            lock (_gate)
            {
                ThrowIfDisposed();

                using (var cmd = _connection.CreateCommand())
                {
                    cmd.CommandText = @"
UPDATE message_state
SET state = @state,
    archive_entry_id = @archive_entry_id,
    archive_record_key_hex = @archive_record_key_hex,
    working_copy_entry_id = NULL,
    working_copy_record_key_hex = NULL,
    updated_utc = @updated_utc
WHERE account_smtp = @account_smtp AND search_key_hex = @search_key_hex;";

                    cmd.Parameters.AddWithValue("@state", (int)ArchiveState.Archived);
                    cmd.Parameters.AddWithValue("@archive_entry_id", archive.EntryId);
                    cmd.Parameters.AddWithValue("@archive_record_key_hex", (object)archive.RecordKeyHex ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@updated_utc", DateTime.UtcNow.ToString("o"));
                    cmd.Parameters.AddWithValue("@account_smtp", accountSmtp);
                    cmd.Parameters.AddWithValue("@search_key_hex", searchKeyHex);
                    RequireSingleRow(cmd.ExecuteNonQuery(), "MarkArchived");
                }
            }
        }

        private MessageState GetWithinTransaction(
            SQLiteTransaction tx,
            string accountSmtp,
            string searchKeyHex)
        {
            using (var cmd = _connection.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = SelectSql + @"
WHERE account_smtp = @account_smtp AND search_key_hex = @search_key_hex;";
                cmd.Parameters.AddWithValue("@account_smtp", accountSmtp);
                cmd.Parameters.AddWithValue("@search_key_hex", searchKeyHex);

                using (var reader = cmd.ExecuteReader())
                {
                    return reader.Read() ? ReadState(reader) : null;
                }
            }
        }

        private void UpdateSourceLocatorWithinTransaction(
            SQLiteTransaction tx,
            SourceMessageDescriptor source)
        {
            using (var cmd = _connection.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = @"
UPDATE message_state
SET store_id = @store_id,
    source_entry_id = @source_entry_id,
    source_record_key_hex = @source_record_key_hex,
    updated_utc = @updated_utc
WHERE account_smtp = @account_smtp AND search_key_hex = @search_key_hex;";

                cmd.Parameters.AddWithValue("@account_smtp", source.AccountSmtp);
                cmd.Parameters.AddWithValue("@store_id", source.StoreId);
                cmd.Parameters.AddWithValue("@source_entry_id", source.EntryId);
                cmd.Parameters.AddWithValue("@source_record_key_hex", source.RecordKeyHex);
                cmd.Parameters.AddWithValue("@updated_utc", DateTime.UtcNow.ToString("o"));
                cmd.Parameters.AddWithValue("@search_key_hex", source.SearchKeyHex);
                RequireSingleRow(cmd.ExecuteNonQuery(), "UpdateSourceLocator");
            }
        }

        private static MessageState ReadState(SQLiteDataReader reader)
        {
            return new MessageState
            {
                StoreId = reader.GetString(0),
                SearchKeyHex = reader.GetString(1),
                AccountSmtp = reader.GetString(2),
                SourceEntryId = reader.GetString(3),
                SourceRecordKeyHex = reader.GetString(4),
                OperationId = reader.GetString(5),
                State = (ArchiveState)reader.GetInt32(6),
                WorkingCopyEntryId = reader.IsDBNull(7) ? null : reader.GetString(7),
                WorkingCopyRecordKeyHex = reader.IsDBNull(8) ? null : reader.GetString(8),
                ArchiveEntryId = reader.IsDBNull(9) ? null : reader.GetString(9),
                ArchiveRecordKeyHex = reader.IsDBNull(10) ? null : reader.GetString(10)
            };
        }

        private const string SelectSql = @"
SELECT store_id, search_key_hex, account_smtp,
       source_entry_id, source_record_key_hex,
       operation_id, state,
       working_copy_entry_id, working_copy_record_key_hex,
       archive_entry_id, archive_record_key_hex
FROM message_state
";

        private void ExecuteNonQuery(string sql)
        {
            lock (_gate)
            {
                ThrowIfDisposed();
                ExecuteNonQueryUnlocked(sql);
            }
        }

        private void ExecuteNonQueryUnlocked(string sql)
        {
            using (var cmd = _connection.CreateCommand())
            {
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
            }
        }

        private bool ColumnExistsUnlocked(string table, string column)
        {
            using (var cmd = _connection.CreateCommand())
            {
                cmd.CommandText = "PRAGMA table_info(" + table + ");";

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string name = Convert.ToString(reader["name"]);
                        if (string.Equals(name, column, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }
            }

            return false;
        }

        private int GetUserVersionUnlocked()
        {
            using (var cmd = _connection.CreateCommand())
            {
                cmd.CommandText = "PRAGMA user_version;";
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private void SetUserVersionUnlocked(int version)
        {
            ExecuteNonQueryUnlocked("PRAGMA user_version=" + version + ";");
        }

        private static void RequireSingleRow(int affected, string operation)
        {
            if (affected != 1)
                throw new InvalidOperationException(
                    operation + " expected to update exactly one state row, but updated " + affected + ".");
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(SqliteStateStore));
        }

        public void Dispose()
        {
            lock (_gate)
            {
                if (_disposed)
                    return;

                _disposed = true;
                _connection.Dispose();
            }
        }
    }
}
