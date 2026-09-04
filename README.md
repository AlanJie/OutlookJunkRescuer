# OutlookJunkRescuer — read-only source / crash-safe journal edition

[English](README.md) | [中文说明](README.zh-CN.md)

A VSTO add-in for **Classic Outlook for Windows**.

It preserves the normal Junk workflow while keeping one long-term copy in a normal folder:

```text
Junk                    Inbox\Junk Archive
────                    ──────────────────
original stays here  -> plugin-owned copy
```

Only accounts whose SMTP domain is exactly `outlook.com` or `hotmail.com` are processed.

---

## Ownership Boundary

The original Junk item is user-owned and treated as strictly read-only.
`OutlookSourceReader` may only:

- Enumerate/read identifiers;
- Reopen a source item as a locator operation;
- Call `Copy()`.

It does **not** call `Save`, `Move`, `Delete`, write `UserProperties`, change flags/categories/read state, or otherwise mutate the source.

Plugin state lives separately in:

```text
%LOCALAPPDATA%\OutlookJunkRescuer\state-v3.sqlite
```

The filename remains `state-v3.sqlite` so an existing journal can be migrated in place. The internal SQLite schema version is 4.

Only plugin-created copies and `Inbox\Junk Archive` receive plugin metadata.

---

## MAPI Identity Model

The implementation deliberately separates identity from location:

- `PR_SEARCH_KEY` — logical message correlation across original/copy.
- `PR_RECORD_KEY` — concrete record validation.
- `EntryID` + `StoreID` — locator only.

`EntryID`s are never compared with string equality. Whenever two `EntryID`s must be compared semantically, the add-in calls `NameSpace.CompareEntryIDs()`.

A locator reopened with `GetItemFromID()` is not trusted merely because lookup succeeded: source/working-copy paths validate `PR_SEARCH_KEY` and/or `PR_RECORD_KEY` before using the object.

---

## Durable State Machine

The durable state machine transitions as follows:

```text
Pending
   |
   | Copy + ownership stamp + Save
   v
CopyCreated
   |
   | SQLite MarkMoving COMMIT   <-- write-ahead barrier
   v
Moving
   |
   | copy.Move(Junk Archive)
   v
Archived
```

There is also an `Uncertain` state:

- `Uncertain` is used for legacy v3 `Pending` rows because older code did not have a durable pre-`Move` barrier. An old `Pending` row therefore cannot safely be assumed replayable.

### State Details

- **`Pending`**:
  Indicates the v4 state machine has not yet begun `Move()`. A crash after raw `Copy()` but before the ownership marker is saved can leave an unmarked orphan in Junk. The add-in never mutates such an unknown object. It may create a new provably owned copy upon recovery; the orphan remains harmless in Junk. If the ownership marker was saved before the crash, recovery reuses that marked copy when it is locally visible.

- **`CopyCreated`**:
  A stamped plugin-owned copy exists and its `EntryID + PR_RECORD_KEY` locator was committed to SQLite. `Move()` has not yet been invoked by the v4 state machine. If that copy cannot currently be located, the add-in does **not** create another copy; it waits for a later recovery pass.

- **`Moving`**:
  Before invoking `Move()`, SQLite commits the `Moving` state. This is the non-replayable edge. After this point:
  - If the exact owned copy is still visible outside Archive, it may be moved;
  - If the exact owned copy is already in Archive, the row becomes `Archived`;
  - If `Junk Archive` positively finds `OJRArchiveId`, the row becomes `Archived`;
  - If neither source nor Archive currently exposes the object, the outcome is `Unknown` and no duplicate copy is created.

  In particular, `Items.Find(...) == null` is *not* interpreted as authoritative absence while recovering a `Moving` operation, because Cached Exchange / Outlook.com mode may not yet have completed server-side folder synchronization.

- **`Uncertain`**:
  On first schema-v4 startup, legacy rows with numeric `state=0` are migrated to `Uncertain`. Recovery may accept positive archive evidence or move a provably plugin-owned existing copy, but will never create duplicate copies from `Uncertain`.

---

## Folder-Level Custom Properties

The Archive folder registers:

- `OJRArchiveId`
- `OJRSearchKey`

in `Folder.UserDefinedProperties`, because Outlook folder-level `Items.Find` queries require the property definition to exist.

`EnsureQueryableFields()` only guarantees that the query schema is legal; it does not claim that the Cached Exchange / Outlook.com folder has completed initial synchronization. An empty Archive folder therefore legitimately returns `null` from `Items.Find`; COM/query failures propagate instead of being converted to `null`.

---

## SQLite Journal Architecture

`SqliteStateStore` maintains a single connection for the lifetime of the add-in.
The connection string explicitly specifies:

```text
Pooling=False
```

and the store serializes access through a private synchronization lock. SQLite is configured with:

```sql
PRAGMA journal_mode=WAL;
PRAGMA synchronous=FULL;
PRAGMA foreign_keys=ON;
PRAGMA busy_timeout=5000;
```

The database remains in `%LOCALAPPDATA%`; it is never stored alongside the VSTO assembly or native SQLite DLLs. If SQLite fails to initialize, the add-in safely disables the sweep for that session rather than operating without a durable journal.

---

## Installation and Distribution

Pre-built binaries are automatically built, signed, and packaged by GitHub Actions for every tagged release (available under GitHub Actions Artifacts and Releases). Two installation methods are supported:

### Option 1: ClickOnce Official Installer (Recommended)

Standard deployment managed directly by Office VSTO runtime with full Windows Settings and Control Panel uninstall support.

1. Download and extract `OutlookJunkRescuer-ClickOnce-Installer.zip` to any temporary directory.
2. Right-click `install-cert.bat` and select **Run as administrator** (this imports the self-signed public certificate into CurrentUser's `Root` and `TrustedPublisher` stores to satisfy Office VSTO trust requirements).
3. Double-click `OutlookJunkRescuer.vsto` and click **Install** in the Microsoft Office prompt.
4. **Uninstall**: Remove directly from Windows **Settings > Apps > Installed apps** (or Control Panel **Programs and Features**).

### Option 2: Portable Registration (`install.bat` / Zero-Certificate Bypass)

Ideal for users who prefer a portable, green installation without importing certificates into system certificate stores. This method registers the add-in using the `|vstolocal` flag under `HKCU`, instructing the VSTO runtime to load the add-in directly with local full trust.

1. Download and extract `OutlookJunkRescuer-Portable-Bat.zip` to a permanent folder (e.g. `C:\Tools\OutlookJunkRescuer`).
2. Double-click `install.bat` to register the add-in in Current User registry.
3. Start or restart Classic Outlook.
4. **Uninstall**: Double-click `uninstall.bat` to clean up the registry keys.

---

## Build from Source

### Prerequisites
- Visual Studio 2022 / 2026 with the **Office/SharePoint development** workload (`Microsoft.VisualStudio.Component.TeamOffice`).
- .NET Framework 4.8 Developer Pack.
- Classic Outlook for Windows.

### Building
```powershell
msbuild OutlookJunkRescuer.csproj /t:Restore
msbuild OutlookJunkRescuer.csproj /t:Build /p:Configuration=Release
```

Native SQLite binaries are deployed automatically to `bin\Release\x64` and `bin\Release\x86` based on the `System.Data.SQLite.Core` NuGet package. Ensure the bitness matches your installed Classic Outlook (x86 or x64).

---

## Project Structure

- `ThisAddIn.cs` — Add-in lifecycle and delayed startup sweep.
- `ArchiveEngine.cs` — Durable state machine and crash recovery policy.
- `OutlookSourceReader.cs` — Narrow read/Copy-only source inspection.
- `OwnedCopyLocator.cs` — Validates and reopens provably plugin-owned copies.
- `ArchiveWriter.cs` — Mutation layer for plugin-owned copies.
- `MapiIdentity.cs` — Helper for extracting `PR_SEARCH_KEY` and `PR_RECORD_KEY`.
- `SqliteStateStore.cs` — Single-connection WAL journal and schema migration.
- `Models.cs` — Data and descriptor types.
- `ComUtil.cs`, `Logger.cs` — Diagnostic and COM cleanup utilities.
- `install.bat`, `uninstall.bat` — Portable registration scripts.
- `install-cert.bat` — Certificate trust helper for ClickOnce.
