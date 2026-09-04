# OutlookJunkRescuer — read-only source / crash-safe journal edition

A VSTO add-in for **Classic Outlook for Windows**.

It preserves the normal Junk workflow while keeping one long-term copy in a
normal folder:

    Junk                    Inbox\Junk Archive
    ────                    ──────────────────
    original stays here  -> plugin-owned copy

Only accounts whose SMTP domain is exactly `outlook.com` or `hotmail.com` are
processed.

## Ownership boundary

The original Junk item is user-owned and treated as read-only.
`OutlookSourceReader` may only:

- enumerate/read identifiers;
- reopen a source item as a locator operation;
- call `Copy()`.

It does **not** call `Save`, `Move`, `Delete`, write `UserProperties`, change
flags/categories/read state, or otherwise mutate the source.

Plugin state lives separately in:

    %LOCALAPPDATA%\OutlookJunkRescuer\state-v3.sqlite

The filename remains `state-v3.sqlite` so an existing journal can be migrated
in place. The internal SQLite schema version is now 4.

Only plugin-created copies and `Inbox\Junk Archive` receive plugin metadata.

## MAPI identity model

The implementation deliberately separates identity from location:

- `PR_SEARCH_KEY` — logical message correlation across original/copy.
- `PR_RECORD_KEY` — concrete record validation.
- `EntryID` + `StoreID` — locator only.

EntryIDs are never compared with string equality. Whenever two EntryIDs must be
compared semantically, the add-in calls `NameSpace.CompareEntryIDs()`.

A locator reopened with `GetItemFromID()` is not trusted merely because lookup
succeeded: source/working-copy paths validate `PR_SEARCH_KEY` and/or
`PR_RECORD_KEY` before using the object.

## State machine

The durable state machine is:

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

There is also:

    Uncertain

`Uncertain` is used for legacy v3 `Pending` rows because the older code did not
have a durable pre-Move barrier. An old `Pending` row therefore cannot safely be
assumed replayable.

### Pending

`Pending` means the new state machine has not begun `Move()`.

A crash after raw `Copy()` but before the ownership marker is saved can leave an
unmarked orphan in Junk. The add-in never mutates such an unknown object. It may
create a new provably owned copy on recovery; the orphan remains harmless in
Junk.

If the ownership marker was saved before the crash, recovery reuses that marked
copy when it is locally visible.

### CopyCreated

A stamped plugin-owned copy exists and its `EntryID + PR_RECORD_KEY` locator was
committed to SQLite. `Move()` has not yet been invoked by the v4 state machine.

If that copy cannot currently be proven/located, the add-in does **not** create
another copy. It waits for a later recovery pass.

### Moving

Before invoking `Move()`, SQLite is committed to `Moving`.

This is the non-replayable edge. After this point:

- if the exact owned copy is still visible outside Archive, it may be moved;
- if the exact owned copy is already in Archive, the row becomes `Archived`;
- if `Junk Archive` positively finds `OJRArchiveId`, the row becomes `Archived`;
- if neither source nor Archive currently exposes the object, the outcome is
  **Unknown** and no new copy is created.

In particular:

    Items.Find(...) == null

is *not* interpreted as globally authoritative absence while recovering a
`Moving` operation. Cached Outlook/Outlook.com may not yet expose all
server-side folder contents.

### Uncertain

On first schema-v4 startup, legacy rows with numeric `state=0` are migrated to
`Uncertain`. Recovery may:

- accept positive archive evidence;
- move a provably plugin-owned existing copy;
- otherwise do nothing and retry later.

It never creates a new copy from `Uncertain`.

## Folder-level custom properties

The Archive folder registers:

- `OJRArchiveId`
- `OJRSearchKey`

in `Folder.UserDefinedProperties`, because Outlook folder-level `Items.Find`
queries require the property definition to exist.

`EnsureQueryableFields()` only guarantees that the query schema is legal. It
does **not** claim that the Cached Exchange/Outlook.com folder has completed
initial synchronization.

An empty Archive folder therefore legitimately returns `null` from `Items.Find`;
a COM/query failure propagates instead of being converted to `null`.

## SQLite

`SqliteStateStore` owns one connection for the add-in lifetime.

The connection string explicitly sets:

    Pooling=False

and the store serializes access through a private lock. SQLite is configured
with:

    PRAGMA journal_mode=WAL;
    PRAGMA synchronous=FULL;
    PRAGMA foreign_keys=ON;
    PRAGMA busy_timeout=5000;

The database remains under `%LOCALAPPDATA%`; it is not stored next to the VSTO
assembly/native SQLite DLLs.

If SQLite cannot initialize, the add-in disables the sweep for that Outlook run
rather than operating without its durable journal.

## System.Data.SQLite deployment

Keep code/native dependencies with the add-in, for example:

    <addin>\
      OutlookJunkRescuer.dll
      System.Data.SQLite.dll
      x86\SQLite.Interop.dll
      x64\SQLite.Interop.dll

Mutable state remains separate under `%LOCALAPPDATA%`.

The SQLite native architecture must match the installed Classic Outlook process
(x86 or x64), not merely the Windows architecture.

## Build

1. Install Visual Studio 2022.
2. Enable the **Office/SharePoint development** workload.
3. Create an **Outlook VSTO Add-in** project targeting .NET Framework.
4. Name it `OutlookJunkRescuer`.
5. Add the source files from this archive to the project.
6. Install/reference `System.Data.SQLite.Core`.
7. Verify the correct `SQLite.Interop.dll` deployment for the Outlook bitness.
8. Build and run with Classic Outlook.

## Files

- `ThisAddIn.cs` — lifecycle and delayed startup sweep.
- `ArchiveEngine.cs` — state machine and recovery policy.
- `OutlookSourceReader.cs` — narrow read/Copy-only source capability.
- `OwnedCopyLocator.cs` — validates/reopens provably plugin-owned copies.
- `ArchiveWriter.cs` — the only Outlook mutation layer for plugin-owned copies.
- `MapiIdentity.cs` — reads `PR_SEARCH_KEY` / `PR_RECORD_KEY`.
- `SqliteStateStore.cs` — single-connection WAL journal and v3->v4 migration.
- `Models.cs` — state/descriptor types.
- `ComUtil.cs`, `Logger.cs` — support utilities.

## Remaining unavoidable crash window

Outlook OOM does not provide a transaction spanning `MailItem.Copy()`, writing
our ownership marker, SQLite, and `MailItem.Move()`.

The only intentionally accepted ownership gap is:

    original.Copy()
        <process killed here>
    stamp copy with OJRArchiveId

A copy left in that window is not provably plugin-owned and is therefore never
moved/deleted by recovery. This is preferred over risking mutation of a
user-owned/unknown item.

## 安装与分发 (Installation & Distribution)

每次代码提交或发布时，GitHub Actions 均会自动编译并提供两种开箱即用的分发包（位于 Actions Artifacts 中）：

### 方式一：ClickOnce 标准安装包（官方标准推荐）

适用于希望像正规 Office 插件一样由 Office 统一托管、支持控制面板标准卸载的用户。

1. 解压 `OutlookJunkRescuer-ClickOnce-Installer.zip` 到任意临时目录；
2. 右键管理员运行 `install-cert.bat`（将自签名公钥导入当前用户的受信任根证书与受信任发布者，避免 Office VSTO 信任提示拦截）；
3. 双击 `OutlookJunkRescuer.vsto`，在弹出的 Microsoft Office 提示框中点击“安装”即可；
4. 卸载：直接在 Windows“设置” -> “应用和功能”（或控制面板“程序和功能”）中找到 `OutlookJunkRescuer` 卸载。

### 方式二：便携免安装版（注册表一键安装，完全无需导入证书）

适用于不希望向系统证书库导入自签名证书、希望绿色免安装的用户。通过 VSTO `|vstolocal` 本地加载标志，VSTO 运行时直接以全信任本地加载，无需任何证书验证。

1. 将 `OutlookJunkRescuer-Portable-Bat.zip` 解压到你打算长期存放的文件夹（如 `C:\Tools\OutlookJunkRescuer`）；
2. 双击运行 `install.bat`，即可一键注册至当前用户的 Outlook 插件列表；
3. 打开 Classic Outlook 即可开始自动工作；
4. 卸载：双击运行 `uninstall.bat` 即可一键清理注册表。

