# OutlookJunkRescuer — 只读源 / 防崩溃日志版

[English](README.md) | [中文说明](README.zh-CN.md)

专为 **Windows 经典版 Outlook (Classic Outlook)** 设计的 VSTO 邮件防误判归档插件。

在保持垃圾邮件箱（Junk）原有工作流程与微软服务器判定完全不变的前提下，自动在常规邮件文件夹中保留一份可长期保存的插件专属副本：

```text
垃圾邮件 (Junk)           收件箱\Junk Archive
───────────────          ──────────────────
原始邮件保持不动      ->   插件托管的安全归档副本
```

仅处理发件箱或账户 SMTP 域名属于微软个人邮箱生态（`@outlook.com`、`@hotmail.com`、`@live.com`、`@live.cn`、`@msn.com`）的账户。

---

## 邮件所有权边界 (Ownership Boundary)

原始垃圾邮件属于用户资产，插件严格将其视为**只读**。
`OutlookSourceReader` 仅允许执行以下安全操作：

- 枚举并读取标识符；
- 将源邮件项作为定位标识符重新打开；
- 调用 `Copy()` 生成副本。

插件**绝不**对原始邮件调用 `Save`、`Move`、`Delete`，绝不修改其 `UserProperties`、标志、分类、已读状态或产生任何变更。

插件状态独立保存在本地数据库中：

```text
%LOCALAPPDATA%\OutlookJunkRescuer\state-v3.sqlite
```

数据库文件名保持为 `state-v3.sqlite` 以便就地迁移旧版日志。内部 SQLite 数据架构版本为 4。

仅插件创建的专属副本以及目标文件夹 `收件箱\Junk Archive` 会被写入插件专属元数据。

---

## MAPI 身份识别模型 (MAPI Identity Model)

实现中严格将“身份识别”与“位置信息”正交解耦：

- `PR_SEARCH_KEY` — 用于跨源邮件与副本的逻辑消息关联。
- `PR_RECORD_KEY` — 用于校验具体的实体记录一致性。
- `EntryID` + `StoreID` — 仅作为重新打开对象的物理定位符。

`EntryID` 绝不使用普通字符串相等性（`==`）进行比较。凡涉及两个 `EntryID` 的语义比对，统一调用 `NameSpace.CompareEntryIDs()`。

通过 `GetItemFromID()` 重新打开的定位项不会仅仅因为“查找成功”就被信任：源路径与工作副本路径在实际使用该对象前，必须严格比对校验 `PR_SEARCH_KEY` 和/或 `PR_RECORD_KEY`。

---

## 防崩溃持久化状态机 (State Machine)

可靠的状态机流转如下：

```text
Pending (等待处理)
   |
   | 复制原始邮件 + 写入防混淆所有权标记 + Save
   v
CopyCreated (副本已创建)
   |
   | SQLite MarkMoving COMMIT   <-- 预写式屏障 (Write-Ahead Barrier)
   v
Moving (正在移动)
   |
   | copy.Move(Junk Archive)
   v
Archived (归档完成)
```

另有 `Uncertain`（不确定）状态：

- `Uncertain` 用于处理从旧版 v3 迁移过来的 `Pending` 记录，因为老版本代码没有持久化的移动前预写屏障。旧的 `Pending` 记录无法确定是否已经执行过移动，因此不可安全重放。

### 状态详细说明

- **`Pending`（待处理）**：
  表示 v4 状态机尚未开始调用 `Move()`。若在 `Copy()` 之后但尚未写入所有权标记时发生异常崩溃，可能会在垃圾箱中遗留一个无标记孤儿项。插件绝对不会修改此类未知对象。系统在下次恢复运行时会重新生成一个可证明所有权的新副本；无标记孤儿项安全保留在垃圾箱中。若所有权标记在崩溃前已持久化保存，恢复引擎会优先复用该已有副本。

- **`CopyCreated`（副本已建）**：
  表明已盖上插件所有权戳记的副本已存在，且其 `EntryID + PR_RECORD_KEY` 定位信息已提交到 SQLite。此时 v4 状态机尚未开始 `Move()`。如果该副本当前暂时无法在本地定位，插件**绝不会**创建另一个新副本，而是等待后续恢复轮次再次重试。

- **`Moving`（移动中）**：
  在调用 `Move()` 之前，SQLite 必须先提交为 `Moving`。这是不可回退的关键边界。在此之后：
  - 若在归档文件夹外能正确定位到该专属副本，则继续执行移动；
  - 若该专属副本已位于归档文件夹内，则状态转为 `Archived`；
  - 若归档文件夹内已正向检测到匹配的 `OJRArchiveId`，则状态转为 `Archived`；
  - 若源文件夹和归档文件夹当前均未暴露该对象，则判定为 `Unknown`，**绝不创建重复副本**。

  特别注意：在恢复 `Moving` 操作时，`Items.Find(...) == null` **并不**代表邮件在全局绝对缺失，因为 Exchange 缓存模式或 Outlook.com 可能尚未完成服务端后台同步。

- **`Uncertain`（不确定状态）**：
  在升级到架构 v4 首次启动时，数值 `state=0` 的旧版记录会自动迁移为 `Uncertain`。恢复策略仅在归档文件夹中找到明确证据或找到已确权的孤立副本时才会尝试移动，绝不会基于 `Uncertain` 凭空创建新副本。

---

## 双轨防护与诊断界面 (Dual-Track Engine & UI)

### 1. 实时拦截保护 (Fast Path)
插件启动后会自动为所有符合域名的有效账户的垃圾邮件箱（Junk Folder）挂载事件监听器（`JunkFolderWatcher`），监听 COM 的 `Items.ItemAdd` 事件：
- **实时响应**：每当有新邮件落入垃圾箱时，无需等待扫描周期，毫秒级在后台无感完成副本创建与移动保护；
- **防 GC 释放机制**：在插件托管生命周期内，对垃圾箱 `MAPIFolder` 与 `Items` 集合持有强引用字段，规避 .NET COM RCW 被垃圾回收器提早回收导致监听失效的问题；
- **状态机保障**：即使是实时拦截路径，完全复用同一套 v4 持久化状态机与预写式事务屏障，确保在任何断电、崩溃场景下的数据幂等与绝对安全。

### 2. 启动与手动对账扫描 (Reconciliation Sweep)
- 经典版 Outlook 启动 15 秒后自动执行一次全量垃圾箱对账扫描，补全客户端离线期间在服务端落入垃圾箱的邮件，并自动重放、恢复任何未完成的事务；
- 支持随时在控制台中手动触发全量对账扫描。

### 3. 功能区与诊断控制台 (Ribbon & Status Console)
- **Ribbon 菜单**：在 Outlook 主界面“开始 (邮件)”选项卡中内置原生 “Junk Rescuer” 分组，包含：
  - **运行状态**：唤起运行状态与诊断控制台；
  - **清理重复项**：唤起跨设备重复副本检测与清理控制台。
- **诊断控制台 (`StatusForm`)**：
  - 实时保护状态与当前受保护的邮箱账户列表；
  - 最近一次扫描时间、耗时、处理状态统计（Archived / Skipped / Uncertain / Failed 以及实时拦截计数）；
  - 本地 SQLite 状态数据库绝对路径与文件大小；
  - 一键执行对账扫描、一键打开清理重复副本窗口、一键打开日志和数据存储目录。

---

## 多设备安全归档与保守重复副本清理

### 1. 产品核心原则：“永不漏备 (Never-Miss)”优先于“绝不重复 (Never-Duplicate)”
当同一邮箱在多台设备（如 PC-A 与 PC-B）上同时运行 OutlookJunkRescuer 时：
- 每台设备各自保证本地的崩溃安全与 SQLite 事务状态机；
- 多台机器对同一封垃圾邮件独立创建归档副本属于模型显式允许的**良性副本**；
- **$O(\text{Junk})$ 日常运行铁律**：日常实时归档与启动对账**完全无需扫描 `Junk Archive` 历史目录**。即便归档文件夹累积了 10 万+ 历史邮件，日常启动与归档性能与历史条目完全解耦，毫秒级响应。

### 2. 副本元数据与稳定设备标识 (ReplicaId)
每份归档副本均打上专用 UserProperties 元数据：
- `OJRPluginId`：`"OutlookJunkRescuer"`（所有权校验）；
- `OJRArchiveKey`：十六进制格式的 `PR_SEARCH_KEY`；
- `OJRCopyId`：单次副本唯一 GUID；
- `OJRReplicaId`：当前机器的持久化稳定 UUID（首次运行自动生成并持久化在本地 SQLite 中）。
- 保持对 v1.0.0 旧副本属性（`OJRArchiveId` 与 `OJRSearchKey`）的完全兼容。

### 3. 保守重复副本清理 (`DuplicateCleanupForm`)
- **人工显式触发**：仅在用户显式点击“扫描重复项”时，才对 `Junk Archive` 执行轻量 `Table` 遍历。
- **保守重验证原则 (Conservative Revalidation)**：
  - 验证留存的 Winner 副本完好且位于 `Junk Archive`；
  - 逐一重新验证每个待移动的 Loser 副本；
  - **Never reduce 1 -> 0 铁律**：硬性保证归档目录中任何邮件的有效副本绝不归零；
  - 多余副本移动至 `收件箱\Junk Archive\Duplicate Trash` 软隔离目录，不执行不可逆的物理删除；
  - 提供一键在 Outlook 中审查垃圾桶、以及确认后清空垃圾桶的功能。

---

## 文件夹级自定义字段 (Folder Properties)

归档文件夹会自动注册：

- `OJRArchiveId`
- `OJRSearchKey`

至 `Folder.UserDefinedProperties` 中，因为 Outlook 文件夹级别的 `Items.Find` 查询语法要求自定义字段必须已预先注册定义。

`EnsureQueryableFields()` 仅确保查询字段定义合法，不代表 Exchange 缓存模式已完成全量同步。因此空文件夹返回 `null` 是正常情况，而底层 COM 错误会被如实抛出而不是被吞掉掩盖为 `null`。

---

## SQLite 持久化架构

`SqliteStateStore` 在插件整个生命周期内只独占一条持久数据库连接。
连接字符串显式关闭连接池：

```text
Pooling=False
```

并且所有数据库访问均由私有同步锁严格序列化执行。SQLite 采用高可靠配置：

```sql
PRAGMA journal_mode=WAL;
PRAGMA synchronous=FULL;
PRAGMA foreign_keys=ON;
PRAGMA busy_timeout=5000;
```

数据库常驻于 `%LOCALAPPDATA%`，绝不随 VSTO 程序集目录存放。如果 SQLite 无法初始化，插件会自动停用该会话的扫描任务，而绝不脱离日志保护裸奔运行。

---

## 安装与分发方式 (Installation & Distribution)

每次代码推送发布标签（Tag）时，GitHub Actions 均会自动执行云端编译、代码签名，并生成两种开箱即用的分发包（位于 GitHub Releases 或 Actions Artifacts 中）：

### 方式一：ClickOnce 官方标准安装包（微软标准推荐）

适用于希望由 Office 统一托管、支持控制面板标准卸载的用户。

1. 下载并解压 `OutlookJunkRescuer-ClickOnce-Installer.zip` 到任意临时目录；
2. 右键管理员身份运行 `install-cert.bat`（将打包中随附的自签名公钥导入当前用户的“受信任的根证书颁发机构”与“受信任的发布者”，避免 Office VSTO 信任安全拦截）；
3. 双击 `OutlookJunkRescuer.vsto`，在弹出的 Microsoft Office 提示框中点击“安装”即可；
4. **卸载**：直接在 Windows“设置” -> “应用和功能”（或控制面板“程序和功能”）中找到 `OutlookJunkRescuer` 卸载。

### 方式二：便携免安装版（注册表一键就地注册）

适用于希望绿色便携、就地加载的用户。该方式通过在注册表 Manifest 路径后添加 `|vstolocal` 标志，指示 VSTO 运行时直接从程序集所在的本地目录加载运行，而不是将其复制到 ClickOnce 缓存目录中。

> [!NOTE]
> `|vstolocal` 改变的是加载位置（本地目录 vs. ClickOnce 缓存），VSTO 仍遵循 Windows 与 Office 信任体系。如果您的系统组策略对未受信任加载项实施了严格阻断，建议优先使用方式一或在 Outlook 信任中心允许加载。

1. 将 `OutlookJunkRescuer-Portable-Bat.zip` 解压到你打算长期存放的文件夹（例如 `C:\Tools\OutlookJunkRescuer`）；
2. 双击运行 `install.bat`，即可一键注册至当前用户的 Outlook 加载项；
3. 打开经典版 Outlook 即可自动开始工作；
4. **卸载**：双击运行 `uninstall.bat` 即可干净清理注册表加载项。

---

## 源码编译说明 (Build from Source)

### 环境要求
- Visual Studio 2022 / 2026，需勾选 **Office/SharePoint 开发** 工作负载（`Microsoft.VisualStudio.Component.TeamOffice`）；
- .NET Framework 4.8 目标包；
- 经典版 Windows Outlook。

### 编译命令
```powershell
msbuild OutlookJunkRescuer.csproj /t:Restore
msbuild OutlookJunkRescuer.csproj /t:Build /p:Configuration=Release
```

通过 NuGet 引入的 `System.Data.SQLite.Core` 会自动根据当前目标输出将 `x64\SQLite.Interop.dll` 与 `x86\SQLite.Interop.dll` 拷贝至输出目录。请确保 SQLite 原生运行时位数与实际安装的 Outlook 位数（32位或64位）一致。

---

## 项目文件结构说明

- `ThisAddIn.cs` — 插件生命周期管理、事件监听挂载与启动延时扫描调度。
- `JunkFolderWatcher.cs` — 垃圾箱实时 `ItemAdd` 事件监听器与 COM 强引用持有。
- `Ribbon.cs` — Outlook Explorer 功能区 (Ribbon) 扩展与菜单按钮定义。
- `StatusForm.cs` — 运行状态与诊断信息可视化控制台窗口。
- `DuplicateCleanupForm.cs` — 跨设备重复归档副本可视化检测与清理窗口。
- `DuplicateCleaner.cs` — 快速扫表与保守重验证重复项清理核心引擎。
- `ArchiveEngine.cs` — 防崩溃持久化状态机与异常恢复引擎。
- `OutlookSourceReader.cs` — 垃圾箱源邮件的狭窄只读与 Copy 抽象。
- `OwnedCopyLocator.cs` — 确权副本校验与安全重新定位器。
- `ArchiveWriter.cs` — 仅针对插件专属托管副本的唯一变更写入层与 Duplicate Trash 维护。
- `MapiIdentity.cs` — MAPI `PR_SEARCH_KEY` 与 `PR_RECORD_KEY` 读取支持。
- `SqliteStateStore.cs` — 单连接 WAL 事务日志持久化存储与架构升级。
- `Models.cs` — 状态枚举、数据契约与描述符定义。
- `ComUtil.cs`, `Logger.cs` — 日志记录与 COM 对象生命周期释放辅助。
- `install.bat`, `uninstall.bat` — 绿色便携版一键安装与卸载脚本。
- `install-cert.bat` — ClickOnce 自签名证书一键导入信任脚本。
