using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkRescuer
{
    internal sealed class DuplicateCleanupForm : Form
    {
        private sealed class AccountItem
        {
            public string SmtpAddress { get; set; }
            public string StoreId { get; set; }
            public string DisplayName { get; set; }
            public override string ToString() => DisplayName;
        }

        private readonly Outlook.NameSpace _session;
        private readonly DuplicateCleaner _cleaner;
        private readonly ArchiveWriter _archiveWriter;

        private Outlook.MAPIFolder _archiveFolder;
        private string _storeId;
        private List<DuplicateGroup> _duplicateGroups;

        private Label _lblHeader;
        private Label _lblSubHeader;
        private Label _lblAccount;
        private ComboBox _cmbAccounts;
        private Label _lblStatus;
        private Label _lblSummary;
        private DataGridView _grid;
        private RadioButton _rdoKeepEarliest;
        private RadioButton _rdoKeepLatest;
        private RadioButton _rdoDestTrash;
        private RadioButton _rdoDestDeleted;
        private ProgressBar _progressBar;
        private Button _btnScan;
        private Button _btnClean;
        private Button _btnOpenTrash;
        private Button _btnEmptyTrash;
        private Button _btnClose;

        public DuplicateCleanupForm(Outlook.NameSpace session)
        {
            _session = session ?? throw new ArgumentNullException(nameof(session));
            _cleaner = new DuplicateCleaner(session);
            _archiveWriter = new ArchiveWriter();

            InitializeComponent();
            PopulateAccounts();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                ComUtil.Release(_archiveFolder);
                _archiveFolder = null;
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.Text = "Outlook Junk Rescuer — 跨设备重复归档副本清理";
            this.Size = new Size(760, 640);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            this.BackColor = Color.FromArgb(248, 249, 250);

            // Header
            _lblHeader = new Label
            {
                Text = "跨设备归档副本检测与清理",
                Location = new Point(16, 14),
                Size = new Size(710, 24),
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = Color.FromArgb(32, 32, 32)
            };
            this.Controls.Add(_lblHeader);

            _lblSubHeader = new Label
            {
                Text = "多设备独立备份可能会产生良性重复副本。本工具遵循 Never-reduce-1->0 铁律，在保留 1 份法定有效副本的前提下，安全将多余副本移至隔离目录或废件箱。",
                Location = new Point(16, 40),
                Size = new Size(710, 32),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                ForeColor = Color.FromArgb(100, 100, 100)
            };
            this.Controls.Add(_lblSubHeader);

            // Account selection
            _lblAccount = new Label
            {
                Text = "选择账户:",
                Location = new Point(16, 78),
                Size = new Size(70, 22),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            this.Controls.Add(_lblAccount);

            _cmbAccounts = new ComboBox
            {
                Location = new Point(90, 77),
                Size = new Size(320, 24),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            _cmbAccounts.SelectedIndexChanged += CmbAccounts_SelectedIndexChanged;
            this.Controls.Add(_cmbAccounts);

            // Group: Results Grid
            var grpResults = new GroupBox
            {
                Text = " 重复归档邮件列表 ",
                Location = new Point(16, 110),
                Size = new Size(712, 210),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            _grid = new DataGridView
            {
                Location = new Point(12, 22),
                Size = new Size(688, 176),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.Fixed3D,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            _grid.Columns.Add("Index", "#");
            _grid.Columns["Index"].Width = 40;
            _grid.Columns.Add("Subject", "邮件主题");
            _grid.Columns["Subject"].FillWeight = 200;
            _grid.Columns.Add("Copies", "副本总数");
            _grid.Columns["Copies"].Width = 80;
            _grid.Columns.Add("Redundant", "冗余副本");
            _grid.Columns["Redundant"].Width = 80;
            _grid.Columns.Add("Replicas", "副本来源设备");
            _grid.Columns["Replicas"].Width = 140;
            _grid.Columns.Add("LatestDate", "接收时间");
            _grid.Columns["LatestDate"].Width = 130;

            grpResults.Controls.Add(_grid);
            this.Controls.Add(grpResults);

            // Policy panel
            var grpPolicy = new GroupBox
            {
                Text = " 保留策略 ",
                Location = new Point(16, 328),
                Size = new Size(712, 48),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            _rdoKeepEarliest = new RadioButton
            {
                Text = "保留最早创建的副本 (推荐)",
                Location = new Point(20, 18),
                Size = new Size(240, 22),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Checked = true
            };
            grpPolicy.Controls.Add(_rdoKeepEarliest);

            _rdoKeepLatest = new RadioButton
            {
                Text = "保留最新创建的副本",
                Location = new Point(280, 18),
                Size = new Size(220, 22),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            grpPolicy.Controls.Add(_rdoKeepLatest);
            this.Controls.Add(grpPolicy);

            // Destination panel
            var grpDestination = new GroupBox
            {
                Text = " 多余副本清理去向 ",
                Location = new Point(16, 382),
                Size = new Size(712, 48),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            _rdoDestTrash = new RadioButton
            {
                Text = "移至专用隔离目录 (Junk Archive\\Duplicate Trash) [推荐]",
                Location = new Point(20, 18),
                Size = new Size(360, 22),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Checked = true
            };
            grpDestination.Controls.Add(_rdoDestTrash);

            _rdoDestDeleted = new RadioButton
            {
                Text = "移至系统「已删除邮件」废件箱 (Deleted Items)",
                Location = new Point(390, 18),
                Size = new Size(300, 22),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            grpDestination.Controls.Add(_rdoDestDeleted);
            this.Controls.Add(grpDestination);

            // Status info
            _lblStatus = new Label
            {
                Location = new Point(16, 436),
                Size = new Size(710, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Text = "状态: 正在初始化..."
            };
            this.Controls.Add(_lblStatus);

            _lblSummary = new Label
            {
                Location = new Point(16, 458),
                Size = new Size(710, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                Text = "检测结果: 尚未开始扫描"
            };
            this.Controls.Add(_lblSummary);

            _progressBar = new ProgressBar
            {
                Location = new Point(16, 482),
                Size = new Size(712, 16),
                Visible = false
            };
            this.Controls.Add(_progressBar);

            // Buttons
            _btnScan = new Button
            {
                Text = "扫描重复项 (&S)",
                Location = new Point(16, 508),
                Size = new Size(125, 34),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            _btnScan.FlatAppearance.BorderSize = 0;
            _btnScan.Click += BtnScan_Click;
            this.Controls.Add(_btnScan);

            _btnClean = new Button
            {
                Text = "执行清理 (&C)",
                Location = new Point(148, 508),
                Size = new Size(125, 34),
                BackColor = Color.FromArgb(216, 59, 1),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Enabled = false,
                Cursor = Cursors.Hand
            };
            _btnClean.FlatAppearance.BorderSize = 0;
            _btnClean.Click += BtnClean_Click;
            this.Controls.Add(_btnClean);

            _btnOpenTrash = new Button
            {
                Text = "打开 Duplicate Trash (&T)",
                Location = new Point(280, 508),
                Size = new Size(165, 34),
                BackColor = Color.FromArgb(225, 225, 225),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Cursor = Cursors.Hand
            };
            _btnOpenTrash.Click += BtnOpenTrash_Click;
            this.Controls.Add(_btnOpenTrash);

            _btnEmptyTrash = new Button
            {
                Text = "清空 Trash (&E)",
                Location = new Point(452, 508),
                Size = new Size(120, 34),
                BackColor = Color.FromArgb(225, 225, 225),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Cursor = Cursors.Hand
            };
            _btnEmptyTrash.Click += BtnEmptyTrash_Click;
            this.Controls.Add(_btnEmptyTrash);

            _btnClose = new Button
            {
                Text = "关闭 (&X)",
                Location = new Point(628, 508),
                Size = new Size(100, 34),
                BackColor = Color.FromArgb(225, 225, 225),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Cursor = Cursors.Hand
            };
            _btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(_btnClose);
        }

        private void PopulateAccounts()
        {
            _cmbAccounts.Items.Clear();

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
                        string smtp = ArchiveEngine.TryGetSmtpAddress(account);
                        if (!string.IsNullOrEmpty(smtp) && ArchiveEngine.IsSupportedAddress(smtp))
                        {
                            store = account.DeliveryStore;
                            if (store != null)
                            {
                                _cmbAccounts.Items.Add(new AccountItem
                                {
                                    SmtpAddress = smtp,
                                    StoreId = store.StoreID,
                                    DisplayName = $"{account.DisplayName} ({smtp})"
                                });
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Write($"[DuplicateCleanupForm] 读取账户信息失败: {ex.Message}");
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

            if (_cmbAccounts.Items.Count > 0)
            {
                _cmbAccounts.SelectedIndex = 0;
            }
            else
            {
                _lblStatus.Text = "状态: 未检测到支持的 Outlook.com / Hotmail 账户。";
                _btnScan.Enabled = false;
            }
        }

        private void CmbAccounts_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_cmbAccounts.SelectedItem is AccountItem item)
            {
                SelectAccount(item);
            }
        }

        private void SelectAccount(AccountItem item)
        {
            ComUtil.Release(_archiveFolder);
            _archiveFolder = null;
            _storeId = item.StoreId;
            _duplicateGroups = null;
            _grid.Rows.Clear();
            _btnClean.Enabled = false;
            _lblSummary.Text = "检测结果: 尚未开始扫描";

            Outlook.Store store = null;
            Outlook.MAPIFolder inbox = null;
            try
            {
                store = _session.GetStoreFromID(_storeId);
                if (store != null)
                {
                    inbox = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    if (inbox != null)
                    {
                        _archiveFolder = _archiveWriter.GetOrCreateArchiveFolder(inbox);
                    }
                }

                if (_archiveFolder == null)
                {
                    _lblStatus.Text = $"状态: 未能在 {item.SmtpAddress} 下定位或创建 Junk Archive 目录。";
                    _btnScan.Enabled = false;
                }
                else
                {
                    _lblStatus.Text = $"状态: 就绪（当前账户: {item.SmtpAddress}）";
                    _btnScan.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = "状态: 获取归档目录异常: " + ex.Message;
                _btnScan.Enabled = false;
            }
            finally
            {
                ComUtil.Release(inbox);
                ComUtil.Release(store);
            }
        }

        private void BtnScan_Click(object sender, EventArgs e)
        {
            if (_archiveFolder == null)
                return;

            SetBusyState(true, "正在轻量扫描 Junk Archive 文件夹...");
            _grid.Rows.Clear();
            _btnClean.Enabled = false;

            try
            {
                int totalScanned = 0;
                _duplicateGroups = _cleaner.ScanDuplicates(_archiveFolder, _storeId, out totalScanned);

                int totalRedundant = _duplicateGroups.Sum(g => g.RedundantCopies);

                for (int i = 0; i < _duplicateGroups.Count; i++)
                {
                    var group = _duplicateGroups[i];
                    var replicas = group.Copies
                        .Select(c => string.IsNullOrEmpty(c.ReplicaId) ? "未知设备" : c.ReplicaId.Substring(0, Math.Min(8, c.ReplicaId.Length)))
                        .Distinct()
                        .ToList();

                    string replicaSummary = string.Join(", ", replicas);
                    DateTime latestDate = group.Copies.Max(c => c.ReceivedTime);
                    string dateStr = latestDate != DateTime.MinValue ? latestDate.ToString("yyyy-MM-dd HH:mm") : "-";

                    _grid.Rows.Add(
                        (i + 1).ToString(),
                        group.Subject,
                        group.TotalCopies.ToString(),
                        group.RedundantCopies.ToString(),
                        replicaSummary,
                        dateStr);
                }

                _lblStatus.Text = $"扫描完成: 遍历 {totalScanned} 封归档邮件，发现 {_duplicateGroups.Count} 组重复。";
                _lblSummary.Text = $"检测结果: 发现 {_duplicateGroups.Count} 组重复邮件，可清理 {totalRedundant} 个多余副本。";

                if (totalRedundant > 0)
                {
                    _btnClean.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "扫描重复副本时发生异常: " + ex.Message,
                    "Outlook Junk Rescuer",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                _lblStatus.Text = "扫描失败: " + ex.Message;
            }
            finally
            {
                SetBusyState(false, null);
            }
        }

        private void BtnClean_Click(object sender, EventArgs e)
        {
            if (_archiveFolder == null || _duplicateGroups == null || _duplicateGroups.Count == 0)
                return;

            bool moveToDeleted = _rdoDestDeleted.Checked;
            var destination = moveToDeleted
                ? DuplicateDestination.DeletedItems
                : DuplicateDestination.DuplicateTrash;

            string destDesc = moveToDeleted
                ? "系统「已删除邮件」废件箱 (Deleted Items)"
                : "「Junk Archive\\Duplicate Trash」软隔离目录";

            int totalRedundant = _duplicateGroups.Sum(g => g.RedundantCopies);
            var confirm = MessageBox.Show(
                $"即将对 {_duplicateGroups.Count} 组重复邮件执行清理。\n\n" +
                $"• 策略: {(_rdoKeepEarliest.Checked ? "保留最早创建的副本" : "保留最新创建的副本")}\n" +
                $"• 目标: 移动至 {destDesc}\n" +
                $"• 预计将 {totalRedundant} 份多余副本安全移出。\n" +
                $"• 系统坚守 Never-reduce-1->0 铁律，确保每组在 Junk Archive 均保留 1 份法定有效副本。\n\n" +
                "是否立即开始清理？",
                "确认执行重复副本清理",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm != DialogResult.Yes)
                return;

            SetBusyState(true, $"正在安全移动冗余副本至 {destDesc}...");
            _progressBar.Visible = true;
            _progressBar.Minimum = 0;
            _progressBar.Maximum = _duplicateGroups.Count;
            _progressBar.Value = 0;

            try
            {
                var policy = _rdoKeepEarliest.Checked
                    ? DuplicateRetentionPolicy.KeepEarliest
                    : DuplicateRetentionPolicy.KeepLatest;

                var result = _cleaner.CleanDuplicates(
                    _archiveFolder,
                    _storeId,
                    _duplicateGroups,
                    policy,
                    destination,
                    (current, total) =>
                    {
                        if (current <= _progressBar.Maximum)
                        {
                            _progressBar.Value = current;
                            _progressBar.Refresh();
                        }
                    });

                MessageBox.Show(
                    $"重复副本清理完成！\n\n" +
                    $"• 清理目标: {destDesc}\n" +
                    $"• 成功移动: {result.MovedToTrash} 封\n" +
                    $"• 跳过 (无多余副本、状态模糊或已留存唯一副本): {result.Skipped} 封\n" +
                    $"• 失败: {result.Failed} 封",
                    "清理完成",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // 重新扫描以刷新列表
                BtnScan_Click(sender, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "执行清理时发生异常: " + ex.Message,
                    "Outlook Junk Rescuer",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                _progressBar.Visible = false;
                SetBusyState(false, null);
            }
        }

        private void BtnOpenTrash_Click(object sender, EventArgs e)
        {
            if (_archiveFolder == null)
                return;

            try
            {
                var trash = _archiveWriter.FindDuplicateTrashFolder(_archiveFolder);
                if (trash == null)
                {
                    MessageBox.Show(
                        "Duplicate Trash 目录当前不存在（尚未移动过多余副本）。",
                        "提示",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                Outlook.Explorer activeExplorer = null;
                try
                {
                    activeExplorer = _session.Application.ActiveExplorer();
                    if (activeExplorer != null)
                    {
                        activeExplorer.CurrentFolder = trash;
                    }
                    else
                    {
                        trash.Display();
                    }
                }
                finally
                {
                    ComUtil.Release(activeExplorer);
                    ComUtil.Release(trash);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "无法打开 Duplicate Trash 目录: " + ex.Message,
                    "Outlook Junk Rescuer",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void BtnEmptyTrash_Click(object sender, EventArgs e)
        {
            if (_archiveFolder == null)
                return;

            Outlook.MAPIFolder trash = null;
            Outlook.Items items = null;

            try
            {
                trash = _archiveWriter.FindDuplicateTrashFolder(_archiveFolder);
                if (trash == null)
                {
                    MessageBox.Show(
                        "Duplicate Trash 目录当前不存在，无需清空。",
                        "清空 Duplicate Trash",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                items = trash.Items;
                int count = items.Count;

                if (count == 0)
                {
                    MessageBox.Show(
                        "Duplicate Trash 目录当前为空，无需清空。",
                        "清空 Duplicate Trash",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                var confirm = MessageBox.Show(
                    $"Duplicate Trash 目录下当前共有 {count} 封邮件。\n\n" +
                    "清空操作将通过所有权正向验证的冗余副本移至系统「已删除邮件」（废件箱）。\n" +
                    "若目录中存在非插件归档的未知项目，系统将予以安全保留，不予触碰。\n\n" +
                    "确定要将隔离副本移至「已删除邮件」吗？",
                    "清空 Duplicate Trash 确认",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (confirm != DialogResult.Yes)
                    return;

                SetBusyState(true, "正在将隔离目录中的冗余副本移至「已删除邮件」...");

                int deleted = 0;
                int skippedUnknown = 0;
                _cleaner.EmptyTrash(trash, out deleted, out skippedUnknown);

                string msg = $"已成功将 {deleted} 封冗余副本移至「已删除邮件」。";
                if (skippedUnknown > 0)
                {
                    msg += $"\n\n（检测到 {skippedUnknown} 个非 OJR 拥有或标记不完整的项目，已安全保留在原处）";
                }

                MessageBox.Show(
                    msg,
                    "清空完成",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "清空 Duplicate Trash 失败: " + ex.Message,
                    "Outlook Junk Rescuer",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                ComUtil.Release(items);
                ComUtil.Release(trash);
                SetBusyState(false, null);
            }
        }

        private void SetBusyState(bool busy, string statusText)
        {
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
            _cmbAccounts.Enabled = !busy;
            _btnScan.Enabled = !busy;
            _btnOpenTrash.Enabled = !busy;
            _btnEmptyTrash.Enabled = !busy;
            _btnClose.Enabled = !busy;

            if (busy)
            {
                _btnClean.Enabled = false;
                if (!string.IsNullOrEmpty(statusText))
                    _lblStatus.Text = "状态: " + statusText;
            }
        }
    }
}
