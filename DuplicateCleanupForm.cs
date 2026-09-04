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
        private readonly Outlook.NameSpace _session;
        private readonly DuplicateCleaner _cleaner;
        private readonly ArchiveWriter _archiveWriter;

        private Outlook.MAPIFolder _archiveFolder;
        private string _storeId;
        private List<DuplicateGroup> _duplicateGroups;

        private Label _lblHeader;
        private Label _lblSubHeader;
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
            ResolveActiveArchive();
        }

        private void ResolveActiveArchive()
        {
            try
            {
                var accounts = _session.Accounts;
                for (int i = 1; i <= accounts.Count; i++)
                {
                    Outlook.Account account = null;
                    Outlook.Store store = null;
                    try
                    {
                        account = accounts[i];
                        store = account.DeliveryStore;
                        if (store != null)
                        {
                            Outlook.MAPIFolder inbox = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                            if (inbox != null)
                            {
                                try
                                {
                                    _archiveFolder = _archiveWriter.GetOrCreateArchiveFolder(inbox);
                                    _storeId = store.StoreID;
                                    if (_archiveFolder != null)
                                        break;
                                }
                                finally
                                {
                                    ComUtil.Release(inbox);
                                }
                            }
                        }
                    }
                    finally
                    {
                        ComUtil.Release(store);
                        ComUtil.Release(account);
                    }
                }

                if (_archiveFolder == null)
                {
                    _lblStatus.Text = "状态: 未能定位到支持的 Junk Archive 归档目录。";
                    _btnScan.Enabled = false;
                }
                else
                {
                    _lblStatus.Text = $"状态: 就绪（目标目录: {_archiveFolder.FolderPath}）";
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = "状态: 解析归档目录失败: " + ex.Message;
                _btnScan.Enabled = false;
            }
        }

        private void InitializeComponent()
        {
            this.Text = "Outlook Junk Rescuer — 跨设备重复归档副本清理";
            this.Size = new Size(760, 620);
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
                Text = "多设备独立备份可能会产生良性重复副本。本维护工具将在保留 1 份法定有效副本的前提下，安全将多余副本移至隔离目录或废件箱。",
                Location = new Point(16, 40),
                Size = new Size(710, 32),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                ForeColor = Color.FromArgb(100, 100, 100)
            };
            this.Controls.Add(_lblSubHeader);

            // Group: Results Grid
            var grpResults = new GroupBox
            {
                Text = " 重复归档邮件列表 ",
                Location = new Point(16, 78),
                Size = new Size(712, 232),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            _grid = new DataGridView
            {
                Location = new Point(12, 22),
                Size = new Size(688, 198),
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
                Location = new Point(16, 318),
                Size = new Size(712, 48),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            _rdoKeepEarliest = new RadioButton
            {
                Text = "保留最早接收/创建的副本 (推荐)",
                Location = new Point(20, 18),
                Size = new Size(240, 22),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Checked = true
            };
            grpPolicy.Controls.Add(_rdoKeepEarliest);

            _rdoKeepLatest = new RadioButton
            {
                Text = "保留最新接收/创建的副本",
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
                Location = new Point(16, 372),
                Size = new Size(712, 48),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            _rdoDestTrash = new RadioButton
            {
                Text = "移至专用软隔离目录 (Junk Archive\\Duplicate Trash) [推荐]",
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
                Location = new Point(16, 428),
                Size = new Size(710, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Text = "状态: 就绪"
            };
            this.Controls.Add(_lblStatus);

            _lblSummary = new Label
            {
                Location = new Point(16, 450),
                Size = new Size(710, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                Text = "检测结果: 尚未开始扫描"
            };
            this.Controls.Add(_lblSummary);

            _progressBar = new ProgressBar
            {
                Location = new Point(16, 474),
                Size = new Size(712, 16),
                Visible = false
            };
            this.Controls.Add(_progressBar);

            // Buttons
            _btnScan = new Button
            {
                Text = "扫描重复项 (&S)",
                Location = new Point(16, 498),
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
                Location = new Point(148, 498),
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
                Location = new Point(280, 498),
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
                Location = new Point(452, 498),
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
                Location = new Point(628, 498),
                Size = new Size(100, 34),
                BackColor = Color.FromArgb(225, 225, 225),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Cursor = Cursors.Hand
            };
            _btnClose.Click += (s, e) => this.Close();
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

            // 如果已有 Duplicate Trash 并且用户选择移至废件箱，询问是否删除已有的 Duplicate Trash
            if (moveToDeleted)
            {
                Outlook.MAPIFolder existingTrash = _archiveWriter.FindDuplicateTrashFolder(_archiveFolder);
                if (existingTrash != null)
                {
                    ComUtil.Release(existingTrash);

                    var askDelete = MessageBox.Show(
                        "检测到当前归档目录下已存在专用的「Duplicate Trash」软隔离目录。\n\n" +
                        "由于您本次选择将多余副本移至系统「已删除邮件」废件箱，是否需要同时删除已有的「Duplicate Trash」目录？\n\n" +
                        "• 点击「是(Yes)」：删除已有的 Duplicate Trash 目录，并继续执行清理\n" +
                        "• 点击「否(No)」：保留该目录，继续执行清理\n" +
                        "• 点击「取消(Cancel)」：中止本次清理操作",
                        "提示：已存在 Duplicate Trash 目录",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question);

                    if (askDelete == DialogResult.Cancel)
                        return;

                    if (askDelete == DialogResult.Yes)
                    {
                        try
                        {
                            _archiveWriter.DeleteDuplicateTrashFolder(_archiveFolder);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                "删除 Duplicate Trash 目录失败: " + ex.Message,
                                "Outlook Junk Rescuer",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                        }
                    }
                }
            }

            int totalRedundant = _duplicateGroups.Sum(g => g.RedundantCopies);
            var confirm = MessageBox.Show(
                $"即将对 {_duplicateGroups.Count} 组重复邮件执行清理。\n\n" +
                $"• 策略: {(_rdoKeepEarliest.Checked ? "保留最早创建/接收的副本" : "保留最新创建/接收的副本")}\n" +
                $"• 目标: 移动至 {destDesc}\n" +
                $"• 预计将 {totalRedundant} 份多余副本安全移出。\n" +
                $"• 系统坚守 Never-reduce-1->0 铁律，确保每组在 Junk Archive 均保留 1 份法定副本。\n\n" +
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
                            _progressBar.Value = current;
                        Application.DoEvents();
                    });

                MessageBox.Show(
                    $"重复副本清理完成！\n\n" +
                    $"• 清理目标: {destDesc}\n" +
                    $"• 成功移动: {result.MovedToTrash} 封\n" +
                    $"• 跳过 (无多余副本或已留存唯一副本): {result.Skipped} 封\n" +
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
                        "Duplicate Trash 目录当前不存在（尚未移动过多余副本或已被移除）。",
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
                    "清空操作将永久删除这些被隔离的冗余副本，此操作不可撤销！\n\n" +
                    "确定要永久清空吗？",
                    "高危操作确认",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (confirm != DialogResult.Yes)
                    return;

                SetBusyState(true, "正在永久清空 Duplicate Trash 目录...");

                int deleted = 0;
                while (items.Count > 0)
                {
                    object item = null;
                    try
                    {
                        item = items[1];
                        if (item is Outlook.MailItem mail)
                        {
                            mail.Delete();
                            deleted++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    finally
                    {
                        ComUtil.Release(item);
                    }
                }

                MessageBox.Show(
                    $"已成功清空 Duplicate Trash 中的 {deleted} 封冗余邮件。",
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
