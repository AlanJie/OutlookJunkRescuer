using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace OutlookJunkRescuer
{
    internal sealed class StatusForm : Form
    {
        private Label _lblStatus;
        private Label _lblAccounts;
        private Label _lblLastSweep;
        private Label _lblSweepStats;
        private Label _lblTotalStats;
        private Label _lblDbInfo;
        private Button _btnScanNow;
        private Button _btnCleanDuplicates;
        private Button _btnOpenFolder;
        private Button _btnClose;

        public StatusForm()
        {
            InitializeComponent();
            RefreshData();
        }

        private void InitializeComponent()
        {
            this.Text = "Outlook Junk Rescuer — 运行状态与诊断控制台";
            this.Size = new Size(640, 520);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            this.BackColor = Color.FromArgb(248, 249, 250);

            // Group 1: Protection Status
            var grpProtection = new GroupBox
            {
                Text = " 保护与监听状态 (Real-Time Protection) ",
                Location = new Point(16, 12),
                Size = new Size(592, 110),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            _lblStatus = new Label
            {
                Location = new Point(15, 24),
                Size = new Size(560, 22),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Text = "运行状态: 正在加载..."
            };

            _lblAccounts = new Label
            {
                Location = new Point(15, 48),
                Size = new Size(560, 52),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Text = "监听账户: 正在检测..."
            };

            grpProtection.Controls.Add(_lblStatus);
            grpProtection.Controls.Add(_lblAccounts);
            this.Controls.Add(grpProtection);

            // Group 2: Statistics
            var grpStats = new GroupBox
            {
                Text = " 归档与统计数据 (Statistics) ",
                Location = new Point(16, 130),
                Size = new Size(592, 135),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            _lblLastSweep = new Label
            {
                Location = new Point(15, 24),
                Size = new Size(560, 22),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Text = "上次全量对账: 从未执行"
            };

            _lblSweepStats = new Label
            {
                Location = new Point(15, 48),
                Size = new Size(560, 42),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Text = "本次扫描成果: -"
            };

            _lblTotalStats = new Label
            {
                Location = new Point(15, 96),
                Size = new Size(560, 24),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Text = "会话累计归档: 0 封 (含实时拦截: 0 封)"
            };

            grpStats.Controls.Add(_lblLastSweep);
            grpStats.Controls.Add(_lblSweepStats);
            grpStats.Controls.Add(_lblTotalStats);
            this.Controls.Add(grpStats);

            // Group 3: Database & Journal
            var grpDb = new GroupBox
            {
                Text = " 持久化存储与日志 (Storage & Logs) ",
                Location = new Point(16, 273),
                Size = new Size(592, 95),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            _lblDbInfo = new Label
            {
                Location = new Point(15, 24),
                Size = new Size(560, 60),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Text = "SQLite 数据库: 正在读取..."
            };

            grpDb.Controls.Add(_lblDbInfo);
            this.Controls.Add(grpDb);

            // Action Buttons
            _btnScanNow = new Button
            {
                Text = "立即执行对账扫描 (&S)",
                Location = new Point(16, 380),
                Size = new Size(180, 36),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            _btnScanNow.FlatAppearance.BorderSize = 0;
            _btnScanNow.Click += BtnScanNow_Click;
            this.Controls.Add(_btnScanNow);

            _btnCleanDuplicates = new Button
            {
                Text = "清理重复归档副本 (&D)...",
                Location = new Point(206, 380),
                Size = new Size(195, 36),
                BackColor = Color.FromArgb(16, 124, 65),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            _btnCleanDuplicates.FlatAppearance.BorderSize = 0;
            _btnCleanDuplicates.Click += BtnCleanDuplicates_Click;
            this.Controls.Add(_btnCleanDuplicates);

            _btnOpenFolder = new Button
            {
                Text = "打开数据与日志目录 (&O)",
                Location = new Point(16, 426),
                Size = new Size(195, 36),
                BackColor = Color.FromArgb(225, 225, 225),
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Cursor = Cursors.Hand
            };
            _btnOpenFolder.Click += BtnOpenFolder_Click;
            this.Controls.Add(_btnOpenFolder);

            _btnClose = new Button
            {
                Text = "关闭 (&C)",
                Location = new Point(488, 426),
                Size = new Size(120, 36),
                BackColor = Color.FromArgb(225, 225, 225),
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Cursor = Cursors.Hand
            };
            _btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(_btnClose);
        }

        private void BtnCleanDuplicates_Click(object sender, EventArgs e)
        {
            var instance = ThisAddIn.Instance;
            if (instance == null || instance.Application == null)
                return;

            try
            {
                using (var form = new DuplicateCleanupForm(instance.Application.Session))
                {
                    form.ShowDialog(this);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "无法打开重复副本清理窗口: " + ex.Message,
                    "Outlook Junk Rescuer",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void RefreshData()
        {
            var instance = ThisAddIn.Instance;
            if (instance == null)
            {
                _lblStatus.Text = "运行状态: 插件宿主未就绪";
                return;
            }

            var stats = ArchiveEngine.Statistics;
            var accounts = instance.GetWatchedAccounts();
            if (accounts.Count > 0)
            {
                _lblStatus.Text = $"运行状态: 🟢 正常保护中（已实时监听 {accounts.Count} 个账户）";
                _lblAccounts.Text = "监听账户:\n  " + string.Join(", ", accounts);
            }
            else if (stats.LastSweepTime == DateTime.MinValue)
            {
                _lblStatus.Text = "运行状态: ⏳ 启动初始化中（启动 15 秒后将自动挂载监听并执行首次对账）";
                _lblAccounts.Text = "监听账户: 等待 15 秒启动延迟中...";
            }
            else
            {
                _lblStatus.Text = "运行状态: ⚪ 未发现支持的 Outlook.com / Hotmail / Live 账户";
                _lblAccounts.Text = "监听账户: 无支持的个人邮箱账户";
            }

            lock (stats)
            {
                if (stats.LastSweepTime != DateTime.MinValue)
                {
                    string durationStr = stats.LastSweepDurationMs > 0
                        ? $" (耗时 {stats.LastSweepDurationMs / 1000.0:F2} 秒)"
                        : string.Empty;
                    _lblLastSweep.Text = $"上次全量对账: {stats.LastSweepTime:yyyy-MM-dd HH:mm:ss}{durationStr}";
                    _lblSweepStats.Text =
                        $"对账成果: 扫描 {stats.LastVisibleCount} 封 | 归档 {stats.LastArchivedCount} 封 (恢复 {stats.LastRecoveredCount} 封)\n" +
                        $"处理明细: 跳过 {stats.LastSkippedCount} 封 | 存疑 {stats.LastUncertainCount} 封 | 失败 {stats.LastFailedCount} 封";
                }
                else
                {
                    _lblLastSweep.Text = "上次全量对账: 启动延时等待中（或尚未执行）";
                    _lblSweepStats.Text = "本次扫描成果: -";
                }

                _lblTotalStats.Text = $"会话累计归档: {stats.TotalArchivedSession} 封 (含实时拦截: {stats.TotalRealtimeIntercepted} 封)";
            }

            string dataDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookJunkRescuer");

            string dbPath = Path.Combine(dataDir, "state-v3.sqlite");
            if (File.Exists(dbPath))
            {
                var fi = new FileInfo(dbPath);
                _lblDbInfo.Text = $"日志数据库: state-v3.sqlite ({fi.Length / 1024.0:F1} KB)\n存储路径: {dataDir}";
            }
            else
            {
                _lblDbInfo.Text = $"日志数据库: 尚未创建或正在初始化\n存储路径: {dataDir}";
            }
        }

        private void BtnScanNow_Click(object sender, EventArgs e)
        {
            var instance = ThisAddIn.Instance;
            if (instance == null)
                return;

            _btnScanNow.Enabled = false;
            Cursor = Cursors.WaitCursor;

            try
            {
                instance.TriggerManualSweep();
                RefreshData();
                MessageBox.Show(
                    "全量对账扫描已执行完毕！统计数据已实时刷新。",
                    "Outlook Junk Rescuer",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "执行扫描时发生异常: " + ex.Message,
                    "Outlook Junk Rescuer",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
                _btnScanNow.Enabled = true;
            }
        }

        private void BtnOpenFolder_Click(object sender, EventArgs e)
        {
            try
            {
                string dataDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OutlookJunkRescuer");

                if (!Directory.Exists(dataDir))
                {
                    Directory.CreateDirectory(dataDir);
                }

                Process.Start("explorer.exe", dataDir);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "无法打开目录: " + ex.Message,
                    "Outlook Junk Rescuer",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
