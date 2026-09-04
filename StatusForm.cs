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
            this.Size = new Size(580, 480);
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
                Location = new Point(15, 12),
                Size = new Size(535, 105),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            _lblStatus = new Label
            {
                Location = new Point(15, 25),
                Size = new Size(505, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Text = "运行状态: 正在加载..."
            };

            _lblAccounts = new Label
            {
                Location = new Point(15, 50),
                Size = new Size(505, 45),
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
                Location = new Point(15, 125),
                Size = new Size(535, 115),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            _lblLastSweep = new Label
            {
                Location = new Point(15, 25),
                Size = new Size(505, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Text = "上次全量对账: 从未执行"
            };

            _lblSweepStats = new Label
            {
                Location = new Point(15, 50),
                Size = new Size(505, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Text = "本次扫描成果: -"
            };

            _lblTotalStats = new Label
            {
                Location = new Point(15, 75),
                Size = new Size(505, 25),
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
                Location = new Point(15, 248),
                Size = new Size(535, 95),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            _lblDbInfo = new Label
            {
                Location = new Point(15, 25),
                Size = new Size(505, 60),
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Text = "SQLite 数据库: 正在读取..."
            };

            grpDb.Controls.Add(_lblDbInfo);
            this.Controls.Add(grpDb);

            // Action Buttons
            _btnScanNow = new Button
            {
                Text = "立即执行对账扫描 (&S)",
                Location = new Point(15, 360),
                Size = new Size(165, 34),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            _btnScanNow.FlatAppearance.BorderSize = 0;
            _btnScanNow.Click += BtnScanNow_Click;
            this.Controls.Add(_btnScanNow);

            _btnOpenFolder = new Button
            {
                Text = "打开数据与日志目录 (&O)",
                Location = new Point(190, 360),
                Size = new Size(175, 34),
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
                Location = new Point(445, 360),
                Size = new Size(105, 34),
                BackColor = Color.FromArgb(225, 225, 225),
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Cursor = Cursors.Hand
            };
            _btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(_btnClose);
        }

        private void RefreshData()
        {
            var instance = ThisAddIn.Instance;
            if (instance == null)
            {
                _lblStatus.Text = "运行状态: 插件宿主未就绪";
                return;
            }

            var accounts = instance.GetWatchedAccounts();
            if (accounts.Count > 0)
            {
                _lblStatus.Text = $"运行状态: 🟢 正常保护中（已实时监听 {accounts.Count} 个账户）";
                _lblAccounts.Text = "监听账户:\n  " + string.Join(", ", accounts);
            }
            else
            {
                _lblStatus.Text = "运行状态: ⚪ 未发现支持的 Outlook.com / Hotmail / Live 账户";
                _lblAccounts.Text = "监听账户: 无支持的个人邮箱账户";
            }

            var stats = ArchiveEngine.Statistics;
            lock (stats)
            {
                if (stats.LastSweepTime != DateTime.MinValue)
                {
                    _lblLastSweep.Text = $"上次全量对账: {stats.LastSweepTime:yyyy-MM-dd HH:mm:ss}";
                    _lblSweepStats.Text = $"本次扫描成果: 发现 {stats.LastVisibleCount} 封 | 归档 {stats.LastArchivedCount} 封 | 跳过 {stats.LastSkippedCount} 封 | 失败 {stats.LastFailedCount} 封";
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
