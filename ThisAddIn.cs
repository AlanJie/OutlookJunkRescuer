using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkRescuer
{
    public partial class ThisAddIn
    {
        private const int StartupDelayMilliseconds = 15000;

        public static ThisAddIn Instance { get; private set; }

        private Timer _startupTimer;
        private SqliteStateStore _stateStore;
        private ArchiveEngine _engine;
        private readonly List<JunkFolderWatcher> _watchers = new List<JunkFolderWatcher>();

        internal SqliteStateStore StateStore => _stateStore;

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Instance = this;

            try
            {
                string statePath = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.LocalApplicationData),
                    "OutlookJunkRescuer",
                    // Keep the v3 filename so existing journals migrate in place.
                    "state-v3.sqlite");

                _stateStore = new SqliteStateStore(statePath);
            }
            catch (Exception ex)
            {
                // State durability is required for safe operation. If SQLite or
                // its native dependency cannot initialize, disable this run rather
                // than operating without a journal or affecting Outlook startup.
                Logger.Write("State store initialization failed; add-in disabled: " + ex);
                return;
            }

            try
            {
                _engine = new ArchiveEngine(Application.Session, _stateStore);
                AttachWatchers();
            }
            catch (Exception ex)
            {
                Logger.Write("Failed to initialize engine or watchers: " + ex);
            }

            _startupTimer = new Timer
            {
                Interval = StartupDelayMilliseconds
            };

            _startupTimer.Tick += StartupTimer_Tick;
            _startupTimer.Start();

            Logger.Write("Add-in loaded; real-time protection active; read-only-source sweep scheduled.");
        }

        private void AttachWatchers()
        {
            Outlook.Accounts accounts = null;

            try
            {
                accounts = Application.Session.Accounts;
                int count = accounts.Count;

                for (int i = 1; i <= count; i++)
                {
                    Outlook.Account account = null;
                    Outlook.Store store = null;
                    Outlook.MAPIFolder junk = null;

                    try
                    {
                        account = accounts[i];
                        string smtp = ArchiveEngine.TryGetSmtpAddress(account);

                        if (!ArchiveEngine.IsSupportedAddress(smtp))
                            continue;

                        store = account.DeliveryStore;
                        if (store == null)
                            continue;

                        string storeId = store.StoreID;
                        if (string.IsNullOrEmpty(storeId))
                            continue;

                        junk = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);
                        if (junk == null)
                            continue;

                        var watcher = new JunkFolderWatcher(
                            smtp.Trim().ToLowerInvariant(),
                            storeId,
                            junk,
                            _engine);

                        _watchers.Add(watcher);
                        junk = null; // ownership transferred to watcher
                    }
                    catch (Exception ex)
                    {
                        Logger.Write($"Account #{i} watcher attach failed: {ex}");
                    }
                    finally
                    {
                        ComUtil.Release(junk);
                        ComUtil.Release(store);
                        ComUtil.Release(account);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Write("AttachWatchers failed: " + ex);
            }
            finally
            {
                ComUtil.Release(accounts);
            }
        }

        public List<string> GetWatchedAccounts()
        {
            var list = new List<string>();
            lock (_watchers)
            {
                foreach (var w in _watchers)
                {
                    list.Add(w.AccountSmtp);
                }
            }
            return list;
        }

        public void TriggerManualSweep()
        {
            if (_engine != null)
            {
                Logger.Write("Manual full sweep triggered by user.");
                _engine.RunStartupSweep();
            }
        }

        private void StartupTimer_Tick(object sender, EventArgs e)
        {
            _startupTimer.Stop();

            try
            {
                _engine?.RunStartupSweep();
            }
            catch (Exception ex)
            {
                Logger.Write("Startup sweep failed: " + ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (_startupTimer != null)
            {
                _startupTimer.Stop();
                _startupTimer.Tick -= StartupTimer_Tick;
                _startupTimer.Dispose();
                _startupTimer = null;
            }

            lock (_watchers)
            {
                foreach (var watcher in _watchers)
                {
                    watcher.Dispose();
                }
                _watchers.Clear();
            }

            _stateStore?.Dispose();
            _stateStore = null;
            _engine = null;
            Instance = null;
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
