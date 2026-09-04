using System;
using System.IO;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkRescuer
{
    public partial class ThisAddIn
    {
        private const int StartupDelayMilliseconds = 15000;

        private Timer _startupTimer;
        private SqliteStateStore _stateStore;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
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

            _startupTimer = new Timer
            {
                Interval = StartupDelayMilliseconds
            };

            _startupTimer.Tick += StartupTimer_Tick;
            _startupTimer.Start();

            Logger.Write("Add-in loaded; read-only-source sweep scheduled.");
        }

        private void StartupTimer_Tick(object sender, EventArgs e)
        {
            _startupTimer.Stop();

            Outlook.NameSpace session = null;

            try
            {
                session = Application.Session;
                var engine = new ArchiveEngine(session, _stateStore);
                engine.RunStartupSweep();
            }
            catch (Exception ex)
            {
                Logger.Write("Startup sweep failed: " + ex);
            }
            finally
            {
                session = null;
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

            _stateStore?.Dispose();
            _stateStore = null;
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
