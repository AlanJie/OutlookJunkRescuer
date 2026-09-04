using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkRescuer
{
    /// <summary>
    /// Listens for real-time ItemAdd events in an account's Junk Email folder.
    /// Retains strong COM references to prevent .NET garbage collection from releasing event sinks.
    /// </summary>
    internal sealed class JunkFolderWatcher : IDisposable
    {
        private readonly string _accountSmtp;
        private readonly string _storeId;
        private readonly ArchiveEngine _engine;
        private Outlook.MAPIFolder _junkFolder;
        private Outlook.Items _junkItems;

        public string AccountSmtp => _accountSmtp;
        public string StoreId => _storeId;

        public JunkFolderWatcher(
            string accountSmtp,
            string storeId,
            Outlook.MAPIFolder junkFolder,
            ArchiveEngine engine)
        {
            _accountSmtp = accountSmtp;
            _storeId = storeId;
            _junkFolder = junkFolder;
            _engine = engine;

            // Retain strong reference to Items to keep event handler alive against GC
            _junkItems = junkFolder.Items;
            _junkItems.ItemAdd += JunkItems_ItemAdd;

            Logger.Write($"[{_accountSmtp}] Real-time Junk folder watcher attached.");
        }

        private void JunkItems_ItemAdd(object item)
        {
            try
            {
                if (item == null)
                    return;

                Logger.Write($"[{_accountSmtp}] New item detected in Junk folder; triggering real-time rescue.");
                _engine.ProcessSingleItem(item, _accountSmtp, _storeId, _junkFolder);
            }
            catch (Exception ex)
            {
                Logger.Write($"[{_accountSmtp}] Error processing ItemAdd event: {ex}");
            }
            finally
            {
                ComUtil.Release(item);
            }
        }

        public void Dispose()
        {
            if (_junkItems != null)
            {
                try
                {
                    _junkItems.ItemAdd -= JunkItems_ItemAdd;
                }
                catch
                {
                }

                ComUtil.Release(_junkItems);
                _junkItems = null;
            }

            if (_junkFolder != null)
            {
                ComUtil.Release(_junkFolder);
                _junkFolder = null;
            }
        }
    }
}
