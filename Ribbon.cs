using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace OutlookJunkRescuer
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public Ribbon()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            // Only inject custom UI for the main Outlook Explorer window
            if (ribbonID == "Microsoft.Outlook.Explorer")
            {
                return GetRibbonXml();
            }

            return null;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void OnShowStatusClick(Office.IRibbonControl control)
        {
            try
            {
                using (var form = new StatusForm())
                {
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "无法打开状态窗口: " + ex.Message,
                    "Outlook Junk Rescuer",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public void OnCleanDuplicatesClick(Office.IRibbonControl control)
        {
            var instance = ThisAddIn.Instance;
            if (instance == null || instance.Application == null)
                return;

            try
            {
                using (var form = new DuplicateCleanupForm(instance.Application.Session))
                {
                    form.ShowDialog();
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

        private static string GetRibbonXml()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""Ribbon_Load"">
  <ribbon>
    <tabs>
      <tab idMso=""TabMail"">
        <group id=""grpJunkRescuer"" label=""Junk Rescuer"">
          <button id=""btnJunkRescuerStatus""
                  label=""运行状态""
                  size=""large""
                  imageMso=""AutoArchiveSettings""
                  screentip=""Outlook Junk Rescuer 运行状态与诊断""
                  supertip=""查看实时垃圾邮件防误判监控状态、本次与累计扫描统计，或手动触发即时归档扫描。""
                  onAction=""OnShowStatusClick"" />
          <button id=""btnJunkRescuerDuplicates""
                  label=""清理重复项""
                  size=""large""
                  imageMso=""CleanUpFolder""
                  screentip=""跨设备重复归档副本清理""
                  supertip=""检测因多台设备同时归档产生的良性重复副本，在坚守 Never-reduce-1->0 铁律下，将多余副本安全隔离至 Duplicate Trash。""
                  onAction=""OnCleanDuplicatesClick"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }
    }
}
