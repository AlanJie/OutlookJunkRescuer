using System;
using System.Runtime.InteropServices;

namespace OutlookJunkRescuer
{
    internal static class ComUtil
    {
        public static void Release(object value)
        {
            if (value == null || !Marshal.IsComObject(value))
                return;

            try
            {
                Marshal.ReleaseComObject(value);
            }
            catch
            {
                // Best effort only. COM cleanup must not take Outlook down.
            }
        }
    }
}
