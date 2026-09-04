using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkRescuer
{
    internal static class MapiIdentity
    {
        public const string SearchKeySchema =
            "http://schemas.microsoft.com/mapi/proptag/0x300B0102";

        public const string RecordKeySchema =
            "http://schemas.microsoft.com/mapi/proptag/0x0FF90102";

        public static string GetSearchKeyHex(Outlook.MailItem mail)
        {
            return ReadBinaryAsHex(mail, SearchKeySchema);
        }

        public static string GetRecordKeyHex(Outlook.MailItem mail)
        {
            return ReadBinaryAsHex(mail, RecordKeySchema);
        }

        public static string ToHex(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                return null;

            char[] chars = new char[bytes.Length * 2];
            const string hex = "0123456789ABCDEF";

            for (int i = 0; i < bytes.Length; i++)
            {
                chars[i * 2] = hex[bytes[i] >> 4];
                chars[i * 2 + 1] = hex[bytes[i] & 0x0F];
            }

            return new string(chars);
        }

        private static string ReadBinaryAsHex(
            Outlook.MailItem mail,
            string schema)
        {
            Outlook.PropertyAccessor accessor = null;

            try
            {
                accessor = mail.PropertyAccessor;
                object value = accessor.GetProperty(schema);

                if (!(value is byte[] bytes) || bytes.Length == 0)
                    return null;

                return ToHex(bytes);
            }
            finally
            {
                ComUtil.Release(accessor);
            }
        }
    }
}
