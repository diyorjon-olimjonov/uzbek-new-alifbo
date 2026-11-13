// Foydalanuvchi sozlamalari: HKEY_CURRENT_USER\Software\Alifbo
//   ProtectedTerms (matn)  - o'girilmaydigan so'zlar, ";" bilan: "Microsoft;Chelsea FC"
//   Exceptions     (matn)  - istisnolar "kalit=qiymat;...": "mushaf=mus'haf"
//   SetUzbekLatinLanguage (DWORD, 1 yoki 0) - "O'zbek (kirill)" tilini "O'zbek (lotin)" ga o'tkazish
using System;
using System.Collections.Generic;
using Microsoft.Win32;

namespace Alifbo.Word
{
    internal static class Settings
    {
        private const string KeyPath = @"Software\Alifbo";

        private static object Read(string name)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    return key == null ? null : key.GetValue(name);
                }
            }
            catch
            {
                return null;
            }
        }

        public static IEnumerable<string> ProtectedTerms()
        {
            var raw = Read("ProtectedTerms") as string;
            return string.IsNullOrEmpty(raw) ? new string[0] : raw.Split(';');
        }

        public static IDictionary<string, string> Exceptions()
        {
            var result = new Dictionary<string, string>();
            var raw = Read("Exceptions") as string;
            if (string.IsNullOrEmpty(raw)) return result;
            foreach (var part in raw.Split(';'))
            {
                int eq = part.IndexOf('=');
                if (eq > 0) result[part.Substring(0, eq).Trim()] = part.Substring(eq + 1).Trim();
            }
            return result;
        }

        public static bool SetUzbekLatinLanguage
        {
            get
            {
                var value = Read("SetUzbekLatinLanguage");
                return !(value is int) || (int)value != 0;
            }
        }
    }
}
