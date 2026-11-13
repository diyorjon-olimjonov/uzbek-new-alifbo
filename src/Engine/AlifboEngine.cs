// Alifbo - o'zbek matnini eski lotin va kirill alifbosidan yangi lotin alifbosiga
// (2026) o'girish dvigateli. https://github.com/azakapro/alifbo kutubxonasidagi
// fromCyrillic / toNewLatin qoidalarining C# nusxasi.
//
// Natija har bir manba belgisi uchun alohida "bo'lak" (piece) sifatida qaytariladi:
// Word qismi faqat o'zgargan belgilarni almashtiradi, shuning uchun har bir harfning
// formati saqlanadi. Word'ga bog'liq emas, .NET Framework 4.0 va .NET 8 da kompilyatsiya bo'ladi.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Alifbo
{
    public sealed class AlifboEngine
    {
        // Kichik yordamchi bufer: belgi kodlari + har birining manbadagi o'rni.
        private sealed class Buf
        {
            public char[] C;
            public int[] O;
            public int N;

            public Buf(int capacity)
            {
                C = new char[Math.Max(capacity, 4)];
                O = new int[C.Length];
            }

            public void Push(char c, int origin)
            {
                if (N == C.Length)
                {
                    Array.Resize(ref C, C.Length * 2);
                    Array.Resize(ref O, C.Length);
                }
                C[N] = c;
                O[N] = origin;
                N++;
            }

            public int At(int i)
            {
                return i >= 0 && i < N ? C[i] : -1;
            }
        }

        private const char OkinaMod = 'ʻ'; // o' / g' belgisi
        private const char Tutuq = 'ʼ';    // tutuq belgisi

        private static readonly Dictionary<char, char> LowerMap = new Dictionary<char, char>();
        private static readonly Dictionary<char, char> UpperMap = new Dictionary<char, char>();
        private static readonly Dictionary<char, string> CyrMap = new Dictionary<char, string>();

        private readonly List<KeyValuePair<string, string>> _exceptions = new List<KeyValuePair<string, string>>();
        private readonly List<string> _terms = new List<string>();
        private readonly bool _hasUserExceptions;

        static AlifboEngine()
        {
            for (int i = 0; i < 26; i++) Pair((char)('a' + i), (char)('A' + i));
            Pair('ö', 'Ö'); // o umlaut
            Pair('ğ', 'Ğ'); // g breve
            Pair('ş', 'Ş'); // s cedilla
            Pair('ç', 'Ç'); // c cedilla
            Pair('ș', 'Ș'); // s comma (o'xshash belgi)
            for (int i = 0; i < 32; i++) Pair((char)(0x430 + i), (char)(0x410 + i));
            Pair('ё', 'Ё'); // yo
            Pair('ў', 'Ў'); // o' (kirill)
            Pair('қ', 'Қ'); // q
            Pair('ғ', 'Ғ'); // g'
            Pair('ҳ', 'Ҳ'); // h

            Cyr('а', "a"); Cyr('б', "b"); Cyr('в', "v"); Cyr('г', "g");
            Cyr('д', "d"); Cyr('ж', "j"); Cyr('з', "z"); Cyr('и', "i");
            Cyr('й', "y"); Cyr('к', "k"); Cyr('л', "l"); Cyr('м', "m");
            Cyr('н', "n"); Cyr('о', "o"); Cyr('п', "p"); Cyr('р', "r");
            Cyr('с', "s"); Cyr('т', "t"); Cyr('у', "u"); Cyr('ф', "f");
            Cyr('х', "x"); Cyr('ч', "ç"); Cyr('ш', "ş");
            Cyr('ы', "i"); Cyr('э', "e"); Cyr('ў', "ö");
            Cyr('қ', "q"); Cyr('ғ', "ğ"); Cyr('ҳ', "h");
        }

        private static void Pair(char lower, char upper)
        {
            LowerMap[upper] = lower;
            UpperMap[lower] = upper;
        }

        private static void Cyr(char c, string latin)
        {
            CyrMap[c] = latin;
        }

        /// <param name="protectedTerms">O'girilmaydigan so'zlar (katta-kichik harf farqlanadi).</param>
        /// <param name="exceptions">Qo'shimcha "kalit -> qiymat" istisnolari.</param>
        public AlifboEngine(IEnumerable<string> protectedTerms = null, IDictionary<string, string> exceptions = null)
        {
            const string a = "ʼ";
            // Kutubxonadagi istisnolar (data/exceptions.json)
            AddException("ashob", "as" + a + "hob");
            AddException("Ashob", "As" + a + "hob");
            AddException("ASHOB", "AS" + a + "HOB");
            AddException("Ishoq", "Is" + a + "hoq");
            AddException("ishoq", "is" + a + "hoq");
            AddException("ISHOQ", "IS" + a + "HOQ");
            if (exceptions != null)
            {
                foreach (var pair in exceptions)
                {
                    if (string.IsNullOrEmpty(pair.Key)) continue;
                    AddException(PrepareString(pair.Key.Trim()), (pair.Value ?? "").Trim());
                    _hasUserExceptions = true;
                }
            }
            // Uzunroq kalitlar birinchi (barqaror saralash)
            var sorted = new List<KeyValuePair<string, string>>(_exceptions);
            _exceptions.Clear();
            for (int len = MaxKeyLength(sorted); len > 0; len--)
                foreach (var e in sorted)
                    if (e.Key.Length == len) _exceptions.Add(e);

            // Kutubxonadagi himoyalangan so'zlar (data/protected-terms.json)
            _terms.Add("Shakespeare");
            _terms.Add("Chelsea");
            _terms.Add("Photoshop");
            if (protectedTerms != null)
                foreach (var t in protectedTerms)
                    if (!string.IsNullOrEmpty(t) && t.Trim().Length > 0) _terms.Add(t.Trim());
        }

        private static int MaxKeyLength(List<KeyValuePair<string, string>> list)
        {
            int m = 0;
            foreach (var e in list) m = Math.Max(m, e.Key.Length);
            return m;
        }

        private void AddException(string key, string value)
        {
            if (key.Length == 0) return;
            for (int i = 0; i < _exceptions.Count; i++)
            {
                if (_exceptions[i].Key == key)
                {
                    _exceptions[i] = new KeyValuePair<string, string>(key, value);
                    return;
                }
            }
            _exceptions.Add(new KeyValuePair<string, string>(key, value));
        }

        // ------------------------------------------------------------------ API

        /// <summary>Matnni yangi lotinga o'giradi.</summary>
        public string Convert(string s)
        {
            if (string.IsNullOrEmpty(s)) return s ?? "";
            var pieces = Pieces(s, null);
            var sb = new StringBuilder(s.Length + 8);
            foreach (var p in pieces) sb.Append(p);
            return sb.ToString();
        }

        /// <summary>
        /// Har bir manba belgisi o'rniga yoziladigan matn ("" - o'chiriladi).
        /// mask[j] = true bo'lgan belgilar o'zgarmaydi (Word maydon kodlari va h.k.).
        /// </summary>
        public string[] Pieces(string s, bool[] mask)
        {
            int n = s.Length;
            var pieces = new string[n];
            var m = new bool[n];
            if (mask != null) Array.Copy(mask, m, Math.Min(n, mask.Length));
            ProtectSpans(s, m);
            for (int j = 0; j < n; j++) pieces[j] = m[j] ? s[j].ToString() : "";

            int k = 0;
            while (k < n)
            {
                if (m[k]) { k++; continue; }
                int a = k;
                while (k < n && !m[k]) k++;
                ConvertSegment(s, a, k, pieces);
            }
            return pieces;
        }

        /// <summary>Tezkor oldindan tekshiruv: matnda o'zgarishi mumkin bo'lgan narsa bormi?</summary>
        public bool MightChange(string s)
        {
            if (_hasUserExceptions) return true;
            int prev = -1;
            foreach (char c in s)
            {
                if (c >= 'Ѐ' && c <= 'ӿ') return true;
                if (c == 'ș' || c == 'Ș') return true;
                if (IsApos(c)) return true;
                if ((c == 'h' || c == 'H') && (prev == 's' || prev == 'S' || prev == 'c' || prev == 'C')) return true;
                prev = c;
            }
            return false;
        }

        /// <summary>So'z ichidagi belgi (harf, raqam, diakritika, apostrof)?</summary>
        public static bool IsWordish(char c)
        {
            return IsWord(c) || IsApos(c);
        }

        // ---------------------------------------------------------- belgi turlari

        private static char Low(char c)
        {
            char r;
            return LowerMap.TryGetValue(c, out r) ? r : c;
        }

        private static int Low(int c)
        {
            return c < 0 ? c : Low((char)c);
        }

        private static char Up(char c)
        {
            char r;
            return UpperMap.TryGetValue(c, out r) ? r : c;
        }

        private static bool IsUp(int c)
        {
            return c >= 0 && LowerMap.ContainsKey((char)c);
        }

        private static bool IsApos(char c)
        {
            switch (c)
            {
                case 'ʻ': case 'ʼ': case '\'': case '‘': case '’':
                case '`': case '´': case '′': case '＇':
                    return true;
            }
            return false;
        }

        private static bool IsLetter(int c)
        {
            if (c < 0) return false;
            switch (CharUnicodeInfo.GetUnicodeCategory((char)c))
            {
                case UnicodeCategory.UppercaseLetter:
                case UnicodeCategory.LowercaseLetter:
                case UnicodeCategory.TitlecaseLetter:
                case UnicodeCategory.ModifierLetter:
                case UnicodeCategory.OtherLetter:
                    return true;
            }
            return false;
        }

        private static bool IsNumber(int c)
        {
            if (c < 0) return false;
            switch (CharUnicodeInfo.GetUnicodeCategory((char)c))
            {
                case UnicodeCategory.DecimalDigitNumber:
                case UnicodeCategory.LetterNumber:
                case UnicodeCategory.OtherNumber:
                    return true;
            }
            return false;
        }

        private static bool IsMark(int c)
        {
            if (c < 0) return false;
            switch (CharUnicodeInfo.GetUnicodeCategory((char)c))
            {
                case UnicodeCategory.NonSpacingMark:
                case UnicodeCategory.SpacingCombiningMark:
                case UnicodeCategory.EnclosingMark:
                    return true;
            }
            return false;
        }

        private static bool IsWord(int c)
        {
            return IsLetter(c) || IsNumber(c) || IsMark(c);
        }

        // JavaScript \s
        private static bool IsSpaceJs(char c)
        {
            if (c >= '\u0009' && c <= '\u000D') return true;
            if (c >= ' ' && c <= ' ') return true;
            switch (c)
            {
                case ' ': case ' ': case ' ': case ' ': case ' ':
                case ' ': case ' ': case '　': case '﻿':
                    return true;
            }
            return false;
        }

        private static bool IsCyrVowel(int c)
        {
            switch (c)
            {
                case 0x430: case 0x435: case 0x451: case 0x438: case 0x43E:
                case 0x443: case 0x45E: case 0x44D: case 0x44E: case 0x44F:
                    return true;
            }
            return false;
        }

        private static int LetterAt(Buf b, int i)
        {
            if (i < 0 || i >= b.N) return -1;
            return IsLetter(b.C[i]) ? Low(b.C[i]) : -1;
        }

        // ------------------------------------------- himoyalangan qismlar

        private static bool IsLocalPart(char c)
        {
            return c == '.' || c == '_' || c == '%' || c == '+' || c == '-' || IsLetter(c) || IsNumber(c);
        }

        private static bool IsDomainPart(char c)
        {
            return c == '.' || c == '-' || IsLetter(c) || IsNumber(c);
        }

        private static void Mark(bool[] mask, int a, int b)
        {
            for (int k = a; k < b; k++) mask[k] = true;
        }

        private static bool StartsAt(string s, int i, string t)
        {
            return i + t.Length <= s.Length && string.CompareOrdinal(s, i, t, 0, t.Length) == 0;
        }

        private void ProtectSpans(string s, bool[] mask)
        {
            int n = s.Length;

            // `kod`
            int i = 0;
            while (i < n)
            {
                if (s[i] == '`')
                {
                    int j = i + 1;
                    bool found = false;
                    for (; j < n; j++)
                    {
                        if (s[j] == '`') { found = true; break; }
                        if (s[j] == '\n' || s[j] == '\u000B' || s[j] == '\r') break;
                    }
                    if (found) { Mark(mask, i, j + 1); i = j + 1; }
                    else i++;
                }
                else i++;
            }

            // e-pochta: [\p{L}\p{N}._%+-]+@[\p{L}\p{N}.-]+\.[\p{L}]{2,}
            int lastEnd = 0;
            for (int a = 0; a < n; a++)
            {
                if (s[a] != '@' || a < lastEnd) continue;
                int s0 = a;
                while (s0 > lastEnd && IsLocalPart(s[s0 - 1])) s0--;
                if (s0 >= a) continue;
                int e = a + 1;
                while (e < n && IsDomainPart(s[e])) e++;
                int q;
                bool ok = false;
                for (q = e - 3; q >= a + 2; q--)
                {
                    if (s[q] == '.' && IsLetter(s[q + 1]) && IsLetter(s[q + 2])) { ok = true; break; }
                }
                if (!ok) continue;
                int p = q + 1;
                while (p < n && IsLetter(s[p])) p++;
                Mark(mask, s0, p);
                lastEnd = p;
            }

            // URL: (?:https?://|www\.)[^\s`]+
            i = 0;
            while (i < n)
            {
                int len = 0;
                if (StartsAt(s, i, "https://")) len = 8;
                else if (StartsAt(s, i, "http://")) len = 7;
                else if (StartsAt(s, i, "www.")) len = 4;
                if (len > 0)
                {
                    int j = i + len;
                    while (j < n && !IsSpaceJs(s[j]) && s[j] != '`') j++;
                    if (j > i + len) { Mark(mask, i, j); i = j; continue; }
                }
                i++;
            }

            // Himoyalangan so'zlar
            foreach (var t in _terms)
            {
                int p = s.IndexOf(t, StringComparison.Ordinal);
                while (p >= 0)
                {
                    Mark(mask, p, p + t.Length);
                    p = s.IndexOf(t, p + t.Length, StringComparison.Ordinal);
                }
            }
        }

        // --------------------------------------------------- o'girish bosqichlari

        private void ConvertSegment(string s, int a, int b, string[] pieces)
        {
            var c1 = new Buf(b - a);
            for (int j = a; j < b; j++) c1.Push(s[j], j);
            // fromCyrillic: tayyorlash -> istisnolar -> kirill harflari
            Prepare(c1);
            var c2 = ApplyExceptions(c1);
            var c3 = CyrillicStage(c2);
            // ... so'ng eski lotin -> yangi lotin
            Prepare(c3);
            var c4 = ApplyExceptions(c3);
            var c5 = MergeStage(c4);

            var sb = new StringBuilder[b - a];
            for (int j = 0; j < c5.N; j++)
            {
                int o = c5.O[j] - a;
                if (sb[o] == null) sb[o] = new StringBuilder(2);
                sb[o].Append(c5.C[j]);
            }
            for (int j = 0; j < sb.Length; j++) pieces[a + j] = sb[j] == null ? "" : sb[j].ToString();
        }

        // s-vergul -> s-sedil; apostrofga o'xshash belgilar -> U+02BB (o/g dan keyin) yoki U+02BC.
        private static void Prepare(Buf b)
        {
            int prev = -1;
            for (int i = 0; i < b.N; i++)
            {
                char raw = b.C[i];
                if (raw == 'ș') raw = 'ş';
                if (raw == 'Ș') raw = 'Ş';
                if (IsApos(raw))
                    b.C[i] = (prev == 'o' || prev == 'O' || prev == 'g' || prev == 'G') ? OkinaMod : Tutuq;
                else
                    b.C[i] = raw;
                prev = raw;
            }
        }

        private static bool MatchAt(Buf b, int i, string key)
        {
            if (i + key.Length > b.N) return false;
            for (int k = 0; k < key.Length; k++)
                if (b.C[i + k] != key[k]) return false;
            return true;
        }

        private Buf ApplyExceptions(Buf c)
        {
            var d = new Buf(c.N + 16);
            int i = 0;
            while (i < c.N)
            {
                int hit = -1;
                if (_exceptions.Count > 0 && (i == 0 || !IsWord(c.C[i - 1])))
                {
                    for (int k = 0; k < _exceptions.Count; k++)
                    {
                        string key = _exceptions[k].Key;
                        if (!MatchAt(c, i, key)) continue;
                        if (i + key.Length >= c.N || !IsWord(c.C[i + key.Length])) { hit = k; break; }
                    }
                }
                if (hit >= 0)
                {
                    EmitException(c, i, _exceptions[hit].Key, _exceptions[hit].Value, d);
                    i += _exceptions[hit].Key.Length;
                }
                else
                {
                    d.Push(c.C[i], c.O[i]);
                    i++;
                }
            }
            return d;
        }

        // Almashtirish natijasini kalit belgilariga moslab joylaydi (format uchun).
        private static void EmitException(Buf c, int i, string key, string value, Buf d)
        {
            int kl = key.Length, vl = value.Length, mx = Math.Min(kl, vl);
            int pl = 0, sl = 0;
            while (pl < mx && key[pl] == value[pl]) pl++;
            while (sl < mx - pl && key[kl - 1 - sl] == value[vl - 1 - sl]) sl++;
            for (int t = 0; t < vl; t++)
            {
                int ko;
                if (t < pl) ko = t;
                else if (t >= vl - sl) ko = kl - (vl - t);
                else if (pl < kl - sl) ko = pl;
                else if (pl > 0) ko = pl - 1;
                else ko = 0;
                d.Push(value[t], c.O[i + ko]);
            }
        }

        // Ko'p harfli almashtirishni manba harfining katta-kichikligiga moslab yozadi.
        private static void EmitCased(Buf c, int i, string lower, Buf d)
        {
            int orig = c.O[i];
            if (!IsUp(c.C[i]))
            {
                foreach (char ch in lower) d.Push(ch, orig);
                return;
            }
            bool inUpper = IsUp(c.At(i + 1)) || (i > 0 && IsUp(c.C[i - 1]));
            for (int k = 0; k < lower.Length; k++)
                d.Push(inUpper || k == 0 ? Up(lower[k]) : lower[k], orig);
        }

        private static Buf CyrillicStage(Buf c)
        {
            var d = new Buf(2 * c.N + 16);
            int prevL = -1;
            for (int i = 0; i < c.N; i++)
            {
                int lo = Low(c.C[i]);
                switch (lo)
                {
                    case 0x435: // e
                        EmitCased(c, i, prevL == -1 || IsCyrVowel(prevL) || prevL == 0x44A || prevL == 0x44C ? "ye" : "e", d);
                        break;
                    case 0x446: // ts
                        EmitCased(c, i, prevL == -1 || prevL == 0x446 ? "s" : "ts", d);
                        break;
                    case 0x449: // shch
                        EmitCased(c, i, "şç", d);
                        break;
                    case 0x44A: // qattiq belgi
                        EmitCased(c, i, "ʼ", d);
                        break;
                    case 0x44C: // yumshoq belgi
                        switch (LetterAt(c, i + 1))
                        {
                            case 0x430: case 0x43E: case 0x443: case 0x44D: case 0x45E:
                                EmitCased(c, i, "y", d);
                                break;
                        }
                        break;
                    case 0x451:
                        EmitCased(c, i, "yo", d);
                        break;
                    case 0x44E:
                        EmitCased(c, i, "yu", d);
                        break;
                    case 0x44F:
                        EmitCased(c, i, "ya", d);
                        break;
                    default:
                        string rep;
                        if (CyrMap.TryGetValue((char)lo, out rep)) EmitCased(c, i, rep, d);
                        else d.Push(c.C[i], c.O[i]);
                        break;
                }
                prevL = LetterAt(c, i);
            }
            return d;
        }

        // Eski lotin -> yangi lotin: o' g' sh ch
        private static Buf MergeStage(Buf c)
        {
            var d = new Buf(c.N + 16);
            int i = 0;
            while (i < c.N)
            {
                char ch = c.C[i];
                int lo = Low(ch);
                int nx = c.At(i + 1);
                if ((lo == 'o' || lo == 'g') && nx == OkinaMod)
                {
                    char r = lo == 'o' ? 'ö' : 'ğ';
                    d.Push(IsUp(ch) ? Up(r) : r, c.O[i]);
                    i += 2;
                }
                else if (lo == 's' && nx == Tutuq && Low(c.At(i + 2)) == 'h')
                {
                    d.Push(c.C[i], c.O[i]);
                    d.Push(c.C[i + 1], c.O[i + 1]);
                    d.Push(c.C[i + 2], c.O[i + 2]);
                    i += 3;
                }
                else if (lo == 's' && Low(nx) == 'h')
                {
                    d.Push(IsUp(ch) ? 'Ş' : 'ş', c.O[i]);
                    i += 2;
                }
                else if (lo == 'c' && Low(nx) == 'h')
                {
                    d.Push(IsUp(ch) ? 'Ç' : 'ç', c.O[i]);
                    i += 2;
                }
                else if (lo == 'n' && Low(nx) == 'g')
                {
                    // ng - o'zgarmaydigan harf birikmasi
                    d.Push(c.C[i], c.O[i]);
                    d.Push(c.C[i + 1], c.O[i + 1]);
                    i += 2;
                }
                else
                {
                    d.Push(ch, c.O[i]);
                    i++;
                }
            }
            return d;
        }

        private static string PrepareString(string s)
        {
            var b = new Buf(s.Length);
            for (int i = 0; i < s.Length; i++) b.Push(s[i], i);
            Prepare(b);
            return new string(b.C, 0, b.N);
        }
    }
}
