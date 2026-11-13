// Word hujjatidagi matnni format saqlagan holda yangi lotinga o'giradi.
//
// Har bir xatboshi matni olinadi, AlifboEngine har bir belgi uchun yangi qiymatni
// hisoblaydi va hujjatda FAQAT o'zgargan belgilar almashtiriladi. Shrift, o'lcham,
// qalin, kursiv, rang, ajratish va boshqa formatlar har bir harfda o'z joyida qoladi.
// Maydon kodlari, URL, e-pochta va `kod` qismlari o'zgarmaydi.
//
// Word bilan kech bog'lanish (dynamic) ishlatiladi - Word PIA kerak emas, shuning
// uchun bitta DLL Word 2010, 2013, 2016, 2019, 2021, 2024 va Microsoft 365 da ishlaydi.
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace Alifbo.Word
{
    internal sealed class WordConverter
    {
        private const int WdUndefined = 9999999;
        private const int WdNoProtection = -1;
        private const int LangUzLatin = 1091;
        private const int LangUzCyrillic = 2115;
        private const int MaxGap = 3; // bitta amalda yoziladigan guruhdagi o'zgarmagan belgilar

        private readonly dynamic _app;
        private readonly AlifboEngine _engine;
        private int _changed;
        private int _skipped;
        private int _paragraphs;

        public WordConverter(object application)
        {
            _app = application;
            _engine = new AlifboEngine(Settings.ProtectedTerms(), Settings.Exceptions());
        }

        public void Run()
        {
            if ((int)_app.Documents.Count == 0)
            {
                MessageBox.Show("Ochiq hujjat yo\u02BBq.", "Alifbo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            dynamic doc = _app.ActiveDocument;
            if ((int)doc.ProtectionType != WdNoProtection)
            {
                MessageBox.Show("Hujjat tahrirdan himoyalangan. Avval himoyani olib tashlang.", "Alifbo",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            dynamic selection = _app.Selection;
            bool selectionOnly = (int)selection.Start != (int)selection.End;
            var timer = Stopwatch.StartNew();

            dynamic undo = null;
            try
            {
                undo = _app.UndoRecord;              // Word 2010+
                undo.StartCustomRecord("Yangi lotinga o\u02BBgirish");
            }
            catch
            {
                undo = null;
            }

            bool screen = true;
            try { screen = (bool)_app.ScreenUpdating; } catch { }
            _app.ScreenUpdating = false;
            try
            {
                if (selectionOnly)
                {
                    dynamic range = selection.Range;
                    ProcessRange(range, range);
                }
                else
                {
                    ProcessAllStories(doc);
                }
            }
            finally
            {
                try { _app.ScreenUpdating = screen; } catch { }
                if (undo != null)
                {
                    try { undo.EndCustomRecord(); } catch { }
                }
            }

            string msg = _changed == 0
                ? "Alifbo: o\u02BBzgartiriladigan matn topilmadi."
                : string.Format("Alifbo: {0} ta belgi yangi lotinga o\u02BBgirildi ({1:0.0} s). Bekor qilish: Ctrl+Z.",
                                _changed, timer.Elapsed.TotalSeconds);
            if (_skipped > 0) msg += " " + _skipped + " ta joy o\u02BBtkazib yuborildi.";
            try { _app.StatusBar = msg; } catch { }
        }

        // Asosiy matn, kolontitullar, izohlar, sharhlar va matn qutilari.
        private void ProcessAllStories(dynamic doc)
        {
            foreach (dynamic first in doc.StoryRanges)
            {
                dynamic story = first;
                while (story != null)
                {
                    ProcessRange(story, null);
                    int type = (int)story.StoryType;
                    if (type >= 6 && type <= 11) ProcessHeaderShapes(story);
                    story = story.NextStoryRange;
                }
            }
        }

        // Kolontitullardagi matn qutilari alohida "story" sifatida ko'rinmaydi.
        private void ProcessHeaderShapes(dynamic story)
        {
            try
            {
                dynamic shapes = story.ShapeRange;
                int count = (int)shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic frame = shapes.Item(i).TextFrame;
                        if ((int)frame.HasText != 0) ProcessRange(frame.TextRange, null);
                    }
                    catch { }
                }
            }
            catch { }
        }

        private void ProcessRange(dynamic range, dynamic limit)
        {
            dynamic para = range.Paragraphs.First;
            while (para != null)
            {
                dynamic pr = para.Range;
                ProcessChunk(pr, limit, true);
                _paragraphs++;
                if (_paragraphs % 50 == 0)
                {
                    try { _app.StatusBar = "Alifbo: " + _paragraphs + " ta xatboshi..."; } catch { }
                }
                if ((int)pr.End >= (int)range.End) break;
                para = para.Next(1);
            }
        }

        // Bitta xatboshi (yoki so'z) ni o'giradi.
        private void ProcessChunk(dynamic source, dynamic limit, bool allowFallback)
        {
            dynamic r = source.Duplicate;
            r.TextRetrievalMode.IncludeFieldCodes = true;
            r.TextRetrievalMode.IncludeHiddenText = true;
            string s = (string)r.Text ?? "";
            int n = s.Length;
            if (n == 0 || !_engine.MightChange(s)) return;
            int start = (int)r.Start, end = (int)r.End;

            // Matn va hujjatdagi o'rinlar mos kelishini tekshiramiz (jadval katagi
            // oxiridagi \r\a bitta o'rin egallaydi).
            if (n != end - start && !(n == end - start + 1 && s.EndsWith("\r\a", StringComparison.Ordinal)))
            {
                if (allowFallback)
                {
                    // So'zma-so'z, oxiridan boshiga (o'zgarishlar oldingi so'zlarni siljitmaydi).
                    dynamic words = r.Words;
                    for (int w = (int)words.Count; w >= 1; w--) ProcessChunk(words.Item(w), limit, false);
                }
                else
                {
                    _skipped++;
                }
                return;
            }

            bool[] mask = FieldMask(s);
            string[] pieces = _engine.Pieces(s, mask);

            // Belgilangan qism chegaralari (so'zlarni bo'lmaslik uchun kengaytiriladi).
            int lo = 0, hi = n;
            if (limit != null)
            {
                lo = Math.Max(0, (int)limit.Start - start);
                hi = Math.Min(n, (int)limit.End - start);
                while (lo > 0 && lo < n && AlifboEngine.IsWordish(s[lo]) && AlifboEngine.IsWordish(s[lo - 1])) lo--;
                while (hi > 0 && hi < n && AlifboEngine.IsWordish(s[hi - 1]) && AlifboEngine.IsWordish(s[hi])) hi++;
                if (lo >= hi) return;
            }

            var changed = new bool[n];
            bool any = false;
            for (int i = lo; i < hi; i++)
            {
                changed[i] = pieces[i].Length != 1 || pieces[i][0] != s[i];
                any |= changed[i];
            }
            if (!any) return;

            // O'zgargan belgilarni guruhlarga birlashtiramiz va oxiridan boshiga qarab
            // qo'llaymiz - shunda oldingi o'rinlar siljimaydi.
            var groups = new System.Collections.Generic.List<int[]>();
            int g = lo;
            while (g < hi)
            {
                if (!changed[g]) { g++; continue; }
                int last = g, k = g + 1;
                while (k < hi)
                {
                    if (changed[k]) { last = k; k++; }
                    else if (k - last > MaxGap || s[k] < ' ') break;
                    else k++;
                }
                groups.Add(new[] { g, last });
                g = last + 1;
            }

            int before = _changed;
            for (int i = groups.Count - 1; i >= 0; i--)
                ApplyGroup(r, start, groups[i][0], groups[i][1], s, pieces);

            if (Settings.SetUzbekLatinLanguage && _changed > before)
            {
                try
                {
                    if ((int)source.LanguageID == LangUzCyrillic) source.LanguageID = LangUzLatin;
                }
                catch { }
            }
        }

        // Maydon kodlari ({ HYPERLINK ... }, { PAGE } va h.k.) o'zgarmaydi.
        private static bool[] FieldMask(string s)
        {
            var mask = new bool[s.Length];
            var inCode = new bool[65];
            int depth = 0, codeDepth = 0;
            for (int i = 0; i < s.Length; i++)
            {
                switch (s[i])
                {
                    case '\u0013': // maydon boshi
                        mask[i] = true;
                        if (depth < 64)
                        {
                            depth++;
                            inCode[depth] = true;
                            codeDepth++;
                        }
                        break;
                    case '\u0014': // maydon natijasi boshlanadi
                        mask[i] = true;
                        if (depth > 0 && inCode[depth])
                        {
                            inCode[depth] = false;
                            codeDepth--;
                        }
                        break;
                    case '\u0015': // maydon oxiri
                        mask[i] = true;
                        if (depth > 0)
                        {
                            if (inCode[depth]) codeDepth--;
                            inCode[depth] = false;
                            depth--;
                        }
                        break;
                    default:
                        if (codeDepth > 0) mask[i] = true;
                        break;
                }
            }
            return mask;
        }

        // s ichidagi [a..b] guruhni hujjatda almashtiradi. Guruh formati bir xil bo'lmasa,
        // ikkiga bo'lib, oxir-oqibat har bir belgini alohida almashtiradi.
        private void ApplyGroup(dynamic story, int baseOffset, int a, int b, string s, string[] pieces)
        {
            while (a <= b && pieces[a].Length == 1 && pieces[a][0] == s[a]) a++;
            while (b >= a && pieces[b].Length == 1 && pieces[b][0] == s[b]) b--;
            if (a > b) return;

            dynamic g = story.Duplicate;
            g.SetRange(baseOffset + a, baseOffset + b + 1);
            g.TextRetrievalMode.IncludeFieldCodes = true;
            g.TextRetrievalMode.IncludeHiddenText = true;
            bool same = (string)g.Text == s.Substring(a, b - a + 1);

            if (same && (a == b || CanRewrite(g)))
            {
                var text = new System.Text.StringBuilder();
                for (int i = a; i <= b; i++)
                {
                    text.Append(pieces[i]);
                    if (pieces[i].Length != 1 || pieces[i][0] != s[i]) _changed++;
                }
                g.Text = text.ToString();
                return;
            }
            if (a == b)
            {
                _skipped++;
                return;
            }
            int mid = (a + b) / 2;
            ApplyGroup(story, baseOffset, mid + 1, b, s, pieces);
            ApplyGroup(story, baseOffset, a, mid, s, pieces);
        }

        // Guruhni bitta amal bilan yozish mumkinmi: barcha belgilar formati bir xil va
        // ichida xatcho'p, sharh, maydon, boshqaruv elementi yoki tuzatish yo'q.
        private static bool CanRewrite(dynamic g)
        {
            try
            {
                if ((int)g.Fields.Count > 0) return false;
                if ((int)g.Bookmarks.Count > 0) return false;
                if ((int)g.Comments.Count > 0) return false;
                if ((int)g.Revisions.Count > 0) return false;
                if ((int)g.InlineShapes.Count > 0) return false;
                try
                {
                    if ((int)g.ContentControls.Count > 0) return false; // Word 2007+
                }
                catch { }
                if (Undef(g.HighlightColorIndex) || Undef(g.LanguageID)) return false;

                dynamic f = g.Font;
                if (string.IsNullOrEmpty((string)f.Name) || string.IsNullOrEmpty((string)f.NameAscii) ||
                    string.IsNullOrEmpty((string)f.NameOther) || string.IsNullOrEmpty((string)f.NameBi)) return false;
                if (Undef(f.Size) || Undef(f.Bold) || Undef(f.Italic) || Undef(f.Underline) ||
                    Undef(f.UnderlineColor) || Undef(f.Color) || Undef(f.StrikeThrough) ||
                    Undef(f.DoubleStrikeThrough) || Undef(f.Superscript) || Undef(f.Subscript) ||
                    Undef(f.Hidden) || Undef(f.SmallCaps) || Undef(f.AllCaps) || Undef(f.Spacing) ||
                    Undef(f.Position) || Undef(f.Scaling) || Undef(f.Kerning) || Undef(f.Emboss) ||
                    Undef(f.Engrave) || Undef(f.Outline) || Undef(f.Shadow)) return false;
                dynamic shading = f.Shading;
                if (Undef(shading.BackgroundPatternColor) || Undef(shading.Texture)) return false;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool Undef(object value)
        {
            if (value == null) return true;
            try
            {
                return Math.Abs(Convert.ToDouble(value) - WdUndefined) < 0.5;
            }
            catch
            {
                return false;
            }
        }
    }
}
