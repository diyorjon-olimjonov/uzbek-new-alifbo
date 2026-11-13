// Dvigatelni asl alifbo kutubxonasi natijalari bilan solishtiradi.
// vectors.txt: har bir qatorda "kirish<TAB>kutilgan", UTF-16 kodlari hex ko'rinishida.
using System;
using System.IO;
using System.Linq;
using System.Text;
using Alifbo;

internal static class Program
{
    private static string FromHex(string h)
    {
        if (h == "x") return "";
        return new string(h.Split(' ').Select(p => (char)Convert.ToInt32(p, 16)).ToArray());
    }

    private static string Show(string s)
    {
        var sb = new StringBuilder();
        foreach (char c in s)
            sb.Append(c < 0x20 || c > 0x7E ? $"\\u{(int)c:X4}" : c.ToString());
        return sb.ToString();
    }

    private static int Main(string[] args)
    {
        string path = args.Length > 0 ? args[0] : Path.Combine(AppContext.BaseDirectory, "vectors.txt");
        var engine = new AlifboEngine();
        int total = 0, failed = 0;
        foreach (var line in File.ReadAllLines(path))
        {
            if (line.Length == 0) continue;
            var parts = line.Split('\t');
            string input = FromHex(parts[0]), expected = FromHex(parts[1]);
            string actual = engine.Convert(input);
            total++;
            if (actual != expected)
            {
                failed++;
                if (failed <= 20)
                    Console.WriteLine($"FAIL\n  in : {Show(input)}\n  exp: {Show(expected)}\n  got: {Show(actual)}");
            }
            // Bo'laklar soni manba uzunligiga teng bo'lishi shart.
            if (engine.Pieces(input, null).Length != input.Length)
            {
                failed++;
                Console.WriteLine($"FAIL pieces length: {Show(input)}");
            }
        }

        // Foydalanuvchi sozlamalari
        var custom = new AlifboEngine(new[] { "Microsoft" },
            new System.Collections.Generic.Dictionary<string, string> { { "mushaf", "musʼhaf" } });
        Check(custom.Convert("Microsoft shahar mushaf"), "Microsoft şahar musʼhaf", ref total, ref failed);
        // Maska: belgilangan belgilar o'zgarmaydi
        var mask = new bool[] { true, true, false, false };
        Check(string.Concat(engine.Pieces("shsh", mask)), "shş", ref total, ref failed);

        Console.WriteLine($"{total - failed}/{total} passed");
        return failed == 0 ? 0 : 1;
    }

    private static void Check(string actual, string expected, ref int total, ref int failed)
    {
        total++;
        if (actual == expected) return;
        failed++;
        Console.WriteLine($"FAIL\n  exp: {Show(expected)}\n  got: {Show(actual)}");
    }
}
