# Alifbo — Word uchun yangi lotin plagini

Word'ning **Home** lentasiga **Alifbo → Yangi lotinga** tugmasini qoʻshadi. Bir bosishda oʻzbek matnini **eski lotin** yoki **kirill** alifbosidan **yangi lotin** alifbosiga (2026) oʻgiradi:
`oʻ → ö`, `gʻ → ğ`, `sh → ş`, `ch → ç`. Kirill harflari [alifbo](https://github.com/azakapro/alifbo) kutubxonasi qoidalari boʻyicha oʻgiriladi (е → ye/e, ц → s/ts, ё, ю, я, ъ, ь…).

- **Format saqlanadi:** faqat oʻzgargan harflar almashtiriladi. Shrift, oʻlcham, qalin, kursiv, rang va highlight har bir harfda oʻz joyida qoladi.
- **Nima oʻgiriladi:** matn belgilangan boʻlsa, faqat oʻsha qism, aks holda butun hujjat: asosiy matn, jadvallar, kolontitullar, izohlar, sharhlar va matn qutilari.
- **Nimaga tegilmaydi:** URL, e-pochta, `` `kod` `` qismlari va maydon kodlari (`{ HYPERLINK }`, `{ PAGE }`).
- **Bekor qilish:** hammasi bitta **Ctrl+Z** bilan bekor qilinadi.

## Qoʻllab-quvvatlanadi

| | |
|---|---|
| Word | 2010, 2013, 2016, 2019, 2021, 2024, Microsoft 365 (32 va 64 bit) |
| Windows | 7 SP1, 8.1, 10, 11 |
| Talab | .NET Framework 4.x (Windows 8 dan boshlab oʻrnatilgan) |
| Ishlamaydi | Mac, Word Online, iPad/Android |

## Oʻrnatish

1. [Releases](../../releases) boʻlimidan `Alifbo-Word-Setup-x.y.z.exe` ni yuklab oling.
2. Ishga tushiring: **Keyingi → Oʻrnatish**. Administrator huquqi kerak emas.
3. Word'ni oching (ochiq boʻlsa, qayta oching).

> Windows “Noma'lum nashriyotchi” (SmartScreen) deb ogohlantirsa: **Batafsil → Baribir ishga tushirish**. Ogohlantirish oʻrnatuvchi kod imzolash sertifikati bilan imzolanmaguncha chiqadi.

Oʻchirish: **Sozlamalar → Ilovalar → Alifbo → Oʻchirish**.

## Sozlamalar (ixtiyoriy)

Sozlamalar `HKEY_CURRENT_USER\Software\Alifbo` reestr kalitida saqlanadi:

| Qiymat | Tur | Misol | Maʼnosi |
|---|---|---|---|
| `ProtectedTerms` | matn | `Microsoft;Chelsea FC` | Oʻgirilmaydigan soʻzlar (katta-kichik harf farqlanadi) |
| `Exceptions` | matn | `mushaf=musʼhaf` | Soʻz istisnolari |
| `SetUzbekLatinLanguage` | DWORD | `1` | “Oʻzbek (kirill)” tilini “Oʻzbek (lotin)” ga oʻtkazish |

## Plagin ishlamasa

- **Word → File → Options → Add-ins → Manage: COM Add-ins → Go…** oynasida **Alifbo – yangi lotin** belgilanganini tekshiring.
- U **Disabled Items** roʻyxatida boʻlsa, oʻsha yerdan yoqing.
- Tashkilotda administrator plaginlarni taqiqlagan boʻlishi mumkin. Bunda IT boʻlimiga murojaat qiling.

## Ishlab chiquvchilar uchun

```
src/Engine/AlifboEngine.cs   oʻgirish dvigateli (Word'ga bogʻliq emas)
src/AddIn/                   COM plagin: lenta, Word bilan ishlash, sozlamalar
tests/Engine.Tests/          asl alifbo kutubxonasi natijalari bilan 436 ta sinov
installer/Alifbo.iss         Inno Setup oʻrnatuvchisi (joriy foydalanuvchi uchun)
.github/workflows/build.yml  sinov, yigʻish, oʻrnatuvchi, COM tekshiruvi, release
```

Plagin Office PIA'larsiz yozilgan: COM interfeyslari oʻzida eʼlon qilingan, Word bilan esa `dynamic` orqali ishlaydi. Shu tufayli bitta `net40` DLL barcha Word versiyalarida ishlaydi.

Yangi versiya chiqarish uchun `v1.0.0` kabi teg qoʻyiladi. GitHub Actions oʻrnatuvchini yigʻib, Releases boʻlimiga joylaydi.

```
git tag v1.0.0 && git push --tags
```

Mahalliy yigʻish (Windows, .NET 8 SDK va Inno Setup 6 kerak):

```
dotnet run --project tests/Engine.Tests -c Release
dotnet build src/AddIn/Alifbo.Word.csproj -c Release
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\Alifbo.iss
```

O'girish qoidalari [azakapro/alifbo](https://github.com/azakapro/alifbo) kutubxonasidan olingan.
