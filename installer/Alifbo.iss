; Alifbo - Word uchun o'rnatuvchi (Inno Setup 6)
; Joriy foydalanuvchi uchun o'rnatadi: administrator huquqi kerak emas.
; Plagin 32 va 64 bitli Office uchun ham ro'yxatdan o'tkaziladi.

#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif
#ifndef BinDir
  #define BinDir "..\src\AddIn\bin\Release\net40"
#endif
#define ClsId "{{5D920E66-8009-4A2B-8956-325639112835}"
#define ProgId "Alifbo.Connect"
#define AsmName "Alifbo.Word, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"

[Setup]
AppId={{E8623E77-628B-424F-97D5-1857C89E3E1C}
AppName=Alifbo - Word uchun yangi lotin
AppVersion={#AppVersion}
AppPublisher=Alifbo
AppPublisherURL=https://github.com/azakapro/alifbo
DefaultDirName={localappdata}\Programs\Alifbo
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableReadyPage=yes
PrivilegesRequired=lowest
OutputBaseFilename=Alifbo-Word-Setup-{#AppVersion}
OutputDir=..\dist
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName=Alifbo - Word uchun yangi lotin
MinVersion=6.1sp1
CloseApplications=no

[Messages]
SetupAppTitle=Alifbo o'rnatuvchisi
SetupWindowTitle=Alifbo o'rnatuvchisi
WelcomeLabel1=Alifbo - Word uchun yangi lotin
WelcomeLabel2=Word'ga "Yangi lotinga" tugmasi qo'shiladi. U o'zbek matnini eski lotin yoki kirill alifbosidan yangi lotin alifbosiga (ö, ğ, ş, ç) formatni saqlagan holda o'giradi.%n%nDavom etish uchun "Keyingi" ni bosing.
ButtonNext=&Keyingi >
ButtonBack=< &Orqaga
ButtonCancel=Bekor qilish
ButtonInstall=&O'rnatish
ButtonFinish=&Tayyor
ExitSetupTitle=O'rnatishni to'xtatish
ExitSetupMessage=O'rnatish tugallanmadi. Chiqilsinmi?
StatusExtractFiles=Fayllar nusxalanmoqda...
FinishedHeadingLabel=Alifbo o'rnatildi
FinishedLabelNoIcons=Word'ni oching (ochiq bo'lsa - yopib, qayta oching). Home lentasida "Alifbo -> Yangi lotinga" tugmasi paydo bo'ladi.
ConfirmUninstall=Alifbo plaginini o'chirasizmi?

[Files]
Source: "{#BinDir}\Alifbo.Word.dll"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; ---- COM ro'yxati (32 bitli ko'rinish) ----
Root: HKCU32; Subkey: "Software\Classes\{#ProgId}"; ValueType: string; ValueData: "{#ProgId}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\{#ProgId}\CLSID"; ValueType: string; ValueData: "{#ClsId}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsId}"; ValueType: string; ValueData: "{#ProgId}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsId}\ProgId"; ValueType: string; ValueData: "{#ProgId}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Alifbo.Word.Connect"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBaseUrl}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsId}\Implemented Categories\{{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}"; ValueType: none
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Alifbo - yangi lotin"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: string; ValueName: "Description"; ValueData: "O'zbek matnini yangi lotin alifbosiga o'giradi"
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"

; ---- COM ro'yxati (64 bitli ko'rinish) ----
Root: HKCU64; Subkey: "Software\Classes\{#ProgId}"; ValueType: string; ValueData: "{#ProgId}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#ProgId}\CLSID"; ValueType: string; ValueData: "{#ClsId}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsId}"; ValueType: string; ValueData: "{#ProgId}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsId}\ProgId"; ValueType: string; ValueData: "{#ProgId}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Alifbo.Word.Connect"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsId}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:CodeBaseUrl}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsId}\Implemented Categories\{{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}"; ValueType: none; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Alifbo - yangi lotin"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: string; ValueName: "Description"; ValueData: "O'zbek matnini yangi lotin alifbosiga o'giradi"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64

; ---- Sozlamalar (mavjud bo'lsa saqlanadi) ----
Root: HKCU; Subkey: "Software\Alifbo"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Alifbo"; ValueType: string; ValueName: "ProtectedTerms"; ValueData: ""; Flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\Alifbo"; ValueType: string; ValueName: "Exceptions"; ValueData: ""; Flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\Alifbo"; ValueType: dword; ValueName: "SetUzbekLatinLanguage"; ValueData: "1"; Flags: createvalueifdoesntexist

[Code]
// COM uchun DLL manzili: file:///C:/Users/.../Alifbo.Word.dll
function CodeBaseUrl(Param: String): String;
begin
  Result := ExpandConstant('{app}\Alifbo.Word.dll');
  StringChangeEx(Result, '\', '/', True);
  Result := 'file:///' + Result;
end;

function HasDotNet4: Boolean;
var
  Install: Cardinal;
begin
  Result :=
    (RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Install', Install) and (Install = 1)) or
    (RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Client', 'Install', Install) and (Install = 1));
end;

function InitializeSetup: Boolean;
begin
  Result := True;
  if not HasDotNet4 then
  begin
    MsgBox('Kompyuterda .NET Framework 4 topilmadi.' + #13#10 +
           'Uni https://dotnet.microsoft.com/download/dotnet-framework sahifasidan o''rnatib, so''ng qayta urinib ko''ring.',
           mbError, MB_OK);
    Result := False;
  end;
end;

// Word plaginni "sekin ishga tushadi" deb o'chirib qo'ymasligi uchun (Office 2013+).
procedure SetDoNotDisable(Version: String; Enable: Boolean);
var
  Key: String;
begin
  Key := 'Software\Microsoft\Office\' + Version + '\Word\Resiliency\DoNotDisableAddinList';
  if Enable then
    RegWriteDWordValue(HKCU, Key, '{#ProgId}', 1)
  else
    RegDeleteValue(HKCU, Key, '{#ProgId}');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    SetDoNotDisable('15.0', True);
    SetDoNotDisable('16.0', True);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    SetDoNotDisable('15.0', False);
    SetDoNotDisable('16.0', False);
  end;
end;
