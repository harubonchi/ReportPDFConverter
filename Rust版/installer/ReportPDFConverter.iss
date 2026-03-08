#define MyAppName "Report PDF Converter"
#define MyAppPublisher "ReportPDFConverter"
#define MyAppExeName "report-pdf-converter.exe"
#define MyAppId "{{8F280A6A-7150-40E6-9985-7B35591A6436}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVerName={#MyAppName}
AppPublisher={#MyAppPublisher}
UninstallDisplayIcon={app}\static\favicon.ico
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\dist\installer
OutputBaseFilename=ReportPDFConverter-Setup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
WizardStyle=modern
SetupIconFile=..\static\installer.ico

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップショートカットを作成"; GroupDescription: "追加タスク:"; Flags: unchecked

[Files]
Source: "..\target\release\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\static\*"; DestDir: "{app}\static"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\lab_members\*"; DestDir: "{localappdata}\ReportPDFConverter\data\lab_members"; Flags: ignoreversion recursesubdirs createallsubdirs onlyifdoesntexist uninsneveruninstall
Source: "..\gmail_api_credentials\*"; DestDir: "{localappdata}\ReportPDFConverter\data\gmail_api_credentials"; Flags: ignoreversion recursesubdirs createallsubdirs skipifsourcedoesntexist onlyifdoesntexist uninsneveruninstall
Source: "..\..\Python版\gmail_api_credentials\*"; DestDir: "{localappdata}\ReportPDFConverter\data\gmail_api_credentials"; Flags: ignoreversion recursesubdirs createallsubdirs skipifsourcedoesntexist onlyifdoesntexist uninsneveruninstall

[InstallDelete]
Type: filesandordirs; Name: "{app}\lab_members"
Type: filesandordirs; Name: "{app}\gmail_api_credentials"
Type: filesandordirs; Name: "{app}\data"

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{#MyAppName} を起動"; Flags: nowait postinstall skipifsilent

[Code]
procedure WriteEnvFiles();
var
  AppDir: String;
  Content: String;
  SamplePath: String;
begin
  AppDir := ExpandConstant('{app}');
  Content :=
    '# このファイル名を .env に変更すると設定が有効になります。' + #13#10 +
    '# サーバーの待受ポート（未指定時は80）' + #13#10 +
    'RPC_PORT=80' + #13#10 +
    '' + #13#10 +
    '# Gmail APIのクライアント認証情報ファイルへのパス' + #13#10 +
    '# GMAIL_CREDENTIALS_JSON=.\gmail_api_credentials\credentials.json' + #13#10 +
    '# Gmail APIのトークンファイルへのパス' + #13#10 +
    '# GMAIL_TOKEN_JSON=.\gmail_api_credentials\token.json' + #13#10 +
    '# メール送信者アドレス（From）' + #13#10 +
    '# EMAIL_SENDER=roboken.report.tool@gmail.com' + #13#10 +
    '# メール送信者の表示名' + #13#10 +
    '# EMAIL_DISPLAY_NAME=ロボ研報告書作成ツール' + #13#10 + #13#10 +
    '# SMTPサーバーのホスト名（Gmail API失敗時のフォールバック）' + #13#10 +
    '# SMTP_HOST=smtp.example.com' + #13#10 +
    '# SMTPサーバーのポート番号' + #13#10 +
    '# SMTP_PORT=587' + #13#10 +
    '# SMTP認証ユーザー名' + #13#10 +
    '# SMTP_USER=user@example.com' + #13#10 +
    '# SMTP認証パスワード' + #13#10 +
    '# SMTP_PASS=change_me' + #13#10 +
    '# SMTP送信元アドレス（未指定時はSMTP_USER）' + #13#10 +
    '# SMTP_FROM=user@example.com' + #13#10 +
    '# SMTP送信元表示名' + #13#10 +
    '# SMTP_FROM_NAME=Report PDF Converter' + #13#10;

  SamplePath := AddBackslash(AppDir) + '.env.sample';
  SaveStringToFile(SamplePath, Content, False);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    WriteEnvFiles();
end;
