; JYYJ助手 v2.18.0.0 合并安装器（Inno Setup 6）—— 一个 exe 同时适配 32 位与 64 位 Excel
;   1) 同时拷贝 x64 与 x86 两套完整 app + 依赖 到 %LocalAppData%\XaiAssistant\app\x64 与 app\x86
;   2) 安装完成由 64 位 PowerShell 执行 register-auto.ps1：探测 Excel 位数，显式写对应注册表视图
;      的 .NET COM 类 + Excel Addins 键（对进程位数免疫，32/64 位 Excel 各加载对应架构）
;   3) 卸载由 unregister-all.ps1 清理 64/32 两个视图的全部注册项
; 环境与依赖：本包自带全部托管程序集 + 各架构 SQLite.Interop + WebView2Loader；
;   系统级依赖 .NET Framework 4.x 与 WebView2 Runtime（Win10/11 自带）由 [Code] 前置检测提示。

#define AppVersion "2.18.0.0"
#define ProductName "JYYJ助手"
#define AppId "ExcelAiAssistant"
#define PkgSrc64 "D:\execl1.0\dist\JYYJ助手-2.18.0.0-x64\app"
#define PkgSrc86 "D:\execl1.0\dist\JYYJ助手-2.18.0.0-x86\app"

[Setup]
AppName={#ProductName}
AppVersion={#AppVersion}
AppId={#AppId}
AppPublisher=ExcelAiAssistant
DefaultDirName={localappdata}\XaiAssistant\app\x64
DefaultGroupName={#ProductName}
DisableProgramGroupPage=yes
AllowNoIcons=yes
CreateAppDir=yes
Compression=lzma2/max
SolidCompression=yes
OutputDir=D:\execl1.0\dist\installer
OutputBaseFilename={#ProductName}-{#AppVersion}-Setup
PrivilegesRequired=lowest
UsePreviousPrivileges=no
SetupLogging=yes
MinVersion=6.1sp1
UninstallDisplayName={#ProductName} {#AppVersion}（Excel 加载项）
UninstallDisplayIcon={localappdata}\XaiAssistant\app\x64\ExcelAiAssistant.Addin.dll
ShowLanguageDialog=no

[Languages]
Name: "chs"; MessagesFile: "compiler:Default.isl"

[Files]
; 64 位与 32 位两套完整 app + 依赖，分别装入对应架构目录
Source: "{#PkgSrc64}\*"; DestDir: "{localappdata}\XaiAssistant\app\x64"; Flags: recursesubdirs createallsubdirs ignoreversion
Source: "{#PkgSrc86}\*"; DestDir: "{localappdata}\XaiAssistant\app\x86"; Flags: recursesubdirs createallsubdirs ignoreversion
; 注册编排与卸载清理脚本（置于用户数据根目录，不随架构变动）
Source: "D:\execl1.0\tools\installer\register-auto.ps1"; DestDir: "{localappdata}\XaiAssistant"; Flags: ignoreversion
Source: "D:\execl1.0\tools\installer\unregister-all.ps1"; DestDir: "{localappdata}\XaiAssistant"; Flags: ignoreversion

; 安装完成后静默执行注册（{sysnative}=64 位 PowerShell，脚本用显式 RegistryView 双视图可写）
[Run]
Filename: "{sysnative}\WindowsPowerShell\v1.0\powershell.exe"; \
  Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{localappdata}\XaiAssistant\register-auto.ps1"""; \
  Flags: runhidden waituntilterminated; \
  StatusMsg: "正在注册 Excel 加载项…"

; 卸载时先清理注册项，再删除文件（脚本显式清 64/32 两个视图）
[UninstallRun]
Filename: "{sysnative}\WindowsPowerShell\v1.0\powershell.exe"; \
  Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{localappdata}\XaiAssistant\unregister-all.ps1"""; \
  Flags: runhidden; RunOnceId: UnregisterAll
