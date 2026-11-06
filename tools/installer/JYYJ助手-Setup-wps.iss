; JYYJ助手 WPS 版安装器（Inno Setup 6）—— 双包双后缀的 WPS 分发包【版本号模板化】
;   版本号单一来源 = Directory.Build.props；本文件不再硬编码版本，
;   由 tools/release/release.py pack 以 ISCC /DAppVersion=x.y.z.w 注入；
;   直接双击 ISCC 编译时使用下方兜底默认值（仅限本地调试，发版必须走 release.py）。
;   与 Excel 版（JYYJ助手-X-Setup）同 base 版本号，-wps 仅在产物文件名；
;   同 AppId / 同安装目录：WPS 包即包含双宿主能力的升级包，覆盖安装 Excel 版互不冲突。
;   1) 同时拷贝 x64 与 x86 两套完整 app + 依赖 到 %LocalAppData%\XaiAssistant\app\x64 与 app\x86
;   2) 安装完成由 64 位命令行执行 register-auto.ps1：探测 Excel 与 WPS（et.exe）位数，
;      显式写对应注册表视图的 .NET COM 类 + Excel Addins 键 + Kingsoft ET Addins/AddinsWL 白名单
;   3) 卸载由 unregister-all.ps1 清理 64/32 两个视图的全部注册项（Excel 与 WPS 一并清理）
; 环境与依赖：本包自带全部托管程序集 + 各架构 SQLite.Interop + WebView2Loader；
;   支持 WPS 表格（个人版白名单加载 / 专业版 v11.8.2+）与 Excel 2013~365（双宿主）；
;   系统级依赖 .NET Framework 4.8 与 WebView2 Runtime（Win10/11 一般自带）由 [Code] 前置检测提示。

#ifndef AppVersion
  #define AppVersion "3.7.0.0"
#endif
#define ProductName "JYYJ助手"
#define AppId "ExcelAiAssistant"
#ifndef RepoRoot
  ; 审查 R7：不再硬编码仓库绝对路径——由 ISCC 的 SourcePath（本 iss 所在目录）向上两级推导，
  ; 这样换机器 / CI / 不同盘符都能编译；release.py 亦可通过 /DRepoRoot=... 显式注入。
  #define RepoRoot AddBackslash(SourcePath) + "..\.."
#endif
#define DistRoot RepoRoot + "\dist"
#define PkgSrc64 DistRoot + "\JYYJ助手-" + AppVersion + "-wps-x64\app"
#define PkgSrc86 DistRoot + "\JYYJ助手-" + AppVersion + "-wps-x86\app"
#define InstallNote DistRoot + "\JYYJ助手-" + AppVersion + "-wps-安装说明.txt"

[Setup]
AppName={#ProductName}
AppVersion={#AppVersion}
AppId={#AppId}
AppPublisher=ExcelAiAssistant
DefaultDirName={localappdata}\XaiAssistant\app\x64
DefaultGroupName={#ProductName}
DisableProgramGroupPage=yes
DisableDirPage=yes
AllowNoIcons=yes
CreateAppDir=yes
Compression=lzma2/max
SolidCompression=yes
OutputDir=D:\execl1.0-wps\dist\installer
OutputBaseFilename={#ProductName}-{#AppVersion}-Setup-wps
PrivilegesRequired=lowest
UsePreviousPrivileges=no
SetupLogging=yes
MinVersion=6.1sp1
UninstallDisplayName={#ProductName} {#AppVersion}（Excel/WPS 双宿主加载项）
UninstallDisplayIcon={localappdata}\XaiAssistant\app\x64\ExcelAiAssistant.Addin.dll
ShowLanguageDialog=no

[Languages]
Name: "chs"; MessagesFile: "compiler:Default.isl"

[Files]
; 64 位与 32 位两套完整 app + 依赖，分别装入对应架构目录
Source: "{#PkgSrc64}\*"; DestDir: "{localappdata}\XaiAssistant\app\x64"; Flags: recursesubdirs createallsubdirs ignoreversion
Source: "{#PkgSrc86}\*"; DestDir: "{localappdata}\XaiAssistant\app\x86"; Flags: recursesubdirs createallsubdirs ignoreversion
Source: "{#InstallNote}"; DestDir: "{localappdata}\XaiAssistant"; Flags: ignoreversion
; 注册编排与卸载清理脚本（WPS 感知版，置于用户数据根目录，不随架构变动）
Source: "{#RepoRoot}\tools\installer\register-auto.ps1"; DestDir: "{localappdata}\XaiAssistant"; Flags: ignoreversion
Source: "{#RepoRoot}\tools\installer\unregister-all.ps1"; DestDir: "{localappdata}\XaiAssistant"; Flags: ignoreversion
; 用户支持文件：小白安装说明 + 一键诊断 + 用户手册（报错时用户自助取证）
Source: "{#RepoRoot}\tools\installer\安装说明.txt"; DestDir: "{localappdata}\XaiAssistant"; Flags: ignoreversion
Source: "{#RepoRoot}\tools\installer\diagnose.ps1"; DestDir: "{localappdata}\XaiAssistant"; Flags: ignoreversion
Source: "{#RepoRoot}\tools\installer\一键诊断.cmd"; DestDir: "{localappdata}\XaiAssistant"; Flags: ignoreversion
; 用户手册 HTML（release.py pack 从 docs 版本化手册拷入，文件名不带版本防 bump 断链）
Source: "{#RepoRoot}\tools\installer\用户手册.html"; DestDir: "{localappdata}\XaiAssistant"; Flags: ignoreversion

; 开始菜单：安装说明 / 用户手册 / 一键诊断入口（用户报错时的第一自救路径）
[Icons]
Name: "{group}\JYYJ助手 安装说明"; Filename: "{localappdata}\XaiAssistant\安装说明.txt"
Name: "{group}\JYYJ助手 用户手册"; Filename: "{localappdata}\XaiAssistant\用户手册.html"
Name: "{group}\JYYJ助手 一键诊断"; Filename: "{localappdata}\XaiAssistant\一键诊断.cmd"; WorkingDir: "{localappdata}\XaiAssistant"

; 安装完成后静默执行注册（{sysnative}=64 位宿主脚本入口，脚本用显式 RegistryView 双视图可写）
[Run]
Filename: "{sysnative}\WindowsPowerShell\v1.0\powershell.exe"; \
  Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{localappdata}\XaiAssistant\register-auto.ps1"""; \
  Flags: runhidden waituntilterminated; \
  StatusMsg: "正在注册 Excel/WPS 加载项…"

; 卸载时先清理注册项，再删除文件（脚本显式清 64/32 两个视图）
[UninstallRun]
Filename: "{sysnative}\WindowsPowerShell\v1.0\powershell.exe"; \
  Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{localappdata}\XaiAssistant\unregister-all.ps1"""; \
  Flags: runhidden; RunOnceId: UnregisterAll

[Code]
// 检测宿主是否运行（Excel 与 WPS 表格均须关闭——两者运行中都会锁 DLL 并回写注册，
// WPS 运行中安装会导致加载项注册被旧路径覆盖，装完也无法正常生效）
// 注意：用 CSV 输出 + 带引号进程名匹配实现【精确】比较——若用裸子串匹配，
// "et.exe" 会误命中 "dotnet.exe" 等名称尾部含 et.exe 的进程，造成假拦截。
function IsHostRunning(): Boolean;
var
  cmd: String;
  buf: AnsiString;
  ec: Integer;
  outFile: String;
  lines: TStringList;
  i, p, memKB: Integer;
  line, memStr: String;
begin
  Result := False;
  outFile := ExpandConstant('{tmp}\hostchk.txt');
  // 不用 /FI 过滤器：实测其可靠性不足（本地化系统/重定向下可能失灵漏报）；
  // /FO CSV 使进程名带双引号成独立字段，"et.exe" 只会精确命中 WPS 表格进程
  cmd := Format('/C tasklist /NH /FO CSV > "%s" 2>&1', [outFile]);
  if Exec(ExpandConstant('{cmd}'), cmd, '', SW_HIDE, ewWaitUntilTerminated, ec) then
    if LoadStringFromFile(outFile, buf) then
    begin
      // 2026-09-12 僵尸进程排除（v3.7.1.34 用户实证）：探针遗留的已终止 EXCEL.EXE
      // 残留在 tasklist 快照里（内存仅 52K、Handles=0），taskkill/wmic 任何权限都
      // 清不掉（Access denied），而它已终止不锁任何文件——若按进程名判"宿主在运行"
      // 会让安装死循环（警告弹窗场景 4"重启电脑才能装"）。真实宿主进程内存为
      // 几十~几百 MB，此处把内存 <5000K 的残留行视为僵尸排除，不触发拦截。
      lines := TStringList.Create;
      try
        lines.Text := String(buf);
        for i := 0 to lines.Count - 1 do
        begin
          line := lines[i];
          if (Pos('"et.exe"', line) > 0) or (Pos('"EXCEL.EXE"', line) > 0) or
             (Pos('"excel.exe"', line) > 0) then
          begin
            p := RPos(',', line);
            if p <= 0 then
              Result := True  // 格式异常宁可信其有（fail-closed）
            else
            begin
              memStr := Trim(Copy(line, p + 1, MaxInt));
              StringChangeEx(memStr, '"', '', True);
              StringChangeEx(memStr, ' K', '', True);
              StringChangeEx(memStr, 'K', '', True);
              StringChangeEx(memStr, ',', '', True);
              memKB := StrToIntDef(Trim(memStr), -1);
              if memKB < 0 then
                Result := True  // 解析不出内存 → 按活进程处理（fail-closed）
              else if memKB >= 5000 then
              begin
                Result := True;  // 真实运行的宿主进程
                break;
              end;
              // memKB < 5000：僵尸残留行，跳过继续扫下一行
            end;
            if Result then break;
          end;
        end;
      finally
        lines.Free;
      end;
      // 命中诊断：把判定为"宿主在运行"时的原始快照留在 tmp，供排查假拦截
      if Result then
        SaveStringToFile(ExpandConstant('{tmp}\hostchk_hit.csv'), buf, False);
    end;
  DeleteFile(outFile);
end;

// 安装前置硬防线：Excel / WPS 运行中会把加载项注册回写为旧路径（真机实测确认），
// 导致"重装也修不好"。必须先关闭全部宿主才允许继续。
// 三按钮方案（2026-09-08 用户反馈"只有重试/取消太麻烦"）：
//   【是】自动关闭全部宿主进程并继续（taskkill /F，未保存内容会丢失，弹窗已警示）
//   【否】用户手动关闭后重新检测（等价原"重试"）
//   【取消】退出安装
//
// 2026-09-09 二次修复"管理员身份运行也杀不死"（用户真机反馈截图实证）——根因两条：
//   A)【主因】PrivilegesRequired=lowest 下 Inno 永远以普通权限运行：即使右键
//      "以管理员身份运行"，Inno 也会降权重启安装器——宿主 Excel 提权启动时，
//      普通权限 taskkill /F 永远"拒绝访问"，而旧弹窗指引"右键管理员运行"完全无效。
//      → 修复：普通权限 3 轮杀不死时，改用 ShellExec('runas') 发起 UAC 提权限强杀
//      （这才是唯一真实生效的提权路径），弹窗文案同步改为真实归因。
//   B) 击杀名单缺 wpsoffice.exe：WPS 2019+ 主控/守护进程，单杀 et.exe 会被它拉起。
//      → 修复：加入击杀名单（/T 进程树一并收割）。
// 附带：静默安装（/SILENT，企业批量部署）自动强杀宿主，杀不死才中止——
//   既是自动化真机验证的前提，也是部署惯例（Office 加载项安装器普遍自动关闭宿主）。
// 2026-09-09 真机验证补充（M17.5）：击杀日志从 {tmp} 改到
//   {localappdata}\XaiAssistant\logs\hostkill.log——{tmp} 安装结束即删，失败后无据可查。
function KillLogPath(): String;
begin
  Result := ExpandConstant('{localappdata}\XaiAssistant\logs\hostkill.log');
  ForceDirectories(ExtractFilePath(Result));
end;

procedure KillHostTree(elevate: Boolean; const tmpLog: String);
var
  args: String;
  ec: Integer;
begin
  // 输出统一追加持久化日志（KillLogPath）供诊断（静默失败是 M17.2 事故根因之首）。
  // 注意：Format 数组参数必须与调用同行——Inno 会把行首 "[" 误判为 section 标签（编译实证）。
  args := Format('/C taskkill /F /T /IM EXCEL.EXE /IM et.exe /IM wps.exe /IM wpp.exe /IM wpsoffice.exe >> "%s" 2>&1', [tmpLog]);
  if elevate then
    // UAC 提权执行：用户在 UAC 弹窗点"是"后，taskkill 以管理员身份运行；
    // 用户点"否"/取消时 ShellExec 返回 False（ec=1223），由上层复查归因。
    ShellExec('runas', ExpandConstant('{cmd}'), args, '', SW_HIDE, ewWaitUntilTerminated, ec)
  else
    Exec(ExpandConstant('{cmd}'), args, '', SW_HIDE, ewWaitUntilTerminated, ec);
end;

function ForceKillHostProcesses(): Boolean;
var
  killRound: Integer;
  tmpLog: String;
begin
  tmpLog := KillLogPath();
  Result := False;
  // 第一阶段：普通权限 3 轮（覆盖宿主与安装器同权限的常规场景）
  for killRound := 1 to 3 do
  begin
    KillHostTree(False, tmpLog);
    Sleep(1000); // 等进程树完全退出、DLL 文件锁释放后再复查
    Result := not IsHostRunning();
    if Result then
      Exit;
  end;
  // 第二阶段：普通权限杀不死 → 宿主提权运行，弹 UAC 以管理员强杀一轮
  // （交互模式先弹窗预告 UAC，避免用户不知道要点"是"——真机验证实证的盲点）
  if not WizardSilent() then
    MsgBox('普通权限无法结束宿主进程——宿主可能以"管理员身份"运行。' #13#10 #13#10
          '接下来会弹出 UAC 提权窗口，请在其中点击"是"，' #13#10
          '以管理员身份强制关闭全部 Excel / WPS 进程。',
          mbInformation, MB_OK);
  KillHostTree(True, tmpLog);
  Sleep(1000);
  Result := not IsHostRunning();
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  // 静默安装（/SILENT，企业批量部署与自动化验证）：不弹交互框，自动强杀宿主，
  // 杀不死才中止（SetupLogging=yes 时归因可在安装日志与 {tmp}\hostkill.log 追溯）。
  if WizardSilent() then
  begin
    if IsHostRunning() and not ForceKillHostProcesses() then
      Result := '自动关闭失败：Excel / WPS 进程无法结束（可能以管理员身份运行或被安全软件拦截），已中止安装。请手动关闭全部 Excel 与 WPS 后重试。';
    Exit;
  end;
  while IsHostRunning() do
  begin
    case MsgBox('检测到 Excel 或 WPS 表格正在运行。' #13#10 #13#10
              '宿主运行中安装会导致注册信息被覆盖，装完也无法正常生效。' #13#10 #13#10
              '【是】自动关闭全部宿主进程并继续安装（未保存的表格内容会丢失）' #13#10
              '【否】我已手动关闭，请重新检测' #13#10
              '【取消】退出安装',
              mbConfirmation, MB_YESNOCANCEL) of
      IDYES:
        if not ForceKillHostProcesses() then
        begin
          // 击杀失败必须归因。注意：本安装器以普通权限运行（PrivilegesRequired=lowest，
          // 右键"以管理员身份运行"对 Inno 无效），提权强杀走的是第二阶段的 UAC 弹窗。
          MsgBox('自动关闭失败：已多轮强制结束（含 UAC 提权限强杀），仍有 Excel / WPS 进程存活。' #13#10 #13#10
                '常见原因与处理：' #13#10
                '  1) 屏幕上出现过 UAC 提权弹窗但被取消/未确认——请再选一次【是】，' #13#10
                '     并在弹出的 UAC 窗口中点击"是"；' #13#10
                '  2) 安全软件拦截了强制结束——请手动关闭 Excel 与 WPS 后选【否】；' #13#10
                '  3) WPS 后台组件自动拉起进程——请退出 WPS（含托盘图标）后选【否】；' #13#10
                '  4) 若连提权强杀也报"没有此任务的实例"——进程可能由系统服务或' #13#10
                '     自动化测试遗留拉起，任何权限都无法结束，请重启电脑后重新安装。' #13#10 #13#10
                '点击"确定"后可重新选择【是】重试。',
                mbError, MB_OK);
        end;
      IDCANCEL:
        begin
          Result := '已取消安装：请先关闭全部 Excel 与 WPS 表格，再重新运行安装包。';
          Exit;
        end;
      // IDNO：回到 while 循环重新检测
    end;
  end;
end;
