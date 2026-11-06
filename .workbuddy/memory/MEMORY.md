# 项目长期记忆（execl1.0-wps · WPS 兼容版）

> 版本史（M13–M20 逐版明细）在 docs/版本史.md 与 .workbuddy/memory/ 日志；此处只留现役事实。当前 v3.8.3.1（2026-10-02），门禁基线：1291 单测 · 156 工具卡点 · Excel 真机 119/119 · WPS full 131 组。**git 现状（2026-09-19 起）**：本目录是 git 仓库，origin=gitclone 镜像 github.com/Asaceoo/ExeclAi，推送走 SSH `git@github.com:Asaceoo/ExeclAi.git`（HTTPS 无写凭据）；本地分支 master、远端 main；⚠️ 截至 2026-10-02 本地 master 停在 3.8.2.9 基线提交、M20 改动未提交且与 origin/main 历史分叉（M20 手册经 worktree 单提交直推远端）。

## 发版与产物规范（用户强制要求）

- 产物命名：安装包 `JYYJ助手-<版本号>-Setup-wps.exe`、便携包 `JYYJ助手-<版本号>-wps-x64|x86.zip`——dist 下所有产物必须带 `wps` 后缀（Excel 主副本 D:\execl1.0 发版流不带，不受此约束）。
- 版本单一来源 = 根 `Directory.Build.props`；升版**只走** `release.py bump --ver x.y.z.w`，打包**只走** `release.py pack`（或 `gate.py --pack` 一条门禁直通发版）。iss 模板无硬编码版本（ISCC /DAppVersion 注入）；PkgSrc 指向 dist 便携目录，改目录名必须同步 iss。
- bump 自动联动（sync_manuals，M17.6/17.7）：四份手册（技术/用户 × md/html）**文件名+版本戳**、CLAUDE.md/AI交接文档/README 手册引用、安装说明模板"适用版本"戳——一次同步，含零残留断言；v 前缀叙事不自动改（归 neat-freak）。
- 手册正文更新流程：改 md → `tools/release/gen_manual_html.py <md> <同md> <同路径html>` 重生成 → 重跑 pack 进包；**每次 Edit 后必须 grep 复核落盘**（编辑报成功但未落盘的坑两次复现）。
- `tools/package-xai.ps1` 是 Excel 主线遗留，本仓库不用。

## 工具与门禁体系（现役）

- verify_* 自检族 18 个；MappedPrimaryTools 26 类 / 豁免 77 条，覆盖审计两执行点 + 一致性锁单测——新写工具不表态发不出去。
- SheetProtectionGuard：非 ReadOnly 工具统一 SHEET_PROTECTED 前置探测（守卫是加法前置，探测异常一律放行）；**探测类 COM 对 Workbook/Worksheet 级对象零 FinalRelease**（RCW 按 IUnknown 身份缓存，释放连带分离外部持有人）。
- gate.py 真机步前快照宿主进程、步后 assert_no_host_leak 断言——harness 中断泄漏的 DCOM Excel 跨 silo 不可杀、会锁死用户安装，泄漏必须当场亮红（M17.5）。

## WPS 关键差异（探针实测）

- 冻结窗格：Excel 的「Select 锚点 + FreezePanes=true」在 WPS ET 被静默忽略，必须 SplitRow/SplitColumn + FreezePanes=true（已统一走 Split 双宿主通用）。
- Excel 错误单元格 Value2 读回 Int32 CVErr（#N/A=-2146826246），必须映射回错误码计入分布（Power Query 同口径）。
- AddPicture(-1,-1) 原生尺寸行为两宿主一致；AddPicture 对垃圾字节文件 Excel 宽容接受、WPS 报 E_FAIL——防线前移做魔数白名单嗅探。

## 安装器与随包支持（M17.5/M17.7 定型）

- Inno `PrivilegesRequired=lowest` 永远普通权限（右键管理员运行也被降权）；提权唯一正路 = `ShellExec('runas')` UAC；击杀名单含 wpsoffice.exe；击杀日志持久化 `%LocalAppData%\XaiAssistant\logs\hostkill.log`；/SILENT 自动强杀宿主。详见踩坑手册 §6.11。
- 三件套随包（M17.7）：diagnose.ps1（**双宿主诊断**：Excel + WPS——Kingsoft ET Addins 双视图/AddinsWL/et.exe PE 位数/KET.Application COM 探测；键位必须与 register-auto.ps1 写入逻辑同源）+ 安装说明.txt（tools/installer/ 为单一来源）+ 用户手册.html（pack 从 docs 版本化手册拷入，文件名不带版本）。开始菜单三入口。
- 地面真值（诊断脚本防虚构基准）：本机 WPS x86 12.1.0.26899、KET.Application ProgID 在、WPS 个人版无 VBA 且**无可远程判定的注册键/文件信号**——诊断对 VBA 只指引 Alt+F11 自查，不断言。

## 本机环境坑（反复踩）

- 沙箱会话缺核心环境变量 → dotnet 全挂：命令内补全 APPDATA/SystemRoot/ProgramFiles(x86)/TMP/TEMP/USERPROFILE 等（gate/release 的 ENV_FIX 已兜底）。
- 同一消息内并行多个 Edit 同一文件会静默丢失部分编辑——同文件多处修改必须串行 Edit。
- Bash 里 grep 管道接后台 exe 输出偶发空（UTF-16 误判）：真机日志先落盘再 `grep -a` 读。
- 沙箱内 taskkill/杀进程类命令会被 SIGTERM 拦截：杀进程逻辑必须实跑目标程序（如安装器本体）验证。
- harness 断言注意匿名对象 snake_case vs 命名类 PascalCase；dynamic 打跨程序集测试假对象必抛 RuntimeBinderException——判定逻辑抽接口接缝才可单测。
- 诱饵宿主保活：`copy cmd.exe et.exe` + `/C ping -t 127.0.0.1`（notepad 改名立即退出、cmd /K 读 EOF 退出）。
