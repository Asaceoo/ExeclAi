# M0-WPS 加载可行性报告

> 里程碑：M0（POC 硬门）· 状态：**通过** · 日期：2026-09-08
> 输入：《设计-WPS兼容版需求文档.md》v1.0 · 产出：本文档 + 可复用注册/回滚脚本
> 结论先行：**WPS 表格 12.1 个人版原生支持本项目的 COM 加载项全链路——OnConnection、Ribbon XML、CustomTaskPane、WebView2 全部工作**，最大风险项（侧边栏任务窗格）被证伪，侧边栏走形态 A，无需浮动窗兜底。

## 1. 环境

| 项 | 值 |
|---|---|
| WPS 版本 | 12.1.0.26899（个人版，`C:\Program Files (x86)\Kingsoft\WPS Office\...\office6\et.exe`） |
| et.exe PE 位数 | x86（machine=0x014C）→ 采用 Registry32 视图 + x86 构建 |
| 被加载产物 | `D:\execl1.0-wps\src\ExcelAiAssistant.Addin\bin\x86\Release\net48\win-x86\ExcelAiAssistant.Addin.dll`（v2.19.6.0） |
| 注册脚本 | `D:\execl1.0\_dev_tmp\wps_register.py`（register / unregister / verify 三动作，全部 HKCU，可随时回滚） |

## 2. 注册层（M0 落地配方）

在现有 `register-auto.ps1` 配方基础上，WPS 需追加两组键（均为 HKCU、32 位视图）：

| 键 | 内容 |
|---|---|
| `HKCU\Software\Kingsoft\Office\ET\Addins\<ProgId>` | FriendlyName / Description / **LoadBehavior=3** |
| `HKCU\SOFTWARE\kingsoft\Office\ET\AddinsWL` | 各 ProgId 空字符串值（个人版白名单） |

COM 类注册（`HKCU\Software\Classes\CLSID\{GUID}` + InprocServer32/ProgId/Implemented Categories）与 Excel 版完全一致，无需改动；3 个类（AddIn / TaskPaneHost / DeepSeekPaneHost）全部注册并回读验证通过。

## 3. 加载验证证据（et.exe 启动后 `%TEMP%\ExcelAiAssistant.log` 新增行）

```
04:20:10.104 [INFO] OnConnection connectMode=ext_cm_Startup          ← IDTExtensibility2 触发 ✓
04:20:10.108 [INFO] GetCustomUI ribbonId=Microsoft.Excel.Workbook resourceFound=True  ← WPS 按 Excel 同名 ribbonId 请求 ✓
04:20:10.115 [INFO] CTPFactoryAvailable                              ← WPS 支持 ICustomTaskPaneConsumer ✓
04:20:11.579 [INFO] CreateCTP completed title=JYYJ助手               ← 任务窗格创建成功 ✓
04:20:11.580 [INFO] Task pane made visible
04:20:12.916 [INFO] WebView2 initialization completed success=True   ← WebView2 在 WPS 进程内正常 ✓
04:20:15.850 [INFO] LoadSkills reload-done count=13                  ← 完整初始化（技能 13 个）
```

- `LoadBehavior` 启动后保持 3（未被 WPS 翻转为 2 = 无加载失败）。
- 日志中 `builtIn=...execl1.0-wps...` 确认加载的是 WPS 副本产物，非主工作区残留。

## 4. 需求文档 5 个开放问题的答案

| # | 问题 | 答案 |
|---|------|------|
| 1 | et.exe 位数 | x86（32 位） |
| 2 | CustomTaskPane 是否触发 | **是**，CTPFactoryAvailable + CreateCTP 全链路成功 |
| 3 | WebView2 能否初始化 | 是，success=True，侧边栏资源正常伺服 |
| 4 | Ribbon 图标渲染效果 | XML 返回成功（resourceFound=True）；**透明底/灰底视觉效果需人工打开 WPS 确认**（唯一未闭环项） |
| 5 | KET.Application 连通性 | Dispatch OK，Version=12.0（harness 连接方式可行；探测经 pywin32，注意沙箱默认 python 无 pywin32，用系统 Python312） |

## 5. 对原计划的影响

- **M2 简化**：侧边栏形态 A（原生任务窗格）成立，浮动窗兜底降级为"仅备而不用"，M2 工作量显著下降。
- **M1**：注册配方已验证，剩余工作是把它产品化进 `register-auto.ps1`/`unregister-all.ps1`（探测到 WPS 时追加写 Kingsoft 键）与 `.iss`。
- **M4**：harness 用 `KET.Application` 连接（pywin32 或 C# Marshal GetActiveObject/CreateObject），连接方式已验证。

## 6. 遗留与风险

1. Ribbon 图标视觉（透明 Bitmap 灰底坑）待人工确认——启动 WPS 看一眼即可。
2. 本次仅验证个人版 12.1.0.26899 单点；专业版/教育版矩阵验证仍按 M4 计划执行。
3. WPS 关闭后 LoadBehavior 行为、异常退出后的恢复未压测。
4. 源码版本号漂移提醒：构建产物为 2.19.6.0，README 仍写 2.19.5.0，发版前按 SOP 同步。

## 7. 回滚

```
python D:\execl1.0\_dev_tmp\wps_register.py unregister
```

仅清理本次写入的 HKCU 键（Classes 3 组 + Kingsoft Addins/AddinsWL），不触碰 Excel 侧任何注册项。

---

## 附：M1 进展（2026-09-08，注册层产品化完成）

M0 配方已产品化进 WPS 副本安装器脚本（`D:\execl1.0-wps\tools\installer\`）：

- **register-auto.ps1**：新增 WPS 段——探测常见安装根下最高版本 `office6\et.exe` → 读 PE 头定位数 → 与 Excel 位数不同时按 WPS 视图补注册 COM 类（CodeBase 指向对应架构 `app\{arch}` 目录）→ Kingsoft `ET\Addins` 键 + `AddinsWL` 白名单双视图（64/32）写入；未装 WPS 时优雅跳过（仅注册 Excel）。`Register-ComClass` 参数化（Assembly/CodeBase 入参）。
- **unregister-all.ps1**：双视图追加清理 Kingsoft Addins 键与 AddinsWL 值。
- **闭环测试证据**（本机 Excel x64 + WPS x86 双宿主真实场景）：
  1. 注册 → 回读：view64 CodeBase=app\x64、view32 CodeBase=app\x86，Kingsoft 双视图 LB=3 + 白名单齐，Excel 键无损；
  2. 卸载 → `CLEAN-CHECK PASS`（双视图 COM/Addins/Kingsoft/白名单零残留）；
  3. 复装 → 重启 WPS 真机加载成功，日志断言 `builtIn=D:\execl1.0-wps\tools\installer\app\x86\skills`（确认加载自产品化注册的 app 目录，非 M0 手工路径）。
- **M1 退出条件达成**：本机个人版注册后自动加载 ✓。
- 脚本均保持 UTF-8 BOM（红线）；`.iss` 的 WPS 分发（双包双后缀、PkgSrc 路径）留待 M5 发版里程碑处理。

---

## 附：M2 进展（2026-09-08，宿主 UI 层完成）

- **HostKind 宿主检测**（`ExcelAiAddIn.cs`）：OnConnection 时按进程名判定宿主（`et`→WPS、`EXCEL*`→Excel），暴露 `HostKind` 静态属性 + 诊断日志 `Host detected kind=... process=...`，为 M3 宿主感知过滤（需求文档 D9）提供依据；Excel 主线行为零改动（N1）。
- **Ribbon IPictureDisp 核查**：`GetImage` 已经由 `OleCreatePictureIndirect`（fOwn=false + 记忆化缓存）返回 IPictureDisp——这正是规避"WPS 透明 Bitmap 灰底"的兼容路径，**免改**；WPS 真机多轮加载日志零 `GetImage failed`。
- **浮动窗兜底**：M0 判形态 A（原生任务窗格可用），按计划跳过。
- **验证**：x86/x64 双平台重建 0 警告 0 错误 → app 目录同步 → WPS 重启日志实测 `Host detected kind=WPS process=et`；无 Ribbon 错误。
- **新踩坑沉淀**（手册 6.13 方向）：WPS 开着时重建 x86 DLL 会被 `WPS Spreadsheets` 进程锁（MSB3027），与既有"每轮 harness 后清 EXCEL 进程"同构——WPS 侧同样要先关 et.exe 再构建。
