using System.Runtime.CompilerServices;

// 仅向单测程序集暴露内部纯函数（ConvertRefsInFormula/DetectDelimiter/TryParseDate 等），
// 让对抗性修复有可脱离 Excel 的回归测试钩子；对生产组装无副作用。
[assembly: InternalsVisibleTo("ExcelAiAssistant.Tests")]
// P3-D2 真机诊断工具：复刻 run_vba_snippet 内部阶段二分定位状态归零点（非生产组装，仅开发期使用）。
[assembly: InternalsVisibleTo("VerifyP3Diag")]