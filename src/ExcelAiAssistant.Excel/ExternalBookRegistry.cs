using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 跨工作簿只读引用（P5-C1）：以只读句柄打开外部工作簿，供 read_external_range 直读。
    /// 句柄容器由 SidebarBridge 持有、跨轮注入 ExcelToolHost，保证多轮复用。
    /// 安全边界：仅读；外部句柄 Workbook 均以 ReadOnly 打开，写仍限当前活动工作簿。
    /// </summary>
    public sealed class ExternalBookRegistry
    {
        private readonly dynamic application;
        private readonly Dictionary<string, object> handles =
            new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> pathToToken =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>标记句柄对应的工作簿是用户自己打开的（借用实例）：Close 时只解除登记、不执行 Close（M16）。</summary>
        private readonly HashSet<string> userOwned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public ExternalBookRegistry(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public sealed class HandleInfo
        {
            public string Token { get; set; }
            public string Name { get; set; }
            public string Path { get; set; }
            public List<string> Sheets { get; set; } = new List<string>();
        }

        /// <summary>以只读句柄打开指定路径工作簿，注册并返回句柄信息；成功后尽量恢复原活动工作簿。</summary>
        public HandleInfo Open(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("请提供工作簿路径。", nameof(path));
            }

            var full = Path.GetFullPath(path);
            // 与 verify_file / 导出链统一：只允许本地固定盘，拒绝 UNC/网络盘/设备路径。
            // 必须先于 File.Exists——否则检查存在性就会先向远端发起 SMB/NTLM 外连。
            full = ExcelAiAssistant.Core.Security.PathGuard.GuardOutputPath(full);
            if (!File.Exists(full))
            {
                throw new FileNotFoundException("工作簿不存在：" + full, full);
            }

            // 重复打开同一文件：复用既有句柄，避免句柄泄漏。
            if (pathToToken.TryGetValue(full, out var reused) && handles.ContainsKey(reused))
            {
                return Describe(reused);
            }

            dynamic original = null;
            try { original = application.ActiveWorkbook; } catch { original = null; }

            // M16 全量审查修复：若该文件已被用户在宿主中打开，COM 的 Workbooks.Open 会返回
            // 用户已打开的实例——此时登记到 handles 后，close_external_book 会把用户工作簿
            // 一并关闭（未保存数据丢失）。因此先按全路径比对，命中则直接复用用户实例：
            // 登记一个"借用"句柄并标记 userOwned，Close 时只解除登记、不执行 Close。
            foreach (dynamic wb in application.Workbooks)
            {
                string openPath = null;
                try { openPath = Convert.ToString(wb.FullName); } catch { openPath = null; }
                var isMatch = false;
                if (!string.IsNullOrEmpty(openPath))
                {
                    try { isMatch = string.Equals(openPath, full, StringComparison.OrdinalIgnoreCase); }
                    catch { isMatch = false; }
                }

                if (isMatch)
                {
                    var name = Path.GetFileNameWithoutExtension(full);
                    var token = MakeToken(name);
                    handles[token] = wb;
                    pathToToken[full] = token;
                    userOwned.Add(token);

                    TryRestoreActive(original);
                    return Describe(token);
                }

                // 枚举得到的 Workbook RCW 不做强制释放（与 ExcelContextService 的共享 RCW 策略一致，
                // 强制释放有"RCW 分离"风险；出循环后交由 GC 回收）。
            }

            var book = application.Workbooks.Open(
                full,
                UpdateLinks: 0,
                ReadOnly: true,
                AddToMru: false,
                IgnoreReadOnlyRecommended: true);

            var openedName = Path.GetFileNameWithoutExtension(full);
            var openedToken = MakeToken(openedName);
            handles[openedToken] = book;
            pathToToken[full] = openedToken;

            TryRestoreActive(original);
            return Describe(openedToken);
        }

        public bool TryGet(string token, out object workbook)
        {
            return handles.TryGetValue(token, out workbook);
        }

        public bool Close(string token)
        {
            if (!handles.TryGetValue(token, out var obj) || obj == null)
            {
                return false;
            }

            // M16 全量审查：借用（用户已打开）的实例只解除登记，不 Close——否则会把用户
            // 未保存的工作簿无提示关闭。本工具自己打开的实例才 Close，并释放 RCW。
            if (!userOwned.Remove(token))
            {
                dynamic book = obj;
                try { book.Close(false); } catch { }
                try
                {
                    if (System.Runtime.InteropServices.Marshal.IsComObject(obj))
                    {
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
                    }
                }
                catch { }
            }

            handles.Remove(token);
            var pathAlias = pathToToken.FirstOrDefault(p => string.Equals(p.Value, token, StringComparison.Ordinal));
            if (pathAlias.Key != null)
            {
                pathToToken.Remove(pathAlias.Key);
            }

            return true;
        }

        public HandleInfo Describe(string token)
        {
            if (!handles.TryGetValue(token, out var obj) || obj == null)
            {
                throw new KeyNotFoundException("句柄不存在或已关闭：" + token);
            }

            dynamic book = obj;
            string full = null;
            try { full = Convert.ToString(book.FullName); } catch { full = null; }
            var sheets = new List<string>();
            try
            {
                foreach (var sheet in book.Worksheets)
                {
                    sheets.Add(Convert.ToString(sheet.Name));
                }
            }
            catch { }

            return new HandleInfo
            {
                Token = token,
                Name = full != null ? Path.GetFileNameWithoutExtension(full) : token,
                Path = Convert.ToString(book.Path),
                Sheets = sheets
            };
        }

        public IReadOnlyList<HandleInfo> List()
        {
            var result = new List<HandleInfo>();
            foreach (var pair in handles)
            {
                if (pair.Value is string)
                {
                    continue;
                }

                try { result.Add(Describe(pair.Key)); } catch { }
            }

            return result;
        }

        private string MakeToken(string fileName)
        {
            var safe = string.Concat(fileName.Select(ch => char.IsLetterOrDigit(ch) ? ch : '_'));
            if (string.IsNullOrWhiteSpace(safe))
            {
                safe = "ext";
            }

            var token = "ext_" + safe;
            var n = 1;
            while (handles.ContainsKey(token))
            {
                token = "ext_" + safe + "_" + n++;
            }

            return token;
        }

        private void TryRestoreActive(dynamic original)
        {
            if (original == null)
            {
                return;
            }

            try
            {
                var sheets = original.Worksheets;
                var first = sheets[1];
                first.Activate();
                // Windows.Activate 使原工作簿最前。
                try { original.Activate(); } catch { }
            }
            catch { }
        }
    }
}