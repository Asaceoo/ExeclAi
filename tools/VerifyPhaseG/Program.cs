using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Configuration;

namespace VerifyPhaseG
{
    /// <summary>
    /// 自动冒烟（第 G 批）：P5-A 会话与记忆增强的纯文件级验证，不依赖 Excel COM。
    /// 断言：
    ///  1) 会话重命名（Rename）持久化并可跨实例读取；
    ///  2) 会话归档（SetArchived）标志往返；
    ///  3) 会话全文检索（Search）按标题/内容命中并给出片段；
    ///  4) 长期记忆（MemoryStore）偏好/备注持久化、最近摘要注入文本生成；
    ///  5) SidebarBridge 暴露 P5-A 桥接方法（反射校验签名）。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static readonly List<string> Failures = new List<string>();

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            try
            {
                RunAll();
            }
            catch (Exception ex)
            {
                _failed++;
                Failures.Add("HARNESS: " + Safe(ex));
                Console.WriteLine("HARNESS EXCEPTION: " + Safe(ex));
            }

            Console.WriteLine();
            Console.WriteLine("== PASSED=" + _passed + " FAILED=" + _failed + " ==");
            if (Failures.Count > 0)
            {
                foreach (var f in Failures)
                {
                    Console.WriteLine("FAILED: " + f);
                }

                return 1;
            }

            Console.WriteLine("ALL PHASE-G SESSION/MEMORY CHECKS OK");
            return 0;
        }

        private static void RunAll()
        {
            var root = Path.Combine(Path.GetTempPath(), "verifyphaseg-" + Guid.NewGuid().ToString("N"));
            var sessionDir = Path.Combine(root, "sessions");
            var memoryPath = Path.Combine(root, "memory.json");
            const string book = "BookG.xlsx";

            try
            {
                CheckSessionRename(sessionDir, book);
                CheckSessionArchive(sessionDir, book);
                CheckSessionSearch(sessionDir, book);
                CheckMemoryStore(memoryPath);
                CheckBridgeMethods();
            }
            finally
            {
                try { if (Directory.Exists(root)) { Directory.Delete(root, true); } }
                catch { }
            }
        }

        private static void CheckSessionRename(string dir, string book)
        {
            var store = new SessionStore(dir);
            var id = store.Create(book);
            store.Save(book, id, new[] { ChatMessage.User("旧标题内容") }, "旧标题");

            Assert("rename-updates-title", store.Rename(book, id, "新标题"),
                "Rename 返回 true");
            Assert("rename-persists", new SessionStore(dir).Title(book, id) == "新标题",
                "重命名跨实例持久化");
            Assert("rename-rejects-blank", !store.Rename(book, id, "   "),
                "空白标题被拒绝");
        }

        private static void CheckSessionArchive(string dir, string book)
        {
            var store = new SessionStore(dir);
            var id = store.Create(book);

            Assert("archive-default-false", !store.List(book)[0].Archived,
                "新会话默认未归档");
            Assert("archive-set-true", store.SetArchived(book, id, true),
                "归档成功");
            Assert("archive-roundtrip", store.List(book)[0].Archived,
                "归档标志可读回");
            Assert("archive-set-false", !store.SetArchived(book, id, false) || !store.List(book)[0].Archived,
                "取消归档成功");
        }

        private static void CheckSessionSearch(string dir, string book)
        {
            var store = new SessionStore(dir);
            var a = store.Create(book);
            store.Save(book, a, new[] { ChatMessage.User("把销售数据做成汇总表") }, "销售汇总");
            var b = store.Create(book);
            store.Save(book, b, new[] { ChatMessage.User("分析库存周转率") }, "库存分析");

            var titleHits = store.Search(book, "销售");
            Assert("search-title-hit", titleHits.Count == 1 && titleHits[0].SessionId == a && titleHits[0].TitleMatched,
                "标题命中 1 条且标记 TitleMatched");

            var contentHits = store.Search(book, "周转率");
            Assert("search-content-hit", contentHits.Count == 1 && contentHits[0].SessionId == b && !contentHits[0].TitleMatched,
                "内容命中 1 条且非标题匹配");
            Assert("search-snippet", contentHits.Count == 1 && contentHits[0].Snippet != null && contentHits[0].Snippet.Contains("周转率"),
                "命中片段包含关键词");

            Assert("search-no-match", store.Search(book, "不存在").Count == 0,
                "无命中返回空");
            Assert("search-blank", store.Search(book, "  ").Count == 0,
                "空白关键词返回空");
        }

        private static void CheckMemoryStore(string path)
        {
            var store = new MemoryStore(path);
            const string book = "BookG.xlsx";
            Assert("memory-defaults-empty", store.Preferences == string.Empty && store.Notes == string.Empty,
                "默认偏好/备注为空");

            store.Preferences = "图表默认中文标题；写入优先公式。";
            store.Notes = "用户偏好简洁回复。";
            var reloaded = new MemoryStore(path);
            Assert("memory-persists", reloaded.Preferences == "图表默认中文标题；写入优先公式。" && reloaded.Notes == "用户偏好简洁回复。",
                "偏好/备注跨实例持久化");

            for (var i = 0; i < 25; i++) { store.AddRecentSummary(book, "摘要 " + i); }
            Assert("memory-cap-20", store.RecentSummaries(book).Count == 20,
                "最近摘要上限 20 条");
            Assert("memory-keep-latest", store.RecentSummaries(book)[0] == "摘要 5" && store.RecentSummaries(book)[19] == "摘要 24",
                "保留最新 20 条");

            store.AddRecentSummary("OtherBook.xlsx", "别的簿摘要");
            var text = store.BuildInjectionText(book, recentSummaryCount: 2);
            Assert("memory-injection", text.Contains("用户长期偏好") && text.Contains("摘要 23") && text.Contains("摘要 24") && !text.Contains("摘要 0"),
                "注入文本含偏好与最近 2 条摘要");
            Assert("memory-per-book", !text.Contains("别的簿摘要"),
                "摘要按工作簿隔离");
        }

        private static void CheckBridgeMethods()
        {
            var bridgeType = Type.GetType("ExcelAiAssistant.Addin.SidebarBridge, ExcelAiAssistant.Addin");
            if (bridgeType == null)
            {
                Fail("bridge-type-found", "未加载 SidebarBridge 类型");
                return;
            }

            var expected = new[] { "RenameSession", "ArchiveSession", "SearchSessions", "ListArchivedSessions", "GetMemory", "SetMemory" };
            foreach (var name in expected)
            {
                var method = bridgeType.GetMethod(name);
                Assert("bridge-" + name, method != null, "桥接方法 " + name + " 存在");
            }
        }

        private static void Assert(string name, bool condition, string detail)
        {
            if (condition)
            {
                _passed++;
                Console.WriteLine("PASS " + name + " — " + detail);
            }
            else
            {
                _failed++;
                Failures.Add(name + ": " + detail);
                Console.WriteLine("FAIL " + name + " — " + detail);
            }
        }

        private static void Fail(string name, string detail)
        {
            _failed++;
            Failures.Add(name + ": " + detail);
            Console.WriteLine("FAIL " + name + " — " + detail);
        }

        private static string Safe(Exception ex)
        {
            return ex == null ? string.Empty : ex.Message;
        }
    }
}
