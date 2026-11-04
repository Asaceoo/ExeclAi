// 决定性实验：验证 VBA「模块级变量是否会被工程重编译重置」。
// 直接通过 VBE 代码模型操控模块，绕过 run_vba_snippet 的封装。
// 关键设计：读取过程 ReadGxP 在一开始就写入模块，此后只「运行」不「增删代码」，
// 避免读取动作本身触发重编译污染实验。
//  1. 模块 M：声明区 Public gX + SetX(写42) + GetX(空) + ReadGxP(把 gX 写入 A1)。
//  2. 运行 SetX → gX=42 → 读 A1 应为 42。
//  3. 对照组：不改代码仅运行 GetX → 读 A1（验证无编辑时保留）。
//  4. 过程区新增 GetX2（声明区保留）→ 触发重编译 → 读 A1（关键：声明保留是否仍被重置）。
//  5. 声明区新增 Public gY → 读 A1（声明被改，预期重置）。
using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace VerifyVbaStateSemantics
{
    internal static class Program
    {
        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("== VerifyVbaStateSemantics 模块级变量 vs 重编译 ==");

            var result = 1;
            var runner = new Thread(() => result = Run()) { IsBackground = true };
            runner.SetApartmentState(ApartmentState.STA);
            runner.Start();
            if (!runner.Join(240000))
            {
                runner.Abort();
                Console.WriteLine("FAIL  240s 超时");
                return 1;
            }

            return result;
        }

        private static int Run()
        {
            dynamic app = null;
            dynamic wb = null;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                wb = app.Workbooks.Add();

                dynamic vbProject = wb.VBProject;
                dynamic components = vbProject.VBComponents;

                dynamic m = components.Add(1);
                m.Name = "M";
                dynamic cm = m.CodeModule;
                cm.AddFromString(
                    "Public gX As Long\n" +
                    "Sub SetX()\n    gX = 42\nEnd Sub\n" +
                    "Sub GetX()\nEnd Sub\n" +
                    "Sub ReadGxP()\n    Range(\"A1\").Value = gX\nEnd Sub\n" +
                    "Sub GetX2()\nEnd Sub");

                dynamic ws = wb.Worksheets.Item[1];

                app.Run("SetX");
                app.Run("ReadGxP");
                Console.WriteLine("步骤1 SetX 后 gX=：  " + ReadCell(ws));

                app.Run("GetX");
                app.Run("ReadGxP");
                Console.WriteLine("步骤2(对照) 仅运行未改代码后 gX=：" + ReadCell(ws));

                // 过程区新增一行注释（不改声明区、不新增过程，改过程体也算改代码 → 重编译）。
                var cnt = Convert.ToInt32(cm.CountOfLines);
                cm.InsertLines(cnt + 1, "' 追加注释行触发重编译");
                app.Run("ReadGxP");
                Console.WriteLine("步骤3 过程区插入注释(声明区保留)后 gX=：" + ReadCell(ws));

                // 声明区新增变量。
                cm.InsertLines(2, "Public gY As Long");
                app.Run("ReadGxP");
                Console.WriteLine("步骤4 声明区新增变量后 gX=：" + ReadCell(ws));

                // 步骤5 跨模块实验：模块 N 声明区变化，是否重置模块 M 的 gX？
                dynamic n = components.Add(1);
                n.Name = "N";
                dynamic ncm = n.CodeModule;
                ncm.AddFromString("Public gN As Long\nSub SetN()\n    gN = 99\nEnd Sub\nSub ReadNP()\n    Range(\"A2\").Value = gN\nEnd Sub");
                app.Run("SetX");   // 重新设置 gX=42
                app.Run("SetN");
                app.Run("ReadGxP");
                app.Run("ReadNP");
                Console.WriteLine("步骤5 设置后 gX=" + ReadCell(ws) + " gN=" + ReadCell2(ws));

                // 改 N 的声明区（新增变量）→ 检查 M 的 gX 是否被重置。
                var ncnt = Convert.ToInt32(ncm.CountOfLines);
                ncm.InsertLines(2, "Public gN2 As Long");
                app.Run("ReadGxP");
                app.Run("ReadNP");
                Console.WriteLine("步骤6 N声明区变化后 gX=" + ReadCell(ws) + " gN=" + ReadCell2(ws));

                // 步骤7 过程区变化（N 追加注释，声明区保留）→ 检查 M 的 gX。
                var ncnt2 = Convert.ToInt32(ncm.CountOfLines);
                ncm.InsertLines(ncnt2 + 1, "' 注释");
                app.Run("ReadGxP");
                app.Run("ReadNP");
                Console.WriteLine("步骤7 N过程区注释后 gX=" + ReadCell(ws) + " gN=" + ReadCell2(ws));

                // 步骤8 破坏性重写实验（复刻 run_vba_snippet 的 ReplaceModuleCode：DeleteLines+AddFromString）：
                // 状态模块 S 声明 gS；scratch 模块 H 先有代码，被整清重写 → 检查 gS 是否存活。
                dynamic s = components.Add(1);
                s.Name = "S";
                dynamic scm = s.CodeModule;
                scm.AddFromString("Private gS As String\nSub SetS()\n    gS = \"abc\"\nEnd Sub\nSub ReadSP()\n    Range(\"A3\").Value = gS\nEnd Sub");
                app.Run("SetS");
                app.Run("ReadSP");
                Console.WriteLine("步骤8a S设置后 gS=" + ReadCell3(ws));

                dynamic h = components.Add(1);
                h.Name = "H";
                dynamic hcm = h.CodeModule;
                hcm.AddFromString("Option Explicit\nPublic gH As Long\nSub T1()\nEnd Sub");
                var hcnt = Convert.ToInt32(hcm.CountOfLines);
                hcm.DeleteLines(1, hcnt);
                hcm.AddFromString("Option Explicit\nPublic gH As Long\nSub T2()\nEnd Sub");
                app.Run("ReadSP");
                Console.WriteLine("步骤8b H破坏性重写后 gS=" + ReadCell3(ws) + "（关键）");

                // 再改 S 的声明区，确认「改 S 才重置 S」。
                var scnt = Convert.ToInt32(scm.CountOfLines);
                scm.InsertLines(2, "Private gS2 As String");
                app.Run("ReadSP");
                Console.WriteLine("步骤8c S声明区变化后 gS=" + ReadCell3(ws));

                // 步骤9 修复方向验证：H2 模块只增量替换「过程区」（保留声明区），S 的 gS 是否存活？
                // 模拟：S 重新 SetS 后，H 先「只删过程行、保留声明行」，再 Append 新过程。
                app.Run("SetS");
                app.Run("ReadSP");
                Console.WriteLine("步骤9a S重新设置后 gS=" + ReadCell3(ws));

                // 当前 H 内容是：声明行(1:Option Explicit, 2:Public gH) + T2 过程。声明区为第1-2行。
                var hTotal = Convert.ToInt32(hcm.CountOfLines);
                var hDecl = 2; // 声明区行数（Option Explicit + Public gH）
                if (hTotal > hDecl)
                {
                    hcm.DeleteLines(hDecl + 1, hTotal - hDecl); // 只删过程区
                }
                hcm.AddFromString("Sub T3()\nEnd Sub");
                app.Run("ReadSP");
                Console.WriteLine("步骤9b H仅过程区替换后 gS=" + ReadCell3(ws) + "（修复方向验证）");

                // ===== 决定性实验：定位真实流程里「仅替换过程区仍重置状态」的真正触发点 =====
                // 真实 run_vba_snippet：编辑草稿模块过程区后，Run 的是「被编辑模块自身」的探针过程（JYYJHelper.XaiProbe9F2E），
                // 而步骤9b 只 Run 了「未编辑模块 S」的 ReadSP。下面区分两个变量：Run 位置 / 编辑规模。

                // 步骤10：小编辑（删2行+加2行）+ Run「刚编辑模块 H 自身」的新过程 → 检查 S。
                app.Run("SetS"); // gS=abc
                var hTot10 = Convert.ToInt32(hcm.CountOfLines);
                if (hTot10 > 2)
                {
                    hcm.DeleteLines(3, hTot10 - 2); // 只删过程区，小编辑
                }
                hcm.AddFromString("Sub T4()\nEnd Sub");
                app.Run("H.T4"); // Run 刚编辑模块自身的过程
                app.Run("ReadSP");
                Console.WriteLine("步骤10 H小编辑+Run H自身过程后 gS=" + ReadCell3(ws) + "（关键：隔离Run位置）");

                // 步骤11：大编辑（仿真实辅助块规模 ~200 行）+ Run「刚编辑模块 H 自身」过程 → 检查 S。
                app.Run("SetS");
                var big = new System.Text.StringBuilder();
                for (var i = 0; i < 100; i++)
                {
                    big.Append("Sub Big" + i + "()\nEnd Sub\n");
                }
                var hTot11 = Convert.ToInt32(hcm.CountOfLines);
                if (hTot11 > 2)
                {
                    hcm.DeleteLines(3, hTot11 - 2); // 只删过程区
                }
                hcm.AddFromString(big.ToString()); // 大追加（~200 行）
                hcm.AddFromString("Sub T5()\nEnd Sub");
                app.Run("H.T5"); // Run 刚编辑模块自身过程
                app.Run("ReadSP");
                Console.WriteLine("步骤11 H大编辑(200行)+Run H自身过程后 gS=" + ReadCell3(ws) + "（关键：隔离编辑规模）");

                // 步骤12：大编辑但 Run「未编辑模块 S」的过程 → 检查 S。
                app.Run("SetS");
                var hTot12 = Convert.ToInt32(hcm.CountOfLines);
                if (hTot12 > 2)
                {
                    hcm.DeleteLines(3, hTot12 - 2); // 只删过程区
                }
                hcm.AddFromString(big.ToString());
                hcm.AddFromString("Sub T6()\nEnd Sub");
                app.Run("ReadSP"); // 仍 Run 未编辑模块 S 的过程
                Console.WriteLine("步骤12 H大编辑但Run S过程后 gS=" + ReadCell3(ws));

                // ===== 决定性实验2：大块 DeleteLines 重置他模块状态的触发边界 =====
                // VerifyP3Diag bisect 实证：生产每轮 ReplaceModuleCode 单次大删 ~398 行过程区 → gXaiState 归零。
                // 下面用「全新载荷模块」隔离变量：状态模块 S 置位后，仅对载荷模块做不同删除模式，
                // 判断触发整工程重编译的是「单次删除行数」还是「删除总量」，据此定修复方向（分块删 vs 常驻辅助模块）。
                // 每次测试用全新模块（创建=纯增量，不做大删），避免历史大删污染 S 的状态基线。

                // 步骤13：单次删除 400 行（复现生产）→ 预期重置。
                app.Run("SetS");
                var p1 = CreatePayload(components, "P1", 400);
                var p1t = Convert.ToInt32(p1.CountOfLines);
                p1.DeleteLines(3, p1t - 2); // 单次删 400
                app.Run("ReadSP");
                Console.WriteLine("步骤13 单次删400行后 gS=" + ReadCell3(ws) + "（复现生产，预期重置）");
                Release((object)p1);

                // 步骤14：4 × 100 分块删除 → 若存活则分块删是修复方向。
                app.Run("SetS");
                var p2 = CreatePayload(components, "P2", 400);
                for (var ci = 0; ci < 4; ci++)
                {
                    var t2 = Convert.ToInt32(p2.CountOfLines);
                    p2.DeleteLines(3, Math.Min(100, t2 - 2));
                }
                app.Run("ReadSP");
                Console.WriteLine("步骤14 4×100分块删后 gS=" + ReadCell3(ws) + "（分块删是否规避重置）");
                Release((object)p2);

                // 步骤15：40 × 10 分块删除 → 更细粒度是否更安全。
                app.Run("SetS");
                var p3 = CreatePayload(components, "P3", 400);
                for (var ci = 0; ci < 40; ci++)
                {
                    var t3 = Convert.ToInt32(p3.CountOfLines);
                    p3.DeleteLines(3, Math.Min(10, t3 - 2));
                }
                app.Run("ReadSP");
                Console.WriteLine("步骤15 40×10分块删后 gS=" + ReadCell3(ws) + "（细粒度分块删）");
                Release((object)p3);

                // 步骤16：单次删除 100 行 → 找单次删除的安全上界。
                app.Run("SetS");
                var p4 = CreatePayload(components, "P4", 100);
                var p4t = Convert.ToInt32(p4.CountOfLines);
                p4.DeleteLines(3, p4t - 2); // 单次删 100
                app.Run("ReadSP");
                Console.WriteLine("步骤16 单次删100行后 gS=" + ReadCell3(ws) + "（单次删安全上界探测）");
                Release((object)p4);

                // 步骤17：单次删除 50 行。
                app.Run("SetS");
                var p5 = CreatePayload(components, "P5", 50);
                var p5t = Convert.ToInt32(p5.CountOfLines);
                p5.DeleteLines(3, p5t - 2); // 单次删 50
                app.Run("ReadSP");
                Console.WriteLine("步骤17 单次删50行后 gS=" + ReadCell3(ws) + "（单次删安全上界探测）");
                Release((object)p5);

                // ===== 决定性实验3：区分「删除规模」与「跨模块引用」两个变量 =====
                // 步骤13 载荷模块(无跨模块引用)单次删 400 行 → S 存活；而 bisect 里 JYYJHelper(辅助块跨模块引用
                // XaiStateRaw/XaiStateRawSet) 大删即归零。下面 P6 的每个 Sub 都调用 S 的 ReadSP（跨模块引用），
                // 单次删 400 行 → 若归零则证明触发源是「含跨模块引用的模块被大删」，而非删除规模本身。

                // 步骤18：含跨模块引用的模块单次删 400 行。
                var p6 = components.Add(1);
                p6.Name = "P6";
                var p6cm = p6.CodeModule;
                var sb6 = new System.Text.StringBuilder();
                for (var i = 0; i < 200; i++)
                {
                    sb6.Append("Sub X" + i + "()\n    ReadSP\nEnd Sub\n"); // 400 行，每 Sub 都引用 S.ReadSP
                }
                p6cm.AddFromString(sb6.ToString());
                app.Run("SetS");
                var p6t = Convert.ToInt32(p6cm.CountOfLines);
                p6cm.DeleteLines(1, p6t); // 单次删 400
                app.Run("ReadSP");
                Console.WriteLine("步骤18 含跨模块引用模块单次删400后 gS=" + ReadCell3(ws) + "（跨模块引用是否触发整工程重编译）");
                Release((object)p6cm);
                Release((object)p6);

                // 步骤19（已移除）：手工行号替换片段区被 VBE InsertLines 行为绊住产生悬空语句，
                // 且其「辅助块常驻+小改片段」概念将由生产实现端到端验证，此处不再用脆弱的行号算术复刻。

                // ===== 决定性实验4：彻底隔离三个变量（跨模块引用 / 模块清空 / 分块删除）=====
                // 步骤18 的 P6 无声明区且「删到空」，与步骤13 的 P1（有声明、保留声明）同时差异了「跨模块引用」
                // 与「是否删空」。下面逐一隔离：
                //  步骤20：跨模块引用 + 保留声明区 + 单次删400 → 若重置，则「跨模块引用」是触发源（排除删空因素）。
                //  步骤21：无跨模块引用 + 删空（删全部402行）→ 若存活，则「删空」本身不触发。
                //  步骤22：跨模块引用 + 4×100分块删 → 若存活，则分块删是修复方向；若重置则需「运行时调用规避编译期跨模块引用」。

                // 步骤20：P8 跨模块引用 + 声明区保留 + 单次删 400。
                var p8 = components.Add(1);
                p8.Name = "P8";
                var p8cm = p8.CodeModule;
                p8cm.AddFromString("Private gP8 As String\n' ==== p8 ===="); // 声明 2 行，保留
                var sb8 = new System.Text.StringBuilder();
                for (var i = 0; i < 200; i++)
                {
                    sb8.Append("Sub X" + i + "()\n    ReadSP\nEnd Sub\n");
                }
                p8cm.AddFromString(sb8.ToString());
                RunApp(app, "SetS");
                var p8t = Convert.ToInt32(p8cm.CountOfLines);
                p8cm.DeleteLines(3, p8t - 2); // 保留声明 2 行，单次删 400
                RunApp(app, "ReadSP");
                Console.WriteLine("步骤20 跨引用+保留声明+单次删400后 gS=" + ReadCell3(ws) + "（隔离：跨模块引用）");
                Release((object)p8cm);
                Release((object)p8);

                // 步骤21：P9 无跨引用 + 删空（删全部）。
                var p9 = components.Add(1);
                p9.Name = "P9";
                var p9cm = p9.CodeModule;
                p9cm.AddFromString("Private gP9 As String\n' ==== p9 ====");
                var sb9 = new System.Text.StringBuilder();
                for (var i = 0; i < 200; i++)
                {
                    sb9.Append("Sub Y" + i + "()\nEnd Sub\n");
                }
                p9cm.AddFromString(sb9.ToString());
                RunApp(app, "SetS");
                var p9t = Convert.ToInt32(p9cm.CountOfLines);
                p9cm.DeleteLines(1, p9t); // 删空
                RunApp(app, "ReadSP");
                Console.WriteLine("步骤21 无跨引用+删空后 gS=" + ReadCell3(ws) + "（隔离：模块清空）");
                Release((object)p9cm);
                Release((object)p9);

                // ===== 决定性实验5：验证修复方向——Application.Run 运行时调度是否规避整工程重编译 =====
                // 步骤18/20 的 P6/P8 直接直呼 S.ReadSP（编译期跨模块引用）→ 大删即重置。
                // P11 改为 Application.Run "S.ReadSP"（运行时调度，字符串宏名无编译期绑定）→ 若大删后 gS 存活，
                // 则生产修复 = 辅助块内对 JYYJXaiState 的访问全部改 Application.Run 运行时调度。
                // 注意：本组实验必须先于步骤22 运行——步骤22 的分块删会搞崩整个工程（后续 SetS 宏不可用）。

                // 步骤23：P11 运行时调度跨引用 + 单次删 400（保留声明）。
                var p11 = components.Add(1);
                p11.Name = "P11";
                var p11cm = p11.CodeModule;
                p11cm.AddFromString("Private gP11 As String\n' ==== p11 ====");
                var sb11 = new System.Text.StringBuilder();
                for (var i = 0; i < 200; i++)
                {
                    sb11.Append("Sub W" + i + "()\n    Application.Run \"S.ReadSP\"\nEnd Sub\n");
                }
                p11cm.AddFromString(sb11.ToString());
                RunApp(app, "SetS");
                var p11t = Convert.ToInt32(p11cm.CountOfLines);
                p11cm.DeleteLines(3, p11t - 2); // 单次删 400，保留声明
                RunApp(app, "ReadSP");
                Console.WriteLine("步骤23 运行时调度跨引用+单次删400后 gS=" + ReadCell3(ws) + "（Application.Run 是否规避整工程重编译）");
                Release((object)p11cm);
                Release((object)p11);

                // ===== 决定性实验6：隔离「可执行语句」变量（不含任何跨模块/运行时引用）=====
                // 步骤23 颠覆了「编译期跨引用触发重置」假设：运行时调度也重置。
                // P13 过程体只有 Debug.Print 1（无跨模块引用、无 Application.Run）→ 单次删 400：
                //  若重置 → 触发源 = 「删除含可执行语句的过程」，任何大删都不可行，只能换载体（如隐藏表/新鲜模块）。
                //  若存活 → 触发源仍与跨模块有关，需继续细分。
                // 注意：必须置于分块删实验之前——分块删（步骤24/22）会把工程搞崩，后续步骤不可运行。

                // 步骤25：P13 非空体（仅 Debug.Print）+ 无任何引用 + 单次删 400。
                var p13 = components.Add(1);
                p13.Name = "P13";
                var p13cm = p13.CodeModule;
                p13cm.AddFromString("Private gP13 As String\n' ==== p13 ====");
                var sb13 = new System.Text.StringBuilder();
                for (var i = 0; i < 200; i++)
                {
                    sb13.Append("Sub U" + i + "()\n    Debug.Print 1\nEnd Sub\n");
                }
                p13cm.AddFromString(sb13.ToString());
                RunApp(app, "SetS");
                var p13t = Convert.ToInt32(p13cm.CountOfLines);
                p13cm.DeleteLines(3, p13t - 2); // 单次删 400，保留声明
                RunApp(app, "ReadSP");
                Console.WriteLine("步骤25 非空体(仅Debug.Print)+无引用+单次删400后 gS=" + ReadCell3(ws) + "（隔离：可执行语句）");
                Release((object)p13cm);
                Release((object)p13);

                // ===== 决定性实验7：寻找不触发重置的「代码替换方式」=====
                // 已证：删除含可执行语句的代码块 → 整工程重编译 → 重置所有模块级变量。
                // 若生产必须替换片段，则只能找「不删除」的路径。测两个候选：
                //  步骤26：整模块 components.Remove（非 DeleteLines）是否触发重置。
                //  步骤27：纯追加可执行代码（不删除）是否触发重置。

                // 步骤26：P14 含可执行代码 → SetS → components.Remove(P14) → 读 gS。
                var p14 = components.Add(1);
                p14.Name = "P14";
                var p14cm = p14.CodeModule;
                p14cm.AddFromString("Private gP14 As String\n' ==== p14 ====\nSub T14()\n    Debug.Print 14\nEnd Sub");
                RunApp(app, "SetS");
                components.Remove(p14);
                RunApp(app, "ReadSP");
                Console.WriteLine("步骤26 components.Remove含可执行体模块后 gS=" + ReadCell3(ws) + "（整模块移除是否规避重置）");
                Release((object)p14cm);

                // 步骤27：P15 已有可执行代码，纯追加（不删除）更多可执行代码 → SetS → 读 gS。
                var p15 = components.Add(1);
                p15.Name = "P15";
                var p15cm = p15.CodeModule;
                p15cm.AddFromString("Private gP15 As String\n' ==== p15 ====\nSub T15a()\n    Debug.Print 15\nEnd Sub");
                RunApp(app, "SetS");
                var sb15 = new System.Text.StringBuilder();
                for (var i = 0; i < 100; i++)
                {
                    sb15.Append("Sub Add" + i + "()\n    Debug.Print 15\nEnd Sub\n");
                }
                p15cm.AddFromString(sb15.ToString()); // 纯追加 200 行可执行代码，不删除
                RunApp(app, "ReadSP");
                Console.WriteLine("步骤27 纯追加可执行代码(200行不删)后 gS=" + ReadCell3(ws) + "（纯追加是否规避重置）");
                Release((object)p15cm);
                Release((object)p15);

                // ===== 决定性实验8：完整生产周期复刻（新鲜运行模块 + components.Remove + 直呼状态访问器）=====
                // 步骤26 已证 components.Remove 安全；这里验证完整周期：独立状态模块 S2，
                // 每轮「新建运行模块(引用S2访问器)→设状态→移除→再新建」，检查 gS2 跨轮存活。

                // 步骤28：S2 状态模块 + 轮1运行模块 → 设状态 → 移除 → 轮2运行模块 → 读状态。
                var s2 = components.Add(1);
                s2.Name = "S2";
                var s2cm = s2.CodeModule;
                s2cm.AddFromString(
                    "Private gS2 As String\n" +
                    "Public Function XaiStateRaw2() As String\n    XaiStateRaw2 = gS2\nEnd Function\n" +
                    "Public Sub XaiStateRawSet2(v As String)\n    gS2 = v\nEnd Sub\n" +
                    "Sub ReadS2P()\n    Range(\"A5\").Value = gS2\nEnd Sub");

                var q = components.Add(1);
                q.Name = "JYYJRun_X";
                var qcm = q.CodeModule;
                qcm.AddFromString(
                    "Public Function GetState() As String\n    GetState = XaiStateRaw2\nEnd Function\n" +
                    "Public Sub SetState()\n    XaiStateRawSet2 \"xyz\"\nEnd Sub");
                RunApp(app, "ReadS2P");
                Console.WriteLine("步骤28a 轮1创建运行模块(引用状态模块)后 gS2=" + ReadCell5(ws));
                RunApp(app, "SetState");
                RunApp(app, "ReadS2P");
                Console.WriteLine("步骤28b 轮1运行模块SetState后 gS2=" + ReadCell5(ws));
                components.Remove(q);
                Release((object)qcm);
                RunApp(app, "ReadS2P");
                Console.WriteLine("步骤28c 轮1移除运行模块后 gS2=" + ReadCell5(ws));

                var q2 = components.Add(1);
                q2.Name = "JYYJRun_Y";
                var q2cm = q2.CodeModule;
                q2cm.AddFromString(
                    "Public Function GetState() As String\n    GetState = XaiStateRaw2\nEnd Function");
                RunApp(app, "ReadS2P");
                Console.WriteLine("步骤28d 轮2创建运行模块后 gS2=" + ReadCell5(ws) + "（跨调用状态保留判据）");
                components.Remove(q2);
                Release((object)q2cm);
                Release((object)s2cm);
                Release((object)s2);

                // 步骤24：运行时调度 + 4×100 分块删（对照，验证分块删是否与运行时调度叠加更稳）。
                var p12 = components.Add(1);
                p12.Name = "P12";
                var p12cm = p12.CodeModule;
                p12cm.AddFromString("Private gP12 As String\n' ==== p12 ====");
                var sb12 = new System.Text.StringBuilder();
                for (var i = 0; i < 200; i++)
                {
                    sb12.Append("Sub V" + i + "()\n    Application.Run \"S.ReadSP\"\nEnd Sub\n");
                }
                p12cm.AddFromString(sb12.ToString());
                RunApp(app, "SetS");
                for (var ci = 0; ci < 4; ci++)
                {
                    var t12 = Convert.ToInt32(p12cm.CountOfLines);
                    p12cm.DeleteLines(3, Math.Min(100, t12 - 2));
                }
                var p12read = SafeReadSP(app, ws);
                Console.WriteLine("步骤24 运行时调度+4×100分块删后 gS=" + p12read + "（叠加对照）");
                Release((object)p12cm);
                Release((object)p12);

                // 步骤22：P10 编译期跨引用 + 4×100 分块删（置于最后：分块删会搞崩工程，后续步骤不可运行）。
                var p10 = components.Add(1);
                p10.Name = "P10";
                var p10cm = p10.CodeModule;
                p10cm.AddFromString("Private gP10 As String\n' ==== p10 ====");
                var sb10 = new System.Text.StringBuilder();
                for (var i = 0; i < 200; i++)
                {
                    sb10.Append("Sub Z" + i + "()\n    ReadSP\nEnd Sub\n");
                }
                p10cm.AddFromString(sb10.ToString());
                RunApp(app, "SetS");
                for (var ci = 0; ci < 4; ci++)
                {
                    var t10 = Convert.ToInt32(p10cm.CountOfLines);
                    p10cm.DeleteLines(3, Math.Min(100, t10 - 2));
                }
                var p10read = SafeReadSP(app, ws);
                Console.WriteLine("步骤22 跨引用+4×100分块删后 gS=" + p10read + "（分块删是否规避跨引用重置）");
                Release((object)p10cm);
                Release((object)p10);

                Release(hcm);
                Release(h);
                Release(scm);
                Release(s);
                Release(ncm);
                Release(n);
                Release(cm);
                Release(m);
                Release(components);
                Release(vbProject);
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL  异常: " + ex.Message);
                return 1;
            }
            finally
            {
                try { if (wb != null) wb.Close(false); } catch { }
                try { if (app != null) { app.Quit(); } } catch { }
                try { if (app != null && Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app); } catch { }
            }
        }

        private static string ReadCell(dynamic ws)
        {
            try
            {
                var v = ws.Cells.Item[1, 1].Value2;
                return v == null ? "(null)" : Convert.ToString(v);
            }
            catch (Exception ex)
            {
                return "(读异常 " + ex.Message + ")";
            }
        }

        private static string ReadCell2(dynamic ws)
        {
            try
            {
                var v = ws.Cells.Item[2, 1].Value2;
                return v == null ? "(null)" : Convert.ToString(v);
            }
            catch (Exception ex)
            {
                return "(读异常 " + ex.Message + ")";
            }
        }

        private static string ReadCell3(dynamic ws)
        {
            try
            {
                var v = ws.Cells.Item[3, 1].Value2;
                return v == null ? "(null)" : Convert.ToString(v);
            }
            catch (Exception ex)
            {
                return "(读异常 " + ex.Message + ")";
            }
        }

        private static string ReadCell5(dynamic ws)
        {
            try
            {
                var v = ws.Cells.Item[5, 1].Value2;
                return v == null ? "(null)" : Convert.ToString(v);
            }
            catch (Exception ex)
            {
                return "(读异常 " + ex.Message + ")";
            }
        }

        /// <summary>运行宏并重试（模拟生产 RunVbaProbeWithRetry 的瞬态自愈）：大编辑后首次 Run 常报「宏不可用」，
        /// 等待 600ms 重试至多 3 次，避免瞬态编译未就绪误判为失败。</summary>
        private static object RunApp(dynamic app, string macro)
        {
            for (var i = 0; i < 3; i++)
            {
                try
                {
                    return app.Run(macro);
                }
                catch (Exception)
                {
                    if (i == 2)
                    {
                        throw;
                    }
                    Thread.Sleep(600);
                }
            }
            return null;
        }

        /// <summary>安全读取 S 的 gS（A3）：把「宏不可用」等瞬态失败降级为字符串输出，避免步骤级失败中止后续实验。</summary>
        private static string SafeReadSP(dynamic app, dynamic ws)
        {
            try
            {
                RunApp(app, "ReadSP");
                return ReadCell3(ws);
            }
            catch (Exception ex)
            {
                return "(读异常 " + ex.Message + ")";
            }
        }

        /// <summary>创建全新载荷模块：2 行声明 + procLines 行纯过程代码（无 Option Explicit 干扰），
        /// 供大删阈值实验使用。创建是纯增量（Add 模块 + AddFromString），不做任何删除，不污染 S 的状态基线。</summary>
        private static dynamic CreatePayload(dynamic components, string name, int procLines)
        {
            var p = components.Add(1);
            p.Name = name;
            var pcm = p.CodeModule;
            pcm.AddFromString("Private gP As String\n' ==== payload ====");
            var sb = new System.Text.StringBuilder();
            var subs = procLines / 2;
            for (var i = 0; i < subs; i++)
            {
                sb.Append("Sub D" + i + "()\nEnd Sub\n");
            }

            pcm.AddFromString(sb.ToString());
            return pcm;
        }

        private static void Release(object o)
        {
            if (o != null && Marshal.IsComObject(o))
            {
                try { Marshal.FinalReleaseComObject(o); } catch { }
            }
        }
    }
}
