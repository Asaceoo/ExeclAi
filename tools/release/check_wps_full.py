# -*- coding: utf-8 -*-
"""check_wps_full.py —— WPS full 巡检结果基线校验（M17 沉淀，清单约束兜底自动化）。

用法：
    python check_wps_full.py --log <full日志路径>            # 校验（CI/门禁用）
    python check_wps_full.py --log <full日志路径> --update   # 重新生成基线（需人工 diff 后提交）

日志格式（WpsToolVerify full 输出，UTF-8/GBK 均可，自动探测）：
    TOOL_COUNT=144
    RESULT\t工具名\tOK|FAIL\t错误码(OK 为 -)\t消息或JSON

基线文件：本脚本同目录 expected_wps_full.tsv（tool/状态/错误码/次数/理由，tab 分隔）。
退出码：0=与基线一致；1=有差异（新增/缺失/次数不符）；2=日志无法解析。
"""
import argparse
import collections
import io
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
BASELINE = os.path.join(HERE, "expected_wps_full.tsv")

# 工具注册总数（与 WpsToolVerify TOOL_COUNT 对账）。新增/移除工具时必须同步本值、
# 基线与两份豁免清单——这是有意设置的"变更卡点"，不要改成自动探测。
# 2026-09-10：144→147（M-D 批次新增 create_combo_chart/create_doughnut/create_truncated_column
# 三个 WPS 侧可见图表工具；7 个看板工具经 D11 过滤不出现在 WPS）。
# 2026-09-10 二修：对账口径改为【排除 ct_ 动态注册】——用户沉淀工具（ct_ 前缀）由
# harness 巡检样例动态写入用户工具库（如 ct_hf_review_tool），计数随用户库浮动，
# 不属于代码注册口径；对账时两侧同步剔除 ct_ 计数（见 TOOL_COUNT 记录行后的说明）。
EXPECTED_TOOL_COUNT = 169  # M28：168 + 1（verify_file 文件核验）；此前 167+1 save_workbook；M26：156 + 11（L2 元素层 6 + import_csv 1 + 草稿模式 4）  # 代码注册数（2026-09-18 +build_formula/eval_formula/diagnose_formula/rewrite_formulas 公式构造与验证闭环；2026-09-15 gate 实测口径：+web_search（联网搜索，ReadOnly）；2026-09-14：+analyze_trend_forecast、+analyze_hypothesis_test、+analyze_cluster、+analyze_regression、+analyze_text_variants、+add_insight_card（ThirdPhase 注册）；其余历史增量见下）


def read_text_any_encoding(path):
    for enc in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            with io.open(path, "r", encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    raise SystemExit("无法解码日志文件：%s" % path)


def parse_log(text):
    """返回 (Counter[(tool, status, code)], tool_count or None)。行数异常直接报错。"""
    counter = collections.Counter()
    tool_count = None
    result_lines = 0
    for line in text.splitlines():
        line = line.strip()
        if line.startswith("TOOL_COUNT="):
            tool_count = int(line.split("=", 1)[1])
            continue
        if not line.startswith("RESULT\t"):
            continue
        parts = line.split("\t")
        if len(parts) < 4:
            raise SystemExit("RESULT 行字段数异常（期望 >=4）：\n%s" % line)
        tool, status, code = parts[1], parts[2], parts[3]
        if status not in ("OK", "FAIL"):
            raise SystemExit("RESULT 行状态异常（期望 OK|FAIL）：\n%s" % line)
        counter[(tool, status, code)] += 1
        result_lines += 1
    if result_lines == 0:
        raise SystemExit("日志中没有任何 RESULT 行——确认传入的是 full 模式日志。")
    return counter, tool_count


def parse_baseline(path):
    counter = collections.Counter()
    with io.open(path, "r", encoding="utf-8-sig") as f:
        for lineno, raw in enumerate(f, 1):
            line = raw.rstrip("\n")
            if not line or line.startswith("#"):
                continue
            parts = line.split("\t")
            if len(parts) != 5:
                raise SystemExit("基线第 %d 行格式异常（期望 5 列）：%s" % (lineno, line))
            tool, status, code, count = parts[0], parts[1], parts[2], int(parts[3])
            counter[(tool, status, code)] += count
    return counter


def diff(log_counter, base_counter):
    """返回 (问题列表, 是否有差异)。新增=日志有基线无；缺失=基线有日志无；次数不符。"""
    problems = []
    keys = set(log_counter) | set(base_counter)
    for key in sorted(keys):
        got, want = log_counter.get(key, 0), base_counter.get(key, 0)
        tool, status, code = key
        if want == 0:
            problems.append("新增结果（未登记，先人工定性再决定 --update 或修复）：%s\t%s\t%s ×%d" % (tool, status, code, got))
        elif got == 0:
            problems.append("基线条目未出现（场景被删或行为变化？）：%s\t%s\t%s（期望 ×%d）" % (tool, status, code, want))
        elif got != want:
            problems.append("次数变化：%s\t%s\t%s 期望 ×%d 实际 ×%d" % (tool, status, code, want, got))
    return problems, bool(problems)


def write_baseline(counter, path):
    with io.open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("# WPS full 巡检结果基线（由 check_wps_full.py --update 生成；reason 列需人工补写并评审）\n")
        f.write("# 格式（tab 分隔）：tool\t状态(OK|FAIL)\t错误码(OK 为 -)\t次数\t理由\n")
        for (tool, status, code), count in sorted(counter.items()):
            f.write("%s\t%s\t%s\t%d\tTODO-人工补写理由\n" % (tool, status, code, count))
    print("基线已重写：%s（请人工 diff 并补写 reason）" % path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--log", required=True, help="WpsToolVerify full 日志路径")
    ap.add_argument("--update", action="store_true", help="用日志结果重写基线（慎用，需评审）")
    ap.add_argument("--min-ok", type=int, default=140, help="OK 总数下限（防止大面积静默丢失）")
    ap.add_argument("--expect-tools", type=int, default=EXPECTED_TOOL_COUNT,
                    help="期望工具注册数（对账 TOOL_COUNT）")
    args = ap.parse_args()

    text = read_text_any_encoding(args.log)
    log_counter, tool_count = parse_log(text)
    ok_total = sum(v for (t, s, c), v in log_counter.items() if s == "OK")
    # 2026-09-10 二修：TOOL_COUNT 含 ct_ 前缀用户沉淀工具的动态注册（随用户库浮动），
    # 与代码注册口径混合会让卡点被巡检样例误击穿。对账剔除方式：记录行在枚举名单中
    # 统计 ct_ 数量并从两侧同步扣除（TOOL_COUNT-ctN vs EXPECTED）。
    ct_count = 0
    for line in text.splitlines():
        if line.startswith("ct_") and line.rstrip("\r\n").count("\t") >= 1:
            ct_count += 1
    if tool_count is not None and ct_count:
        print("对账口径：TOOL_COUNT=%s 剔除 ct_ 动态注册 %d 个 → 代码注册 %d"
              % (tool_count, ct_count, tool_count - ct_count))
        tool_count -= ct_count
    ok_total = ok_total  # ok_total 保持原语义（含 ct 用例的 RESULT）
    print("日志解析：RESULT 条目 %d 组，OK 总数 %d，代码注册 TOOL_COUNT=%s（剔除 ct_ %d 个）"
          % (len(log_counter), ok_total, tool_count, ct_count))

    if args.update:
        write_baseline(log_counter, BASELINE)
        return 0

    base_counter = parse_baseline(BASELINE)
    problems, changed = diff(log_counter, base_counter)

    if ok_total < args.min_ok:
        problems.append("OK 总数 %d 低于下限 %d——疑似大面积场景未执行" % (ok_total, args.min_ok))
    if tool_count is not None and tool_count != args.expect_tools:
        problems.append("TOOL_COUNT=%s ≠ %d——工具注册数变化，需同步基线与豁免清单"
                        % (tool_count, args.expect_tools))

    # M17.8 修复：TOOL_COUNT/min_ok 卡点此前只 append 不判定， problems 非空也放行（卡点失效）。
    # 现在基线 diff、OK 下限、工具数对账任一命中即拦截。
    if changed or problems:
        print("== 校验失败：与基线存在 %d 处差异 ==" % len(problems))
        for p in problems:
            print("  [DIFF] " + p)
        print("处置：新 FAIL 先人工定性；有意的行为变更走 --update 重生成基线并逐行评审；"
              "工具数变化须同步 EXPECTED_TOOL_COUNT。")
        return 1

    print("== 校验通过：结果分布与基线完全一致（%d 组 / OK %d）==" % (len(log_counter), ok_total))
    return 0


if __name__ == "__main__":
    sys.exit(main())
