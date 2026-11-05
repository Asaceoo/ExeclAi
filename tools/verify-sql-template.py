# -*- coding: utf-8 -*-
"""层1 真机闭环：从线上 index.html 提取 SQL 模板真实文本，起真实 Excel + xlwings 跑通并断言。
   已安装 Python3.12.4 + xlwings 0.36.17 + sqlite3（与侧边栏 RunPythonExcel worker 同环境）。
   PS: 本脚本为临时验证工具，非发布物。
"""
import re, json, sys

HTML = r"D:\execl1.0\src\ExcelAiAssistant.Addin\Sidebar\index.html"

def extract_sql_template():
    src = open(HTML, encoding="utf-8").read()
    start = src.index("const sqlBody = [")
    head = src[start:]
    end = head.index('\n    ].join("\\n");')
    arr_src = head[:end].split("[", 1)[1]
    tokens = re.findall(r'"((?:[^"\\]|\\.)*)"', arr_src)
    lines = [json.loads('"' + t + '"') for t in tokens]
    return "\n".join(lines)

def main():
    template = extract_sql_template()
    assert "sqlite3.connect(':memory:')" in template, "未提取到 SQLite 模板"
    print("[extract] 已提取模板行数:", len(template.splitlines()))

    import xlwings as xw
    app = xw.App(visible=False)
    failures = []
    passed = 0
    try:
        wb = app.books.add()
        sht = wb.sheets[0]
        sht.range("A1").value = [["品名", "金额"], ["苹果", 500], ["梨", 1500], ["香蕉", 200], ["橙子", 3000]]

        sel = "'Sheet1'!$A$1:$B$5"
        marker = "# XAI_SEL: " + sel + "\nsel_range = \"" + sel + "\""
        ns = {}
        captured = {}
        def xai_result(d): captured.update(d)
        ns["xai_result"] = xai_result
        exec(marker + "\n" + template, ns)
        got = captured
        try:
            assert got.get("count") == 4, f"count={got.get('count')} 期望4"
            assert got.get("table_columns") == ["品名", "金额"], f"table_columns={got.get('table_columns')}"
            assert got.get("columns") == ["c0", "c1"], f"columns={got.get('columns')}"
            assert isinstance(got.get("rows"), list) and len(got["rows"]) == 4, "rows 长度不符"
            passed += 1
            print("[PASS] 真机 SQL 模板闭环：count=%s table_columns=%s rows=%s" % (got.get("count"), got.get("table_columns"), got.get("rows")))
        except AssertionError as ex:
            failures.append("SQL 模板断言失败: %s" % ex)

        import sqlite3
        rows = sht.range("A1:B5").value
        names = [str(c) for c in rows[0]]
        data = rows[1:]
        cols = ["c%d" % i for i in range(len(names))]
        conn = sqlite3.connect(":memory:")
        conn.execute("CREATE TABLE t (" + ", ".join('"%s"' % c for c in cols) + ")")
        conn.executemany("INSERT INTO t VALUES (" + ",".join("?" * len(cols)) + ")", [list(r) for r in data])
        conn.commit()
        cur = conn.execute("SELECT * FROM t WHERE c1 > 1000 LIMIT 20")
        filtered = [list(r) for r in cur.fetchall()]
        conn.close()
        assert [r[0] for r in filtered] == ["梨", "橙子"], "WHERE c1>1000 应命中 梨/橙子，实际 %s" % filtered
        passed += 1
        print("[PASS] 真机 WHERE 过滤：%s" % filtered)
    except Exception as ex:
        failures.append("真机执行异常: %s" % ex)
    finally:
        try: app.quit()
        except Exception: pass

    print("\n== SQL 模板真机结果：通过 %d，失败 %d ==" % (passed, len(failures)))
    for f in failures:
        print("[FAIL]", f)
    return 0 if not failures else 1

if __name__ == "__main__":
    sys.exit(main())