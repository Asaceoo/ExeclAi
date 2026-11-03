# -*- coding: utf-8 -*-
"""gen_manual_html.py —— 从版本化 md 手册重生成同风格 HTML（M17.6）。

复用现有 HTML 的 <style>（从旧 HTML 提取），正文按 md 子集渲染：
h2 分节（显式闭合）、表格、``` 代码块、> 引用块、有序/无序列表、段落、
行内 **加粗**/`代码`/[链接](url)。
用法：python gen_manual_html.py <手册md路径> <旧HTML路径（取CSS）> <输出HTML路径>
"""
import html
import io
import re
import sys


def load_style(old_html_path):
    text = io.open(old_html_path, "r", encoding="utf-8-sig").read()
    m = re.search(r"<style>(.*?)</style>", text, re.S)
    if not m:
        sys.exit("旧 HTML 中未找到 <style>：%s" % old_html_path)
    return m.group(1)


def inline(s):
    s = html.escape(s, quote=False)
    s = re.sub(r"`([^`]+)`", r"<code>\1</code>", s)
    s = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", s)
    s = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r'<a href="\2" target="_blank">\1</a>', s)
    return s


def render(md_path):
    lines = io.open(md_path, "r", encoding="utf-8-sig").read().splitlines()
    title = "JYYJ助手 手册"
    blocks = []        # [(h2_title, id, [html lines])]
    cur_title, cur_id, cur = None, None, None
    sec_idx = 0

    def new_section(t):
        nonlocal cur_title, cur_id, cur, sec_idx
        cur_title, cur = t, []
        cur_id = "sec%d" % sec_idx
        sec_idx += 1
        blocks.append((t, cur_id, cur))

    def flush_list(state):
        kind, items = state
        if kind and items:
            tag = "ol" if kind == "ol" else "ul"
            cur.append("<%s>" % tag)
            cur.extend("<li>%s</li>" % inline(it) for it in items)
            cur.append("</%s>" % tag)
        state[0], state[1] = None, []

    def flush_table(buf):
        if not buf:
            return
        rows = [r for r in buf if not re.match(r"^\|[\s:\-|]+\|$", r)]
        grid = [[c.strip() for c in r.strip().strip("|").split("|")] for r in rows]
        cur.append("<table><thead><tr>%s</tr></thead>" % "".join("<th>%s</th>" % inline(c) for c in grid[0]))
        cur.append("<tbody>")
        for row in grid[1:]:
            cur.append("<tr>%s</tr>" % "".join("<td>%s</td>" % inline(c) for c in row))
        cur.append("</tbody></table>")
        del buf[:]

    def flush_quote(buf):
        if cur is None:
            del buf[:]
            return
        if not buf:
            return
        paras, para = [], []
        for q in buf:
            if q.strip():
                para.append(inline(q.strip()))
            elif para:
                paras.append(para)
                para = []
        if para:
            paras.append(para)
        cur.append("<blockquote>%s</blockquote>"
                   % "".join("<p>%s</p>" % "<br>".join(p) for p in paras))
        del buf[:]

    lst = [None, []]     # [kind, items]
    tbl, quote = [], []
    pre = [False, []]    # [in_code, code_lines]

    for raw in lines:
        stripped = raw.strip()
        if pre[0]:
            if stripped.startswith("```"):
                cur.append("<pre><code>%s</code></pre>" % html.escape("\n".join(pre[1]), quote=False))
                pre = [False, []]
            else:
                pre[1].append(raw)
            continue

        if stripped.startswith("|"):
            flush_quote(quote)
            flush_list(lst)
            tbl.append(stripped)
            continue
        flush_table(tbl)

        if stripped.startswith(">"):
            flush_list(lst)
            quote.append(stripped[1:])
            continue
        flush_quote(quote)

        if stripped.startswith("```"):
            flush_list(lst)
            pre = [True, []]
            continue

        m = re.match(r"^(#{1,4})\s+(.*)$", stripped)
        if m:
            flush_list(lst)
            level, text = len(m.group(1)), m.group(2).strip()
            if level == 1:
                title = text.replace("`", "")
            elif level == 2:
                new_section(text)
            else:
                cur.append("<h%d>%s</h%d>" % (level, inline(text), level))
            continue

        if stripped == "---" or not stripped:
            flush_list(lst)
            continue

        m = re.match(r"^(\d+)[.、)]\s+(.*)$", stripped)
        if m:
            flush_quote(quote)
            if lst[0] and lst[0] != "ol":
                flush_list(lst)
            lst[0] = "ol"
            lst[1].append(m.group(2))
            continue
        if stripped.startswith("- "):
            flush_quote(quote)
            if lst[0] and lst[0] != "ul":
                flush_list(lst)
            lst[0] = "ul"
            lst[1].append(stripped[2:])
            continue

        flush_list(lst)
        if cur is not None:
            cur.append("<p>%s</p>" % inline(stripped))

    flush_table(tbl)
    flush_quote(quote)
    flush_list(lst)
    if pre[0]:
        cur.append("<pre><code>%s</code></pre>" % html.escape("\n".join(pre[1]), quote=False))
    return title, blocks


TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<style>{style}</style>
</head>
<body>
<header><div class="wrap">
  <div class="logo">JY</div>
  <div><h1>{title}</h1><div class="sub">{subtitle}</div></div>
</div></header>
<nav><div class="wrap">{nav}</div></nav>
<div class="wrap">
{body}
</div>
<footer><div class="wrap">JYYJ助手 · {subtitle}</div></footer>
</body>
</html>
"""


def main():
    md_path, old_html, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
    style = load_style(old_html)
    title, blocks = render(md_path)
    subtitle = ""
    for line in io.open(md_path, "r", encoding="utf-8-sig").read().splitlines():
        if line.strip().startswith(">"):
            subtitle = html.escape(line.strip().lstrip(">").strip(), quote=False)
            break
    body = "\n".join(
        '<section id="%s"><h2>%s</h2>\n%s\n</section>' % (sid, inline(t), "\n".join(ls))
        for t, sid, ls in blocks)
    nav = "".join('<a href="#%s">%s</a>' % (sid, html.escape(t, quote=False)) for t, sid, _ in blocks)
    page = TEMPLATE.format(title=html.escape(title, quote=False), style=style,
                           subtitle=subtitle, nav=nav, body=body)
    io.open(out_path, "w", encoding="utf-8").write(page)
    print("[html] %s（%.1f KB，%d 节）" % (out_path, len(page.encode("utf-8")) / 1024, len(blocks)))


if __name__ == "__main__":
    main()
