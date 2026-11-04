# -*- coding: utf-8 -*-
"""release.py —— 发版自动化（M17 沉淀，替代每轮手工 sed 三处版本号）。

版本号单一来源 = 仓库根 Directory.Build.props 的 <Version>。
本脚本是唯一允许修改版本号的入口；csproj 继承 props，iss 由 ISCC /DAppVersion 注入，
README 由 bump 子命令同步。禁止手工 sed 三处版本号（历史红线：M15 增量编译旧戳事故）。

用法：
    python release.py bump --ver 3.8.0.0 [--dry-run]   # 升版：props + README（唯一改版本入口）
    python release.py pack [--from 3.7.0.0] [--dry-run] # 打包：构建→组装→zip→戳断言→ISCC
                                                         # 版本号自动读 props；--from 指定旧版
                                                         # （说明文件来源），缺省自动取 dist 下最高旧版
前置：dist 下已有上一版本的便携目录与安装说明 txt（首版之后必然满足）。
"""
import argparse
import glob
import io
import os
import re
import shutil
import subprocess
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
PROPS = os.path.join(ROOT, "Directory.Build.props")
README = os.path.join(ROOT, "README.md")
ISS = os.path.join(ROOT, "tools", "installer", "JYYJ助手-Setup-wps.iss")
DIST = os.path.join(ROOT, "dist")
DOCS = os.path.join(ROOT, "docs")
# 四份手册（技术/用户 × md/html）随 bump 同步：文件名 + 版本戳 + 指针引用（M17.6 补自动化缺口）。
# PDF 不在自动范围（内容需人工重生成，bump 只提醒）。
MANUALS = ("JYYJ助手技术手册", "JYYJ助手用户手册")
MANUAL_EXTS = (".md", ".html")
MANUAL_POINTER_FILES = (
    os.path.join(ROOT, "CLAUDE.md"),
    os.path.join(DOCS, "AI交接文档.md"),
    README,
)
BUILDS = {
    "x64": os.path.join(ROOT, "src", "ExcelAiAssistant.Addin", "bin", "x64", "Release", "net48", "win-x64"),
    "x86": os.path.join(ROOT, "src", "ExcelAiAssistant.Addin", "bin", "x86", "Release", "net48", "win-x86"),
}
VER_RE = re.compile(r"^\d+\.\d+\.\d+\.\d+$")
VER_TAG_RE = re.compile(r"<(Version|InformationalVersion)>([\d.]+)</\1>")
README_TAG_RE = re.compile(r"（v\d+\.\d+\.\d+\.\d+）")

# 沙箱/CI 环境缺失变量兜底（本机 Git Bash 会话曾因缺这些导致 dotnet 全挂）
ENV_FIX = {
    "APPDATA": r"C:\Users\iamly\AppData\Roaming",
    "LOCALAPPDATA": r"C:\Users\iamly\AppData\Local",
    "PROGRAMDATA": r"C:\ProgramData",
    "ALLUSERSPROFILE": r"C:\ProgramData",
    "SystemRoot": r"C:\WINDOWS",
    "windir": r"C:\WINDOWS",
    "COMSPEC": r"C:\WINDOWS\system32\cmd.exe",
    "ProgramFiles": r"C:\Program Files",
    "ProgramW6432": r"C:\Program Files",
    "ProgramFiles(x86)": r"C:\Program Files (x86)",
    # 裸环境（沙箱/CI）缺失时 MSBuild 会试图写 C:\WINDOWS\MSBuildTemp / NuGet 挂在
    # "未设置 USERPROFILE"——真实踩坑（M17.4 变异测试实证），一并兜底。
    "TMP": r"C:\Users\iamly\AppData\Local\Temp",
    "TEMP": r"C:\Users\iamly\AppData\Local\Temp",
    "USERPROFILE": r"C:\Users\iamly",
    "HOMEDRIVE": "C:",
    "HOMEPATH": r"\Users\iamly",
}


def run(cmd, **kw):
    env = dict(os.environ)
    for k, v in ENV_FIX.items():
        env.setdefault(k, v)
    print("[run] " + " ".join(cmd))
    r = subprocess.run(cmd, env=env, cwd=ROOT, **kw)
    if r.returncode != 0:
        sys.exit("命令失败（exit %d）：%s" % (r.returncode, " ".join(cmd)))
    return r


def read_version():
    with open(PROPS, "r", encoding="utf-8-sig") as f:
        m = VER_TAG_RE.search(f.read())
    if not m:
        sys.exit("Directory.Build.props 中未找到 <Version>——版本号单一来源被破坏。")
    return m.group(2)


def ver_key(v):
    return tuple(int(x) for x in v.split("."))


def find_iscc():
    for p in (r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
              r"C:\Program Files\Inno Setup 6\ISCC.exe"):
        if os.path.isfile(p):
            return p
    sys.exit("未找到 ISCC.exe（Inno Setup 6）。")


def sync_manuals(old_ver, new_ver):
    """四份手册（技术/用户 × md/html）版本同步：改名 + 版本戳 + 指针引用，并断言零残留。

    版本戳替换规则：裸版本号（前面不是 v/数字/字母）才替换——v 前缀历史条目不改动。
    HTML 正文与 md 正文按约定人工同步，本函数只保证版本标识一致（M17.6）。
    """
    stamp_re = re.compile(r"(?<![\wv])%s" % re.escape(old_ver))
    failures = []
    for base in MANUALS:
        for ext in MANUAL_EXTS:
            src = os.path.join(DOCS, "%s-%s%s" % (base, old_ver, ext))
            dst = os.path.join(DOCS, "%s-%s%s" % (base, new_ver, ext))
            if not os.path.isfile(src):
                if not os.path.isfile(dst):
                    failures.append("手册缺失且无法同步：%s" % src)
                continue
            text = io.open(src, "r", encoding="utf-8-sig").read()
            text, n = stamp_re.subn(new_ver, text)
            io.open(dst, "w", encoding="utf-8-sig").write(text)
            os.remove(src)
            if n == 0 and not os.path.isfile(dst):
                failures.append("版本戳未命中：%s" % src)
            print("[manual] %s-%s%s → %s（戳替换 %d 处）" % (base, old_ver, ext, new_ver, n))
    for p in MANUAL_POINTER_FILES:
        if not os.path.isfile(p):
            failures.append("指针文件缺失：%s" % p)
            continue
        text = io.open(p, "r", encoding="utf-8-sig").read()
        text, n = stamp_re.subn(new_ver, text)
        io.open(p, "w", encoding="utf-8-sig").write(text)
        print("[manual] 指针同步 %s（%d 处）" % (os.path.relpath(p, ROOT), n))
    # 安装说明模板（打包单一来源）内的"适用版本"戳一并同步
    note_tpl = os.path.join(ROOT, "tools", "installer", "安装说明.txt")
    if not os.path.isfile(note_tpl):
        failures.append("安装说明模板缺失：%s（打包单一来源，M17.7 起 required）" % note_tpl)
    else:
        text = io.open(note_tpl, "r", encoding="utf-8-sig").read()
        text, n = stamp_re.subn(new_ver, text)
        io.open(note_tpl, "w", encoding="utf-8-sig").write(text)
        print("[manual] 安装说明模板（%d 处）" % n)
    # 断言：新手册在位、新版本戳在位、旧裸版本零残留
    for base in MANUALS:
        for ext in MANUAL_EXTS:
            dst = os.path.join(DOCS, "%s-%s%s" % (base, new_ver, ext))
            if not os.path.isfile(dst):
                failures.append("同步后手册缺失：%s" % dst)
                continue
            text = io.open(dst, "r", encoding="utf-8-sig").read()
            if text.count(new_ver) < 1:
                failures.append("%s 缺新版本戳 %s" % (dst, new_ver))
            if stamp_re.search(text):
                failures.append("%s 残留旧版本戳 %s" % (dst, old_ver))
    if failures:
        sys.exit("手册版本同步失败：\n  " + "\n  ".join(failures))
    pdf = os.path.join(DOCS, "JYYJ助手用户手册-%s.pdf" % old_ver)
    if os.path.isfile(pdf):
        print("[manual] 提醒：PDF 手册不在自动同步范围（%s），需人工重生成。" % os.path.basename(pdf))


def cmd_bump(args):
    if not VER_RE.match(args.ver):
        sys.exit("版本号格式非法：%s（期望 x.y.z.w）" % args.ver)
    old = read_version()
    if ver_key(args.ver) <= ver_key(old):
        sys.exit("新版本 %s 必须大于当前 %s" % (args.ver, old))
    print("升版：%s → %s" % (old, args.ver))
    if args.dry_run:
        print("[dry-run] 将改写：%s（Version/InformationalVersion）、%s（第 3 行 v 标记）、四份手册（改名+版本戳）与指针引用" % (PROPS, README))
        return 0

    with open(PROPS, "r", encoding="utf-8-sig") as f:
        props = f.read()
    props = VER_TAG_RE.sub(lambda m: "<%s>%s</%s>" % (m.group(1), args.ver, m.group(1)), props)
    with open(PROPS, "w", encoding="utf-8-sig") as f:
        f.write(props)

    with open(README, "r", encoding="utf-8") as f:
        readme = f.read()
    readme, n = README_TAG_RE.subn("（v%s）" % args.ver, readme, count=1)
    if n != 1:
        sys.exit("README 未命中版本标记（（vX.Y.Z.W）），请人工检查。")
    with open(README, "w", encoding="utf-8") as f:
        f.write(readme)

    sync_manuals(old, args.ver)
    sync_npm_package(args.ver)

    print("完成。下一步：python release.py pack（版本号自动读取 %s）" % os.path.basename(PROPS))
    return 0


def build_app_dir(arch, ver, old_ver, old_dir, new_dir):
    """组装便携目录：复用旧目录非 app 文件（说明等），app/ 从构建输出重建。"""
    app_new = os.path.join(new_dir, "app")
    if os.path.exists(app_new):
        shutil.rmtree(app_new)
    os.makedirs(app_new)
    if not os.path.isdir(new_dir):
        if not old_dir or not os.path.isdir(old_dir):
            sys.exit("新目录不存在且找不到旧版目录（--from %s）：%s" % (old_ver, new_dir))
        shutil.copytree(old_dir, new_dir)
        if os.path.exists(app_new):
            shutil.rmtree(app_new)
        os.makedirs(app_new)
    for name in sorted(os.listdir(BUILDS[arch])):
        src = os.path.join(BUILDS[arch], name)
        if name == "runtimes" or name.endswith(".pdb"):
            continue
        dst = os.path.join(app_new, name)
        if os.path.isdir(src):
            shutil.copytree(src, dst)
        else:
            shutil.copy2(src, dst)


def make_zip(dir_path, flat=False):
    # flat=True：arcname 以 dir_path 自身为根（zip 内平铺）——MCP server 独立包用，
    # 与 npm 外壳 locate 的 cacheDir 根候选对齐（M27R3：原默认带一层目录名，解压后 locate 必不中）。
    zip_path = dir_path + ".zip"
    if os.path.exists(zip_path):
        os.remove(zip_path)
    arc_root = dir_path if flat else os.path.dirname(dir_path)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for base, _dirs, files in os.walk(dir_path):
            for fn in files:
                full = os.path.join(base, fn)
                zf.write(full, os.path.relpath(full, arc_root))
    return zip_path


def published_asset_name(dist_basename):
    """dist 文件名 -> GitHub Release 发布资产名（与 publish-github.ps1 改名表同源）。"""
    return dist_basename.replace("JYYJ助手-", "JYYJ-").replace("单文件便携", "portable-single")


def write_sha256_manifest(ver, files):
    """M28：生成 SHA256SUMS-<ver>.txt 供 npm 外壳安装前校验（下载链完整性闭环）。
    行为：逐文件流式 sha256；跳过不存在项；按 pypi 风格 "hash  filename" 输出。"""
    import hashlib
    lines = []
    mapping = {}
    for f in files:
        if not f or not os.path.isfile(f):
            continue
        digest = hashlib.sha256()
        with open(f, "rb") as fh:
            for chunk in iter(lambda: fh.read(1024 * 1024), b""):
                digest.update(chunk)
        published = published_asset_name(os.path.basename(f))
        if not re.match(r"^[A-Za-z0-9._-]+$", published):
            sys.exit("[sha256] 发布资产名含非法字符：" + published)
        lines.append("%s  %s" % (digest.hexdigest(), published))
        mapping[published] = f
    if not lines:
        sys.exit("[sha256] 没有可校验的产物，拒绝生成空清单")
    out = os.path.join(DIST, "SHA256SUMS-%s.txt" % ver)
    with open(out, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n")
    print("[sha256] %s（%d 个资产，键=发布名）" % (os.path.basename(out), len(lines)))
    return out, mapping


def assert_manifest_matches_publish(ver, manifest_path, mapping):
    """M28 审查：清单键必须是发布脚本负责上传的发布名，且发布脚本必须包含清单上传项。"""
    if not any(name.startswith("JYYJ-%s-mcp-server-x64" % ver) for name in mapping):
        sys.exit("[sha256] 清单缺少 mcp-server 独立包（npm install 必需资产）")
    publish_script = os.path.join(ROOT, "tools", "release", "publish-github.ps1")
    if os.path.isfile(publish_script):
        text = io.open(publish_script, "r", encoding="utf-8-sig").read()
        if ("SHA256SUMS-%s.txt" % ver) not in text and "SHA256SUMS-$ver.txt" not in text:
            sys.exit("[sha256] publish-github.ps1 未包含 SHA256SUMS 清单上传项——npm install 将 fail-closed")
    for name in mapping:
        if not name.startswith("JYYJ-"):
            sys.exit("[sha256] 清单键不是发布名：" + name)


def assert_npm_version(ver):
    """M28 审查（M17.1 单一来源）：npm 外壳 package.json 版本必须与 props 一致。"""
    pkg = os.path.join(ROOT, "tools", "npm-jyyj-mcp", "package.json")
    if not os.path.isfile(pkg):
        return
    text = io.open(pkg, "r", encoding="utf-8").read()
    marker = '"version"'
    i = text.find(marker)
    if i < 0:
        sys.exit("[npm] package.json 缺 version 字段")
    colon = text.find(":", i + len(marker))
    q1 = text.find('"', colon + 1)
    q2 = text.find('"', q1 + 1)
    current = text[q1 + 1:q2]
    if current != ver:
        sys.exit("[npm] package.json 版本 %s != props %s（bump 未同步？）" % (current, ver))


def sync_npm_package(ver):
    """M28：bump 同步 npm 外壳版本（package.json + README 引用），防 npx 拉旧 tag。"""
    pkg = os.path.join(ROOT, "tools", "npm-jyyj-mcp", "package.json")
    if os.path.isfile(pkg):
        text = io.open(pkg, "r", encoding="utf-8").read()
        marker = '"version"'
        i = text.find(marker)
        if i < 0:
            sys.exit("[npm] package.json 缺 version 字段")
        colon = text.find(":", i + len(marker))
        q1 = text.find('"', colon + 1)
        q2 = text.find('"', q1 + 1)
        text = text[:q1 + 1] + ver + text[q2:]
        io.open(pkg, "w", encoding="utf-8", newline="").write(text)
        print("[npm] package.json version -> %s" % ver)
    readme = os.path.join(ROOT, "tools", "npm-jyyj-mcp", "README.md")
    if os.path.isfile(readme):
        rtext = io.open(readme, "r", encoding="utf-8-sig").read()
        key = "GitHub Release v"
        i = rtext.find(key)
        if i >= 0:
            j = i + len(key)
            k = j
            while k < len(rtext) and (rtext[k].isdigit() or rtext[k] == "."):
                k += 1
            rtext = rtext[:j] + ver + rtext[k:]
            io.open(readme, "w", encoding="utf-8-sig", newline="").write(rtext)
            print("[npm] README 版本引用同步")


def assert_version_stamps(ver, old_ver):
    """双架构版本戳断言：新版本戳在位、旧版本零残留、守卫符号在位（M15 红线的自动化）。"""
    import codecs
    failures = []
    for arch in ("x64", "x86"):
        addin = os.path.join(DIST, "JYYJ助手-%s-wps-%s" % (ver, arch), "app", "ExcelAiAssistant.Addin.dll")
        excel = os.path.join(DIST, "JYYJ助手-%s-wps-%s" % (ver, arch), "app", "ExcelAiAssistant.Excel.dll")
        data = open(addin, "rb").read()
        v_new = data.count(ver.encode("utf-16-le"))
        v_old = data.count(old_ver.encode("utf-16-le")) if old_ver else 0
        guard = open(excel, "rb").read().count("SHEET_PROTECTED".encode("utf-16-le"))
        print("[stamp] %s Addin %s戳=%d 残留%s=%d | Excel.dll SHEET_PROTECTED戳=%d" % (arch, ver, v_new, old_ver, v_old, guard))
        if v_new < 3:
            failures.append("%s Addin.dll 新版本戳不足（%d < 3）——疑似旧 DLL 或增量编译" % (arch, v_new))
        if v_old != 0:
            failures.append("%s Addin.dll 残留旧版本戳 %s ×%d" % (arch, old_ver, v_old))
        if guard < 2:
            failures.append("%s Excel.dll 缺 SHEET_PROTECTED 符号——守卫未编译进去？" % arch)
    if failures:
        sys.exit("版本戳断言失败：\n  " + "\n  ".join(failures))


def auto_find_old_dir(ver, arch):
    """dist 下找最高旧版本的便携目录（排除当前版本）。"""
    best = None
    for p in glob.glob(os.path.join(DIST, "JYYJ助手-*-wps-" + arch)):
        m = re.search(r"JYYJ助手-([\d.]+)-wps-", os.path.basename(p))
        if not m or m.group(1) == ver:
            continue
        if VER_RE.match(m.group(1)) and (best is None or ver_key(m.group(1)) > ver_key(best[0])):
            best = (m.group(1), p)
    return best


def cmd_pack(args):
    ver = read_version()
    old_ver = args.from_version
    print("打包版本：%s（来源 Directory.Build.props）；旧版参考：%s" % (ver, old_ver or "自动探测"))
    if args.dry_run:
        print("[dry-run] 将执行：双架构 build --no-incremental → 组装 dist/JYYJ助手-%s-wps-{x64,x86} → "
              "安装说明 txt 替换 → zip → 版本戳断言 → ISCC /DAppVersion=%s" % (ver, ver))
        return 0

    # 1) 双架构强制全量重建（M15 红线：升版后增量编译不重生 DLL）
    for arch in ("x64", "x86"):
        run(["dotnet", "build", os.path.join("src", "ExcelAiAssistant.Addin", "ExcelAiAssistant.Addin.csproj"),
             "-c", "Release", "-p:Platform=" + arch, "--no-incremental", "--nologo", "-v", "q"])

    # 2) 组装双架构便携目录 + 安装说明 txt + 用户手册
    note_tpl = os.path.join(ROOT, "tools", "installer", "安装说明.txt")
    manual_src = os.path.join(DOCS, "JYYJ助手用户手册-%s.html" % ver)
    for arch in ("x64", "x86"):
        ref = old_ver or (auto_find_old_dir(ver, arch) or (None, None))[0]
        old_dir = os.path.join(DIST, "JYYJ助手-%s-wps-%s" % (ref, arch)) if ref else None
        new_dir = os.path.join(DIST, "JYYJ助手-%s-wps-%s" % (ver, arch))
        build_app_dir(arch, ver, ref, old_dir, new_dir)
        # 安装说明 txt：单一来源 = tools/installer/安装说明.txt（M17.7 起，杜绝随包版与
        # dist 版内容分叉）；模板缺失时回退旧版继承（ref 替换版本号），两者皆无则 fail-fast。
        txt_new = os.path.join(DIST, "JYYJ助手-%s-wps-安装说明.txt" % ver)
        if os.path.isfile(note_tpl):
            shutil.copyfile(note_tpl, txt_new)
        else:
            txt_old = os.path.join(DIST, "JYYJ助手-%s-wps-安装说明.txt" % ref) if ref else None
            if txt_old and os.path.isfile(txt_old):
                with open(txt_old, "r", encoding="utf-8-sig") as f:
                    content = f.read()
                with open(txt_new, "w", encoding="utf-8-sig") as f:
                    f.write(content.replace(ref, ver))
            elif not os.path.isfile(txt_new):
                # 审查者视角修复（M17.4）：ref=None 时静默跳过会让发版缺安装说明——fail-fast
                sys.exit("安装说明模板缺失（%s）且无旧版可继承：%s" % (note_tpl, txt_new))
        print("[pack] %s → %s" % (arch, new_dir))

    # 2b) 用户手册 HTML 集合进安装包与便携包（M17.7）：
    #     docs/ 版本化手册 → tools/installer/用户手册.html（ISCC 捆绑源，文件名不带版本防 bump 断链）
    #                     → dist 双架构便携目录根（与安装说明并列）
    if not os.path.isfile(manual_src):
        sys.exit("当前版本用户手册不存在：%s——请先更新手册再打包。" % manual_src)
    shutil.copyfile(manual_src, os.path.join(ROOT, "tools", "installer", "用户手册.html"))
    for arch in ("x64", "x86"):
        shutil.copyfile(manual_src,
                        os.path.join(DIST, "JYYJ助手-%s-wps-%s" % (ver, arch), "用户手册.html"))
    print("[pack] 用户手册 → 安装包 + 便携目录（%s）" % os.path.basename(manual_src))

    # 2c) MCP server 独立包（M27R3 审查：原先只有手工产物，pack 不重产出 → npx 分发会断）
    MCP_BUILD = os.path.join(ROOT, "tools", "McpServer", "bin", "Release", "net48")
    MCP_FILES = ("ExcelAiMcpServer.exe", "ExcelAiMcpServer.exe.config",
                 "ExcelAiAssistant.Core.dll", "ExcelAiAssistant.Excel.dll",
                 "Newtonsoft.Json.dll", "System.Data.SQLite.dll")
    run(["dotnet", "build", os.path.join("tools", "McpServer", "ExcelAiMcpServer.csproj"),
         "-c", "Release", "--no-incremental", "--nologo", "-v", "q"])
    mcp_dir = os.path.join(DIST, "JYYJ助手-%s-mcp-server-x64" % ver)
    if os.path.isdir(mcp_dir):
        shutil.rmtree(mcp_dir, ignore_errors=True)
    os.makedirs(mcp_dir)
    for name in MCP_FILES:
        src = os.path.join(MCP_BUILD, name)
        if not os.path.exists(src):
            sys.exit("[pack] mcp-server 依赖缺失：%s（先构建 tools/McpServer）" % src)
        shutil.copy2(src, os.path.join(mcp_dir, name))
    for arch in ("x64", "x86"):
        os.makedirs(os.path.join(mcp_dir, arch), exist_ok=True)
        shutil.copy2(os.path.join(MCP_BUILD, arch, "SQLite.Interop.dll"),
                     os.path.join(mcp_dir, arch, "SQLite.Interop.dll"))
    # 平铺打包：外壳 candidatePaths 含 cacheDir 根，解压后即可定位 exe
    mcp_zip = make_zip(mcp_dir, flat=True)
    print("[zip] %s %.2f MB（MCP server 独立包，供 npx jyyj-mcp 下载）" %
          (os.path.basename(mcp_zip), os.path.getsize(mcp_zip) / 1048576))
    # 版本戳断言：exe 内必须含新版本（UTF-16LE）
    with open(os.path.join(mcp_dir, "ExcelAiMcpServer.exe"), "rb") as fh:
        if ver.encode("utf-16-le") not in fh.read():
            sys.exit("[pack] mcp-server 包版本戳缺失：%s" % ver)

    # 便携单文件包（与 dist 其余形态同名规则对齐，供外壳第二档兜底）
    single_zip = None
    single_dir = os.path.join(DIST, "JYYJ助手-%s-wps-x64" % ver, "app")
    if os.path.isdir(single_dir):
        single_tmp = os.path.join(DIST, "JYYJ助手-%s-wps-单文件便携-x64" % ver)
        if os.path.isdir(single_tmp):
            shutil.rmtree(single_tmp, ignore_errors=True)
        shutil.copytree(single_dir, single_tmp)
        single_zip = make_zip(single_tmp, flat=True)
        print("[zip] %s %.2f MB（单文件便携）" %
              (os.path.basename(single_zip), os.path.getsize(single_zip) / 1048576))

    # 3) zip（进程内压缩，不用 PowerShell Compress-Archive——沙箱红线）
    zips = []
    for arch in ("x64", "x86"):
        zp = make_zip(os.path.join(DIST, "JYYJ助手-%s-wps-%s" % (ver, arch)))
        zips.append((arch, zp, os.path.getsize(zp)))
        print("[zip] %s %.2f MB" % (os.path.basename(zp), os.path.getsize(zp) / 1048576))

    # 4) 版本戳断言
    assert_version_stamps(ver, old_ver or (auto_find_old_dir(ver, "x64") or (None, None))[0])

    # 5) ISCC（/DAppVersion 注入模板 iss）；输出 exe 可能被杀软/索引暂态锁定（Error 32），自动退避重试
    iscc = find_iscc()
    for attempt, delay in enumerate((0, 5, 10, 20), 1):
        if delay:
            import time
            time.sleep(delay)
            print("[iscc] 重试第 %d 次（输出文件暂态锁定退避 %ds）" % (attempt, delay))
        env = dict(os.environ)
        for k, v in ENV_FIX.items():
            env.setdefault(k, v)
        r = subprocess.run([iscc, "/DAppVersion=" + ver, ISS], env=env, cwd=ROOT,
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        out = r.stdout.decode("utf-8", "replace")
        if r.returncode == 0:
            print(out.strip().splitlines()[-1])
            break
        if r.returncode == 2 and "Error 32" in out and attempt < 4:
            print(out.strip().splitlines()[-2] if "Error in" in out else out[-200:])
            continue
        sys.exit("ISCC 失败（exit %d）：\n%s" % (r.returncode, out[-1500:]))

    # 6) 下载链完整性清单（M28）：安装器 + 全部 zip 的 SHA256，随 Release 一同上传；
    #    npm 外壳 install 必须在解压前校验，未知来源包一律拒绝。
    installer_path = os.path.join(DIST, "installer", "JYYJ助手-%s-Setup-wps.exe" % ver)
    manifest_path, manifest_mapping = write_sha256_manifest(
        ver, [installer_path] + [zp for _arch, zp, _sz in zips] + [mcp_zip, single_zip])
    assert_manifest_matches_publish(ver, manifest_path, manifest_mapping)
    assert_npm_version(ver)

    print("== pack 完成 ==")
    for arch, zp, sz in zips:
        print("  %s %s（%.2f MB）" % (arch, zp, sz / 1048576))
    print("  安装器：%s" % installer_path)
    return 0


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p1 = sub.add_parser("bump", help="升版：props + README（唯一改版本入口）")
    p1.add_argument("--ver", required=True)
    p1.add_argument("--dry-run", action="store_true")
    p1.set_defaults(fn=cmd_bump)

    p2 = sub.add_parser("pack", help="打包：构建→组装→zip→戳断言→ISCC")
    p2.add_argument("--from", dest="from_version", help="旧版本号（安装说明来源），缺省自动探测 dist 最高旧版")
    p2.add_argument("--dry-run", action="store_true")
    p2.set_defaults(fn=cmd_pack)

    args = ap.parse_args()
    return args.fn(args)


if __name__ == "__main__":
    sys.exit(main())
