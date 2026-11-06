# -*- coding: utf-8 -*-
"""gate.py —— 一键门禁（M17.1 grill 清单 ③，把"人肉按序跑"固化成一条命令）。

完整链：构建全解 → 全量单测 → Excel 真机 harness → WPS full 真机 harness →
        WPS full 基线校验（check_wps_full.py）→（可选 --pack 直通发版）。
任一步失败立即中止（fail-fast），退出码非 0。

用法：
    python tools/release/gate.py [--skip-build] [--skip-excel] [--skip-wps] [--pack]

产物/日志：_dev_tmp/gate_*.log；末尾打印门禁摘要（单测数/真机断言数/基线校验结论）。
注意：真机 harness 会启动独立 Excel/WPS 进程，运行前请关闭正在编辑的宿主文件。
"""
import argparse
import io
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
from release import ENV_FIX, ROOT as R_ROOT  # noqa: E402  复用环境修复与根目录

assert R_ROOT == ROOT
DEV_TMP = os.path.join(ROOT, "_dev_tmp")
SLN = os.path.join(ROOT, "ExcelAiAssistant.sln")
EXCEL_EXE = os.path.join(ROOT, "tools", "RealMachineVerify", "bin", "Release", "net48", "RealMachineVerify.exe")
WPS_EXE = os.path.join(ROOT, "tools", "WpsToolVerify", "bin", "Release", "net48", "WpsToolHarness.exe")
SUMMARY = []


def run(cmd, log_name=None, timeout=900, must=True):
    """运行命令；log_name 给定时把 stdout/stderr 重定向到 _dev_tmp/gate_<log_name>.log。"""
    env = dict(os.environ)
    for k, v in ENV_FIX.items():
        env.setdefault(k, v)
    printable = " ".join(cmd)
    print("[gate] " + printable, flush=True)
    stdout = stderr = None
    log_path = None
    if log_name:
        os.makedirs(DEV_TMP, exist_ok=True)
        log_path = os.path.join(DEV_TMP, "gate_%s.log" % log_name)
        stdout = open(log_path, "wb")
        stderr = subprocess.STDOUT
    r = subprocess.run(cmd, env=env, cwd=ROOT, stdout=stdout, stderr=stderr, timeout=timeout)
    if log_path:
        stdout.close()
    if must and r.returncode != 0:
        if log_path:
            tail = io.open(log_path, "rb").read().decode("utf-8", "replace").splitlines()[-15:]
            print("\n".join("    " + t for t in tail))
        sys.exit("[gate] 失败（exit %d）：%s" % (r.returncode, printable))
    return r, log_path


def parse_excel_summary(log_path):
    """解析 Excel 真机汇总行：== 结果：通过 119，失败 0，无进程泄漏…=="""
    text = io.open(log_path, "rb").read().decode("utf-8", "replace")
    m = re.search(r"通过\s*(\d+)，失败\s*(\d+)", text)
    if not m:
        sys.exit("[gate] Excel 真机日志缺少汇总行（通过 N，失败 M）：%s" % log_path)
    passed, failed = int(m.group(1)), int(m.group(2))
    if failed != 0:
        sys.exit("[gate] Excel 真机存在 %d 个失败断言，详见 %s" % (failed, log_path))
    SUMMARY.append(("Excel 真机", "%d 断言全过" % passed))


HOST_NAMES = ("excel.exe", "et.exe", "wps.exe", "wpp.exe", "wpsoffice.exe")


def snap_hosts():
    """快照当前宿主进程 (name_lower, pid) 集合（tasklist CSV 解析，不依赖 /FI）。"""
    env = dict(os.environ)
    for k, v in ENV_FIX.items():
        env.setdefault(k, v)
    r = subprocess.run(["tasklist", "/NH", "/FO", "CSV"], env=env, cwd=ROOT,
                       capture_output=True, timeout=60)
    hosts = set()
    for line in r.stdout.decode("utf-8", "replace").splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) >= 2 and parts[0].lower() in HOST_NAMES:
            hosts.add((parts[0].lower(), parts[1]))
    return hosts


def assert_no_host_leak(before, step):
    """真机步后断言：不新增宿主进程泄漏。M17.5 事故实证——harness 被中断时
    KillSpawnedExcel 兜底会静默失效，遗留的 DCOM 拉起 EXCEL 跨 silo 不可杀，
    会锁死用户后续安装（"管理员也装不上"）。泄漏必须当场亮红，不许静默带过。"""
    leaked = sorted(snap_hosts() - before)
    if leaked:
        detail = ", ".join("%s(pid=%s)" % (n, p) for n, p in leaked)
        sys.exit("[gate] %s 真机步泄漏宿主进程：%s——先手动清理（必要时重启电脑）"
                 "再重跑门禁；遗留进程会锁死用户安装包的宿主检测。" % (step, detail))
    SUMMARY.append((step + " 泄漏检查", "无新增宿主进程"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--skip-build", action="store_true", help="跳过 sln 构建（刚构建过）")
    ap.add_argument("--skip-excel", action="store_true", help="跳过 Excel 真机")
    ap.add_argument("--skip-wps", action="store_true", help="跳过 WPS full 真机")
    ap.add_argument("--pack", action="store_true", help="门禁全绿后直通 release.py pack")
    args = ap.parse_args()
    py = sys.executable

    # 1) 构建 + 单测
    if not args.skip_build:
        run(["dotnet", "build", SLN, "-c", "Release", "--nologo", "-v", "q"], log_name="build")
    r, log = run(["dotnet", "test", SLN, "-c", "Release", "--no-build", "--nologo"], log_name="test")
    text = io.open(log, "rb").read().decode("utf-8", "replace")
    m = re.search(r"已通过!?\s*-\s*失败:\s*(\d+)，通过:\s*(\d+)", text)
    if not m:
        sys.exit("[gate] 单测日志缺少汇总行，详见 %s" % log)
    if int(m.group(1)) != 0:
        sys.exit("[gate] 单测失败 %d 个，详见 %s" % (int(m.group(1)), log))
    SUMMARY.append(("单测", "%s 全过" % m.group(2)))

    # 2) Excel 真机（步前快照宿主、步后断言无泄漏——M17.5 事故兜底）
    if not args.skip_excel:
        run(["dotnet", "build", os.path.join(ROOT, "tools", "RealMachineVerify", "RealMachineVerify.csproj"),
             "-c", "Release", "--nologo", "-v", "q"], log_name="build_rmv")
        before_excel = snap_hosts()
        run([EXCEL_EXE], log_name="excel", timeout=1200)
        parse_excel_summary(os.path.join(DEV_TMP, "gate_excel.log"))
        assert_no_host_leak(before_excel, "Excel")

    # 3) WPS full 真机 + 基线校验（同泄漏断言）
    if not args.skip_wps:
        run(["dotnet", "build", os.path.join(ROOT, "tools", "WpsToolVerify", "WpsToolHarness.csproj"),
             "-c", "Release", "--nologo", "-v", "q"], log_name="build_wps")
        before_wps = snap_hosts()
        run([WPS_EXE, "full"], log_name="wps_full", timeout=1200)
        assert_no_host_leak(before_wps, "WPS")
        run([py, os.path.join(HERE, "check_wps_full.py"), "--log", os.path.join(DEV_TMP, "gate_wps_full.log")])
        SUMMARY.append(("WPS full 基线", "校验通过"))

    # 4) 摘要
    print("== 门禁摘要 ==")
    for name, verdict in SUMMARY:
        print("  [PASS] %s：%s" % (name, verdict))

    # 5) 可选直通发版
    if args.pack:
        # 门禁直通发版：默认 --no-bump（版本应由 release.py bump 显式决定），
        # 避免"门禁代跑一次 pack"把版本再抬一级、与已发文档版本错位。
        run([py, os.path.join(HERE, "release.py"), "pack", "--no-bump"])
        SUMMARY.append(("pack", "完成"))

    print("== GATE PASS ==")
    return 0


if __name__ == "__main__":
    sys.exit(main())
