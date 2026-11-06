#!/usr/bin/env node
/*
 * jyyj-mcp：JYYJ助手 MCP server 分发外壳（P1-2，对齐 xlsx-for-ai 的 npx 一装模式）。
 * 运行时是 net48 原生 ExcelAiMcpServer.exe（需要本机装有 Excel/WPS），本壳只做三件事：
 *   1) locate  定位本机 exe（安装目录 / dist / --path 指定）
 *   2) fetch   找不到时从 GitHub Release 拉单文件便携 zip 并解压到缓存
 *   3) config  打印 Claude Desktop / DSH 等客户端的 MCP 配置片段
 * 子命令：jyyj-mcp [locate|config|run|install] [--path <exe>] [--version <x.y.z.w>] [--ro]
 */
const fs = require("fs");
const path = require("path");
const os = require("os");
const https = require("https");
const http = require("http");
const crypto = require("crypto");
const { execFileSync, spawnSync } = require("child_process");

const NET_TIMEOUT_MS = 60000; // M28：单请求超时（连接/空闲都算）
const PROXY = process.env.HTTPS_PROXY || process.env.https_proxy || process.env.HTTP_PROXY || process.env.http_proxy;

const REPO = "Asaceoo/ExeclAi";
const args = process.argv.slice(2);
const cmd = args[0] || "config";
const flag = (name) => { const i = args.indexOf("--" + name); return i >= 0 ? args[i + 1] : null; };
const has = (name) => args.includes("--" + name);
// M28 审查（M17.1 单一来源）：默认版本从本包 package.json 读取，bump 由 release.py 同步；
// 这里的 fallback 仅在 package.json 缺失/损坏时兜底（release.py pack 会断言两者一致，防 npx 拉旧 tag）。
let defaultVersion = null;
try { defaultVersion = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "package.json"), "utf8")).version; } catch (e) { }
if (!defaultVersion) defaultVersion = "3.9.1.2";
const version = flag("version") || defaultVersion;
// FIX（评审 A7）：版本号必须匹配 x.y.z.w——否则会拼进缓存目录与 API tag URL（目录穿越）
if (!/^\d+\.\d+\.\d+\.\d+$/.test(version)) {
  console.error("非法版本号（需 x.y.z.w）：" + version);
  process.exit(2);
}

const cacheDir = path.join(os.homedir(), ".jyyj-mcp", version);
const exeName = "ExcelAiMcpServer.exe";

function candidatePaths() {
  const local = process.env.LOCALAPPDATA || path.join(os.homedir(), "AppData", "Local");
  return [
    flag("path"),
    path.join(cacheDir, exeName),           // 独立 mcp-server 包：zip 根平铺
    path.join(cacheDir, "app", exeName),     // 兼容老结构
    path.join(local, "XaiAssistant", "app", "x64", exeName),
    path.join(process.cwd(), "tools", "McpServer", "bin", "Release", "net48", exeName)
  ].filter(Boolean);
}

function locate() {
  for (const p of candidatePaths()) { try { if (fs.existsSync(p)) return p; } catch (e) {} }
  return null;
}

function parseProxy(raw) {
  // 支持 https://user:pass@host:port 与 http://host:port
  const u = new URL(raw);
  return { host: u.hostname, port: parseInt(u.port || (u.protocol === "https:" ? "443" : "80"), 10), auth: u.username ? decodeURIComponent(u.username) + ":" + decodeURIComponent(u.password || "") : null };
}

function openUrl(url, headers, depth) {
  // M28：代理（HTTPS_PROXY/HTTP_PROXY 的 CONNECT 隧道）+ 超时 + 重定向；仅 https
  return new Promise((resolve, reject) => {
    if (depth > 5) return reject(new Error("重定向过多: " + url));
    let target;
    try { target = new URL(url); } catch (e) { return reject(new Error("非法 URL: " + url)); }
    if (target.protocol !== "https:") return reject(new Error("仅支持 https 下载: " + url));
    const onRes = (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        res.resume();
        return openUrl(new URL(res.headers.location, url).toString(), headers, depth + 1).then(resolve, reject);
      }
      if (res.statusCode !== 200) { res.resume(); return reject(new Error("HTTP " + res.statusCode + " " + url)); }
      resolve(res);
    };
    const options = { hostname: target.hostname, port: target.port ? parseInt(target.port, 10) : 443, path: target.pathname + target.search, headers: headers || {}, timeout: NET_TIMEOUT_MS };
    if (PROXY) {
      const px = parseProxy(PROXY);
      const connectReq = http.request({ host: px.host, port: px.port, method: "CONNECT", path: target.hostname + ":443", headers: px.auth ? { "Proxy-Authorization": "Basic " + Buffer.from(px.auth).toString("base64") } : {} });
      connectReq.setTimeout(NET_TIMEOUT_MS, () => connectReq.destroy(new Error("代理连接超时")));
      connectReq.on("connect", (cres, socket) => {
        if (cres.statusCode !== 200) { socket.destroy(); return reject(new Error("代理 CONNECT 失败 HTTP " + cres.statusCode)); }
        const tlsReq = https.request(Object.assign({}, options, { socket, servername: target.hostname }), onRes);
        tlsReq.setTimeout(NET_TIMEOUT_MS, () => tlsReq.destroy(new Error("下载超时")));
        tlsReq.on("error", reject).end();
      });
      connectReq.on("error", reject).end();
      return;
    }
    const req = https.get(options, onRes);
    req.setTimeout(NET_TIMEOUT_MS, () => req.destroy(new Error("下载超时")));
    req.on("error", reject);
  });
}

function download(url, dest) {
  // M28：先写 .part，完成才改名；失败/超时/中断删除半途文件（不留下“看似完整”的坏包）
  const part = dest + ".part";
  const cleanup = () => { try { fs.unlinkSync(part); } catch (e) {} };
  return openUrl(url, { "User-Agent": "jyyj-mcp" }, 0).then((res) => new Promise((resolve, reject) => {
    const out = fs.createWriteStream(part);
    let settled = false;
    const fail = (err) => { if (settled) return; settled = true; cleanup(); try { out.destroy(); } catch (e) {} reject(err); };
    res.on("error", fail);
    out.on("error", fail);
    out.on("finish", () => {
      out.close(() => {
        if (settled) return;
        try { fs.renameSync(part, dest); settled = true; resolve(dest); } catch (e) { fail(e); }
      });
    });
    res.pipe(out);
  }), (err) => { cleanup(); throw err; });
}

function sha256File(file) {
  return new Promise((resolve, reject) => {
    const h = crypto.createHash("sha256");
    const rs = fs.createReadStream(file);
    rs.on("error", reject);
    rs.on("data", (c) => h.update(c));
    rs.on("end", () => resolve(h.digest("hex")));
  });
}

function parseManifest(text) {
  const map = new Map(); // filename -> sha256
  for (const line of String(text).split(/\r?\n/)) {
    const m = line.trim().match(/^([0-9a-fA-F]{64})\s+[*]?(.+)$/);
    if (m) map.set(m[2].trim(), m[1].toLowerCase());
  }
  return map;
}

function findChecksumsAsset(assets, ver) {
  return assets.find((a) => /^SHA256SUMS-.+\.txt$/i.test(a.name) && a.name.includes(ver)) ||
         assets.find((a) => /^SHA256SUMS-.+\.txt$/i.test(a.name)) || null;
}

function verifyFile(file, expectedSha, expectedSize) {
  return new Promise((resolve, reject) => {
    const size = fs.statSync(file).size;
    if (expectedSize && size !== expectedSize) return reject(new Error("大小不符：期望 " + expectedSize + "，实际 " + size));
    sha256File(file).then((actual) => {
      if (!expectedSha) return reject(new Error("校验清单缺少该文件摘要：" + path.basename(file)));
      if (actual !== expectedSha.toLowerCase()) return reject(new Error("SHA256 不符：期望 " + expectedSha + "，实际 " + actual + "（下载被篡改或包已损坏）"));
      resolve(actual);
    }, reject);
  });
}

function findAsset(assets) {
  // 优先级：独立 mcp-server 包（含 ExcelAiMcpServer.exe）> 单文件便携 > 加载项 x64 包。
  // 注意：加载项包（wps-x64 / portable-single）里没有 MCP server exe，只是兜底提示用。
  const patterns = [/mcp-server-x64\.zip$/i, /portable-single-x64\.zip$/i]; // M27R3：删 wps-x64 兜底——该包不含 exe，命中只会浪费时间
  for (const pat of patterns) {
    const hit = assets.find(a => pat.test(a.name));
    if (hit) return hit;
  }
  return null;
}

function listReleaseAssets(tag) {
  // M28：走统一 openUrl（代理/超时/重定向），非 200 明确报错
  return openUrl("https://api.github.com/repos/" + REPO + "/releases/tags/" + tag,
    { "User-Agent": "jyyj-mcp", Accept: "application/vnd.github+json" }, 0).then((res) => new Promise((resolve, reject) => {
      let d = "";
      res.on("data", c => d += c);
      res.on("end", () => { try { resolve(JSON.parse(d).assets || []); } catch (e) { reject(new Error("Release API 返回非 JSON")); } });
      res.on("error", reject);
    }));
}

async function install() {
  fs.mkdirSync(cacheDir, { recursive: true });
  const tag = "v" + version;
  process.stderr.write("查询 Release " + tag + " …\n");
  const assets = await listReleaseAssets(tag);
  const zipAsset = findAsset(assets);
  if (!zipAsset) throw new Error("Release " + tag + " 未找到可用资产（期望 mcp-server-x64.zip）");
  // FIX（评审 A6）：资产名来自远端 API，必须白名单化后才可拼进命令行/文件名
  if (!/^[A-Za-z0-9._-]+$/.test(zipAsset.name) || zipAsset.name === "." || zipAsset.name === "..") throw new Error("资产名含非法字符，拒绝处理：" + zipAsset.name);
  // M28：下载链完整性——必须拿到同一 Release 的 SHA256SUMS 清单并校验大小+SHA256，
  // 清单缺失即拒绝自动安装（fail-closed，避免把未知来源的 exe 解压进缓存）。
  const sumsAsset = findChecksumsAsset(assets, version);
  if (!sumsAsset) throw new Error("Release " + tag + " 缺少 SHA256SUMS 校验清单，拒绝自动安装。请从 GitHub Releases 手动下载并核对，或用 --path 指定已有 exe。");
  if (!/^[A-Za-z0-9._-]+$/.test(sumsAsset.name)) throw new Error("校验清单资产名含非法字符：" + sumsAsset.name);
  const sumsPath = path.join(cacheDir, "SHA256SUMS-" + version + ".txt");
  process.stderr.write("下载校验清单 " + sumsAsset.name + " …\n");
  await download(sumsAsset.browser_download_url, sumsPath);
  const manifest = parseManifest(fs.readFileSync(sumsPath, "utf8"));
  if (!manifest.has(zipAsset.name)) throw new Error("校验清单缺少 " + zipAsset.name + " 的 SHA256，拒绝安装");
  const zipPath = path.join(cacheDir, zipAsset.name);
  process.stderr.write("下载 " + zipAsset.name + " (" + (zipAsset.size / 1048576).toFixed(1) + " MB) …\n");
  await download(zipAsset.browser_download_url, zipPath);
  await verifyFile(zipPath, manifest.get(zipAsset.name), zipAsset.size);
  process.stderr.write("校验通过：大小 " + fs.statSync(zipPath).size + " B，SHA256 " + manifest.get(zipAsset.name).slice(0, 16) + "…\n");
  // 解压：PowerShell Expand-Archive（Win10+ 自带，零依赖）
  // M28（第四轮续）：zip-slip 防线——解压后扫描缓存目录，任何条目解析后逃出 cacheDir（../、绝对路径、
  // 符号链接）即删缓存整体拒绝安装（fail-closed）；SHA256 只保证包完整性，不保证包内路径安全。
  process.stderr.write("解压到 " + cacheDir + " …\n");
  const r = spawnSync("powershell", ["-NoProfile", "-Command",
    "Expand-Archive -LiteralPath $env:JYYJ_ZIP -DestinationPath $env:JYYJ_DIR -Force"],
    { stdio: "inherit", env: (() => { const w = {}; for (const k of ["PATH","SystemRoot","COMSPEC","TEMP","TMP","USERPROFILE"]) { if (process.env[k]) w[k] = process.env[k]; } w.JYYJ_ZIP = zipPath; w.JYYJ_DIR = cacheDir; return w; })() });
  if (r.status !== 0) throw new Error("解压失败（exit " + r.status + "）");
  // zip-slip 校验：所有解压产物必须仍在 cacheDir 内（ realpath 防符号链接逃逸）
  const cacheReal = fs.realpathSync(cacheDir);
  for (const entry of walkDir(cacheReal)) {
    const real = fs.realpathSync(entry);
    if (real !== cacheReal && !real.startsWith(cacheReal + path.sep)) {
      throw new Error("检测到解压条目逃出缓存目录（疑似 zip-slip）： " + entry + " -> " + real + "。已拒绝安装，请手动清理 " + cacheDir);
    }
  }
  const exe = locate();
  if (!exe) throw new Error("解压后未找到 ExcelAiMcpServer.exe。可能原因：①该 Release 没有 mcp-server 独立包；②解压目录结构变化。请手动下载解压后加 --path 指定。");
  process.stderr.write("完成：" + exe + "\n");
  return exe;
}

// M28（第四轮续）：递归收集目录下全部文件/目录（zip-slip 校验用）
function walkDir(root) {
  const out = [];
  const stack = [root];
  while (stack.length) {
    const cur = stack.pop();
    for (const entry of fs.readdirSync(cur, { withFileTypes: true })) {
      const full = path.join(cur, entry.name);
      out.push(full);
      if (entry.isDirectory()) stack.push(full);
    }
  }
  return out;
}

function ensureExe() {
  let exe = locate();
  if (exe) return exe;
  if (has("no-download")) return null;
  // FIX（评审 A4）：execFileSync 在 stdio:inherit 下返回 null（不是 Buffer）——
  // 原来的 `execFileSync(...) && locate()` 恒为 null，导致下载成功也报「自动下载失败」。
  process.stderr.write("本机未找到 ExcelAiMcpServer.exe，尝试从 GitHub Release 下载便携版（--no-download 可关闭）…\n");
  try {
    execFileSync(process.execPath, [__filename, "install", "--version", version], { stdio: ["ignore", "ignore", "inherit"] });
  } catch (e) {
    return null;
  }
  return locate();
}

function configSnippet(exe, readOnly, confirmMode) {
  const configArgs = readOnly ? [] : ["--allow-write"]; // 默认只读；写模式需显式 --allow-write（评审 E10）
  if (confirmMode) { configArgs.push("--confirm-mode", confirmMode); } // M28：确认门模式透传（stderr=带外硬门）
  return JSON.stringify({
    mcpServers: {
      "jyyj-excel": {
        command: exe,
        args: configArgs,
        description: "JYYJ助手 Excel 工具链（MCP stdio）；需本机装有 Excel 或 WPS ET"
      }
    }
  }, null, 2);
}

(async () => {
  try {
    if (cmd === "locate") {
      const exe = locate();
      console.log(exe || "NOT-FOUND（可运行 jyyj-mcp install 下载便携版）");
    } else if (cmd === "install") {
      await install();
    } else if (cmd === "config") {
      const exe = ensureExe();
      if (!exe) { console.error("未找到 ExcelAiMcpServer.exe，且自动下载失败。装 JYYJ助手 或加 --path。"); process.exit(2); }
      console.log("# Claude Desktop: claude_desktop_config.json 追加以下片段即可（DSH 用户走 MCP 连接器→粘贴 JSON）：\n");
      // FIX（评审 E10）：默认输出只读配置（server 缺省只读更安全）；需要写能力显式加 --allow-write
      console.log(configSnippet(exe, !has("allow-write"), flag("confirm-mode")));
    } else if (cmd === "verify") {
      const manifestFile = flag("manifest");
      const file = flag("file");
      if (!manifestFile || !file) { console.error("用法：jyyj-mcp verify --manifest <SHA256SUMS-x.y.z.w.txt> --file <待校验文件>"); process.exit(2); }
      if (!fs.existsSync(manifestFile)) { console.error("清单不存在：" + manifestFile); process.exit(2); }
      if (!fs.existsSync(file)) { console.error("待校验文件不存在：" + file); process.exit(2); }
      const manifest = parseManifest(fs.readFileSync(manifestFile, "utf8"));
      const expected = manifest.get(path.basename(file));
      if (!expected) { console.error("清单中没有 " + path.basename(file) + " 的摘要"); process.exit(3); }
      const actual = await verifyFile(file, expected, 0);
      console.log("OK " + path.basename(file) + " SHA256=" + actual);
    } else if (cmd === "run") {
      const exe = ensureExe();
      if (!exe) { console.error("未找到 ExcelAiMcpServer.exe"); process.exit(2); }
      // M28 审查：run 默认只读（与 config/MCP 默认一致）；写模式需显式 --allow-write
      const runArgs = has("allow-write") ? ["--allow-write"] : [];
      const confirmMode = flag("confirm-mode");
      if (confirmMode) { runArgs.push("--confirm-mode", confirmMode); }
      const child = spawnSync(exe, runArgs, { stdio: "inherit" });
      process.exit(child.status || 0);
    } else {
      console.log("用法：jyyj-mcp [locate|install|config|run|verify] [--version x.y.z.w] [--path <exe>] [--allow-write] [--confirm-mode file|stderr] [--no-download]\n" +
        "  install 从 Release 下载 mcp-server 包（必须通过 SHA256SUMS 大小+SHA256 校验；\n" +
        "          代理经 HTTPS_PROXY/HTTP_PROXY 的 CONNECT 隧道，单请求 60s 超时，失败自动清理半途文件）\n" +
        "  verify  手动校验下载产物：--manifest <SHA256SUMS-x.txt> --file <待校验文件>\n" +
        "  config  打印 MCP 客户端配置片段（默认只读；加 --allow-write 才输出可写配置）\n" +
        "  run     直接以 stdio 启动 server（客户端 command 也可指本 npm 包）");
    }
  } catch (e) {
    console.error("错误：" + e.message);
    process.exit(1);
  }
})();
