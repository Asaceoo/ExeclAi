/* 运行时冒烟：node vm 中 mock 最小 DOM，从头部执行整个 <script>，定位首个运行时异常。
   仅用于定位"函数库加载不出"类运行时崩溃，不替代真机验证。
   用法：node tools/runtime-smoke.js */
const fs = require("fs");
const vm = require("vm");
const src = fs.readFileSync("src/ExcelAiAssistant.Addin/Sidebar/index.html", "utf8");
const m = src.match(/<script>([\s\S]*?)<\/script>/);
if (!m) { console.error("未找到 <script>"); process.exit(1); }
const code = m[1];

function makeEl(id) {
  const listeners = {};
  return {
    id, innerHTML: "", textContent: "", value: "", style: { display: "block" }, dataset: {},
    classList: { add(){}, remove(){}, toggle(){} }, disabled: false,
    querySelector(){ return makeEl(id + ":q"); },
    querySelectorAll(){ return []; },
    addEventListener(ev, fn){ (listeners[ev] = listeners[ev] || []).push(fn); },
    dispatchEvent(){}, focus(){}, setSelectionRange(){}, click(){}, remove(){}, appendChild(){},
    setAttribute(){}, getAttribute(){ return null; }, contains(){ return false; },
  };
}
const store = {};
for (const mm of src.matchAll(/\bid="([^"]+)"/g)) store[mm[1]] = makeEl(mm[1]);
const doc = {
  getElementById(id){ if(!store[id]) store[id] = makeEl(id); return store[id]; },
  querySelector(sel){ return makeEl(String(sel)); },
  querySelectorAll(){ return []; },
  addEventListener(){}, createElement(){ return makeEl("c"); },
  body: makeEl("body"), documentElement: makeEl("html"),
  execCommand(){}, hasFocus(){ return true; },
};
const bridge = new Proxy({}, { get: (t, name) => (...args) => Promise.resolve(
  name === "GetContext" ? '{"ok":true,"context":{"Selection":{"Address":"$C$13:$L$13"}}}' :
  name === "PreviewFormula" ? '{"ok":true,"value":"3.14","cell":"$C$13"}' :
  name === "GetVersion" ? '{"ok":true,"version":"smoke"}' : '{"ok":true}'
) });
const win = { chrome: { webview: { addEventListener(){}, postMessage(){}, hostObjects: { sync: { excelAiBridge: bridge }, excelAiBridge: bridge } } },
  addEventListener(){}, localStorage: { getItem(){ return null; }, setItem(){} },
  location: { href: "", reload(){} }, navigator: { clipboard: { writeText(){ return Promise.resolve(); } } },
  matchMedia(){ return { matches: false, addEventListener(){}, addListener(){} }; },
  setTimeout, clearTimeout, setInterval, clearInterval };
const ctx = { window: win, document: doc, console, setTimeout, clearTimeout, setInterval, clearInterval,
  navigator: win.navigator, location: win.location, localStorage: win.localStorage, screen: {},
  Promise, JSON, Math, Date, Number, String, Boolean, RegExp, Error, Array, Object,
  parseInt, parseFloat, isNaN, isFinite, decodeURIComponent, encodeURIComponent };
ctx.globalThis = ctx;
vm.createContext(ctx);
try {
  vm.runInContext(code, ctx, { timeout: 10000 });
  console.log("=== 冒烟通过：整个脚本从头执行到末尾，同步+已触发异步无异常 ===");
} catch (e) {
  console.error("=== 运行时异常 ===");
  console.error("Message:", e.message);
  console.error("Stack:", String(e.stack).split("\n").slice(0, 6).join("\n"));
  process.exit(1);
}