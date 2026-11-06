/* 函数库详情三功能定向冒烟：参数逐步讲解 / 示例试运行 / 相似函数对比
   断言代码注入脚本作用域内执行（访问顶层 const = FX_LIB/SIMILAR/fxFillDet）。
   用法：node tools/fxdet-smoke.js */
const fs = require("fs");
const vm = require("vm");

const src = fs.readFileSync("src/ExcelAiAssistant.Addin/Sidebar/index.html", "utf8");
const m = src.match(/<script>([\s\S]*?)<\/script>/);
if (!m) { console.error("未找到 <script>"); process.exit(1); }
const code = m[1];

const idCache = {};
function makeEl(id) {
  const listeners = {};
  return {
    id, innerHTML: "", textContent: "", value: "", style: {}, dataset: {},
    className: "", classList: { add(){}, remove(){}, toggle(){} }, disabled: false,
    querySelector(){ return makeEl(id + ":q"); },
    querySelectorAll(){ return []; },
    addEventListener(ev, fn){ (listeners[ev] = listeners[ev] || []).push(fn); },
    dispatchEvent(){}, focus(){}, setSelectionRange(){}, click(){}, remove(){},
    setAttribute(){}, getAttribute(){ return null; }, contains(){ return false; },
    appendChild(){},
  };
}
const doc = {
  getElementById(id){ if(!idCache[id]) idCache[id] = makeEl(id); return idCache[id]; },
  querySelector(s){ return makeEl(String(s)); },
  querySelectorAll(){ return []; },
  addEventListener(){}, createElement(){ return makeEl("c"); },
  body: makeEl("body"), documentElement: makeEl("html"),
  execCommand(){}, hasFocus(){ return true; },
};
const bridgeCtl = { fail: false };
const bridge = new Proxy({}, { get: (t, name) => (name === "PreviewFormula" && bridgeCtl.fail)
  ? () => Promise.reject(new Error("host down"))
  : (...args) => Promise.resolve(
  name === "GetContext" ? '{"ok":true,"context":{"Selection":{"Address":"$C$13:$L$13","SampleValues":[[""]]}}}' :
  name === "PreviewFormula" ? '{"ok":true,"value":"3.14","cell":"$C$13"}' :
  name === "GetVersion" ? '{"ok":true,"version":"fxdet"}' :
  name === "InsertFormula" ? JSON.stringify({ ok: true }) : '{"ok":true}'
) });
const win = {
  chrome: { webview: { addEventListener(){}, postMessage(){}, hostObjects: { sync: { excelAiBridge: bridge }, excelAiBridge: bridge } } },
  addEventListener(){}, localStorage: { getItem(){ return null; }, setItem(){} },
  location: { href: "", reload(){} }, navigator: { clipboard: { writeText(){ return Promise.resolve(); } } },
  matchMedia(){ return { matches: false, addEventListener(){}, addListener(){} }; },
  alert(){}, setTimeout, clearTimeout, setInterval, clearInterval,
};
const ctx = { window: win, document: doc, console, setTimeout, clearTimeout, setInterval, clearInterval,
  navigator: win.navigator, location: win.location, localStorage: win.localStorage, screen: {},
  Promise, JSON, Math, Date, Number, String, Boolean, RegExp, Error, Array, Object,
  parseInt, parseFloat, isNaN, isFinite, decodeURIComponent, encodeURIComponent };
ctx.globalThis = ctx;
ctx.__bridgeCtl = bridgeCtl;
vm.createContext(ctx);

// 断言代码在脚本作用域内执行，可访问顶层 const
const testCode = `
;(async () => {
  const log = [];
  let pass = 0, fail = 0;
  function check(name, cond, detail) {
    if (cond) { pass++; log.push("[PASS] " + name); }
    else { fail++; log.push("[FAIL] " + name + (detail ? " -- " + detail : "")); }
  }
  if (!FX_LIB || !FX_LIB.length) { log.push("[FAIL] FX_LIB 为 null/空"); fail++; return { pass, fail, log }; }
  log.push("FX_LIB 函数数 = " + FX_LIB.length);

  log.push("=== 1) 参数逐步讲解 ===");
  const sumif = FX_LIB.find(x => x.name === "SUMIF");
  if (!sumif) { check("SUMIF 在库", false); }
  else {
    fxFillDet(sumif);
    check("fxLibDet 显示", document.getElementById("fxLibDet").style.display === "block");
    const pms = document.getElementById("fxLibPms").innerHTML;
    check("参数讲解编号数=参数数", (pms.match(/fx-step-n/g) || []).length === sumif.params.length,
          "steps=" + (pms.match(/fx-step-n/g)||[]).length + " params=" + sumif.params.length);
    check("必填标记渲染", pms.indexOf("必填") !== -1);
    check("参数名+中文渲染", pms.indexOf("<b>") !== -1 && pms.indexOf("—") !== -1);
    check("示例文案写入", document.getElementById("fxLibExp").textContent === sumif.example);
    check("试运行按钮挂载公式", document.getElementById("fxRunDemo").dataset.formula === sumif.example);
  }

  log.push("=== 2) 相似函数对比 ===");
  const cases = ["SUMIF","ROUND","VLOOKUP","SUM"];
  for (const nm of cases) {
    const f = FX_LIB.find(x => x.name === nm);
    if (!f) continue;
    fxFillDet(f);
    const wrap = document.getElementById("fxCmpWrap");
    const ok = wrap.style.display === "block" && wrap.innerHTML.indexOf("current") !== -1 && wrap.innerHTML.indexOf("选型口诀") !== -1;
    check(nm + " 对比卡(共" + (SIMILAR[nm] ? SIMILAR[nm].length : 0) + "项)", ok, "display=" + wrap.style.display);
  }
  const plain = FX_LIB.find(x => !(SIMILAR && SIMILAR[x.name]));
  if (plain) {
    fxFillDet(plain);
    check(plain.name + " 无相似→隐藏对比块", document.getElementById("fxCmpWrap").style.display !== "block");
  }

  log.push("=== 3) 示例试运行 ===");
  const g = FX_LIB.find(x => x.name === "SUMIF") || FX_LIB[0];
  fxFillDet(g);
  await fxRunDemoFn();
  const res = document.getElementById("fxDemoResult");
  check("结果区显示", res.style.display === "inline-block");
  check("值=3.14", res.textContent.indexOf("3.14") !== -1, res.textContent);
  check("非 err 样式", res.className.indexOf("err") === -1, res.className);
  check("按钮恢复", document.getElementById("fxRunDemo").disabled === false);

  globalThis.__bridgeCtl.fail = true;
  await fxRunDemoFn();
  check("桥异常走 err 分支", document.getElementById("fxDemoResult").className.indexOf("err") !== -1);
  globalThis.__bridgeCtl.fail = false;

  log.push("=== 4) 插入占位公式（空选区） ===");
  fxState.sel = g;
  const prevConfirmed = document.getElementById("fxLibConfirm");
  await fxInsertFn();
  check("空选区直接插入无确认", prevConfirmed.style.display !== "block");

  globalThis.__fxdetPromise_resolve({ pass, fail, log });
})();
`;

// 追加断言，并把结果回传到宿主
let resolveOut;
const donePromise = new Promise(r => { resolveOut = r; });
ctx.__fxdetPromise_resolve = resolveOut;
vm.runInContext(code + testCode + ";\nsetTimeout(()=>globalThis.__fxdetPromise_resolve({pass:0,fail:1,log:['断言脚本未完成']}),30000);", ctx, { timeout: 8000 });

donePromise.then(r => {
  console.log(r.log.join("\n"));
  console.log("\n结果：" + r.pass + " 通过, " + r.fail + " 失败");
  process.exit(r.fail ? 1 : 0);
}).catch(e => { console.error("异常:", e); process.exit(1); });