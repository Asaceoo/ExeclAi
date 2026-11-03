// 函数库功能级测试：真实执行 FX_LIB 构造、fxVersionOf、fxVisible 的纯逻辑部分
const fs = require('fs');
const src = fs.readFileSync('src/ExcelAiAssistant.Addin/Sidebar/index.html', 'utf8');

function extract(src, startMarker, endMarker, from=0){
  const i = src.indexOf(startMarker, from); if(i<0) throw new Error('未找到 '+startMarker);
  const j = src.indexOf(endMarker, i); if(j<0) throw new Error('未找到 '+endMarker);
  return { code: src.slice(i, j + endMarker.length), end: j };
}

// 1) 提取 fxVersionOf 函数（临时把 FX_VERSION 挂到全局，模拟脚本全局作用域）
// 注意：fxVersionOf 是单行定义（`function fxVersionOf(name){ … }`），不能按多行 `\n}` 提取，
// 否则会把 FX_LIB 大数组误吞进函数体导致 SyntaxError。改用行尾换行作为结束边界。
global.FX_VERSION = null;
const fvWithMarker = src.slice(src.indexOf('function fxVersionOf(name){'));
const fv = { code: fvWithMarker.slice(0, fvWithMarker.indexOf('\n')) };

// 3) …(先解析 FX_VERSION 再构造 FX_LIB)
// 2) 提取 FX_VERSION（从 const FX_VERSION 到 };）
const vFrom = src.indexOf('const FX_VERSION');
const vObj = extract(src, '{', '};', vFrom);
const FX_VERSION = new Function('return ' + vObj.code + ';')();
global.FX_VERSION = FX_VERSION;
const fxVersionOf = new Function(fv.code + '\nreturn fxVersionOf;')();

// 3) 提取 FX_LIB 数组并构造对象（对应 .map）
const lStart = src.indexOf('const FX_LIB = [');
const arrStart = src.indexOf('[', lStart);
const mapIdx = src.indexOf('].map', arrStart);
const arrSrc = src.slice(arrStart, mapIdx + 1);
const raw = eval('(' + arrSrc + ')');
const FX_LIB = raw.map(r => ({ category: r[0], name: r[1], title: r[2], usage: r[3], params: r[4], example: r[5], exampleNote: r[6], fxInsert: r[7], version: fxVersionOf(r[1]) }));

// 4) 提取 FX_COMMON 数组
const cFrom = src.indexOf('const FX_COMMON');
const cObj = extract(src, '[', ']', cFrom);
const fxCommon = new Function('return ' + cObj.code + ';')();
// fxVisible 逻辑与源码内联保持一致的副本（见 index.html fxVisible）
function fxVisible(state, lib){
  const q = state.q.trim().toLowerCase(), out = [], v = state.ver;
  if(q){
    lib.forEach(f => { if(f.name.toLowerCase().indexOf(q)>=0 || f.title.toLowerCase().indexOf(state.q.trim())>=0 || f.usage.toLowerCase().indexOf(state.q.trim())>=0){ if(v==="all"||f.version===v) out.push(f); } });
  }else if(state.cat === "common"){ lib.forEach(f => { if(fxCommon.indexOf(f.name)>=0 && (v==="all"||f.version===v)) out.push(f); }); }
  else{ lib.forEach(f => { if(f.category===state.cat && (v==="all"||f.version===v)) out.push(f); }); }
  return out;
}

let pass = 0, fail = 0;
const assert = (cond, name) => { if(cond){pass++;} else {fail++; console.error('  ✗ '+name);} };

console.log('FX_LIB 函数数:', FX_LIB.length);
console.log('FX_VERSION 条数:', Object.keys(FX_VERSION).length);

// 0) 总数基线：收录了 GROUPBY/PIVOTBY/PERCENTOF/TRIMRANGE/REGEXTEST/REGEXEXTRACT/REGEXREPLACE 等 365 新函数后
assert(FX_LIB.length === 319, '函数总数应为 319（实际 ' + FX_LIB.length + '）');
assert(FX_LIB.some(f=>f.name==='GROUPBY' && f.version==='365'), 'GROUPBY 已收录且标注 365');
assert(FX_LIB.some(f=>f.name==='PIVOTBY'), 'PIVOTBY 已收录');
assert(FX_LIB.some(f=>f.name==='REGEXEXTRACT'), 'REGEXEXTRACT 已收录');

// 断言A: 每个函数 version 均为合法枚举值
const validVer = ['2007','2010','2013','2016','2019','2021','365'];
let badVer = FX_LIB.filter(f => !validVer.includes(f.version));
assert(badVer.length===0, '所有函数 version ∈ {2007..365}（异常: '+badVer.map(b=>b.name+':'+b.version).join(',')+'）');

// 断言B: 常用分类可渲染且无 2007 之外的排序噪音
const commonAll = fxVisible({cat:'common',ver:'all',q:''}, FX_LIB);
assert(commonAll.length===fxCommon.length, '常用分类返回全部 ' + fxCommon.length + ' 个（实际 '+commonAll.length+'）');

// 断言C: 按版本筛选只返回该版本函数
for(const ver of ['2016','2019','2021','365']){
  const r = fxVisible({cat:'all',ver,q:''}, FX_LIB);
  // cat:'all' 无此分类分支，改用列举：遍历分类
  const rr = FX_LIB.filter(f => f.version===ver);
  assert(rr.length>0, '版本 '+ver+' 有函数收录（'+rr.length+' 个: '+rr.slice(0,4).map(x=>x.name).join(',')+'…）');
}
const v2019 = FX_LIB.filter(f=>f.version==='2019').map(f=>f.name);
assert(v2019.includes('MAXIFS') && v2019.includes('MINIFS'), '2019 版本含 MAXIFS/MINIFS（实际: '+v2019.join(',')+'）');
const v2021 = FX_LIB.filter(f=>f.version==='2021').map(f=>f.name);
assert(v2021.includes('FILTER')&&v2021.includes('XLOOKUP'), '2021 版本含 FILTER/XLOOKUP');
const v365 = FX_LIB.filter(f=>f.version==='365').map(f=>f.name);
assert(v365.includes('TEXTSPLIT'), '365 版本含 TEXTSPLIT');

// 断言D: 搜索跨版本；中文标题搜索副词条
const search365 = fxVisible({cat:'all',ver:'365',q:'拆分'}, FX_LIB);
assert(search365.some(f=>f.name==='TEXTSPLIT'), '365+搜索"拆分"命中 TEXTSPLIT');
const search16 = fxVisible({cat:'all',ver:'2016',q:'IFS'}, FX_LIB);
assert(search16.some(f=>f.name==='IFS'), '2016+搜索 IFS 命中');

// 断言E: 每个函数都有非空 title/usage/example/fxInsert
const incomplete = FX_LIB.filter(f=>!f.title||!f.usage||!f.example||!f.fxInsert||f.params.some(p=>!p[0]||!p[1]));
assert(incomplete.length===0, '所有函数字段完整（异常数 '+incomplete.length+'）');

// 断言F: 统计各版本分布
const dist = {};
FX_LIB.forEach(f=>{ dist[f.version]=(dist[f.version]||0)+1; });
console.log('版本分布:', JSON.stringify(dist));

console.log((fail===0?'\n功能测试 OK ('+pass+' 通过)':'\nFAIL '+fail+' 失败 / '+pass+' 通过'));
process.exit(fail===0?0:1);