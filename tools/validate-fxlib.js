// 函数库结构校验：真实解析 FX_LIB 数组文字量，检测重复/字段缺失/版本映射，剔除嵌套参数数组干扰。
// 用法: node tools/validate-fxlib.js [--fix-syntax]
const fs = require('fs');
const path = 'src/ExcelAiAssistant.Addin/Sidebar/index.html';
const src = fs.readFileSync(path, 'utf8');

const libStart = src.indexOf('const FX_LIB = [');
if (libStart < 0) { console.error('FAIL: const FX_LIB = [ not found'); process.exit(1); }
const arrStart = src.indexOf('[', libStart);
const mapIdx = src.indexOf('].map', arrStart);
if (mapIdx < 0) { console.error('FAIL: "].map" not found to close FX_LIB'); process.exit(1); }
const arrSrc = src.slice(arrStart, mapIdx + 1);

let rows;
try {
  // 纯数据数组，安全 eval
  rows = eval('(' + arrSrc + ')');
} catch (e) {
  console.error('FAIL: FX_LIB 数组无法解析 →', e.message);
  process.exit(1);
}

const CATS = ['math','text','logic','lookup','date','stats','fin','other'];
const issues = [];
const byCat = {};
const seen = new Set();

rows.forEach(r => {
  if (!Array.isArray(r)) { issues.push(`非数组行: ${JSON.stringify(r)}`); return; }
  if (r.length !== 8) { issues.push(`字段数不为8(${r.length}): ${r[0]},${r[1]}`); return; }
  const [cat, name, title, usage, params, example, note, fxInsert] = r;
  if (!CATS.includes(cat)) { issues.push(`未知分类 ${cat}:${name}`); return; }
  byCat[cat] = (byCat[cat] || 0) + 1;
  if (seen.has(name)) { issues.push(`重复函数名: ${name}`); }
  seen.add(name);
  if (typeof title !== 'string' || !title) issues.push(`title缺失: ${name}`);
  if (typeof usage !== 'string' || !usage) issues.push(`usage缺失: ${name}`);
  if (!Array.isArray(params)) issues.push(`params非数组: ${name}`);
  else params.forEach(p => { if (!Array.isArray(p) || p.length < 3) issues.push(`param结构错误: ${name}/${JSON.stringify(p)}`); });
  if (typeof example !== 'string' || !example.startsWith('=')) issues.push(`example非公式: ${name} -> ${example}`);
  if (typeof fxInsert !== 'string' || !fxInsert.startsWith('=')) issues.push(`fxInsert非公式: ${name} -> ${fxInsert}`);
});

console.log('=== FX_LIB 结构校验 ===');
console.log('函数总数:', rows.length);
console.log('按分类:', JSON.stringify(byCat));
if (issues.length) {
  console.error(`\nFAIL: ${issues.length} 个问题`);
  issues.forEach(i => console.error('  -', i));
  process.exit(1);
} else {
  console.log('OK: 无重复、字段完整、示例公式合法');
}

// 版本映射完整性：用 Function 构造解析 FX_VERSION 对象字面量
const vs = src.indexOf('const FX_VERSION');
const objStart = src.indexOf('{', vs);
const objEndMatch = src.indexOf('};', objStart);
let versionMap = {};
if (objStart > 0 && objEndMatch > objStart) {
  const objSrc = src.slice(objStart, objEndMatch + 1);
  try { versionMap = new Function('return ' + objSrc + ';')(); }
  catch (e) { console.warn('FX_VERSION 解析失败:', e.message); }
} else { console.warn('FX_VERSION 对象边界未找到'); }
const inLib = new Set(rows.map(r => r[1]));
const verOnly = Object.keys(versionMap).filter(k => !inLib.has(k));
const libDefault2007 = rows.filter(r => !versionMap[r[1]] && !['SUM','AVERAGE'].includes(r[1]));
console.log('\n=== 版本映射完整性 ===');
console.log('FX_VERSION 定义函数数:', Object.keys(versionMap).length);
console.log('FX_VERSION 有、但 FX_LIB 未收录(提示,可后续扩充):', verOnly.length > 0 ? verOnly.join(', ') : '无');
console.log('FX_LIB 中未标注版本(默认2007)函数数:', libDefault2007.length);