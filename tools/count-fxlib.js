const fs = require('fs');
const src = fs.readFileSync('d:/execl1.0/src/ExcelAiAssistant.Addin/Sidebar/index.html', 'utf8');
const start = src.indexOf('const FX_LIB');
if (start < 0) { console.log('FX_LIB not found'); process.exit(1); }
// find the closing of the array: locate '.map(' after start
const mapIdx = src.indexOf('.map', start);
if (mapIdx < 0) { console.log('.map not found'); process.exit(1); }
const arrSrc = src.slice(start, mapIdx);
// count top-level array rows by looking for lines that start with  ["cat"
const rows = arrSrc.match(/\[\s*"(math|text|logic|lookup|date|stat|fin|info|more|megafunc|all)"/g);
const allRows = arrSrc.match(/\[\s*"/g);
const verRows = arrSrc.match(/\[\s*"(math|text|logic|lookup|date|stat|fin|info|more)"/g);
const namedCats = {};
(arrSrc.match(/\[\s*"(math|text|logic|lookup|date|stat|fin|info|more)"/g) || []).forEach(r => {
  const cat = r.replace(/[\["\s]/g,'');
  namedCats[cat] = (namedCats[cat]||0)+1;
});
console.log('rows start (cat-literal):', rows ? rows.length : 0);
console.log('all bracket-string rows:', allRows ? allRows.length : 0);
console.log('by category:', JSON.stringify(namedCats));
// FX_VERSION size
const vs = src.indexOf('const FX_VERSION');
const ve = src.indexOf(';', vs);
console.log('FX_VERSION block len:', (ve>vs?ve-vs:0));
// FX_COMMON
const cs = src.indexOf('const FX_COMMON');
const ce = src.indexOf(';', cs);
console.log('FX_COMMON block len:', (ce>cs?ce-cs:0), cs>0?'found':'missing');