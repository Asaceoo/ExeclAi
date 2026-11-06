const fs = require('fs');
const src = fs.readFileSync('d:/execl1.0/src/ExcelAiAssistant.Addin/Sidebar/index.html','utf8');
const lines = src.split('\n');
for (const k of ['const FX_VERSION','const FX_COMMON','const FX_LIB','function fxVisible','function fxRenderList','function fxRenderDetail','function fxVersionOf','fxInsert']) {
  const idx = lines.findIndex(l => l.includes(k));
  console.log((idx<0?'MISSING':'line '+(idx+1))+'  '+k);
}