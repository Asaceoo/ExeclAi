// 提取 index.html 的全部 <script> 块并做 JS 语法校验
const fs = require('fs');
const vm = require('vm');
const src = fs.readFileSync('src/ExcelAiAssistant.Addin/Sidebar/index.html','utf8');
const re = /<script[^>]*>([\s\S]*?)<\/script>/g;
let m, i = 0, fail = 0;
while ((m = re.exec(src)) !== null) {
  i++;
  const body = m[1];
  if (!body.trim()) continue;
  try {
    new vm.Script(body, { filename: 'inline-script-' + i + '.js' });
    console.log('script[' + i + '] 语法 OK (' + body.length + ' 字符)');
  } catch (e) {
    fail++;
    console.error('script[' + i + '] 语法错误:', e.message);
  }
}
console.log(fail ? ('\nFAIL: ' + fail + ' 个脚本语法错误') : '\nOK: 全部脚本语法通过');
process.exit(fail ? 1 : 0);