// 函数库版本标注审计：核对 FX_LIB 中所有函数是否被正确标注为非-2007 版本。
// 用权威的「已知新版函数清单」逐一在 FX_VERSION 中检查；缺失视为误标 2007，输出待补充清单。
const fs = require('fs');
const src = fs.readFileSync('src/ExcelAiAssistant.Addin/Sidebar/index.html', 'utf8');

function extract(src, startMarker, endMarker, from=0){
  const i = src.indexOf(startMarker, from); if(i<0) throw new Error('未找到 '+startMarker);
  const j = src.indexOf(endMarker, i); if(j<0) throw new Error('未找到 '+endMarker);
  return { code: src.slice(i, j + endMarker.length), end: j };
}

// FX_VERSION
const vFrom = src.indexOf('const FX_VERSION');
const vObj = extract(src, '{', '};', vFrom);
const FX_VERSION = new Function('return ' + vObj.code + ';')();

// FX_LIB 函数名集合
const lStart = src.indexOf('const FX_LIB = [');
const arrStart = src.indexOf('[', lStart);
const mapIdx = src.indexOf('].map', arrStart);
const arrSrc = src.slice(arrStart, mapIdx + 1);
const raw = eval('(' + arrSrc + ')');

// 权威：Excel 2010/2013/2016/2019/2021/365 新增函数（按其首发版本分桶）
// 若 FX_LIB 收录了这些函数但 FX_VERSION 未标注，即为误标 2007。
const NEW_FUNCS = {
  '2010': [
    'AGGREGATE','NETWORKDAYS.INTL','WORKDAY.INTL','RANK.EQ','RANK.AVG','QUARTILE.INC','QUARTILE.EXC',
    'PERCENTILE.INC','PERCENTILE.EXC','PERCENTRANK.INC','PERCENTRANK.EXC','STDEV.S','STDEV.P','VAR.S','VAR.P',
    'CONFIDENCE.NORM','CONFIDENCE.T','COVARIANCE.P','COVARIANCE.S','BETA.DIST','BETA.INV','BINOM.DIST','BINOM.INV',
    'CHISQ.DIST','CHISQ.DIST.RT','CHISQ.INV','CHISQ.INV.RT','CHISQ.TEST','EXPON.DIST','F.DIST','F.DIST.RT','F.INV',
    'F.INV.RT','F.TEST','GAMMA.DIST','GAMMA.INV','GAMMALN.PRECISE','HYPGEOM.DIST','LOGNORM.DIST','LOGNORM.INV',
    'MODE.SNGL','NEGBINOM.DIST','NORM.DIST','NORM.INV','NORM.S.DIST','NORM.S.INV','POISSON.DIST','T.DIST.2T',
    'T.DIST.RT','T.INV.2T','T.TEST','WEIBULL.DIST','Z.TEST','CEILING.PRECISE','FLOOR.PRECISE'
  ],
  '2013': [
    'BINOM.DIST.RANGE','MODE.MULT','T.DIST','T.INV','GAUSS','SKEW.P','SEC','SECH','CSC','CSCH','COT','COTH',
    'ISO.CEILING','ARABIC','BASE','DECIMAL','FORMULATEXT','ISFORMULA','NUMBERVALUE','WEBSERVICE','FILTERXML',
    'BITAND','BITOR','BITXOR','BITLSHIFT','BITRSHIFT','IFNA','XOR','SHEET','SHEETS','MUNIT','PERMUTATIONA','PHI',
    'GAMMA','IMCOSH','IMCOT','IMCSCH','IMSEC','IMSECH','IMSINH','IMTAN','ISOWEEKNUM'
  ],
  '2016': [
    'CONCAT','TEXTJOIN','IFS','SWITCH','FORECAST.LINEAR','FORECAST.ETS','FORECAST.ETS.CONFINT','FORECAST.ETS.SEASONALITY','FORECAST.ETS.STAT'
  ],
  '2019': ['MAXIFS','MINIFS'],
  '2021': [
    'FILTER','SORT','SORTBY','UNIQUE','SEQUENCE','RANDARRAY','XLOOKUP','XMATCH','LET'
  ],
  '365': [
    'LAMBDA','MAP','REDUCE','SCAN','BYROW','BYCOL','MAKEARRAY','ISOMITTED',
    'TOCOL','TOROW','TAKE','DROP','HSTACK','VSTACK','EXPAND','WRAPROWS','WRAPCOLS','CHOOSECOLS','CHOOSEROWS',
    'TEXTBEFORE','TEXTAFTER','TEXTSPLIT','VALUETOTEXT','ARRAYTOTEXT','IMAGE',
    'GROUPBY','PIVOTBY','PERCENTOF','TRIMRANGE','REGEXTEST','REGEXEXTRACT','REGEXREPLACE','TEXTAFTER'
  ]
};

const libNames = new Set(raw.map(r => r[1]));
let ok = true;
let missing = [];
for (const [ver, funcs] of Object.entries(NEW_FUNCS)) {
  for (const fn of funcs) {
    if (libNames.has(fn) && !(fn in FX_VERSION)) {
      missing.push(`${fn}(应为${ver})`);
      ok = false;
    }
  }
}

if (missing.length === 0) {
  console.log('版本标注审计 OK：所有新版函数均已正确标注。');
} else {
  console.log('版本标注审计发现以下函数被误标/缺标注（待补充到 FX_VERSION）：');
  missing.forEach(m => console.log('  ✗ ' + m));
}
process.exit(ok ? 0 : 1);