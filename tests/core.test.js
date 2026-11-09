// 零依赖回归测试：直接从各 HTML 页面的内联脚本中抽取纯函数并在 vm 中运行。
// 运行：node --test tests/
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const read = f => fs.readFileSync(path.join(__dirname, '..', f), 'utf8');
const script = f => { const s = read(f); return s.slice(s.indexOf('<script>') + 8, s.lastIndexOf('</script>')); };
const between = (s, a, b) => { const i = s.indexOf(a), j = s.indexOf(b, i); assert.ok(i >= 0 && j > i, `marker not found: ${a} .. ${b}`); return s.slice(i, j); };
const run = (code, expose, pre = 'const t=k=>k;') => { const ctx = {}; vm.createContext(ctx); vm.runInContext(`${pre}${code};this.X={${expose}}`, ctx); return ctx.X; };

const text = (() => {
  const s = script('text-compare.html');
  return run(between(s, 'const EQ=0', '/* ===================== syntax'), 'diffSeq,inlineDiff,computeDiff');
})();
const excel = (() => {
  const s = script('excel-compare.html');
  const code = between(s, 'function colIndex', 'function parseXlsx') + between(s, 'function parseCsv', '/* ===================== diff core') + between(s, 'const EQ=0', '/* ===================== UI: loading');
  return run(code, 'compare,parseCsv,opts,serialToDate');
})();

const lcs = (a, b) => { const d = Array.from({ length: a.length + 1 }, () => new Int32Array(b.length + 1)); for (let i = a.length - 1; i >= 0; i--) for (let j = b.length - 1; j >= 0; j--) d[i][j] = a[i] === b[j] ? d[i + 1][j + 1] + 1 : Math.max(d[i + 1][j], d[i][j + 1]); return d[0][0]; };
const rnd = n => Math.floor(Math.random() * n);

test('diffSeq: 结果可还原两侧序列，且 EQ 数量等于 LCS 长度', () => {
  for (let it = 0; it < 3000; it++) {
    const al = 1 + rnd(5), A = Array.from({ length: rnd(14) }, () => String(rnd(al))), B = Array.from({ length: rnd(14) }, () => String(rnd(al)));
    let ia = 0, ib = 0, eq = 0;
    for (const [op, i, j] of text.diffSeq(A, B)) {
      if (op === 0) { assert.equal(i, ia); assert.equal(j, ib); assert.equal(A[i], B[j]); ia++; ib++; eq++; }
      else if (op === 1) { assert.equal(i, ia); ia++; }
      else { assert.equal(j, ib); ib++; }
    }
    assert.equal(ia, A.length); assert.equal(ib, B.length); assert.equal(eq, lcs(A, B));
  }
});

test('diffSeq: 另一侧不存在的元素被预先剔除时仍保持正确', () => {
  const A = Array.from({ length: 80 }, (_, i) => (i % 4 ? 'x' + i : 'k' + (i % 7)));
  const B = Array.from({ length: 80 }, (_, i) => (i % 5 ? 'y' + i : 'k' + (i % 7)));
  let ia = 0, ib = 0;
  for (const [op, i, j] of text.diffSeq(A, B)) { if (op === 0) { assert.equal(A[i], B[j]); ia++; ib++; } else if (op === 1) ia++; else ib++; }
  assert.equal(ia, A.length); assert.equal(ib, B.length);
});

test('inlineDiff: 三种粒度下两侧文本都能无损还原（含中文与代理对）', () => {
  const w = ['foo', 'bar', ' ', '中文', '😀x', 'a', 'B', ',', '  '];
  const mk = () => Array.from({ length: rnd(8) }, () => w[rnd(w.length)]).join('');
  for (const lvl of ['smart', 'word', 'char']) for (let it = 0; it < 500; it++) {
    const l = mk(), r = mk(), segs = text.inlineDiff(l, r, lvl, s => s);
    assert.equal(segs.map(s => s.a).join(''), l); assert.equal(segs.map(s => s.b).join(''), r);
  }
});

test('computeDiff: 相同文本无差异，单行修改计为 1 删 1 增', () => {
  const o = { rules: [] };
  const same = text.computeDiff('a\nb\nc', 'a\nb\nc', o); assert.equal(same.nDel + same.nIns, 0);
  const one = text.computeDiff('a\nb\nc', 'a\nX\nc', o); assert.equal(one.nDel, 1); assert.equal(one.nIns, 1);
  const ws = text.computeDiff('a  b', 'a b', { rules: [], hideWs: true }); assert.equal(ws.nDel + ws.nIns, 0);
});

test('excel.compare: 每一行恰好出现一次（表头 / 逐行、排序 / 不排序）', () => {
  const G = rows => ({ rows: rows.map(r => r.map(v => ({ v: String(v), f: null, d: false }))), cols: Math.max(0, ...rows.map(r => r.length)) });
  for (let it = 0; it < 800; it++) {
    const nc = 1 + rnd(4), gen = () => [Array.from({ length: nc }, (_, i) => 'H' + i), ...Array.from({ length: rnd(8) }, () => Array.from({ length: nc }, () => rnd(3)))];
    const A = gen(), B = gen();
    for (const mode of ['table', 'rows']) for (const sortRows of [false, true]) {
      Object.assign(excel.opts, { mode, sortRows });
      const R = excel.compare(G(A), G(B), 0, 0), off = mode === 'table' ? 1 : 0;
      const ua = R.rows.filter(r => r.ai != null).map(r => r.ai), ub = R.rows.filter(r => r.bi != null).map(r => r.bi);
      assert.equal(ua.length, A.length - off); assert.equal(new Set(ua).size, ua.length);
      assert.equal(ub.length, B.length - off); assert.equal(new Set(ub).size, ub.length);
    }
  }
});

test('excel.parseCsv: 引号、转义引号与分隔符自动识别', () => {
  const g = JSON.parse(JSON.stringify(excel.parseCsv('a;b\n"x;1";"y""z"\n1;2').sheets[0].grid.rows.map(r => r.map(c => c.v))));
  assert.deepEqual(g, [['a', 'b'], ['x;1', 'y"z'], ['1', '2']]);
});

test('excel.serialToDate: 日期、日期时间与纯时间', () => {
  assert.equal(excel.serialToDate(45000), '2023-03-15');
  assert.equal(excel.serialToDate(45000.75), '2023-03-15 18:00:00');
  assert.equal(excel.serialToDate(0.5), '12:00:00');
});

test('页面内联脚本语法正确，且 i18n 中英文键一致', () => {
  for (const f of ['text', 'image', 'excel']) {
    const s = script(f + '-compare.html'); assert.doesNotThrow(() => new Function(s), f);
    const I = run(s.slice(0, s.indexOf('let lang=')), 'I18N', '').I18N;
    assert.deepEqual([...Object.keys(I.en)].sort(), [...Object.keys(I.zh)].sort(), f);
  }
});

test('excel.parseCsv: 字段中间的引号按字面处理，分隔符识别忽略引号内的内容，兼容仅 \r 换行', () => {
  const g = s => JSON.parse(JSON.stringify(excel.parseCsv(s).sheets[0].grid.rows.map(r => r.map(c => c.v))));
  assert.deepEqual(g('id,desc\n1,5" pipe\n2,x'), [['id', 'desc'], ['1', '5" pipe'], ['2', 'x']]);
  assert.deepEqual(g('"a,b,c";d;e\n1;2;3'), [['a,b,c', 'd', 'e'], ['1', '2', '3']]);
  assert.deepEqual(g('a;b\r1;2'), [['a', 'b'], ['1', '2']]);
  assert.deepEqual(g('a,"x ""q"" y",c'), [['a', 'x "q" y', 'c']]);
});
