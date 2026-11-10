// 零依赖回归测试：直接从各 HTML 页面的内联脚本中抽取纯函数并在 vm 中运行。
// 运行：npm test（即 node --test）
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
  return run(between(s, 'const EQ=0', '/* ===================== syntax'), 'diffSeq,inlineDiff,computeDiff,makeNorm');
})();
const excel = (() => {
  const s = script('excel-compare.html');
  const code = between(s, 'function colIndex', 'function parseXlsx') + between(s, 'function parseCsv', '/* ===================== diff core') + between(s, 'const EQ=0', '/* ===================== UI: loading');
  return run(code, 'compare,parseCsv,opts,serialToDate,serialToElapsed,isoDateCell,isZipSig,pairBlock,normalize,inlineDiff,shiftFormula');
})();

const inflate = (() => {
  const s = script('excel-compare.html');
  return run(between(s, 'function inflate', '/* ===================== XLSX parsing'), 'inflate').inflate;
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

test('excel.serialToDate: 1900 系统中 1900-03-01 之前的序列号（Excel 的闰年错误）', () => {
  assert.equal(excel.serialToDate(1), '1900-01-01');
  assert.equal(excel.serialToDate(59), '1900-02-28');
  assert.equal(excel.serialToDate(61), '1900-03-01');
  assert.equal(excel.serialToDate(1, true), '1904-01-02');
});

test('excel.isZipSig: 以 PK 开头的 CSV 不会被当作 xlsx', () => {
  assert.equal(excel.isZipSig(Buffer.from([0x50, 0x4B, 3, 4])), true);
  assert.equal(excel.isZipSig(Buffer.from([0x50, 0x4B, 5, 6])), true);
  assert.equal(excel.isZipSig(Buffer.from('PK,name\n1,a')), false);
  assert.equal(excel.isZipSig(Buffer.from('a,b')), false);
});

test('excel.pairBlock: 配对不交叉，A / B 两侧的行序都保持原有顺序且各出现一次', () => {
  const A = [['a', 'b', 'c'], ['x', 'y', 'z']], B = [['x', 'y', 'q'], ['a', 'b', 'd']];
  const out = excel.pairBlock([0, 1], [0, 1], null, null, A, B);
  assert.deepEqual(Array.from(out.filter(e => e.bi != null), e => e.bi), [0, 1]);
  for (let it = 0; it < 1500; it++) {
    const na = rnd(12), nb = rnd(12), mk = n => Array.from({ length: n }, () => Array.from({ length: 3 }, () => String(rnd(3))));
    const ca = mk(na), cb = mk(nb), da = ca.map((_, i) => i), db = cb.map((_, i) => i);
    const r = excel.pairBlock(da, db, null, null, ca, cb);
    const as = Array.from(r.filter(e => e.ai != null), e => e.ai), bs = Array.from(r.filter(e => e.bi != null), e => e.bi);
    assert.deepEqual(as, da); assert.deepEqual(bs, db);
  }
});

test('computeDiff: 行数不等的改动块按相似度配对，且每行恰好出现一次、顺序不变', () => {
  const o = { rules: [] };
  const d = text.computeDiff('foo bar baz\nqux quux corge\nhello world one', 'hello world 1', o);
  const pr = Array.from(d.blocks[0].pairs, p => [p.l, p.r]);
  assert.deepEqual(pr, [[0, null], [1, null], [2, 0]]);
  const w = ['a b', 'c d', 'e f', 'a c', 'x y z', 'p q'];
  for (let it = 0; it < 800; it++) {
    const gen = () => Array.from({ length: rnd(9) }, () => w[rnd(w.length)] + ' ' + rnd(4)).join('\n');
    const r = text.computeDiff(gen(), gen(), o);
    for (const b of r.blocks) if (b.type === 'chg') {
      assert.deepEqual(Array.from(b.pairs.filter(p => p.l != null), p => p.l), Array.from(b.del));
      assert.deepEqual(Array.from(b.pairs.filter(p => p.r != null), p => p.r), Array.from(b.ins));
    }
  }
});

test('excel.inflate: 与 zlib 的 deflate 输出逐字节一致（存储 / 固定 / 动态 Huffman 块）', () => {
  const zlib = require('node:zlib'), crypto = require('node:crypto');
  const samples = [
    Buffer.alloc(0), Buffer.from('a'), Buffer.from('hello hello hello '.repeat(400)), crypto.randomBytes(70000),
    Buffer.from(Array.from({ length: 50000 }, (_, i) => (i * i) % 251)),
    Buffer.from(JSON.stringify(Array.from({ length: 3000 }, (_, i) => ({ id: i, name: 'n' + (i % 13) })))),
  ];
  for (const d of samples) for (const level of [0, 1, 6, 9]) for (const strategy of [0, 1, 2, 3, 4]) {
    const z = zlib.deflateRawSync(d, { level, strategy });
    assert.ok(Buffer.from(inflate(new Uint8Array(z))).equals(d), `level ${level} strategy ${strategy} size ${d.length}`);
  }
});

test('excel.serialToElapsed: 时长格式的小时数可以超过 24，负值带负号', () => {
  assert.equal(excel.serialToElapsed(1.25), '30:00:00');
  assert.equal(excel.serialToElapsed(0.5), '12:00:00');
  assert.equal(excel.serialToElapsed(100 / 24 + 61 / 86400), '100:01:01');
  assert.equal(excel.serialToElapsed(-0.5), '-12:00:00');
});

test('inlineDiff: 超长行只裁剪公共前后缀，仍能无损还原且不会卡死', () => {
  const words = Array.from({ length: 30000 }, (_, i) => 'w' + i);
  const l = words.join(' '), r = words.slice(0, 15000).concat(['CHANGED', 'tokens'], words.slice(15003)).join(' ');
  const t0 = Date.now();
  for (const level of ['smart', 'word', 'char']) {
    const segs = text.inlineDiff(l, r, level, s => s);
    assert.equal(segs.map(x => x.a).join(''), l);
    assert.equal(segs.map(x => x.b).join(''), r);
    assert.ok(segs.some(x => x.t === 1) && segs.some(x => x.t === 2));
  }
  assert.ok(Date.now() - t0 < 3000, '超长行的行内差异应在 3 秒内完成');
});

const img = (() => {
  const s = script('image-compare.html');
  const code = between(s, 'function dilate', 'function scan') + between(s, 'function fft1d', 'function grayPatch') + between(s, 'function phaseCorr', 'async function autoAlign') + between(s, 'function invertAlign', "$('swapBtn')");
  return run(code, 'dilate,detect,phaseCorr,invertAlign', 'const defAlign=()=>({dx:0,dy:0,rot:0,scale:1,fh:false,fv:false});');
})();
const lcg = seed => () => (seed = (seed * 1664525 + 1013904223) >>> 0) / 4294967296;

test('image.dilate: 与逐像素暴力膨胀结果一致', () => {
  const r = lcg(7), W = 37, H = 23, m = new Uint8Array(W * H);
  for (let i = 0; i < m.length; i++) m[i] = r() < 0.04 ? 1 : 0;
  for (const rad of [1, 2, 5]) {
    const d = img.dilate(m, W, H, rad);
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      let e = 0;
      for (let yy = Math.max(0, y - rad); yy <= Math.min(H - 1, y + rad); yy++) for (let xx = Math.max(0, x - rad); xx <= Math.min(W - 1, x + rad); xx++) if (m[yy * W + xx]) e = 1;
      assert.equal(d[y * W + x], e, `r=${rad} (${x},${y})`);
    }
  }
});

test('image.detect: 分组半径决定相邻差异块是否合并，过小的块被丢弃', () => {
  const W = 60, H = 30, m = new Uint8Array(W * H), fill = (x0, y0, w, h) => { for (let y = y0; y < y0 + h; y++) for (let x = x0; x < x0 + w; x++) m[y * W + x] = 1; };
  fill(5, 5, 6, 6); fill(14, 5, 6, 6); fill(45, 20, 5, 5); fill(30, 2, 1, 1);
  const o = { minBlobArea: 4, maxBlobAreaFraction: 1 }, rect = { x0: 0, y0: 0, x1: W, y1: H };
  const sep = [], grouped = [];
  img.detect(m, W, H, o, rect, 0, sep);
  img.detect(m, W, H, o, rect, 3, grouped);
  assert.equal(sep.length, 3);
  assert.equal(grouped.length, 2);
  const big = grouped.find(r => r.w > 10);
  assert.equal([big.x, big.y, big.w, big.h, big.px].join(), '5,5,15,6,72');
});

test('image.phaseCorr: 峰值给出 B 相对 A 的平移（符号为 -d）', () => {
  const N = 64, r = lcg(3), a = new Float64Array(N * N);
  for (let i = 0; i < a.length; i++) a[i] = r();
  for (const [dx, dy] of [[3, 5], [-4, 2], [0, 0], [7, -9]]) {
    const b = new Float64Array(N * N);
    for (let y = 0; y < N; y++) for (let x = 0; x < N; x++) b[((y + dy + N) % N) * N + ((x + dx + N) % N)] = a[y * N + x];
    const p = img.phaseCorr(a, b, N);
    assert.ok(Math.abs(p.dx + dx) < 0.05 && Math.abs(p.dy + dy) < 0.05, `shift ${dx},${dy} -> ${p.dx},${p.dy}`);
  }
});

test('image.invertAlign: 交换两张图后的对齐参数与原参数互为逆变换', () => {
  const A = { w: 300, h: 200 }, B = { w: 260, h: 210 }, al = { dx: 12, dy: -7, rot: 3, scale: 1.1, fh: false, fv: false };
  const map = (p, a, b) => { const th = a.rot * Math.PI / 180, c = Math.cos(th), s = Math.sin(th), x = (p.x - b.w / 2) * a.scale, y = (p.y - b.h / 2) * a.scale; return { x: b.w / 2 + a.dx + c * x - s * y, y: b.h / 2 + a.dy + s * x + c * y }; };
  const p = { x: 50, y: 80 }, q = map(p, al, B), back = map(q, img.invertAlign(al, A, B), A);
  assert.ok(Math.abs(back.x - p.x) < 1e-6 && Math.abs(back.y - p.y) < 1e-6);
});

test('decodeBuf: UTF-8 优先，其次严格的 GB18030，其余（如 Latin-1 / windows-1252）不再被误读成 GBK 乱码', () => {
  for (const f of ['text-compare.html', 'excel-compare.html']) {
    const code = between(script(f), 'function decodeBuf', f === 'text-compare.html' ? '\ntry{mergeOpts' : '\nasync function loadFile');
    const ctx = { TextDecoder }; vm.createContext(ctx); vm.runInContext(code + ';this.decodeBuf=decodeBuf', ctx);
    const dec = b => ctx.decodeBuf(Uint8Array.from(b).buffer);
    assert.equal(dec(Buffer.from('héllo 中文', 'utf8')), 'héllo 中文', f);
    assert.equal(dec([0xD6, 0xD0, 0xCE, 0xC4, 0x20, 0x74]), '中文 t', f); // GBK “中文 t”
    assert.equal(dec(Buffer.from('café au lait\nnaïve', 'latin1')), 'café au lait\nnaïve', f);
  }
});

test('excel.normalize: 日期统一成 ISO；x/y/yyyy 在第二位大于 12 时按 月/日/年，其余按 日/月/年', () => {
  assert.equal(excel.normalize('2024/3/5'), '2024-03-05');
  assert.equal(excel.normalize('03/15/2024'), '2024-03-15');
  assert.equal(excel.normalize('13/02/2024'), '2024-02-13');
  assert.equal(excel.normalize('03/04/2024'), '2024-04-03');
  assert.equal(excel.normalize('2024-03-15 8:05'), '2024-03-15 08:05:00');
  assert.equal(excel.normalize('1.2.3'), '1.2.3');
  assert.equal(excel.normalize('2024-03-05 00:00:00'), '2024-03-05');
  assert.equal(excel.normalize('2024-03-05 00:00'), '2024-03-05');
  assert.equal(excel.normalize('2024-03-05 08:05:09'), '2024-03-05 08:05:09');
  assert.equal(excel.normalize('9:30'), '09:30:00');
  assert.equal(excel.normalize('09:30:15'), '09:30:15');
  assert.equal(excel.normalize('1:2'), '1:2');
  assert.equal(excel.normalize('03/15/2024'), '2024-03-15');
});

test('inflate: 无压缩块的长度与补码不符或数据被截断时抛错，而不是静默补零', () => {
  assert.equal(Buffer.from(inflate(Uint8Array.from([0x01, 0x03, 0x00, 0xfc, 0xff, 0x61, 0x62, 0x63]))).toString(), 'abc');
  assert.throws(() => inflate(Uint8Array.from([0x01, 0x10, 0x00, 0xef, 0xff, 0x61])), /badZip/);
  assert.throws(() => inflate(Uint8Array.from([0x01, 0x03, 0x00, 0x00, 0x00, 0x61, 0x62, 0x63])), /badZip/);
});

test('makeNorm: 忽略规则与正文使用相同的引号 / 小写规范化', () => {
  const lower = text.computeDiff('Foo bar', 'bar', { rules: ['Foo '], lower: true });
  assert.equal(lower.nDel + lower.nIns, 0);
  const quote = text.computeDiff('say \u201chi\u201d now', 'say now', { rules: ['\u201chi\u201d '], igQuotes: true });
  assert.equal(quote.nDel + quote.nIns, 0);
  const plain = text.computeDiff('Foo bar', 'bar', { rules: ['Foo '] });
  assert.equal(plain.nDel + plain.nIns, 0);
  const diff = text.computeDiff('Foo bar', 'baz', { rules: ['Foo '], lower: true });
  assert.equal(diff.nDel + diff.nIns, 2);
});

test('excel.compare: 整行新增 / 删除的内容所在列算作有改动（隐藏未修改的列时不能丢数据）', () => {
  const G = rows => ({ rows: rows.map(r => r.map(v => ({ v: String(v), f: null, d: false }))), cols: Math.max(0, ...rows.map(r => r.length)) });
  Object.assign(excel.opts, { mode: 'table', sortRows: false, matchHeaders: true });
  const add = excel.compare(G([['K', 'V', 'W'], [1, 2, 3]]), G([['K', 'V', 'W'], [1, 2, 3], [4, 5, '']]), 0, 0);
  assert.deepEqual([...add.colChanged].sort(), [0, 1]);
  const del = excel.compare(G([['K', 'V', 'W'], [1, 2, 3], ['', '', 9]]), G([['K', 'V', 'W'], [1, 2, 3]]), 0, 0);
  assert.deepEqual([...del.colChanged], [2]);
  const same = excel.compare(G([['K', 'V'], [1, 2]]), G([['K', 'V'], [1, 2]]), 0, 0);
  assert.equal(same.colChanged.size, 0);
});

test('excel.parseCsv: 末尾没有换行的 "" 空字段行不会丢失', () => {
  const rows = s => JSON.parse(JSON.stringify(excel.parseCsv(s, ',').sheets[0].grid.rows.map(r => r.map(c => c.v))));
  assert.deepEqual(rows('a\n""'), [['a'], ['']]);
  assert.deepEqual(rows('a,""'), [['a', '']]);
  assert.deepEqual(rows('a\n"x"'), [['a'], ['x']]);
  assert.deepEqual(rows('a\n'), [['a']]);
});

test('excel.inlineDiff: 超长单元格只裁剪公共前后缀，仍能无损还原且不会卡死', () => {
  const rndWord = () => Math.random().toString(36).slice(2, 8);
  const mid = n => Array.from({ length: n }, rndWord).join(' ');
  const l = 'head ' + mid(12000) + ' tail', r = 'head ' + mid(12000) + ' tail';
  const t0 = Date.now(), segs = excel.inlineDiff(l, r);
  assert.ok(Date.now() - t0 < 3000);
  assert.equal(segs.map(s => s.a).join(''), l); assert.equal(segs.map(s => s.b).join(''), r);
  assert.equal(segs[0].t, 0); assert.equal(segs[segs.length - 1].t, 0);
  const small = excel.inlineDiff('abcdef', 'abXdef');
  assert.equal(small.map(s => s.a).join(''), 'abcdef'); assert.equal(small.map(s => s.b).join(''), 'abXdef');
});

test('excel.shiftFormula: 共享公式按位移平移相对引用，$ 锁定、字符串、函数名与带引号的工作表名保持不变', () => {
  const s = excel.shiftFormula;
  assert.equal(s('A1*2', 1, 0), 'A2*2');
  assert.equal(s('SUM($A1:B$2)', 2, 1), 'SUM($A3:C$2)');
  assert.equal(s('"A1"&A1', 1, 0), '"A1"&A2');
  assert.equal(s('LOG10(A1)', 1, 0), 'LOG10(A2)');
  assert.equal(s("Sheet2!A1+'My A1'!B2", 1, 1), "Sheet2!B2+'My A1'!C3");
  assert.equal(s('1E5+A1', 0, 1), '1E5+B1');
  assert.equal(s('Z9+AA1', 0, 1), 'AA9+AB1');
  assert.equal(s('A1+B2', 0, 0), 'A1+B2');
  assert.equal(s('A1', -1, 0), 'A1');
});

test('diffSeq: 差异量很大且重复元素很多时有代价上限，结果仍可还原两侧序列且很快返回', () => {
  const al = 20, A = Array.from({ length: 8000 }, () => 'k' + rnd(al)), B = Array.from({ length: 8000 }, () => 'k' + rnd(al));
  const t0 = Date.now(), ops = text.diffSeq(A, B);
  assert.ok(Date.now() - t0 < 5000, 'diffSeq too slow: ' + (Date.now() - t0) + 'ms');
  let ia = 0, ib = 0;
  for (const [op, i, j] of ops) {
    if (op === 0) { assert.equal(i, ia); assert.equal(j, ib); assert.equal(A[i], B[j]); ia++; ib++; }
    else if (op === 1) { assert.equal(i, ia); ia++; } else { assert.equal(j, ib); ib++; }
  }
  assert.equal(ia, A.length); assert.equal(ib, B.length);
});

test('diffSeq: 差异量在上限以内时仍然是最优的（EQ 数量等于 LCS 长度）', () => {
  for (let it = 0; it < 5; it++) {
    const A = Array.from({ length: 700 }, () => String(rnd(3))), B = Array.from({ length: 700 }, () => String(rnd(3)));
    assert.equal(text.diffSeq(A, B).filter(o => o[0] === 0).length, lcs(A, B));
  }
});

// 导出的 .xlsx 用的是自带的 STORE 压缩 zip 写入器：用页面自己的读取器读回，保证容器格式（本地头 / 中央目录 / CRC）自洽。
const zipLib = (() => {
  const s = script('excel-compare.html');
  const code = between(s, 'function inflate', '/* ===================== XLSX parsing') + between(s, 'const CRC_T', 'const xe=');
  const ctx = { Blob, TextEncoder, TextDecoder, Date };
  vm.createContext(ctx);
  vm.runInContext(`const t=k=>k;${code};this.X={readZip,zipStore,crc32}`, ctx);
  return ctx.X;
})();

test('zipStore → readZip: 写出的 zip 能被读取器原样读回（含 UTF-8 文件名、中文内容与较大文件）', async () => {
  const enc = new TextEncoder();
  const big = 'x'.repeat(200000) + '尾';
  const files = [{ name: 'a.txt', data: enc.encode('你好 hello') }, { name: 'dir/文件.xml', data: enc.encode(big) }, { name: 'empty', data: new Uint8Array(0) }];
  const buf = await zipLib.zipStore(files).arrayBuffer();
  const z = zipLib.readZip(buf);
  assert.deepEqual([...z.names], ['a.txt', 'dir/文件.xml', 'empty']);
  assert.equal(z.text('a.txt'), '你好 hello');
  assert.equal(z.text('dir/文件.xml'), big);
  assert.equal(z.get('empty').length, 0);
  assert.equal(z.get('missing'), null);
  assert.equal(zipLib.crc32(enc.encode('123456789')), 0xCBF43926);
});

test('excel.parseCsv: 按 RFC 4180 序列化的随机表格能原样解析回来（含逗号、引号、换行、CRLF）', () => {
  let seed = 12345;
  const rand = n => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
  const alpha = ['a', 'b', ',', '"', '\n', '\r\n', ' ', 'x', '1', ';', '\t', 'é'];
  const norm = g => JSON.stringify(g.map(r => r.map(v => v.replace(/\r\n/g, '\n'))));
  for (let it = 0; it < 3000; it++) {
    const R = 1 + rand(4), C = 1 + rand(4);
    const g = Array.from({ length: R }, () => Array.from({ length: C }, () => Array.from({ length: rand(4) }, () => alpha[rand(alpha.length)]).join('')));
    const text = g.map(r => r.map(v => (/[",\r\n]/.test(v) || (v === '' && C === 1)) ? '"' + v.replace(/"/g, '""') + '"' : v).join(',')).join(rand(2) ? '\r\n' : '\n') + (rand(2) ? '\n' : '');
    const got = excel.parseCsv(text, ',').sheets[0].grid.rows.map(r => r.map(c => c.v));
    assert.equal(norm(got), norm(g), JSON.stringify(text));
  }
});

test('excel.pairBlock: 改动块很大时按位置配对也要求相似度，完全不同的两行记为删除 + 新增', () => {
  const n = 60, A = [], B = [];
  for (let i = 0; i < n; i++) { A.push(['a' + i, 'x', 'y']); B.push(i % 2 ? ['a' + i, 'x', 'z'] : ['b' + i, 'p', 'q']); }
  const idx = Array.from({ length: n }, (_, i) => i);
  const out = excel.pairBlock(idx, idx, null, null, A, B);
  assert.equal(out.filter(e => e.type === 'chg').length, n / 2);
  assert.equal(out.filter(e => e.type === 'del').length, n / 2);
  assert.equal(out.filter(e => e.type === 'ins').length, n / 2);
  const seenA = out.filter(e => e.ai != null).map(e => e.ai), seenB = out.filter(e => e.bi != null).map(e => e.bi);
  assert.deepEqual([...seenA], idx); assert.deepEqual([...seenB], idx);
});

test('excel.inflate: 头部声明的解压后大小被伪造得极大时不会按它分配内存，仍能正确解压', () => {
  const zlib = require('node:zlib');
  const raw = Buffer.from('hello hello hello '.repeat(5000));
  const z = zlib.deflateRawSync(raw);
  for (const claimed of [0xFFFFFFFF, 1, 0]) assert.equal(Buffer.compare(Buffer.from(inflate(new Uint8Array(z), claimed)), raw), 0);
});

test('inlineDiff: 自定义忽略规则与忽略空白在行内差异中不再被标为删改，且仍可无损还原', () => {
  for (const level of ['smart', 'word', 'char']) {
    const s1 = text.inlineDiff('Foo bar 1', 'bar 2', level, text.makeNorm({ rules: ['Foo '] }));
    assert.equal(s1.map(s => s.a).join(''), 'Foo bar 1');
    assert.equal(s1.map(s => s.b).join(''), 'bar 2');
    assert.ok(!s1.some(s => s.t !== 0 && s.a.includes('Foo')), `rule "Foo " should not be highlighted in ${level}`);

    const s2 = text.inlineDiff('  hello   world 1', 'hello world 2', level, text.makeNorm({ rules: [], hideWs: true }));
    assert.equal(s2.map(s => s.a).join(''), '  hello   world 1');
    assert.equal(s2.map(s => s.b).join(''), 'hello world 2');
    assert.ok(!s2.some(s => s.t !== 0 && /^\s+$/.test(s.a || s.b)), `ignored whitespace should not be highlighted in ${level}`);
  }
});

test('excel.compare: 列顺序变动、按位置匹配时表头修改、以及无公共列时的行差异处理', () => {
  const G = rows => ({ rows: rows.map(r => r.map(v => ({ v: String(v), f: null, d: false }))), cols: Math.max(0, ...rows.map(r => r.length)) });
  // 1) 仅列移动时不应误判为完全相同
  Object.assign(excel.opts, { mode: 'table', matchHeaders: true, sortRows: false, sortCols: false });
  const mov = excel.compare(G([['A', 'B'], [1, 2]]), G([['B', 'A'], [2, 1]]), 0, 0);
  assert.equal(mov.identical, false);
  assert.equal(mov.stats.colsMov, 1);

  // 2) matchHeaders 关闭时表头文字改动不应被丢弃或误判相同
  Object.assign(excel.opts, { mode: 'table', matchHeaders: false, sortRows: false, sortCols: false });
  const hdr = excel.compare(G([['ID', 'Price'], [1, 10]]), G([['ID', 'Cost'], [1, 10]]), 0, 0);
  assert.equal(hdr.identical, false);
  assert.equal(hdr.columns[1].name, 'Price → Cost');
  assert.ok(hdr.colChanged.has(1));

  // 3) 表头完全无交集时，两侧数据行不应被当作未修改的相同行配对
  Object.assign(excel.opts, { mode: 'table', matchHeaders: true, sortRows: false, sortCols: false });
  const disj = excel.compare(G([['A'], [1], [2]]), G([['B'], [3], [4]]), 0, 0);
  assert.equal(disj.identical, false);
  assert.equal(disj.stats.rowsDel, 2);
  assert.equal(disj.stats.rowsIns, 2);
  assert.ok(disj.rows.every(r => r.type !== 'eq'));
});

test('excel.normalize: 支持带时间的日月年/月日年日期，且过滤越界的伪日期与伪时间', () => {
  Object.assign(excel.opts, { normDates: true, ignoreWs: false, ignoreCase: false });
  assert.equal(excel.normalize('03/15/2024 8:05'), '2024-03-15 08:05:00');
  assert.equal(excel.normalize('15/03/2024 00:00:00'), '2024-03-15');
  assert.equal(excel.normalize('2024.99.88'), '2024.99.88');
  assert.equal(excel.normalize('99/88/2024'), '99/88/2024');
  assert.equal(excel.normalize('12:99'), '12:99');
  assert.equal(excel.normalize('25:00'), '25:00');
});

test('image.phaseCorr: 启用参考图 FFT 缓存时与未缓存结果一致', () => {
  const N = 32, r = lcg(11), a = new Float64Array(N * N), b = new Float64Array(N * N);
  for (let i = 0; i < a.length; i++) a[i] = r();
  for (let y = 0; y < N; y++) for (let x = 0; x < N; x++) b[((y + 3 + N) % N) * N + ((x - 2 + N) % N)] = a[y * N + x];
  const p1 = img.phaseCorr(a, b, N);
  a._c = 1;
  const p2 = img.phaseCorr(a, b, N);
  const p3 = img.phaseCorr(a, b, N);
  assert.ok(a._fft && a._fft.N === N);
  assert.deepEqual(p1, p2);
  assert.deepEqual(p1, p3);
});

test('image.phaseCorr: 纯色 / 空白图片相关面全零时 psr 为 0，而不是 NaN（否则 autoAlign 的阈值判断会误报对齐成功）', () => {
  const z = new Float64Array(32 * 32);
  const p = img.phaseCorr(z, z, 32);
  assert.ok(Number.isFinite(p.psr) && Number.isFinite(p.peak));
  assert.equal(p.psr, 0);
  assert.equal(p.peak, 0);
});

test('excel.compare: 按行排序只以共有列为键，删除 / 新增的列不打乱两侧的行对齐', () => {
  const G = rows => ({ rows: rows.map(r => r.map(v => ({ v: String(v), f: null, d: false }))), cols: Math.max(0, ...rows.map(r => r.length)) });
  Object.assign(excel.opts, { mode: 'table', matchHeaders: true, sortRows: true, sortCols: false });
  const R = excel.compare(G([['X', 'K'], ['3', 'a'], ['1', 'b'], ['2', 'c']]), G([['K', 'Y'], ['c', 'z'], ['a', 'x'], ['b', 'y']]), 0, 0);
  assert.equal(R.stats.rowsDel, 0); assert.equal(R.stats.rowsIns, 0); assert.equal(R.stats.rowsChg, 0);
  assert.ok(R.rows.every(r => r.type === 'eq'));
  Object.assign(excel.opts, { sortRows: false });
});

test('excel.isoDateCell: t="d" 的 ISO 日期与日期序列号的显示一致', () => {
  assert.equal(excel.isoDateCell('2024-03-05T00:00:00'), excel.serialToDate(45356));
  assert.equal(excel.isoDateCell('2024-03-05T00:00:00Z'), '2024-03-05');
  assert.equal(excel.isoDateCell('2024-03-05T08:05:09.000'), '2024-03-05 08:05:09');
  assert.equal(excel.isoDateCell('2024-03-05T08:05'), '2024-03-05 08:05:00');
  assert.equal(excel.isoDateCell('10:30:00'), excel.serialToDate(0.4375));
  assert.equal(excel.isoDateCell('not a date'), 'not a date');
  assert.equal(excel.isoDateCell(null), '');
});

test('image.invertAlign: 带翻转的对齐参数求逆后仍能把点映射回原位', () => {
  // 与 drawAligned 相同的映射：以图片中心为原点，先翻转、再缩放、再旋转，最后平移
  const map = (al, w, h, [x, y]) => {
    const th = al.rot * Math.PI / 180, c = Math.cos(th), s = Math.sin(th);
    const px = (x - w / 2) * al.scale * (al.fh ? -1 : 1), py = (y - h / 2) * al.scale * (al.fv ? -1 : 1);
    return [w / 2 + al.dx + c * px - s * py, h / 2 + al.dy + s * px + c * py];
  };
  const r = lcg(5), a = { w: 640, h: 480 }, b = { w: 500, h: 700 };
  for (const [fh, fv] of [[false, false], [true, false], [false, true], [true, true]]) for (let it = 0; it < 20; it++) {
    const al = { dx: r() * 80 - 40, dy: r() * 80 - 40, rot: r() * 20 - 10, scale: 0.7 + r() * 0.6, fh, fv };
    const inv = img.invertAlign(al, a, b);
    assert.equal(inv.fh, fh); assert.equal(inv.fv, fv);
    const p = [r() * b.w, r() * b.h], q = map(al, b.w, b.h, p), back = map(inv, a.w, a.h, q);
    assert.ok(Math.abs(back[0] - p[0]) < 1e-6 && Math.abs(back[1] - p[1]) < 1e-6, `${fh},${fv}: ${back} != ${p}`);
  }
});

test('excel.serialToDate: 贴近午夜的纯时间按最近整秒回绕为 00:00:00，与 Excel 及日期时间路径一致', () => {
  assert.equal(excel.serialToDate(0.9999999, false), '00:00:00');
  assert.equal(excel.serialToDate(0.5, false), '12:00:00');
  assert.equal(excel.serialToDate(45443.9999999653, false), '2024-06-01');
});

test('excel.serialToDate: 伪造的超大序列号不会显示成 NaN-NaN-NaN', () => {
  const r = excel.serialToDate(1e300, false);
  assert.ok(!/NaN/.test(r));
});

test('excel.compare: 开启“对列排序”后，仅列顺序不同的两张表视为相同，不再报告列移动', () => {
  const G = rows => ({ rows: rows.map(r => r.map(v => ({ v: String(v), f: null, d: false }))), cols: Math.max(0, ...rows.map(r => r.length)) });
  Object.assign(excel.opts, { mode: 'table', matchHeaders: true, sortRows: false, sortCols: true });
  const R = excel.compare(G([['A', 'B', 'C'], [1, 2, 3]]), G([['C', 'A', 'B'], [3, 1, 2]]), 0, 0);
  assert.equal(R.stats.colsMov, 0);
  assert.equal(R.identical, true);
  Object.assign(excel.opts, { sortCols: false });
  assert.equal(excel.compare(G([['A', 'B', 'C'], [1, 2, 3]]), G([['C', 'A', 'B'], [3, 1, 2]]), 0, 0).stats.colsMov, 1);
});

test('excel.compare: 整列为空的尾部列（如 CSV 行尾多余的分隔符）不算数据列，不会误报删除 / 新增了一列', () => {
  const grid = s => excel.parseCsv(s, ',').sheets[0].grid;
  for (const mode of ['table', 'rows']) {
    Object.assign(excel.opts, { mode, matchHeaders: true, sortRows: false, sortCols: false });
    const R = excel.compare(grid('a,b,\n1,2,\n'), grid('a,b\n1,2\n'), 0, 0);
    assert.equal(R.identical, true, mode);
    assert.equal(R.columns.length, 2, mode);
    const real = excel.compare(grid('a,b,\n1,2,x\n'), grid('a,b\n1,2\n'), 0, 0);
    assert.equal(real.stats.colsDel, 1, mode);
    assert.equal(real.identical, false, mode);
  }
  Object.assign(excel.opts, { mode: 'table' });
});

test('excel.inlineDiff: 忽略大小写时，仅大小写不同的字符不标为改动，且仍可无损还原', () => {
  Object.assign(excel.opts, { ignoreCase: true });
  const segs = excel.inlineDiff('Hello World', 'hello Earth');
  assert.equal(segs.map(s => s.a).join(''), 'Hello World'); assert.equal(segs.map(s => s.b).join(''), 'hello Earth');
  assert.equal([segs[0].t, segs[0].a, segs[0].b].join('|'), '0|Hello |hello ');
  Object.assign(excel.opts, { ignoreCase: false });
  assert.notEqual(excel.inlineDiff('Hello World', 'hello Earth')[0].t, 0);
});
