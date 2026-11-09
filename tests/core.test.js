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
  return run(code, 'compare,parseCsv,opts,serialToDate,serialToElapsed,isZipSig,pairBlock');
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
  assert.equal(JSON.stringify(img.invertAlign({ ...al, fh: true }, A, B)), JSON.stringify({ dx: 0, dy: 0, rot: 0, scale: 1, fh: false, fv: false }));
});
