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
  const code = between(s, 'function colIndex', 'function parseXlsx') + between(s, 'const MAX_IMPORTED_CELLS=', '/* ===================== diff core') + between(s, 'const EQ=0', '/* ===================== UI: loading');
  return run(code, 'compare,parseCsv,scanCsv,opts,serialToDate,serialToElapsed,isoDateCell,isZipSig,pairBlock,normalize,inlineDiff,shiftFormula,xstr');
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
  return run(code, 'dilate,detect,dedupeRegions,phaseCorr,invertAlign', 'const defAlign=()=>({dx:0,dy:0,rot:0,scale:1,fh:false,fv:false});');
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

test('image.dedupeRegions: 与暴力 IoU 去重等价，不相交区域不再触发平方级比较', () => {
  const reference = all => {
    const keep = [];
    for (const r of all) {
      let dup = false;
      for (const k of keep) {
        const ix = Math.max(0, Math.min(r.x + r.w, k.x + k.w) - Math.max(r.x, k.x));
        const iy = Math.max(0, Math.min(r.y + r.h, k.y + k.h) - Math.max(r.y, k.y));
        const inter = ix * iy;
        if (inter && inter / (r.w * r.h + k.w * k.h - inter) >= .4) { dup = true; break; }
      }
      if (!dup) keep.push(r);
    }
    return keep;
  };
  const shape = rows => Array.from(rows, r => [r.x, r.y, r.w, r.h, r.px]);
  const r = lcg(19);
  for (let it = 0; it < 100; it++) {
    const all = Array.from({ length: 1 + Math.floor(r() * 45) }, () => ({
      x: Math.floor(r() * 120), y: Math.floor(r() * 100), w: 1 + Math.floor(r() * 25), h: 1 + Math.floor(r() * 20), px: 1,
    }));
    assert.deepEqual(shape(img.dedupeRegions(all)), shape(reference(all)));
  }
  const many = Array.from({ length: 30000 }, (_, i) => ({ x: (i % 250) * 65, y: Math.floor(i / 250) * 65, w: 1, h: 1, px: 1 }));
  const started = Date.now(), kept = img.dedupeRegions(many);
  assert.equal(kept.length, many.length);
  assert.ok(Date.now() - started < 1500, `去重耗时 ${Date.now() - started}ms`);
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

test('decodeBuf: 无 BOM 的 UTF-16 含 é、中文等非 ASCII 字符时同样能识别，不会读成夹 NUL 的乱码', () => {
  for (const f of ['text-compare.html', 'excel-compare.html']) {
    const code = between(script(f), 'function decodeBuf', f === 'text-compare.html' ? '\ntry{mergeOpts' : '\nasync function loadFile');
    const ctx = { TextDecoder }; vm.createContext(ctx); vm.runInContext(code + ';this.decodeBuf=decodeBuf', ctx);
    const dec = b => ctx.decodeBuf(Uint8Array.from(b).buffer);
    for (const s of ['name,城市\ncafé,北京\nnaïve,上海', 'hello\nworld']) {
      assert.equal(dec(Buffer.from(s, 'utf16le')), s, f);
      assert.equal(dec(Buffer.from(s, 'utf16le').swap16()), s, f);
    }
    assert.equal(dec([0x61]), 'a', f);
  }
});

test('excel.shiftFormula: 平移后越过最后一行（1048576）的单元格引用保持原样，与整行 / 整列引用的处理一致', () => {
  assert.equal(excel.shiftFormula('A1048576+B2', 5, 0), 'A1048576+B7');
  assert.equal(excel.shiftFormula('A$1048576+A1048571', 5, 1), 'B$1048576+B1048576');
  assert.equal(excel.shiftFormula('SUM(1048576:1048576)', 1, 0), 'SUM(1048576:1048576)');
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

test('excel.normalize: 只规范化真实存在的日期，非法月末和非闰年不能产生伪相等', () => {
  const saved = { ...excel.opts };
  try {
    Object.assign(excel.opts, { normDates: true, ignoreWs: false, ignoreCase: false, mode: 'table', matchHeaders: true, sortRows: false, sortCols: false });
    for (const v of ['2023/02/29', '1900/02/29', '2024/04/31', '2024/06/31', '2024/13/01']) assert.equal(excel.normalize(v), v);
    for (const [v, expected] of [['2024/02/29', '2024-02-29'], ['2000/02/29', '2000-02-29'], ['2024/04/30', '2024-04-30']]) assert.equal(excel.normalize(v), expected);
    const grid = v => ({ rows: [[{ v: 'Date', f: null, d: false }], [{ v, f: null, d: false }]], cols: 1 });
    const bad = excel.compare(grid('2023/02/29'), grid('2023-02-29'), 0, 0);
    assert.equal(bad.identical, false);
    assert.equal(bad.stats.rowsDel, 1);
    assert.equal(bad.stats.rowsIns, 1);
  } finally { Object.assign(excel.opts, saved); }
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

test('readZip: ZIP 注释含伪结束标记时仍定位真实目录，损坏的条目不能静默读入', async () => {
  const zip = Buffer.from(await zipLib.zipStore([{ name: 'xl/workbook.xml', data: new TextEncoder().encode('workbook data') }]).arrayBuffer());
  const read = b => zipLib.readZip(Uint8Array.from(b).buffer);
  const commented = Buffer.from(zip);
  commented.writeUInt16LE(24, commented.length - 2);
  const comment = Buffer.alloc(24);
  comment.writeUInt32LE(0x06054b50, 2); // 恰好位于文件最后 22 字节的伪 EOCD
  const z = read(Buffer.concat([commented, comment]));
  assert.equal(z.text('xl/workbook.xml'), 'workbook data');
  assert.equal(z.get('constructor'), null); // ZIP 文件名不能继承 Object.prototype
  const namedUndefined = Buffer.from(await zipLib.zipStore([{ name: 'undefined', data: new TextEncoder().encode('real entry') }]).arrayBuffer());
  assert.equal(read(namedUndefined).get('missing'), null); // 缺失的大小写别名不能误查到名为 undefined 的文件
  assert.equal(read(namedUndefined).text('undefined'), 'real entry');

  const corrupt = Buffer.from(zip);
  const dataStart = 30 + corrupt.readUInt16LE(26) + corrupt.readUInt16LE(28);
  corrupt[dataStart] ^= 1; // 长度不变，只修改内容
  assert.throws(() => read(corrupt).text('xl/workbook.xml'), /badZip/);

  const badDir = Buffer.from(zip);
  badDir.writeUInt16LE(2, badDir.length - 22 + 10); // 声称有两项，实际上只有一项
  assert.throws(() => read(badDir), /badZip/);

  // 真正的 .xlsx 通常使用 DEFLATE；同样须验证解压后的内容而不只是压缩长度
  const packed = require('node:zlib').deflateRawSync(Buffer.from('workbook data'));
  const oldCd = zip.readUInt32LE(zip.length - 22 + 16);
  const local = Buffer.from(zip.subarray(0, dataStart)), central = Buffer.from(zip.subarray(oldCd, zip.length - 22)), end = Buffer.from(zip.subarray(zip.length - 22));
  local.writeUInt16LE(8, 8); local.writeUInt32LE(packed.length, 18);
  central.writeUInt16LE(8, 10); central.writeUInt32LE(packed.length, 20);
  end.writeUInt32LE(dataStart + packed.length, 16);
  const deflated = Buffer.concat([local, packed, central, end]);
  assert.equal(read(deflated).text('xl/workbook.xml'), 'workbook data');
  const wrongCrc = Buffer.from(deflated);
  wrongCrc[dataStart + packed.length + 16] ^= 1;
  assert.throws(() => read(wrongCrc).text('xl/workbook.xml'), /badZip/);
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

test('excel.inflate: 在输出前校验声明长度，解压结果必须准确匹配', () => {
  const zlib = require('node:zlib'), raw = Buffer.from('hello '.repeat(5000));
  const z = new Uint8Array(zlib.deflateRawSync(raw));
  assert.deepEqual(Buffer.from(inflate(z, raw.length)), raw);
  for (const size of [0xffffffff, 1, 0, raw.length + 1]) assert.throws(() => inflate(z, size), /badZip/);
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

test('Excel 按位置匹配：空表头改动在比较标签中写出两侧的值', () => {
  const grid = name => ({ rows: [[{ v: name }], [{ v: 'item' }]], cols: 1 });
  Object.assign(excel.opts, { mode: 'table', matchHeaders: false, sortRows: false, sortCols: false });
  try {
    for (const [a, b, label] of [['Price', '', '"Price" → ""'], ['', 'Price', '"" → "Price"']]) {
      const r = excel.compare(grid(a), grid(b), 0, 0);
      assert.equal(r.stats.cells, 1);
      assert.equal(r.columns[0].name, label);
      assert.ok(r.columns[0].hdrChg);
      assert.equal(r.rows[0].type, 'eq');
    }
  } finally { excel.opts.matchHeaders = true; }
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
  assert.equal(excel.isoDateCell('2024-03-05T00:00:00Z'), '2024-03-05T00:00:00Z');
  assert.equal(excel.isoDateCell('2024-03-05T08:05:09.000'), '2024-03-05 08:05:09');
  assert.equal(excel.isoDateCell('2024-03-05T08:05'), '2024-03-05 08:05:00');
  assert.equal(excel.isoDateCell('10:30:00'), excel.serialToDate(0.4375));
  assert.equal(excel.isoDateCell('not a date'), 'not a date');
  assert.equal(excel.isoDateCell(null), '');
});

test('excel.isoDateCell: 非法日历日期、越界时间与时区保持原文，不能截断后伪相等', () => {
  for (const value of [
    '2023-02-29', '2024-02-30T00:00:00', '2024-04-31T08:00:00Z', '2024-13-01',
    '2024-01-01T24:00:00', '2024-01-01T12:60:00', '25:00:00', '2024-01-01T00:00:00+25:00',
  ]) assert.equal(excel.isoDateCell(value), value);
  assert.equal(excel.isoDateCell('2024-02-29T00:00:00'), '2024-02-29');
  assert.equal(excel.isoDateCell('2024-02-29T08:05:09+05:30'), '2024-02-29T08:05:09+05:30');
});

test('Excel 的 ISO 日期和日期文本保留有效的小数秒，不与整秒伪相等', () => {
  Object.assign(excel.opts, { normDates: true });
  const date = excel.isoDateCell('2024-02-29T00:00:00.125');
  assert.equal(date, '2024-02-29 00:00:00.125');
  assert.equal(excel.normalize('2024/2/29 0:00:00.125'), date);
  assert.equal(excel.isoDateCell('08:30:00.100'), '08:30:00.1');
  assert.equal(excel.normalize('8:30:00.100'), '08:30:00.1');
  assert.notEqual(date, excel.isoDateCell('2024-02-29T00:00:00.000'));
  assert.equal(excel.isoDateCell('0000-02-29T00:00:00'), '0000-02-29T00:00:00');
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

test('excel.xstr: 解码 xlsx 字符串里的 _xHHHH_ 转义，_x005F_ 表示字面的下划线', () => {
  assert.equal(excel.xstr('line1_x000D_\nline2'), 'line1\r\nline2');
  assert.equal(excel.xstr('_x005F_x000D_'), '_x000D_');
  assert.equal(excel.xstr('a_x0041_b_x4e2d_'), 'aAb中');
  assert.equal(excel.xstr('plain_text_x12'), 'plain_text_x12');
});

test('computeDiff: 空文本按 0 行处理，和空白输入对比时不会多报一处空行的删除 / 新增', () => {
  const o = { rules: [] };
  const add = text.computeDiff('', 'a\nb', o); assert.equal(add.nDel, 0); assert.equal(add.nIns, 2);
  const del = text.computeDiff('a', '', o); assert.equal(del.nDel, 1); assert.equal(del.nIns, 0);
  const none = text.computeDiff('', '', o); assert.equal(none.nDel + none.nIns, 0);
  const nl = text.computeDiff('a\n', 'a', o); assert.equal(nl.nDel, 1); assert.equal(nl.nIns, 0); // 末尾换行的差异仍会报告
});

test('xe: 导出 xlsx 时转义字面的 _xHHHH_，经读取器的 xstr 解码后与原文一致', () => {
  for (const [f, end] of [['text-compare.html', 'function colName'], ['excel-compare.html', 'function makeXlsx']]) {
    const xe = run(between(script(f), 'const xe=', end), 'xe', '').xe;
    assert.equal(xe('a_x000D_b'), 'a_x005F_x000D_b', f);
    assert.equal(xe('a_b & <c>'), 'a_b &amp; &lt;c&gt;', f);
    const al = ['_', 'x', '0', '0', 'D', '5', 'F', 'a'];
    for (const s of ['_x000D_x000D_', '_x005F_', '__x000A__', ...Array.from({ length: 3000 }, () => Array.from({ length: rnd(16) }, () => al[rnd(al.length)]).join(''))]) {
      assert.equal(excel.xstr(xe(s)), s, f + ' ' + s);
    }
  }
});

test('excel.parseCsv: 首行是标题 / 空行或只有一列时，仍按后续行识别分隔符', () => {
  const g = s => JSON.parse(JSON.stringify(excel.parseCsv(s).sheets[0].grid.rows.map(r => r.map(c => c.v))));
  assert.deepEqual(g('\na;b\n1;2'), [[''], ['a', 'b'], ['1', '2']]);
  assert.deepEqual(g('Report title\na;b\n1;2'), [['Report title'], ['a', 'b'], ['1', '2']]);
  // 字段中间的引号不开启引用，不能吞掉后面的行而让分隔符识别退回逗号
  assert.deepEqual(g('Report title\n5" x;1\n2;3'), [['Report title'], ['5" x', '1'], ['2', '3']]);
  // 标题行里的逗号比数据行的分号还多，仍以出现行数最多的分号为准
  assert.deepEqual(g('Report, A, B, C\na;b\n1;2'), [['Report, A, B, C'], ['a', 'b'], ['1', '2']]);
  assert.deepEqual(g('a\tb\n1\t2'), [['a', 'b'], ['1', '2']]);
});

test('excel.compare: 表头行号超出数据范围时按无表头比较，不会把两侧所有行都丢掉', () => {
  const G = rows => ({ rows: rows.map(r => r.map(v => ({ v: String(v), f: null, d: false }))), cols: Math.max(0, ...rows.map(r => r.length)) });
  Object.assign(excel.opts, { mode: 'table', matchHeaders: true, sortRows: false, sortCols: false });
  const R = excel.compare(G([['A'], [1], [2]]), G([['B'], [3], [4]]), 9, 9);
  assert.equal(R.identical, false);
  assert.equal(R.stats.rowsDel, 3);
  assert.equal(R.stats.rowsIns, 3);
  assert.equal(R.rows.length, 6);
  // 合法的表头行仍然只比较表头以下的行（第一行是第 2 行，不是表头）
  const ok = excel.compare(G([['A'], [1], [2]]), G([['A'], [1], [3]]), 0, 0);
  assert.equal(ok.rows[0].na, 2);
  assert.equal(ok.stats.rowsDel, 1);
  assert.equal(ok.stats.rowsIns, 1);
});

test('excel.parseCsv: 出现行数相同时，优先选每行个数一致的分隔符，而不是正文里逗号更多的那个', () => {
  const g = s => JSON.parse(JSON.stringify(excel.parseCsv(s).sheets[0].grid.rows.map(r => r.map(c => c.v))));
  const body = [['id', 'name, title', 'note'], ['1', 'apples, pears, plums, figs', 'x'], ['2', 'red, green, blue, pink', 'y'], ['3', 'one, two, three, four', 'z']];
  for (const d of ['\t', ';', '|']) assert.deepEqual(g(body.map(r => r.join(d)).join('\n')), body);
  // 普通逗号分隔的文件不受影响，正文里偶尔出现的分号也不会抢走分隔符
  assert.deepEqual(g('a,b;c\n1,2;3\n4,5,6'), [['a', 'b;c'], ['1', '2;3'], ['4', '5', '6']]);
});

test('excel.shiftFormula: 整列（A:A）与整行（1:1）引用同样按位移平移，$ 锁定、工作表名与大写 3D 引用保持不变', () => {
  const s = excel.shiftFormula;
  assert.equal(s('SUM(A:A)', 0, 1), 'SUM(B:B)');
  assert.equal(s('SUM(A:A)', 3, 0), 'SUM(A:A)');
  assert.equal(s('INDEX(B:B,ROW(1:1))', 1, 0), 'INDEX(B:B,ROW(2:2))');
  assert.equal(s('SUM($A:B,C:$D)', 0, 2), 'SUM($A:D,E:$D)');
  assert.equal(s('SUM($1:2,3:$4)', 2, 0), 'SUM($1:4,5:$4)');
  assert.equal(s('Sheet2!A:A+A1', 1, 1), 'Sheet2!B:B+B2');
  assert.equal(s('\'A:B\'!A1&"1:1"', 1, 0), '\'A:B\'!A2&"1:1"');
  assert.equal(s('JAN:MAR!A1', 1, 1), 'JAN:MAR!B2');
  // 平移后越界的引用保持原样
  assert.equal(s('SUM(XFD:XFD)', 0, 1), 'SUM(XFD:XFD)');
});

test('image.autoAlign: 沿用并带回调用方当前的翻转状态；带翻转时不走 OpenCV（仿射估计不含镜像）', async () => {
  const s = script('image-compare.html');
  const pre = 'const t=k=>k,notice=()=>{},setTimeout=f=>f(),window={},calls=[],ocv=[];'
    + 'function grayPatch(src,N,sc,rot,fh,fv){calls.push([src.id,fh,fv]);return{g:[],f:1}}'
    + 'function phaseCorr(){return{dx:3,dy:-2,peak:1,psr:50}}'
    + 'function alignWithOpenCV(){ocv.push(1);return{dx:9,dy:9,rot:0,scale:1,fh:false,fv:false,conf:.9}}';
  const X = run(between(s, 'async function autoAlign', 'function alignWithOpenCV'), 'autoAlign,calls,ocv,window', pre);
  const A = { id: 'A', width: 64, height: 48 }, B = { id: 'B', width: 64, height: 48 };
  for (const allow of [{ rot: false, scale: false }, { rot: true, scale: true }]) for (const [fh, fv] of [[false, false], [true, false], [false, true], [true, true]]) {
    X.calls.length = 0;
    const r = await X.autoAlign(A, B, { ...allow, fh, fv });
    assert.equal(r.fh, fh); assert.equal(r.fv, fv); assert.equal(r.dx, 3); assert.equal(r.dy, -2);
    assert.ok(X.calls.length > 0);
    for (const [id, h, v] of X.calls) { assert.equal(h, id === 'B' ? fh : false); assert.equal(v, id === 'B' ? fv : false); }
  }
  assert.equal((await X.autoAlign(A, B, { rot: true, scale: true })).fh, false); // 不传翻转时与原先一致
  X.window.cv = { ORB: 1 };
  assert.equal((await X.autoAlign(A, B, { rot: true, scale: true })).dx, 9); assert.equal(X.ocv.length, 1);
  assert.equal((await X.autoAlign(A, B, { rot: false, scale: true })).dx, 3); assert.equal(X.ocv.length, 1);
  const flipped = await X.autoAlign(A, B, { rot: true, scale: true, fh: true });
  assert.equal(X.ocv.length, 1); assert.equal(flipped.fh, true); assert.equal(flipped.dx, 3);
});

test('excel.compare: 按位置匹配列时，“对列排序”先把两侧的列各自按表头排序再逐个对应', () => {
  const G = rows => ({ rows: rows.map(r => r.map(v => ({ v: String(v), f: null, d: false }))), cols: Math.max(0, ...rows.map(r => r.length)) });
  const A = [['A', 'B', 'C'], [1, 2, 3], [4, 5, 6]], B = [['C', 'A', 'B'], [3, 1, 2], [6, 4, 5]];
  Object.assign(excel.opts, { mode: 'table', matchHeaders: false, sortRows: false, sortCols: true });
  const R = excel.compare(G(A), G(B), 0, 0);
  assert.equal(R.identical, true);
  assert.deepEqual([...R.columns].map(c => [c.name, c.ia, c.ib]), [['A', 0, 1], ['B', 1, 2], ['C', 2, 0]]);
  // 列数不同时，排序后多出来的那一列仍按位置记为新增，其余列照常对应
  const R2 = excel.compare(G([['B', 'A'], [2, 1]]), G([['A', 'B', 'C'], [1, 2, 3]]), 0, 0);
  assert.deepEqual([...R2.columns].map(c => c.st), ['eq', 'eq', 'ins']);
  assert.equal(R2.stats.cells, 0); assert.equal(R2.stats.rowsChg, 0);
  // 关闭后仍按原位置逐列对应，行为不变
  Object.assign(excel.opts, { sortCols: false });
  const R3 = excel.compare(G(A), G(B), 0, 0);
  assert.equal(R3.identical, false);
  assert.deepEqual([...R3.columns].map(c => [c.ia, c.ib, c.hdrChg]), [[0, 0, true], [1, 1, true], [2, 2, true]]);
  Object.assign(excel.opts, { matchHeaders: true });
});

test('文本页 CSS 的自动识别与着色：结果正确，且不会在超长的连续字母数字串上回溯成 O(n²)', () => {
  const s = script('text-compare.html');
  const syn = run(between(s, 'const SYN={', '/* ===================== state'), 'detectSyntax,synClasses');
  assert.equal(syn.detectSyntax('.btn-primary:hover {\n  color: #fff;\n  margin: 0 -4px;\n}'), 'css');
  assert.equal(syn.detectSyntax('a{b:c}'), 'css');
  assert.equal(syn.detectSyntax('{ color: red }'), '');
  assert.equal(syn.detectSyntax('plain words, a {brace} and a colon: nothing else'), '');
  const css = '.a-b{color:red;width:10px}\nh1 > h2 { --x-1: #fff }', cls = syn.synClasses(css, 'css');
  const of = (str, from = 0) => { const i = css.indexOf(str, from); return [...new Set(cls.slice(i, i + str.length))].join('|'); };
  assert.equal(of('.a-b'), 'tg'); assert.equal(of('color'), 'at'); assert.equal(of('red'), ''); assert.equal(of('width'), 'at'); assert.equal(of('10px'), 'n');
  assert.equal(of('h1'), ''); // 标识符里的数字不是数值
  assert.equal(of('h2'), 'tg'); assert.equal(of('--x-1'), 'at'); assert.equal(of('#fff'), 'n');
  // 32 万字符的连续标识符（base64url、哈希之类）：原先逐位回溯需要数分钟
  const big = 'token ' + 'a1B2-c3_'.repeat(40000), t0 = Date.now();
  assert.equal(syn.detectSyntax(big), '');
  assert.equal(syn.synClasses(big, 'css').length, big.length);
  assert.ok(Date.now() - t0 < 3000, `took ${Date.now() - t0}ms`);
});

test('文本语法着色把未闭合的跨行注释或三引号字符串持续到文末', () => {
  const code = between(script('text-compare.html'), 'const SYN=', 'const esc=');
  const { synClasses } = run(code, 'synClasses');
  for (const [lang, open, close, kind] of [
    ['js', '/*', '*/', 'c'], ['css', '/*', '*/', 'c'], ['sql', '/*', '*/', 'c'],
    ['xml', '<!--', '-->', 'c'], ['py', '"""', '"""', 's'],
  ]) {
    const prefix = '  ', closed = open + '\ncomment\n' + close, rest = ' ' + open + '\ncontinued\n';
    const cls = synClasses(prefix + closed + rest, lang);
    assert.equal(cls[0], '', lang);
    assert.ok(cls.slice(prefix.length, prefix.length + closed.length).every(c => c === kind), lang);
    assert.equal(cls[prefix.length + closed.length], '', lang);
    assert.ok(cls.slice(prefix.length + closed.length + 1).every(c => c === kind), lang);
  }
});

test('文本页语法自动识别：缩进的 Python / JS 仍能识别，且不会在大量连续空行上回溯成 O(n²)', () => {
  const s = script('text-compare.html');
  const { detectSyntax } = run(between(s, 'const SYN={', '/* ===================== state'), 'detectSyntax');
  assert.equal(detectSyntax('x = 1\n\n\n    \tdef run(a, b):\n        return a'), 'py');
  assert.equal(detectSyntax('notes\n\n  elif done:\n    pass'), 'py');
  assert.equal(detectSyntax('\n\n  from os import path, sep as s\n'), 'py');
  assert.equal(detectSyntax('// header\n\n\n  export default class Foo {}'), 'js');
  assert.equal(detectSyntax('def is a word here\nexported goods'), '');
  const t0 = Date.now();
  assert.equal(detectSyntax('x' + '\n'.repeat(200000) + 'y'), ''); // 原先约需数分钟
  assert.ok(Date.now() - t0 < 3000, `took ${Date.now() - t0}ms`);
});

test('文本页 countText: 行 / 词 / 字符统计与原先的正则口径一致，且不物化词数组', () => {
  const s = script('text-compare.html');
  const X = run(between(s, 'const isWs=', 'function updateMeta'), 'countText', ''), countText = v => ({ ...X.countText(v) });/* vm 里创建的对象原型不同，摊平后再 deepEqual */
  const W = /[\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF\uAC00-\uD7AF]|[^\s\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF\uAC00-\uD7AF]+/g;
  const ref = v => ({ lines: v.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').length, words: (v.match(W) || []).length, chars: v.length });
  const al = ['a', 'Z', '1', '_', ',', ' ', '  ', '\t', '\n', '\r', '\r\n', '\u00a0', '\u3000', '\u2028', '中', '文', 'あ', '한', '😀', '\ufeff', '\u200b'];
  for (let it = 0; it < 3000; it++) {
    const v = Array.from({ length: 1 + rnd(12) }, () => al[rnd(al.length)]).join('');
    assert.deepEqual(countText(v), ref(v), JSON.stringify(v));
  }
  for (const v of ['a', 'a\n', 'a\n\n', '\n', '\r', 'a\r\nb\r\n', 'x\ry', ' 中 文 ', 'foo_bar,baz 😀x'] ) assert.deepEqual(countText(v), ref(v), JSON.stringify(v));
});

test('excel.inlineDiff: 忽略空白时空白差异不标成改动，且仍能无损还原', () => {
  const saved = excel.opts.ignoreWs;
  try {
    excel.opts.ignoreWs = true;
    const join = (segs, k) => segs.map(s => s[k]).join('');
    const bad = s => s.t !== 0 && !(s.t === 1 ? s.a : s.b).trim();
    for (const [l, r] of [['a  b x', 'a b y'], ['  a b x', 'a b y'], ['x  y', 'x y z'], ['a\tb', 'a b'], ['ab', 'a b']]) {
      const segs = excel.inlineDiff(l, r);
      assert.equal(join(segs, 'a'), l); assert.equal(join(segs, 'b'), r);
      assert.ok(!segs.some(bad), JSON.stringify([l, r, segs]));
    }
    assert.ok(excel.inlineDiff('x  y', 'x y z').every(s => s.t === 0 || s.b === 'z'));
    assert.ok(excel.inlineDiff('a  b', 'a b').every(s => s.t === 0));
    excel.opts.ignoreWs = false;
    assert.ok(excel.inlineDiff('a  b', 'a b').some(s => s.t !== 0));
  } finally { excel.opts.ignoreWs = saved; }
});

test('文本页语法自动识别：SQL / CSS / import…from 的判断不再在大量 SELECT / { / import 上回溯成 O(n²)，且与原正则等价', () => {
  const s = script('text-compare.html');
  const { detectSyntax } = run(between(s, 'const SYN={', '/* ===================== state'), 'detectSyntax');
  assert.equal(detectSyntax('select 1\n-- notes\nSELECT a FROM t'), 'sql');
  assert.equal(detectSyntax('a{color:red}'), 'css');
  assert.equal(detectSyntax('b{}\na{margin:0}'), 'css');
  assert.equal(detectSyntax("import x from\n'./x'"), 'js'); // from 与引号之间的空白可以跨行
  for (const [name, src] of [['SELECT', 'SELECT\n'.repeat(60000)], ['{', 'a{'.repeat(60000)], ['import', 'import '.repeat(30000)]]) {
    const t0 = Date.now();
    assert.equal(detectSyntax(src), '', name); // 原先各需数秒到数十秒
    assert.ok(Date.now() - t0 < 1500, `${name}: ${Date.now() - t0}ms`);
  }
  // 与原来的整体正则逐一对照
  const ref = str => {
    str = str.trim();
    if (/^[\[{][\s\S]*[\]}]$/.test(str)) { try { JSON.parse(str); return 'json'; } catch (e) {} }
    if (/^<\?xml|^<!doctype|^<[a-z!]/i.test(str)) return 'xml';
    if (/^[ \t]*(def|class)\s+\w+.*:\s*$|^[ \t]*elif\b|\bself\.|^[ \t]*print\(|^[ \t]*(from\s+[\w.]+\s+)?import\s+[\w.]+(\s*,\s*[\w.]+)*(\s+as\s+\w+)?\s*$/m.test(str)) return 'py';
    if (/\b(?:const|let|var)\s+[\w$]+\s*[=;:,]|\bfunction\s*[\w$]*\s*\(|=>|\bimport\b[^\n]*\bfrom\s+['"]|^[ \t]*export\s+(?:default|const|function|class)\b/m.test(str)) return 'js';
    if (/\bSELECT\b[\s\S]*\bFROM\b/i.test(str)) return 'sql';
    if (/[\w-]\s*\{[^}]*:[^}]*\}/.test(str)) return 'css';
    return '';
  };
  const tok = ['SELECT', 'select', ' FROM ', 'from', '{', '}', ':', ';', 'a', '-', '\n', 'import ', ' from ', "'", '"', 'x', ' ', '\r\n', '=', '(', ')', 'color', 'b1'];
  for (let it = 0; it < 30000; it++) {
    const str = Array.from({ length: 1 + rnd(14) }, () => tok[rnd(tok.length)]).join('');
    assert.equal(detectSyntax(str), ref(str), JSON.stringify(str));
  }
});

test('excel.compare: 新增或删除列含数据时，对应行不能被隐藏未修改的行折叠', () => {
  const G = rows => ({ rows: rows.map(r => r.map(v => ({ v: String(v), f: null, d: false }))), cols: Math.max(0, ...rows.map(r => r.length)) });
  const saved = { ...excel.opts };
  try {
    Object.assign(excel.opts, { mode: 'table', matchHeaders: true, hideRows: true, sortRows: false, sortCols: false });
    const left = G([['ID'], ['1'], ['2'], ['3']]);
    const right = G([['ID', 'Note'], ['1', 'new'], ['2', ''], ['3', 'also new']]);
    for (const [a, b, status] of [[left, right, 'cins'], [right, left, 'cdel']]) {
      const result = excel.compare(a, b, 0, 0);
      assert.deepEqual(Array.from(result.rows, r => r.type), ['eq', 'eq', 'eq']);
      assert.deepEqual(Array.from(result.rows, r => r.hasColData), [true, false, true]);
      assert.equal(result.stats.rowsChg, 0); // 新列有独立统计，不计入共有列的修改数
      assert.equal(result.rows[0].cells[1].st, status);
      assert.equal(result.rows[2].cells[1].st, status);
      assert.ok(result.colChanged.has(1));
      const nodes = { stats: { innerHTML: '' }, modeSeg: { querySelectorAll: () => [] }, valSeg: { querySelectorAll: () => [] }, gridWrap: { dataset: {}, querySelectorAll: () => [] } };
      const ctx = { result, options: excel.opts, nodes };
      vm.createContext(ctx);
      vm.runInContext(`const RES=result,opts=options,focusIdx=-1,expanded=new Set(),t=k=>k,esc=String,$=id=>nodes[id];${between(script('excel-compare.html'), 'function renderCell', 'function goChange')}render()`, ctx);
      assert.match(nodes.gridWrap.innerHTML, /also new/);
      assert.doesNotMatch(nodes.gridWrap.innerHTML, /class="collapsed"/);
      assert.match(nodes.gridWrap.innerHTML, /<th[^>]+id="focus-0"/);
      assert.equal(nodes.gridWrap.dataset.changes, 3); // 列标题与两行数据均可定位。
    }
  } finally { Object.assign(excel.opts, saved); }
});

test('文本行内差异：先在整行应用规则，再按粒度比较且还原原文', () => {
  for (const level of ['char', 'word', 'smart']) {
    const norm = text.makeNorm({ hideWs: true, igQuotes: true, rules: ['“tag” here', 'OMIT'] });
    const a = '  left  “tag”   here old😀', b = 'left new😀 OMIT ';
    const spans = text.inlineDiff(a, b, level, norm);
    assert.equal(spans.map(s => s.a).join(''), a);
    assert.equal(spans.map(s => s.b).join(''), b);
    assert.equal(spans.filter(s => s.t === 1).map(s => s.a).join(''), 'old');
    assert.equal(spans.filter(s => s.t === 2).map(s => s.b).join(''), 'new');
    assert.equal(norm.project(a).key, norm(a));
    assert.equal(norm.project(b).key, norm(b));
  }
  const w = text.makeNorm({ hideWs: true });
  assert.equal(text.inlineDiff('a b', 'ab', 'char', w).filter(s => s.t === 1).map(s => s.a).join(''), ' ');
});

test('decodeBuf: BOM 优先，奇数字节或无偏向的 NUL 不作为 UTF-16 猜测', () => {
  for (const f of ['text-compare.html', 'excel-compare.html']) {
    const code = between(script(f), 'function decodeBuf', f.startsWith('text') ? '\ntry{mergeOpts' : '\nasync function loadFile');
    const ctx = { TextDecoder };
    vm.createContext(ctx);
    vm.runInContext(code + ';this.decodeBuf=decodeBuf', ctx);
    const dec = b => ctx.decodeBuf(Uint8Array.from(b).buffer);
    assert.equal(dec([0xef, 0xbb, 0xbf, 0x00, 0x00, 0x61, 0x00, 0x62, 0x00]), '\0\0a\0b\0');
    assert.equal(dec([0x61, 0x00, 0x62, 0x00, 0x63]), 'a\0b\0c');
    assert.equal(dec([0x00, 0x61, 0x62, 0x00]), '\0ab\0');
    assert.equal(dec([0xff, 0xfe, 0x61, 0x00]), 'a');
  }
});

test('excel.compare: 表头以前的列不参与匹配；空表头不能与真实名称冲突', () => {
  const grid = s => excel.parseCsv(s, ',').sheets[0].grid;
  const saved = { ...excel.opts };
  try {
    Object.assign(excel.opts, { mode: 'table', matchHeaders: true, sortCols: false, ignoreWs: false });
    const a = grid('Report,,,Reference\nName,Price\nTea,1'), b = grid('Report\nName,Price\nTea,1');
    const r = excel.compare(a, b, 1, 1);
    assert.equal(r.identical, true);
    assert.equal(r.columns.length, 2);
    const named = excel.compare(grid(',Name\n1,Tea'), grid('__col0,Name\n1,Tea'), 0, 0);
    assert.equal(named.stats.colsDel, 1);
    assert.equal(named.stats.colsIns, 1);
  } finally { Object.assign(excel.opts, saved); }
});

test('excel.compare: 仅含空白的尾部行列只在开启忽略空白时裁掉', () => {
  const grid = s => excel.parseCsv(s, ',').sheets[0].grid;
  const saved = { ...excel.opts };
  try {
    Object.assign(excel.opts, { mode: 'rows', sortRows: false, ignoreWs: false });
    const a = grid('Tea, '), b = grid('Tea');
    assert.equal(excel.compare(a, b, 0, 0).stats.colsDel, 1);
    assert.equal(excel.compare(grid('Tea\n '), b, 0, 0).stats.rowsDel, 1);
    excel.opts.ignoreWs = true;
    assert.equal(excel.compare(a, b, 0, 0).identical, true);
  } finally { Object.assign(excel.opts, saved); }
});

test('excel.compare: 在展开稀疏行列之前限制比较规模', () => {
  const saved = { ...excel.opts };
  try {
    Object.assign(excel.opts, { mode: 'rows', sortRows: false });
    const row = []; row[16383] = { v: 'value', f: null };
    const cells = Array.from({ length: 40 }, () => row);
    assert.throws(() => excel.compare({ rows: cells }, { rows: [] }, 0, 0), /compareTooLarge/);
  } finally { Object.assign(excel.opts, saved); }
});

test('表格会话：保留实际行宽与公式，拒绝结构无效的单元格', () => {
  const code = between(script('excel-compare.html'), 'function validateSavedRows', 'function setWorkbook');
  const { fromJsonGrid, toSavedGrid } = run(code, 'fromJsonGrid,toSavedGrid', 'const t=k=>k,MAX_COMPARE_CELLS=500000;');
  const cell = []; cell[2] = { v: '3', f: '=1+2' };
  const rows = [[{ v: ' ' }], cell, []];
  const saved = JSON.parse(JSON.stringify(toSavedGrid({ name: 'Data' }, { cols: 16384, rows }, 1)));
  assert.deepEqual(saved.rows.map(r => r.length), [1, 3, 0]);
  const b = fromJsonGrid(saved), grid = b.load(b.sheets[0]);
  assert.equal(grid.cols, 3);
  assert.equal(grid.rows[0][0].v, ' ');
  assert.equal(grid.rows[1][2].f, '=1+2');
  for (const bad of [{ rows: [[[]]] }, { rows: [[{ v: {} }]] }, { rows: [['ok']], header: -1 }]) {
    assert.throws(() => fromJsonGrid(bad), /wrongFile/);
  }
});

test('Excel 日期系统：日期格式的 0、虚拟闰日与超过一天的纯时间', () => {
  const code = between(script('excel-compare.html'), 'function secondFraction', '// t="d"');
  const { serialToDate, customDateKind, serialToElapsed } = run(code, 'serialToDate,customDateKind,serialToElapsed');
  assert.equal(serialToDate(0, true, false), '1904-01-01');
  assert.equal(serialToDate(0.5, true, false), '1904-01-01 12:00:00');
  assert.equal(serialToDate(1.5, false, true), '12:00:00');
  assert.equal(serialToDate(60.25, false, false), '1900-02-29 06:00:00');
  assert.equal(serialToDate(60, false, false, false), '1900-02-28');
  assert.equal(serialToElapsed(1e300), '1e+300');
  for (const [format, kind] of [['yyyy-mm-dd hh:mm', 'date'], ['mm', 'date'], ['hh:mm:ss.000', 'time'], ['h:mm AM/PM', 'time'], ['[h]:mm', 'elapsed'], ['0.00 "days"', '']]) {
    assert.equal(customDateKind(format), kind, format);
  }
});

test('Excel 数值时间：格式精度决定小数秒，跨秒及午夜正确进位', () => {
  const code = between(script('excel-compare.html'), 'function secondFraction', '// t="d"');
  const { serialToDate, serialToElapsed, dateNumberStyle } = run(code, 'serialToDate,serialToElapsed,dateNumberStyle');
  assert.deepEqual({ ...dateNumberStyle('yyyy-mm-dd hh:mm:ss.000') }, { kind: 'date', precision: 3 });
  assert.deepEqual({ ...dateNumberStyle('[s].00') }, { kind: 'elapsed', precision: 2 });
  assert.deepEqual({ ...dateNumberStyle('hh:mm:ss ".000"') }, { kind: 'time', precision: 0 });
  assert.equal(serialToDate(45356 + 0.125 / 86400, false, false, true, 3), '2024-03-05 00:00:00.125');
  assert.equal(serialToDate(0.5 + 0.2 / 86400, false, true, true, 1), '12:00:00.2');
  assert.equal(serialToDate(1 - 0.0004 / 86400, false, true, true, 3), '00:00:00');
  assert.equal(serialToElapsed(-(90001.125 / 86400), 3), '-25:00:01.125');
});

test('Excel 自定义数值格式：只按实际生效的格式段识别日期', () => {
  const code = between(script('excel-compare.html'), 'function secondFraction', '// t="d"');
  const { customNumberStyle } = run(code, 'customNumberStyle');
  const signed = customNumberStyle('yyyy-mm-dd ";";0.00;0;@');
  assert.equal(signed(45356).kind, 'date');
  assert.equal(signed(-1).kind, '');
  assert.equal(signed(0).kind, '');
  const conditional = customNumberStyle('[Blue][>=1]yyyy-mm-dd;[>=0]hh:mm:ss.000;0');
  assert.equal(conditional(1).kind, 'date');
  assert.deepEqual({ ...conditional(0.5) }, { kind: 'time', precision: 3 });
  assert.equal(conditional(-2).kind, '');
});

test('XLSX 部件路径：以当前部件目录为基准解析相对路径', () => {
  const { resolvePart } = run(between(script('excel-compare.html'), 'function resolvePart', 'function parseXlsx'), 'resolvePart');
  assert.equal(resolvePart('xl/', 'xl/sheet.xml'), 'xl/xl/sheet.xml');
  assert.equal(resolvePart('custom/nested/', '../sheets/data.xml'), 'custom/sheets/data.xml');
  assert.equal(resolvePart('xl/', '/worksheets/data.xml'), 'worksheets/data.xml');
  assert.equal(resolvePart('xl/', '../../outside.xml'), '');
  assert.equal(resolvePart('', 'https://example.invalid/sheet.xml'), '');
});

test('readZip: 本地条目名、长度必须与中央目录相符', async () => {
  const buf = Buffer.from(await zipLib.zipStore([{ name: 'data', data: new TextEncoder().encode('tea') }]).arrayBuffer());
  for (const offset of [22, 30]) {
    const changed = Buffer.from(buf); changed[offset] ^= 1;
    assert.throws(() => zipLib.readZip(Uint8Array.from(changed).buffer).text('data'), /badZip/);
  }
});

test('XLSX XML：仅序言里的文档类型声明被拒绝，不误判注释和正文', () => {
  const { hasXmlDoctype } = run(between(script('excel-compare.html'), 'function hasXmlDoctype', 'const xml='), 'hasXmlDoctype');
  assert.equal(hasXmlDoctype('<?xml version="1.0"?> \n<!-- intro -->\n<!DOCTYPE workbook><workbook/>'), true);
  assert.equal(hasXmlDoctype('<!-- <!DOCTYPE workbook> --><workbook><![CDATA[<!DOCTYPE sample>]]></workbook>'), false);
  assert.equal(hasXmlDoctype('<?xml version="1.0"?><workbook/>'), false);
});

test('XLSX ZIP：XML 部件可以使用带 BOM 的 UTF-16LE 或 UTF-16BE', async () => {
  const xml = '<?xml version="1.0" encoding="UTF-16"?><worksheet>中文 😀</worksheet>';
  const le = Buffer.from('\ufeff' + xml, 'utf16le'), be = Buffer.from(le).swap16();
  const blob = zipLib.zipStore([{ name: 'le.xml', data: le }, { name: 'be.xml', data: be }, { name: 'broken.xml', data: Uint8Array.of(255) }]);
  const zip = zipLib.readZip(await blob.arrayBuffer());
  assert.equal(zip.text('le.xml'), xml);
  assert.equal(zip.text('be.xml'), xml);
  assert.throws(() => zip.text('broken.xml'), /badZip/);
});

test('CSV: 按候选分隔符解析引号，拒绝未闭合字段', () => {
  const read = (s, d) => JSON.parse(JSON.stringify(excel.parseCsv(s, d).sheets[0].grid.rows.map(r => r.map(c => c.v))));
  assert.deepEqual(read('ID;Label\n1;"first, line\nsecond"\n2;"last, row"'), [['ID', 'Label'], ['1', 'first, line\nsecond'], ['2', 'last, row']]);
  assert.deepEqual(read('a,tag;ok\nb,"one|two"\nc,x', ','), [['a', 'tag;ok'], ['b', 'one|two'], ['c', 'x']]);
  assert.throws(() => read('ID,Label\n1,"not closed'), /badCsvQuotes/);
});

test('CSV：多行字段内部的分隔符不会抢走真实分隔符', () => {
  const values = (s, d) => Array.from(excel.parseCsv(s, d).sheets[0].grid.rows, r => Array.from(r, c => c.v));
  for (const [d, other] of [[',', ';'], [';', ','], ['|', ';']]) {
    const notes = Array.from({ length: 8 }, (_, i) => `label${i}${other}detail`).join('\n');
    assert.deepEqual(values(`ID${d}Notes\n1${d}"${notes}"\n2${d}end`), [['ID', 'Notes'], ['1', notes], ['2', 'end']]);
  }
  assert.throws(() => values('a,"b"junk\nc,d', ','), /badCsvQuotes/);
});

test('CSV 导出：前导空白的公式样式文本被保留为文本，数值保持原样', () => {
  const { csvField } = run(between(script('excel-compare.html'), 'function csvField', 'function exportDiffFile'), 'csvField');
  assert.equal(csvField('  =1+2'), "'  =1+2");
  assert.equal(csvField('＝1+2'), "'＝1+2");
  assert.equal(csvField('-12.5'), '-12.5');
  assert.equal(csvField('-.5e+2'), '-.5e+2');
  assert.equal(csvField('Tea; Coffee | Water'), '"Tea; Coffee | Water"');
});

test('共享公式：移动相对引用但保留结构化列名和工作表范围', () => {
  const shift = excel.shiftFormula;
  assert.equal(shift('SUM(Table1[[#This Row],[A1]])+a1', 2, 1), 'SUM(Table1[[#This Row],[A1]])+b3');
  assert.equal(shift('S1:S3!C4+[Book.xlsx]S1!d2', 1, 1), 'S1:S3!D5+[Book.xlsx]S1!e3');
  assert.equal(shift('A1+A1.local+Table1[A1]', 1, 0), 'A2+A1.local+Table1[A1]');
});

test('XLSX 字符串导出保留回车、控制字符及有效代理对', () => {
  for (const [f, end] of [['text-compare.html', 'function colName'], ['excel-compare.html', 'function makeXlsx']]) {
    const { xe } = run(between(script(f), 'const xe=', end), 'xe');
    const s = 'first\r\nsecond\0\x1b 😀\ud800\ufffe_x000D_';
    assert.equal(excel.xstr(xe(s)), s);
    assert.doesNotMatch(xe(s), /[\x00\r\x1b\ufffe]/);
  }
});

test('XLSX 写入拒绝超出行、列、文本规格的结果，而不截断数据', () => {
  for (const f of ['text-compare.html', 'excel-compare.html']) {
    const code = between(script(f), 'const xe=', 'function exportXlsx');
    const { makeXlsx } = run(code, 'makeXlsx', 'const TextEncoder=class{encode(){}};' + (f.startsWith('excel') ? 'function colName(){return "A"}' : ''));
    assert.throws(() => makeXlsx([{ name: 'Test', rows: new Array(1048577) }]), /exportXlsxLimit/);
    assert.throws(() => makeXlsx([{ name: 'Test', rows: [new Array(16385)] }]), /exportXlsxLimit/);
    assert.throws(() => makeXlsx([{ name: 'Test', rows: [[{ v: 'x'.repeat(32768) }]] }]), /exportXlsxLimit/);
  }
});

test('图片像素差异同时比较颜色和透明度，忽略全透明像素隐藏的颜色', () => {
  const { diffPixels } = run(between(script('image-compare.html'), 'function diffPixels', 'function computeDiffMap'), 'diffPixels');
  const a = [255, 255, 255, 0, 255, 0, 0, 0, 20, 30, 40, 255];
  const b = [255, 255, 255, 255, 0, 0, 255, 0, 20, 35, 40, 255];
  const map = new Uint8Array(3), rgba = new Uint8Array(12);
  assert.equal(diffPixels(a, b, map, null, rgba), 2);
  assert.deepEqual([...map], [255, 0, 5]);
  assert.deepEqual([...rgba], [255, 255, 255, 255, 0, 0, 0, 255, 0, 5, 0, 255]);
});

test('图片区域检测：环形区域不重复统计内部区域，并始终执行最大面积过滤', () => {
  const W = 21, H = 21, mask = new Uint8Array(W * H);
  for (let j = 2; j <= 18; j++) for (let i = 2; i <= 18; i++) if (i === 2 || i === 18 || j === 2 || j === 18) mask[j * W + i] = 1;
  mask[10 * W + 10] = 1;
  const rect = { x0: 0, y0: 0, x1: W, y1: H };
  for (const radius of [0, 2]) {
    const all = [], limited = [];
    img.detect(mask, W, H, { minBlobArea: 1, maxBlobAreaFraction: 1 }, rect, radius, all);
    assert.deepEqual(all.map(r => r.px), [64, 1]);
    img.detect(mask, W, H, { minBlobArea: 1, maxBlobAreaFraction: 0.25 }, rect, radius, limited);
    assert.deepEqual(limited.map(r => r.px), [1]);
  }
  const partial = [];
  assert.equal(img.detect(mask, W, H, { minBlobArea: 1, maxBlobAreaFraction: 1 }, rect, 0, partial, 1), false);
  assert.equal(partial.length, 1);
});

test('图片工作画布：极窄的输入仍至少保留一个像素', () => {
  const { workingFrame, fitsCanvas } = run(between(script('image-compare.html'), 'const MAXDIM=', 'let baseA'), 'workingFrame,fitsCanvas');
  assert.deepEqual({ ...workingFrame(12000, 1) }, { w: 2600, h: 1, scale: 2600 / 12000 });
  assert.deepEqual({ ...workingFrame(1, 12000) }, { w: 1, h: 2600, scale: 2600 / 12000 });
  assert.equal(fitsCanvas(3000, 6000, 32 * 1024 * 1024), true);
  assert.equal(fitsCanvas(16384, 4096, 32 * 1024 * 1024), false);
});

test('图片对比画幅包含原图和变换后的四角，不剪掉负坐标或超出的内容', () => {
  const { comparisonFrame } = run(between(script('image-compare.html'), 'const MAXDIM=', 'let baseA'), 'comparisonFrame');
  const a = { w: 4, h: 3 }, b = { w: 8, h: 4 }, al = { dx: 0, dy: 0, rot: 0, scale: 1 };
  assert.deepEqual({ ...comparisonFrame(a, b, al) }, { x: 0, y: 0, w: 8, h: 4 });
  assert.deepEqual({ ...comparisonFrame(a, b, { ...al, dx: -3, dy: 2 }) }, { x: -3, y: 0, w: 8, h: 6 });
  assert.deepEqual({ ...comparisonFrame(a, b, { ...al, rot: 90, fh: true }) }, { x: 0, y: -2, w: 6, h: 8 });
  assert.throws(() => comparisonFrame(a, b, { ...al, dx: 1e10 }), /alignOutOfRange/);
});

test('image.autoAlign: 同时允许旋转和缩放时，终步仍评估仅缩放候选', async () => {
  const code = between(script('image-compare.html'), 'async function autoAlign', 'function alignWithOpenCV');
  const pre = 'const window={},setTimeout=f=>f();'
    + 'function grayPatch(src,N,sc,rot){return {g:{N,sc,rot},f:1}}'
    + 'function phaseCorr(a,b){return {dx:0,dy:0,psr:b.N===512&&b.sc===1.01&&b.rot===0?100:50}}';
  const { autoAlign } = run(code, 'autoAlign', pre), dim = { width: 64, height: 64 };
  const r = await autoAlign(dim, dim, { rot: true, scale: true });
  assert.equal(r.rot, 0);
  assert.equal(r.scale, 1.01);
});

test('image.autoAlign: 中等分辨率通过而最终峰值不足时返回失败', async () => {
  const code = between(script('image-compare.html'), 'async function autoAlign', 'function alignWithOpenCV');
  const pre = 'const window={};function grayPatch(src,N){return {g:N,f:1}}'
    + 'function phaseCorr(a,b){return {dx:0,dy:0,psr:b===512?5:30}}';
  const { autoAlign } = run(code, 'autoAlign', pre), dims = { width: 96, height: 64 };
  assert.equal(await autoAlign(dims, dims, { rot: false, scale: false }), null);
});

test('文本打印视图可以使用不同布局，但不改当前布局、不复制导航标识', () => {
  const s = script('text-compare.html'), d = text.computeDiff('old line', 'new line', { rules: [] });
  const ctx = { d, opts: { type: 'split', level: 'smart', syntax: '' }, inlineDiff: text.inlineDiff };
  vm.createContext(ctx);
  vm.runInContext(`const EQ=0,DEL=1,INS=2,t=k=>k,expanded=new Set();${between(s, 'const SYN=', '/* ===================== state')}${between(s, 'const CTX=', '/* 展开折叠块')}this.output=diffView(d,'unified',false);this.screen=diffView(d,opts.type);`, ctx);
  assert.equal(ctx.opts.type, 'split');
  assert.match(ctx.output.html, /class="sign">−/);
  assert.doesNotMatch(ctx.output.html, /id="focus-/);
  assert.match(ctx.screen.html, /class="split chg" id="focus-0"/);
});

test('文本 PNG 的尺寸和画布像素面积都有限', () => {
  const { pngLineBudget } = run(between(script('text-compare.html'), 'function pngLineBudget', 'async function exportPng'), 'pngLineBudget');
  const width = 1600, header = 72, line = 20, max = pngLineBudget(width, header, line);
  assert.ok((header + max * line) * width <= 16 * 1024 * 1024);
  assert.ok((header + (max + 1) * line) * width > 16 * 1024 * 1024);
  assert.ok(header + pngLineBudget(100, header, line) * line <= 16384);
});

test('文本 PNG 排版保留长行和换行后的高亮，不切断代理对', () => {
  const { wrapPngLine } = run(between(script('text-compare.html'), 'function wrapPngLine', 'async function exportPng'), 'wrapPngLine');
  const rows = wrapPngLine([{ text: 'abcd', mark: false }, { text: '😀efgh', mark: true }], 4, () => 1, 10);
  assert.deepEqual(JSON.parse(JSON.stringify(rows)).map(r => r.map(p => [p.text, p.mark])), [[['abcd', false]], [['😀efg', true]], [['h', true]]]);
  assert.throws(() => wrapPngLine([{ text: 'x'.repeat(20), mark: true }], 4, () => 1, 3), /image too tall/);
});

test('共享公式：长标识符保持原样，非 ASCII 字母不会误认成列引用', () => {
  const name = 'X'.repeat(60000) + '_note';
  assert.equal(excel.shiftFormula(`${name}+K1+a1`, 1, 1), `${name}+K1+b2`);
});

test('共享公式：超出原工作表网格的名称不能被移进网格', () => {
  assert.equal(excel.shiftFormula('XFE1+A1048577+XFE:XFF+1048577:1048578+B2', -1, -1),
    'XFE1+A1048577+XFE:XFF+1048577:1048578+A1');
});

test('XLSX 单元格地址保留实际坐标并拒绝越界或非法地址', () => {
  const { cellAddress } = run(between(script('excel-compare.html'), 'function colIndex', '// OOXML 字符串'), 'cellAddress');
  assert.deepEqual({ ...cellAddress('xfd1048576') }, { row: 1048575, col: 16383 });
  assert.deepEqual({ ...cellAddress('B7') }, { row: 6, col: 1 });
  for (const ref of ['XFE1', 'A1048577', 'A0', 'A7B', '1A', '']) assert.equal(cellAddress(ref), null);
});


const worksheet = (() => {
  const s = script('excel-compare.html');
  const code = between(s, 'function colIndex', 'function parseXlsx') + between(s, 'const MAX_IMPORTED_CELLS=', '// 检测和实际解析');
  const { parseSheetRows } = run(code, 'parseSheetRows');
  const node = (localName, attrs = {}, children = [], textContent = '') => ({
    localName, children, textContent, getAttribute: name => attrs[name] ?? null, hasAttribute: name => Object.hasOwn(attrs, name),
  });
  const row = (n, ...cells) => node('row', n == null ? {} : { r: n }, cells);
  const cell = (ref, value, type) => node('c', { ...(ref == null ? {} : { r: ref }), t: type }, value === undefined ? [] : [node('v', {}, [], value)]);
  const readRows = (rows, sst = [], limit) => parseSheetRows({ children: rows }, sst, new Map(), false, true, n => n.textContent, limit);
  return { node, row, cell, readRows };
})();

test('XLSX 工作表拒绝重复的空行和空单元格坐标，保留合法的隐式坐标', () => {
  const { row, cell, readRows } = worksheet;
  assert.throws(() => readRows([row('2'), row('2', cell('A2', '7'))]), /badZip/);
  assert.throws(() => readRows([row('1', cell('B1'), cell('B1', '7'))]), /badZip/);
  assert.throws(() => readRows([row('1e2')]), /badZip/);
  const { rows, cols } = readRows([row(null, cell('B3', '7'), cell(null, '8'))]);
  assert.equal(rows.length, 3);
  assert.equal(cols, 3);
  assert.equal(rows[2][1].v, '7');
  assert.equal(rows[2][2].v, '8');
});


test('XLSX 工作表：缺省公式缓存为空；非法共享字符串索引和布尔值不能伪装成正常值', () => {
  const { node, row, cell, readRows } = worksheet;
  const formula = (r, t) => node('c', { r, t }, [node('f', {}, [], 'B2')]);
  const { rows } = readRows([row('1', formula('A1', 's'), formula('B1', 'b'), cell('C1', ' 0 ', 's'), cell('D1', 'false', 'b'))], ['label']);
  assert.deepEqual(Array.from(rows[0], c => c.v), ['', '', 'label', 'FALSE']);
  assert.equal(rows[0][0].f, '=B2');
  for (const [value, type] of [['', 's'], ['1', 's'], ['1e0', 's'], ['other', 'b']]) {
    assert.throws(() => readRows([row('1', cell('A1', value, type))], ['label']), /badZip/);
  }
});


test('共享公式：组号按数值匹配，同一组内必须有唯一的主公式', () => {
  const { node, row, readRows } = worksheet;
  const f = (r, si, text) => node('c', { r }, [node('f', { t: 'shared', si }, [], text)]);
  const { rows } = readRows([row('1', f('A1', '01', '')), row('2', f('A2', '1', 'B2'))]);
  assert.equal(rows[0][0].f, '=B1');
  assert.throws(() => readRows([row('1', f('A1', '01', 'B1'), f('B1', '1', 'C1'))]), /badZip/);
  assert.throws(() => readRows([row('1', f('A1', undefined, 'B1'))]), /badZip/);
});


test('表格导入预算：收集 CSV 字段和 XLSX 单元格时即拒绝超限', () => {
  assert.equal(excel.scanCsv('a,b\nc,d', ',', true, Infinity, 4).rows.length, 2);
  assert.throws(() => excel.scanCsv('a,b\nc,d', ',', true, Infinity, 3), /importTooLarge/);
  assert.throws(() => excel.parseCsv(','.repeat(16384), ','), /importTooLarge/);
  const { row, cell, readRows } = worksheet;
  const rows = [row('3', cell('A3'), cell('B3', 'data'))];
  assert.equal(readRows(rows, [], 2).rows[2][1].v, 'data');
  assert.throws(() => readRows(rows, [], 1), /importTooLarge/);
});
