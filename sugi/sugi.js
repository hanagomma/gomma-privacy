/* 수기 세금계산서 집계 — 브라우저(및 Node 테스트)에서 동작하는 판독·집계·엑셀 생성 로직
 * 필요 전역: XLSX(SheetJS), ExcelJS, JSZip
 */
(function (root) {
  'use strict';

  var norm = function (s) { return String(s == null ? '' : s).replace(/[\s 　]/g, ''); };
  var LABEL_RE = /^(등록번호|사업자등록번호|종사업장|상호|성명|대표자|사업장|주소|업태|종목|공급자|공급받는자|공급받는|공급가액|세액|비고|작성|품목|합계|이메일|전자우편|담당)/;
  var IMG_RE = /\.(jpe?g|png|gif|bmp|heic|heif|tiff?|webp|pdf)$/i;
  var XL_RE = /\.(xlsx|xlsm|xls)$/i;

  function isLabel(v) { var n = norm(v); return n !== '' && LABEL_RE.test(n); }
  function cell(g, r, c) { var row = g[r]; return row ? row[c] : null; }
  function empty(v) { return v == null || norm(v) === ''; }
  function toNum(v) {
    if (v == null) return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = String(v).replace(/[,\s원₩]/g, '');
    if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
    return parseFloat(s);
  }

  // ---------- 사업자등록번호 ----------
  function bizCheck(d) {
    if (!/^\d{10}$/.test(d)) return false;
    var w = [1, 3, 7, 1, 3, 7, 1, 3, 5], s = 0;
    for (var i = 0; i < 9; i++) s += (+d[i]) * w[i];
    s += Math.floor((+d[8]) * 5 / 10);
    return (10 - s % 10) % 10 === +d[9];
  }
  function fmtBiz(d) { return /^\d{10}$/.test(d) ? d.slice(0, 3) + '-' + d.slice(3, 5) + '-' + d.slice(5) : d; }

  // ---------- 그리드 ----------
  function sheetGrid(ws) {
    if (!ws || !ws['!ref']) return [];
    var g = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    return g;
  }

  // 라벨 오른쪽 값 모으기 (다음 라벨 또는 경계 전까지)
  function rightValues(g, r, c, limit) {
    var row = g[r] || [], out = [];
    var end = Math.min(row.length, limit == null ? row.length : limit);
    for (var j = c + 1; j < end; j++) {
      var v = row[j];
      if (empty(v)) continue;
      if (isLabel(v)) break;
      out.push(v);
    }
    return out;
  }

  function findCells(g, r0, r1, pred) {
    var out = [];
    for (var r = r0; r <= r1; r++) {
      var row = g[r]; if (!row) continue;
      for (var c = 0; c < row.length; c++) { var v = row[c]; if (!empty(v) && pred(norm(v), v)) out.push({ r: r, c: c, v: v }); }
    }
    return out;
  }

  // ---------- 블록(세금계산서 1장) 분리 ----------
  function splitBlocks(g) {
    var titles = [];
    for (var r = 0; r < g.length; r++) {
      var row = g[r] || [];
      for (var c = 0; c < row.length; c++) {
        var n = norm(row[c]);
        if (n.length <= 30 && /세금계산서/.test(n) && !/합계|집계|요약|목록|명세/.test(n)) {
          titles.push({ r: r, text: n }); break;
        }
      }
    }
    // 연속 행 제목(제목 + 부제)은 하나로
    var merged = [];
    titles.forEach(function (t) {
      var last = merged[merged.length - 1];
      if (last && t.r - last.r <= 2) { last.text += t.text; } else merged.push({ r: t.r, text: t.text });
    });
    if (!merged.length) return [];
    return merged.map(function (t, i) {
      return { r0: Math.max(0, t.r - 1), r1: (i + 1 < merged.length ? merged[i + 1].r - 2 : g.length - 1), title: t.text };
    });
  }

  // ---------- 블록 판독 ----------
  function parseBlock(g, b) {
    var r0 = b.r0, r1 = b.r1, res = { notes: [] };
    var regs = findCells(g, r0, r1, function (n) { return /^(사업자)?등록번호/.test(n); });
    if (!regs.length) return null;
    regs.sort(function (a, b2) { return a.c - b2.c || a.r - b2.r; });
    var sup = regs[0];
    // 공급받는자 경계: 같은 행(±1)에서 오른쪽에 있는 등록번호, 또는 '공급받는자' 라벨
    var boundary = null;
    regs.forEach(function (x) { if (x !== sup && Math.abs(x.r - sup.r) <= 1 && x.c > sup.c && (boundary == null || x.c < boundary)) boundary = x.c; });
    var recvLab = findCells(g, r0, r1, function (n) { return /^공급받/.test(n) && !/보관/.test(n); });
    recvLab.forEach(function (x) { if (x.c > sup.c && (boundary == null || x.c < boundary)) boundary = x.c; });
    res.boundary = boundary;

    // 등록번호
    var raw = rightValues(g, sup.r, sup.c, boundary).map(function (v) { return typeof v === 'number' ? String(Math.round(v)) : String(v); }).join('');
    var lab = norm(sup.v).replace(/^(사업자)?등록번호/, '');
    if (/\d/.test(lab)) raw = lab + raw; // "등록번호 123-45-67890" 한 칸
    var digits = raw.replace(/\D/g, '');
    res.regRaw = raw.trim();
    res.reg = digits.length > 10 ? digits.slice(0, 10) : digits;

    function labelValue(re) {
      var cands = findCells(g, r0, r1, function (n) { return re.test(n); })
        .filter(function (x) { return x.c >= sup.c - 1 && (boundary == null || x.c < boundary); })
        .sort(function (a, b2) { return a.r - b2.r || a.c - b2.c; });
      for (var i = 0; i < cands.length; i++) {
        var inline = norm(cands[i].v).replace(re, '').replace(/^\(?(법인명|대표자)\)?/, '').replace(/^[:：]/, '');
        if (inline) return String(cands[i].v).replace(/^.*?[:：)]\s*/, '').trim() || inline;
        var vals = rightValues(g, cands[i].r, cands[i].c, boundary);
        if (vals.length) return String(vals[0]).trim();
      }
      return '';
    }
    res.name = labelValue(/^상호/);
    res.rep = labelValue(/^(성명|대표자)/);

    // 작성일
    res.date = findDate(g, r0, r1, res);

    // 금액
    var items = itemTable(g, r0, r1);
    var dig = digitAmounts(g, r0, r1);
    res.items = items ? items.names : [];
    var supI = items ? items.sup : 0, vatI = items ? items.vat : 0;
    if (items && (supI || vatI)) {
      res.sup = supI; res.vat = vatI; res.src = '품목';
      if (dig && (dig.sup || dig.vat) && (dig.sup !== supI || dig.vat !== vatI)) res.notes.push('자릿수칸 금액 불일치');
    } else if (dig && (dig.sup || dig.vat)) {
      res.sup = dig.sup; res.vat = dig.vat; res.src = '자릿수칸';
    } else {
      var tot = labeledAmount(g, r0, r1);
      res.sup = tot.sup; res.vat = tot.vat; res.src = tot.sup || tot.vat ? '금액칸' : '';
    }
    if (!res.date && items && items.firstMD) res.dateMD = items.firstMD;
    res.copy = /공급받는자보관/.test(b.title) ? '공급받는자' : (/공급자보관/.test(b.title) ? '공급자' : '');
    return res;
  }

  function mkDate(y, m, d) {
    y = +y; m = +m; d = +d || 1;
    if (!y || !m || m < 1 || m > 12) return null;
    if (y < 100) y += 2000;
    if (y < 1990 || y > 2100) return null;
    if (d < 1 || d > 31) d = 1;
    return { y: y, m: m, d: d };
  }
  function fromAny(v) {
    if (v instanceof Date && !isNaN(v)) return mkDate(v.getFullYear(), v.getMonth() + 1, v.getDate());
    if (typeof v === 'number' && v > 30000 && v < 80000) { var dt = XLSX.SSF.parse_date_code(v); return mkDate(dt.y, dt.m, dt.d); }
    var s = String(v == null ? '' : v);
    var m = s.match(/(\d{2,4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})?\s*일?/) || s.match(/(\d{4})\s*[.\-\/]\s*(\d{1,2})\s*[.\-\/]\s*(\d{1,2})/);
    if (m) return mkDate(m[1], m[2], m[3]);
    return null;
  }
  function findDate(g, r0, r1) {
    // 1) '년' '월' '일' 머리행 아래 값
    for (var r = r0; r <= r1; r++) {
      var row = g[r] || [], cy = -1, cm = -1, cd = -1;
      for (var c = 0; c < row.length; c++) {
        var n = norm(row[c]);
        if (n === '년' && cy < 0) cy = c; else if (n === '월' && cy >= 0 && cm < 0) cm = c; else if (n === '일' && cm >= 0 && cd < 0) cd = c;
      }
      if (cy >= 0 && cm >= 0) {
        for (var k = 1; k <= 2; k++) {
          var y = toNum(cell(g, r + k, cy)), mo = toNum(cell(g, r + k, cm)), d = cd >= 0 ? toNum(cell(g, r + k, cd)) : 1;
          var dt = mkDate(y, mo, d); if (dt) return dt;
          var whole = fromAny(cell(g, r + k, cy)); if (whole) return whole;
        }
      }
    }
    // 2) '작성일자' 라벨 옆/아래
    var labs = findCells(g, r0, r1, function (n) { return /^작성(일자|일|년월일)?/.test(n); });
    for (var i = 0; i < labs.length; i++) {
      var x = labs[i], inl = fromAny(x.v); if (inl) return inl;
      var row2 = g[x.r] || [];
      for (var j = x.c + 1; j < Math.min(row2.length, x.c + 8); j++) { var d1 = fromAny(row2[j]); if (d1) return d1; }
      for (var k2 = 1; k2 <= 2; k2++) { var d2 = fromAny(cell(g, x.r + k2, x.c)); if (d2) return d2; }
    }
    // 3) 아무 칸이나 'YYYY년 M월 D일'
    var any = findCells(g, r0, r1, function (n, v) { return !!fromAny(v) && (/년.*월/.test(n) || v instanceof Date); });
    return any.length ? fromAny(any[0].v) : null;
  }

  // 품목 표: '품목' 머리행 → 공급가액/세액 열 합계
  function itemTable(g, r0, r1) {
    for (var r = r0; r <= r1; r++) {
      var row = g[r] || [], cP = -1, cS = -1, cV = -1, cM = -1, cD = -1;
      for (var c = 0; c < row.length; c++) {
        var n = norm(row[c]);
        if (/^품목/.test(n) && cP < 0) cP = c;
        else if (/^공급가액/.test(n) && cS < 0) cS = c;
        else if (/^세액/.test(n) && cV < 0) cV = c;
        else if (n === '월' && cM < 0) cM = c;
        else if (n === '일' && cD < 0 && cM >= 0) cD = c;
      }
      if (cP < 0 || cS < 0) continue;
      var sup = 0, vat = 0, names = [], firstMD = null, rows = 0;
      for (var k = r + 1; k <= r1; k++) {
        var rr = g[k] || [];
        var first = rr.map(norm).filter(Boolean)[0] || '';
        if (/^(합계금액|합계|계$|총계)/.test(first) || rr.some(function (v) { return /^합계금액/.test(norm(v)); })) break;
        var s = toNum(rr[cS]), v = cV >= 0 ? toNum(rr[cV]) : null;
        if (s == null && v == null) { if (rows > 0 && rr.every(empty)) continue; else continue; }
        sup += s || 0; vat += v || 0; rows++;
        if (!empty(rr[cP])) names.push(String(rr[cP]).trim());
        if (!firstMD && cM >= 0) { var mm = toNum(rr[cM]); if (mm >= 1 && mm <= 12) firstMD = { m: mm, d: toNum(rr[cD]) || 1 }; }
      }
      return { sup: Math.round(sup), vat: Math.round(vat), names: names, firstMD: firstMD };
    }
    return null;
  }

  // 상단 자릿수칸 (공급가액 | 세액 머리 아래 숫자칸 이어붙이기)
  function digitAmounts(g, r0, r1) {
    var heads = findCells(g, r0, r1, function (n) { return /^공급가액/.test(n); }).sort(function (a, b) { return a.r - b.r; });
    for (var i = 0; i < heads.length; i++) {
      var h = heads[i], row = g[h.r] || [];
      if (row.some(function (v) { return /^품목/.test(norm(v)); })) continue; // 품목표 머리행 제외
      var cV = -1, cE = row.length;
      for (var c = h.c + 1; c < row.length; c++) {
        var n = norm(row[c]);
        if (cV < 0 && /^세액/.test(n)) cV = c;
        else if (cV >= 0 && n && !/^세액/.test(n)) { cE = c; break; }
      }
      if (cV < 0) continue;
      // 값 행: 머리 아래 1~3행 중 단위행(백/십/억...) 아닌 첫 행
      for (var k = 1; k <= 3; k++) {
        var vr = g[h.r + k] || [];
        var seg = vr.slice(h.c, cE).map(norm).filter(Boolean);
        if (!seg.length) continue;
        if (seg.every(function (x) { return /^(백|십|억|천|만|일|조|공란수)$/.test(x); })) continue;
        var join = function (a, b) {
          var s = ''; for (var j = a; j < b; j++) { var v = vr[j]; if (empty(v)) continue; s += typeof v === 'number' ? String(Math.round(v)) : String(v).replace(/[,\s]/g, ''); }
          return /^\d+$/.test(s) ? parseInt(s, 10) : (toNum(s) || 0);
        };
        return { sup: join(h.c, cV), vat: join(cV, cE) };
      }
    }
    return null;
  }

  function labeledAmount(g, r0, r1) {
    var pick = function (re) {
      var cs = findCells(g, r0, r1, function (n) { return re.test(n); });
      for (var i = 0; i < cs.length; i++) {
        var row = g[cs[i].r] || [];
        for (var j = cs[i].c + 1; j < Math.min(row.length, cs[i].c + 6); j++) { var n = toNum(row[j]); if (n) return n; }
        for (var k = 1; k <= 2; k++) { var n2 = toNum(cell(g, cs[i].r + k, cs[i].c)); if (n2) return n2; }
      }
      return 0;
    };
    return { sup: Math.round(pick(/^공급가액/)), vat: Math.round(pick(/^세액/)) };
  }

  // ---------- 파일 → 세금계산서 목록 ----------
  function parseWorkbook(wb, fileName) {
    var out = [], dupCount = 0, seen = {};
    wb.SheetNames.forEach(function (sn) {
      var ws = wb.Sheets[sn];
      if (ws && ws['!state'] === 'hidden') return;
      var g = sheetGrid(ws); if (!g.length) return;
      var blocks = splitBlocks(g);
      if (!blocks.length && findCells(g, 0, g.length - 1, function (n) { return /^(사업자)?등록번호/.test(n); }).length &&
          findCells(g, 0, g.length - 1, function (n) { return /^공급가액/.test(n); }).length) blocks = [{ r0: 0, r1: g.length - 1, title: '' }];
      blocks.forEach(function (b) {
        var inv = parseBlock(g, b);
        if (!inv) return;
        if (!inv.sup && !inv.vat) return; // 빈 양식
        if (!inv.date && inv.dateMD) inv.dateGuess = inv.dateMD;
        inv.file = fileName; inv.sheet = sn;
        var dkey = inv.date ? inv.date.y + '-' + inv.date.m + '-' + inv.date.d : 'nodate:' + sn;
        var key = [inv.reg, norm(inv.name), dkey, inv.sup, inv.vat].join('|');
        if (seen[key]) { dupCount++; return; }
        seen[key] = true;
        out.push(inv);
      });
    });
    return { invoices: out, dup: dupCount };
  }

  function decodeName(bytes) {
    try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
    catch (e) { try { return new TextDecoder('euc-kr').decode(bytes); } catch (e2) { return String.fromCharCode.apply(null, bytes); } }
  }

  // inputs: [{name, data(ArrayBuffer|Uint8Array)}]
  async function expandInputs(inputs) {
    var files = [], images = [], others = [];
    async function walk(name, data, depth) {
      var base = name.split('/').pop();
      if (/^(~\$|\._)/.test(base) || /__MACOSX/.test(name) || base === '.DS_Store') return;
      if (/\.zip$/i.test(name) && depth < 3) {
        var zip = await JSZip.loadAsync(data, { decodeFileName: decodeName });
        var names = Object.keys(zip.files).sort();
        for (var i = 0; i < names.length; i++) {
          var e = zip.files[names[i]]; if (e.dir) continue;
          await walk(e.name, await e.async('uint8array'), depth + 1);
        }
      } else if (XL_RE.test(name)) files.push({ name: name, data: data });
      else if (IMG_RE.test(name)) images.push(name);
      else others.push(name);
    }
    for (var i = 0; i < inputs.length; i++) await walk(inputs[i].name, inputs[i].data, 0);
    return { files: files, images: images, others: others };
  }

  // ---------- 집계 ----------
  function process(parsedFiles, opts) {
    opts = opts || {};
    var today = opts.today || new Date();
    var all = [], skipped = [], dupTotal = 0, errors = [];
    parsedFiles.forEach(function (pf) {
      if (pf.error) { errors.push(pf.name + ' (' + pf.error + ')'); return; }
      if (!pf.invoices.length) { skipped.push(pf.name); return; }
      dupTotal += pf.dup; all = all.concat(pf.invoices);
    });

    var warns = {}, notes = {};
    function add(bag, kind, sample) { (bag[kind] = bag[kind] || { kind: kind, count: 0, samples: [] }); bag[kind].count++; if (sample && bag[kind].samples.length < 5) bag[kind].samples.push(sample); }
    var where = function (inv) { return inv.file.split('/').pop() + ' [' + inv.sheet + ']'; };

    // 작성일 보정 (연도 없음 → 다른 계산서 연도)
    var years = {}; all.forEach(function (i) { if (i.date) years[i.date.y] = (years[i.date.y] || 0) + 1; });
    var commonYear = +Object.keys(years).sort(function (a, b) { return years[b] - years[a]; })[0] || today.getFullYear();
    all.forEach(function (inv) {
      if (!inv.date && inv.dateGuess) { inv.date = { y: commonYear, m: inv.dateGuess.m, d: inv.dateGuess.d }; add(warns, '작성연도 추정', where(inv) + ' → ' + commonYear + '년'); }
      if (!inv.date) add(warns, '작성일 판독 불가', where(inv) + ' ' + (inv.name || ''));
    });

    // 사업자 묶기
    var people = {}, order = [], badSeen = {};
    all.forEach(function (inv) {
      var valid = bizCheck(inv.reg);
      var bk = inv.reg + '|' + norm(inv.name);
      if (!valid && !badSeen[bk]) { badSeen[bk] = 1; add(warns, '등록번호 이상(형식·검증번호 오류)', (inv.name || '') + ' ' + (inv.rep || '') + ' ' + (inv.regRaw || '(없음)') + ' — ' + inv.file.split('/').pop()); }
      var key = valid || inv.reg.length === 10 ? 'R' + inv.reg : 'N' + norm(inv.name) + '|' + norm(inv.rep);
      var p = people[key];
      if (!p) { p = people[key] = { key: key, reg: inv.reg, names: {}, reps: {}, months: {}, cnt: 0, sup: 0, vat: 0 }; order.push(key); }
      if (inv.name) p.names[inv.name] = (p.names[inv.name] || 0) + 1;
      if (inv.rep) p.reps[inv.rep] = (p.reps[inv.rep] || 0) + 1;
      var mk = inv.date ? inv.date.y + '-' + ('0' + inv.date.m).slice(-2) : '미상';
      var mo = p.months[mk] = p.months[mk] || { sup: 0, vat: 0, cnt: 0 };
      mo.sup += inv.sup; mo.vat += inv.vat; mo.cnt++;
      p.cnt++; p.sup += inv.sup; p.vat += inv.vat;
      inv.pkey = key;

      if (inv.date && new Date(inv.date.y, inv.date.m - 1, inv.date.d) > today) add(warns, '미래 작성일', where(inv) + ' ' + (inv.name || '') + ' ' + inv.date.y + '.' + inv.date.m + '.' + inv.date.d);
      if (inv.sup && Math.abs(Math.round(inv.sup / 10) - inv.vat) > 10 && inv.vat !== 0) add(notes, '세액이 공급가액의 10%와 다름', where(inv) + ' ' + (inv.name || ''));
      if (inv.vat === 0 && inv.sup) add(notes, '세액 0원(영세율·누락 여부 확인)', where(inv) + ' ' + (inv.name || ''));
      inv.notes.forEach(function (n) { add(notes, n, where(inv)); });
    });
    var top = function (o) { return Object.keys(o).sort(function (a, b) { return o[b] - o[a]; })[0] || ''; };
    var list = order.map(function (k) { var p = people[k]; p.name = top(p.names); p.rep = top(p.reps); return p; });

    // 사업자 갈림: 같은 상호·성명인데 등록번호가 다름 / 같은 번호에 상호 여러 개
    var byName = {};
    list.forEach(function (p) { var nk = norm(p.name) + '|' + norm(p.rep); (byName[nk] = byName[nk] || []).push(p); });
    Object.keys(byName).forEach(function (nk) {
      var arr = byName[nk]; if (arr.length > 1 && nk !== '|') add(warns, '사업자 갈림(같은 이름, 다른 등록번호)', arr[0].name + ' ' + arr[0].rep + ': ' + arr.map(function (p) { return fmtBiz(p.reg) || '번호없음'; }).join(' / '));
    });
    list.forEach(function (p) { if (Object.keys(p.names).length > 1) add(notes, '같은 등록번호에 상호 표기가 여러 개', fmtBiz(p.reg) + ': ' + Object.keys(p.names).join(' / ')); });

    // 다른 파일 간 중복 의심
    var cross = {};
    all.forEach(function (inv) {
      if (!inv.date) return;
      var k = [inv.pkey, inv.date.y, inv.date.m, inv.date.d, inv.sup, inv.vat].join('|');
      if (cross[k] && cross[k] !== inv.file) add(warns, '다른 파일에 같은 계산서(중복 의심)', inv.name + ' ' + inv.date.y + '.' + inv.date.m + '.' + inv.date.d + ' ' + inv.sup.toLocaleString('ko-KR'));
      else cross[k] = inv.file;
    });
    if (dupTotal) add(notes, '같은 파일 안의 사본(보관용 2벌 등) 제외', dupTotal + '장');

    list.sort(function (a, b) { return (a.name || '').localeCompare(b.name || '', 'ko'); });
    var months = {}; list.forEach(function (p) { Object.keys(p.months).forEach(function (m) { months[m] = 1; }); });
    months = Object.keys(months).sort(function (a, b) { return a === '미상' ? 1 : b === '미상' ? -1 : a < b ? -1 : 1; });

    var sum = { people: list.length, cnt: 0, sup: 0, vat: 0 };
    list.forEach(function (p) { sum.cnt += p.cnt; sum.sup += p.sup; sum.vat += p.vat; });
    var vals = function (o) { return Object.keys(o).map(function (k) { return o[k]; }); };
    return { people: list, months: months, invoices: all, summary: sum, warns: vals(warns), notes: vals(notes), skipped: skipped, errors: errors };
  }

  // ---------- 엑셀 생성 ----------
  function monthLabel(m, multiYear) { if (m === '미상') return '작성월 미상'; var p = m.split('-'); return (multiYear ? p[0].slice(2) + '년 ' : '') + (+p[1]) + '월'; }
  function periodLabel(months) {
    var ms = months.filter(function (m) { return m !== '미상'; });
    if (!ms.length) return '';
    var y0 = ms[0].slice(0, 4), y1 = ms[ms.length - 1].slice(0, 4);
    if (y0 === y1) return y0 + '년 ' + (+ms[0].slice(5)) + '~' + (+ms[ms.length - 1].slice(5)) + '월';
    return y0 + '.' + ms[0].slice(5) + '~' + y1 + '.' + ms[ms.length - 1].slice(5);
  }

  async function buildXlsx(client, R) {
    var wb = new ExcelJS.Workbook();
    wb.creator = '금마세무회계 수기 세금계산서 집계';
    var multiYear = {}; R.months.forEach(function (m) { if (m !== '미상') multiYear[m.slice(0, 4)] = 1; });
    multiYear = Object.keys(multiYear).length > 1;
    var BLUE = 'FF1F4E79', HEAD = 'FFDDEBF7', SUB = 'FFF2F2F2', TOT = 'FFFFF2CC';
    var thin = { style: 'thin', color: { argb: 'FFBFBFBF' } };
    var border = { top: thin, left: thin, bottom: thin, right: thin };

    function summarySheet(title, months) {
      var ws = wb.addWorksheet(title.slice(0, 31), { views: [{ state: 'frozen', xSplit: 4, ySplit: 4 }] });
      var nCols = 4 + months.length * 3 + 4;
      ws.mergeCells(1, 1, 1, nCols);
      ws.getCell(1, 1).value = client + ' 수기 세금계산서 집계 — ' + title;
      ws.getCell(1, 1).font = { bold: true, size: 14, color: { argb: BLUE } };
      ws.getCell(2, 1).value = '단위: 원 · 작성 ' + new Date().toLocaleDateString('ko-KR');
      ws.getCell(2, 1).font = { size: 9, color: { argb: 'FF7F7F7F' } };
      var h1 = ws.getRow(3), h2 = ws.getRow(4);
      ['순번', '거래처명(상호)', '대표자', '등록번호'].forEach(function (t, i) { ws.mergeCells(3, i + 1, 4, i + 1); h1.getCell(i + 1).value = t; });
      months.forEach(function (m, i) {
        var c = 5 + i * 3; ws.mergeCells(3, c, 3, c + 2); h1.getCell(c).value = monthLabel(m, multiYear);
        h2.getCell(c).value = '공급가액'; h2.getCell(c + 1).value = '부가세'; h2.getCell(c + 2).value = '합계';
      });
      var tc = 5 + months.length * 3;
      ws.mergeCells(3, tc, 3, tc + 3); h1.getCell(tc).value = '총계';
      ['매수', '공급가액', '부가세', '총합계'].forEach(function (t, i) { h2.getCell(tc + i).value = t; });
      [h1, h2].forEach(function (row) { for (var c = 1; c <= nCols; c++) { var x = row.getCell(c); x.font = { bold: true }; x.alignment = { horizontal: 'center', vertical: 'middle' }; x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: c >= tc ? TOT : HEAD } }; x.border = border; } });

      var rowsP = R.people.filter(function (p) { return months.some(function (m) { return p.months[m]; }); });
      var tot = { cnt: 0, sup: 0, vat: 0, m: {} };
      rowsP.forEach(function (p, idx) {
        var row = ws.getRow(5 + idx), cnt = 0, s = 0, v = 0;
        row.getCell(1).value = idx + 1; row.getCell(2).value = p.name; row.getCell(3).value = p.rep; row.getCell(4).value = fmtBiz(p.reg);
        months.forEach(function (m, i) {
          var x = p.months[m], c = 5 + i * 3;
          if (x) { row.getCell(c).value = x.sup; row.getCell(c + 1).value = x.vat; row.getCell(c + 2).value = x.sup + x.vat; cnt += x.cnt; s += x.sup; v += x.vat;
            var t = tot.m[m] = tot.m[m] || { sup: 0, vat: 0 }; t.sup += x.sup; t.vat += x.vat; }
        });
        row.getCell(tc).value = cnt; row.getCell(tc + 1).value = s; row.getCell(tc + 2).value = v; row.getCell(tc + 3).value = s + v;
        tot.cnt += cnt; tot.sup += s; tot.vat += v;
        for (var c = 1; c <= nCols; c++) { var x2 = row.getCell(c); x2.border = border; if (c >= 5) x2.numFmt = '#,##0'; if (c === 1 || c === 3 || c === 4 || c === tc) x2.alignment = { horizontal: 'center' }; }
      });
      var tr = ws.getRow(5 + rowsP.length);
      ws.mergeCells(tr.number, 1, tr.number, 4); tr.getCell(1).value = '합계 (' + rowsP.length + '명)';
      months.forEach(function (m, i) { var t = tot.m[m] || { sup: 0, vat: 0 }, c = 5 + i * 3; tr.getCell(c).value = t.sup; tr.getCell(c + 1).value = t.vat; tr.getCell(c + 2).value = t.sup + t.vat; });
      tr.getCell(tc).value = tot.cnt; tr.getCell(tc + 1).value = tot.sup; tr.getCell(tc + 2).value = tot.vat; tr.getCell(tc + 3).value = tot.sup + tot.vat;
      for (var c = 1; c <= nCols; c++) { var x3 = tr.getCell(c); x3.font = { bold: true }; x3.border = border; x3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TOT } }; if (c >= 5) x3.numFmt = '#,##0'; }
      tr.getCell(1).alignment = { horizontal: 'center' }; tr.getCell(tc).alignment = { horizontal: 'center' };
      ws.getColumn(1).width = 6; ws.getColumn(2).width = 22; ws.getColumn(3).width = 10; ws.getColumn(4).width = 14;
      for (var c2 = 5; c2 <= nCols; c2++) ws.getColumn(c2).width = c2 === tc ? 7 : 13;
      ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
    }

    var period = periodLabel(R.months);
    summarySheet('통합 (' + (period || '전체') + ')', R.months);
    var qs = {};
    R.months.forEach(function (m) { if (m === '미상') return; var q = m.slice(0, 4) + '-' + Math.ceil(+m.slice(5) / 3); (qs[q] = qs[q] || []).push(m); });
    var qKeys = Object.keys(qs).sort();
    if (qKeys.length > 1 || (qKeys.length === 1 && qs[qKeys[0]].length < R.months.length)) {
      qKeys.forEach(function (q) { var p = q.split('-'); summarySheet((multiYear ? p[0] + ' ' : '') + p[1] + '분기', qs[q]); });
    }

    // 전체명세
    var ws = wb.addWorksheet('전체명세 (검토용)', { views: [{ state: 'frozen', ySplit: 1 }] });
    var hd = ['순번', '거래처명(상호)', '대표자', '등록번호', '작성일', '품목', '공급가액', '부가세', '합계', '판독', '파일', '시트', '비고'];
    ws.addRow(hd).eachCell(function (x) { x.font = { bold: true }; x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEAD } }; x.border = border; x.alignment = { horizontal: 'center' }; });
    var inv = R.invoices.slice().sort(function (a, b) {
      return (a.name || '').localeCompare(b.name || '', 'ko') || ((a.date ? a.date.y * 1e4 + a.date.m * 100 + a.date.d : 0) - (b.date ? b.date.y * 1e4 + b.date.m * 100 + b.date.d : 0));
    });
    inv.forEach(function (i, k) {
      var row = ws.addRow([k + 1, i.name, i.rep, fmtBiz(i.reg) || i.regRaw, i.date ? new Date(Date.UTC(i.date.y, i.date.m - 1, i.date.d)) : '', i.items.slice(0, 3).join(', ') + (i.items.length > 3 ? ' 외 ' + (i.items.length - 3) : ''),
        i.sup, i.vat, i.sup + i.vat, i.src, i.file.split('/').pop(), i.sheet, i.notes.join(', ')]);
      row.eachCell({ includeEmpty: true }, function (x, c) { x.border = border; if (c >= 7 && c <= 9) x.numFmt = '#,##0'; if (c === 5) x.numFmt = 'yyyy-mm-dd'; });
    });
    [6, 22, 10, 14, 12, 26, 13, 12, 13, 8, 28, 10, 22].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    ws.autoFilter = { from: 'A1', to: 'M1' };

    // 확인 필요
    if (R.warns.length || R.notes.length || R.skipped.length || R.errors.length) {
      var wc = wb.addWorksheet('확인 필요');
      wc.addRow(['구분', '내용', '건수', '예시']).eachCell(function (x) { x.font = { bold: true }; x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEAD } }; });
      R.warns.forEach(function (g) { wc.addRow(['⚠ 확인', g.kind, g.count, g.samples.join('\n')]); });
      R.notes.forEach(function (g) { wc.addRow(['참고', g.kind, g.count, g.samples.join('\n')]); });
      if (R.skipped.length) wc.addRow(['제외', '세금계산서 내용이 없는 파일', R.skipped.length, R.skipped.join('\n')]);
      if (R.errors.length) wc.addRow(['오류', '열 수 없는 파일', R.errors.length, R.errors.join('\n')]);
      wc.getColumn(1).width = 8; wc.getColumn(2).width = 36; wc.getColumn(3).width = 7; wc.getColumn(4).width = 80;
      wc.eachRow(function (r) { r.getCell(4).alignment = { wrapText: true, vertical: 'top' }; r.getCell(2).alignment = { vertical: 'top' }; });
    }

    var buf = await wb.xlsx.writeBuffer();
    var safe = String(client).replace(/[\\/:*?"<>|]/g, '_');
    return { buffer: buf, filename: safe + '_수기세금계산서_집계' + (period ? '_' + period.replace(/\s/g, '') : '') + '.xlsx' };
  }

  // ---------- 전체 실행 ----------
  async function run(client, inputs, opts) {
    var ex = await expandInputs(inputs);
    var parsed = ex.files.map(function (f) {
      try {
        var wb = XLSX.read(f.data instanceof ArrayBuffer ? new Uint8Array(f.data) : f.data, { type: 'array', cellDates: true, cellNF: false, cellStyles: false });
        var r = parseWorkbook(wb, f.name); r.name = f.name; return r;
      } catch (e) { return { name: f.name, error: /password|encrypt/i.test(e.message) ? '암호 걸린 파일' : e.message }; }
    });
    var R = process(parsed, opts);
    R.images = ex.images; R.others = ex.others; R.fileCount = ex.files.length;
    if (!R.invoices.length) return R;
    var x = await buildXlsx(client, R);
    R.xlsx = x.buffer; R.filename = x.filename;
    return R;
  }

  var api = { run: run, parseWorkbook: parseWorkbook, process: process, buildXlsx: buildXlsx, bizCheck: bizCheck, fmtBiz: fmtBiz };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Sugi = api;
})(typeof window !== 'undefined' ? window : globalThis);
