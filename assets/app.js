/* ════════════════════════════════════════════════════════════
   Doc Workbench — 문서 자동화 워크벤치
   모든 처리는 브라우저 안에서 이루어집니다. 서버 전송 없음.
   ════════════════════════════════════════════════════════════ */
(function () {
'use strict';

/* ─────────────────────────────────────────────
   0. 공통 유틸
   ───────────────────────────────────────────── */
const $ = (s, r) => (r || document).querySelector(s);
const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));

const esc = s => String(s).replace(/[&<>"']/g, c =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));

const nf = n => n.toLocaleString('ko-KR');

function bytes(n) {
  if (n < 1024) return n + ' B';
  if (n < 1048576) return (n / 1024).toFixed(1) + ' KB';
  return (n / 1048576).toFixed(1) + ' MB';
}

function baseName(name) { return name.replace(/\.[^.]+$/, ''); }
function extOf(name) { return (name.match(/\.([^.]+)$/) || [, ''])[1].toLowerCase(); }

/** 안전한 파일명 — 윈도 금지문자 제거 */
function safeName(s) {
  return String(s).replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim().slice(0, 120) || 'document';
}

function download(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

/* ─────────────────────────────────────────────
   1. 진행 표시
   ───────────────────────────────────────────── */
const Prog = {
  el: $('#prog'), t: $('#prog-t'), s: $('#prog-s'), f: $('#prog-fill'),
  open(title) { this.t.textContent = title; this.s.textContent = ''; this.f.style.width = '0%'; this.el.hidden = false; },
  set(pct, note) { this.f.style.width = Math.max(0, Math.min(100, pct)) + '%'; if (note != null) this.s.textContent = note; },
  close() { this.el.hidden = true; }
};
/** 무거운 루프 중간에 화면이 갱신되도록 양보 */
const breathe = () => new Promise(r => setTimeout(r, 0));

/* ─────────────────────────────────────────────
   2. 라이브러리 준비 확인
   ───────────────────────────────────────────── */
const Lib = {
  pdf: typeof pdfjsLib !== 'undefined',
  zip: typeof JSZip !== 'undefined',
  docx: typeof mammoth !== 'undefined'
};
if (Lib.pdf) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = window.DW_PDF_WORKER ||
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}
(function reportLibs() {
  const miss = [];
  if (!Lib.pdf) miss.push('PDF');
  if (!Lib.zip) miss.push('ZIP');
  if (!Lib.docx) miss.push('DOCX');
  const el = $('#net-state');
  if (!miss.length) { el.textContent = 'LIBRARIES READY'; el.style.color = 'var(--ok)'; }
  else { el.textContent = '로드 실패: ' + miss.join(', '); el.style.color = 'var(--warn)'; }
})();

/* ─────────────────────────────────────────────
   3. 화면 전환 · 드롭존
   ───────────────────────────────────────────── */
$$('.nav-i').forEach(btn => {
  btn.addEventListener('click', () => {
    $$('.nav-i').forEach(b => { b.classList.remove('is-on'); b.removeAttribute('aria-current'); });
    btn.classList.add('is-on'); btn.setAttribute('aria-current', 'page');
    $$('.mod').forEach(m => m.classList.remove('is-on'));
    $('#' + btn.dataset.go).classList.add('is-on');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });
});

/** 드롭존 + 파일 입력을 하나의 핸들러에 연결 */
function wireDrop(dropSel, inputSel, handler) {
  const drop = $(dropSel), input = $(inputSel);
  const fire = files => {
    const list = Array.from(files || []);
    if (list.length) handler(list);
    input.value = '';
  };
  drop.addEventListener('click', () => input.click());
  drop.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
  });
  input.addEventListener('change', e => fire(e.target.files));
  ['dragenter', 'dragover'].forEach(t =>
    drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('is-over'); }));
  ['dragleave', 'drop'].forEach(t =>
    drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('is-over'); }));
  drop.addEventListener('drop', e => fire(e.dataTransfer.files));
}

function showError(outSel, msg) {
  $(outSel).innerHTML = '<div class="msg err">' + esc(msg) + '</div>';
}

/* ─────────────────────────────────────────────
   4. 파일 읽기 — 형식별 텍스트 추출
   ───────────────────────────────────────────── */
async function readAsText(file) {
  const ext = extOf(file.name);

  if (['txt', 'md', 'csv', 'log', 'json', 'xml', 'html'].includes(ext)) {
    return await file.text();
  }

  if (ext === 'docx') {
    if (!Lib.docx) throw new Error('DOCX 처리 라이브러리를 불러오지 못했습니다.');
    const buf = await file.arrayBuffer();
    const r = await mammoth.extractRawText({ arrayBuffer: buf });
    return r.value;
  }

  if (ext === 'hwpx') return await readHwpx(file);

  if (ext === 'pdf') {
    const r = await pdfText(file, { clean: true, dehyphen: true, pageMark: false });
    return r.text;
  }

  if (ext === 'hwp') {
    throw new Error('구형 .hwp(한글 5.0 바이너리)는 브라우저에서 열 수 없습니다. 한글에서 .hwpx 또는 .pdf로 저장한 뒤 올려 주세요.');
  }

  // 알 수 없는 형식은 일단 텍스트로 시도
  return await file.text();
}

/** HWPX는 ZIP+XML — 본문 섹션에서 글자만 모읍니다 */
async function readHwpx(file) {
  if (!Lib.zip) throw new Error('ZIP 처리 라이브러리를 불러오지 못했습니다.');
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const names = Object.keys(zip.files)
    .filter(n => /Contents\/section\d+\.xml$/i.test(n))
    .sort((a, b) => {
      const na = +(a.match(/section(\d+)/i)[1]), nb = +(b.match(/section(\d+)/i)[1]);
      return na - nb;
    });
  if (!names.length) throw new Error('HWPX 본문(section) 파일을 찾지 못했습니다.');

  let out = [];
  for (const n of names) {
    const xml = await zip.file(n).async('string');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    // 문단(hp:p) 단위로 글자 조각(hp:t)을 이어 붙입니다
    const paras = Array.from(doc.getElementsByTagName('*'));
    let buf = [];
    for (const el of paras) {
      const tag = el.localName;
      if (tag === 'p') { if (buf.length) out.push(buf.join('')); buf = []; }
      else if (tag === 't') buf.push(el.textContent);
    }
    if (buf.length) out.push(buf.join(''));
  }
  return out.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

/* ─────────────────────────────────────────────
   5. 문서 분석
   ───────────────────────────────────────────── */
const RE = {
  hangul: /[가-힣ᄀ-ᇿ㄰-㆏]/,
  hanja: /[一-鿿㐀-䶿豈-﫿]/,
  latin: /[A-Za-z]/,
  digit: /[0-9]/,
  space: /\s/
};

const STOP = new Set([
  '그리고', '그러나', '하지만', '또는', '있다', '없다', '하는', '되는', '이다', '것이',
  '위해', '대한', '통해', '따라', '경우', '등의', '등을', '및', '수', '등', '이', '그', '저', '것',
  // 공문서에서 뜻을 거의 싣지 않는 상투어
  '있음', '없음', '통한', '위한', '관한', '따른', '대하여', '관하여', '하여', '되어', '함', '임', '됨',
  'the', 'and', 'for', 'with', 'that', 'this', 'from', 'are', 'was', 'were', 'not', 'but']);
const JOSA = ['으로써', '에서는', '으로는', '에게서', '이라고', '으로', '에서', '에게', '까지', '부터',
  '보다', '처럼', '이나', '라도', '은', '는', '이', '가', '을', '를', '의', '에', '와', '과', '도', '로', '만'];

function analyze(text) {
  const chars = Array.from(text);
  const total = chars.length;

  let hangul = 0, hanja = 0, latin = 0, digit = 0, space = 0, special = 0;
  const specialMap = new Map();

  for (const ch of chars) {
    if (RE.space.test(ch)) space++;
    else if (RE.hangul.test(ch)) hangul++;
    else if (RE.hanja.test(ch)) hanja++;
    else if (RE.latin.test(ch)) latin++;
    else if (RE.digit.test(ch)) digit++;
    else { special++; specialMap.set(ch, (specialMap.get(ch) || 0) + 1); }
  }

  const noSpace = total - space;
  const words = text.split(/\s+/).filter(Boolean);
  const paras = text.split(/\n\s*\n/).map(s => s.trim()).filter(Boolean);
  const lines = text.split(/\n/).filter(s => s.trim()).length;

  // 문장 분리 — 종결부호 기준
  const sentences = (text.match(/[^.!?。？！\n]+[.!?。？！]?/g) || [])
    .map(s => s.trim()).filter(s => s.length > 1);
  const sentLens = sentences.map(s => Array.from(s).length);
  const avgSent = sentLens.length ? sentLens.reduce((a, b) => a + b, 0) / sentLens.length : 0;
  const maxSent = sentLens.length ? Math.max(...sentLens) : 0;

  // 자주 쓴 단어 — 조사를 떼고 집계
  const freq = new Map();
  for (const w0 of words) {
    let w = w0.replace(/[^가-힣A-Za-z0-9]/g, '');
    if (w.length < 2) continue;
    if (/^\d+$/.test(w)) continue;   // 숫자만 남은 토큰(금액·비율 등)은 제외
    if (RE.hangul.test(w)) {
      for (const j of JOSA) {
        if (w.length > j.length + 1 && w.endsWith(j)) { w = w.slice(0, -j.length); break; }
      }
    }
    const k = RE.latin.test(w) && !RE.hangul.test(w) ? w.toLowerCase() : w;
    if (k.length < 2 || STOP.has(k)) continue;
    freq.set(k, (freq.get(k) || 0) + 1);
  }

  const topWords = Array.from(freq.entries()).sort((a, b) => b[1] - a[1]).slice(0, 24);
  const topSpecial = Array.from(specialMap.entries()).sort((a, b) => b[1] - a[1]).slice(0, 28);

  return {
    total, noSpace, space, hangul, hanja, latin, digit, special,
    words: words.length, paras: paras.length, lines,
    sentences: sentences.length, avgSent, maxSent,
    manuscript: total / 200,          // 원고지 200자 기준 매수
    a4: noSpace / 1600,               // A4 한 장 대략 1,600자(공백 제외) 기준
    doubleSpace: (text.match(/ {2,}/g) || []).length,
    topWords, topSpecial
  };
}

const COMPO = [
  ['hangul', '한글', '#6fd7ef'],
  ['hanja', '한자', '#a78bfa'],
  ['latin', '영문', '#5ede9c'],
  ['digit', '숫자', '#f3c969'],
  ['special', '특수문자', '#f08a9b'],
  ['space', '공백', '#3d434e']
];

function renderStats(st) {
  const tile = (k, v, u, hi) =>
    `<div class="tile${hi ? ' hi' : ''}"><div class="tile-k">${k}</div>
     <div class="tile-v">${v}${u ? `<span class="tile-u">${u}</span>` : ''}</div></div>`;

  let h = '<div class="tiles">' +
    tile('전체 글자수', nf(st.total), '자', true) +
    tile('공백 제외', nf(st.noSpace), '자') +
    tile('단어수', nf(st.words), '개', true) +
    tile('특수문자', nf(st.special), '자') +
    tile('문장수', nf(st.sentences), '개') +
    tile('문단수', nf(st.paras), '개') +
    tile('평균 문장 길이', st.avgSent.toFixed(1), '자') +
    tile('가장 긴 문장', nf(st.maxSent), '자') +
    tile('원고지', st.manuscript.toFixed(1), '매') +
    tile('A4 예상', st.a4.toFixed(1), '쪽') +
    '</div>';

  // 문자 구성 막대
  const bar = COMPO.map(([k, , c]) =>
    st.total ? `<i style="width:${(st[k] / st.total * 100).toFixed(2)}%;background:${c}"></i>` : ''
  ).join('');
  const legend = COMPO.map(([k, label, c]) =>
    `<div><i style="background:${c}"></i>${label}
     <b>${nf(st[k])} · ${st.total ? (st[k] / st.total * 100).toFixed(1) : '0.0'}%</b></div>`
  ).join('');

  h += `<div class="sub">문자 구성</div><div class="compo">${bar}</div><div class="legend">${legend}</div>`;

  if (st.topSpecial.length) {
    h += '<div class="sub">특수문자 종류별</div><div class="chips">' +
      st.topSpecial.map(([c, n]) =>
        `<span class="chip">${c === ' ' ? '␠' : esc(c)}<b>${nf(n)}</b></span>`).join('') +
      '</div>';
    if (st.doubleSpace) {
      h += `<p class="hint">연속된 공백이 ${nf(st.doubleSpace)}곳 있습니다. 공문서에서는 대개 정리 대상입니다.</p>`;
    }
  }

  if (st.topWords.length) {
    h += '<div class="sub">자주 쓴 단어</div><div class="scroll"><table class="tbl">' +
      '<tr><th>단어</th><th style="text-align:right">횟수</th><th style="text-align:right">비중</th></tr>' +
      st.topWords.map(([w, n]) =>
        `<tr><td>${esc(w)}</td><td class="n">${nf(n)}</td>` +
        `<td class="n">${st.words ? (n / st.words * 100).toFixed(2) : '0.00'}%</td></tr>`).join('') +
      '</table></div>' +
      '<p class="hint">한글은 조사를 떼어 내고 집계했습니다. 복합어나 고유명사는 정확히 분리되지 않을 수 있습니다.</p>';
  }
  return h;
}

/* ─────────────────────────────────────────────
   6. PDF 텍스트 추출
   ───────────────────────────────────────────── */
/** 페이지 구분선 — 삽입과 되읽기가 같은 정의를 쓰도록 한곳에 둡니다 */
const BAR = '─'.repeat(9);
const PAGE_MARK = n => BAR + ' p.' + n + ' ' + BAR;
const PAGE_MARK_RE = new RegExp('\\u2500{9} p\\.\\d+ \\u2500{9}\\n?', 'g');

/**
 * PDF에서 나온 텍스트 다듬기.
 * 글꼴에 글자가 없을 때 들어가는 널 문자, 줄바꿈 없는 공백(NBSP),
 * 서식용 보이지 않는 문자를 정리합니다.
 */
function normalizeText(s) {
  return s
    .replace(/\u0000/g, '')                       // 널 문자 — 글꼴에 글자가 없을 때
    .replace(/[   ]/g, ' ')        // 줄바꿈 없는 공백 → 보통 공백
    .replace(/[​-‍﻿]/g, '')        // 폭 없는 문자
    .replace(/�/g, '');                      // 대체 문자
}

async function pdfText(file, opt, onProg) {
  if (!Lib.pdf) throw new Error('PDF 라이브러리를 불러오지 못했습니다.');
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const pages = [];

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const tc = await page.getTextContent();

    // y좌표로 줄을 묶고, x좌표로 정렬해 원래 읽기 순서를 복원
    const rows = new Map();
    for (const it of tc.items) {
      if (!it.str) continue;
      const y = Math.round(it.transform[5]);
      let key = y;
      for (const k of rows.keys()) { if (Math.abs(k - y) <= 2) { key = k; break; } }
      if (!rows.has(key)) rows.set(key, []);
      rows.get(key).push({ x: it.transform[4], s: it.str, w: it.width || 0 });
    }

    const lines = Array.from(rows.entries())
      .sort((a, b) => b[0] - a[0])
      .map(([, items]) => {
        items.sort((a, b) => a.x - b.x);
        let s = '';
        let prevEnd = null;
        for (const it of items) {
          if (prevEnd !== null && it.x - prevEnd > 1.2 && !/\s$/.test(s)) s += ' ';
          s += it.s;
          prevEnd = it.x + it.w;
        }
        return normalizeText(s).replace(/\s+/g, ' ').trim();
      })
      .filter(Boolean);

    pages.push(lines);
    if (onProg) onProg(p, pdf.numPages);
    if (p % 8 === 0) await breathe();
  }

  let removed = [];
  if (opt.clean) removed = stripRunningHeads(pages);

  let out = pages.map((lines, i) => {
    let body = lines.join('\n');
    return opt.pageMark ? PAGE_MARK(i + 1) + '\n' + body : body;
  }).join('\n\n');

  // 줄 끝 하이픈 병합 — 일반 하이픈(-)과 소프트 하이픈(U+00AD) 모두 대상
  if (opt.dehyphen) out = out.replace(/([A-Za-z])[-­]\n([a-z])/g, '$1$2');
  out = out.replace(/­/g, '');   // 남은 소프트 하이픈은 보이지 않는 문자이므로 제거
  out = out.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();

  return { text: out, pages: pdf.numPages, removed };
}

/** 페이지마다 되풀이되는 머리말·꼬리말과 쪽번호를 제거 */
function stripRunningHeads(pages) {
  const removed = [];
  if (pages.length < 3) return removed;

  const tally = (getter) => {
    const m = new Map();
    pages.forEach(p => {
      const l = getter(p);
      if (!l) return;
      const key = l.replace(/\d+/g, '#').trim();
      if (key.length < 2 || key.length > 90) return;
      m.set(key, (m.get(key) || 0) + 1);
    });
    return m;
  };

  const threshold = Math.max(2, Math.ceil(pages.length * 0.5));

  [[p => p[0], 'head'], [p => p[p.length - 1], 'foot']].forEach(([get, kind]) => {
    const m = tally(get);
    for (const [key, n] of m) {
      if (n < threshold) continue;
      removed.push({ kind, sample: key, count: n });
      pages.forEach(p => {
        const idx = kind === 'head' ? 0 : p.length - 1;
        const l = p[idx];
        if (l && l.replace(/\d+/g, '#').trim() === key) p.splice(idx, 1);
      });
    }
  });

  // 숫자만 남은 줄(쪽번호)은 앞뒤 어디든 제거
  pages.forEach(p => {
    for (const idx of [0, p.length - 1]) {
      if (p[idx] && /^[-–—\s]*\d{1,4}[-–—\s]*$/.test(p[idx])) p.splice(idx, 1);
    }
  });

  return removed;
}

/* ─────────────────────────────────────────────
   7. PDF 이미지 추출
   ───────────────────────────────────────────── */

/** pdf.js 이미지 객체를 캔버스로 — 여러 픽셀 포맷 대응 */
function imgToCanvas(img) {
  const w = img.width, h = img.height;
  if (!w || !h) return null;
  const cv = document.createElement('canvas');
  cv.width = w; cv.height = h;
  const ctx = cv.getContext('2d');

  // JPG는 알파가 없으므로 흰 배경을 먼저 깔아 둡니다
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, w, h);

  if (img.bitmap) { ctx.drawImage(img.bitmap, 0, 0); return cv; }
  if (!img.data) return null;

  const id = ctx.createImageData(w, h);
  const out = id.data, src = img.data;

  if (img.kind === 3 || src.length === w * h * 4) {          // RGBA
    out.set(src.subarray(0, w * h * 4));
  } else if (img.kind === 2 || src.length === w * h * 3) {   // RGB
    for (let i = 0, j = 0; i < w * h; i++, j += 3) {
      out[i * 4] = src[j]; out[i * 4 + 1] = src[j + 1]; out[i * 4 + 2] = src[j + 2]; out[i * 4 + 3] = 255;
    }
  } else if (img.kind === 1) {                               // 1bit 흑백
    const rowBytes = (w + 7) >> 3;
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const bit = (src[y * rowBytes + (x >> 3)] >> (7 - (x & 7))) & 1;
        const v = bit ? 255 : 0, o = (y * w + x) * 4;
        out[o] = out[o + 1] = out[o + 2] = v; out[o + 3] = 255;
      }
    }
  } else if (src.length === w * h) {                          // 8bit 그레이
    for (let i = 0; i < w * h; i++) {
      out[i * 4] = out[i * 4 + 1] = out[i * 4 + 2] = src[i]; out[i * 4 + 3] = 255;
    }
  } else return null;

  // 알파를 흰 배경과 합성
  const tmp = document.createElement('canvas');
  tmp.width = w; tmp.height = h;
  tmp.getContext('2d').putImageData(id, 0, 0);
  ctx.drawImage(tmp, 0, 0);
  return cv;
}

const toJpeg = (cv, q) => new Promise(res => cv.toBlob(res, 'image/jpeg', q));

/**
 * 이미지 객체 찾기.
 * pdf.js는 그림을 page.objs 또는 commonObjs(여러 쪽이 공유하는 자원) 중 한곳에 둡니다.
 * 없는 쪽에 콜백을 걸면 영영 돌아오지 않으므로, 보유 여부를 먼저 확인하고 시간 제한도 둡니다.
 */
function getPdfObj(page, name) {
  return new Promise(resolve => {
    let done = false;
    const fin = v => { if (!done) { done = true; clearTimeout(timer); resolve(v || null); } };
    const timer = setTimeout(() => fin(null), 5000);

    const tryStore = store => {
      if (!store) return false;
      try { if (store.has && store.has(name)) { fin(store.get(name)); return true; } } catch (e) { }
      return false;
    };

    if (tryStore(page.objs)) return;
    if (tryStore(page.commonObjs)) return;

    // 아직 도착하지 않았을 수 있으므로 양쪽에 콜백을 걸어 둡니다
    try { page.objs.get(name, fin); } catch (e) { }
    try { page.commonObjs.get(name, fin); } catch (e) { }
  });
}

async function pdfImages(file, opt, onProg) {
  if (!Lib.pdf) throw new Error('PDF 라이브러리를 불러오지 못했습니다.');
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const shots = [];
  const seen = new Set();
  const OPS = pdfjsLib.OPS;

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);

    if (opt.mode === 'page') {
      const vp = page.getViewport({ scale: opt.dpi / 72 });
      const cv = document.createElement('canvas');
      cv.width = Math.round(vp.width); cv.height = Math.round(vp.height);
      const ctx = cv.getContext('2d');
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cv.width, cv.height);
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
      const blob = await toJpeg(cv, opt.quality);
      if (blob) shots.push({ page: p, w: cv.width, h: cv.height, blob, url: URL.createObjectURL(blob), on: true });
    } else {
      const ops = await page.getOperatorList();
      let idx = 0;
      for (let i = 0; i < ops.fnArray.length; i++) {
        const fn = ops.fnArray[i];
        const isX = fn === OPS.paintImageXObject || fn === OPS.paintJpegXObject;
        const isInline = fn === OPS.paintInlineImageXObject;
        if (!isX && !isInline) continue;

        let img = null;
        if (isInline) {
          img = ops.argsArray[i][0];
        } else {
          const name = ops.argsArray[i][0];
          // 이름은 문서 안에서 고유하므로, 여러 쪽에 반복 등장하는 로고 등은 한 번만 꺼냅니다
          if (seen.has(name)) continue;
          seen.add(name);
          img = await getPdfObj(page, name);
        }
        if (!img) continue;
        if (img.width < opt.min || img.height < opt.min) continue;

        const cv = imgToCanvas(img);
        if (!cv) continue;
        const blob = await toJpeg(cv, opt.quality);
        if (!blob) continue;
        idx++;
        shots.push({ page: p, index: idx, w: cv.width, h: cv.height, blob, url: URL.createObjectURL(blob), on: true });
      }
    }

    if (onProg) onProg(p, pdf.numPages, shots.length);
    await breathe();
  }
  return { shots, pages: pdf.numPages };
}

/* ─────────────────────────────────────────────
   8. 문서 빌더 — DOCX / HWPX / HTML
   ───────────────────────────────────────────── */
const EMU_IN = 914400;
const MAX_W_EMU = Math.round(6.0 * EMU_IN);   // 본문 폭 약 6인치

function xmlDecl(s) { return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s; }

/** 텍스트와 이미지를 담은 DOCX 생성 */
async function buildDocx(title, blocks) {
  if (!Lib.zip) throw new Error('ZIP 라이브러리를 불러오지 못했습니다.');
  const zip = new JSZip();
  const rels = [];
  const body = [];
  let imgN = 0;

  const para = (text, style) => {
    const rPr = '<w:rPr><w:rFonts w:ascii="맑은 고딕" w:eastAsia="맑은 고딕" w:hAnsi="맑은 고딕"/>' +
      (style === 'h1' ? '<w:b/><w:sz w:val="32"/>' : '<w:sz w:val="20"/>') + '</w:rPr>';
    return '<w:p><w:pPr>' + (style === 'h1' ? '<w:spacing w:after="240"/>' : '<w:spacing w:after="120"/>') +
      '</w:pPr><w:r>' + rPr + '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>';
  };

  body.push(para(title, 'h1'));

  for (const b of blocks) {
    if (b.type === 'text') {
      for (const line of b.value.split(/\n/)) {
        body.push(line.trim() ? para(line) : '<w:p/>');
      }
    } else if (b.type === 'image') {
      imgN++;
      const rid = 'rIdImg' + imgN;
      const fname = 'image' + imgN + '.jpg';
      zip.file('word/media/' + fname, b.blob);
      rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${fname}"/>`);

      const ratio = b.h / b.w;
      const cx = Math.min(MAX_W_EMU, b.w * EMU_IN / 96);
      const cy = Math.round(cx * ratio);
      body.push(
        '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="160"/></w:pPr><w:r><w:drawing>' +
        `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
        `<wp:extent cx="${Math.round(cx)}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
        `<wp:docPr id="${imgN}" name="Picture ${imgN}"/>` +
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        `<pic:nvPicPr><pic:cNvPr id="${imgN}" name="${fname}"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(cx)}" cy="${cy}"/></a:xfrm>` +
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
        '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
      );
      if (b.caption) body.push(para(b.caption));
    }
  }

  const doc = xmlDecl(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
    '<w:body>' + body.join('') +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1418" w:right="1134" w:bottom="1418" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>' +
    '</w:sectPr></w:body></w:document>');

  zip.file('[Content_Types].xml', xmlDecl(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="jpg" ContentType="image/jpeg"/>' +
    '<Default Extension="jpeg" ContentType="image/jpeg"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>'));

  zip.file('_rels/.rels', xmlDecl(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>'));

  zip.file('word/_rels/document.xml.rels', xmlDecl(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    rels.join('') + '</Relationships>'));

  zip.file('word/document.xml', doc);

  return await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
}

/** 최소 사양 HWPX 생성 — 본문 텍스트 전용(실험적) */
async function buildHwpx(title, text) {
  if (!Lib.zip) throw new Error('ZIP 라이브러리를 불러오지 못했습니다.');
  const zip = new JSZip();

  // mimetype은 압축하지 않고 가장 먼저 넣습니다
  zip.file('mimetype', 'application/hwp+zip', { compression: 'STORE' });

  zip.file('version.xml', xmlDecl(
    '<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" ' +
    'tagetApplication="WORDPROCESSOR" major="5" minor="0" micro="5" buildNumber="0" ' +
    'os="10" xmlVersion="1.4" application="Doc Workbench" appVersion="1.0"/>'));

  zip.file('META-INF/container.xml', xmlDecl(
    '<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container">' +
    '<ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" ' +
    'media-type="application/hwpml-package+xml"/></ocf:rootfiles></ocf:container>'));

  zip.file('META-INF/manifest.xml', xmlDecl(
    '<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" version="1.2">' +
    '<odf:file-entry odf:full-path="/" odf:media-type="application/hwp+zip"/>' +
    '<odf:file-entry odf:full-path="Contents/content.hpf" odf:media-type="application/hwpml-package+xml"/>' +
    '<odf:file-entry odf:full-path="Contents/header.xml" odf:media-type="application/xml"/>' +
    '<odf:file-entry odf:full-path="Contents/section0.xml" odf:media-type="application/xml"/>' +
    '</odf:manifest>'));

  zip.file('Contents/content.hpf', xmlDecl(
    '<opf:package xmlns:opf="http://www.idpf.org/2007/opf/" ' +
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" version="" unique-identifier="" id="">' +
    '<opf:metadata><opf:title>' + esc(title) + '</opf:title>' +
    '<opf:language>ko</opf:language></opf:metadata>' +
    '<opf:manifest>' +
    '<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>' +
    '<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>' +
    '</opf:manifest>' +
    '<opf:spine><opf:itemref idref="header" linear="yes"/>' +
    '<opf:itemref idref="section0" linear="yes"/></opf:spine>' +
    '</opf:package>'));

  const LANGS = ['HANGUL', 'LATIN', 'HANJA', 'JAPANESE', 'OTHER', 'SYMBOL', 'USER'];
  const fontfaces = LANGS.map(l =>
    `<hh:fontface lang="${l}" fontCnt="1">` +
    '<hh:font id="0" face="함초롬바탕" type="TTF" isEmbedded="0">' +
    '<hh:typeInfo familyType="FCAT_UNKNOWN" weight="0" proportion="0" contrast="0" ' +
    'strokeVariation="0" armStyle="0" letterform="0" midline="0" xHeight="0"/>' +
    '</hh:font></hh:fontface>').join('');

  zip.file('Contents/header.xml', xmlDecl(
    '<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" ' +
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" version="1.4" secCnt="1">' +
    '<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>' +
    '<hh:refList>' +
    `<hh:fontfaces itemCnt="${LANGS.length}">${fontfaces}</hh:fontfaces>` +
    '<hh:borderFills itemCnt="1"><hh:borderFill id="1" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">' +
    '<hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>' +
    '<hh:leftBorder type="NONE" width="0.1 mm" color="#000000"/><hh:rightBorder type="NONE" width="0.1 mm" color="#000000"/>' +
    '<hh:topBorder type="NONE" width="0.1 mm" color="#000000"/><hh:bottomBorder type="NONE" width="0.1 mm" color="#000000"/>' +
    '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/></hh:borderFill></hh:borderFills>' +
    '<hh:charProperties itemCnt="1">' +
    '<hh:charPr id="0" height="1000" textColor="#000000" shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="1">' +
    '<hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>' +
    '<hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>' +
    '<hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>' +
    '<hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>' +
    '<hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>' +
    '</hh:charPr></hh:charProperties>' +
    '<hh:tabProperties itemCnt="1"><hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/></hh:tabProperties>' +
    '<hh:numberings itemCnt="0"/>' +
    '<hh:paraProperties itemCnt="1">' +
    '<hh:paraPr id="0" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="1" suppressLineNumbers="0" checked="0">' +
    '<hh:align horizontal="JUSTIFY" vertical="BASELINE"/>' +
    '<hh:heading type="NONE" idRef="0" level="0"/>' +
    '<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="KEEP_WORD" widowOrphan="0" ' +
    'keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>' +
    '<hh:margin><hc:intent value="0" unit="HWPUNIT"/><hc:left value="0" unit="HWPUNIT"/>' +
    '<hc:right value="0" unit="HWPUNIT"/><hc:prev value="0" unit="HWPUNIT"/>' +
    '<hc:next value="0" unit="HWPUNIT"/></hh:margin>' +
    '<hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>' +
    '<hh:border borderFillIDRef="1" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" ' +
    'connect="0" ignoreMargin="0"/>' +
    '</hh:paraPr></hh:paraProperties>' +
    '<hh:styles itemCnt="1">' +
    '<hh:style id="0" type="PARA" name="바탕글" engName="Normal" paraPrIDRef="0" charPrIDRef="0" ' +
    'nextStyleIDRef="0" langID="1042" lockForm="0"/></hh:styles>' +
    '</hh:refList></hh:head>'));

  const paras = text.split(/\n/).map((line, i) =>
    '<hp:p id="' + i + '" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">' +
    '<hp:run charPrIDRef="0">' +
    (line.trim() ? '<hp:t>' + esc(line) + '</hp:t>' : '<hp:t/>') +
    '</hp:run></hp:p>').join('');

  zip.file('Contents/section0.xml', xmlDecl(
    '<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" ' +
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" ' +
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">' + paras + '</hs:sec>'));

  return await zip.generateAsync({ type: 'blob', mimeType: 'application/hwp+zip', compression: 'DEFLATE' });
}

/** 어디서나 열리는 HTML 문서 */
function buildHtmlDoc(title, blocks) {
  const parts = blocks.map(b => {
    if (b.type === 'text') {
      return b.value.split(/\n\s*\n/).filter(s => s.trim())
        .map(p => '<p>' + esc(p).replace(/\n/g, '<br>') + '</p>').join('\n');
    }
    return '<figure><img src="' + b.dataUrl + '" alt="' + esc(b.caption || '그림') + '">' +
      (b.caption ? '<figcaption>' + esc(b.caption) + '</figcaption>' : '') + '</figure>';
  }).join('\n');

  return '<!DOCTYPE html>\n<html lang="ko"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>' + esc(title) + '</title><style>' +
    'body{font-family:Pretendard,-apple-system,system-ui,sans-serif;max-width:820px;margin:48px auto;' +
    'padding:0 20px;line-height:1.85;color:#1a1d21}' +
    'h1{font-size:26px;margin-bottom:28px;padding-bottom:14px;border-bottom:2px solid #1a1d21}' +
    'p{margin:0 0 14px}figure{margin:26px 0;text-align:center}' +
    'img{max-width:100%;height:auto;border:1px solid #dcdfe4}' +
    'figcaption{font-size:13px;color:#6b7480;margin-top:8px}' +
    '@media print{body{margin:0}}</style></head><body>' +
    '<h1>' + esc(title) + '</h1>' + parts + '</body></html>';
}

const blobToDataUrl = blob => new Promise(res => {
  const r = new FileReader(); r.onload = () => res(r.result); r.readAsDataURL(blob);
});

/* ═════════════════════════════════════════════
   MODULE 01 — 문서 분석
   ═════════════════════════════════════════════ */
wireDrop('#d-analyze', '#f-analyze', async files => {
  const out = $('#o-analyze');
  out.innerHTML = '';
  Prog.open('문서를 읽는 중…');

  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    Prog.set((i / files.length) * 100, f.name);
    const card = document.createElement('div');
    card.className = 'res';

    try {
      const text = await readAsText(f);
      if (!text.trim()) throw new Error('읽어낼 텍스트가 없습니다. 스캔본 PDF이거나 이미지로만 구성된 문서일 수 있습니다.');
      const st = analyze(text);
      card.innerHTML =
        '<div class="res-h"><span class="res-n">' + esc(f.name) + '</span>' +
        '<span class="res-m">' + bytes(f.size) + '</span></div>' +
        '<div class="res-b">' + renderStats(st) +
        '<div class="btns">' +
        '<button class="btn sm" data-act="csv">통계 CSV로 저장</button>' +
        '<button class="btn sm" data-act="txt">추출 텍스트 저장</button>' +
        '</div></div>';

      card.querySelector('[data-act="csv"]').addEventListener('click', () => {
        const rows = [
          ['항목', '값'],
          ['파일명', f.name],
          ['전체 글자수', st.total], ['공백 제외 글자수', st.noSpace],
          ['단어수', st.words], ['문장수', st.sentences], ['문단수', st.paras], ['줄수', st.lines],
          ['한글', st.hangul], ['한자', st.hanja], ['영문', st.latin], ['숫자', st.digit],
          ['특수문자', st.special], ['공백', st.space],
          ['평균 문장 길이', st.avgSent.toFixed(1)], ['가장 긴 문장', st.maxSent],
          ['원고지(200자)', st.manuscript.toFixed(1)], ['A4 예상 쪽수', st.a4.toFixed(1)],
          [], ['특수문자', '횟수'], ...st.topSpecial,
          [], ['자주 쓴 단어', '횟수'], ...st.topWords
        ];
        const csv = '﻿' + rows.map(r => r.map(c => '"' + String(c == null ? '' : c).replace(/"/g, '""') + '"').join(',')).join('\r\n');
        download(new Blob([csv], { type: 'text/csv;charset=utf-8' }), safeName(baseName(f.name)) + '_통계.csv');
      });

      card.querySelector('[data-act="txt"]').addEventListener('click', () => {
        download(new Blob(['﻿' + text], { type: 'text/plain;charset=utf-8' }),
          safeName(baseName(f.name)) + '.txt');
      });

    } catch (err) {
      card.innerHTML =
        '<div class="res-h"><span class="res-n">' + esc(f.name) + '</span></div>' +
        '<div class="res-b"><div class="msg err">' + esc(err.message || String(err)) + '</div></div>';
    }
    out.appendChild(card);
    await breathe();
  }
  Prog.close();
});

/* ═════════════════════════════════════════════
   MODULE 02 — 텍스트 추출
   ═════════════════════════════════════════════ */
wireDrop('#d-text', '#f-text', async files => {
  const out = $('#o-text');
  out.innerHTML = '';
  const opt = { clean: $('#t-clean').checked, dehyphen: $('#t-dehyphen').checked, pageMark: $('#t-pagemark').checked };
  const made = [];
  Prog.open('PDF에서 텍스트를 꺼내는 중…');

  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    const card = document.createElement('div');
    card.className = 'res';
    try {
      const r = await pdfText(f, opt, (p, tot) => {
        Prog.set(((i + p / tot) / files.length) * 100, f.name + ' — ' + p + '/' + tot + '쪽');
      });
      const name = safeName(baseName(f.name)) + '.txt';
      const blob = new Blob(['﻿' + r.text], { type: 'text/plain;charset=utf-8' });
      made.push({ name, blob });

      const st = analyze(r.text);
      let note = '';
      if (!r.text.trim()) {
        note = '<div class="msg warn">텍스트가 검출되지 않았습니다. 스캔해서 만든 이미지 PDF로 보입니다. 이 경우 문자 인식(OCR)이 필요하며, 지금 버전에서는 지원하지 않습니다.</div>';
      } else if (r.removed.length) {
        note = '<div class="msg ok">반복 머리말·꼬리말 ' + r.removed.length + '종을 제거했습니다 — ' +
          r.removed.map(x => '“' + esc(x.sample.slice(0, 40)) + '” ' + x.count + '쪽').join(', ') + '</div>';
      }

      card.innerHTML =
        '<div class="res-h"><span class="res-n">' + esc(f.name) + '</span>' +
        '<span class="res-m">' + r.pages + '쪽 · ' + nf(st.total) + '자</span></div>' +
        '<div class="res-b">' + note +
        '<div class="pre">' + esc(r.text.slice(0, 2600)) + (r.text.length > 2600 ? '\n\n… 이하 생략' : '') + '</div>' +
        '<div class="btns"><button class="btn pri" data-act="dl">.txt 저장</button></div></div>';

      card.querySelector('[data-act="dl"]').addEventListener('click', () => download(blob, name));
    } catch (err) {
      card.innerHTML = '<div class="res-h"><span class="res-n">' + esc(f.name) + '</span></div>' +
        '<div class="res-b"><div class="msg err">' + esc(err.message || String(err)) + '</div></div>';
    }
    out.appendChild(card);
  }

  if (made.length > 1 && Lib.zip) {
    const bar = document.createElement('div');
    bar.className = 'btns';
    bar.innerHTML = '<button class="btn pri">전체 ' + made.length + '건 ZIP으로 저장</button>';
    bar.firstChild.addEventListener('click', async () => {
      const zip = new JSZip();
      made.forEach(m => zip.file(m.name, m.blob));
      download(await zip.generateAsync({ type: 'blob' }), 'extracted_text.zip');
    });
    out.appendChild(bar);
  }
  Prog.close();
});

/* ═════════════════════════════════════════════
   MODULE 03 — 이미지 추출
   ═════════════════════════════════════════════ */
$('#i-q').addEventListener('input', e => { $('#i-qv').value = e.target.value; });
$('#i-mode').addEventListener('change', e => { $('#i-dpi-wrap').hidden = e.target.value !== 'page'; });

wireDrop('#d-image', '#f-image', async files => {
  const out = $('#o-image');
  out.innerHTML = '';
  const opt = {
    mode: $('#i-mode').value,
    min: Math.max(0, +$('#i-min').value || 0),
    quality: (+$('#i-q').value || 88) / 100,
    dpi: +$('#i-dpi').value || 150
  };
  Prog.open('PDF에서 그림을 꺼내는 중…');

  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    const card = document.createElement('div');
    card.className = 'res';
    try {
      const r = await pdfImages(f, opt, (p, tot, found) => {
        Prog.set(((i + p / tot) / files.length) * 100, f.name + ' — ' + p + '/' + tot + '쪽 · ' + found + '개 발견');
      });

      if (!r.shots.length) {
        card.innerHTML = '<div class="res-h"><span class="res-n">' + esc(f.name) + '</span></div>' +
          '<div class="res-b"><div class="msg warn">조건에 맞는 그림을 찾지 못했습니다. 최소 크기를 낮추거나, 추출 방식을 “페이지 전체를 그림으로”로 바꿔 보세요. 도표나 수식이 선·글자로 그려진 PDF는 삽입 그림으로 잡히지 않습니다.</div></div>';
        out.appendChild(card);
        continue;
      }

      const grid = r.shots.map((s, idx) =>
        '<div class="shot" data-i="' + idx + '">' +
        '<input class="shot-c" type="checkbox" checked aria-label="' + (idx + 1) + '번 그림 선택">' +
        '<button class="shot-t" data-open="' + idx + '" aria-label="원본 크기로 열기">' +
        '<img src="' + s.url + '" alt="' + (s.page) + '쪽 그림" loading="lazy"></button>' +
        '<div class="shot-m"><span>p.' + s.page + '</span><span>' + s.w + '×' + s.h + '</span></div></div>'
      ).join('');

      card.innerHTML =
        '<div class="res-h"><span class="res-n">' + esc(f.name) + '</span>' +
        '<span class="res-m">' + r.pages + '쪽 · 그림 ' + r.shots.length + '개</span></div>' +
        '<div class="res-b"><div class="imgs">' + grid + '</div>' +
        '<div class="btns">' +
        '<button class="btn sm" data-act="all">전체 선택</button>' +
        '<button class="btn sm" data-act="none">전체 해제</button>' +
        '<button class="btn pri" data-act="zip">선택한 그림 ZIP으로 저장</button>' +
        '</div></div>';

      const shotEls = $$('.shot', card);
      const sync = () => shotEls.forEach((el, idx) => {
        r.shots[idx].on = $('.shot-c', el).checked;
        el.classList.toggle('off', !r.shots[idx].on);
      });
      shotEls.forEach(el => $('.shot-c', el).addEventListener('change', sync));
      $$('[data-open]', card).forEach(b => b.addEventListener('click', () => {
        window.open(r.shots[+b.dataset.open].url, '_blank', 'noopener');
      }));
      card.querySelector('[data-act="all"]').addEventListener('click', () => {
        shotEls.forEach(el => { $('.shot-c', el).checked = true; }); sync();
      });
      card.querySelector('[data-act="none"]').addEventListener('click', () => {
        shotEls.forEach(el => { $('.shot-c', el).checked = false; }); sync();
      });
      card.querySelector('[data-act="zip"]').addEventListener('click', async () => {
        sync();
        const picked = r.shots.filter(s => s.on);
        if (!picked.length) return alert('선택된 그림이 없습니다.');
        if (!Lib.zip) return alert('ZIP 라이브러리를 불러오지 못했습니다.');
        const zip = new JSZip();
        const base = safeName(baseName(f.name));
        picked.forEach((s, k) => {
          const nm = base + '_p' + String(s.page).padStart(3, '0') + '_' + String(k + 1).padStart(2, '0') + '.jpg';
          zip.file(nm, s.blob);
        });
        download(await zip.generateAsync({ type: 'blob' }), base + '_images.zip');
      });

    } catch (err) {
      card.innerHTML = '<div class="res-h"><span class="res-n">' + esc(f.name) + '</span></div>' +
        '<div class="res-b"><div class="msg err">' + esc(err.message || String(err)) + '</div></div>';
    }
    out.appendChild(card);
  }
  Prog.close();
});

/* ═════════════════════════════════════════════
   MODULE 04 — 문서 조립
   ═════════════════════════════════════════════ */
wireDrop('#d-build', '#f-build', async files => {
  const f = files[0];
  const out = $('#o-build');
  out.innerHTML = '';
  const title = ($('#b-title').value || '').trim() || baseName(f.name);
  const place = $('#b-place').value;
  const minPx = Math.max(0, +$('#b-min').value || 0);

  Prog.open('텍스트와 그림을 꺼내는 중…');
  try {
    const tr = await pdfText(f, { clean: $('#b-clean').checked, dehyphen: true, pageMark: false },
      (p, tot) => Prog.set((p / tot) * 45, '텍스트 ' + p + '/' + tot + '쪽'));

    const ir = await pdfImages(f, { mode: 'embedded', min: minPx, quality: 0.9, dpi: 150 },
      (p, tot, n) => Prog.set(45 + (p / tot) * 45, '그림 ' + p + '/' + tot + '쪽 · ' + n + '개'));

    Prog.set(92, '문서를 조립하는 중…');

    // 본문과 그림을 하나의 흐름으로 배치
    const blocks = [];
    if (place === 'inline' && ir.shots.length) {
      // 쪽 단위로 텍스트를 나눈 뒤 해당 쪽 그림을 뒤에 붙입니다
      const perPage = await pdfText(f, { clean: $('#b-clean').checked, dehyphen: true, pageMark: true });
      // 빈 쪽을 걸러내면 쪽 번호와 그림이 어긋나므로, 앞쪽 빈 조각만 떼어 냅니다
      const chunks = perPage.text.split(PAGE_MARK_RE);
      if (chunks.length && !chunks[0].trim()) chunks.shift();
      const byPage = new Map();
      ir.shots.forEach(s => {
        if (!byPage.has(s.page)) byPage.set(s.page, []);
        byPage.get(s.page).push(s);
      });
      chunks.forEach((chunk, idx) => {
        if (chunk.trim()) blocks.push({ type: 'text', value: chunk.trim() });
        (byPage.get(idx + 1) || []).forEach((s, k) =>
          blocks.push({ type: 'image', blob: s.blob, w: s.w, h: s.h, caption: '[그림] ' + (idx + 1) + '쪽 ' + (k + 1) }));
      });
    } else {
      blocks.push({ type: 'text', value: tr.text });
      ir.shots.forEach((s, k) =>
        blocks.push({ type: 'image', blob: s.blob, w: s.w, h: s.h, caption: '[그림 ' + (k + 1) + '] ' + s.page + '쪽' }));
    }

    const st = analyze(tr.text);

    out.innerHTML =
      '<div class="res"><div class="res-h"><span class="res-n">' + esc(f.name) + '</span>' +
      '<span class="res-m">' + tr.pages + '쪽 · ' + nf(st.total) + '자 · 그림 ' + ir.shots.length + '개</span></div>' +
      '<div class="res-b">' +
      (ir.shots.length ? '' : '<div class="msg warn">삽입된 그림을 찾지 못해 본문만 담았습니다.</div>') +
      '<div class="pre">' + esc(tr.text.slice(0, 1400)) + (tr.text.length > 1400 ? '\n\n… 이하 생략' : '') + '</div>' +
      '<div class="btns">' +
      '<button class="btn pri" data-b="docx">DOCX로 저장 <span style="opacity:.7">(권장)</span></button>' +
      '<button class="btn" data-b="hwpx">HWPX로 저장 (실험적)</button>' +
      '<button class="btn" data-b="html">HTML로 저장</button>' +
      '<button class="btn" data-b="zip">원본 조각 ZIP (txt + jpg)</button>' +
      '</div>' +
      '<p class="hint">HWPX는 본문 텍스트만 담깁니다. 그림까지 함께 넣으려면 DOCX나 HTML을 쓰시고, 한글에서 불러오기로 여시면 됩니다.</p>' +
      '</div></div>';

    const base = safeName(title);

    $('[data-b="docx"]', out).addEventListener('click', async () => {
      Prog.open('DOCX를 만드는 중…'); Prog.set(50);
      try { download(await buildDocx(title, blocks), base + '.docx'); }
      catch (e) { alert('DOCX 생성 실패: ' + e.message); }
      Prog.close();
    });

    $('[data-b="hwpx"]', out).addEventListener('click', async () => {
      Prog.open('HWPX를 만드는 중…'); Prog.set(50);
      try { download(await buildHwpx(title, tr.text), base + '.hwpx'); }
      catch (e) { alert('HWPX 생성 실패: ' + e.message); }
      Prog.close();
    });

    $('[data-b="html"]', out).addEventListener('click', async () => {
      Prog.open('HTML을 만드는 중…'); Prog.set(40);
      const withData = [];
      for (const b of blocks) {
        if (b.type === 'image') withData.push(Object.assign({}, b, { dataUrl: await blobToDataUrl(b.blob) }));
        else withData.push(b);
      }
      download(new Blob([buildHtmlDoc(title, withData)], { type: 'text/html;charset=utf-8' }), base + '.html');
      Prog.close();
    });

    $('[data-b="zip"]', out).addEventListener('click', async () => {
      if (!Lib.zip) return alert('ZIP 라이브러리를 불러오지 못했습니다.');
      Prog.open('ZIP으로 묶는 중…'); Prog.set(50);
      const zip = new JSZip();
      zip.file(base + '.txt', '﻿' + tr.text);
      const folder = zip.folder('images');
      ir.shots.forEach((s, k) =>
        folder.file(base + '_p' + String(s.page).padStart(3, '0') + '_' + String(k + 1).padStart(2, '0') + '.jpg', s.blob));
      download(await zip.generateAsync({ type: 'blob' }), base + '_추출.zip');
      Prog.close();
    });

  } catch (err) {
    showError('#o-build', err.message || String(err));
  }
  Prog.close();
});

})();
