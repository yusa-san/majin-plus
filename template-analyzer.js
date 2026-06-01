/**
 * PPTX テンプレート解析エンジン（定型処理）
 *
 * テンプレ .pptx を解凍し、theme / slideMaster / slideLayouts に加えて
 * 「実スライド（ppt/slides/*.xml）＝例スライド」まで読み取り、
 *   - メタ（サイズ/形式）
 *   - カラーパレット（theme色＋9論理名対応＋役割推定）
 *   - フォント（major/minor/scriptフォント＋役割）
 *   - 装飾（master/layout の静的図形：フッターバー/ロゴ/ページ番号 等）
 *   - レイアウトカタログ（原型ごとの分類・推奨まじん式タイプ・枠座標・テキスト許容量）
 *   - 例スライド構成サマリ（本文/図表の中身は引用しない・要素種別と数のみ）
 * を構造化した「解析オブジェクト」を返す。
 *
 * さらに解析オブジェクトから template_analysis.md（外部LLM投入用）を生成する。
 *
 * 依存: JSZip (CDN), DOMParser (ブラウザ標準)
 * 公開API:
 *   window.analyzeTemplate(file) -> Promise<Analysis>
 *   window.buildTemplateAnalysisMarkdown(Analysis) -> string
 */
(function () {
  'use strict';

  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const EMU_PER_IN = 914400;
  const EMU_PER_PT = 12700;

  // 9論理パレット（まじん式+ 6.1）。テンプレ色の最近傍で対応付け。
  const PALETTE = {
    skyblue: '#61B0E2', green: '#238966', deepgreen: '#1A664C', yellow: '#F7B515',
    blue: '#3271AD', orange: '#E95541', red: '#D82430', white: '#FFFFFF', black: '#000000'
  };

  // ===== 汎用XMLヘルパー（template-loader.js と共通の実装） =====
  function parseXml(str) { return new DOMParser().parseFromString(str, 'application/xml'); }
  function emuToIn(v) { return Number(v || 0) / EMU_PER_IN; }
  function stripHash(c) { return c ? String(c).replace('#', '') : c; }

  function elementChildren(el) {
    const out = [];
    if (!el || !el.childNodes) return out;
    for (let i = 0; i < el.childNodes.length; i++) { if (el.childNodes[i].nodeType === 1) out.push(el.childNodes[i]); }
    return out;
  }
  function getByTag(parent, tag) {
    let els = parent.getElementsByTagName(tag);
    if (els.length) return els[0];
    const local = tag.indexOf(':') >= 0 ? tag.split(':')[1] : tag;
    els = parent.getElementsByTagNameNS('*', local);
    return els.length ? els[0] : null;
  }
  function getAllByLocal(parent, local) {
    let els = parent.getElementsByTagName(local);
    if (els && els.length) return Array.prototype.slice.call(els);
    els = parent.getElementsByTagNameNS('*', local.indexOf(':') >= 0 ? local.split(':')[1] : local);
    return Array.prototype.slice.call(els);
  }
  function directChild(el, local) {
    if (!el) return null;
    const kids = elementChildren(el);
    for (let i = 0; i < kids.length; i++) {
      if ((kids[i].localName || (kids[i].tagName || '').replace(/^[a-z]+:/, '')) === local) return kids[i];
    }
    return null;
  }
  function localName(node) { return node.localName || (node.tagName || '').replace(/^[a-z]+:/, ''); }

  function colorOf(el) {
    if (!el) return null;
    const srgb = el.getElementsByTagName('a:srgbClr')[0] || el.getElementsByTagNameNS(A, 'srgbClr')[0];
    if (srgb && srgb.getAttribute('val')) return '#' + srgb.getAttribute('val').toUpperCase();
    const sys = el.getElementsByTagName('a:sysClr')[0] || el.getElementsByTagNameNS(A, 'sysClr')[0];
    if (sys) return '#' + (sys.getAttribute('lastClr') || '000000').toUpperCase();
    return null;
  }

  // ===== テーマ（色・フォント） =====
  function isRealTypeface(name) { const n = (name || '').trim(); return !!n && n.charAt(0) !== '+'; }
  function parseTheme(themeDoc) {
    const colors = {};
    const clrScheme = getByTag(themeDoc, 'a:clrScheme');
    if (clrScheme) {
      elementChildren(clrScheme).forEach(function (child) {
        const name = child.localName || (child.tagName || '').replace(/^a:/, '');
        const c = colorOf(child);
        if (c) colors[name] = c;
      });
    }
    function scriptFontTypeface(scheme, script) {
      const list = getAllByLocal(scheme, 'a:font');
      for (let i = 0; i < list.length; i++) {
        const f = list[i];
        if (script == null || f.getAttribute('script') === script) {
          const t = f.getAttribute('typeface');
          if (isRealTypeface(t)) return t.trim();
        }
      }
      return '';
    }
    function tfOf(scheme, tag) {
      const el = scheme.getElementsByTagName('a:' + tag)[0] || scheme.getElementsByTagNameNS(A, tag)[0];
      const t = el && el.getAttribute('typeface');
      return isRealTypeface(t) ? t.trim() : '';
    }
    function fontFromScheme(tag) {
      const scheme = getByTag(themeDoc, tag);
      if (!scheme) return '';
      return tfOf(scheme, 'ea') || scriptFontTypeface(scheme, 'Jpan') || scriptFontTypeface(scheme, null) || tfOf(scheme, 'latin') || '';
    }
    function collectFonts(tag, out) {
      const scheme = getByTag(themeDoc, tag);
      if (!scheme) return;
      ['latin', 'ea', 'cs'].forEach(function (t) { const v = tfOf(scheme, t); if (v) out[v] = true; });
      getAllByLocal(scheme, 'a:font').forEach(function (f) { const v = f.getAttribute('typeface'); if (isRealTypeface(v)) out[v.trim()] = true; });
    }
    const majorFont = fontFromScheme('a:majorFont');
    const minorFont = fontFromScheme('a:minorFont');
    const fontSet = {};
    collectFonts('a:majorFont', fontSet);
    collectFonts('a:minorFont', fontSet);
    return { colors: colors, majorFont: majorFont, minorFont: minorFont, fonts: Object.keys(fontSet) };
  }
  function collectTypefacesFromDoc(doc, out) {
    if (!doc) return;
    ['a:latin', 'a:ea', 'a:cs', 'a:font'].forEach(function (tag) {
      getAllByLocal(doc, tag).forEach(function (el) {
        const n = (el.getAttribute('typeface') || '').trim();
        if (n && n.charAt(0) !== '+') out[n] = true;
      });
    });
  }

  function clrMapFrom(masterDoc) {
    const clrMap = {};
    const el = getByTag(masterDoc, 'p:clrMap');
    if (el) ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].forEach(function (k) { const v = el.getAttribute(k); if (v) clrMap[k] = v; });
    return clrMap;
  }
  function resolveSchemeColor(val, colors, clrMap) {
    if (!val) return null;
    if (clrMap && clrMap[val]) val = clrMap[val];
    return colors[val] || null;
  }
  function colorFromContainer(container, colors, clrMap) {
    if (!container) return null;
    const srgb = directChild(container, 'srgbClr');
    if (srgb && srgb.getAttribute('val')) return '#' + srgb.getAttribute('val').toUpperCase();
    const sch = directChild(container, 'schemeClr');
    if (sch) return resolveSchemeColor(sch.getAttribute('val'), colors, clrMap);
    return null;
  }
  function fillFrom(sp, spPr, colors, clrMap) {
    if (spPr) {
      if (directChild(spPr, 'noFill')) return null;
      const solid = directChild(spPr, 'solidFill');
      if (solid) { const c = colorFromContainer(solid, colors, clrMap); if (c) return c; }
      const grad = directChild(spPr, 'gradFill');
      if (grad) { const gsLst = directChild(grad, 'gsLst'); const gs = gsLst ? directChild(gsLst, 'gs') : null; if (gs) { const c = colorFromContainer(gs, colors, clrMap); if (c) return c; } }
    }
    const style = directChild(sp, 'style');
    if (style) { const fillRef = directChild(style, 'fillRef'); if (fillRef) { const c = colorFromContainer(fillRef, colors, clrMap); if (c) return c; } }
    return undefined;
  }
  function lineFrom(spPr, colors, clrMap) {
    if (!spPr) return undefined;
    const ln = directChild(spPr, 'ln');
    if (!ln) return undefined;
    if (directChild(ln, 'noFill')) return null;
    const solid = directChild(ln, 'solidFill');
    const color = solid ? colorFromContainer(solid, colors, clrMap) : null;
    if (!color) return undefined;
    const wEmu = ln.getAttribute('w');
    return { color: stripHash(color), width: wEmu ? Number(wEmu) / EMU_PER_PT : 1 };
  }
  function xfrmFrom(spPr) {
    const xfrm = spPr ? directChild(spPr, 'xfrm') : null;
    if (!xfrm) return null;
    const off = directChild(xfrm, 'off'), ext = directChild(xfrm, 'ext');
    if (!off || !ext) return null;
    const r = { x: emuToIn(off.getAttribute('x')), y: emuToIn(off.getAttribute('y')), w: emuToIn(ext.getAttribute('cx')), h: emuToIn(ext.getAttribute('cy')) };
    const rot = xfrm.getAttribute('rot'); if (rot) r.rotate = Number(rot) / 60000;
    if (xfrm.getAttribute('flipH') === '1') r.flipH = true;
    if (xfrm.getAttribute('flipV') === '1') r.flipV = true;
    return r;
  }
  function placeholderInfo(sp) {
    const nvSpPr = directChild(sp, 'nvSpPr');
    const nvPr = nvSpPr ? directChild(nvSpPr, 'nvPr') : null;
    const ph = nvPr ? directChild(nvPr, 'ph') : null;
    if (!ph) return null;
    return { type: ph.getAttribute('type') || '', idx: ph.getAttribute('idx') || '' };
  }
  function textFrom(sp, colors, clrMap) {
    const txBody = directChild(sp, 'txBody');
    if (!txBody) return null;
    const paras = elementChildren(txBody).filter(function (e) { return localName(e) === 'p'; });
    let str = '', size = null, color = null, bold = false;
    paras.forEach(function (p, pi) {
      if (pi > 0) str += '\n';
      elementChildren(p).forEach(function (node) {
        if (localName(node) === 'r') {
          const t = directChild(node, 't');
          if (t && t.textContent) str += t.textContent;
          const rPr = directChild(node, 'rPr');
          if (rPr && size == null) {
            const sz = rPr.getAttribute('sz'); if (sz) size = Number(sz) / 100;
            if (rPr.getAttribute('b') === '1') bold = true;
            const c = colorFromContainer(directChild(rPr, 'solidFill'), colors, clrMap);
            if (c) color = c;
          }
        }
      });
    });
    str = str.trim();
    return str ? { string: str, size: size, color: color, bold: bold } : null;
  }
  function mimeFromExt(path) {
    const m = /\.([a-zA-Z0-9]+)$/.exec(path);
    const ext = (m ? m[1] : 'png').toLowerCase();
    if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg';
    if (ext === 'gif') return 'image/gif';
    if (ext === 'svg') return 'image/svg+xml';
    if (ext === 'bmp') return 'image/bmp';
    if (ext === 'emf') return 'image/x-emf';
    if (ext === 'wmf') return 'image/x-wmf';
    return 'image/png';
  }
  function resolvePath(baseDir, target) {
    if (!target) return null;
    if (target.charAt(0) === '/') return target.slice(1);
    const parts = (baseDir + '/' + target).split('/');
    const stack = [];
    parts.forEach(function (p) { if (p === '' || p === '.') return; if (p === '..') stack.pop(); else stack.push(p); });
    return stack.join('/');
  }
  async function parseRels(zip, relsPath) {
    const map = {};
    const f = zip.file(relsPath);
    if (!f) return map;
    const doc = parseXml(await f.async('string'));
    let list = doc.getElementsByTagName('Relationship');
    if (!list.length) list = doc.getElementsByTagNameNS('*', 'Relationship');
    for (let i = 0; i < list.length; i++) {
      const id = list[i].getAttribute('Id'), target = list[i].getAttribute('Target'), type = list[i].getAttribute('Type');
      if (id && target) map[id] = { target: target, type: type || '' };
    }
    return map;
  }
  async function parseBgFill(doc, baseDir, zip, relsMap, colors, clrMap) {
    const bg = getByTag(doc, 'p:bg');
    if (!bg) return null;
    const bgPr = directChild(bg, 'bgPr');
    if (bgPr) {
      const solid = directChild(bgPr, 'solidFill');
      if (solid) { const c = colorFromContainer(solid, colors, clrMap); if (c) return { color: stripHash(c) }; }
      const bf = directChild(bgPr, 'blipFill');
      const blip = bf ? directChild(bf, 'blip') : null;
      if (blip) {
        const rel = relsMap && (relsMap[blip.getAttribute('r:embed')] || relsMap[blip.getAttribute('embed')]);
        const target = rel && rel.target;
        if (target) { const path = resolvePath(baseDir, target); const file = path && zip.file(path); if (file) { const b64 = await file.async('base64'); return { data: 'data:' + mimeFromExt(path) + ';base64,' + b64 }; } }
      }
    }
    const bgRef = directChild(bg, 'bgRef');
    if (bgRef) { const c = colorFromContainer(bgRef, colors, clrMap); if (c) return { color: stripHash(c) }; }
    return null;
  }

  // spTree から静的(非プレースホルダ)要素を PptxGenJS master object 配列へ変換
  async function buildStaticObjects(doc, baseDir, zip, relsMap, colors, clrMap) {
    const objects = [];
    const spTree = getByTag(doc, 'p:spTree');
    if (!spTree) return objects;
    const kids = elementChildren(spTree);
    for (let i = 0; i < kids.length; i++) {
      const node = kids[i], local = localName(node);
      if (local === 'sp') {
        if (placeholderInfo(node)) continue;
        const spPr = directChild(node, 'spPr');
        const r = xfrmFrom(spPr);
        if (!r) continue;
        const fill = fillFrom(node, spPr, colors, clrMap);
        const line = lineFrom(spPr, colors, clrMap);
        const o = { x: r.x, y: r.y, w: r.w, h: r.h };
        if (r.rotate) o.rotate = r.rotate;
        if (r.flipH) o.flipH = true;
        if (r.flipV) o.flipV = true;
        if (fill === null) o.fill = { type: 'none' };
        else if (fill) o.fill = { color: stripHash(fill) };
        else o.fill = { type: 'none' };
        if (line) o.line = line;
        objects.push({ rect: o });
        const tx = textFrom(node, colors, clrMap);
        if (tx) {
          const topt = { x: r.x, y: r.y, w: r.w, h: r.h, align: 'center', valign: 'middle' };
          if (tx.size) topt.fontSize = tx.size;
          if (tx.color) topt.color = stripHash(tx.color);
          if (tx.bold) topt.bold = true;
          objects.push({ text: { text: tx.string, options: topt } });
        }
      } else if (local === 'cxnSp') {
        const spPr = directChild(node, 'spPr');
        const r = xfrmFrom(spPr);
        if (!r) continue;
        const line = lineFrom(spPr, colors, clrMap) || { color: '888888', width: 1 };
        const o = { x: r.x, y: r.y, w: r.w, h: r.h, line: line };
        if (r.flipH) o.flipH = true;
        if (r.flipV) o.flipV = true;
        objects.push({ line: o });
      } else if (local === 'pic') {
        const spPr = directChild(node, 'spPr');
        const r = xfrmFrom(spPr);
        if (!r) continue;
        const bf = directChild(node, 'blipFill');
        const blip = bf ? directChild(bf, 'blip') : null;
        const rel = blip && relsMap && (relsMap[blip.getAttribute('r:embed')] || relsMap[blip.getAttribute('embed')]);
        const target = rel && rel.target;
        if (!target) continue;
        const path = resolvePath(baseDir, target);
        const file = path && zip.file(path);
        if (!file) continue;
        try { const b64 = await file.async('base64'); objects.push({ image: { x: r.x, y: r.y, w: r.w, h: r.h, data: 'data:' + mimeFromExt(path) + ';base64,' + b64 } }); }
        catch (e) { /* skip */ }
      }
    }
    return objects;
  }

  // spTree 直下のプレースホルダ {type,idx,rect}
  function placeholdersOf(doc) {
    const result = [];
    const spTree = getByTag(doc, 'p:spTree');
    if (!spTree) return result;
    elementChildren(spTree).forEach(function (node) {
      if (localName(node) !== 'sp') return;
      const ph = placeholderInfo(node);
      if (!ph) return;
      result.push({ type: ph.type, idx: ph.idx, rect: xfrmFrom(directChild(node, 'spPr')) });
    });
    return result;
  }

  // 非プレースホルダ要素の種別カウント（中身は持たない）
  function countNonPhElements(doc) {
    const c = { shapes: 0, pics: 0, lines: 0, tables: 0, charts: 0 };
    const spTree = getByTag(doc, 'p:spTree');
    if (!spTree) return c;
    elementChildren(spTree).forEach(function (node) {
      const local = localName(node);
      if (local === 'sp') { if (!placeholderInfo(node)) c.shapes++; }
      else if (local === 'pic') c.pics++;
      else if (local === 'cxnSp') c.lines++;
      else if (local === 'graphicFrame') {
        if (node.getElementsByTagName('a:tbl').length || node.getElementsByTagNameNS(A, 'tbl').length) c.tables++;
        else c.charts++;
      }
    });
    return c;
  }

  // ===== 分類・許容量 =====
  // OOXML layout type / 名前 / プレースホルダ・要素から まじん式カテゴリを推定
  function classify(layoutType, name, phs, counts) {
    const t = (layoutType || '').toLowerCase();
    const nm = (name || '');
    const bodyCount = phs.filter(function (p) { return p.type === 'body' || p.type === '' || p.type === 'subTitle'; }).length;
    if (/closing|結び|おわり|thank/i.test(nm)) return 'closing';
    if (t === 'title' || /title\s*slide|表紙|hyoshi/i.test(nm)) return 'title';
    if (t === 'sectionheader' || t === 'sechead' || /section|章|区切|divider/i.test(nm)) return 'section';
    if (/agenda|目次|アジェンダ/i.test(nm)) return 'content';
    if (/compar|比較|対比/i.test(nm)) return 'compare';
    if (counts && counts.tables > 0) return 'table';
    if (counts && (counts.lines >= 3 || counts.charts > 0)) return 'diagram';
    if (t === 'twoobj' || t === 'twotxtwoobj' || bodyCount >= 2) return 'twoCol';
    if (counts && counts.pics >= 3) return 'cards';
    if (t === 'title only' || t === 'titleonly') return 'content';
    if (t === 'obj' || t === 'tx' || t === 'objtx' || t === 'objonly') return 'content';
    if (t === 'blank') return 'other';
    return 'content';
  }
  const CATEGORY_TO_MAJIN = {
    title: 'title', section: 'section', content: 'content', twoCol: 'content（twoColumn）',
    compare: 'compare', cards: 'cards / headerCards', table: 'table',
    diagram: 'mermaid / flowChart / cycle', closing: 'closing', other: 'content'
  };

  // 枠サイズからテキスト許容量を概算（和文基準・あくまで目安）
  function estimateCapacity(titleFrame, bodyFrame) {
    function rough(widthIn, heightIn, fontPt) {
      if (!widthIn) return null;
      const charW = fontPt * 0.5 / 72;          // 全角1字 ≈ 0.5em
      const lineH = fontPt * 1.4 / 72;
      const perLine = Math.max(1, Math.floor(widthIn / charW));
      const lines = heightIn ? Math.max(1, Math.floor(heightIn / lineH)) : 1;
      return { perLine: perLine, lines: lines };
    }
    const cap = {};
    if (titleFrame) { const r = rough(titleFrame.width, titleFrame.height, 28); if (r) cap.titleChars = r.perLine * r.lines; }
    if (bodyFrame) { const r = rough(bodyFrame.width, bodyFrame.height, 18); if (r) { cap.bodyCharsPerLine = r.perLine; cap.bodyLines = r.lines; } }
    return cap;
  }

  function rectOf(r) { return r ? { left: r.x, top: r.y, width: r.w, height: r.h } : null; }

  // プレースホルダ群から title枠 / body枠（複数）/ 図枠 を解決（無xfrmはmasterから継承）
  const FIGURE_PH_TYPES = { pic: 1, chart: 1, tbl: 1, clipArt: 1, dgm: 1, media: 1 };
  function framesFromPhs(phs, masterPhs) {
    // 継承: 同 type+idx → 同 idx → 同 type の順でマスターから座標を引く（layoutにxfrmが無い場合）
    function inherit(type, idx) {
      let m = masterPhs.find(function (p) { return p.type === type && p.idx === idx && p.rect; });
      if (m) return m.rect;
      if (idx) { m = masterPhs.find(function (p) { return p.idx === idx && p.rect; }); if (m) return m.rect; }
      m = masterPhs.find(function (p) { return p.type === type && p.rect; });
      if (m) return m.rect;
      // 汎用(body/'')はマスターの body で代替
      if (type === '' || type === 'body') { m = masterPhs.find(function (p) { return (p.type === 'body' || p.type === '') && p.rect; }); if (m) return m.rect; }
      return null;
    }
    function resolve(p) { return p.rect || inherit(p.type, p.idx); }
    const titlePh = phs.find(function (p) { return p.type === 'ctrTitle' || p.type === 'title'; });
    const bodyPhs = phs.filter(function (p) { return p.type === 'body' || p.type === 'subTitle' || p.type === ''; });
    const figurePhs = phs.filter(function (p) { return FIGURE_PH_TYPES[p.type]; });
    const titleFrame = titlePh ? rectOf(resolve(titlePh)) : null;
    const bodyFrames = bodyPhs.map(function (p) { return rectOf(resolve(p)); }).filter(Boolean);
    const figureFrames = figurePhs.map(function (p) {
      const r = rectOf(resolve(p));
      return r ? Object.assign({ phType: p.type }, r) : null;
    }).filter(Boolean);
    return { titleFrame: titleFrame, bodyFrames: bodyFrames, figureFrames: figureFrames };
  }

  // body枠が無いレイアウト用に、タイトル下〜フッター上の領域を本文枠として合成する
  function synthBodyFrame(titleFrame, slideSize) {
    const W = (slideSize && slideSize.wIn) ? slideSize.wIn : 13.333;
    const H = (slideSize && slideSize.hIn) ? slideSize.hIn : 7.5;
    const left = titleFrame ? titleFrame.left : W * 0.06;
    const width = titleFrame ? titleFrame.width : W * 0.88;
    const top = titleFrame ? (titleFrame.top + titleFrame.height + H * 0.03) : H * 0.22;
    const bottom = H * 0.88; // フッター帯の上で止める
    return { left: left, top: top, width: width, height: Math.max(H * 0.2, bottom - top) };
  }

  // 9論理名へ最近傍対応（テンプレ色 → palette名）
  function hexToRgb(h) { h = String(h).replace('#', ''); return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]; }
  function nearestPaletteName(hex) {
    if (!hex) return null;
    const rgb = hexToRgb(hex);
    let best = null, bestD = Infinity;
    Object.keys(PALETTE).forEach(function (name) {
      const p = hexToRgb(PALETTE[name]);
      const d = (rgb[0] - p[0]) ** 2 + (rgb[1] - p[1]) ** 2 + (rgb[2] - p[2]) ** 2;
      if (d < bestD) { bestD = d; best = name; }
    });
    return best;
  }

  // ===== メイン解析 =====
  async function analyzeTemplate(file) {
    if (typeof JSZip === 'undefined') throw new Error('JSZip が読み込まれていません（CDN取得失敗の可能性）。');
    const zip = await JSZip.loadAsync(file);

    // --- メタ（サイズ/形式） ---
    let slideSize = null, format = '16:9';
    try {
      const presFile = zip.file('ppt/presentation.xml');
      if (presFile) {
        const sldSz = getByTag(parseXml(await presFile.async('string')), 'p:sldSz');
        if (sldSz) {
          const cx = Number(sldSz.getAttribute('cx') || 0), cy = Number(sldSz.getAttribute('cy') || 0);
          if (cx > 0 && cy > 0) {
            const wIn = cx / EMU_PER_IN, hIn = cy / EMU_PER_IN;
            slideSize = { wIn: wIn, hIn: hIn };
            const aspect = wIn / hIn, tol = 0.04;
            if (Math.abs(aspect - 16 / 9) <= tol) format = '16:9';
            else if (Math.abs(aspect - 10.83 / 7.5) <= tol) format = 'a4';
            else format = 'unsupported';
          }
        }
      }
    } catch (e) { console.warn('slide size detect failed', e); }

    // --- テーマ ---
    let themeStr = null;
    const themeFiles = zip.file(/^ppt\/theme\/theme\d+\.xml$/);
    if (zip.file('ppt/theme/theme1.xml')) themeStr = await zip.file('ppt/theme/theme1.xml').async('string');
    else if (themeFiles && themeFiles.length) themeStr = await themeFiles[0].async('string');
    if (!themeStr) throw new Error('テーマ(theme1.xml)が見つかりません。有効な.pptxファイルか確認してください。');
    const themeInfo = parseTheme(parseXml(themeStr));
    const colors = themeInfo.colors;
    const fontSet = {};
    (themeInfo.fonts || []).forEach(function (f) { fontSet[f] = true; });

    // --- マスター ---
    let masterPath = 'ppt/slideMasters/slideMaster1.xml';
    if (!zip.file(masterPath)) { const mf = zip.file(/^ppt\/slideMasters\/slideMaster\d+\.xml$/); masterPath = (mf && mf.length) ? mf[0].name : null; }
    let clrMap = {}, masterObjects = [], masterPhs = [], masterBg = null, backgroundColor = null;
    let masterDoc = null;
    if (masterPath) {
      masterDoc = parseXml(await zip.file(masterPath).async('string'));
      collectTypefacesFromDoc(masterDoc, fontSet);
      clrMap = clrMapFrom(masterDoc);
      const baseDir = masterPath.substring(0, masterPath.lastIndexOf('/'));
      const fileNm = masterPath.substring(masterPath.lastIndexOf('/') + 1);
      const rels = await parseRels(zip, baseDir + '/_rels/' + fileNm + '.rels');
      try { masterObjects = await buildStaticObjects(masterDoc, baseDir, zip, rels, colors, clrMap); } catch (e) { console.warn('master objects', e); }
      masterPhs = placeholdersOf(masterDoc);
      try { masterBg = await parseBgFill(masterDoc, baseDir, zip, rels, colors, clrMap); } catch (e) { console.warn('master bg', e); }
      backgroundColor = (masterBg && masterBg.color) ? ('#' + masterBg.color) : null;
    }

    // --- レイアウトカタログ ---
    const layoutFiles = (zip.file(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/) || []).sort(function (a, b) { return a.name.localeCompare(b.name); });
    const layoutByPath = {};
    const layoutCatalog = [];
    for (let i = 0; i < layoutFiles.length; i++) {
      const lf = layoutFiles[i];
      const doc = parseXml(await lf.async('string'));
      collectTypefacesFromDoc(doc, fontSet);
      const type = (doc.documentElement && doc.documentElement.getAttribute('type')) || '';
      const cSld = getByTag(doc, 'p:cSld');
      const name = cSld ? (cSld.getAttribute('name') || '') : '';
      const baseDir = lf.name.substring(0, lf.name.lastIndexOf('/'));
      const fileNm = lf.name.substring(lf.name.lastIndexOf('/') + 1);
      const rels = await parseRels(zip, baseDir + '/_rels/' + fileNm + '.rels');
      const phs = placeholdersOf(doc);
      const counts = countNonPhElements(doc);
      let lObjs = [];
      try { lObjs = await buildStaticObjects(doc, baseDir, zip, rels, colors, clrMap); } catch (e) { /* skip */ }
      const bg = await parseBgFill(doc, baseDir, zip, rels, colors, clrMap);
      const frames = framesFromPhs(phs, masterPhs);
      const category = classify(type, name, phs, counts);
      const id = 'L' + (i + 1);
      // body枠が無いレイアウトはタイトル下〜フッター上を合成（Phase B が確実に効くように）
      let bodyFrames = frames.bodyFrames;
      if (!bodyFrames.length && category !== 'title' && category !== 'section' && category !== 'closing') {
        bodyFrames = [synthBodyFrame(frames.titleFrame, slideSize)];
      }
      const entry = {
        id: id, source: 'layout', layoutType: type, name: name, category: category,
        suggestedType: CATEGORY_TO_MAJIN[category] || 'content',
        titleFrame: frames.titleFrame,
        bodyFrame: bodyFrames[0] || null,
        bodyFrames: bodyFrames,
        figureFrames: frames.figureFrames,        // 図/表/グラフ プレースホルダ位置（ダミー図描画用）
        columns: Math.max(1, bodyFrames.length),
        objects: masterObjects.concat(lObjs),     // 装飾（master＋layout 静的図形）
        background: bg || masterBg || null,
        capacity: estimateCapacity(frames.titleFrame, bodyFrames[0] || null),
        usedBy: []                                 // 例スライドからの被参照
      };
      layoutCatalog.push(entry);
      layoutByPath[lf.name] = entry;
    }

    // --- 例スライド（構成サマリ。本文/図表の中身は持たない） ---
    const exampleSlides = [];
    const slideFiles = (zip.file(/^ppt\/slides\/slide\d+\.xml$/) || []).sort(function (a, b) {
      function n(f) { const m = /slide(\d+)\.xml$/.exec(f.name); return m ? Number(m[1]) : 0; }
      return n(a) - n(b);
    });
    for (let i = 0; i < slideFiles.length; i++) {
      const sf = slideFiles[i];
      const doc = parseXml(await sf.async('string'));
      const baseDir = sf.name.substring(0, sf.name.lastIndexOf('/'));
      const fileNm = sf.name.substring(sf.name.lastIndexOf('/') + 1);
      const rels = await parseRels(zip, baseDir + '/_rels/' + fileNm + '.rels');
      // slide -> layout
      let layoutEntry = null;
      Object.keys(rels).forEach(function (rid) {
        const rel = rels[rid];
        if (rel.type && rel.type.indexOf('slideLayout') >= 0) {
          const path = resolvePath(baseDir, rel.target);
          if (path && layoutByPath[path]) layoutEntry = layoutByPath[path];
        }
      });
      const phs = placeholdersOf(doc);
      const counts = countNonPhElements(doc);
      const baseCat = layoutEntry ? layoutEntry.category : 'content';
      // スライド固有要素で分類を補正
      const category = (function () {
        if (counts.tables > 0) return 'table';
        if (counts.lines >= 3 || counts.charts > 0) return 'diagram';
        if (counts.pics >= 3 || counts.shapes >= 4) return 'cards';
        return baseCat;
      })();
      const hasTitle = phs.some(function (p) { return p.type === 'title' || p.type === 'ctrTitle'; });
      const bodyCount = phs.filter(function (p) { return p.type === 'body' || p.type === '' || p.type === 'subTitle'; }).length;
      const entry = {
        index: i + 1,
        layoutId: layoutEntry ? layoutEntry.id : null,
        layoutName: layoutEntry ? layoutEntry.name : '',
        category: category,
        suggestedType: CATEGORY_TO_MAJIN[category] || 'content',
        elements: { placeholders: phs.length, bodies: bodyCount, shapes: counts.shapes, pics: counts.pics, lines: counts.lines, tables: counts.tables, charts: counts.charts },
        hasTitle: hasTitle
      };
      exampleSlides.push(entry);
      if (layoutEntry) layoutEntry.usedBy.push(i + 1);
    }

    // --- カラー役割 + 9論理名対応 ---
    const colorRoles = [];
    Object.keys(colors).forEach(function (key) {
      let role = '';
      if (key === 'dk1' || key === 'dk2' || key === 'tx1' || key === 'tx2') role = '文字/濃色';
      else if (key === 'lt1' || key === 'lt2' || key === 'bg1' || key === 'bg2') role = '背景/淡色';
      else if (key === 'accent1' || key === 'accent2') role = 'メインカラー候補';
      else if (/^accent/.test(key)) role = 'アクセント候補';
      else if (/hlink/i.test(key)) role = 'リンク';
      colorRoles.push({ key: key, hex: colors[key], role: role, palette: nearestPaletteName(colors[key]) });
    });

    // --- 重要原型のマーキング（座標まで詳細にmd記載する対象） ---
    const importantOrder = ['title', 'section', 'content', 'twoCol', 'compare', 'closing'];
    const seen = {};
    layoutCatalog.forEach(function (e) { e.important = false; });
    importantOrder.forEach(function (cat) {
      if (seen[cat]) return;
      // 例スライドで使われているものを優先
      const used = layoutCatalog.filter(function (e) { return e.category === cat && e.usedBy.length; });
      const cand = (used[0]) || layoutCatalog.find(function (e) { return e.category === cat; });
      if (cand) { cand.important = true; seen[cat] = true; }
    });

    return {
      meta: { slideSize: slideSize, format: format, layoutCount: layoutCatalog.length, slideCount: exampleSlides.length },
      colors: colors,
      colorRoles: colorRoles,
      fonts: Object.keys(fontSet),
      majorFont: themeInfo.majorFont,
      minorFont: themeInfo.minorFont,
      backgroundColor: backgroundColor,
      decorations: masterObjects,
      layoutCatalog: layoutCatalog,
      exampleSlides: exampleSlides
    };
  }

  // ===== Markdown 生成 =====
  function fmtFrame(f) {
    if (!f) return '—';
    function r(n) { return Math.round(n * 100) / 100; }
    return 'left ' + r(f.left) + ', top ' + r(f.top) + ', w ' + r(f.width) + ', h ' + r(f.height) + ' (in)';
  }
  function buildTemplateAnalysisMarkdown(an) {
    const L = [];
    L.push('# template_analysis.md');
    L.push('');
    L.push('> このファイルは、テンプレートPPTXを定型解析して書き出した「設計情報」です。');
    L.push('> まじん式+.md と元データに**この情報を併せて**与え、テンプレの構成・配色・レイアウトを活かした slideData を生成してください。');
    L.push('> ※テンプレ例スライドの本文・図表の中身は含みません（構成・型・配置のみ）。');
    L.push('');

    // 1. メタ
    L.push('## 1. メタ情報');
    const sz = an.meta.slideSize;
    L.push('- スライドサイズ: ' + (sz ? (Math.round(sz.wIn * 100) / 100 + ' × ' + Math.round(sz.hIn * 100) / 100 + ' in') : '不明'));
    L.push('- 形式: ' + an.meta.format + (an.meta.format === 'unsupported' ? '（16:9 / A4横 のみ対応・対象外）' : ''));
    L.push('- レイアウト数: ' + an.meta.layoutCount + ' / 例スライド数: ' + an.meta.slideCount);
    L.push('');

    // 2. カラー
    L.push('## 2. カラーパレット（テンプレ色 → 9論理名 → 役割）');
    L.push('');
    L.push('| theme | HEX | 最近傍の論理名 | 役割 |');
    L.push('| --- | --- | --- | --- |');
    an.colorRoles.forEach(function (c) { L.push('| ' + c.key + ' | ' + c.hex + ' | ' + (c.palette || '—') + ' | ' + (c.role || '') + ' |'); });
    L.push('');
    L.push('> slideData の `accentColor` は上表の**論理名**で指定。メインカラーは accent1またはaccent2（または役割「メインカラー候補」）を基調に統一してください。');
    L.push('');

    // 3. フォント
    L.push('## 3. フォント');
    L.push('- 見出し(major): ' + (an.majorFont || '—'));
    L.push('- 本文(minor): ' + (an.minorFont || '—'));
    if (an.fonts && an.fonts.length) L.push('- テンプレに現れる全フォント: ' + an.fonts.join(' / '));
    L.push('');

    // 4. レイアウトカタログ
    L.push('## 4. レイアウトカタログ（まじん式タイプへの対応）');
    L.push('');
    L.push('| id | 分類 | 推奨まじん式タイプ | 例スライドでの使用 | タイトル枠 | 本文枠 | 図枠 | 列 |');
    L.push('| --- | --- | --- | --- | --- | --- | --- | --- |');
    an.layoutCatalog.forEach(function (e) {
      const figN = (e.figureFrames && e.figureFrames.length) ? (e.figureFrames.length + '枠') : '—';
      L.push('| ' + e.id + ' | ' + e.category + ' | ' + e.suggestedType + ' | ' + (e.usedBy.length ? ('slide ' + e.usedBy.join(',')) : '—') + ' | ' + (e.titleFrame ? 'あり' : '—') + ' | ' + (e.bodyFrames.length ? e.bodyFrames.length + '枠' : '—') + ' | ' + figN + ' | ' + e.columns + ' |');
    });
    L.push('');
    // 重要原型は座標・許容量を詳細に
    L.push('### 重要原型の詳細（幾何座標・テキスト許容量）');
    L.push('');
    an.layoutCatalog.filter(function (e) { return e.important; }).forEach(function (e) {
      L.push('#### ' + e.id + '（' + e.category + ' / ' + e.suggestedType + '）' + (e.name ? ' — ' + e.name : ''));
      L.push('- タイトル枠: ' + fmtFrame(e.titleFrame));
      e.bodyFrames.forEach(function (bf, i) { L.push('- 本文枠' + (i + 1) + ': ' + fmtFrame(bf)); });
      (e.figureFrames || []).forEach(function (ff, i) { L.push('- 図枠' + (i + 1) + '(' + (ff.phType || 'pic') + '): ' + fmtFrame(ff)); });
      const cap = e.capacity || {};
      const capParts = [];
      if (cap.titleChars) capParts.push('タイトル≈' + cap.titleChars + '字');
      if (cap.bodyLines) capParts.push('本文≈' + cap.bodyCharsPerLine + '字×' + cap.bodyLines + '行');
      L.push('- テキスト許容量(目安): ' + (capParts.length ? capParts.join(' / ') : '—'));
      L.push('- 装飾図形数: ' + (e.objects ? e.objects.length : 0) + (e.background ? '（背景あり）' : ''));
      L.push('');
    });

    // 5. 例スライド構成
    L.push('## 5. 例スライド構成サマリ（中身なし・構成のみ）');
    L.push('');
    L.push('| # | 参照レイアウト | 分類 | 推奨まじん式タイプ | 要素(本文枠/図形/画像/線/表) |');
    L.push('| --- | --- | --- | --- | --- |');
    an.exampleSlides.forEach(function (s) {
      const el = s.elements;
      L.push('| ' + s.index + ' | ' + (s.layoutId || '—') + ' | ' + s.category + ' | ' + s.suggestedType + ' | ' + el.bodies + '/' + el.shapes + '/' + el.pics + '/' + el.lines + '/' + el.tables + ' |');
    });
    L.push('');

    // 6. LLMへの指示
    L.push('## 6. slideData 生成時の指示（このテンプレに合わせる）');
    L.push('');
    L.push('- **テンプレが持つ型を優先**: 上のカタログ「推奨まじん式タイプ」と例スライドの並び・粒度に倣って slideData の type を選ぶ。');
    L.push('- **配色**: `accentColor`/`color` は §2 の論理名で指定。メインカラー1色を全 `section`/`title`/`closing`/章末 `iconBanner` に統一（まじん式+ §6.2）。');
    L.push('- **フォント**: 文字数は §4 の「テキスト許容量(目安)」を超えない。タイトル/本文の字数上限を厳守。');
    L.push('- **レイアウト指定（任意）**: 特定のテンプレ原型に強く寄せたいスライドは、共通プロパティ `templateLayout` にカタログの id（例 "' + (an.layoutCatalog[0] ? an.layoutCatalog[0].id : 'L1') + '"）を指定してよい。**座標は書かない**（描画ツールがテンプレ実枠へ配置する）。');
    (function () {
      const figLayouts = an.layoutCatalog.filter(function (e) { return e.figureFrames && e.figureFrames.length; }).map(function (e) { return e.id; });
      L.push('- **図の差し込み**: 図/写真/図表を入れたいスライドは共通プロパティ `figure` に「何を入れるか」を記述（ツールがダミー枠を描画）。'
        + (figLayouts.length ? '図枠ありレイアウト: ' + figLayouts.join(', ') + ' を `templateLayout` に指定するとテンプレ図位置に入る。' : '図枠ありレイアウトは無いため、本文枠の右側に図ダミーが入る。')
        + ' **画像URLは生成・推定しない**。');
    })();
    L.push('- **例スライドの本文・図表は引用しない**（構成だけを参考にする）。');
    L.push('');
    return L.join('\n');
  }

  window.analyzeTemplate = analyzeTemplate;
  window.buildTemplateAnalysisMarkdown = buildTemplateAnalysisMarkdown;
})();
