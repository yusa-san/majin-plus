/**
 * PPTX テンプレート・スタイル抽出
 *
 * テンプレートとなる .pptx を解凍し、テーマ(theme1.xml)・スライドマスター
 * (slideMaster1.xml)・レイアウト(slideLayout*.xml)から
 * 「カラースキーム / フォント / マスター背景色 / 静的要素・プレースホルダ座標」を読み取る。
 *
 * 依存: JSZip (CDN), DOMParser (ブラウザ標準)
 * 公開API: window.loadPptxTemplate(file) -> Promise<ThemeInfo>
 *
 * ThemeInfo = {
 *   colors: { dk1, lt1, dk2, lt2, accent1..accent6, hlink, folHlink },  // '#RRGGBB'
 *   majorFont, minorFont,         // フォント名（日本語フォント優先）
 *   backgroundColor,              // マスター背景色 '#RRGGBB' or null
 *   masters: {                    // レイアウト取り込み用（無ければ null）
 *     title:   MasterDef,         // 表紙系（Title Slide レイアウト）
 *     content: MasterDef          // 本文系（Title and Content レイアウト）
 *   }
 * }
 * MasterDef = {
 *   background: {color}|{data}|null,        // PptxGenJS defineSlideMaster.background 互換
 *   objects: [ {rect|line|image|text:{...}} ],  // 同 objects 互換（座標はinch）
 *   placeholders: { title:RectIn|null, body:RectIn|null }  // RectIn={left,top,width,height}(inch)
 * }
 */
(function () {
  'use strict';

  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

  function parseXml(str) {
    return new DOMParser().parseFromString(str, 'application/xml');
  }

  // 要素配下の最初の色定義(srgbClr/sysClr)を '#RRGGBB' で返す
  function colorOf(el) {
    if (!el) return null;
    const srgb = el.getElementsByTagName('a:srgbClr')[0] || el.getElementsByTagNameNS(A, 'srgbClr')[0];
    if (srgb && srgb.getAttribute('val')) return '#' + srgb.getAttribute('val').toUpperCase();
    const sys = el.getElementsByTagName('a:sysClr')[0] || el.getElementsByTagNameNS(A, 'sysClr')[0];
    if (sys) return '#' + (sys.getAttribute('lastClr') || '000000').toUpperCase();
    return null;
  }

  // 要素ノードの子のみを配列で返す（.children 非対応環境にも対応）
  function elementChildren(el) {
    const out = [];
    if (!el || !el.childNodes) return out;
    for (let i = 0; i < el.childNodes.length; i++) {
      const c = el.childNodes[i];
      if (c.nodeType === 1) out.push(c);
    }
    return out;
  }

  function getByTag(parent, tag) {
    let els = parent.getElementsByTagName(tag);
    if (els.length) return els[0];
    const local = tag.indexOf(':') >= 0 ? tag.split(':')[1] : tag;
    els = parent.getElementsByTagNameNS('*', local);
    return els.length ? els[0] : null;
  }

  function parseTheme(themeDoc) {
    const colors = {};
    const clrScheme = getByTag(themeDoc, 'a:clrScheme');
    if (clrScheme) {
      // 子要素 a:dk1, a:lt1, a:dk2, a:lt2, a:accent1.. を走査
      elementChildren(clrScheme).forEach(function (child) {
        const name = child.localName || (child.tagName || '').replace(/^a:/, ''); // dk1, lt1, accent1 ...
        const c = colorOf(child);
        if (c) colors[name] = c;
      });
    }

    // typeface 文字列が実フォント名か判定（空 / +mn-* +mj-* 等のテーマ参照トークンを除外）
    function isRealTypeface(name) {
      const n = (name || '').trim();
      return !!n && n.charAt(0) !== '+';
    }
    // scheme(majorFont/minorFont) 直下の <a:font script="Jpan"> 等から typeface を取得
    function scriptFontTypeface(scheme, script) {
      const fonts = scheme.getElementsByTagName('a:font');
      const list = (fonts && fonts.length) ? fonts : scheme.getElementsByTagNameNS(A, 'font');
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

    // フォント（本文=minorFont を採用）。和文は ea(非空) → <a:font script="Jpan"> → 任意のscriptフォント → latin の順で採用。
    function fontFromScheme(tag) {
      const scheme = getByTag(themeDoc, tag);
      if (!scheme) return '';
      return tfOf(scheme, 'ea')
        || scriptFontTypeface(scheme, 'Jpan')
        || scriptFontTypeface(scheme, null)
        || tfOf(scheme, 'latin')
        || '';
    }
    // scheme に現れる全 typeface（latin/ea/cs + 全 a:font）を収集
    function collectFontsFromScheme(tag, out) {
      const scheme = getByTag(themeDoc, tag);
      if (!scheme) return;
      ['latin', 'ea', 'cs'].forEach(function (t) { const v = tfOf(scheme, t); if (v) out[v] = true; });
      const fonts = scheme.getElementsByTagName('a:font');
      const list = (fonts && fonts.length) ? fonts : scheme.getElementsByTagNameNS(A, 'font');
      for (let i = 0; i < list.length; i++) {
        const v = list[i].getAttribute('typeface');
        if (isRealTypeface(v)) out[v.trim()] = true;
      }
    }
    const majorFont = fontFromScheme('a:majorFont');
    const minorFont = fontFromScheme('a:minorFont');
    const fontSet = {};
    collectFontsFromScheme('a:majorFont', fontSet);
    collectFontsFromScheme('a:minorFont', fontSet);

    return { colors, majorFont, minorFont, fonts: Object.keys(fontSet) };
  }

  // schemeClr の val(bg1/tx1/accent1/dk1...) を実色に解決
  function resolveSchemeColor(val, colors, clrMap) {
    if (!val) return null;
    // clrMap 経由（bg1 -> lt1 等）
    if (clrMap && clrMap[val]) val = clrMap[val];
    // phClr 等は解決不能
    if (colors[val]) return colors[val];
    return null;
  }

  function parseMasterBackground(masterDoc, colors) {
    // clrMap（属性: bg1, tx1, bg2, tx2, accent1..6, hlink, folHlink）
    const clrMap = {};
    const clrMapEl = getByTag(masterDoc, 'p:clrMap');
    if (clrMapEl) {
      ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].forEach(k => {
        const v = clrMapEl.getAttribute(k);
        if (v) clrMap[k] = v;
      });
    }

    const bg = getByTag(masterDoc, 'p:bg');
    if (!bg) return null;

    // 1) bgPr > solidFill（srgbClr or schemeClr）
    const bgPr = bg.getElementsByTagName('p:bgPr')[0] || bg.getElementsByTagNameNS('*', 'bgPr')[0];
    if (bgPr) {
      const solid = bgPr.getElementsByTagName('a:solidFill')[0] || bgPr.getElementsByTagNameNS(A, 'solidFill')[0];
      if (solid) {
        const srgb = solid.getElementsByTagName('a:srgbClr')[0] || solid.getElementsByTagNameNS(A, 'srgbClr')[0];
        if (srgb && srgb.getAttribute('val')) return '#' + srgb.getAttribute('val').toUpperCase();
        const sch = solid.getElementsByTagName('a:schemeClr')[0] || solid.getElementsByTagNameNS(A, 'schemeClr')[0];
        if (sch) {
          const r = resolveSchemeColor(sch.getAttribute('val'), colors, clrMap);
          if (r) return r;
        }
      }
    }

    // 2) bgRef > schemeClr（塗りスタイル参照。色は近似でschemeClrを採用）
    const bgRef = bg.getElementsByTagName('p:bgRef')[0] || bg.getElementsByTagNameNS('*', 'bgRef')[0];
    if (bgRef) {
      const sch = bgRef.getElementsByTagName('a:schemeClr')[0] || bgRef.getElementsByTagNameNS(A, 'schemeClr')[0];
      if (sch) {
        const r = resolveSchemeColor(sch.getAttribute('val'), colors, clrMap);
        if (r) return r;
      }
      const srgb = bgRef.getElementsByTagName('a:srgbClr')[0] || bgRef.getElementsByTagNameNS(A, 'srgbClr')[0];
      if (srgb && srgb.getAttribute('val')) return '#' + srgb.getAttribute('val').toUpperCase();
    }

    return null;
  }

  // ========================================
  // マスター/レイアウト（ジオメトリ）抽出ヘルパー
  // ========================================
  const EMU_PER_IN = 914400;
  const EMU_PER_PT = 12700;
  function emuToIn(v) { return Number(v || 0) / EMU_PER_IN; }
  function stripHash(c) { return c ? String(c).replace('#', '') : c; }

  // 直下の子要素から localName 一致の最初を返す
  function directChild(el, local) {
    if (!el) return null;
    const kids = elementChildren(el);
    for (let i = 0; i < kids.length; i++) {
      if ((kids[i].localName || (kids[i].tagName || '').replace(/^[a-z]+:/, '')) === local) return kids[i];
    }
    return null;
  }

  function clrMapFrom(masterDoc) {
    const clrMap = {};
    const clrMapEl = getByTag(masterDoc, 'p:clrMap');
    if (clrMapEl) {
      ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].forEach(k => {
        const v = clrMapEl.getAttribute(k);
        if (v) clrMap[k] = v;
      });
    }
    return clrMap;
  }

  // solidFill / schemeClr / srgbClr を含むコンテナ(直下)から色を解決
  function colorFromContainer(container, colors, clrMap) {
    if (!container) return null;
    const srgb = directChild(container, 'srgbClr');
    if (srgb && srgb.getAttribute('val')) return '#' + srgb.getAttribute('val').toUpperCase();
    const sch = directChild(container, 'schemeClr');
    if (sch) return resolveSchemeColor(sch.getAttribute('val'), colors, clrMap);
    return null;
  }

  // spPr から塗り色を解決（無ければ style.fillRef を参照）。'#RRGGBB' / null(明示noFill) / undefined(不明)
  function fillFrom(sp, spPr, colors, clrMap) {
    if (spPr) {
      if (directChild(spPr, 'noFill')) return null;
      const solid = directChild(spPr, 'solidFill');
      if (solid) { const c = colorFromContainer(solid, colors, clrMap); if (c) return c; }
      const grad = directChild(spPr, 'gradFill');
      if (grad) {
        const gsLst = directChild(grad, 'gsLst');
        const gs = gsLst ? directChild(gsLst, 'gs') : null;
        if (gs) { const c = colorFromContainer(gs, colors, clrMap); if (c) return c; }
      }
    }
    // style > fillRef
    const style = directChild(sp, 'style');
    if (style) {
      const fillRef = directChild(style, 'fillRef');
      if (fillRef) { const c = colorFromContainer(fillRef, colors, clrMap); if (c) return c; }
    }
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
    const width = wEmu ? Number(wEmu) / EMU_PER_PT : 1;
    return { color: stripHash(color), width: width };
  }

  function xfrmFrom(spPr) {
    const xfrm = spPr ? directChild(spPr, 'xfrm') : null;
    if (!xfrm) return null;
    const off = directChild(xfrm, 'off'), ext = directChild(xfrm, 'ext');
    if (!off || !ext) return null;
    const r = {
      x: emuToIn(off.getAttribute('x')), y: emuToIn(off.getAttribute('y')),
      w: emuToIn(ext.getAttribute('cx')), h: emuToIn(ext.getAttribute('cy'))
    };
    const rot = xfrm.getAttribute('rot'); if (rot) r.rotate = Number(rot) / 60000;
    if (xfrm.getAttribute('flipH') === '1') r.flipH = true;
    if (xfrm.getAttribute('flipV') === '1') r.flipV = true;
    return r;
  }

  // 非プレースホルダ sp 内の txBody から簡易テキストを取得
  function textFrom(sp, colors, clrMap) {
    const txBody = directChild(sp, 'txBody');
    if (!txBody) return null;
    const paras = elementChildren(txBody).filter(e => (e.localName || '') === 'p');
    let str = '';
    let size = null, color = null, bold = false;
    paras.forEach((p, pi) => {
      if (pi > 0) str += '\n';
      elementChildren(p).forEach(node => {
        if ((node.localName || '') === 'r') {
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
    if (!str) return null;
    return { string: str, size: size, color: color, bold: bold };
  }

  function placeholderInfo(sp) {
    const nvSpPr = directChild(sp, 'nvSpPr');
    const nvPr = nvSpPr ? directChild(nvSpPr, 'nvPr') : null;
    const ph = nvPr ? directChild(nvPr, 'ph') : null;
    if (!ph) return null;
    return { type: ph.getAttribute('type') || '', idx: ph.getAttribute('idx') || '' };
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
    parts.forEach(p => { if (p === '' || p === '.') return; if (p === '..') stack.pop(); else stack.push(p); });
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
      const id = list[i].getAttribute('Id');
      const target = list[i].getAttribute('Target');
      if (id && target) map[id] = target;
    }
    return map;
  }

  // 背景塗りを {color} / {data} / null で返す
  async function parseBgFill(doc, baseDir, zip, relsMap, colors, clrMap) {
    const bg = getByTag(doc, 'p:bg');
    if (!bg) return null;
    const bgPr = directChild(bg, 'bgPr');
    if (bgPr) {
      const solid = directChild(bgPr, 'solidFill');
      if (solid) { const c = colorFromContainer(solid, colors, clrMap); if (c) return { color: stripHash(c) }; }
      const blip = (function () {
        const bf = directChild(bgPr, 'blipFill');
        return bf ? directChild(bf, 'blip') : null;
      })();
      if (blip) {
        const rid = blip.getAttribute('r:embed') || blip.getAttribute('embed');
        const target = rid && relsMap ? relsMap[rid] : null;
        if (target) {
          const path = resolvePath(baseDir, target);
          const file = path && zip.file(path);
          if (file) { const b64 = await file.async('base64'); return { data: 'data:' + mimeFromExt(path) + ';base64,' + b64 }; }
        }
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
      const node = kids[i];
      const local = node.localName || (node.tagName || '').replace(/^[a-z]+:/, '');
      if (local === 'sp') {
        if (placeholderInfo(node)) continue; // プレースホルダは静的要素から除外
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
        // 図形内の静的テキスト
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
        const blipFill = directChild(node, 'blipFill');
        const blip = blipFill ? directChild(blipFill, 'blip') : null;
        const rid = blip ? (blip.getAttribute('r:embed') || blip.getAttribute('embed')) : null;
        const target = rid && relsMap ? relsMap[rid] : null;
        if (!target) continue;
        const path = resolvePath(baseDir, target);
        const file = path && zip.file(path);
        if (!file) continue;
        try {
          const b64 = await file.async('base64');
          objects.push({ image: { x: r.x, y: r.y, w: r.w, h: r.h, data: 'data:' + mimeFromExt(path) + ';base64,' + b64 } });
        } catch (e) { /* skip */ }
      }
      // grpSp 等は近似対象外（スキップ）
    }
    return objects;
  }

  // doc のプレースホルダを type/idx -> rect(inch) のマップで返す
  function placeholdersOf(doc) {
    const result = [];
    const spTree = getByTag(doc, 'p:spTree');
    if (!spTree) return result;
    elementChildren(spTree).forEach(node => {
      if ((node.localName || '') !== 'sp') return;
      const ph = placeholderInfo(node);
      if (!ph) return;
      const spPr = directChild(node, 'spPr');
      const r = xfrmFrom(spPr);
      result.push({ type: ph.type, idx: ph.idx, rect: r });
    });
    return result;
  }

  // レイアウトのプレースホルダ(無xfrmはマスターから継承)から title/body rect(engine形式)を返す
  function placeholderRects(layoutPhs, masterPhs) {
    function findMaster(type, idx) {
      let m = masterPhs.find(p => p.type === type && p.idx === idx && p.rect);
      if (m) return m.rect;
      m = masterPhs.find(p => p.type === type && p.rect);
      return m ? m.rect : null;
    }
    function rectFor(matchTypes) {
      const cand = layoutPhs.find(p => matchTypes.indexOf(p.type) >= 0) ||
        (matchTypes.indexOf('') >= 0 ? layoutPhs.find(p => p.type === '') : null);
      if (!cand) return null;
      const r = cand.rect || findMaster(cand.type, cand.idx);
      if (!r) return null;
      return { left: r.x, top: r.y, width: r.w, height: r.h };
    }
    return {
      title: rectFor(['ctrTitle', 'title']),
      body: rectFor(['body', 'subTitle', ''])
    };
  }

  function readEntry(zip, candidates) {
    for (const name of candidates) {
      const f = zip.file(name);
      if (f) return f.async('string');
    }
    // フォールバック: パターン一致の先頭
    return null;
  }

  // doc 内の全 typeface（a:latin/a:ea/a:cs/a:font）を out{name:true} に収集（テーマ参照トークン/空は除外）
  function collectTypefacesFromDoc(doc, out) {
    if (!doc) return;
    ['a:latin', 'a:ea', 'a:cs', 'a:font'].forEach(function (tag) {
      let list = doc.getElementsByTagName(tag);
      if (!list || !list.length) list = doc.getElementsByTagNameNS(A, tag.split(':')[1]);
      for (let i = 0; i < list.length; i++) {
        const t = list[i].getAttribute('typeface');
        const n = (t || '').trim();
        if (n && n.charAt(0) !== '+') out[n] = true;
      }
    });
  }

  async function loadPptxTemplate(file) {
    if (typeof JSZip === 'undefined') {
      throw new Error('JSZip が読み込まれていません（CDN取得失敗の可能性）。');
    }
    const zip = await JSZip.loadAsync(file);

    // スライドサイズ / 形式判定（16:9 / A4横 / 対象外）
    // presentation.xml の <p:sldSz cx cy>(EMU) を読み、アスペクト比で分類する。
    let slideSize = null;
    let format = '16:9'; // sldSz が読めない場合のフォールバック
    try {
      const presFile = zip.file('ppt/presentation.xml');
      if (presFile) {
        const sldSz = getByTag(parseXml(await presFile.async('string')), 'p:sldSz');
        if (sldSz) {
          const cx = Number(sldSz.getAttribute('cx') || 0);
          const cy = Number(sldSz.getAttribute('cy') || 0);
          if (cx > 0 && cy > 0) {
            const wIn = cx / EMU_PER_IN, hIn = cy / EMU_PER_IN;
            slideSize = { wIn: wIn, hIn: hIn };
            const aspect = wIn / hIn;
            const tol = 0.04;
            if (Math.abs(aspect - (16 / 9)) <= tol) format = '16:9';        // 10×5.625 / 13.333×7.5 等
            else if (Math.abs(aspect - (10.83 / 7.5)) <= tol) format = 'a4'; // PowerPoint A4横 / 真A4横
            else format = 'unsupported';
          }
        }
      }
    } catch (e) { console.warn('slide size detect failed', e); }

    // theme1.xml（無ければ theme フォルダの先頭）
    let themeStr = await (readEntry(zip, ['ppt/theme/theme1.xml']) || Promise.resolve(null));
    if (!themeStr) {
      const themeFiles = zip.file(/^ppt\/theme\/theme\d+\.xml$/);
      if (themeFiles && themeFiles.length) themeStr = await themeFiles[0].async('string');
    }
    if (!themeStr) throw new Error('テーマ(theme1.xml)が見つかりません。有効な.pptxファイルか確認してください。');
    const themeInfo = parseTheme(parseXml(themeStr));

    // テンプレ由来フォントの集合（テーマ fontScheme をシード。後でマスター/レイアウトもマージ）
    const fontSet = {};
    (themeInfo.fonts || []).forEach(function (f) { fontSet[f] = true; });

    // スライドマスター（背景色・静的要素・プレースホルダ用）
    let masterPath = 'ppt/slideMasters/slideMaster1.xml';
    if (!zip.file(masterPath)) {
      const mf = zip.file(/^ppt\/slideMasters\/slideMaster\d+\.xml$/);
      masterPath = (mf && mf.length) ? mf[0].name : null;
    }

    let clrMap = {}, masterObjects = [], masterPhs = [], masterBg = null, backgroundColor = null, masters = null;

    if (masterPath) {
      const masterDoc = parseXml(await zip.file(masterPath).async('string'));
      collectTypefacesFromDoc(masterDoc, fontSet); // run/プレースホルダ直接指定フォントも拾う
      clrMap = clrMapFrom(masterDoc);
      const masterBaseDir = masterPath.substring(0, masterPath.lastIndexOf('/'));
      const masterFile = masterPath.substring(masterPath.lastIndexOf('/') + 1);
      const masterRels = await parseRels(zip, masterBaseDir + '/_rels/' + masterFile + '.rels');
      try { masterObjects = await buildStaticObjects(masterDoc, masterBaseDir, zip, masterRels, themeInfo.colors, clrMap); }
      catch (e) { console.warn('master objects parse failed', e); }
      masterPhs = placeholdersOf(masterDoc);
      try { masterBg = await parseBgFill(masterDoc, masterBaseDir, zip, masterRels, themeInfo.colors, clrMap); }
      catch (e) { console.warn('master bg parse failed', e); }
      backgroundColor = (masterBg && masterBg.color) ? ('#' + masterBg.color) : null;

      // レイアウト列挙・選別
      const layoutFiles = zip.file(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/) || [];
      const layouts = [];
      for (const lf of layoutFiles) {
        const doc = parseXml(await lf.async('string'));
        const type = (doc.documentElement && doc.documentElement.getAttribute('type')) || '';
        const cSld = getByTag(doc, 'p:cSld');
        const name = cSld ? (cSld.getAttribute('name') || '') : '';
        collectTypefacesFromDoc(doc, fontSet); // レイアウト直接指定フォントも拾う
        layouts.push({ path: lf.name, doc: doc, type: type, name: name });
      }
      function pickLayout(typeList, nameRe) {
        let l = layouts.find(x => typeList.indexOf(x.type) >= 0);
        if (l) return l;
        if (nameRe) { l = layouts.find(x => nameRe.test(x.name)); if (l) return l; }
        return null;
      }
      let titleLayout = pickLayout(['title'], /title\s*slide|表紙/i);
      let contentLayout = pickLayout(['obj', 'tx', 'twoObj', 'secHead'], /title and content|content|本文/i);
      if (!contentLayout) contentLayout = layouts.find(x => ['title', 'blank'].indexOf(x.type) < 0) || null;
      if (!titleLayout) titleLayout = contentLayout;

      async function buildMasterDef(layout) {
        let objects = masterObjects.slice();
        let bg = masterBg;
        let placeholders = placeholderRects(masterPhs, masterPhs);
        if (layout) {
          const baseDir = layout.path.substring(0, layout.path.lastIndexOf('/'));
          const file = layout.path.substring(layout.path.lastIndexOf('/') + 1);
          const rels = await parseRels(zip, baseDir + '/_rels/' + file + '.rels');
          let lObjs = [];
          try { lObjs = await buildStaticObjects(layout.doc, baseDir, zip, rels, themeInfo.colors, clrMap); }
          catch (e) { console.warn('layout objects parse failed', e); }
          objects = masterObjects.concat(lObjs);
          const lbg = await parseBgFill(layout.doc, baseDir, zip, rels, themeInfo.colors, clrMap);
          if (lbg) bg = lbg;
          placeholders = placeholderRects(placeholdersOf(layout.doc), masterPhs);
        }
        return { background: bg, objects: objects, placeholders: placeholders };
      }

      masters = {
        title: await buildMasterDef(titleLayout),
        content: await buildMasterDef(contentLayout)
      };
    }

    return {
      colors: themeInfo.colors,
      majorFont: themeInfo.majorFont,
      minorFont: themeInfo.minorFont,
      backgroundColor: backgroundColor,
      masters: masters,
      slideSize: slideSize,   // {wIn,hIn} or null（生のテンプレ寸法。正規化はエンジン側）
      format: format,         // '16:9' | 'a4' | 'unsupported'
      fonts: Object.keys(fontSet) // テーマ＋マスター/レイアウト由来の全フォント名（重複排除）
    };
  }

  window.loadPptxTemplate = loadPptxTemplate;
})();
