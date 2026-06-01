/**
 * PPTX Slide Generator Engine (local, browser)
 *
 * GAS版「Google Slide Generator」(コード.gs)の描画ロジックを、
 * CDN配信の PptxGenJS を用いてローカルでPPTXを生成するよう移植したもの。
 *
 * - 座標基準: 960x540px (元コードと同じ) を 10in x 5.625in (16:9) にマッピング。
 * - フォントサイズは pt のまま。
 * - 画像(ロゴ・背景)はURLを直挿し。取得失敗時は当該画像のみスキップ。
 *
 * 公開API: window.generatePptx(slideDataString, settings) -> Promise (writeFile)
 *
 * @author まじん (original GAS) yusa-san (ported to PptxGenJS)
 * @license CC BY-NC 4.0
 */
(function () {
  'use strict';

  // 現在生成中の PptxGenJS インスタンス（ShapeType参照用）
  let PPTX = null;

  // ========================================
  // 色彩操作ヘルパー関数（元コードからそのまま移植）
  // ========================================
  function hexToRgb(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
      r: parseInt(result[1], 16),
      g: parseInt(result[2], 16),
      b: parseInt(result[3], 16)
    } : null;
  }

  function rgbToHsl(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h, s, l = (max + min) / 2;
    if (max === min) { h = s = 0; }
    else {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      switch (max) {
        case r: h = (g - b) / d + (g < b ? 6 : 0); break;
        case g: h = (b - r) / d + 2; break;
        case b: h = (r - g) / d + 4; break;
      }
      h /= 6;
    }
    return { h: h * 360, s: s * 100, l: l * 100 };
  }

  function hslToHex(h, s, l) {
    s /= 100; l /= 100;
    const c = (1 - Math.abs(2 * l - 1)) * s,
      x = c * (1 - Math.abs((h / 60) % 2 - 1)),
      m = l - c / 2;
    let r = 0, g = 0, b = 0;
    if (0 <= h && h < 60) { r = c; g = x; b = 0; }
    else if (60 <= h && h < 120) { r = x; g = c; b = 0; }
    else if (120 <= h && h < 180) { r = 0; g = c; b = x; }
    else if (180 <= h && h < 240) { r = 0; g = x; b = c; }
    else if (240 <= h && h < 300) { r = x; g = 0; b = c; }
    else if (300 <= h && h < 360) { r = c; g = 0; b = x; }
    r = Math.round((r + m) * 255);
    g = Math.round((g + m) * 255);
    b = Math.round((b + m) * 255);
    return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
  }

  function generateTintedGray(tintColorHex, saturation, lightness) {
    const rgb = hexToRgb(tintColorHex);
    if (!rgb) return '#F8F9FA';
    const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
    return hslToHex(hsl.h, saturation, lightness);
  }

  function lightenColor(color, amount) {
    const rgb = hexToRgb(color);
    if (!rgb) return color;
    const lighten = (c) => Math.min(255, Math.round(c + (255 - c) * amount));
    const newR = lighten(rgb.r), newG = lighten(rgb.g), newB = lighten(rgb.b);
    return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
  }

  function darkenColor(color, amount) {
    const rgb = hexToRgb(color);
    if (!rgb) return color;
    const darken = (c) => Math.max(0, Math.round(c * (1 - amount)));
    const newR = darken(rgb.r), newG = darken(rgb.g), newB = darken(rgb.b);
    return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
  }

  function generatePyramidColors(baseColor, levels) {
    const colors = [];
    for (let i = 0; i < levels; i++) {
      const lightenAmount = (i / Math.max(1, levels - 1)) * 0.6;
      colors.push(lightenColor(baseColor, lightenAmount));
    }
    return colors;
  }

  function generateStepUpColors(baseColor, steps) {
    const colors = [];
    for (let i = 0; i < steps; i++) {
      const lightenAmount = 0.6 * (1 - (i / Math.max(1, steps - 1)));
      colors.push(lightenColor(baseColor, lightenAmount));
    }
    return colors;
  }

  function generateProcessColors(baseColor, steps) {
    const colors = [];
    for (let i = 0; i < steps; i++) {
      const lightenAmount = 0.5 * (1 - (i / Math.max(1, steps - 1)));
      colors.push(lightenColor(baseColor, lightenAmount));
    }
    return colors;
  }

  function generateTimelineCardColors(baseColor, milestones) {
    const colors = [];
    for (let i = 0; i < milestones; i++) {
      const lightenAmount = 0.4 * (1 - (i / Math.max(1, milestones - 1)));
      colors.push(lightenColor(baseColor, lightenAmount));
    }
    return colors;
  }

  function generateCompareColors(baseColor) {
    return { left: darkenColor(baseColor, 0.3), right: baseColor };
  }

  // ========================================
  // まじん式+ カラーパレット（accentColor / color の9値・main のみ。light/dark は派生）
  // ========================================
  const PALETTE = {
    skyblue:   '#42A5F5',
    green:     '#4CAF50',
    deepgreen: '#2E7D32',
    yellow:    '#FFC107',
    blue:      '#1565C0',
    orange:    '#FF7043',
    red:       '#F44336',
    white:     '#FFFFFF', // ベースカラー
    black:     '#000000'  // ベースカラー
  };
  // gradient モードで濃淡を割り当てるクロマ系の固定順（white/black 除く）
  const PALETTE_CHROMA_ORDER = ['skyblue', 'blue', 'green', 'deepgreen', 'yellow', 'orange', 'red'];

  // main 色から light/dark を派生（white/black は専用値）
  function deriveLight(key, mainHex) {
    if (key === 'white') return '#FFFFFF';
    if (key === 'black') return '#F1F3F4';
    return generateTintedGray(mainHex, 30, 94);
  }
  function deriveDark(key, mainHex) {
    if (key === 'white') return '#333333';
    if (key === 'black') return '#000000';
    return darkenColor(mainHex, 0.25);
  }

  // gradient モード: 基色(primary)の濃淡シェードを名前のインデックスから生成
  function gradientShade(key) {
    const base = CONFIG.COLORS.primary_color || '#3271AD';
    if (key === 'white') return '#FFFFFF';
    if (key === 'black') return '#000000';
    const n = PALETTE_CHROMA_ORDER.length;
    let idx = PALETTE_CHROMA_ORDER.indexOf(key);
    if (idx < 0) idx = Math.floor((n - 1) / 2); // 不明名は中央
    const factor = n > 1 ? idx / (n - 1) : 0.5;
    if (factor < 0.5) return darkenColor(base, (0.5 - factor) * 0.5);   // 濃いめ（最大25%暗く）
    return lightenColor(base, (factor - 0.5) * 1.0);                    // 淡いめ（最大50%明るく）
  }

  function paletteHex(name, kind) {
    const key = String(name || '').toLowerCase();
    // 配色モード: プライマリカラーの濃淡グラデーション
    if (__COLOR_MODE === 'gradient') {
      const shade = gradientShade(key);
      if (!shade) return null;
      if (kind === 'light') return deriveLight(key, shade);
      if (kind === 'dark') return deriveDark(key, shade);
      return shade;
    }
    // 手動マッピング（テンプレ色割当）優先、無ければ既定パレット
    const mainHex = (__ACCENT_MAP && __ACCENT_MAP[key]) || PALETTE[key];
    if (!mainHex) return null;
    if (kind === 'light') return deriveLight(key, mainHex);
    if (kind === 'dark') return deriveDark(key, mainHex);
    return mainHex;
  }
  function accentMain(name, fallback) {
    return paletteHex(name, 'main') || fallback;
  }

  // ========================================
  // マスターデザイン設定（元コード CONFIG を移植）
  // ========================================
  const CONFIG = {
    BASE_PX: { W: 960, H: 540 },
    BACKGROUND_IMAGES: { title: '', closing: '', section: '', main: '' },
    POS_PX: {
      titleSlide: {
        logo: { left: 55, top: 60, width: 135 },
        title: { left: 50, top: 200, width: 830, height: 90 },
        date: { left: 50, top: 450, width: 250, height: 40 }
      },
      contentSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        body: { left: 25, top: 132, width: 910, height: 330 },
        twoColLeft: { left: 25, top: 132, width: 440, height: 330 },
        twoColRight: { left: 495, top: 132, width: 440, height: 330 }
      },
      compareSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        leftBox: { left: 25, top: 112, width: 445, height: 350 },
        rightBox: { left: 490, top: 112, width: 445, height: 350 }
      },
      processSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 132, width: 910, height: 330 }
      },
      timelineSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 132, width: 910, height: 330 }
      },
      diagramSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        lanesArea: { left: 25, top: 132, width: 910, height: 330 }
      },
      cardsSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        gridArea: { left: 25, top: 120, width: 910, height: 340 }
      },
      tableSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 130, width: 910, height: 330 }
      },
      progressSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 132, width: 910, height: 330 }
      },
      quoteSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 88, width: 260, height: 4 },
        subhead: { left: 25, top: 100, width: 910, height: 40 }
      },
      kpiSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        gridArea: { left: 25, top: 132, width: 910, height: 330 }
      },
      triangleSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        area: { left: 25, top: 110, width: 910, height: 350 }
      },
      flowChartSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        singleRow: { left: 25, top: 160, width: 910, height: 180 },
        upperRow: { left: 25, top: 150, width: 910, height: 120 },
        lowerRow: { left: 25, top: 290, width: 910, height: 120 }
      },
      stepUpSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        stepArea: { left: 25, top: 130, width: 910, height: 330 }
      },
      imageTextSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 80, width: 260, height: 4 },
        subhead: { left: 25, top: 90, width: 910, height: 40 },
        leftImage: { left: 25, top: 150, width: 440, height: 270 },
        leftImageCaption: { left: 25, top: 430, width: 440, height: 30 },
        rightText: { left: 485, top: 150, width: 450, height: 310 },
        leftText: { left: 25, top: 150, width: 450, height: 310 },
        rightImage: { left: 495, top: 150, width: 440, height: 270 },
        rightImageCaption: { left: 495, top: 430, width: 440, height: 30 }
      },
      pyramidSlide: {
        headerLogo: { right: 20, top: 20, width: 75 },
        title: { left: 25, top: 20, width: 830, height: 65 },
        titleUnderline: { left: 25, top: 88, width: 260, height: 4 },
        subhead: { left: 25, top: 100, width: 910, height: 40 },
        pyramidArea: { left: 25, top: 120, width: 910, height: 360 }
      },
      sectionSlide: {
        title: { left: 55, top: 230, width: 840, height: 80 },
        ghostNum: { left: 35, top: 120, width: 400, height: 200 }
      },
      footer: {
        leftText: { left: 15, top: 505, width: 250, height: 20 },
        rightPage: { right: 15, top: 505, width: 50, height: 20 }
      },
      bottomBar: { left: 0, top: 534, width: 960, height: 6 }
    },
    FONTS: {
      family: 'Noto Sans JP',
      sizes: {
        title: 40, date: 16, sectionTitle: 38, contentTitle: 24, subhead: 16,
        body: 14, footer: 9, chip: 11, laneTitle: 13, small: 10,
        processStep: 14, axis: 12, ghostNum: 180
      }
    },
    COLORS: {
      primary_color: '#4285F4',
      text_primary: '#333333',
      background_white: '#FFFFFF',
      card_bg: '#f6e9f0',
      background_gray: '', faint_gray: '', ghost_gray: '', table_header_bg: '',
      lane_border: '', card_border: '', neutral_gray: '', process_arrow: '',
      success_green: '#1e8e3e', error_red: '#d93025',
      slide_bg: '#FFFFFF' // スライド全面の背景色（テンプレート取込で上書き可）
    },
    DIAGRAM: {
      laneGap_px: 24, lanePad_px: 10, laneTitle_h_px: 30, cardGap_px: 12,
      cardMin_h_px: 48, cardMax_h_px: 70, arrow_h_px: 10, arrowGap_px: 8
    },
    LOGOS: { header: '', closing: '' },
    FOOTER_TEXT: ''
  };

  function updateDynamicColors(settings) {
    const primary = settings.primaryColor;
    CONFIG.COLORS.background_gray = generateTintedGray(primary, 10, 98);
    CONFIG.COLORS.faint_gray = generateTintedGray(primary, 10, 93);
    CONFIG.COLORS.ghost_gray = generateTintedGray(primary, 38, 88);
    CONFIG.COLORS.table_header_bg = generateTintedGray(primary, 20, 94);
    CONFIG.COLORS.lane_border = generateTintedGray(primary, 15, 85);
    CONFIG.COLORS.card_border = generateTintedGray(primary, 15, 85);
    CONFIG.COLORS.neutral_gray = generateTintedGray(primary, 5, 62);
    CONFIG.COLORS.process_arrow = CONFIG.COLORS.ghost_gray;
  }

  // ========================================
  // レイアウト管理（px -> inch）
  // ========================================
  const PX_TO_IN = 10 / 960; // 960px -> 10in (540px -> 5.625in も同係数)

  function createLayoutManager() {
    const pxToIn = (px) => px * PX_TO_IN;
    const getPositionFromPath = (path) => path.split('.').reduce((obj, key) => (obj ? obj[key] : undefined), CONFIG.POS_PX);
    return {
      pageW_pt: 10,      // 互換のため名称維持（実体は inch）
      pageH_pt: 5.625,
      pxToPt: pxToIn,    // 互換のため名称維持（実体は px->inch）
      getRect: (spec) => {
        // Phase B（テンプレ駆動）: 本文コンテンツ領域はテンプレ本文枠(bodyFrame)へ確定配置する。
        // インライン矩形指定（quote/mermaid/callout/stats/bar 等）は本文領域とみなし bodyFrame に、
        // 既知のコンテンツ領域キーは bodyFrame の分割（full/left/right/top/bottom）にマップ。
        if (__templateActive && __TPL.bodyRect) {
          if (typeof spec !== 'string') return Object.assign({}, __TPL.bodyRect);
          const kind = CONTENT_AREA_KEYS[spec];
          if (kind) return remapToBodyFrame(kind, __TPL.bodyRect);
        }
        const pos = typeof spec === 'string' ? getPositionFromPath(spec) : spec;
        if (!pos) return { left: undefined, top: 0, width: 0, height: 0 };
        let left_px = pos.left;
        if (pos.right !== undefined && pos.left === undefined) {
          left_px = CONFIG.BASE_PX.W - pos.right - pos.width;
        }
        if (left_px === undefined && pos.right === undefined) left_px = 0;
        return {
          left: left_px !== undefined ? pxToIn(left_px) : undefined,
          top: pos.top !== undefined ? pxToIn(pos.top) : 0,
          width: pos.width !== undefined ? pxToIn(pos.width) : 0,
          height: pos.height !== undefined ? pxToIn(pos.height) : 0
        };
      }
    };
  }

  function safeGetRect(layout, path) {
    try {
      const rect = layout.getRect(path);
      if (rect && typeof rect.top === 'number' && typeof rect.width === 'number' && typeof rect.height === 'number') {
        if (rect.left === undefined) return null;
        return rect;
      }
      return null;
    } catch (e) { return null; }
  }

  function offsetRect(rect, dx, dy) {
    const r = { left: rect.left + (dx || 0), top: rect.top + (dy || 0), width: rect.width, height: rect.height };
    // テンプレ駆動: コンテンツ枠（高さ1in以上）がテンプレ本文枠の下端（=フッター帯の上）を超える分は切り詰め、
    // フッターロゴ/装飾への重なりを防ぐ。全コンテンツ領域は offsetRect(...,0,dy) を通るためここで一括処理。
    if (__templateActive && __TPL.bodyRect && typeof r.top === 'number' && typeof r.height === 'number' && r.height >= 1.0) {
      const safeBottom = __TPL.bodyRect.top + __TPL.bodyRect.height;
      if (r.top < safeBottom && r.top + r.height > safeBottom + 0.02) {
        r.height = Math.max(0.5, safeBottom - r.top);
      }
    }
    return r;
  }

  // ========================================
  // PptxGenJS プリミティブ・ヘルパー
  // ========================================
  function hx(c) { return String(c == null ? '000000' : c).replace('#', ''); }

  const ST = {
    rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', line: 'line',
    rightArrow: 'rightArrow', downArrow: 'downArrow', bentArrow: 'bentArrow'
  };

  /** 塗り図形を描画 */
  function box(slide, type, x, y, w, h, opt = {}) {
    const o = { x: x, y: y, w: w, h: h };
    if (opt.fill === null) {
      o.fill = { type: 'none' };
    } else if (opt.fill) {
      o.fill = (opt.fillTransparency != null)
        ? { color: hx(opt.fill), transparency: opt.fillTransparency }
        : { color: hx(opt.fill) };
    }
    if (opt.line) o.line = { color: hx(opt.line), width: opt.lineWidth || 1 };
    else o.line = { type: 'none' };
    if (opt.rotate) o.rotate = opt.rotate;
    if (opt.flipH) o.flipH = true;
    if (opt.flipV) o.flipV = true;
    try { slide.addShape(type, o); } catch (e) { console.warn('box error', e); }
  }

  /** スタイル付きテキスト（必要なら塗り図形を兼ねる） */
  function txt(slide, x, y, w, h, rawText, opt = {}, ctx = {}) {
    const runs = toRuns(rawText, opt, ctx);
    const o = {
      x: x, y: y, w: w, h: h,
      margin: (opt.margin != null ? opt.margin : 2),
      fontFace: CONFIG.FONTS.family,
      fontSize: opt.size || CONFIG.FONTS.sizes.body,
      color: hx(opt.color || CONFIG.COLORS.text_primary),
      bold: !!opt.bold,
      align: opt.align || 'left',
      valign: opt.valign || 'top',
      wrap: true
    };
    if (opt.fontFace) o.fontFace = opt.fontFace;
    if (opt.shape) o.shape = opt.shape;
    if (opt.fill) o.fill = { color: hx(opt.fill) };
    if (opt.line) o.line = { color: hx(opt.line), width: opt.lineWidth || 1 };
    if (opt.lineSpacingMultiple) o.lineSpacingMultiple = opt.lineSpacingMultiple;
    if (opt.paraSpaceAfter != null) o.paraSpaceAfter = opt.paraSpaceAfter;
    if (opt.shrink) o.fit = 'shrink';
    try { slide.addText(runs, o); } catch (e) { console.warn('txt error', e); }
  }

  /** 直線（任意で終端矢印） */
  function lineSeg(slide, x1, y1, x2, y2, opt = {}) {
    const x = Math.min(x1, x2), y = Math.min(y1, y2);
    const w = Math.abs(x2 - x1), h = Math.abs(y2 - y1);
    const o = { x: x, y: y, w: w, h: h, line: { color: hx(opt.color), width: opt.width || 1 } };
    if (x2 < x1) o.flipH = true;
    if (y2 < y1) o.flipV = true;
    if (opt.arrow) o.line.endArrowType = 'triangle';
    try { slide.addShape(ST.line, o); } catch (e) { console.warn('line error', e); }
  }

  /** 画像をフレーム内にフィット（失敗時スキップ） */
  function img(slide, x, y, w, h, url, sizing = 'contain') {
    if (!url) return;
    try {
      const o = { path: url, x: x, y: y, w: w, h: h };
      if (sizing) o.sizing = { type: sizing, w: w, h: h };
      slide.addImage(o);
    } catch (e) { console.warn('image skipped:', url, e); }
  }

  /** data URI 画像を指定位置に配置（mermaid PNG 等） */
  function imgData(slide, x, y, w, h, dataUri) {
    if (!dataUri) return;
    try { slide.addImage({ data: dataUri, x: x, y: y, w: w, h: h }); }
    catch (e) { console.warn('imgData skipped:', e); }
  }

  /**
   * スライドをラップし、addShape/addText/addImage/addTable のオプション x,y に (dx,dy) を加算する。
   * コンテンツは常に 10×5.625 で作図し、A4等の拡大ページでは中央寄せをここで一括適用する。
   * background= / addNotes 等は素通し。dx=dy=0 のときは元のスライドをそのまま返す（無回帰）。
   */
  function makeOffsetSlide(slide, dx, dy) {
    if (!dx && !dy) return slide;
    const shift = (o) => {
      if (o && typeof o === 'object') {
        const c = Object.assign({}, o);
        if (typeof c.x === 'number') c.x = c.x + dx;
        if (typeof c.y === 'number') c.y = c.y + dy;
        return c;
      }
      return o;
    };
    return new Proxy(slide, {
      get: function (target, prop) {
        const val = target[prop];
        if (typeof val !== 'function') return val;
        if (prop === 'addShape') return function (type, opts) { return target.addShape(type, shift(opts)); };
        if (prop === 'addText') return function (t, opts) { return target.addText(t, shift(opts)); };
        if (prop === 'addImage') return function (opts) { return target.addImage(shift(opts)); };
        if (prop === 'addTable') return function (rows, opts) { return target.addTable(rows, shift(opts)); };
        return val.bind(target);
      },
      set: function (target, prop, value) { target[prop] = value; return true; }
    });
  }

  /**
   * マスター定義 objects の座標(x,y,w,h)を一律 s 倍（テンプレ装飾を出力キャンバス全面へ正規化）。
   * 元配列・元オブジェクトは破壊せず新配列を返す（data URI 文字列は参照共有）。
   */
  function scaleObjects(objects, s) {
    if (!objects) return objects;
    if (!s || s === 1) return objects.slice();
    return objects.map(function (item) {
      if (!item || typeof item !== 'object') return item;
      if (item.text) {
        const opts = Object.assign({}, (item.text.options || {}));
        if (typeof opts.x === 'number') opts.x *= s;
        if (typeof opts.y === 'number') opts.y *= s;
        if (typeof opts.w === 'number') opts.w *= s;
        if (typeof opts.h === 'number') opts.h *= s;
        return { text: Object.assign({}, item.text, { options: opts }) };
      }
      const key = item.rect ? 'rect' : item.line ? 'line' : item.image ? 'image' : null;
      if (!key) return item;
      const o = Object.assign({}, item[key]);
      if (typeof o.x === 'number') o.x *= s;
      if (typeof o.y === 'number') o.y *= s;
      if (typeof o.w === 'number') o.w *= s;
      if (typeof o.h === 'number') o.h *= s;
      const wrap = {}; wrap[key] = o; return wrap;
    });
  }

  /** プレースホルダ rect {left,top,width,height} を一律 s 倍（null安全） */
  function scaleRect(rect, s) {
    if (!rect || !s || s === 1) return rect;
    return { left: rect.left * s, top: rect.top * s, width: rect.width * s, height: rect.height * s };
  }

  // まじん式タイプ → テンプレ・レイアウトカテゴリ（LayoutResolver 用）
  const TYPE_TO_CATEGORY = {
    title: 'title', section: 'section', closing: 'closing',
    compare: 'compare', statsCompare: 'compare', barCompare: 'compare',
    table: 'table',
    mermaid: 'diagram', flowChart: 'diagram', cycle: 'diagram', diagram: 'diagram', pyramid: 'diagram', triangle: 'diagram',
    cards: 'cards', headerCards: 'cards', bulletCards: 'cards', kpi: 'cards'
    // 既定（content）: content, agenda, process, processList, timeline, quote, faq, stepUp, progress, callout, calloutGrid, iconBanner, imageText
  };

  // LayoutResolver: スライドtype（または明示 templateLayout id）→ テンプレ レイアウトカタログの1件
  function resolveTemplateLayout(type, layoutId, tpl) {
    const cat = (tpl && tpl.layoutCatalog) || [];
    if (!cat.length) return null;
    if (layoutId) { const e = cat.find(function (x) { return x.id === layoutId; }); if (e) return e; }
    const want = TYPE_TO_CATEGORY[type] || 'content';
    return cat.find(function (x) { return x.category === want && (x.titleFrame || x.bodyFrame); })
      || cat.find(function (x) { return x.category === want; })
      || cat.find(function (x) { return x.category === 'content' && (x.titleFrame || x.bodyFrame); })
      || cat.find(function (x) { return x.category === 'content'; })
      || cat[0];
  }

  // Phase B: 各 generator の本文コンテンツ領域キー → テンプレ本文枠(bodyFrame)の分割種別
  const CONTENT_AREA_KEYS = {
    'contentSlide.body': 'full', 'contentSlide.twoColLeft': 'left', 'contentSlide.twoColRight': 'right',
    'compareSlide.leftBox': 'left', 'compareSlide.rightBox': 'right',
    'processSlide.area': 'full', 'timelineSlide.area': 'full', 'diagramSlide.lanesArea': 'full',
    'cardsSlide.gridArea': 'full', 'tableSlide.area': 'full', 'progressSlide.area': 'full',
    'kpiSlide.gridArea': 'full', 'triangleSlide.area': 'full', 'pyramidSlide.pyramidArea': 'full',
    'stepUpSlide.area': 'full',
    'flowChartSlide.upperRow': 'top', 'flowChartSlide.lowerRow': 'bottom', 'flowChartSlide.singleRow': 'full'
  };
  function remapToBodyFrame(kind, BF) {
    const gap = 0.18;
    const halfW = Math.max(0.4, (BF.width - gap) / 2);
    const halfH = Math.max(0.4, (BF.height - gap) / 2);
    switch (kind) {
      case 'left': return { left: BF.left, top: BF.top, width: halfW, height: BF.height };
      case 'right': return { left: BF.left + halfW + gap, top: BF.top, width: halfW, height: BF.height };
      case 'top': return { left: BF.left, top: BF.top, width: BF.width, height: halfH };
      case 'bottom': return { left: BF.left, top: BF.top + halfH + gap, width: BF.width, height: halfH };
      default: return { left: BF.left, top: BF.top, width: BF.width, height: BF.height };
    }
  }

  // ========================================
  // テキスト・インラインスタイル
  // ========================================
  function parseInlineStyles(s) {
    const ranges = [];
    let out = '';
    let i = 0;
    while (i < s.length) {
      if (s[i] === '*' && s[i + 1] === '*' && s[i + 2] === '[' && s[i + 3] === '[') {
        const contentStart = i + 4;
        const close = s.indexOf(']]**', contentStart);
        if (close !== -1) {
          const content = s.substring(contentStart, close);
          const start = out.length; out += content; const end = out.length;
          ranges.push({ start, end, bold: true, color: CONFIG.COLORS.primary_color });
          i = close + 4; continue;
        }
      }
      if (s[i] === '[' && s[i + 1] === '[') {
        const close = s.indexOf(']]', i + 2);
        if (close !== -1) {
          const content = s.substring(i + 2, close);
          const start = out.length; out += content; const end = out.length;
          ranges.push({ start, end, bold: true, color: CONFIG.COLORS.primary_color });
          i = close + 2; continue;
        }
      }
      if (s[i] === '*' && s[i + 1] === '*') {
        const close = s.indexOf('**', i + 2);
        if (close !== -1) {
          const content = s.substring(i + 2, close);
          if (content.indexOf('[[') === -1) {
            const start = out.length; out += content; const end = out.length;
            ranges.push({ start, end, bold: true });
            i = close + 2; continue;
          } else { i += 2; continue; }
        }
      }
      out += s[i]; i++;
    }
    return { output: out, ranges };
  }

  /** parseInlineStyles の結果を PptxGenJS のテキストラン配列へ変換（改行対応） */
  function toRuns(rawText, baseOpt = {}, ctx = {}) {
    const parsed = parseInlineStyles(String(rawText == null ? '' : rawText));
    const text = parsed.output;
    const ranges = parsed.ranges;
    const defColor = baseOpt.color || CONFIG.COLORS.text_primary;
    const defBold = !!baseOpt.bold;

    const optsAt = (idx) => {
      let bold = defBold;
      let color = defColor;
      for (const r of ranges) {
        if (idx >= r.start && idx < r.end) {
          if (r.bold) bold = true;
          if (r.color) {
            let fc = r.color;
            if (ctx.bgColor && ctx.primaryColor &&
              String(ctx.bgColor).toLowerCase() === String(ctx.primaryColor).toLowerCase()) {
              fc = CONFIG.COLORS.background_white;
            }
            color = fc;
          }
        }
      }
      return { bold, color: hx(color) };
    };

    const n = text.length;
    if (n === 0) return [{ text: '', options: {} }];

    const runs = [];
    let cur = null; // { text, key, options }
    for (let idx = 0; idx < n; idx++) {
      const ch = text[idx];
      if (ch === '\n') {
        if (cur) { cur.options.breakLine = true; runs.push(cur); cur = null; }
        else { runs.push({ text: '', options: { breakLine: true } }); }
        continue;
      }
      const o = optsAt(idx);
      const key = (o.bold ? '1' : '0') + '|' + o.color;
      if (cur && cur.key === key) {
        cur.text += ch;
      } else {
        if (cur) runs.push(cur);
        cur = { text: ch, key: key, options: { bold: o.bold, color: o.color } };
      }
    }
    if (cur) runs.push(cur);
    return runs.map(r => ({ text: r.text, options: r.options }));
  }

  /** 箇条書き（各pointを段落とし、間隔をあける） */
  function bulletText(slide, x, y, w, h, points, opt = {}) {
    const list = (points || []);
    const runs = [];
    if (list.length === 0) {
      runs.push({ text: '—', options: {} });
    } else {
      list.forEach((pt, idx) => {
        const r = toRuns(String(pt == null ? '' : pt), { size: CONFIG.FONTS.sizes.body });
        r.forEach(run => runs.push(run));
        // 各項目末尾に改行（最終項目以外）
        if (idx < list.length - 1) {
          runs.push({ text: '', options: { breakLine: true } });
        }
      });
    }
    try {
      slide.addText(runs, {
        x: x, y: y, w: w, h: h, margin: 2,
        fontFace: CONFIG.FONTS.family,
        fontSize: CONFIG.FONTS.sizes.body,
        color: hx(CONFIG.COLORS.text_primary),
        align: 'left', valign: opt.valign || 'top', wrap: true,
        lineSpacingMultiple: 1.15, paraSpaceAfter: 8
      });
    } catch (e) { console.warn('bulletText error', e); }
  }

  function cleanSpeakerNotes(notesText) {
    if (!notesText) return '';
    let cleaned = notesText;
    cleaned = cleaned.replace(/\*\*([^*]+)\*\*/g, '$1');
    cleaned = cleaned.replace(/\[\[([^\]]+)\]\]/g, '$1');
    cleaned = cleaned.replace(/\*([^*]+)\*/g, '$1');
    cleaned = cleaned.replace(/_([^_]+)_/g, '$1');
    cleaned = cleaned.replace(/~~([^~]+)~~/g, '$1');
    cleaned = cleaned.replace(/`([^`]+)`/g, '$1');
    return cleaned;
  }

  function estimateTextWidthPt(text, fontSizePt) {
    const multipliers = { ascii: 0.62, japanese: 1.0, other: 0.85 };
    return String(text || '').split('').reduce((acc, char) => {
      if (char.match(/[ -~]/)) return acc + multipliers.ascii;
      else if (char.match(/[぀-ゟ゠-ヿ一-龯]/)) return acc + multipliers.japanese;
      else return acc + multipliers.other;
    }, 0) * fontSizePt;
  }

  function parseNumericValue(str) {
    if (typeof str !== 'string') return 0;
    const match = str.match(/(\d+(\.\d+)?)/);
    return match ? parseFloat(match[1]) : 0;
  }

  // 折返し行数の概算（段落=\n 区切り、各段落を幅で割って切上げ）
  function estimateLines(text, widthIn, sizePt) {
    const widthPt = Math.max(1, widthIn * 72);
    const paras = String(text == null ? '' : text).split('\n');
    let lines = 0;
    paras.forEach(function (p) {
      const wpt = estimateTextWidthPt(p, sizePt);
      lines += Math.max(1, Math.ceil(wpt / widthPt));
    });
    return Math.max(1, lines);
  }

  // 概算レンダ高（inch、行高1.25）
  function blockHeightIn(text, widthIn, sizePt) {
    return estimateLines(text, widthIn, sizePt) * sizePt * 1.25 / 72;
  }

  // 枠(幅×高さ inch)に収まる最大フォントサイズ(pt)を base から min まで縮小して算出
  function autoSizePt(text, widthIn, heightIn, baseSizePt, minSizePt) {
    let size = baseSizePt;
    const minS = minSizePt || 12;
    while (size > minS && blockHeightIn(text, widthIn, size) > heightIn) {
      size -= 1;
    }
    return size;
  }

  function isAgendaTitle(title) {
    return /(agenda|アジェンダ|目次|本日お伝えすること)/i.test(String(title || ''));
  }

  let __SECTION_COUNTER = 0;
  let __SLIDE_DATA_FOR_AGENDA = [];

  // テンプレート・マスター取り込み状態
  let __templateActive = false;
  let __TPL = { titleRect: null, bodyRect: null, titleSlideTitleRect: null, category: null };

  // 出力形式に応じたコンテンツ中央オフセット（16:9=0,0 / A4横=0.415,0.9375）。Proxyで全描画に加算。
  let __CONTENT_DX = 0;
  let __CONTENT_DY = 0;

  // まじん式+ : パレット名→実色の手動マッピング（テンプレ色への対応付け）。null/未指定はパレット既定
  let __ACCENT_MAP = null;
  // まじん式+ : 配色モード 'accent'（対応表どおり）| 'gradient'（プライマリ濃淡）
  let __COLOR_MODE = 'accent';

  // まじん式+ : 現在スライドのタイトル絵文字（ループで設定）
  let __CURRENT_ICON = '';
  function withIcon(title) {
    return (__CURRENT_ICON ? __CURRENT_ICON + ' ' : '') + (title == null ? '' : title);
  }

  function buildAgendaFromSlideData() {
    return __SLIDE_DATA_FOR_AGENDA.filter(d => d && d.type === 'section' && d.title).map(d => d.title.trim());
  }

  // ========================================
  // 共通描画ヘルパー
  // ========================================
  function setBackgroundImageFromUrl(slide, layout, imageUrl, fallbackColor) {
    if (imageUrl) {
      slide.background = { color: hx(fallbackColor) };
      // 背景画像はスライド最初の要素として配置（最背面）
      try {
        slide.addImage({
          path: imageUrl, x: 0, y: 0, w: layout.pageW_pt, h: layout.pageH_pt,
          sizing: { type: 'cover', w: layout.pageW_pt, h: layout.pageH_pt }
        });
      } catch (e) { console.warn('bg image skipped:', imageUrl, e); }
      return;
    }
    // テンプレート適用中は、マスターの背景を見せるため単色で覆わない
    if (__templateActive) return;
    slide.background = { color: hx(fallbackColor) };
  }

  function setMainSlideBackground(slide, layout) {
    setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.main, CONFIG.COLORS.slide_bg);
  }

  function createGradientRectangle(slide, x, y, width, height, colors) {
    const numStrips = 40;
    const stripWidth = width / numStrips;
    const startColor = hexToRgb(colors[0]), endColor = hexToRgb(colors[1]);
    if (!startColor || !endColor) return;
    for (let i = 0; i < numStrips; i++) {
      const ratio = i / (numStrips - 1);
      const r = Math.round(startColor.r + (endColor.r - startColor.r) * ratio);
      const g = Math.round(startColor.g + (endColor.g - startColor.g) * ratio);
      const b = Math.round(startColor.b + (endColor.b - startColor.b) * ratio);
      const hexc = '#' + [r, g, b].map(v => v.toString(16).padStart(2, '0')).join('');
      box(slide, ST.rect, x + (i * stripWidth), y, stripWidth + 0.005, height, { fill: hexc });
    }
  }

  function applyFill(slide, x, y, width, height, settings) {
    if (settings.enableGradient) {
      createGradientRectangle(slide, x, y, width, height, [settings.gradientStart, settings.gradientEnd]);
    } else {
      box(slide, ST.rect, x, y, width, height, { fill: settings.primaryColor });
    }
  }

  function createContentCushion(slide, area, settings, layout) {
    if (!area || !area.width || !area.height || area.width <= 0 || area.height <= 0) return;
    box(slide, ST.rect, area.left, area.top, area.width, area.height,
      { fill: CONFIG.COLORS.background_gray, fillTransparency: 50 });
  }

  function placeLogo(slide, rect, url) {
    if (!url || !rect) return;
    img(slide, rect.left, rect.top, rect.width, rect.width, url, 'contain'); // 正方フレームにcontain
  }

  // 図/表/画像の差し込み位置を示すダミー枠（淡グレー＋枠線＋「ここに図を挿入：説明」）
  function drawFigurePlaceholder(slide, rect, label) {
    if (!rect || !(rect.width > 0) || !(rect.height > 0)) return;
    box(slide, ST.roundRect, rect.left, rect.top, rect.width, rect.height,
      { fill: CONFIG.COLORS.background_gray, fillTransparency: 20, line: CONFIG.COLORS.neutral_gray, lineWidth: 1.5 });
    const pad = 0.12;
    const text = '🖼 ここに図を挿入' + (label ? '\n' + label : '');
    txt(slide, rect.left + pad, rect.top + pad, Math.max(0.5, rect.width - pad * 2), Math.max(0.4, rect.height - pad * 2), text,
      { size: 12, color: CONFIG.COLORS.neutral_gray, align: 'center', valign: 'middle' });
  }

  // テンプレ図枠（entry.figureFrames）または予約枠(figureReserve)にダミー図を配置。
  // ラベルは data.figure（string または配列）。data.figure 未指定時は描画しない。
  function drawFigurePlaceholders(slide, data, entry, tplScale, figureReserve) {
    const rawLabels = Array.isArray(data.figure) ? data.figure
      : (data.figure != null && data.figure !== '' ? [data.figure] : []);
    if (!rawLabels.length) return;
    let frames = [];
    if (entry && Array.isArray(entry.figureFrames) && entry.figureFrames.length) {
      frames = entry.figureFrames.map(function (ff) { return scaleRect({ left: ff.left, top: ff.top, width: ff.width, height: ff.height }, tplScale); });
    } else if (figureReserve) {
      frames = [figureReserve];
    }
    if (!frames.length) return;
    frames.forEach(function (fr, i) {
      const label = rawLabels[i] != null ? rawLabels[i] : (rawLabels.length === 1 ? rawLabels[0] : '');
      drawFigurePlaceholder(slide, fr, label);
    });
  }

  function drawStandardTitleHeader(slide, layout, key, title, settings) {
    const logoRect = safeGetRect(layout, `${key}.headerLogo`);
    if (CONFIG.LOGOS.header && logoRect) placeLogo(slide, logoRect, CONFIG.LOGOS.header);

    let titleRect = safeGetRect(layout, `${key}.title`);
    if (!titleRect) return;
    // テンプレート取り込み: タイトル位置をプレースホルダに合わせる
    const usingTplFrame = !!(__templateActive && __TPL.titleRect);
    if (usingTplFrame) titleRect = __TPL.titleRect;
    const fontSize = CONFIG.FONTS.sizes.contentTitle;
    const titleText = withIcon(title || '');
    // テンプレ枠ではタイトル枠の高さいっぱいに描画（テキストが収まる）、それ以外は従来の最適高
    const boxH = (usingTplFrame && titleRect.height) ? titleRect.height : layout.pxToPt(fontSize + 8);
    txt(slide, titleRect.left, titleRect.top, titleRect.width, boxH, titleText,
      { size: fontSize, bold: true, valign: 'top', shrink: true });

    if (settings.showTitleUnderline && title) {
      let uRect = safeGetRect(layout, `${key}.titleUnderline`);
      if (!uRect) return;
      let uLeft = uRect.left, uTop = uRect.top, uH = uRect.height;
      if (usingTplFrame) {
        // 下線の縦位置 = タイトル文字の「実レンダ高さ」の直下（重なり防止）。
        // 枠高でのクランプは行わない（枠がテキストにぴったりだと線が文字に食い込むため）。
        // 行高はやや余裕をもって見積もり、確実に文字の下へ配置する。
        const textH = blockHeightIn(titleText, titleRect.width, fontSize) * 1.05;
        uLeft = titleRect.left;
        uTop = titleRect.top + textH + layout.pxToPt(3);
      }
      // 下線はタイトル文字幅（アイコン込み）に合わせる。上限はタイトル枠幅（無ければ余白考慮の本文幅）。
      const estimatedWidthIn = estimateTextWidthPt(titleText, fontSize) / 72;
      const frameCap = (titleRect.width && titleRect.width > 0)
        ? titleRect.width
        : (layout.pageW_pt - uLeft - layout.pxToPt(25));
      const finalWidth = Math.max(layout.pxToPt(40), Math.min(estimatedWidthIn, frameCap));
      applyFill(slide, uLeft, uTop, finalWidth, uH, settings);
    }
  }

  function drawSubheadIfAny(slide, layout, key, subhead) {
    // テンプレ駆動時: subheadの有無にかかわらず常に1行分オフセットを確保
    if (__templateActive && __TPL.bodyRect) {
      const br = __TPL.bodyRect;
      const subH = layout.pxToPt(40);
      if (subhead) {
        txt(slide, br.left, br.top, br.width, subH, subhead,
          { size: CONFIG.FONTS.sizes.subhead, color: CONFIG.COLORS.text_primary, valign: 'top' });
      }
      return subH + layout.pxToPt(8);
    }
    if (!subhead) return 0;
    const rect = safeGetRect(layout, `${key}.subhead`);
    if (!rect) return 0;
    txt(slide, rect.left, rect.top, rect.width, rect.height, subhead,
      { size: CONFIG.FONTS.sizes.subhead, color: CONFIG.COLORS.text_primary, valign: 'top' });
    return layout.pxToPt(36);
  }

  function drawBottomBar(slide, layout, settings) {
    const barRect = layout.getRect('bottomBar');
    applyFill(slide, barRect.left, barRect.top, barRect.width, barRect.height, settings);
  }

  function addCucFooter(slide, layout, pageNum) {
    if (CONFIG.FOOTER_TEXT && CONFIG.FOOTER_TEXT.trim() !== '') {
      const leftRect = layout.getRect('footer.leftText');
      txt(slide, leftRect.left, leftRect.top, leftRect.width, leftRect.height, CONFIG.FOOTER_TEXT,
        { size: CONFIG.FONTS.sizes.footer, color: CONFIG.COLORS.text_primary, valign: 'middle' });
    }
    if (pageNum > 0) {
      const rightRect = layout.getRect('footer.rightPage');
      txt(slide, rightRect.left, rightRect.top, rightRect.width, rightRect.height, String(pageNum),
        { size: CONFIG.FONTS.sizes.footer, color: CONFIG.COLORS.primary_color, align: 'right', valign: 'middle' });
    }
  }

  function drawBottomBarAndFooter(slide, layout, pageNum, settings) {
    if (settings.showBottomBar && !__templateActive) drawBottomBar(slide, layout, settings);
    addCucFooter(slide, layout, pageNum);
  }

  function drawArrowBetweenRects(slide, a, b, settings) {
    const fromX = a.left + a.width;
    const fromY = a.top + a.height / 2;
    const toX = b.left;
    const toY = b.top + b.height / 2;
    if (toX - fromX <= 0) return;
    lineSeg(slide, fromX, fromY, toX, toY, { color: settings.primaryColor, width: 1.5, arrow: true });
  }

  function drawNumberedItems(slide, layout, area, items, settings) {
    createContentCushion(slide, area, settings, layout);
    const n = Math.max(1, items.length);
    const topPadding = layout.pxToPt(30);
    const bottomPadding = layout.pxToPt(10);
    const drawableHeight = area.height - topPadding - bottomPadding;
    const gapY = drawableHeight / Math.max(1, n - 1);
    const cx = area.left + layout.pxToPt(44);
    const top0 = area.top + topPadding;

    box(slide, ST.rect, cx - layout.pxToPt(1), top0 + layout.pxToPt(6), layout.pxToPt(2), gapY * (n - 1),
      { fill: CONFIG.COLORS.faint_gray });

    for (let i = 0; i < n; i++) {
      const cy = top0 + gapY * i;
      const sz = layout.pxToPt(28);
      txt(slide, cx - sz / 2, cy - sz / 2, sz, sz, String(i + 1),
        { size: 12, bold: true, color: CONFIG.COLORS.background_white, align: 'center', valign: 'middle', shape: ST.rect, fill: settings.primaryColor });
      let cleanText = String(items[i] || '').replace(/^\s*\d+[\.\s]*/, '');
      txt(slide, cx + layout.pxToPt(28), cy - layout.pxToPt(16), area.width - layout.pxToPt(70), layout.pxToPt(32), cleanText,
        { size: CONFIG.FONTS.sizes.processStep, valign: 'middle' });
    }
  }

  function drawCompareBox(slide, layout, rect, title, items, settings, isLeft, headerColorOverride) {
    box(slide, ST.rect, rect.left, rect.top, rect.width, rect.height,
      { fill: CONFIG.COLORS.background_gray, line: CONFIG.COLORS.lane_border });
    const th = layout.pxToPt(40);
    const compareColors = generateCompareColors(settings.primaryColor);
    const headerColor = headerColorOverride || (isLeft ? compareColors.left : compareColors.right);
    txt(slide, rect.left, rect.top, rect.width, th, title,
      { size: CONFIG.FONTS.sizes.laneTitle, bold: true, color: CONFIG.COLORS.background_white, align: 'center', valign: 'middle', shape: ST.rect, fill: headerColor });
    const pad = layout.pxToPt(12);
    bulletText(slide, rect.left + pad, rect.top + th + pad, rect.width - pad * 2, rect.height - th - pad * 2, items);
  }

  function renderSingleImageInArea(slide, layout, area, imageUrl, caption, position) {
    if (!imageUrl) return;
    img(slide, area.left, area.top, area.width, area.height, imageUrl, 'contain');
    if (caption && caption.trim()) {
      const captionHeight = layout.pxToPt(30);
      const captionTop = area.top + area.height + layout.pxToPt(8);
      txt(slide, area.left, captionTop, area.width, captionHeight, caption.trim(),
        { size: CONFIG.FONTS.sizes.small, color: CONFIG.COLORS.neutral_gray, align: 'center', valign: 'top' });
    }
  }

  function smartFormatTriangleText(text) {
    if (!text || text.length <= 30) return { text: text, isSimple: true, headerLength: 0 };
    const separators = [
      { pattern: '：', priority: 1 }, { pattern: ':', priority: 2 }, { pattern: '。', priority: 3 },
      { pattern: 'について', priority: 4, keepSeparator: true }, { pattern: 'における', priority: 5, keepSeparator: true }
    ];
    for (let sep of separators) {
      const index = text.indexOf(sep.pattern);
      if (index > 5 && index < text.length * 0.6) {
        const headerEnd = sep.keepSeparator ? index + sep.pattern.length : index;
        const header = text.substring(0, headerEnd).trim();
        const body = text.substring(index + sep.pattern.length).trim();
        if (header.length >= 3 && body.length >= 3) {
          return { text: `${header}\n${body}`, isSimple: false, headerLength: header.length };
        }
      }
    }
    if (text.length > 50) {
      const midPoint = Math.floor(text.length * 0.4);
      const header = text.substring(0, midPoint).trim();
      const body = text.substring(midPoint).trim();
      return { text: `${header}\n${body}`, isSimple: false, headerLength: header.length };
    }
    return { text: text, isSimple: true, headerLength: 0 };
  }

  // ========================================
  // スライドジェネレーター
  // ========================================
  function createTitleSlide(slide, data, layout, pageNum, settings) {
    setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.title, CONFIG.COLORS.slide_bg);
    const logoRect = layout.getRect('titleSlide.logo');
    if (CONFIG.LOGOS.header) placeLogo(slide, logoRect, CONFIG.LOGOS.header);
    let titleRect = layout.getRect('titleSlide.title');
    if (__templateActive && __TPL.titleSlideTitleRect) titleRect = __TPL.titleSlideTitleRect;
    const titleText = withIcon(data.title);
    // 長いタイトルは枠に収まるよう自動縮小（許容高さ ≈ 150px 相当、最小24pt）
    const allotIn = Math.max(titleRect.height, layout.pxToPt(150));
    const titleSize = autoSizePt(titleText, titleRect.width, allotIn, CONFIG.FONTS.sizes.title, 24);
    const titleH = Math.max(titleRect.height, blockHeightIn(titleText, titleRect.width, titleSize));
    txt(slide, titleRect.left, titleRect.top, titleRect.width, titleH, titleText,
      { size: titleSize, bold: true, valign: 'top', shrink: true });
    // まじん式+: サブタイトル。テンプレ表紙レイアウトのサブタイトル枠があればそこへ（座標忠実）、
    // 無ければタイトルの実レンダ高に追従して重なりを防止。
    if (data.subtitle) {
      let subRect;
      if (__templateActive && __TPL.category === 'title' && __TPL.bodyRect) {
        subRect = __TPL.bodyRect;
      } else {
        subRect = { left: titleRect.left, top: titleRect.top + titleH + layout.pxToPt(8), width: titleRect.width, height: layout.pxToPt(40) };
      }
      txt(slide, subRect.left, subRect.top, subRect.width, subRect.height, data.subtitle,
        { size: 20, color: CONFIG.COLORS.neutral_gray, valign: 'top' });
    }
    if (settings.showDateColumn) {
      const dateRect = layout.getRect('titleSlide.date');
      txt(slide, dateRect.left, dateRect.top, dateRect.width, dateRect.height, data.date || '',
        { size: CONFIG.FONTS.sizes.date, valign: 'top' });
    }
    if (settings.showBottomBar && !__templateActive) drawBottomBar(slide, layout, settings);
  }

  function createSectionSlide(slide, data, layout, pageNum, settings) {
    setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.section, settings.backgroundColor || CONFIG.COLORS.background_gray);
    __SECTION_COUNTER++;
    const parsedNum = (() => {
      if (Number.isFinite(data.sectionNo)) return Number(data.sectionNo);
      const m = String(data.title || '').match(/^\s*(\d+)[\.．]/);
      return m ? Number(m[1]) : __SECTION_COUNTER;
    })();
    const num = String(parsedNum).padStart(2, '0');
    const ghostRect = layout.getRect('sectionSlide.ghostNum');
    txt(slide, ghostRect.left, ghostRect.top, ghostRect.width, ghostRect.height, num,
      { size: CONFIG.FONTS.sizes.ghostNum, bold: true, color: CONFIG.COLORS.ghost_gray, valign: 'middle', align: 'left' });
    let titleRect = layout.getRect('sectionSlide.title');
    // テンプレの章扉レイアウトにタイトル枠があればそこへ（座標忠実）
    if (__templateActive && __TPL.category === 'section' && __TPL.titleRect) titleRect = __TPL.titleRect;
    txt(slide, titleRect.left, titleRect.top, titleRect.width, titleRect.height, withIcon(data.title),
      { size: CONFIG.FONTS.sizes.sectionTitle, bold: true, align: 'center', valign: 'middle', shrink: true });
    addCucFooter(slide, layout, pageNum);
  }

  function createClosingSlide(slide, data, layout, pageNum, settings) {
    setBackgroundImageFromUrl(slide, layout, CONFIG.BACKGROUND_IMAGES.closing, CONFIG.COLORS.slide_bg);
    if (CONFIG.LOGOS.closing) {
      const wIn = layout.pxToPt(450);
      const hIn = layout.pxToPt(data.message ? 220 : 300);
      const topShift = data.message ? layout.pxToPt(40) : 0;
      img(slide, (layout.pageW_pt - wIn) / 2, (layout.pageH_pt - hIn) / 2 - topShift, wIn, hIn, CONFIG.LOGOS.closing, 'contain');
    }
    // まじん式+: クロージングメッセージ。テンプレの結びレイアウト枠があればそこへ（座標忠実）。
    if (data.message) {
      let msgRect;
      if (__templateActive && __TPL.category === 'closing' && (__TPL.titleRect || __TPL.bodyRect)) {
        msgRect = __TPL.titleRect || __TPL.bodyRect;
      } else {
        const my = CONFIG.LOGOS.closing ? layout.pageH_pt * 0.66 : layout.pageH_pt * 0.42;
        msgRect = { left: layout.pxToPt(60), top: my, width: layout.pageW_pt - layout.pxToPt(120), height: layout.pxToPt(80) };
      }
      txt(slide, msgRect.left, msgRect.top, msgRect.width, msgRect.height, data.message,
        { size: 28, bold: true, align: 'center', valign: 'middle', color: CONFIG.COLORS.text_primary, shrink: true });
    }
  }

  function createContentSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'contentSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead);
    const isAgenda = isAgendaTitle(data.title || '');
    let points = Array.isArray(data.points) ? data.points.slice(0) : [];
    if (isAgenda && points.length === 0) {
      points = buildAgendaFromSlideData();
      if (points.length === 0) points = ['本日の目的', '進め方', '次のアクション'];
    }
    const isTwo = !!(data.twoColumn || data.columns);
    const padding = layout.pxToPt(20);
    if ((isTwo && (data.columns || points)) || (!isTwo && points && points.length > 0)) {
      if (isTwo) {
        let L = [], R = [];
        if (Array.isArray(data.columns) && data.columns.length === 2) {
          L = data.columns[0] || []; R = data.columns[1] || [];
        } else {
          const mid = Math.ceil(points.length / 2);
          L = points.slice(0, mid); R = points.slice(mid);
        }
        let leftRect, rightRect;
        if (__templateActive && __TPL.bodyRect) {
          // テンプレ本文枠を左右2分割
          const b = offsetRect(__TPL.bodyRect, 0, dy);
          const colGap = layout.pxToPt(30);
          const colW = (b.width - colGap) / 2;
          leftRect = { left: b.left, top: b.top, width: colW, height: b.height };
          rightRect = { left: b.left + colW + colGap, top: b.top, width: colW, height: b.height };
        } else {
          leftRect = offsetRect(layout.getRect('contentSlide.twoColLeft'), 0, dy);
          rightRect = offsetRect(layout.getRect('contentSlide.twoColRight'), 0, dy);
        }
        createContentCushion(slide, leftRect, settings, layout);
        createContentCushion(slide, rightRect, settings, layout);
        bulletText(slide, leftRect.left + padding, leftRect.top + padding, leftRect.width - padding * 2, leftRect.height - padding * 2, L);
        bulletText(slide, rightRect.left + padding, rightRect.top + padding, rightRect.width - padding * 2, rightRect.height - padding * 2, R);
      } else {
        const bodyRect = (__templateActive && __TPL.bodyRect)
          ? offsetRect(__TPL.bodyRect, 0, dy)
          : offsetRect(layout.getRect('contentSlide.body'), 0, dy);
        createContentCushion(slide, bodyRect, settings, layout);
        if (isAgenda) {
          drawNumberedItems(slide, layout, bodyRect, points, settings);
        } else {
          bulletText(slide, bodyRect.left + padding, bodyRect.top + padding, bodyRect.width - padding * 2, bodyRect.height - padding * 2, points);
        }
      }
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createCompareSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'compareSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'compareSlide', data.subhead);
    const leftBox = offsetRect(layout.getRect('compareSlide.leftBox'), 0, dy);
    const rightBox = offsetRect(layout.getRect('compareSlide.rightBox'), 0, dy);
    // まじん式+: leftColor/rightColor でヘッダー色を上書き
    const leftHeaderColor = paletteHex(data.leftColor, 'main');
    const rightHeaderColor = paletteHex(data.rightColor, 'main');
    drawCompareBox(slide, layout, leftBox, data.leftTitle || '選択肢A', data.leftItems || [], settings, true, leftHeaderColor);
    drawCompareBox(slide, layout, rightBox, data.rightTitle || '選択肢B', data.rightItems || [], settings, false, rightHeaderColor);
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createProcessSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'processSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'processSlide', data.subhead);
    const area = offsetRect(layout.getRect('processSlide.area'), 0, dy);
    const steps = Array.isArray(data.steps) ? data.steps.slice(0, 4) : [];
    if (steps.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum, settings); return; }

    const processBodyBgColor = generateTintedGray(settings.primaryColor, 30, 94);
    const n = steps.length;
    let boxHPx, arrowHPx, fontSize;
    if (n <= 2) { boxHPx = 100; arrowHPx = 25; fontSize = 16; }
    else if (n === 3) { boxHPx = 80; arrowHPx = 20; fontSize = 16; }
    else { boxHPx = 65; arrowHPx = 15; fontSize = 14; }

    const processColors = generateProcessColors(settings.primaryColor, n);
    let currentY = area.top + layout.pxToPt(10);
    const boxHPt = layout.pxToPt(boxHPx), arrowHPt = layout.pxToPt(arrowHPx);
    const headerWPt = layout.pxToPt(120);
    const bodyLeft = area.left + headerWPt;
    const bodyWPt = area.width - headerWPt;

    for (let i = 0; i < n; i++) {
      const cleanText = String(steps[i] || '').replace(/^\s*\d+[\.\s]*/, '');
      txt(slide, area.left, currentY, headerWPt, boxHPt, `STEP ${i + 1}`,
        { size: fontSize, bold: true, color: CONFIG.COLORS.background_white, align: 'center', valign: 'middle', shape: ST.rect, fill: processColors[i] });
      box(slide, ST.rect, bodyLeft, currentY, bodyWPt, boxHPt, { fill: processBodyBgColor });
      txt(slide, bodyLeft + layout.pxToPt(20), currentY, bodyWPt - layout.pxToPt(40), boxHPt, cleanText,
        { size: fontSize, valign: 'middle' });
      currentY += boxHPt;
      if (i < n - 1) {
        const arrowLeft = area.left + headerWPt / 2 - layout.pxToPt(8);
        box(slide, ST.downArrow, arrowLeft, currentY, layout.pxToPt(16), arrowHPt, { fill: CONFIG.COLORS.process_arrow });
        currentY += arrowHPt;
      }
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createProcessListSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'processSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'processSlide', data.subhead);
    const area = offsetRect(layout.getRect('processSlide.area'), 0, dy);
    const steps = Array.isArray(data.steps) ? data.steps : [];
    if (steps.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum, settings); return; }
    const n = Math.max(1, steps.length);
    const topPadding = layout.pxToPt(30), bottomPadding = layout.pxToPt(10);
    const drawableHeight = area.height - topPadding - bottomPadding;
    const gapY = drawableHeight / Math.max(1, n - 1);
    const cx = area.left + layout.pxToPt(44);
    const top0 = area.top + topPadding;

    box(slide, ST.rect, cx - layout.pxToPt(1), top0 + layout.pxToPt(6), layout.pxToPt(2), gapY * (n - 1), { fill: CONFIG.COLORS.faint_gray });

    for (let i = 0; i < n; i++) {
      const cy = top0 + gapY * i;
      const sz = layout.pxToPt(28);
      txt(slide, cx - sz / 2, cy - sz / 2, sz, sz, String(i + 1),
        { size: 12, bold: true, color: CONFIG.COLORS.background_white, align: 'center', valign: 'middle', shape: ST.rect, fill: settings.primaryColor });
      let cleanText = String(steps[i] || '').replace(/^\s*\d+[\.\s]*/, '');
      txt(slide, cx + layout.pxToPt(28), cy - layout.pxToPt(16), area.width - layout.pxToPt(70), layout.pxToPt(32), cleanText,
        { size: CONFIG.FONTS.sizes.processStep, valign: 'middle' });
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createTimelineSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'timelineSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'timelineSlide', data.subhead);
    const area = offsetRect(layout.getRect('timelineSlide.area'), 0, dy);
    const milestones = Array.isArray(data.milestones) ? data.milestones : [];
    if (milestones.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum, settings); return; }

    const inner = layout.pxToPt(80), baseY = area.top + area.height * 0.50;
    const leftX = area.left + inner, rightX = area.left + area.width - inner;
    lineSeg(slide, leftX, baseY, rightX, baseY, { color: CONFIG.COLORS.faint_gray, width: 2 });
    const dotR = layout.pxToPt(10);
    const gap = (milestones.length > 1) ? (rightX - leftX) / (milestones.length - 1) : 0;
    const cardW_pt = layout.pxToPt(180);
    const vOffset = layout.pxToPt(40);
    const headerHeight = layout.pxToPt(28);
    const bodyHeight = layout.pxToPt(80);
    const timelineColors = generateTimelineCardColors(settings.primaryColor, milestones.length);

    milestones.forEach((m, i) => {
      const x = leftX + gap * i;
      const isAbove = i % 2 === 0;
      const dateText = String(m.date || '');
      const labelText = String(m.label || '');
      const cardH_pt = headerHeight + bodyHeight;
      const cardLeft = x - (cardW_pt / 2);
      const cardTop = isAbove ? (baseY - vOffset - cardH_pt) : (baseY + vOffset);

      const connectorY_start = isAbove ? (cardTop + cardH_pt) : baseY;
      const connectorY_end = isAbove ? baseY : cardTop;
      lineSeg(slide, x, connectorY_start, x, connectorY_end, { color: CONFIG.COLORS.neutral_gray, width: 1 });

      // ボディ（白背景）
      box(slide, ST.rect, cardLeft, cardTop + headerHeight, cardW_pt, bodyHeight, { fill: CONFIG.COLORS.background_white, line: CONFIG.COLORS.card_border });
      // ヘッダー（日付）
      txt(slide, cardLeft, cardTop, cardW_pt, headerHeight, dateText,
        { size: CONFIG.FONTS.sizes.body, bold: true, color: CONFIG.COLORS.background_white, align: 'center', valign: 'middle', shape: ST.rect, fill: timelineColors[i], line: CONFIG.COLORS.card_border });
      // ドット
      box(slide, ST.ellipse, x - dotR / 2, baseY - dotR / 2, dotR, dotR, { fill: timelineColors[i] });

      let bodyFontSize = CONFIG.FONTS.sizes.body;
      const textLength = labelText.length;
      if (textLength > 40) bodyFontSize = 10;
      else if (textLength > 30) bodyFontSize = 11;
      else if (textLength > 20) bodyFontSize = 12;
      txt(slide, cardLeft, cardTop + headerHeight, cardW_pt, bodyHeight, labelText,
        { size: bodyFontSize, align: 'center', valign: 'middle' });
    });
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createDiagramSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'diagramSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'diagramSlide', data.subhead);
    const lanes = Array.isArray(data.lanes) ? data.lanes : [];
    const area = offsetRect(layout.getRect('diagramSlide.lanesArea'), 0, dy);
    const px = (p) => layout.pxToPt(p);
    const { laneGap_px, lanePad_px, laneTitle_h_px, cardGap_px, cardMin_h_px, cardMax_h_px, arrow_h_px, arrowGap_px } = CONFIG.DIAGRAM;
    const n = Math.max(1, lanes.length);
    const laneW = (area.width - px(laneGap_px) * (n - 1)) / n;
    const cardBoxes = [];
    for (let j = 0; j < n; j++) {
      const lane = lanes[j] || { title: '', items: [] };
      const left = area.left + j * (laneW + px(laneGap_px));
      const laneFill = paletteHex(lane.color, 'main') || settings.primaryColor; // まじん式+: lane color
      txt(slide, left, area.top, laneW, px(laneTitle_h_px), lane.title || '',
        { size: CONFIG.FONTS.sizes.laneTitle, bold: true, color: CONFIG.COLORS.background_white, align: 'center', valign: 'middle', shape: ST.rect, fill: laneFill, line: CONFIG.COLORS.lane_border });
      const laneBodyTop = area.top + px(laneTitle_h_px), laneBodyHeight = area.height - px(laneTitle_h_px);
      box(slide, ST.rect, left, laneBodyTop, laneW, laneBodyHeight, { fill: CONFIG.COLORS.background_gray, line: CONFIG.COLORS.lane_border });
      const items = Array.isArray(lane.items) ? lane.items : [];
      const availH = laneBodyHeight - px(lanePad_px) * 2, rows = Math.max(1, items.length);
      const idealH = (availH - px(cardGap_px) * (rows - 1)) / rows;
      const cardH = Math.max(px(cardMin_h_px), Math.min(px(cardMax_h_px), idealH));
      const firstTop = laneBodyTop + px(lanePad_px) + Math.max(0, (availH - (cardH * rows + px(cardGap_px) * (rows - 1))) / 2);
      cardBoxes[j] = [];
      for (let i = 0; i < rows; i++) {
        const cardTop = firstTop + i * (cardH + px(cardGap_px));
        const cLeft = left + px(lanePad_px), cW = laneW - px(lanePad_px) * 2;
        txt(slide, cLeft, cardTop, cW, cardH, items[i] || '',
          { size: CONFIG.FONTS.sizes.body, valign: 'middle', align: 'center', shape: ST.roundRect, fill: CONFIG.COLORS.background_white, line: CONFIG.COLORS.card_border });
        cardBoxes[j][i] = { left: cLeft, top: cardTop, width: cW, height: cardH };
      }
    }
    const maxRows = Math.max(0, ...cardBoxes.map(a => a ? a.length : 0));
    for (let j = 0; j < n - 1; j++) {
      for (let i = 0; i < maxRows; i++) {
        if (cardBoxes[j] && cardBoxes[j][i] && cardBoxes[j + 1] && cardBoxes[j + 1][i]) {
          drawArrowBetweenRects(slide, cardBoxes[j][i], cardBoxes[j + 1][i], settings);
        }
      }
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createCycleSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'contentSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead);
    const area = offsetRect(layout.getRect('contentSlide.body'), 0, dy);
    const items = Array.isArray(data.items) && data.items.length === 4 ? data.items : [];
    if (items.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum, settings); return; }

    const textLengths = items.map(item => ((item.label || '').length + (item.subLabel || '').length));
    const maxLength = Math.max(...textLengths);
    const avgLength = textLengths.reduce((s, l) => s + l, 0) / textLengths.length;
    const centerX = area.left + area.width / 2;
    const centerY = area.top + area.height / 2;
    const radiusX = area.width / 3.2;
    const radiusY = area.height / 2.6;
    const maxCardW = Math.min(layout.pxToPt(220), radiusX * 0.8);
    const maxCardH = Math.min(layout.pxToPt(100), radiusY * 0.6);

    let cardW, cardH, fontSize;
    if (maxLength > 25 || avgLength > 18) { cardW = Math.min(layout.pxToPt(230), maxCardW); cardH = Math.min(layout.pxToPt(105), maxCardH); fontSize = 13; }
    else if (maxLength > 15 || avgLength > 10) { cardW = Math.min(layout.pxToPt(215), maxCardW); cardH = Math.min(layout.pxToPt(95), maxCardH); fontSize = 14; }
    else { cardW = layout.pxToPt(200); cardH = layout.pxToPt(90); fontSize = 16; }

    // 矢印（先に描いて背面に）
    const arrowRadiusX = radiusX * 0.75, arrowRadiusY = radiusY * 0.80;
    const arrowSize = layout.pxToPt(80);
    const arrowPositions = [
      { left: centerX + arrowRadiusX, top: centerY - arrowRadiusY, rotation: 90 },
      { left: centerX + arrowRadiusX, top: centerY + arrowRadiusY, rotation: 180 },
      { left: centerX - arrowRadiusX, top: centerY + arrowRadiusY, rotation: 270 },
      { left: centerX - arrowRadiusX, top: centerY - arrowRadiusY, rotation: 0 }
    ];
    arrowPositions.forEach(pos => {
      box(slide, ST.bentArrow, pos.left - arrowSize / 2, pos.top - arrowSize / 2, arrowSize, arrowSize,
        { fill: CONFIG.COLORS.ghost_gray, rotate: pos.rotation });
    });

    if (data.centerText) {
      txt(slide, centerX - layout.pxToPt(100), centerY - layout.pxToPt(50), layout.pxToPt(200), layout.pxToPt(100), data.centerText,
        { size: 20, bold: true, align: 'center', valign: 'middle', color: CONFIG.COLORS.text_primary });
    }

    const positions = [
      { x: centerX + radiusX, y: centerY },
      { x: centerX, y: centerY + radiusY },
      { x: centerX - radiusX, y: centerY },
      { x: centerX, y: centerY - radiusY }
    ];
    positions.forEach((pos, i) => {
      const cardX = pos.x - cardW / 2, cardY = pos.y - cardH / 2;
      const item = items[i] || {};
      const subLabelText = item.subLabel || `${i + 1}番目`;
      const labelText = item.label || '';
      // subLabel小さめ + label。PptxGenJSのrunで2段組
      const runs = [
        { text: subLabelText, options: { bold: true, color: hx(CONFIG.COLORS.background_white), fontSize: Math.max(10, fontSize - 2), breakLine: true } },
        { text: labelText, options: { bold: true, color: hx(CONFIG.COLORS.background_white), fontSize: fontSize } }
      ];
      try {
        slide.addText(runs, {
          x: cardX, y: cardY, w: cardW, h: cardH, shape: ST.roundRect, fill: { color: hx(settings.primaryColor) },
          fontFace: CONFIG.FONTS.family, align: 'center', valign: 'middle', margin: 2, wrap: true
        });
      } catch (e) { console.warn('cycle card error', e); }
    });

    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createCardsSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'cardsSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'cardsSlide', data.subhead);
    const area = offsetRect(layout.getRect('cardsSlide.gridArea'), 0, dy);
    const items = Array.isArray(data.items) ? data.items : [];
    const cols = Math.min(3, Math.max(2, Number(data.columns) || (items.length <= 4 ? 2 : 3)));
    const gap = layout.pxToPt(16), rows = Math.ceil(items.length / cols);
    const cardW = (area.width - gap * (cols - 1)) / cols, cardH = Math.max(layout.pxToPt(92), (area.height - gap * (rows - 1)) / rows);
    for (let idx = 0; idx < items.length; idx++) {
      const r = Math.floor(idx / cols), c = idx % cols;
      const left = area.left + c * (cardW + gap), top = area.top + r * (cardH + gap);
      const obj = items[idx];
      if (typeof obj === 'string') {
        txt(slide, left, top, cardW, cardH, obj, { size: CONFIG.FONTS.sizes.body, valign: 'middle', align: 'center', shape: ST.roundRect, fill: CONFIG.COLORS.background_gray, line: CONFIG.COLORS.card_border });
      } else {
        // まじん式+: item.icon をタイトル先頭に付与
        const title = (obj.icon ? obj.icon + ' ' : '') + String(obj.title || ''), desc = String(obj.desc || '');
        let runs;
        if (title && desc) {
          runs = [
            { text: title, options: { bold: true, breakLine: true } },
            { text: '', options: { breakLine: true } },
            { text: desc, options: {} }
          ];
        } else if (title) {
          runs = [{ text: title, options: { bold: true } }];
        } else {
          runs = [{ text: desc, options: {} }];
        }
        try {
          slide.addText(runs, {
            x: left, y: top, w: cardW, h: cardH, shape: ST.roundRect, fill: { color: hx(CONFIG.COLORS.background_gray) },
            line: { color: hx(CONFIG.COLORS.card_border), width: 1 }, fontFace: CONFIG.FONTS.family,
            fontSize: CONFIG.FONTS.sizes.body, color: hx(CONFIG.COLORS.text_primary), align: 'center', valign: 'middle', margin: 6, wrap: true
          });
        } catch (e) { console.warn('cards error', e); }
      }
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createHeaderCardsSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'cardsSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'cardsSlide', data.subhead);
    const area = offsetRect(layout.getRect('cardsSlide.gridArea'), 0, dy);
    const items = Array.isArray(data.items) ? data.items : [];
    const cols = Math.min(3, Math.max(2, Number(data.columns) || (items.length <= 4 ? 2 : 3)));
    const gap = layout.pxToPt(16), rows = Math.ceil(items.length / cols);
    const cardW = (area.width - gap * (cols - 1)) / cols, cardH = Math.max(layout.pxToPt(92), (area.height - gap * (rows - 1)) / rows);
    for (let idx = 0; idx < items.length; idx++) {
      const r = Math.floor(idx / cols), c = idx % cols;
      const left = area.left + c * (cardW + gap), top = area.top + r * (cardH + gap);
      const item = items[idx] || {};
      // まじん式+: item.icon をヘッダー先頭に、item.color でヘッダー色を上書き
      const titleText = (item.icon ? item.icon + ' ' : '') + String(item.title || ''), descText = String(item.desc || '');
      const headerFill = paletteHex(item.color, 'main') || settings.primaryColor;
      const headerHeight = layout.pxToPt(40);
      box(slide, ST.rect, left, top + headerHeight, cardW, cardH - headerHeight, { fill: CONFIG.COLORS.background_gray, line: CONFIG.COLORS.card_border });
      txt(slide, left, top, cardW, headerHeight, titleText,
        { size: CONFIG.FONTS.sizes.body, bold: true, color: CONFIG.COLORS.background_white, align: 'center', valign: 'middle', shape: ST.rect, fill: headerFill, line: CONFIG.COLORS.card_border });
      txt(slide, left + layout.pxToPt(12), top + headerHeight, cardW - layout.pxToPt(24), cardH - headerHeight, descText,
        { size: CONFIG.FONTS.sizes.body, align: 'center', valign: 'middle' });
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createTableSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'tableSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'tableSlide', data.subhead);
    const area = offsetRect(layout.getRect('tableSlide.area'), 0, dy);
    const headers = Array.isArray(data.headers) ? data.headers : [];
    const rows = Array.isArray(data.rows) ? data.rows : [];
    if (headers.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum, settings); return; }

    const tableRows = [];
    tableRows.push(headers.map(hd => ({
      text: String(hd || ''),
      options: { bold: true, color: hx(CONFIG.COLORS.text_primary), fill: hx(CONFIG.COLORS.table_header_bg), align: 'center', valign: 'middle' }
    })));
    const hlColor = paletteHex(data.accentColor, 'light') || CONFIG.COLORS.faint_gray; // まじん式+: 最終行ハイライト色
    for (let r = 0; r < rows.length; r++) {
      const isLast = data.highlightLastRow && r === rows.length - 1;
      tableRows.push(headers.map((_, c) => ({
        text: String((rows[r] || [])[c] || ''),
        options: { color: hx(CONFIG.COLORS.text_primary), bold: !!isLast, fill: hx(isLast ? hlColor : CONFIG.COLORS.background_white), align: 'center', valign: 'middle' }
      })));
    }
    try {
      slide.addTable(tableRows, {
        x: area.left, y: area.top, w: area.width, h: area.height,
        fontFace: CONFIG.FONTS.family, fontSize: CONFIG.FONTS.sizes.body,
        border: { type: 'solid', color: hx(CONFIG.COLORS.card_border), pt: 1 },
        valign: 'middle', autoPage: false
      });
    } catch (e) { console.warn('table error', e); }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createProgressSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'progressSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'progressSlide', data.subhead);
    const area = offsetRect(layout.getRect('progressSlide.area'), 0, dy);
    const items = Array.isArray(data.items) ? data.items : [];
    const n = Math.max(1, items.length);
    const cardGap = layout.pxToPt(12);
    const cardHeight = Math.max(layout.pxToPt(80), (area.height - cardGap * (n - 1)) / n);
    const cardPadding = layout.pxToPt(15);
    const barHeight = layout.pxToPt(12);
    const percentHeight = layout.pxToPt(30);
    const percentWidth = layout.pxToPt(120);

    for (let i = 0; i < n; i++) {
      const cardTop = area.top + i * (cardHeight + cardGap);
      const p = Math.max(0, Math.min(100, Number((items[i] || {}).percent || 0)));
      box(slide, ST.roundRect, area.left, cardTop, area.width, cardHeight, { fill: CONFIG.COLORS.background_white, line: CONFIG.COLORS.card_border });
      const labelHeight = layout.pxToPt(20);
      const labelWidth = area.width - percentWidth - cardPadding * 3;
      txt(slide, area.left + cardPadding, cardTop + cardPadding, labelWidth, labelHeight, String((items[i] || {}).label || ''),
        { size: CONFIG.FONTS.sizes.body, bold: true, align: 'left', valign: 'top' });
      txt(slide, area.left + area.width - percentWidth - cardPadding, cardTop + cardPadding - layout.pxToPt(2), percentWidth, percentHeight, `${p}%`,
        { size: 20, bold: true, color: settings.primaryColor, align: 'right', valign: 'top' });
      const barTop = cardTop + cardHeight - cardPadding - barHeight;
      const barWidth = area.width - cardPadding * 2;
      box(slide, ST.roundRect, area.left + cardPadding, barTop, barWidth, barHeight, { fill: CONFIG.COLORS.faint_gray });
      if (p > 0) {
        const filledBarWidth = Math.max(layout.pxToPt(6), barWidth * (p / 100));
        box(slide, ST.roundRect, area.left + cardPadding, barTop, filledBarWidth, barHeight, { fill: settings.primaryColor });
      }
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createQuoteSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'quoteSlide', data.title || '引用', settings);
    const dy = drawSubheadIfAny(slide, layout, 'quoteSlide', data.subhead);
    const baseTop = 120;
    const subheadHeight = data.subhead ? layout.pxToPt(40) : 0;
    const margin = layout.pxToPt(10);
    const area = offsetRect(layout.getRect({
      left: 40, top: baseTop + (data.subhead ? 40 : 0) + 10, width: 880, height: 320 - (data.subhead ? 40 : 0) - 10
    }), 0, dy);
    box(slide, ST.roundRect, area.left, area.top, area.width, area.height, { fill: CONFIG.COLORS.background_white, line: CONFIG.COLORS.card_border, lineWidth: 2 });
    const padding = layout.pxToPt(40);
    const textLeft = area.left + padding, textTop = area.top + padding;
    const textWidth = area.width - (padding * 2), textHeight = area.height - (padding * 2);
    const quoteTextHeight = textHeight - layout.pxToPt(30);
    txt(slide, textLeft, textTop, textWidth, quoteTextHeight, data.text || '',
      { size: 24, align: 'center', valign: 'middle', color: CONFIG.COLORS.text_primary });
    txt(slide, textLeft, textTop + quoteTextHeight, textWidth, layout.pxToPt(30), `— ${data.author || ''}`,
      { size: 16, color: CONFIG.COLORS.neutral_gray, align: 'right', valign: 'middle' });
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createKpiSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'kpiSlide', data.title || '主要指標', settings);
    const dy = drawSubheadIfAny(slide, layout, 'kpiSlide', data.subhead);
    const area = offsetRect(layout.getRect('kpiSlide.gridArea'), 0, dy);
    const items = Array.isArray(data.items) ? data.items : [];
    const cols = Math.min(4, Math.max(2, Number(data.columns) || (items.length <= 4 ? items.length : 4)));
    const gap = layout.pxToPt(16);
    const cardW = (area.width - gap * (cols - 1)) / cols, cardH = layout.pxToPt(240);
    for (let idx = 0; idx < items.length; idx++) {
      const c = idx % cols, r = Math.floor(idx / cols);
      const left = area.left + c * (cardW + gap), top = area.top + r * (cardH + gap);
      box(slide, ST.rect, left, top, cardW, cardH, { fill: CONFIG.COLORS.background_gray, line: CONFIG.COLORS.card_border });
      const item = items[idx] || {};
      const pad = layout.pxToPt(15);
      txt(slide, left + pad, top + layout.pxToPt(25), cardW - pad * 2, layout.pxToPt(35), item.label || 'KPI',
        { size: 14, color: CONFIG.COLORS.neutral_gray, valign: 'top' });
      txt(slide, left + pad, top + layout.pxToPt(80), cardW - pad * 2, layout.pxToPt(80), item.value || '0',
        { size: 32, bold: true, align: 'center', valign: 'middle' });
      let changeColor = CONFIG.COLORS.text_primary;
      if (item.status === 'bad') changeColor = '#d93025';
      if (item.status === 'good') changeColor = '#1e8e3e';
      if (item.status === 'neutral') changeColor = CONFIG.COLORS.neutral_gray;
      txt(slide, left + pad, top + layout.pxToPt(180), cardW - pad * 2, layout.pxToPt(40), item.change || '',
        { size: 14, color: changeColor, bold: true, align: 'right', valign: 'top' });
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createBulletCardsSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'contentSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead);
    const area = offsetRect(layout.getRect('contentSlide.body'), 0, dy);
    const items = Array.isArray(data.items) ? data.items.slice(0, 3) : [];
    if (items.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum, settings); return; }
    const gap = layout.pxToPt(16);
    const cardHeight = (area.height - gap * (items.length - 1)) / items.length;
    const padding = layout.pxToPt(20);
    for (let i = 0; i < items.length; i++) {
      const top = area.top + i * (cardHeight + gap);
      box(slide, ST.rect, area.left, top, area.width, cardHeight, { fill: CONFIG.COLORS.background_gray, line: CONFIG.COLORS.card_border });
      // まじん式+: item.icon をタイトル先頭に付与
      const title = (items[i].icon ? items[i].icon + ' ' : '') + String(items[i].title || ''), desc = String(items[i].desc || '');
      if (title && desc) {
        const titleFontSize = 14;
        const titleHeight = layout.pxToPt(titleFontSize + 4);
        txt(slide, area.left + padding, top + layout.pxToPt(12), area.width - padding * 2, titleHeight, title,
          { size: titleFontSize, bold: true, valign: 'top' });
        const descTop = top + layout.pxToPt(12) + titleHeight + layout.pxToPt(8);
        const descHeight = cardHeight - layout.pxToPt(12) - titleHeight - layout.pxToPt(8);
        let descFontSize = 14;
        if (desc.length > 100) descFontSize = 12;
        else if (desc.length > 80) descFontSize = 13;
        txt(slide, area.left + padding, descTop, area.width - padding * 2, descHeight, desc,
          { size: descFontSize, valign: 'middle' });
      } else {
        txt(slide, area.left + padding, top, area.width - padding * 2, cardHeight, title || desc,
          { size: 14, bold: !!title, valign: 'middle' });
      }
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createAgendaSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'processSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'processSlide', data.subhead);
    const area = offsetRect(layout.getRect('processSlide.area'), 0, dy);
    let items = Array.isArray(data.items) ? data.items : [];
    if (items.length === 0) {
      items = buildAgendaFromSlideData();
      if (items.length === 0) items = ['本日の目的', '進め方', '次のアクション'];
    }
    const n = Math.max(1, items.length);
    const topPadding = layout.pxToPt(30), bottomPadding = layout.pxToPt(10);
    const drawableHeight = area.height - topPadding - bottomPadding;
    const gapY = drawableHeight / Math.max(1, n - 1);
    const cx = area.left + layout.pxToPt(44);
    const top0 = area.top + topPadding;
    for (let i = 0; i < n; i++) {
      const cy = top0 + gapY * i;
      const sz = layout.pxToPt(28);
      txt(slide, cx - sz / 2, cy - sz / 2, sz, sz, String(i + 1),
        { size: 12, bold: true, color: CONFIG.COLORS.background_white, align: 'center', valign: 'middle', shape: ST.rect, fill: settings.primaryColor });
      let cleanText = String(items[i] || '').replace(/^\s*\d+[\.\s]*/, '');
      txt(slide, cx + layout.pxToPt(28), cy - layout.pxToPt(16), area.width - layout.pxToPt(70), layout.pxToPt(32), cleanText,
        { size: CONFIG.FONTS.sizes.processStep, valign: 'middle' });
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createFaqSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'contentSlide', data.title || 'よくあるご質問', settings);
    const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead);
    const area = offsetRect(layout.getRect('contentSlide.body'), 0, dy);
    const items = Array.isArray(data.items) ? data.items.slice(0, 4) : [];
    if (items.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum, settings); return; }

    const cardGap = layout.pxToPt(12);
    const availableHeight = area.height - cardGap * (items.length - 1);
    const cardHeight = availableHeight / items.length;
    const baseFontSize = items.length >= 4 ? 12 : 14;
    const aTintColor = generateTintedGray(settings.primaryColor, 15, 70);
    let currentY = area.top;

    items.forEach((item) => {
      box(slide, ST.roundRect, area.left, currentY, area.width, cardHeight, { fill: CONFIG.COLORS.background_gray, line: CONFIG.COLORS.card_border });
      let cardPadding, qAreaRatio, qAGap;
      if (items.length <= 2) { cardPadding = layout.pxToPt(16); qAreaRatio = 0.30; qAGap = layout.pxToPt(6); }
      else if (items.length === 3) { cardPadding = layout.pxToPt(12); qAreaRatio = 0.35; qAGap = layout.pxToPt(4); }
      else { cardPadding = layout.pxToPt(8); qAreaRatio = 0.40; qAGap = layout.pxToPt(2); }

      const availH = cardHeight - cardPadding * 2;
      const qAreaHeight = availH * qAreaRatio;
      const aAreaHeight = availH - qAreaHeight - qAGap;
      const qTop = currentY + cardPadding;

      // Q行
      const qRuns = [
        { text: 'Q. ', options: { bold: true, color: hx(settings.primaryColor) } }
      ].concat(toRuns(item.q || '', { bold: true, color: CONFIG.COLORS.text_primary }));
      try {
        slide.addText(qRuns, {
          x: area.left + cardPadding, y: qTop, w: area.width - cardPadding * 2, h: qAreaHeight,
          fontFace: CONFIG.FONTS.family, fontSize: baseFontSize, align: 'left', valign: 'top', margin: 1, wrap: true
        });
      } catch (e) { console.warn('faq q error', e); }

      // A行
      const aTop = qTop + qAreaHeight + qAGap;
      const aIndent = layout.pxToPt(16);
      const aRuns = [
        { text: 'A. ', options: { bold: true, color: hx(aTintColor) } }
      ].concat(toRuns(item.a || '', { color: CONFIG.COLORS.text_primary }));
      try {
        slide.addText(aRuns, {
          x: area.left + cardPadding + aIndent, y: aTop, w: area.width - cardPadding * 2 - aIndent, h: aAreaHeight,
          fontFace: CONFIG.FONTS.family, fontSize: baseFontSize, align: 'left', valign: 'top', margin: 1, wrap: true
        });
      } catch (e) { console.warn('faq a error', e); }

      currentY += cardHeight + cardGap;
    });
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createStatsCompareSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'compareSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'compareSlide', data.subhead);
    const area = offsetRect(layout.getRect({ left: 25, top: 130, width: 910, height: 330 }), 0, dy);
    const stats = Array.isArray(data.stats) ? data.stats : [];
    if (stats.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum, settings); return; }

    box(slide, ST.rect, area.left, area.top, area.width, area.height, { fill: CONFIG.COLORS.background_white });
    const headerHeight = layout.pxToPt(40);
    const totalContentWidth = area.width;
    const centerColWidth = totalContentWidth * 0.25;
    const sideColWidth = (totalContentWidth - centerColWidth) / 2;
    const leftValueColX = area.left;
    const centerLabelColX = leftValueColX + sideColWidth;
    const rightValueColX = centerLabelColX + centerColWidth;
    const labelColor = generateTintedGray(settings.primaryColor, 35, 70);
    const compareColors = generateCompareColors(settings.primaryColor);

    txt(slide, leftValueColX, area.top, sideColWidth, headerHeight, data.leftTitle || '',
      { size: 14, bold: true, color: CONFIG.COLORS.background_white, align: 'center', valign: 'middle', shape: ST.rect, fill: compareColors.left });
    txt(slide, rightValueColX, area.top, sideColWidth, headerHeight, data.rightTitle || '',
      { size: 14, bold: true, color: CONFIG.COLORS.background_white, align: 'center', valign: 'middle', shape: ST.rect, fill: compareColors.right });

    const contentAreaHeight = area.height - headerHeight;
    const rowHeight = contentAreaHeight / stats.length;
    let currentY = area.top + headerHeight;
    stats.forEach((stat, index) => {
      const centerY = currentY + rowHeight / 2;
      const valueHeight = layout.pxToPt(40);
      txt(slide, centerLabelColX, centerY - valueHeight / 2, centerColWidth, valueHeight, stat.label || '',
        { size: 14, align: 'center', color: labelColor, bold: true, valign: 'middle' });
      txt(slide, leftValueColX, centerY - valueHeight / 2, sideColWidth, valueHeight, stat.leftValue || '',
        { size: 22, bold: true, align: 'center', valign: 'middle' });
      txt(slide, rightValueColX, centerY - valueHeight / 2, sideColWidth, valueHeight, stat.rightValue || '',
        { size: 22, bold: true, align: 'center', valign: 'middle' });
      if (index < stats.length - 1) {
        const lineY = currentY + rowHeight;
        lineSeg(slide, area.left + layout.pxToPt(15), lineY, area.left + area.width - layout.pxToPt(15), lineY, { color: CONFIG.COLORS.faint_gray, width: 1 });
      }
      currentY += rowHeight;
    });
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createBarCompareSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'compareSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'compareSlide', data.subhead);
    const area = offsetRect(layout.getRect({ left: 40, top: 130, width: 880, height: 340 }), 0, dy);
    const stats = Array.isArray(data.stats) ? data.stats : [];
    if (stats.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum, settings); return; }

    let blockMargin, titleHeight, titleFontSize, barHeight, valueFontSize, valueWidth;
    if (stats.length <= 2) { blockMargin = layout.pxToPt(30); titleHeight = layout.pxToPt(40); titleFontSize = 18; barHeight = layout.pxToPt(20); valueFontSize = 20; valueWidth = layout.pxToPt(120); }
    else if (stats.length <= 3) { blockMargin = layout.pxToPt(25); titleHeight = layout.pxToPt(35); titleFontSize = 16; barHeight = layout.pxToPt(18); valueFontSize = 18; valueWidth = layout.pxToPt(110); }
    else { blockMargin = layout.pxToPt(20); titleHeight = layout.pxToPt(30); titleFontSize = 15; barHeight = layout.pxToPt(16); valueFontSize = 16; valueWidth = layout.pxToPt(100); }

    const totalContentHeight = area.height - (blockMargin * (stats.length - 1));
    const blockHeight = totalContentHeight / stats.length;
    let currentY = area.top;

    stats.forEach(stat => {
      const blockTop = currentY;
      const barAreaHeight = blockHeight - titleHeight;
      const barRowHeight = barAreaHeight / 2;
      txt(slide, area.left, blockTop, area.width, titleHeight, stat.label || '', { size: titleFontSize, bold: true, valign: 'bottom' });
      const asIsY = blockTop + titleHeight;
      const toBeY = asIsY + barRowHeight;
      const labelWidth = layout.pxToPt(90);
      const barWidth = Math.max(layout.pxToPt(50), area.width - labelWidth - valueWidth - layout.pxToPt(10));
      const barLeft = area.left + labelWidth;
      const val1 = parseNumericValue(stat.leftValue), val2 = parseNumericValue(stat.rightValue);
      const maxValue = Math.max(val1, val2, 1);

      txt(slide, area.left, asIsY, labelWidth, barRowHeight, '現状', { size: 12, color: CONFIG.COLORS.neutral_gray, valign: 'middle' });
      txt(slide, barLeft + barWidth, asIsY, valueWidth, barRowHeight, stat.leftValue || '', { size: valueFontSize, bold: true, align: 'right', valign: 'middle' });
      box(slide, ST.roundRect, barLeft, asIsY + barRowHeight / 2 - barHeight / 2, barWidth, barHeight, { fill: CONFIG.COLORS.faint_gray });
      const asIsFillWidth = Math.max(layout.pxToPt(2), barWidth * (val1 / maxValue));
      box(slide, ST.roundRect, barLeft, asIsY + barRowHeight / 2 - barHeight / 2, asIsFillWidth, barHeight, { fill: CONFIG.COLORS.neutral_gray });

      txt(slide, area.left, toBeY, labelWidth, barRowHeight, '導入後', { size: 12, color: settings.primaryColor, bold: true, valign: 'middle' });
      txt(slide, barLeft + barWidth, toBeY, valueWidth, barRowHeight, stat.rightValue || '', { size: valueFontSize, bold: true, color: settings.primaryColor, align: 'right', valign: 'middle' });
      box(slide, ST.roundRect, barLeft, toBeY + barRowHeight / 2 - barHeight / 2, barWidth, barHeight, { fill: generateTintedGray(settings.primaryColor, 20, 96) });
      const toBeFillWidth = Math.max(layout.pxToPt(2), barWidth * (val2 / maxValue));
      box(slide, ST.roundRect, barLeft, toBeY + barRowHeight / 2 - barHeight / 2, toBeFillWidth, barHeight, { fill: settings.primaryColor });

      currentY += blockHeight + blockMargin;
    });
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createTriangleSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'triangleSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'triangleSlide', data.subhead);
    const area = offsetRect(layout.getRect('triangleSlide.area'), 0, dy);
    const items = Array.isArray(data.items) ? data.items.slice(0, 3) : [];
    if (items.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum, settings); return; }

    const textLengths = items.map(item => (typeof item === 'string' ? item : (item.title || '') + (item.desc || '')).length);
    const maxLength = Math.max(...textLengths);
    const avgLength = textLengths.reduce((s, l) => s + l, 0) / textLengths.length;
    let cardW, cardH, fontSize;
    if (maxLength > 60 || avgLength > 40) { cardW = layout.pxToPt(340); cardH = layout.pxToPt(160); fontSize = 13; }
    else if (maxLength > 35 || avgLength > 25) { cardW = layout.pxToPt(290); cardH = layout.pxToPt(135); fontSize = 14; }
    else { cardW = layout.pxToPt(250); cardH = layout.pxToPt(115); fontSize = 15; }
    const maxCardW = (area.width - layout.pxToPt(160)) / 1.5;
    const maxCardH = (area.height - layout.pxToPt(80)) / 2;
    cardW = Math.min(cardW, maxCardW); cardH = Math.min(cardH, maxCardH);

    const positions = [
      { x: area.left + area.width / 2, y: area.top + layout.pxToPt(40) + cardH / 2 },
      { x: area.left + area.width - layout.pxToPt(80) - cardW / 2, y: area.top + area.height - cardH / 2 },
      { x: area.left + layout.pxToPt(80) + cardW / 2, y: area.top + area.height - cardH / 2 }
    ];

    // 矢印（背面）
    const arrowPadding = cardW > layout.pxToPt(300) ? layout.pxToPt(25) : layout.pxToPt(20);
    const edges = positions.map(p => ({
      right: { x: p.x + cardW / 2, y: p.y }, left: { x: p.x - cardW / 2, y: p.y },
      top: { x: p.x, y: p.y - cardH / 2 }, bottom: { x: p.x, y: p.y + cardH / 2 }
    }));
    const arrowCurves = [
      { sx: edges[0].right.x + arrowPadding, sy: edges[0].right.y, ex: edges[1].top.x, ey: edges[1].top.y - arrowPadding },
      { sx: edges[1].left.x - arrowPadding, sy: edges[1].left.y, ex: edges[2].right.x + arrowPadding, ey: edges[2].right.y },
      { sx: edges[2].top.x, sy: edges[2].top.y - arrowPadding, ex: edges[0].left.x - arrowPadding, ey: edges[0].left.y }
    ];
    arrowCurves.forEach(c => lineSeg(slide, c.sx, c.sy, c.ex, c.ey, { color: CONFIG.COLORS.ghost_gray, width: 4, arrow: true }));

    positions.forEach((pos, i) => {
      if (!items[i]) return;
      const cardX = pos.x - cardW / 2, cardY = pos.y - cardH / 2;
      const item = items[i] || {};
      const itemTitle = typeof item === 'string' ? '' : (item.title || '');
      const itemDesc = typeof item === 'string' ? item : (item.desc || '');
      let runs;
      if (typeof item === 'string' || !itemTitle) {
        const itemText = typeof item === 'string' ? item : itemDesc;
        const processed = smartFormatTriangleText(itemText);
        if (processed.isSimple) {
          runs = [{ text: processed.text, options: { bold: true, color: hx(CONFIG.COLORS.background_white) } }];
        } else {
          const lines = processed.text.split('\n');
          runs = [
            { text: lines[0] || '', options: { bold: true, color: hx(CONFIG.COLORS.background_white), fontSize: Math.max(fontSize - 1, 13), breakLine: true } },
            { text: lines.slice(1).join('\n') || '', options: { color: hx(CONFIG.COLORS.background_white), fontSize: Math.max(fontSize - 3, 11) } }
          ];
        }
      } else {
        runs = [
          { text: itemTitle, options: { bold: true, color: hx(CONFIG.COLORS.background_white), fontSize: Math.max(fontSize - 1, 13), breakLine: !!itemDesc } }
        ];
        if (itemDesc) runs.push({ text: itemDesc, options: { color: hx(CONFIG.COLORS.background_white), fontSize: Math.max(fontSize - 3, 11) } });
      }
      try {
        slide.addText(runs, {
          x: cardX, y: cardY, w: cardW, h: cardH, shape: ST.roundRect,
          fill: { color: hx(settings.primaryColor) }, line: { color: hx(CONFIG.COLORS.card_border), width: 1 },
          fontFace: CONFIG.FONTS.family, fontSize: fontSize, align: 'center', valign: 'middle', margin: 4, wrap: true
        });
      } catch (e) { console.warn('triangle card error', e); }
    });
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createPyramidSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'pyramidSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'pyramidSlide', data.subhead);
    const area = offsetRect(layout.getRect('pyramidSlide.pyramidArea'), 0, dy);
    const levels = Array.isArray(data.levels) ? data.levels.slice(0, 4) : [];
    if (levels.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum, settings); return; }

    const levelHeight = layout.pxToPt(70), levelGap = layout.pxToPt(2);
    const totalHeight = (levelHeight * levels.length) + (levelGap * (levels.length - 1));
    const startY = area.top + (area.height - totalHeight) / 2;
    const pyramidWidth = layout.pxToPt(480);
    const textColumnWidth = layout.pxToPt(400);
    const gap = layout.pxToPt(30);
    const pyramidLeft = area.left;
    const textColumnLeft = pyramidLeft + pyramidWidth + gap;
    const pyramidColors = generatePyramidColors(settings.primaryColor, levels.length);
    const baseWidth = pyramidWidth;
    const widthIncrement = baseWidth / levels.length;
    const centerX = pyramidLeft + pyramidWidth / 2;

    levels.forEach((level, index) => {
      const levelWidth = baseWidth - (widthIncrement * (levels.length - 1 - index));
      const levelX = centerX - levelWidth / 2;
      const levelY = startY + index * (levelHeight + levelGap);
      txt(slide, levelX, levelY, levelWidth, levelHeight, level.title || `レベル${index + 1}`,
        { size: CONFIG.FONTS.sizes.body, bold: true, color: CONFIG.COLORS.background_white, align: 'center', valign: 'middle', shape: ST.roundRect, fill: pyramidColors[index] });
      const connectionStartX = levelX + levelWidth;
      if (textColumnLeft > connectionStartX) {
        const connectionY = levelY + levelHeight / 2;
        lineSeg(slide, connectionStartX, connectionY, textColumnLeft, connectionY, { color: '#D0D7DE', width: 1.5 });
      }
      const levelDesc = level.description || '';
      let formattedText;
      if (levelDesc.includes('•') || levelDesc.includes('・')) formattedText = levelDesc;
      else if (levelDesc.includes('\n')) {
        const lines = levelDesc.split('\n').filter(l => l.trim()).slice(0, 2);
        formattedText = lines.map(l => `• ${l.trim()}`).join('\n');
      } else formattedText = levelDesc;
      txt(slide, textColumnLeft, levelY, textColumnWidth, levelHeight, formattedText,
        { size: CONFIG.FONTS.sizes.body - 1, align: 'left', valign: 'middle', color: CONFIG.COLORS.text_primary });
    });
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function drawFlowRow(slide, flow, area, settings, layout, maxStepsPerRow) {
    if (!flow || !flow.steps || !Array.isArray(flow.steps)) return;
    const steps = flow.steps.filter(step => step && String(step).trim());
    if (steps.length === 0) return;
    const actualSteps = maxStepsPerRow || steps.length;
    const baseArrowSpace = layout.pxToPt(25);
    const arrowSpace = Math.max(baseArrowSpace, area.width * 0.04);
    const totalArrowSpace = (actualSteps - 1) * arrowSpace;
    const cardW = (area.width - totalArrowSpace) / actualSteps;
    const cardH = area.height;
    const arrowHeight = Math.min(cardH * 0.3, layout.pxToPt(40));
    const arrowWidth = arrowSpace;
    steps.forEach((step, index) => {
      const cardX = area.left + index * (cardW + arrowSpace);
      const stepText = String(step || '').trim() || 'ステップ';
      txt(slide, cardX, area.top, cardW, cardH, stepText,
        { size: CONFIG.FONTS.sizes.body, align: 'center', valign: 'middle', shape: ST.roundRect, fill: CONFIG.COLORS.background_gray, line: CONFIG.COLORS.card_border });
      if (index < steps.length - 1) {
        const arrowStartX = cardX + cardW;
        const arrowTop = area.top + cardH / 2 - arrowHeight / 2;
        box(slide, ST.rightArrow, arrowStartX, arrowTop, arrowWidth, arrowHeight, { fill: settings.primaryColor });
      }
    });
  }

  function createFlowChartSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'flowChartSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'flowChartSlide', data.subhead);
    const flows = Array.isArray(data.flows) ? data.flows : [{ steps: data.steps || [] }];
    let isDouble = flows.length > 1;
    let upperFlow, lowerFlow, maxStepsPerRow;
    if (!isDouble && flows[0] && flows[0].steps && flows[0].steps.length >= 5) {
      isDouble = true;
      const allSteps = flows[0].steps;
      const midPoint = Math.ceil(allSteps.length / 2);
      upperFlow = { steps: allSteps.slice(0, midPoint) };
      lowerFlow = { steps: allSteps.slice(midPoint) };
      maxStepsPerRow = midPoint;
    } else {
      upperFlow = flows[0];
      lowerFlow = flows.length > 1 ? flows[1] : null;
      maxStepsPerRow = Math.max((upperFlow && upperFlow.steps ? upperFlow.steps.length : 0), (lowerFlow && lowerFlow.steps ? lowerFlow.steps.length : 0));
    }
    if (isDouble) {
      const upperArea = offsetRect(layout.getRect('flowChartSlide.upperRow'), 0, dy);
      const lowerArea = offsetRect(layout.getRect('flowChartSlide.lowerRow'), 0, dy);
      drawFlowRow(slide, upperFlow, upperArea, settings, layout, maxStepsPerRow);
      if (lowerFlow && lowerFlow.steps && lowerFlow.steps.length > 0) drawFlowRow(slide, lowerFlow, lowerArea, settings, layout, maxStepsPerRow);
    } else {
      const singleArea = offsetRect(layout.getRect('flowChartSlide.singleRow'), 0, dy);
      drawFlowRow(slide, flows[0], singleArea, settings, layout);
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createStepUpSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'stepUpSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'stepUpSlide', data.subhead);
    const area = offsetRect(layout.getRect('stepUpSlide.stepArea'), 0, dy);
    const items = Array.isArray(data.items) ? data.items : [];
    if (items.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum, settings); return; }
    const numSteps = Math.min(5, items.length);
    const headerHeight = layout.pxToPt(40);
    const maxHeight = area.height * 0.95;
    let minHeightRatio;
    if (numSteps <= 2) minHeightRatio = 0.70;
    else if (numSteps === 3) minHeightRatio = 0.60;
    else minHeightRatio = 0.50;
    const minHeight = maxHeight * minHeightRatio;
    const cardW = area.width / numSteps;
    const stepUpColors = generateStepUpColors(settings.primaryColor, numSteps);

    for (let idx = 0; idx < numSteps; idx++) {
      const item = items[idx] || {};
      const titleText = String(item.title || `STEP ${idx + 1}`);
      const descText = String(item.desc || '');
      const heightRatio = (idx / Math.max(1, numSteps - 1));
      const cardH = minHeight + (maxHeight - minHeight) * heightRatio;
      const left = area.left + idx * cardW;
      const top = area.top + area.height - cardH;
      box(slide, ST.rect, left, top + headerHeight, cardW, cardH - headerHeight, { fill: CONFIG.COLORS.background_gray, line: CONFIG.COLORS.card_border });
      txt(slide, left, top, cardW, headerHeight, titleText,
        { size: CONFIG.FONTS.sizes.body, bold: true, color: CONFIG.COLORS.background_white, align: 'center', valign: 'middle', shape: ST.rect, fill: stepUpColors[idx], line: CONFIG.COLORS.card_border, shrink: true });
      txt(slide, left + layout.pxToPt(8), top + headerHeight, cardW - layout.pxToPt(16), cardH - headerHeight, descText,
        { size: CONFIG.FONTS.sizes.body, align: 'center', valign: 'middle', shrink: true });
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  function createImageTextSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'imageTextSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'imageTextSlide', data.subhead);
    const imageUrl = data.image || '';
    const imageCaption = data.imageCaption || '';
    const points = Array.isArray(data.points) ? data.points : [];
    const imagePosition = data.imagePosition === 'right' ? 'right' : 'left';
    const padding = layout.pxToPt(20);

    if (imagePosition === 'left') {
      const imageArea = offsetRect(layout.getRect('imageTextSlide.leftImage'), 0, dy);
      const textArea = offsetRect(layout.getRect('imageTextSlide.rightText'), 0, dy);
      if (imageUrl) renderSingleImageInArea(slide, layout, imageArea, imageUrl, imageCaption, 'left');
      if (points.length > 0) {
        createContentCushion(slide, textArea, settings, layout);
        bulletText(slide, textArea.left + padding, textArea.top + padding, textArea.width - padding * 2, textArea.height - padding * 2, points);
      }
    } else {
      const textArea = offsetRect(layout.getRect('imageTextSlide.leftText'), 0, dy);
      const imageArea = offsetRect(layout.getRect('imageTextSlide.rightImage'), 0, dy);
      if (points.length > 0) {
        createContentCushion(slide, textArea, settings, layout);
        bulletText(slide, textArea.left + padding, textArea.top + padding, textArea.width - padding * 2, textArea.height - padding * 2, points);
      }
      if (imageUrl) renderSingleImageInArea(slide, layout, imageArea, imageUrl, imageCaption, 'right');
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  // ========================================
  // まじん式+ : 新規スライドタイプ
  // ========================================

  // callout（単一の強調ボックス）
  function createCalloutSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'contentSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead);
    const name = String(data.accentColor || 'blue').toLowerCase();
    const main = paletteHex(name, 'main') || settings.primaryColor;
    const light = paletteHex(name, 'light') || CONFIG.COLORS.background_gray;
    const dark = paletteHex(name, 'dark') || main;
    const area = offsetRect(layout.getRect({ left: 50, top: 150, width: 860, height: 300 }), 0, dy);

    box(slide, ST.roundRect, area.left, area.top, area.width, area.height, { fill: light, line: main, lineWidth: 1.5 });

    const iconW = layout.pxToPt(130);
    txt(slide, area.left + layout.pxToPt(20), area.top, iconW, area.height, data.icon || '💡',
      { size: 54, align: 'center', valign: 'middle' });

    const tx = area.left + layout.pxToPt(20) + iconW + layout.pxToPt(10);
    const tw = area.width - (tx - area.left) - layout.pxToPt(30);
    txt(slide, tx, area.top + layout.pxToPt(34), tw, layout.pxToPt(96), data.headline || '',
      { size: 22, bold: true, color: dark, valign: 'top', shrink: true });
    if (data.body) {
      txt(slide, tx, area.top + layout.pxToPt(132), tw, area.height - layout.pxToPt(150), data.body,
        { size: 14, color: CONFIG.COLORS.text_primary, valign: 'top' });
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  // calloutGrid（複数カラム並列）
  function createCalloutGridSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'cardsSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'cardsSlide', data.subhead);
    const area = offsetRect(layout.getRect('cardsSlide.gridArea'), 0, dy);
    const items = Array.isArray(data.items) ? data.items.slice(0, 4) : [];
    if (items.length === 0) { drawBottomBarAndFooter(slide, layout, pageNum, settings); return; }
    const cols = Math.min(3, Math.max(2, Number(data.columns) || (items.length <= 2 ? 2 : 3)));
    const gap = layout.pxToPt(18), rows = Math.ceil(items.length / cols);
    const cardW = (area.width - gap * (cols - 1)) / cols;
    const cardH = (area.height - gap * (rows - 1)) / rows;

    items.forEach((item, idx) => {
      const r = Math.floor(idx / cols), c = idx % cols;
      const left = area.left + c * (cardW + gap), top = area.top + r * (cardH + gap);
      const name = String(item.color || 'blue').toLowerCase();
      const main = paletteHex(name, 'main') || settings.primaryColor;
      const light = paletteHex(name, 'light') || CONFIG.COLORS.background_gray;
      const dark = paletteHex(name, 'dark') || main;
      box(slide, ST.roundRect, left, top, cardW, cardH, { fill: light, line: main, lineWidth: 1.25 });
      const iconH = layout.pxToPt(56);
      txt(slide, left, top + layout.pxToPt(12), cardW, iconH, item.icon || '◆',
        { size: 36, align: 'center', valign: 'middle' });
      const pad = layout.pxToPt(14);
      txt(slide, left + pad, top + iconH + layout.pxToPt(14), cardW - pad * 2, layout.pxToPt(30), item.headline || '',
        { size: 16, bold: true, color: dark, align: 'center', valign: 'top', shrink: true });
      txt(slide, left + pad, top + iconH + layout.pxToPt(48), cardW - pad * 2, cardH - iconH - layout.pxToPt(58), item.body || '',
        { size: 12, color: CONFIG.COLORS.text_primary, align: 'center', valign: 'top', shrink: true });
    });
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  // iconBanner（全面の印象づけバナー）
  function createIconBannerSlide(slide, data, layout, pageNum, settings) {
    const name = String(data.accentColor || 'skyblue').toLowerCase();
    const main = paletteHex(name, 'main') || settings.primaryColor;
    // 全面塗り（マスター背景も覆う）
    box(slide, ST.rect, 0, 0, layout.pageW_pt, layout.pageH_pt, { fill: main });
    const W = layout.pageW_pt, H = layout.pageH_pt;
    txt(slide, 0, H * 0.16, W, layout.pxToPt(130), data.icon || '✨',
      { size: 96, align: 'center', valign: 'middle' });
    txt(slide, layout.pxToPt(50), H * 0.50, W - layout.pxToPt(100), layout.pxToPt(90), data.message || '',
      { size: 34, bold: true, color: CONFIG.COLORS.background_white, align: 'center', valign: 'middle', shrink: true });
    if (data.submessage) {
      txt(slide, layout.pxToPt(70), H * 0.72, W - layout.pxToPt(140), layout.pxToPt(60), data.submessage,
        { size: 16, color: CONFIG.COLORS.background_white, align: 'center', valign: 'top', shrink: true });
    }
  }

  // mermaid（事前レンダリング済みPNGを配置、無ければ等幅テキスト）
  function createMermaidSlide(slide, data, layout, pageNum, settings) {
    setMainSlideBackground(slide, layout);
    drawStandardTitleHeader(slide, layout, 'contentSlide', data.title, settings);
    const dy = drawSubheadIfAny(slide, layout, 'contentSlide', data.subhead);
    const hasCaption = !!data.caption;
    const area = offsetRect(layout.getRect({ left: 40, top: 130, width: 880, height: hasCaption ? 300 : 330 }), 0, dy);
    const png = data.__mermaidPng;
    if (png && png.data) {
      const ar = (png.w && png.h) ? (png.w / png.h) : 1.6;
      let w = area.width, h = w / ar;
      if (h > area.height) { h = area.height; w = h * ar; }
      imgData(slide, area.left + (area.width - w) / 2, area.top + (area.height - h) / 2, w, h, png.data);
    } else {
      // フォールバック: コードを等幅テキストで表示
      const pad = layout.pxToPt(16);
      box(slide, ST.roundRect, area.left, area.top, area.width, area.height, { fill: CONFIG.COLORS.background_gray, line: CONFIG.COLORS.card_border });
      txt(slide, area.left + pad, area.top + pad, area.width - pad * 2, area.height - pad * 2, String(data.code || ''),
        { size: 11, valign: 'top', align: 'left', fontFace: 'Consolas', color: CONFIG.COLORS.text_primary });
    }
    if (hasCaption) {
      txt(slide, area.left, area.top + area.height + layout.pxToPt(6), area.width, layout.pxToPt(28), data.caption,
        { size: CONFIG.FONTS.sizes.small, color: CONFIG.COLORS.neutral_gray, align: 'center', valign: 'top' });
    }
    drawBottomBarAndFooter(slide, layout, pageNum, settings);
  }

  // ========================================
  // ジェネレーター・ディスパッチ表
  // ========================================
  const slideGenerators = {
    title: createTitleSlide,
    section: createSectionSlide,
    content: createContentSlide,
    agenda: createAgendaSlide,
    compare: createCompareSlide,
    process: createProcessSlide,
    processList: createProcessListSlide,
    timeline: createTimelineSlide,
    diagram: createDiagramSlide,
    cycle: createCycleSlide,
    cards: createCardsSlide,
    headerCards: createHeaderCardsSlide,
    table: createTableSlide,
    progress: createProgressSlide,
    quote: createQuoteSlide,
    kpi: createKpiSlide,
    closing: createClosingSlide,
    bulletCards: createBulletCardsSlide,
    faq: createFaqSlide,
    statsCompare: createStatsCompareSlide,
    barCompare: createBarCompareSlide,
    triangle: createTriangleSlide,
    pyramid: createPyramidSlide,
    flowChart: createFlowChartSlide,
    stepUp: createStepUpSlide,
    imageText: createImageTextSlide,
    callout: createCalloutSlide,
    calloutGrid: createCalloutGridSlide,
    iconBanner: createIconBannerSlide,
    mermaid: createMermaidSlide
  };

  // ========================================
  // ファイル名生成（元 createPresentation の命名規則）
  // ========================================
  function buildFileName(slideData, settings) {
    const rawTitle = (slideData[0] && slideData[0].type === 'title' ? String(slideData[0].title || '') : 'Presentation');
    const singleLineTitle = rawTitle.replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();
    let finalName;
    if (settings.showDateColumn) {
      const d = new Date();
      const dateStr = `${d.getFullYear()}.${String(d.getMonth() + 1).padStart(2, '0')}.${String(d.getDate()).padStart(2, '0')}`;
      finalName = singleLineTitle ? (singleLineTitle + ' ' + dateStr) : ('Presentation ' + dateStr);
    } else {
      finalName = singleLineTitle || 'Presentation';
    }
    // ファイル名に使えない文字を除去
    finalName = finalName.replace(/[\\/:*?"<>|]/g, '_');
    return finalName + '.pptx';
  }

  // ========================================
  // まじん式+ : Mermaid 描画（SVG→PNG data URI）
  // ========================================
  // SVG の実寸(px)を viewBox / width・height 属性から推定（無ければ 800×600）
  function svgIntrinsicSize(svgStr) {
    let w = 0, h = 0;
    const vb = svgStr.match(/viewBox\s*=\s*"([\d.\-\s,]+)"/i);
    if (vb) {
      const p = vb[1].trim().split(/[\s,]+/).map(Number);
      if (p.length === 4 && p[2] > 0 && p[3] > 0) { w = p[2]; h = p[3]; }
    }
    if (!w || !h) {
      const wm = svgStr.match(/\bwidth\s*=\s*"([\d.]+)(?:px)?"/i);
      const hm = svgStr.match(/\bheight\s*=\s*"([\d.]+)(?:px)?"/i);
      if (wm) w = parseFloat(wm[1]);
      if (hm) h = parseFloat(hm[1]);
    }
    return { w: w || 800, h: h || 600 };
  }

  // ルート <svg> の width/height を実寸(px)へ上書き（width="100%" 等で縮小描画されるのを防ぐ）
  function svgWithExplicitSize(svgStr, w, h) {
    return svgStr.replace(/<svg([^>]*)>/i, function (m, attrs) {
      attrs = attrs.replace(/\swidth\s*=\s*"[^"]*"/i, '').replace(/\sheight\s*=\s*"[^"]*"/i, '');
      return '<svg' + attrs + ' width="' + w + '" height="' + h + '">';
    });
  }

  function svgToPng(svgStr) {
    return new Promise(function (resolve) {
      try {
        if (typeof document === 'undefined' || typeof Image === 'undefined') { resolve(null); return; }
        const sz = svgIntrinsicSize(svgStr);
        const blob = new Blob([svgWithExplicitSize(svgStr, sz.w, sz.h)], { type: 'image/svg+xml;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const image = new Image();
        image.onload = function () {
          try {
            const w = image.width || sz.w, h = image.height || sz.h;
            // 高DPIラスタライズ（従来2→6 = 線形3倍・面積約9倍）。巨大化防止に長辺8000pxで上限。
            const MAX_SIDE = 8000;
            let scale = 6;
            scale = Math.min(scale, MAX_SIDE / Math.max(w, h));
            scale = Math.max(scale, 2); // 念のため下限（従来比で劣化させない）
            const canvas = document.createElement('canvas');
            canvas.width = Math.max(1, Math.round(w * scale));
            canvas.height = Math.max(1, Math.round(h * scale));
            const ctx = canvas.getContext('2d');
            if (ctx.imageSmoothingEnabled !== undefined) { ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high'; }
            ctx.fillStyle = '#FFFFFF';
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            ctx.setTransform(scale, 0, 0, scale, 0, 0);
            ctx.drawImage(image, 0, 0, w, h);
            URL.revokeObjectURL(url);
            // w/h は論理サイズ（アスペクト比用）を返す＝配置サイズは不変、画素密度のみ向上
            resolve({ data: canvas.toDataURL('image/png'), w: w, h: h });
          } catch (e) { URL.revokeObjectURL(url); resolve(null); }
        };
        image.onerror = function () { URL.revokeObjectURL(url); resolve(null); };
        image.src = url;
      } catch (e) { resolve(null); }
    });
  }

  async function renderMermaidToPng(code, theme) {
    if (typeof mermaid === 'undefined' || !code) return null;
    try {
      mermaid.initialize({
        startOnLoad: false,
        theme: theme || 'default',
        securityLevel: 'loose',
        flowchart: { htmlLabels: false },
        fontFamily: CONFIG.FONTS.family
      });
      const id = 'mmd' + Math.random().toString(36).slice(2);
      const res = await mermaid.render(id, code);
      const svg = res && res.svg ? res.svg : res;
      if (!svg) return null;
      return await svgToPng(svg);
    } catch (e) {
      console.warn('mermaid render failed:', e);
      return null;
    }
  }

  // ========================================
  // 公開API
  // ========================================
  window.generatePptx = async function (slideDataString, settings) {
    if (typeof PptxGenJS === 'undefined') {
      throw new Error('PptxGenJS が読み込まれていません（CDN取得失敗の可能性）。');
    }
    settings = settings || {};
    __ACCENT_MAP = (settings.accentMap && typeof settings.accentMap === 'object') ? settings.accentMap : null;
    __COLOR_MODE = (settings.colorMode === 'gradient') ? 'gradient' : 'accent';
    const slideData = JSON.parse(slideDataString);
    if (!Array.isArray(slideData)) throw new Error('slideData は配列である必要があります。');

    const pptx = new PptxGenJS();
    PPTX = pptx;

    // 出力形式（テンプレ判定結果）に応じてキャンバス寸法とコンテンツ中央オフセットを決定。
    // 16:9 → 10×5.625in / A4横 → 10.83×7.5in（PowerPoint「A4用紙」既定）。
    const tplFormat = (settings.template && settings.template.meta && settings.template.meta.format) || '16:9';
    let OUT_W = 10, OUT_H = 5.625;
    if (tplFormat === 'a4') { OUT_W = 10.83; OUT_H = 7.5; }
    pptx.defineLayout({ name: 'CANVAS', width: OUT_W, height: OUT_H });
    pptx.layout = 'CANVAS';
    // コンテンツは常に 10×5.625 で作図し、拡大ページでは中央寄せ（A4: dx≈0.415, dy≈0.9375 / 16:9: 0,0）
    __CONTENT_DX = (OUT_W - 10) / 2;
    __CONTENT_DY = (OUT_H - 5.625) / 2;

    // CONFIG 初期化
    updateDynamicColors(settings);
    CONFIG.COLORS.primary_color = settings.primaryColor || CONFIG.COLORS.primary_color;
    CONFIG.COLORS.slide_bg = settings.backgroundColor || '#FFFFFF';
    CONFIG.FOOTER_TEXT = settings.footerText || '';
    CONFIG.FONTS.family = settings.fontFamily || CONFIG.FONTS.family;
    CONFIG.LOGOS.header = settings.headerLogoUrl || '';
    CONFIG.LOGOS.closing = settings.closingLogoUrl || '';
    CONFIG.BACKGROUND_IMAGES.title = settings.titleBgUrl || '';
    CONFIG.BACKGROUND_IMAGES.closing = settings.closingBgUrl || '';
    CONFIG.BACKGROUND_IMAGES.section = settings.sectionBgUrl || '';
    CONFIG.BACKGROUND_IMAGES.main = settings.mainBgUrl || '';

    // テンプレート解析駆動（template-analyzer.js の解析オブジェクト）。
    // type/templateLayout → テンプレ実レイアウト枠＋装飾マスターに配置する。
    __templateActive = false;
    __TPL = { titleRect: null, bodyRect: null, titleSlideTitleRect: null, category: null };
    const tpl = settings.template;
    const templateUsable = !!(tpl && tpl.meta && tplFormat !== 'unsupported'
      && Array.isArray(tpl.layoutCatalog) && tpl.layoutCatalog.length);
    // テンプレ実寸 → 出力キャンバスへの正規化スケール
    const tplSlideW = (tpl && tpl.meta && tpl.meta.slideSize && tpl.meta.slideSize.wIn) ? tpl.meta.slideSize.wIn : OUT_W;
    const tplScale = OUT_W / tplSlideW;
    const __definedMasters = {};
    // レイアウトカタログ1件 → 装飾マスター（遅延定義）。マスター名を返す。
    function ensureTemplateMaster(entry) {
      if (!entry) return null;
      const name = 'TPL_' + entry.id;
      if (__definedMasters[name]) return name;
      try {
        pptx.defineSlideMaster({
          title: name,
          background: entry.background || { color: 'FFFFFF' },
          objects: scaleObjects(entry.objects || [], tplScale)
        });
        __definedMasters[name] = true;
        return name;
      } catch (e) { console.warn('defineSlideMaster failed:', e); return null; }
    }
    if (templateUsable) {
      __templateActive = true;
      // テンプレ実枠へ配置するため中央オフセットは無効化（枠は既に全面座標）
      __CONTENT_DX = 0; __CONTENT_DY = 0;
    }
    const TITLE_FAMILY = { title: 1, section: 1, closing: 1 };
    const ICON_HERO = { callout: 1, iconBanner: 1 }; // icon をタイトルに付けない（本体で描画）
    const basePrimary = settings.primaryColor;

    // まじん式+: mermaid を事前レンダリング（非同期）
    for (const data of slideData) {
      if (data && data.type === 'mermaid' && data.code) {
        try { data.__mermaidPng = await renderMermaidToPng(data.code, data.theme); }
        catch (e) { data.__mermaidPng = null; }
      }
    }

    __SLIDE_DATA_FOR_AGENDA = slideData;
    __SECTION_COUNTER = 0;
    const layout = createLayoutManager();
    let pageCounter = 0;

    for (const data of slideData) {
      try {
        const generator = slideGenerators[data.type];
        if (data.type !== 'title' && data.type !== 'closing') pageCounter++;
        if (generator) {
          // まじん式+: accentColor によるスライド単位の再色
          const eff = accentMain(data.accentColor, basePrimary);
          updateDynamicColors({ primaryColor: eff });
          CONFIG.COLORS.primary_color = eff;
          const slideSettings = (eff === basePrimary) ? settings : Object.assign({}, settings, { primaryColor: eff });
          // まじん式+: タイトル絵文字
          __CURRENT_ICON = ICON_HERO[data.type] ? '' : (data.icon || '');

          // テンプレ駆動: type/templateLayout から実レイアウトを解決し、装飾マスター＋枠を適用
          let masterName = null;
          let curEntry = null;
          let figureReserve = null;
          if (__templateActive) {
            curEntry = resolveTemplateLayout(data.type, data.templateLayout, tpl);
            masterName = ensureTemplateMaster(curEntry);
            const tFrame = curEntry ? scaleRect(curEntry.titleFrame, tplScale) : null;
            let bFrame = curEntry ? scaleRect(curEntry.bodyFrame, tplScale) : null;
            // 図枠の無いレイアウトで data.figure 指定時は、本文枠を左右分割して右に図ダミーを予約
            const hasFigSlots = !!(curEntry && Array.isArray(curEntry.figureFrames) && curEntry.figureFrames.length);
            if (data.figure && !hasFigSlots && bFrame) {
              const gap = 0.2, figW = bFrame.width * 0.4, bodyW = bFrame.width - figW - gap;
              figureReserve = { left: bFrame.left + bodyW + gap, top: bFrame.top, width: figW, height: bFrame.height };
              bFrame = { left: bFrame.left, top: bFrame.top, width: bodyW, height: bFrame.height };
            }
            __TPL.titleRect = tFrame;            // 本文系タイトル枠（drawStandardTitleHeader）
            __TPL.bodyRect = bFrame;             // 本文枠（必要に応じ左に縮小）
            __TPL.titleSlideTitleRect = tFrame;  // 表紙の大タイトル枠（createTitleSlide）
            __TPL.category = curEntry ? curEntry.category : null; // 構造スライドの座標忠実用
          }
          const rawSlide = masterName ? pptx.addSlide({ masterName: masterName }) : pptx.addSlide();
          // 拡大ページ（A4等）ではコンテンツ描画を中央へ一括オフセット（16:9 は dx=dy=0 で素通し）
          const slide = makeOffsetSlide(rawSlide, __CONTENT_DX, __CONTENT_DY);
          generator(slide, data, layout, pageCounter, slideSettings);

          // 図/表プレースホルダ: テンプレ実レイアウトの図枠 or 予約枠にダミーを配置（差し込み位置を明示）
          drawFigurePlaceholders(slide, data, curEntry, tplScale, figureReserve);
          if (data.notes) {
            try { slide.addNotes(cleanSpeakerNotes(data.notes)); } catch (e) { /* skip */ }
          }
        }
      } catch (e) {
        console.error('スライド生成エラー (type=' + (data && data.type) + '):', e);
      } finally {
        __CURRENT_ICON = '';
      }
    }

    const fileName = buildFileName(slideData, settings);
    await pptx.writeFile({ fileName: fileName });
    return fileName;
  };

})();
