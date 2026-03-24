/**
 * ShapR Tabler Icons — PowerPoint Add-in
 *
 * Architecture:
 *  - Fetches Tabler outline sprite SVG (one request, all icons)
 *  - Parses <symbol> elements to build icon registry
 *  - Virtual scroll renders only visible cards (~100 at a time)
 *  - Debounced search filters the registry
 *  - Office.js inserts coloured SVG into active slide
 */

/* ═══════════════════════════════════════════════════════════════════
   CONSTANTS & CONFIG
═══════════════════════════════════════════════════════════════════ */

const SPRITE_URLS = [
  'https://cdn.jsdelivr.net/npm/@tabler/icons@latest/tabler-sprite.svg',
  'https://unpkg.com/@tabler/icons@latest/tabler-sprite.svg',
];

/** px dimensions matched to S / M / L labels */
const SIZE_MAP = { '32': 32, '64': 64, '96': 96 };

/** Fallback colours used when Office theme is unavailable */
const DEFAULT_COLORS = [
  { hex: '#000000', name: 'Black' },
  { hex: '#343A40', name: 'Dark Gray' },
  { hex: '#868E96', name: 'Gray' },
  { hex: '#FFFFFF', name: 'White' },
  { hex: '#1971C2', name: 'Blue' },
  { hex: '#C92A2A', name: 'Red' },
];

/** Approx card height including gap — used for virtual scroll row math */
const CARD_SIZE   = 72;
const CARD_GAP    = 6;
const GRID_PADDING = 10;

/** How many extra rows to render above/below viewport (buffer) */
const OVERSCAN_ROWS = 3;

/* ═══════════════════════════════════════════════════════════════════
   STATE
═══════════════════════════════════════════════════════════════════ */

const state = {
  /** Full list of parsed icon objects: { name, viewBox, pathData } */
  allIcons: [],

  /** Currently displayed subset (after search filter) */
  filteredIcons: [],

  /** Hex string of the active colour swatch */
  selectedColor: '#0B2B45',

  /** Insertion size in px */
  selectedSize: 64,

  /** Whether Office.js is available */
  officeAvailable: false,

  /** Virtual scroll: number of columns in the grid */
  columns: 4,

  /** Scroll position cache */
  lastScrollTop: 0,

  /** Debounce timer handle */
  searchTimer: null,

  /** Toast auto-hide timer */
  toastTimer: null,

  /** Whether icons have loaded */
  loaded: false,
};

/* ═══════════════════════════════════════════════════════════════════
   ELEMENT REFERENCES
═══════════════════════════════════════════════════════════════════ */

const els = {
  officeWarning: () => document.getElementById('office-warning'),
  loadingState:  () => document.getElementById('loading-state'),
  errorState:    () => document.getElementById('error-state'),
  errorDetail:   () => document.getElementById('error-detail'),
  emptyState:    () => document.getElementById('empty-state'),
  gridContainer: () => document.getElementById('grid-container'),
  gridSpacer:    () => document.getElementById('grid-spacer'),
  gridViewport:  () => document.getElementById('grid-viewport'),
  searchInput:   () => document.getElementById('search-input'),
  searchClear:   () => document.getElementById('search-clear'),
  iconCount:     () => document.getElementById('icon-count'),
  resultCount:   () => document.getElementById('result-count'),
  sizeDisplay:   () => document.getElementById('size-display'),
  retryBtn:      () => document.getElementById('retry-btn'),
  toast:         () => document.getElementById('toast'),
  colorSwatches: () => document.getElementById('color-swatches'),
  colorPicker:   () => document.getElementById('color-picker'),
  hexInput:      () => document.getElementById('hex-input'),
};

/* ═══════════════════════════════════════════════════════════════════
   OFFICE.JS INITIALISATION
═══════════════════════════════════════════════════════════════════ */

function initOffice() {
  if (typeof Office === 'undefined' || !Office.onReady) {
    // Running in a plain browser — show the warning, continue anyway
    // so developers can test the UI without PowerPoint.
    console.warn('[ShapR] Office.js not available — UI-only mode.');
    state.officeAvailable = false;
    els.officeWarning().classList.remove('hidden');
    startApp();
    return;
  }

  Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
      state.officeAvailable = true;
    } else if (info.host !== null) {
      // Loaded inside an Office host that's NOT PowerPoint
      state.officeAvailable = false;
      els.officeWarning().classList.remove('hidden');
    } else {
      // host is null — opened in a browser for testing
      state.officeAvailable = false;
      els.officeWarning().classList.remove('hidden');
    }
    startApp();
  });
}

/* ═══════════════════════════════════════════════════════════════════
   APP BOOTSTRAP
═══════════════════════════════════════════════════════════════════ */

async function startApp() {
  await initColorSwatches();
  bindControls();
  loadIcons();
}

/* ═══════════════════════════════════════════════════════════════════
   ICON LOADING — Sprite approach
   One request fetches ALL icons as a single SVG sprite.
   We parse <symbol> elements to extract viewBox + path data.
═══════════════════════════════════════════════════════════════════ */

async function loadIcons() {
  showState('loading');

  let spriteText = null;
  let lastError  = null;

  for (const url of SPRITE_URLS) {
    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      spriteText = await res.text();
      break;
    } catch (err) {
      lastError = err;
      console.warn(`[ShapR] Sprite fetch failed from ${url}:`, err.message);
    }
  }

  if (!spriteText) {
    showError(`Could not fetch the Tabler Icons sprite from CDN. Check your internet connection.\n(${lastError?.message || 'Unknown error'})`);
    return;
  }

  try {
    const icons = parseSprite(spriteText);
    if (icons.length === 0) throw new Error('No icons found in sprite');

    state.allIcons     = icons;
    state.filteredIcons = icons;
    state.loaded       = true;

    updateIconCount(icons.length, icons.length);
    showState('grid');
    recalcColumns();
    renderGrid();

  } catch (err) {
    console.error('[ShapR] Sprite parse error:', err);
    showError(`Failed to parse icon data: ${err.message}`);
  }
}

/**
 * Parse the SVG sprite into a flat icon array.
 * Each <symbol id="tabler-{name}"> contains the icon paths.
 *
 * @param {string} spriteText  Raw SVG text
 * @returns {{ name:string, viewBox:string, pathData:string }[]}
 */
function parseSprite(spriteText) {
  // Use DOMParser — fast and reliable in modern browsers
  const parser = new DOMParser();
  const doc    = parser.parseFromString(spriteText, 'image/svg+xml');

  const symbols = doc.querySelectorAll('symbol[id^="tabler-"]');
  const icons   = [];

  symbols.forEach((sym) => {
    const id   = sym.getAttribute('id');          // e.g. "tabler-home"
    const name = id.replace(/^tabler-/, '');       // e.g. "home"

    // Skip variant suffixes like -filled if present (keep outline only)
    if (name.endsWith('-filled')) return;

    const viewBox  = sym.getAttribute('viewBox') || '0 0 24 24';
    const pathData = sym.innerHTML.trim();         // inner SVG elements

    icons.push({ name, viewBox, pathData });
  });

  // Sort alphabetically
  icons.sort((a, b) => a.name.localeCompare(b.name));

  return icons;
}

/* ═══════════════════════════════════════════════════════════════════
   SEARCH
═══════════════════════════════════════════════════════════════════ */

function onSearchInput(e) {
  clearTimeout(state.searchTimer);
  const query = e.target.value.trim();

  // Toggle clear button visibility
  els.searchClear().classList.toggle('hidden', query.length === 0);

  state.searchTimer = setTimeout(() => searchIcons(query), 150);
}

function searchIcons(query) {
  if (!state.loaded) return;

  if (!query) {
    state.filteredIcons = state.allIcons;
  } else {
    const terms = query.toLowerCase().split(/\s+/).filter(Boolean);
    state.filteredIcons = state.allIcons.filter((icon) =>
      terms.every((term) => icon.name.includes(term))
    );
  }

  const total    = state.allIcons.length;
  const filtered = state.filteredIcons.length;

  updateIconCount(filtered, total);

  if (filtered === 0) {
    showState('empty');
  } else {
    showState('grid');
    // Reset scroll to top on new search
    const container = els.gridContainer();
    if (container) container.scrollTop = 0;
    state.lastScrollTop = 0;
    renderGrid();
  }
}

function clearSearch() {
  const input = els.searchInput();
  input.value = '';
  els.searchClear().classList.add('hidden');
  input.focus();
  searchIcons('');
}

/* ═══════════════════════════════════════════════════════════════════
   VIRTUAL SCROLL GRID
   Only renders the icons currently visible in the viewport
   plus OVERSCAN_ROWS rows above and below.
═══════════════════════════════════════════════════════════════════ */

/** Recalculate column count based on container width */
function recalcColumns() {
  const container = els.gridContainer();
  if (!container) return;

  const availableWidth = container.clientWidth - GRID_PADDING * 2;
  const cellWidth      = CARD_SIZE + CARD_GAP;
  state.columns        = Math.max(1, Math.floor((availableWidth + CARD_GAP) / cellWidth));
}

function renderGrid() {
  const container = els.gridContainer();
  const spacer    = els.gridSpacer();
  const viewport  = els.gridViewport();
  if (!container || !spacer || !viewport) return;

  const icons   = state.filteredIcons;
  const cols    = state.columns;
  const rowH    = CARD_SIZE + CARD_GAP;
  const rows    = Math.ceil(icons.length / cols);
  const totalH  = rows * rowH + GRID_PADDING * 2;

  // Set the spacer to full scroll height
  spacer.style.height = `${totalH}px`;

  renderVisibleRows(container, viewport, icons, cols, rowH);
}

function renderVisibleRows(container, viewport, icons, cols, rowH) {
  const scrollTop     = container.scrollTop;
  const containerH    = container.clientHeight;
  const totalRows     = Math.ceil(icons.length / cols);

  // Which rows are visible?
  const firstVisRow = Math.max(0, Math.floor((scrollTop - GRID_PADDING) / rowH) - OVERSCAN_ROWS);
  const lastVisRow  = Math.min(totalRows - 1,
    Math.ceil((scrollTop + containerH - GRID_PADDING) / rowH) + OVERSCAN_ROWS
  );

  // Translate the viewport div so cards appear at the right position
  const offsetY = firstVisRow * rowH + GRID_PADDING;
  viewport.style.transform = `translateY(${offsetY}px)`;

  const firstIconIdx = firstVisRow * cols;
  const lastIconIdx  = Math.min(icons.length - 1, (lastVisRow + 1) * cols - 1);

  // Build fragment for visible icons
  const fragment = document.createDocumentFragment();
  for (let i = firstIconIdx; i <= lastIconIdx; i++) {
    fragment.appendChild(createIconCard(icons[i]));
  }

  viewport.replaceChildren(fragment);
}

function onGridScroll(e) {
  const container = e.currentTarget;
  const icons     = state.filteredIcons;
  const cols      = state.columns;
  const rowH      = CARD_SIZE + CARD_GAP;
  const viewport  = els.gridViewport();

  // Throttle: only re-render if scrolled by at least half a row
  if (Math.abs(container.scrollTop - state.lastScrollTop) < rowH / 2) return;
  state.lastScrollTop = container.scrollTop;

  renderVisibleRows(container, viewport, icons, cols, rowH);
}

/* ═══════════════════════════════════════════════════════════════════
   ICON CARD CREATION
═══════════════════════════════════════════════════════════════════ */

function createIconCard(icon) {
  const card      = document.createElement('div');
  card.className  = 'icon-card';
  card.setAttribute('role', 'listitem');
  card.setAttribute('tabindex', '0');
  card.setAttribute('title', icon.name);
  card.setAttribute('aria-label', `Insert ${icon.name} icon`);
  card.dataset.iconName = icon.name;

  // Inline SVG preview at 28px, coloured with selected colour
  const svgEl = buildPreviewSvg(icon, state.selectedColor, 28);
  card.appendChild(svgEl);

  // Name label
  const nameEl       = document.createElement('span');
  nameEl.className   = 'icon-card-name';
  nameEl.textContent = icon.name;
  card.appendChild(nameEl);

  // Click — insert into PowerPoint
  card.addEventListener('click', () => handleIconClick(card, icon));

  // Keyboard — Enter or Space
  card.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      handleIconClick(card, icon);
    }
  });

  return card;
}

/**
 * Build an inline SVG element for card preview.
 * Uses the icon's pathData from the sprite symbol.
 */
function buildPreviewSvg(icon, color, size) {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', icon.viewBox);
  svg.setAttribute('width', size);
  svg.setAttribute('height', size);
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', color);
  svg.setAttribute('stroke-width', '1.5');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  svg.setAttribute('aria-hidden', 'true');
  svg.innerHTML = icon.pathData;
  return svg;
}

/* ═══════════════════════════════════════════════════════════════════
   COLOUR & SIZE CONTROLS
═══════════════════════════════════════════════════════════════════ */

function selectColor(hex) {
  state.selectedColor = hex;
  // Refresh the visible cards' SVG previews with the new colour
  refreshCardPreviews();
}

function selectSize(px) {
  state.selectedSize = parseInt(px, 10);
  els.sizeDisplay().textContent = `${px}px`;
}

/** Update all currently-rendered SVG previews to the active colour */
function refreshCardPreviews() {
  const viewport = els.gridViewport();
  if (!viewport) return;
  viewport.querySelectorAll('.icon-card').forEach((card) => {
    const name = card.dataset.iconName;
    const icon = state.allIcons.find((i) => i.name === name);
    if (!icon) return;
    const oldSvg = card.querySelector('svg');
    const newSvg = buildPreviewSvg(icon, state.selectedColor, 28);
    if (oldSvg) card.replaceChild(newSvg, oldSvg);
    else card.insertBefore(newSvg, card.firstChild);
  });
}

/* ═══════════════════════════════════════════════════════════════════
   DYNAMIC COLOUR SWATCHES
═══════════════════════════════════════════════════════════════════ */

/**
 * Build colour swatches from the presentation theme (if available) or defaults.
 * Called once during startApp(), before bindControls.
 */
async function initColorSwatches() {
  const colors    = await getThemeColors();
  const container = els.colorSwatches();
  if (!container) return;

  colors.forEach((c, i) => {
    const label      = document.createElement('label');
    label.className  = 'swatch-label';
    label.title      = `${c.name} ${c.hex}`;

    const radio   = document.createElement('input');
    radio.type    = 'radio';
    radio.name    = 'icon-color';
    radio.value   = c.hex;
    radio.className = 'sr-only';
    if (i === 0) radio.checked = true;

    const span      = document.createElement('span');
    span.className  = 'swatch';
    span.style.background = c.hex;
    span.setAttribute('aria-label', c.name);
    if (isLightColor(c.hex)) span.classList.add('swatch--light');

    label.appendChild(radio);
    label.appendChild(span);
    container.appendChild(label);

    radio.addEventListener('change', (e) => {
      if (e.target.checked) {
        selectColor(e.target.value);
        syncHexInput(e.target.value);
      }
    });
  });

  // Set initial colour
  state.selectedColor = colors[0].hex;
  syncHexInput(colors[0].hex);
}


/**
 * Read colours from the active presentation's OOXML theme file.
 * Opens the PPTX as a ZIP via JSZip, parses ppt/theme/theme1.xml,
 * and extracts the <a:clrScheme> entries.  Falls back to DEFAULT_COLORS
 * when Office or JSZip is unavailable.
 */
async function getThemeColors() {
  if (!state.officeAvailable || typeof JSZip === 'undefined') return DEFAULT_COLORS;

  try {
    /* ── 1. Get PPTX bytes via the Common API ── */
    const fileBytes = await new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 65536 },
        async (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            return reject(new Error(result.error.message));
          }
          const file       = result.value;
          const sliceCount = file.sliceCount;
          const parts      = [];

          for (let i = 0; i < sliceCount; i++) {
            const slice = await new Promise((res, rej) => {
              file.getSliceAsync(i, (sr) => {
                if (sr.status === Office.AsyncResultStatus.Succeeded) res(sr.value.data);
                else rej(new Error(sr.error.message));
              });
            });
            parts.push(slice);
          }
          file.closeAsync();

          /* Concatenate Uint8Arrays */
          const total = parts.reduce((n, p) => n + p.byteLength, 0);
          const buf   = new Uint8Array(total);
          let offset  = 0;
          for (const p of parts) {
            buf.set(new Uint8Array(p), offset);
            offset += p.byteLength;
          }
          resolve(buf);
        },
      );
    });

    /* ── 2. Unzip and find the theme XML ── */
    const zip = await JSZip.loadAsync(fileBytes);
    const themeFile = zip.file('ppt/theme/theme1.xml');
    if (!themeFile) return DEFAULT_COLORS;

    const themeXml = await themeFile.async('text');

    /* ── 3. Parse <a:clrScheme> ── */
    const parser  = new DOMParser();
    const doc     = parser.parseFromString(themeXml, 'application/xml');
    const ns      = 'http://schemas.openxmlformats.org/drawingml/2006/main';

    /* Tag → friendly name mapping (same order PowerPoint uses) */
    const COLOR_TAGS = [
      { tag: 'dk1',     name: 'Dark 1' },
      { tag: 'lt1',     name: 'Light 1' },
      { tag: 'dk2',     name: 'Dark 2' },
      { tag: 'lt2',     name: 'Light 2' },
      { tag: 'accent1', name: 'Accent 1' },
      { tag: 'accent2', name: 'Accent 2' },
      { tag: 'accent3', name: 'Accent 3' },
      { tag: 'accent4', name: 'Accent 4' },
      { tag: 'accent5', name: 'Accent 5' },
      { tag: 'accent6', name: 'Accent 6' },
      { tag: 'hlink',   name: 'Hyperlink' },
      { tag: 'folHlink', name: 'Followed Hyperlink' },
    ];

    const scheme = doc.getElementsByTagNameNS(ns, 'clrScheme')[0];
    if (!scheme) return DEFAULT_COLORS;

    const result = [];
    for (const { tag, name } of COLOR_TAGS) {
      const el = scheme.getElementsByTagNameNS(ns, tag)[0];
      if (!el) continue;

      /* Colour can be <a:srgbClr val="4472C4"/> or <a:sysClr lastClr="000000"/> */
      const srgb = el.getElementsByTagNameNS(ns, 'srgbClr')[0];
      const sys  = el.getElementsByTagNameNS(ns, 'sysClr')[0];
      const hex  = srgb
        ? normalizeHex(srgb.getAttribute('val'))
        : sys
          ? normalizeHex(sys.getAttribute('lastClr'))
          : null;

      if (hex) result.push({ hex, name });
    }

    if (result.length >= 2) {
      /* Deduplicate while preserving order */
      const seen = new Set();
      return result.filter((c) => {
        const key = c.hex.toUpperCase();
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    }
  } catch (err) {
    console.warn('[ShapR] Could not read theme colors:', err.message);
  }

  return DEFAULT_COLORS;
}

/** Ensure hex string is #RRGGBB */
function normalizeHex(raw) {
  let h = String(raw).trim().replace(/^#/, '');
  if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  return '#' + h.toUpperCase();
}

/** True when a colour is light enough to need a visible border */
function isLightColor(hex) {
  const h = hex.replace('#', '');
  const r = parseInt(h.substring(0, 2), 16);
  const g = parseInt(h.substring(2, 4), 16);
  const b = parseInt(h.substring(4, 6), 16);
  return (r * 0.299 + g * 0.587 + b * 0.114) > 186;
}

/** Keep the hex text input and native picker in sync */
function syncHexInput(hex) {
  const input  = els.hexInput();
  if (input) input.value = hex.replace('#', '').toUpperCase();
  const picker = els.colorPicker();
  if (picker) picker.value = hex.length === 7 ? hex : '#000000';
}

/** Uncheck all swatch radios (used when a custom colour is picked) */
function uncheckSwatches() {
  document.querySelectorAll('input[name="icon-color"]').forEach((r) => { r.checked = false; });
}

/* ═══════════════════════════════════════════════════════════════════
   ICON INSERTION INTO POWERPOINT
═══════════════════════════════════════════════════════════════════ */

async function handleIconClick(cardEl, icon) {
  cardEl.classList.add('inserting');
  cardEl.removeEventListener('animationend', onCardAnimEnd);
  cardEl.addEventListener('animationend', onCardAnimEnd, { once: true });

  if (!state.officeAvailable) {
    showToast(`"${icon.name}" — open in PowerPoint to insert`, 'error');
    return;
  }

  await insertIcon(icon);
}

function onCardAnimEnd(e) {
  e.currentTarget.classList.remove('inserting');
}

/**
 * Build the full SVG string for insertion (not just the preview).
 * The SVG is sized to selectedSize × selectedSize.
 */
function buildInsertSvg(icon, color, size) {
  return [
    `<svg xmlns="http://www.w3.org/2000/svg"`,
    ` viewBox="${icon.viewBox}"`,
    ` width="${size}" height="${size}"`,
    ` fill="none"`,
    ` stroke="${color}"`,
    ` stroke-width="1.5"`,
    ` stroke-linecap="round"`,
    ` stroke-linejoin="round"`,
    `>`,
    icon.pathData,
    `</svg>`,
  ].join('');
}

/**
 * Apply a colour to an SVG string.
 * Handles both attribute-style and inline-style stroke declarations.
 */
function applyColorToSvg(svgString, color) {
  let result = svgString;

  // Replace explicit stroke="..." attributes (skip stroke="none")
  result = result.replace(/\bstroke="(?!none")[^"]*"/g, `stroke="${color}"`);

  // Replace CSS inline stroke: ... declarations
  result = result.replace(/\bstroke\s*:\s*[^;}"']+/g, `stroke:${color}`);

  // If no stroke attribute on the root <svg>, add it
  if (!/\bstroke=/.test(result.substring(0, result.indexOf('>')))) {
    result = result.replace('<svg', `<svg stroke="${color}"`);
  }

  // Ensure fill="none" on root SVG element
  if (!/\bfill=/.test(result.substring(0, result.indexOf('>')))) {
    result = result.replace('<svg', `<svg fill="none"`);
  }

  return result;
}

async function insertIcon(icon) {
  const size     = state.selectedSize;
  const color    = state.selectedColor;
  const svgRaw   = buildInsertSvg(icon, color, size);
  const svgFinal = applyColorToSvg(svgRaw, color);

  // Dimensions in points (1px ≈ 0.75pt, Office uses points/EMUs)
  // Office.js addSvg left/top/width/height are in points
  const ptSize = Math.round(size * 0.75);

  try {
    // ── Primary: Office.context.document.setSelectedDataAsync ────
    // Works in both PowerPoint Desktop and Online — insert as PNG image
    if (
      typeof Office !== 'undefined' &&
      Office.context?.document?.setSelectedDataAsync
    ) {
      const dataUrl = await svgToDataUrl(svgFinal, size, size);
      await setSelectedDataAsync(dataUrl);
      showToast(`Inserted "${icon.name}" (${size}px)`, 'success');
      return;
    }

    throw new Error('No compatible Office.js insertion API found.');

  } catch (err) {
    console.error('[ShapR] Insert error:', err);
    showToast(`Insert failed: ${err.message}`, 'error');
  }
}

/** Convert px to EMU (English Metric Units) — 1px = 9525 EMU */
function convertPxToEmu(px) {
  return Math.round(px * 9525);
}

/**
 * Rasterise SVG string to a PNG data URL via an off-screen canvas.
 * Used as fallback when addSvg is unavailable.
 */
function svgToDataUrl(svgString, width, height) {
  return new Promise((resolve, reject) => {
    const blob  = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
    const url   = URL.createObjectURL(blob);
    const img   = new Image();
    const scale = 64;  // 8x for crisp presentation-quality output

    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width  = width  * scale;
      canvas.height = height * scale;
      const ctx    = canvas.getContext('2d');
      ctx.scale(scale, scale);
      ctx.drawImage(img, 0, 0, width, height);
      URL.revokeObjectURL(url);
      resolve(canvas.toDataURL('image/png'));
    };

    img.onerror = (e) => {
      URL.revokeObjectURL(url);
      reject(new Error('SVG-to-PNG conversion failed'));
    };

    img.src = url;
  });
}

/**
 * Wrap Office setSelectedDataAsync in a Promise
 */
function setSelectedDataAsync(dataUrl) {
  return new Promise((resolve, reject) => {
    const base64 = dataUrl.replace(/^data:image\/png;base64,/, '');
    Office.context.document.setSelectedDataAsync(
      base64,
      { coercionType: Office.CoercionType.Image },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message || 'setSelectedDataAsync failed'));
        }
      }
    );
  });
}

/* ═══════════════════════════════════════════════════════════════════
   UI STATE MANAGEMENT
═══════════════════════════════════════════════════════════════════ */

/**
 * Show one of: 'loading' | 'error' | 'empty' | 'grid'
 */
function showState(which) {
  els.loadingState().classList.toggle('hidden', which !== 'loading');
  els.errorState().classList.toggle('hidden',   which !== 'error');
  els.emptyState().classList.toggle('hidden',   which !== 'empty');
  els.gridContainer().classList.toggle('hidden', which !== 'grid');
}

function showError(message) {
  els.errorDetail().textContent = message;
  showState('error');
}

function updateIconCount(shown, total) {
  const countEl  = els.iconCount();
  const resultEl = els.resultCount();

  if (shown === total) {
    countEl.textContent  = `${total.toLocaleString()} icons`;
    resultEl.textContent = `${total.toLocaleString()} icons`;
  } else {
    countEl.textContent  = `${shown.toLocaleString()} of ${total.toLocaleString()}`;
    resultEl.textContent = `${shown.toLocaleString()} of ${total.toLocaleString()} icons`;
  }
}

/* ═══════════════════════════════════════════════════════════════════
   TOAST NOTIFICATIONS
═══════════════════════════════════════════════════════════════════ */

function showToast(message, type = 'success') {
  const toast = els.toast();
  clearTimeout(state.toastTimer);

  toast.textContent = message;
  toast.className   = `toast toast--${type}`;
  toast.classList.remove('hidden');

  state.toastTimer = setTimeout(() => {
    toast.classList.add('hidden');
  }, 2800);
}

/* ═══════════════════════════════════════════════════════════════════
   CONTROL BINDING
═══════════════════════════════════════════════════════════════════ */

function bindControls() {
  // Search
  const searchInput = els.searchInput();
  searchInput.addEventListener('input', onSearchInput);
  searchInput.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') clearSearch();
  });

  els.searchClear().addEventListener('click', clearSearch);

  // Native colour picker (the rainbow swatch)
  const colorPicker = els.colorPicker();
  if (colorPicker) {
    colorPicker.addEventListener('input', (e) => {
      selectColor(e.target.value);
      syncHexInput(e.target.value);
      uncheckSwatches();
    });
  }

  // Hex text input
  const hexInput = els.hexInput();
  if (hexInput) {
    hexInput.addEventListener('input', (e) => {
      let val = e.target.value.replace(/[^0-9a-fA-F]/g, '').substring(0, 6);
      e.target.value = val.toUpperCase();
      if (val.length === 6) {
        const hex = '#' + val;
        selectColor(hex);
        const pk = els.colorPicker();
        if (pk) pk.value = hex;
        uncheckSwatches();
      }
    });
    hexInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        const val = e.target.value.replace(/[^0-9a-fA-F]/g, '');
        if (val.length >= 3) {
          const norm = val.length === 3
            ? val[0] + val[0] + val[1] + val[1] + val[2] + val[2]
            : val.substring(0, 6).padEnd(6, '0');
          const hex = '#' + norm;
          selectColor(hex);
          syncHexInput(hex);
          uncheckSwatches();
        }
      }
    });
  }

  // Size buttons — update active style manually since we use labels
  document.querySelectorAll('input[name="icon-size"]').forEach((radio) => {
    radio.addEventListener('change', (e) => {
      if (e.target.checked) {
        // Remove active class from all size buttons
        document.querySelectorAll('.size-btn').forEach((btn) => {
          btn.classList.remove('size-btn--active');
        });
        // Add to the selected one
        e.target.nextElementSibling?.classList.add('size-btn--active');
        selectSize(e.target.value);
      }
    });
  });

  // Retry button
  els.retryBtn().addEventListener('click', () => {
    state.loaded = false;
    state.allIcons = [];
    state.filteredIcons = [];
    loadIcons();
  });

  // Scroll listener on grid container
  const container = els.gridContainer();
  container.addEventListener('scroll', onGridScroll, { passive: true });

  // Resize observer — recalculate columns if pane width changes
  if (typeof ResizeObserver !== 'undefined') {
    const ro = new ResizeObserver(() => {
      if (!state.loaded) return;
      const oldCols = state.columns;
      recalcColumns();
      if (state.columns !== oldCols) renderGrid();
    });
    ro.observe(container);
  }
}

/* ═══════════════════════════════════════════════════════════════════
   ENTRY POINT
═══════════════════════════════════════════════════════════════════ */

// Wait for DOM, then initialise
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initOffice);
} else {
  initOffice();
}
