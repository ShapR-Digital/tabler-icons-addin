/**
 * ShapR Tabler Icons — PowerPoint Add-in
 *
 * Architecture:
 *  - Resolves latest @tabler/icons version from npm registry
 *  - Fetches icons.json (metadata/tags/categories) + tabler-nodes-outline.json (path data)
 *  - Virtual scroll renders only visible cards (~100 at a time)
 *  - Debounced search filters by name + tags; dropdown filters by category
 *  - Office.js inserts coloured SVG into active slide
 */

/* ═══════════════════════════════════════════════════════════════════
   CONSTANTS & CONFIG
═══════════════════════════════════════════════════════════════════ */

/** npm registry endpoint to resolve the true latest version */
const NPM_LATEST_URL = 'https://registry.npmjs.org/@tabler/icons/latest';

/** Known-good fallback version if the registry is unreachable */
const FALLBACK_VERSION = '3.41.1';

/** CDN base URLs (tried in order) */
const CDN_BASES = [
  'https://cdn.jsdelivr.net/npm/@tabler/icons',
  'https://unpkg.com/@tabler/icons',
];

/** Insertion sizes in points (1 cm ≈ 28.35 pt) */
const SIZE_MAP = {
  '32':  28.35,   // S = 1 × 1 cm
  '64':  85.04,   // M = 3 × 3 cm
  '96': 170.08,   // L = 6 × 6 cm
};

/** ShapR brand colours */
const DEFAULT_COLORS = [
  { hex: '#0B2B45', name: 'ShapR Navy' },
  { hex: '#F2BE22', name: 'ShapR Yellow' },
  { hex: '#185359', name: 'ShapR Teal' },
  { hex: '#30A5BF', name: 'ShapR Cyan' },
  { hex: '#D8256A', name: 'ShapR Pink' },
  { hex: '#000000', name: 'Black' },
  { hex: '#FFFFFF', name: 'White' },
];

/**
 * OfficeThemes.css class names to probe for document theme colours.
 * Office.js rewrites these CSS rules at runtime to match the active theme.
 */
const THEME_COLOR_PROBES = [
  { className: 'office-contentAccent1-color', name: 'Accent 1' },
  { className: 'office-contentAccent2-color', name: 'Accent 2' },
  { className: 'office-contentAccent3-color', name: 'Accent 3' },
  { className: 'office-contentAccent4-color', name: 'Accent 4' },
  { className: 'office-contentAccent5-color', name: 'Accent 5' },
  { className: 'office-contentAccent6-color', name: 'Accent 6' },
  { className: 'office-docTheme-primary-fontColor',   name: 'Text Primary' },
  { className: 'office-docTheme-secondary-fontColor',  name: 'Text Secondary' },
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
  /** Full list of parsed icon objects: { name, viewBox, pathData, tags, category } */
  allIcons: [],

  /** Currently displayed subset (after search filter) */
  filteredIcons: [],

  /** Hex string of the active colour swatch */
  selectedColor: '#0B2B45',

  /** Insertion size in px */
  selectedSize: 85.04,

  /** Whether Office.js is available */
  officeAvailable: false,

  /** Virtual scroll: number of columns in the grid */
  columns: 4,

  /** Scroll position cache */
  lastScrollTop: 0,

  /** Active category filter (empty string = all) */
  selectedCategory: '',

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
  categoryFilter: () => document.getElementById('category-filter'),
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
  // Use the presentation's theme colours when running in PowerPoint;
  // fall back to ShapR brand colours otherwise.
  const themeColors = state.officeAvailable ? extractThemeColors() : null;
  initColorSwatches(themeColors || DEFAULT_COLORS);
  bindControls();
  loadIcons();
}

/* ═══════════════════════════════════════════════════════════════════
   ICON LOADING — JSON approach (Tabler Icons v3+)
   Two requests: icons.json (metadata + tags + categories) and
   tabler-nodes-outline.json (SVG path data per icon).
   Version is resolved dynamically from the npm registry so we always
   get the latest icons without relying on CDN @latest caching.
═══════════════════════════════════════════════════════════════════ */

/**
 * Resolve the true latest @tabler/icons version from the npm registry.
 * Falls back to FALLBACK_VERSION if the registry is unreachable.
 */
async function resolveLatestVersion() {
  try {
    const res = await fetch(NPM_LATEST_URL);
    if (res.ok) {
      const { version } = await res.json();
      if (version) return version;
    }
  } catch (_) { /* ignore — use fallback */ }
  return FALLBACK_VERSION;
}

/**
 * Fetch a file from the Tabler CDN, trying each CDN_BASES entry in order.
 */
async function fetchFromCdn(path, version) {
  let lastErr = null;
  for (const base of CDN_BASES) {
    try {
      const res = await fetch(`${base}@${version}/${path}`);
      if (res.ok) return res;
      lastErr = new Error(`HTTP ${res.status} from ${base}`);
    } catch (err) {
      lastErr = err;
    }
  }
  throw lastErr || new Error(`Failed to fetch ${path}`);
}

async function loadIcons() {
  showState('loading');

  try {
    const version = await resolveLatestVersion();
    console.log(`[ShapR] Loading Tabler Icons v${version}`);

    const [metaRes, nodesRes] = await Promise.all([
      fetchFromCdn('icons.json', version),
      fetchFromCdn('tabler-nodes-outline.json', version),
    ]);

    const [meta, nodes] = await Promise.all([
      metaRes.json(),
      nodesRes.json(),
    ]);

    const icons = buildIconRegistry(meta, nodes);
    if (icons.length === 0) throw new Error('No icons found in data');

    state.allIcons      = icons;
    state.filteredIcons = icons;
    state.loaded        = true;

    populateCategoryFilter(icons);
    updateIconCount(icons.length, icons.length);
    showState('grid');
    recalcColumns();
    renderGrid();

  } catch (err) {
    console.error('[ShapR] Load error:', err);
    showError(`Could not load Tabler Icons from CDN. Check your internet connection.\n(${err.message})`);
  }
}

/**
 * Merge icons.json (metadata) with tabler-nodes-outline.json (path data)
 * into a flat array of icon objects.
 *
 * @param {Object} meta   icons.json — { [name]: { category, tags, ... } }
 * @param {Object} nodes  tabler-nodes-outline.json — { [name]: [[tag, attrs], ...] }
 * @returns {{ name, viewBox, pathData, tags, category }[]}
 */
function buildIconRegistry(meta, nodes) {
  const icons = [];

  for (const [name, nodeList] of Object.entries(nodes)) {
    const iconMeta = meta[name] || {};
    icons.push({
      name,
      viewBox:  '0 0 24 24',
      pathData: nodesToPathData(nodeList),
      tags:     Array.isArray(iconMeta.tags) ? iconMeta.tags.map(String) : [],
      category: iconMeta.category || '',
    });
  }

  icons.sort((a, b) => a.name.localeCompare(b.name));
  return icons;
}

/**
 * Convert a node list from tabler-nodes-outline.json into an SVG innerHTML string.
 * Each node is [tagName, attrsObject], e.g. ["path", { "d": "M4 4h16..." }]
 */
function nodesToPathData(nodeList) {
  return nodeList.map(([tag, attrs]) => {
    const attrStr = Object.entries(attrs)
      .map(([k, v]) => `${k}="${v}"`)
      .join(' ');
    return `<${tag} ${attrStr}/>`;
  }).join('');
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

  const terms    = query ? query.toLowerCase().split(/\s+/).filter(Boolean) : [];
  const category = state.selectedCategory;

  state.filteredIcons = state.allIcons.filter((icon) => {
    // Category filter
    if (category && icon.category !== category) return false;

    // Text search — match name OR any tag
    if (terms.length === 0) return true;
    return terms.every((term) =>
      icon.name.includes(term) ||
      icon.tags.some((tag) => tag.toLowerCase().includes(term))
    );
  });

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
   CATEGORY FILTER
═══════════════════════════════════════════════════════════════════ */

/**
 * Populate the category <select> with all unique categories from the icon set,
 * sorted alphabetically. Called once after icons are loaded.
 */
function populateCategoryFilter(icons) {
  const select = els.categoryFilter();
  if (!select) return;

  const categories = [...new Set(icons.map((i) => i.category).filter(Boolean))].sort();

  // Keep the "All categories" option and append the rest
  select.innerHTML = '<option value="">All categories</option>';
  categories.forEach((cat) => {
    const opt       = document.createElement('option');
    opt.value       = cat;
    opt.textContent = cat;
    select.appendChild(opt);
  });
}

/* ═══════════════════════════════════════════════════════════════════
   COLOUR & SIZE CONTROLS
═══════════════════════════════════════════════════════════════════ */

function selectColor(hex) {
  state.selectedColor = hex;
  // When white is selected the icon preview is invisible on a light card bg — swap to navy
  const viewport = els.gridViewport();
  if (viewport) {
    viewport.classList.toggle('white-color-selected', hex.toUpperCase() === '#FFFFFF');
  }
  // Refresh the visible cards' SVG previews with the new colour
  refreshCardPreviews();
}

const SIZE_CM_LABELS = { '32': '1 cm', '64': '3 cm', '96': '6 cm' };

function selectSize(key) {
  state.selectedSize = SIZE_MAP[key];
  els.sizeDisplay().textContent = SIZE_CM_LABELS[key] || key;
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
 * Populate the colour swatch strip from the given colour array.
 */
function initColorSwatches(colors) {
  const container = els.colorSwatches();
  if (!container) return;

  container.innerHTML = '';

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

/**
 * Convert a CSS computed colour string (e.g. "rgb(48, 165, 191)") to "#RRGGBB".
 * Returns null if the string cannot be parsed.
 */
function rgbToHex(rgbStr) {
  const match = rgbStr.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
  if (!match) return null;
  const [, r, g, b] = match;
  return '#' + [r, g, b].map((v) => Number(v).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/**
 * Read document theme colours by probing OfficeThemes.css computed styles.
 * Creates temporary hidden <span> elements, reads getComputedStyle().color,
 * converts to hex, deduplicates, appends Black/White, and returns
 * an array of { hex, name } or null if extraction fails.
 */
function extractThemeColors() {
  try {
    const container = document.createElement('div');
    container.style.cssText = 'position:absolute;left:-9999px;top:-9999px;visibility:hidden;pointer-events:none;';
    document.body.appendChild(container);

    const colors  = [];
    const seenHex = new Set();

    for (const probe of THEME_COLOR_PROBES) {
      const span = document.createElement('span');
      span.className = probe.className;
      container.appendChild(span);

      const computed = window.getComputedStyle(span).color;
      const hex      = rgbToHex(computed);

      if (hex && !seenHex.has(hex)) {
        seenHex.add(hex);
        colors.push({ hex, name: probe.name });
      }
    }

    document.body.removeChild(container);

    // Always offer Black and White
    if (!seenHex.has('#000000')) colors.push({ hex: '#000000', name: 'Black' });
    if (!seenHex.has('#FFFFFF')) colors.push({ hex: '#FFFFFF', name: 'White' });

    // Need at least one non-BW colour to consider the extraction valid
    const nonBW = colors.filter((c) => c.hex !== '#000000' && c.hex !== '#FFFFFF');
    if (nonBW.length < 1) return null;

    return colors;
  } catch (err) {
    console.warn('[ShapR] Theme colour extraction failed:', err);
    return null;
  }
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
  const ptSize   = state.selectedSize;           // already in points via SIZE_MAP
  const color    = state.selectedColor;
  const renderPx = 256;                          // high-res raster for crisp output
  const svgRaw   = buildInsertSvg(icon, color, renderPx);
  const svgFinal = applyColorToSvg(svgRaw, color);

  try {
    if (
      typeof Office !== 'undefined' &&
      Office.context?.document?.setSelectedDataAsync
    ) {
      const dataUrl = await svgToDataUrl(svgFinal, renderPx, renderPx);
      await setSelectedDataAsync(dataUrl, ptSize);

      const cmLabel = { [28.35]: '1 cm', [85.04]: '3 cm', [170.08]: '6 cm' }[ptSize] || `${Math.round(ptSize)}pt`;
      showToast(`Inserted "${icon.name}" (${cmLabel})`, 'success');
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
    const scale = 2;   // 2x for crisp output (256px → 512px canvas)

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

    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error('SVG-to-PNG conversion failed'));
    };

    img.src = url;
  });
}

/**
 * Wrap Office setSelectedDataAsync in a Promise
 */
function setSelectedDataAsync(dataUrl, ptSize) {
  return new Promise((resolve, reject) => {
    const base64 = dataUrl.replace(/^data:image\/png;base64,/, '');
    Office.context.document.setSelectedDataAsync(
      base64,
      {
        coercionType: Office.CoercionType.Image,
        imageWidth: ptSize,
        imageHeight: ptSize,
      },
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

  // Category filter
  const categoryFilter = els.categoryFilter();
  if (categoryFilter) {
    categoryFilter.addEventListener('change', (e) => {
      state.selectedCategory = e.target.value;
      // Re-run search with the current query text so both filters apply
      searchIcons(els.searchInput().value.trim());
    });
  }

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
    state.loaded           = false;
    state.allIcons         = [];
    state.filteredIcons    = [];
    state.selectedCategory = '';
    const cf = els.categoryFilter();
    if (cf) cf.value = '';
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
