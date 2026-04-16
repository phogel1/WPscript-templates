// ==UserScript==
// @name         INU WebPort-Plus
// @namespace    http://tampermonkey.net/
// @version      7.3.20260416.1215
// @description  Enhanced UI for Kiona WebPort
// @match        *://*/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      phogel1.github.io
// @updateURL    https://phogel1.github.io/static-assets/wpp.meta.js
// @downloadURL  https://phogel1.github.io/static-assets/wpp.user.js
// ==/UserScript==

(function () {
    'use strict';

    // Only run on Kiona WebPort pages
    function isWebPort() {
        return !!(document.querySelector('link[href*="webport"]') || document.querySelector('link[href*="designtokens"]') || (typeof unsafeWindow.wp === 'object' && unsafeWindow.wp.Global));
    }

    // ============================================================
    // CONFIG
    // ============================================================
    // Full version is read from the @version header (auto-stamped by sync script).
    // Falls back to the base version if the header parse fails.
    const _fullVersion = (function () {
        try {
            const m = GM_info && GM_info.script && GM_info.script.version;
            if (m) return m;
        } catch (e) {}
        return '7.3';
    })();
    const CFG = {
        version: _fullVersion,
        logPrefix: '[INU WP+]',
        colOffset: 3,
        pollMs: 1000,
        storageKeys: { custom: 'inu_presets_custom', undo: 'inu_presets_undo', monitor: 'inu_monitor_prefs' },
        endpoints: { tagEdit: '/tag/ActionEdit', tagSave: '/tag/actionedit', tagRead: '/tag/read' },
    };

    // ============================================================
    // PAGE DETECTION
    // ============================================================
    function isInuTagPage() {
        return !!(
            document.getElementById('tagtable') &&
            document.querySelector('#tagtable tbody tr.tag') &&
            typeof unsafeWindow.SendFormMulti === 'function'
        );
    }

    // ============================================================
    // DEFAULT PRESETS
    // ============================================================
    const DEFAULT_PRESETS = {
        'Temp °C (-50…150)':      { rawmin:'-500', rawmax:'1500', engmin:'-50', engmax:'150', unit:'°C', format:'0.0' },
        'Temp °C (-50…150) 0.00': { rawmin:'-5000', rawmax:'15000', engmin:'-50', engmax:'150', unit:'°C', format:'0.00' },
        'Temp °C (-50…400)':      { rawmin:'-500', rawmax:'4000', engmin:'-50', engmax:'400', unit:'°C', format:'0.0' },
        'Temp ΔK (-100…100)':     { rawmin:'-1000', rawmax:'1000', engmin:'-100', engmax:'100', unit:'°K', format:'0.0' },
        'Fukt RH% (0…100)':      { rawmin:'0', rawmax:'1000', engmin:'0', engmax:'100', unit:'RH%', format:'0.0' },
        'Tryck Pa (0…1000)':     { rawmin:'0', rawmax:'10000', engmin:'0', engmax:'1000', unit:'Pa', format:'0' },
        'Tryck kPa (0…100)':     { rawmin:'0', rawmax:'10000', engmin:'0', engmax:'100', unit:'kPa', format:'0.0' },
        'Tryck bar (0…10) 0.00': { rawmin:'0', rawmax:'1000', engmin:'0', engmax:'10', unit:'bar', format:'0.00' },
        'Flöde l/s (0…10000)':   { rawmin:'0', rawmax:'10000', engmin:'0', engmax:'10000', unit:'l/s', format:'0' },
        'CO₂ ppm (0…2000)':      { rawmin:'0', rawmax:'20000', engmin:'0', engmax:'2000', unit:'ppm', format:'0' },
        'Procent % (0…100)':     { rawmin:'0', rawmax:'10000', engmin:'0', engmax:'100', unit:'%', format:'0.0' },
        'Effekt kW (0…100)':     { rawmin:'0', rawmax:'10000', engmin:'0', engmax:'100', unit:'kW', format:'0.0' },
        'Energi kWh (0…99999)':  { rawmin:'0', rawmax:'99999', engmin:'0', engmax:'99999', unit:'kWh', format:'0' },
        'Tid s (0…3600)':        { rawmin:'0', rawmax:'3600', engmin:'0', engmax:'3600', unit:'s', format:'0' },
        'Tid min (0…1440)':      { rawmin:'0', rawmax:'1440', engmin:'0', engmax:'1440', unit:'min', format:'0' },
        'Tid h (0…24)':          { rawmin:'0', rawmax:'24', engmin:'0', engmax:'24', unit:'h', format:'0' },
        'Klocka HH:MM (0…2359)': { rawmin:'0', rawmax:'2359', engmin:'0', engmax:'2359', unit:'', format:'##.##' },
        'Digital 0/1':       { rawmin:'0', rawmax:'1', engmin:'0', engmax:'1', unit:'', format:'0' },
    };

    // ============================================================
    // AUTO-SUGGEST (Swedish BMS naming: GT=temp, GP=tryck, GF=fukt, STV=ventil)
    // ============================================================
    const SUGGEST_RULES = [
        { p: /GT\d+.*_PV$/,  s: 'Temp °C (-50…150)' },
        { p: /GT\d+.*_SP/,   s: 'Temp °C (-50…150)' },
        { p: /GT\d+.*_OPM$/, s: 'Temp °C (-50…150)' },
        { p: /GT\d+.*_ALL$/, s: 'Temp °C (-50…150)' },
        { p: /GT\d+/,        s: 'Temp °C (-50…150)' },
        { p: /GP\d+/,        s: 'Tryck Pa (0…1000)' },
        { p: /GF\d+/,        s: 'Flöde l/s (0…10000)' },
        { p: /GM\d+/,        s: 'Fukt % (0…100)' },
        { p: /STV\d+/,       s: 'Procent % (0…100)' },
        { p: /GQ\d+/,        s: 'CO₂ ppm (0…2000)' },
        { p: /_OP_FB$/,      s: 'Procent % (0…100)' },
        { p: /_FB$/,         s: 'Procent % (0…100)' },
        { p: /_CMD$/,        s: 'Digital 0/1' },
        { p: /_M$/,          s: 'Digital 0/1' },
        { p: /_FAULT$/,      s: 'Digital 0/1' },
        { p: /_DI\d*$/,      s: 'Digital 0/1' },
        { p: /_AL\d*$/,      s: 'Digital 0/1' },
        { p: /_LAL$/,        s: 'Digital 0/1' },
        { p: /_HAL$/,        s: 'Digital 0/1' },
        { p: /_AD$/,         s: 'Tid s (0…3600)' },
    ];
    function suggest(tag) { for (const r of SUGGEST_RULES) if (r.p.test(tag)) return r.s; return null; }

    // ============================================================
    // LOCALSTORAGE
    // ============================================================
    function loadCustom()  { try { return JSON.parse(GM_getValue(CFG.storageKeys.custom, '{}')) || {}; } catch(e) { console.warn(CFG.logPrefix, 'loadCustom', e); return {}; } }
    function saveCustom(o) { GM_setValue(CFG.storageKeys.custom, JSON.stringify(o)); }
    function allPresets()   { return { ...DEFAULT_PRESETS, ...loadCustom() }; }
    // Undo is session-only (in-memory) — resets on page refresh
    const undoMap = {};
    function saveUndo(tag, s) { undoMap[tag] = s; }
    function getUndo(tag) { return undoMap[tag] || null; }
    function delUndo(tag) { delete undoMap[tag]; }

    // Notification log — persisted in sessionStorage so navigation doesn't wipe it
    const _logEntries = [];
    const LOG_MAX = 150;
    const LOG_MAX_AGE_MS = 8 * 60 * 60 * 1000; // discard entries older than 8 hours
    const LOG_SS_KEY = 'inu_wp_log';
    let _logEntriesEl = null, _logBodyEl = null;
    let _logFilter = 'all'; // 'all' | 'success' | 'error' | 'warning'
    let _ownToast = false;  // flag: true while our own toast() calls toastr

    // Load persisted entries (skip anything older than LOG_MAX_AGE_MS)
    (function _loadLog() {
        try {
            const raw = sessionStorage.getItem(LOG_SS_KEY);
            if (!raw) return;
            const now = Date.now();
            JSON.parse(raw).forEach(e => {
                const ts = new Date(e.ts);
                if (now - ts.getTime() < LOG_MAX_AGE_MS)
                    _logEntries.push({ ts, level: e.level, msg: e.msg, src: e.src });
            });
        } catch(_) { /* corrupted / quota, ignore */ }
    })();

    function _saveLog() {
        try {
            sessionStorage.setItem(LOG_SS_KEY,
                JSON.stringify(_logEntries.map(e => ({ ts: e.ts.toISOString(), level: e.level, msg: e.msg, src: e.src }))));
        } catch(_) { /* quota exceeded, ignore */ }
    }

    // ============================================================
    // STYLES
    // ============================================================
    let _si = false;
    // INU logo — inlined from INUlogo_Black.svg, uses currentColor for theming
    const INU_LOGO_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 546.84 135.95" style="height:12px;vertical-align:middle;fill:currentColor;"><g transform="matrix(0.13333333,0,0,-0.13333333,0,135.94667)"><path d="m 2770.98,393.707 c 10.05,46.367 15.77,66.434 25.13,106.367 l 66.63,244.832 c 44.31,191.399 -44.42,252.969 -184.13,252.969 h -147.04 l -162.47,-597 c -0.67,-2.402 -1.13,-4.777 -1.71,-7.168 h 403.59"/><path d="M 2358.14,317.586 C 2322.51,141.43 2529.86,0 2784.76,0 h 512.63 c 292.81,0 581.48,180.395 641.48,400.875 l 162.47,597 H 3701.33 L 3565.86,500.074 C 3538.41,399.242 3406.4,316.742 3272.5,316.742 l -914.36,0.844"/><path d="M 368.418,15.4883 H 170.434 C 47.3047,15.4883 -25.3438,115.309 8.16797,238.434 L 116.559,636.695 H 537.477 L 368.418,15.4883"/><path d="M 555.551,707.961 H 134.633 l 82.762,289.914 H 638.316 L 555.551,707.961"/><path d="m 703.086,630.82 c -1.27,-4.047 -2.996,-8.007 -4.109,-12.097 L 536.484,15.4883 H 936.5 l 137.11,509.9607 c 10.66,39.172 18.57,67.231 30.43,105.371 H 703.086"/><path d="m 1845.48,1019.61 h -601.55 c -298.387,0 -445.125,-107.762 -507.438,-312.672 l 913.298,1.851 c 133.9,0 221.01,-82.5 193.57,-183.328 l -74.62,-274.195 c -34.5,-126.77 40.3,-229.5433 167.07,-229.5433 h 170.48 l 162.47,597.0043 c 60,220.484 -130.46,400.883 -423.28,400.883"/></g></svg>';

    function injectBrandPill() {
        if (document.getElementById('inu-wp-pill')) return;
        const topMenu = document.getElementById('top-menu');
        if (!topMenu) return;
        const pill = document.createElement('span');
        pill.id = 'inu-wp-pill';
        pill.innerHTML = INU_LOGO_SVG + '<span style="margin-left:5px;">WebPort+</span>';
        pill.style.cssText = 'padding:3px 10px;border-radius:3px;font-size:10px;font-weight:600;color:#fff;background:#1b5e20;user-select:none;cursor:default;align-self:center;display:inline-flex;align-items:center;';
        pill.title = 'v' + CFG.version;
        const nav = topMenu.parentElement;
        if (nav) nav.insertBefore(pill, nav.firstChild);
    }

    function injectStyles() {
        if (_si) return; _si = true;
        injectBrandPill();
        const s = document.createElement('style');
        s.textContent = `
#tagtable { table-layout:fixed; width:100% !important; }
#tagtable thead th { z-index:10; }
#tagtable th, #tagtable td { padding:3px 4px !important; font-size:11px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
#tagtable .p-larm, #tagtable .p-trend, #tagtable .p-chk { overflow:visible !important; }
#tagtable th:nth-child(15), #tagtable td:nth-child(15) { display:none !important; }
#tagtable th:nth-child(2), #tagtable td:nth-child(2) { width:80px !important; }
#tagtable th:nth-child(3), #tagtable td:nth-child(3) { width:46px !important; }
#tagtable th:nth-child(5), #tagtable td:nth-child(5) { width:100px !important; }
#tagtable th:nth-child(6), #tagtable td:nth-child(6) { width:55px !important; }
#tagtable th:nth-child(7), #tagtable td:nth-child(7) { width:65px !important; }
#tagtable th:nth-child(8), #tagtable td:nth-child(8) { width:55px !important; }
#tagtable th:nth-child(9), #tagtable td:nth-child(9) { width:55px !important; }
#tagtable th:nth-child(10), #tagtable td:nth-child(10) { width:55px !important; }
#tagtable th:nth-child(11), #tagtable td:nth-child(11) { width:55px !important; }
#tagtable th:nth-child(12), #tagtable td:nth-child(12) { width:55px !important; }
#tagtable th:nth-child(13), #tagtable td:nth-child(13) { width:50px !important; }
#tagtable th:nth-child(14), #tagtable td:nth-child(14) { width:150px !important; }
.p-chk { width:20px !important; text-align:center; vertical-align:middle; padding:0 2px !important; }
.p-chk input { cursor:pointer; width:13px; height:13px; margin:0; }
.itb { display:flex; align-items:center; gap:8px; padding:6px 10px; background:#f5f5f8; border-bottom:1px solid #ddd; flex-wrap:wrap; font-size:11px; z-index:50; }
.itb .td { width:1px; height:18px; background:#ccc; margin:0 2px; }
.itb button { background:#5b6abf; color:#fff; border:none; border-radius:4px; padding:3px 9px; cursor:pointer; font-size:11px; font-weight:600; }
.itb button:hover { background:#4a58a8; }
.itb button:disabled { background:#ccc; cursor:not-allowed; }
.itb button.sec { background:#777; }
.itb button.sec:hover { background:#555; }
.itb select { padding:2px 5px; border:1px solid #ccc; border-radius:4px; font-size:11px; max-width:200px; background:#fff; color:#333; }
.itb .inf { color:#666; }
.ism { display:flex; align-items:center; gap:10px; padding:3px 10px; font-size:11px; color:#666; z-index:49; background:#fff; border-bottom:1px solid #eee; }
.ism .dt { width:9px; height:9px; border-radius:50%; display:inline-block; }
.ism .dt.g { background:#4caf50; } .ism .dt.w { background:#ff9800; } .ism .dt.x { background:#90a4ae; }
tr.tag.ru > td { background:#fff8e1 !important; }
tr.tag.ru:hover > td { background:#fff3cd !important; }
tr.tag.p-hidden { display:none !important; }
tr.tag.inu-del > td { background:#fef2f2 !important; color:#bbb !important; }
tr.tag.inu-del:hover > td { background:#fee2e2 !important; }
tr.tag.inu-del .larm-wrap, tr.tag.inu-del .tog-grp { opacity:.2 !important; pointer-events:none; }
tr.tag.inu-del .p-chk input { opacity:.2 !important; pointer-events:none; }
.inu-del-undo { background:#ef4444; color:#fff !important; border:none; border-radius:3px; padding:1px 5px; font-size:10px; font-weight:600; cursor:pointer; margin-right:4px; vertical-align:middle; line-height:1.6; flex-shrink:0; }
.inu-del-undo:hover { background:#dc2626; }
.itb button.danger { background:#ef4444; }
.itb button.danger:hover { background:#dc2626; }
.itb select.fil { padding:2px 5px; border:1px solid #ccc; border-radius:4px; font-size:11px; background:#fff; color:#333; }
.inu-col-filter-btn { font-size:9px !important; margin-left:4px; cursor:pointer; opacity:.4; color:inherit; vertical-align:middle; }
.inu-col-filter-btn:hover, .inu-col-filter-btn.active { opacity:1; color:#4a6cf7; }
.inu-fpill { display:inline-flex; align-items:center; gap:4px; padding:2px 8px 2px 6px; background:#4a6cf7; color:#fff; border-radius:10px; font-size:10px; font-weight:600; cursor:pointer; user-select:none; white-space:nowrap; }
.inu-fpill:hover { background:#3a5ce5; }
.inu-fpill .fa { font-size:9px; opacity:.8; }
.inu-col-filter-dd { position:fixed; z-index:99999; background:#fff; border:1px solid #ccc; border-radius:6px; box-shadow:0 4px 16px rgba(0,0,0,.18); min-width:180px; max-width:260px; font-size:12px; overflow:hidden; }
.inu-cfd-search { padding:6px 8px; border-bottom:1px solid #eee; }
.inu-cfd-search input { width:100%; padding:3px 6px; border:1px solid #ccc; border-radius:4px; font-size:11px; outline:none; box-sizing:border-box; }
.inu-cfd-list { max-height:200px; overflow-y:auto; padding:4px 0; }
.inu-cfd-item { display:flex; align-items:center; gap:6px; padding:3px 10px; cursor:pointer; user-select:none; }
.inu-cfd-item:hover { background:#f0f4ff; }
.inu-cfd-item input[type=checkbox] { margin:0; cursor:pointer; }
.inu-cfd-footer { display:flex; gap:6px; padding:6px 8px; border-top:1px solid #eee; justify-content:flex-end; }
.inu-cfd-footer button { padding:3px 10px; border-radius:4px; border:none; cursor:pointer; font-size:11px; }
.inu-cfd-ok { background:#4a6cf7; color:#fff; }
.inu-cfd-ok:hover { background:#3a5ce5; }
.inu-cfd-cancel { background:#eee; color:#333; }
.inu-cfd-cancel:hover { background:#ddd; }
.larm-wrap { display:flex; align-items:center; justify-content:center; gap:4px; }
.p-larm { padding:0 4px !important; }
.larm-wrap.off .larm-sel { opacity:.35; pointer-events:none; }
.tog-grp { display:flex; align-items:center; gap:2px; }
.tog-ico { font-size:12px; line-height:1; opacity:.4; user-select:none; cursor:pointer; transition:opacity .2s; }
.tog-grp.on .tog-ico { opacity:1; }
.larm-tog { position:relative; display:block; width:32px; height:18px; flex-shrink:0; cursor:pointer; margin:0; padding:0; line-height:0; font-size:0; }
.larm-tog input { opacity:0; width:0; height:0; position:absolute; }
.larm-tog .sl { position:absolute; inset:0; background:#999; border-radius:9px; transition:background .2s; box-shadow:0 0 0 1px rgba(0,0,0,.15); }
.larm-tog .sl::after { content:''; position:absolute; left:2px; top:2px; width:14px; height:14px; background:#fff; border-radius:50%; transition:transform .2s; }
.larm-tog input:checked+.sl { background:#2e7d32; box-shadow:0 0 0 1px rgba(0,0,0,.2); }
.larm-tog input:checked+.sl::after { transform:translateX(14px); }
.larm-tog.trend input:checked+.sl { background:#2196f3; }
.larm-sel { padding:0 4px; border:none; border-radius:3px; font-size:10px; font-weight:600; cursor:pointer; min-width:28px; appearance:none; -webkit-appearance:none; text-align:center; color:#fff; background:#90a4ae; line-height:18px; height:18px; display:block; margin:0; }
.larm-sel:focus { outline:2px solid #5b6abf; outline-offset:1px; }
.larm-sel.lk-a { background:#e53935; } .larm-sel.lk-b { background:#fb8c00; } .larm-sel.lk-c { background:#fdd835; color:#333; } .larm-sel.lk-d { background:#5b6abf; }
.p-help { position:relative; display:inline-flex; align-items:center; justify-content:center; width:20px; height:20px; border-radius:50%; background:#5b6abf; color:#fff; font-size:12px; font-weight:700; cursor:default; margin-left:auto; flex-shrink:0; user-select:none; }
.p-help:hover .p-tip { display:block; }
.p-tip { display:none; position:absolute; right:0; top:28px; width:340px; background:#fff; color:#333; border:1px solid #ccc; border-radius:8px; box-shadow:0 6px 24px rgba(0,0,0,.2); padding:14px 16px; font-size:11px; line-height:1.55; z-index:100001; font-weight:400; text-align:left; white-space:normal; }
.p-tip h4 { margin:0 0 6px; font-size:13px; color:#5b6abf; font-weight:700; }
.p-tip h5 { margin:10px 0 4px; font-size:11px; color:#888; font-weight:700; text-transform:uppercase; letter-spacing:.4px; }
.p-tip .pk { display:inline-block; background:#eee; border:1px solid #ccc; border-radius:3px; padding:0 5px; font-family:monospace; font-size:10px; min-width:16px; text-align:center; }
.p-tip table { width:100%; border-collapse:collapse; }
.p-tip td { padding:2px 0; vertical-align:top; }
.p-tip td:first-child { width:80px; white-space:nowrap; padding-right:8px; }
.ub { background:none; border:none; cursor:pointer; font-size:12px; opacity:.35; padding:0 2px; transition:opacity .15s; }
.inu-dirty { font-size:11px; font-weight:600; color:#e53935; padding:3px 10px; border-radius:4px; background:rgba(229,57,53,.1); white-space:nowrap; animation:inu-pulse-dirty 2s ease-in-out infinite; }
@keyframes inu-pulse-dirty { 50% { opacity:.6; } }
tr.tag.inu-saving > td { opacity:.6; }
tr.tag.inu-sel > td { background:#e3f2fd !important; }
tr.tag.inu-dupe > td:nth-child(${OFF+3}) { background:rgba(255,152,0,.25) !important; font-weight:600; }
.inu-dupe-info { font-size:11px; font-weight:600; color:#e65100; padding:3px 10px; border-radius:4px; background:rgba(255,152,0,.15); cursor:pointer; white-space:nowrap; }
.inu-dupe-info:hover { background:rgba(255,152,0,.25); }
.inu-src-pill { padding:3px 9px; border-radius:3px; font-size:10px; font-weight:600; color:#fff; background:#b84700; cursor:pointer; align-self:center; display:inline-flex; align-items:center; gap:5px; text-decoration:none; white-space:nowrap; margin-right:8px; }
.inu-src-pill:hover { background:#e65100; color:#fff; }
.tog-ico.inu-spin { animation:inu-spin 1s linear infinite; }
@keyframes inu-spin { 100% { transform:rotate(360deg); } }
.ub:hover { opacity:1; }
.mo { position:fixed; top:0; left:0; right:0; bottom:0; background:rgba(0,0,0,.4); z-index:100000; display:flex; align-items:flex-start; justify-content:center; overflow-y:auto; padding:20px 0; }
.inu-log-panel { border-top:1px solid rgba(255,255,255,.08); flex-shrink:0; display:flex; flex-direction:column; font-family:monospace; margin-top:auto; }
.inu-log-hdr { display:flex; align-items:center; gap:4px; padding:4px 8px; background:#111827; color:#9ca3af; font-size:10px; font-weight:600; cursor:pointer; user-select:none; flex-shrink:0; }
.inu-log-hdr:hover { background:#1f2937; }
.inu-log-title { display:flex; align-items:center; gap:5px; flex-shrink:0; }
.inu-log-filters { display:flex; gap:2px; margin-left:6px; }
.inu-lf { background:transparent; border:1px solid rgba(255,255,255,.15); color:#6b7280; border-radius:3px; padding:1px 5px; font-size:9px; font-family:monospace; cursor:pointer; font-weight:600; }
.inu-lf:hover { border-color:rgba(255,255,255,.35); color:#d1d5db; }
.inu-lf.active { background:#374151; border-color:rgba(255,255,255,.35); color:#f9fafb; }
.inu-log-spacer { flex:1; }
.inu-log-clear,.inu-log-tog { background:none; border:none; color:#6b7280; cursor:pointer; padding:0 3px; font-size:11px; line-height:1; }
.inu-log-clear:hover,.inu-log-tog:hover { color:#d1d5db; }
.inu-log-body { background:#0f172a; overflow-y:auto; max-height:180px; min-height:40px; }
.inu-log-empty { color:#374151; font-size:10px; font-family:monospace; padding:8px 10px; font-style:italic; }
.inu-log-row { display:flex; align-items:flex-start; gap:5px; padding:2px 8px; border-bottom:1px solid rgba(255,255,255,.04); line-height:1.45; font-size:10px; }
.inu-log-row:last-child { border-bottom:none; }
.inu-log-ts { color:#9ca3af; flex-shrink:0; white-space:nowrap; }
.inu-log-src { flex-shrink:0; font-size:8px; padding:1px 3px; border-radius:2px; line-height:1.6; margin-top:1px; }
.inu-log-src.inu { background:#1e3a5f; color:#60a5fa; }
.inu-log-src.wp { background:#1c1c1c; color:#6b7280; }
.inu-log-msg { color:#d1d5db; white-space:pre-wrap; word-break:break-all; }
.inu-log-row.success .inu-log-msg { color:#86efac; }
.inu-log-row.error .inu-log-msg { color:#fca5a5; }
.inu-log-row.warning .inu-log-msg { color:#fcd34d; }
.mb { background:#fff; border-radius:8px; padding:20px; min-width:540px; max-width:90vw; width:900px; box-shadow:0 8px 32px rgba(0,0,0,.3); margin:auto; }
.mb h3 { margin:0 0 12px; font-size:15px; }
.mb label { display:block; margin:6px 0 2px; font-size:11px; font-weight:600; color:#555; }
.mb input,.mb textarea { width:100%; padding:5px 7px; border:1px solid #ccc; border-radius:4px; font-size:12px; box-sizing:border-box; }
.mb textarea { height:120px; font-family:monospace; font-size:11px; }
.mb .fr { display:flex; gap:8px; } .mb .fr>div { flex:1; }
.mb .bt { display:flex; gap:8px; margin-top:14px; justify-content:flex-end; }
.mb .bt button { padding:6px 14px; border:none; border-radius:4px; cursor:pointer; font-size:12px; font-weight:600; }
.mb .bok { padding:6px 14px; border:none; border-radius:4px; cursor:pointer; font-size:12px; font-weight:600; background:#5b6abf; color:#fff; } .mb .bok:hover { background:#4a58a8; }
.mb .bx { background:#eee; color:#333; } .mb .bx:hover { background:#ddd; }
.kr-ov { position:fixed; inset:0; background:#f5f6fa; z-index:100001; overflow-y:auto; font-size:13px; }
.kr-hdr { display:flex; align-items:center; gap:12px; padding:12px 24px; background:#fff; border-bottom:2px solid #e0e0e0; position:sticky; top:0; z-index:2; }
.kr-hdr h2 { margin:0; font-size:15px; font-weight:700; color:#222; }
.kr-hdr .kr-sub { font-size:11px; color:#888; }
.kr-hdr .kr-close { margin-left:auto; background:none; border:none; cursor:pointer; font-size:18px; color:#666; padding:2px 6px; border-radius:4px; }
.kr-hdr .kr-close:hover { background:#f0f0f0; color:#333; }
.kr-body { padding:20px 24px; max-width:1400px; margin:0 auto; }
.kr-stats { display:flex; gap:10px; margin-bottom:16px; }
.kr-stat { flex:1; background:#fff; border-radius:8px; padding:12px 16px; border:1px solid #e8e8e8; text-align:center; }
.kr-stat .kn { font-size:26px; font-weight:700; color:#1b5e20; line-height:1.1; }
.kr-stat .kn.warn { color:#e65100; }
.kr-stat .kn.info { color:#5b6abf; }
.kr-stat .kn.gray { color:#90a4ae; }
.kr-stat .kl { font-size:11px; color:#666; margin-top:3px; }
.kr-stat .ks { font-size:10px; color:#aaa; }
.kr-charts { display:grid; grid-template-columns:1fr 1fr 1fr; gap:14px; margin-bottom:14px; }
.kr-card { background:#fff; border-radius:8px; padding:14px 16px; border:1px solid #e8e8e8; }
.kr-card h4 { margin:0 0 10px; font-size:10px; font-weight:700; color:#888; text-transform:uppercase; letter-spacing:.5px; }
.kr-br { display:flex; align-items:center; gap:7px; margin-bottom:5px; }
.kr-br-lbl { font-size:11px; color:#333; width:100px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; flex-shrink:0; }
.kr-br-trk { flex:1; height:13px; background:#f0f0f0; border-radius:3px; overflow:hidden; min-width:40px; }
.kr-br-fill { height:100%; border-radius:3px; min-width:3px; }
.kr-br-cnt { font-size:11px; color:#666; width:32px; text-align:right; flex-shrink:0; }
.kr-cov { background:#fff; border-radius:8px; padding:14px 16px; border:1px solid #e8e8e8; margin-bottom:14px; }
.kr-cov h4 { margin:0 0 10px; font-size:10px; font-weight:700; color:#888; text-transform:uppercase; letter-spacing:.5px; }
.kr-covr { display:flex; align-items:center; gap:10px; margin-bottom:7px; }
.kr-covr-lbl { font-size:12px; font-weight:600; width:50px; flex-shrink:0; }
.kr-covr-trk { flex:1; height:18px; background:#f0f0f0; border-radius:4px; overflow:hidden; }
.kr-covr-fill { height:100%; border-radius:4px; min-width:3px; }
.kr-covr-pct { font-size:12px; font-weight:700; width:38px; text-align:right; flex-shrink:0; }
.kr-covr-ct { font-size:11px; color:#888; }
.kr-hyg { background:#fff; border-radius:8px; padding:14px 16px; border:1px solid #e8e8e8; }
.kr-hyg h4 { margin:0 0 10px; font-size:10px; font-weight:700; color:#888; text-transform:uppercase; letter-spacing:.5px; }
.kr-hyg-sec { margin-bottom:6px; }
.kr-hyg-hdr { display:flex; align-items:center; gap:8px; padding:7px 10px; background:#f5f5f8; border-radius:5px; cursor:pointer; user-select:none; }
.kr-hyg-hdr:hover { background:#ebebf5; }
.kr-hyg-ico { font-size:10px; color:#aaa; transition:transform .15s; flex-shrink:0; }
.kr-hyg-ttl { font-size:12px; font-weight:600; flex:1; color:#333; }
.kr-hyg-badge { font-size:10px; font-weight:700; padding:1px 8px; border-radius:10px; background:#ef4444; color:#fff; }
.kr-hyg-badge.ok { background:#4caf50; }
.kr-hyg-badge.warn { background:#ff9800; }
.kr-hyg-badge.info { background:#5b6abf; }
.kr-hyg-body { padding:8px 10px 4px; display:none; }
.kr-hyg-body.open { display:block; }
.kr-tag { display:inline-block; background:#f0f0f0; border-radius:3px; padding:1px 6px; margin:2px; font-size:10px; font-family:monospace; color:#333; }
.kr-sys-tbl { width:100%; border-collapse:collapse; font-size:12px; }
.kr-sys-tbl th { text-align:left; padding:5px 8px; font-size:10px; font-weight:700; color:#888; border-bottom:2px solid #eee; white-space:nowrap; background:#fff; }
.kr-sys-tbl td { padding:4px 8px; border-bottom:1px solid #f4f4f4; vertical-align:middle; }
.kr-sys-tbl tbody tr:hover td { background:#f8f9ff; }
.kr-sys-wrap { overflow-x:auto; max-height:420px; overflow-y:auto; }
.tpl-modal .mb { width:980px; max-width:95vw; }
.tpl-hdr { display:flex; align-items:center; gap:10px; margin-bottom:12px; }
.tpl-hdr h3 { margin:0; font-size:15px; flex:1; }
.tpl-hdr .tpl-close { background:none; border:none; font-size:18px; cursor:pointer; color:#666; padding:2px 8px; border-radius:4px; }
.tpl-hdr .tpl-close:hover { background:#f0f0f0; }
.tpl-pickers { display:grid; grid-template-columns:1.2fr 1.5fr 1fr; gap:10px; margin-bottom:14px; padding-bottom:12px; border-bottom:1px solid #eee; }
.tpl-pickers > div { min-width:0; }
.tpl-pickers label { display:block; font-size:10px; font-weight:700; color:#555; text-transform:uppercase; letter-spacing:.4px; margin-bottom:3px; }
.tpl-pickers select, .tpl-pickers input { width:100%; padding:5px 7px; border:1px solid #999; border-radius:4px; font-size:13px; font-weight:500; box-sizing:border-box; background:#fff; color:#222 !important; }
.tpl-pickers select:disabled, .tpl-pickers input:disabled { background:#f5f5f5; color:#666 !important; }
.tpl-pickers select:focus, .tpl-pickers input:focus { outline:2px solid #5b6abf; outline-offset:1px; border-color:#5b6abf; }
.tpl-status { font-size:11px; color:#5b6abf; padding:8px 12px; background:#eef1ff; border-radius:4px; margin-bottom:10px; display:none; }
.tpl-status.tpl-err { background:#fef2f2; color:#b91c1c; }
.tpl-cfg { max-height:52vh; overflow-y:auto; padding-right:8px; margin-bottom:12px; }
.tpl-dropdown-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(200px, 1fr)); gap:8px 12px; padding:10px 12px; margin-bottom:14px; background:#f8f9ff; border:1px solid #e3e6f3; border-radius:6px; }
.tpl-dd-cell { display:flex; flex-direction:column; }
.tpl-dd-lbl { font-size:10px; font-weight:700; color:#555; text-transform:uppercase; letter-spacing:.4px; margin-bottom:3px; }
.tpl-dd { padding:5px 7px; border:1px solid #999; border-radius:4px; font-size:13px; font-weight:500; background:#fff; color:#222 !important; cursor:pointer; }
.tpl-dd:focus { outline:2px solid #5b6abf; outline-offset:1px; border-color:#5b6abf; }
.tpl-dd-help { font-size:10px; color:#888; margin-top:2px; font-style:italic; }
.tpl-sec { margin-bottom:14px; padding-bottom:10px; border-bottom:1px dashed #eee; }
.tpl-sec:last-child { border-bottom:none; }
.tpl-sec-lbl { font-size:12px; font-weight:700; color:#333; margin-bottom:3px; }
.tpl-sec-help { font-size:10px; color:#888; margin-bottom:6px; font-style:italic; }
.tpl-opts { display:flex; flex-wrap:wrap; gap:8px 12px; align-items:stretch; }
.tpl-opt { display:inline-flex !important; align-items:center; gap:8px; font-size:12px; line-height:18px !important; cursor:pointer; user-select:none; padding:6px 10px !important; margin:0 !important; border-radius:4px; transition:background .1s; min-height:30px; box-sizing:border-box; }
.tpl-opt:hover { background:#f4f6ff; }
.tpl-opt input[type=checkbox] { margin:0 !important; padding:0 !important; cursor:pointer; width:15px !important; height:15px !important; flex-shrink:0; vertical-align:middle; accent-color:#5b6abf; }
.tpl-opt > span { line-height:18px; display:inline-block; vertical-align:middle; }
.tpl-opt.tpl-chk-on { background:#e3f2fd; font-weight:600; }
.tpl-opt .tpl-info-btn { margin-left:4px !important; align-self:center !important; }
.tpl-resolved { margin-top:10px; border-top:1px solid #eee; padding-top:10px; }
.tpl-res-hdr { display:flex; align-items:center; gap:8px; font-size:12px; font-weight:600; color:#333; cursor:pointer; user-select:none; padding:5px 0; }
.tpl-res-hdr .tpl-res-count { color:#5b6abf; font-size:13px; }
.tpl-res-hdr .tpl-res-tri { font-size:10px; color:#888; transition:transform .15s; }
.tpl-res-hdr.open .tpl-res-tri { transform:rotate(90deg); }
.tpl-res-body { display:none; max-height:32vh; overflow-y:auto; margin-top:6px; border:1px solid #eee; border-radius:4px; }
.tpl-res-body.open { display:block; }
.tpl-res-search { padding:6px 8px; border-bottom:1px solid #eee; background:#fafafa; }
.tpl-res-search input { width:100%; padding:4px 7px; border:1px solid #ccc; border-radius:3px; font-size:11px; box-sizing:border-box; }
.tpl-res-row { display:grid !important; grid-template-columns:18px 260px 80px 60px 1fr; align-items:center; gap:14px; padding:6px 12px !important; font-size:12px; border-bottom:1px solid #eee; cursor:pointer; margin:0 !important; }
.tpl-res-row:hover { background:#f4f6ff; }
.tpl-res-row input { margin:0 !important; cursor:pointer; width:14px; height:14px; }
.tpl-res-row.tpl-r-off { opacity:.4; }
.tpl-res-row.tpl-r-off .tpl-res-name, .tpl-res-row.tpl-r-off .tpl-res-desc { text-decoration:line-through; }
.tpl-res-name { font-family:monospace; font-weight:600; color:#222; font-size:12px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.tpl-res-addr { font-family:monospace; color:#666; font-size:11px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.tpl-res-unit { color:#5b6abf; font-size:10px; font-weight:600; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.tpl-res-desc { color:#555; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.tpl-info-btn { display:inline-block !important; width:15px !important; height:15px !important; line-height:15px !important; text-align:center; border-radius:50%; background:#5b6abf; color:#fff !important; font-size:10px !important; font-weight:700 !important; font-family:Arial, sans-serif !important; font-style:normal !important; cursor:help; user-select:none; flex-shrink:0; margin:0 0 0 4px !important; padding:0 !important; box-sizing:border-box; vertical-align:middle; transition:background .12s; }
.tpl-info-btn:hover, .tpl-info-btn.open { background:#3a4ba0; }
.tpl-dd-row { display:flex; align-items:center; gap:6px; }
.tpl-dd-row .tpl-dd { flex:1; min-width:0; }
.tpl-info-popup { position:absolute; z-index:100002; background:#fff; border:1px solid #5b6abf; border-radius:6px; box-shadow:0 6px 24px rgba(0,0,0,.2); padding:0; max-width:680px; max-height:50vh; overflow:auto; font-size:11px; color:#333; }
.tpl-info-popup .tpl-info-hdr { padding:6px 12px; background:#5b6abf; color:#fff; font-weight:700; font-size:11px; position:sticky; top:0; }
.tpl-info-popup .tpl-info-tbl { border-collapse:collapse; width:100%; }
.tpl-info-popup .tpl-info-tbl th { text-align:left; padding:5px 10px; background:#f4f6ff; color:#555; font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:.4px; border-bottom:1px solid #e3e6f3; position:sticky; top:24px; }
.tpl-info-popup .tpl-info-tbl td { padding:4px 10px; border-bottom:1px solid #f4f4f4; vertical-align:top; }
.tpl-info-popup .tpl-info-tbl tr:last-child td { border-bottom:none; }
.tpl-info-popup .tpl-info-tbl code { font-family:monospace; font-size:11px; color:#222; background:#f8f9ff; padding:1px 4px; border-radius:2px; }
.tpl-info-popup .tpl-info-type { display:inline-block; padding:1px 6px; border-radius:3px; background:#eef1ff; color:#3a4ba0; font-size:10px; font-weight:600; white-space:nowrap; }
.tpl-info-popup .tpl-info-empty { padding:12px 14px; font-size:11px; color:#888; font-style:italic; }
.tpl-foot { display:flex; align-items:center; gap:8px; margin-top:12px; padding-top:12px; border-top:1px solid #eee; justify-content:flex-end; }
.tpl-foot .tpl-prog { flex:1; font-size:11px; color:#666; display:none; }
.tpl-foot .tpl-prog.on { display:block; }
.tpl-foot .tpl-prog-bar { height:4px; background:#eee; border-radius:2px; overflow:hidden; margin-top:3px; }
.tpl-foot .tpl-prog-fill { height:100%; background:#5b6abf; width:0%; transition:width .15s; }
.tpl-foot button { padding:6px 14px; border:none; border-radius:4px; cursor:pointer; font-size:12px; font-weight:600; }
.tpl-foot .tpl-cancel { background:#eee; color:#333; }
.tpl-foot .tpl-cancel:hover { background:#ddd; }
.tpl-foot .tpl-add { background:#5b6abf; color:#fff; }
.tpl-foot .tpl-add:hover { background:#4a58a8; }
.tpl-foot .tpl-add:disabled { background:#ccc; cursor:not-allowed; }
`;
        document.head.appendChild(s);
    }

    // ============================================================
    // HELPERS
    // ============================================================
    function toast(m, ms=2000) {
        _ownToast = true;
        const t = unsafeWindow.toastr || window.toastr;
        if (t) t.info(m, '', { timeOut: ms, positionClass: 'toast-bottom-right' });
        _ownToast = false;
    }
    function toastOk(m) {
        _ownToast = true;
        const t = unsafeWindow.toastr || window.toastr;
        if (t) t.success(m, '', { timeOut: 2000, positionClass: 'toast-bottom-right' });
        _ownToast = false;
    }
    function toastErr(m) {
        _ownToast = true;
        const t = unsafeWindow.toastr || window.toastr;
        if (t) t.error(m, '', { timeOut: 4000, positionClass: 'toast-bottom-right' });
        _ownToast = false;
    }

    // ============================================================
    // NOTIFICATION LOG
    // ============================================================
    function _logRenderRow(entry) {
        if (!_logEntriesEl) return;
        if (_logFilter !== 'all' && entry.level !== _logFilter) return;
        const d = entry.ts;
        const ts = String(d.getHours()).padStart(2,'0') + ':' + String(d.getMinutes()).padStart(2,'0') + ':' + String(d.getSeconds()).padStart(2,'0');
        const row = document.createElement('div');
        row.className = 'inu-log-row ' + entry.level;
        row.innerHTML =
            '<span class="inu-log-ts">' + ts + '</span>' +
            '<span class="inu-log-src ' + entry.src + '">' + entry.src + '</span>' +
            '<span class="inu-log-msg">' + escHtml(entry.msg) + '</span>';
        _logEntriesEl.appendChild(row);
        if (_logBodyEl) _logBodyEl.scrollTop = _logBodyEl.scrollHeight;
        // Remove empty placeholder if present
        const empty = _logEntriesEl.querySelector('.inu-log-empty');
        if (empty) empty.remove();
    }

    function logAppend(level, msg, src) {
        const entry = { ts: new Date(), level, msg, src: src || 'wp' };
        _logEntries.push(entry);
        if (_logEntries.length > LOG_MAX) _logEntries.shift();
        _saveLog();
        _logRenderRow(entry);
    }

    // Silent log — writes to log panel only, no toast
    function logInfo(msg) { logAppend('info', msg, 'inu'); }

    function _rebuildLog() {
        if (!_logEntriesEl) return;
        _logEntriesEl.innerHTML = '';
        const visible = _logFilter === 'all' ? _logEntries : _logEntries.filter(e => e.level === _logFilter);
        if (!visible.length) {
            _logEntriesEl.innerHTML = '<div class="inu-log-empty">Ingen aktivitet ännu</div>';
            return;
        }
        visible.forEach(_logRenderRow);
    }

    function hookToastr() {
        const tr = unsafeWindow.toastr || window.toastr;
        if (!tr || tr._inuHooked) return;
        tr._inuHooked = true;
        ['info','success','error','warning'].forEach(lvl => {
            const orig = tr[lvl].bind(tr);
            tr[lvl] = function(msg, title, opts) {
                const plain = String(msg || '').replace(/<br\s*\/?>/gi,' | ').replace(/<[^>]+>/g,'').trim();
                if (plain) logAppend(lvl, plain, _ownToast ? 'inu' : 'wp');
                return orig(msg, title, opts);
            };
        });
    }

    function initLogPanel() {
        if (document.getElementById('inu-log-panel')) return;
        // Inject log CSS if injectStyles() hasn't run (non-tag/device pages)
        if (!document.getElementById('inu-log-style')) {
            const s = document.createElement('style');
            s.id = 'inu-log-style';
            s.textContent = '.inu-log-panel{border-top:1px solid rgba(255,255,255,.08);flex-shrink:0;display:flex;flex-direction:column;font-family:monospace;margin-top:auto}.inu-log-hdr{display:flex;align-items:center;gap:4px;padding:4px 8px;background:#111827;color:#9ca3af;font-size:10px;font-weight:600;cursor:pointer;user-select:none;flex-shrink:0}.inu-log-hdr:hover{background:#1f2937}.inu-log-title{display:flex;align-items:center;gap:5px;flex-shrink:0}.inu-log-filters{display:flex;gap:2px;margin-left:6px}.inu-lf{background:transparent;border:1px solid rgba(255,255,255,.15);color:#6b7280;border-radius:3px;padding:1px 5px;font-size:9px;font-family:monospace;cursor:pointer;font-weight:600}.inu-lf:hover{border-color:rgba(255,255,255,.35);color:#d1d5db}.inu-lf.active{background:#374151;border-color:rgba(255,255,255,.35);color:#f9fafb}.inu-log-spacer{flex:1}.inu-log-clear,.inu-log-tog{background:none;border:none;color:#6b7280;cursor:pointer;padding:0 3px;font-size:11px;line-height:1}.inu-log-clear:hover,.inu-log-tog:hover{color:#d1d5db}.inu-log-body{background:#0f172a;overflow-y:auto;max-height:180px;min-height:40px}.inu-log-empty{color:#374151;font-size:10px;font-family:monospace;padding:8px 10px;font-style:italic}.inu-log-row{display:flex;align-items:flex-start;gap:5px;padding:2px 8px;border-bottom:1px solid rgba(255,255,255,.04);line-height:1.45;font-size:10px}.inu-log-row:last-child{border-bottom:none}.inu-log-ts{color:#9ca3af;flex-shrink:0;white-space:nowrap}.inu-log-src{flex-shrink:0;font-size:8px;padding:1px 3px;border-radius:2px;line-height:1.6;margin-top:1px}.inu-log-src.inu{background:#1e3a5f;color:#60a5fa}.inu-log-src.wp{background:#1c1c1c;color:#6b7280}.inu-log-msg{color:#d1d5db;white-space:pre-wrap;word-break:break-all}.inu-log-row.success .inu-log-msg{color:#86efac}.inu-log-row.error .inu-log-msg{color:#fca5a5}.inu-log-row.warning .inu-log-msg{color:#fcd34d}';
            document.head.appendChild(s);
        }
        // Target the tree nav (second child of nav.sidebar) so panel sits below the tree.
        const treeNav = document.querySelector('nav.sidebar > nav');
        if (!treeNav) return;
        // The SPA renders the tree UL asynchronously, so we may append the panel before the
        // UL exists. Use a MutationObserver to re-pin the layout whenever children change.
        treeNav.style.cssText += ';display:flex;flex-direction:column;overflow:hidden;';
        const _pinLog = () => {
            const ul = treeNav.querySelector('ul');
            const p  = document.getElementById('inu-log-panel');
            if (ul) { ul.style.flex = '1'; ul.style.overflowY = 'auto'; ul.style.minHeight = '0'; }
            if (p && treeNav.lastElementChild !== p) treeNav.appendChild(p); // keep panel last
        };
        new MutationObserver(_pinLog).observe(treeNav, { childList: true });
        _pinLog();
        const collapsed = _logEntries.length === 0 ? true : GM_getValue('inu_log_collapsed', false);
        const panel = document.createElement('div');
        panel.id = 'inu-log-panel';
        panel.className = 'inu-log-panel';
        panel.innerHTML =
            '<div class="inu-log-hdr">' +
              '<span class="inu-log-title"><i class="fa fa-terminal"></i> Aktivitetslogg</span>' +
              '<span class="inu-log-filters">' +
                '<button class="inu-lf active" data-f="all">Alla</button>' +
                '<button class="inu-lf" data-f="success">OK</button>' +
                '<button class="inu-lf" data-f="error">Fel</button>' +
                '<button class="inu-lf" data-f="warning">Varning</button>' +
              '</span>' +
              '<span class="inu-log-spacer"></span>' +
              '<button class="inu-log-clear" title="Rensa logg"><i class="fa fa-trash"></i></button>' +
              '<button class="inu-log-tog" title="Dölj/Visa"><i class="fa fa-chevron-' + (collapsed ? 'up' : 'down') + '"></i></button>' +
            '</div>' +
            '<div class="inu-log-body">' +
              '<div class="inu-log-entries"><div class="inu-log-empty">Ingen aktivitet ännu</div></div>' +
            '</div>';
        treeNav.appendChild(panel);
        _logBodyEl = panel.querySelector('.inu-log-body');
        _logEntriesEl = panel.querySelector('.inu-log-entries');
        if (collapsed) _logBodyEl.style.display = 'none';
        // Render any entries buffered before panel was ready
        if (_logEntries.length) _rebuildLog();
        // Toggle collapse
        const togBtn = panel.querySelector('.inu-log-tog');
        panel.querySelector('.inu-log-hdr').addEventListener('click', () => {
            const isCollapsed = _logBodyEl.style.display === 'none';
            _logBodyEl.style.display = isCollapsed ? '' : 'none';
            togBtn.innerHTML = '<i class="fa fa-chevron-' + (isCollapsed ? 'down' : 'up') + '"></i>';
            GM_setValue('inu_log_collapsed', !isCollapsed);
        });
        // Clear button — stop propagation so it doesn't toggle collapse
        panel.querySelector('.inu-log-clear').addEventListener('click', e => {
            e.stopPropagation();
            _logEntries.length = 0;
            sessionStorage.removeItem(LOG_SS_KEY);
            _logEntriesEl.innerHTML = '<div class="inu-log-empty">Ingen aktivitet ännu</div>';
        });
        // Filter buttons
        panel.querySelectorAll('.inu-lf').forEach(btn => {
            btn.addEventListener('click', e => {
                e.stopPropagation();
                _logFilter = btn.dataset.f;
                panel.querySelectorAll('.inu-lf').forEach(b => b.classList.toggle('active', b.dataset.f === _logFilter));
                _rebuildLog();
            });
        });
    }

    // Column indices (shifted by CFG.colOffset because our columns are prepended)
    const OFF = CFG.colOffset;
    const C = { NAME:OFF, IO:OFF+1, ADDR:OFF+2, DTYPE:OFF+3, RMIN:OFF+4, RMAX:OFF+5, EMIN:OFF+6, EMAX:OFF+7, UNIT:OFF+8, FMT:OFF+9, DESC:OFF+10 };

    // HTML escape helper — used anywhere values are interpolated into innerHTML
    function escHtml(val) { const d = document.createElement('div'); d.textContent = String(val ?? ''); return d.innerHTML; }

    function scl(row) {
        return {
            rawmin: row.cells[C.RMIN]?.textContent?.trim()||'0',
            rawmax: row.cells[C.RMAX]?.textContent?.trim()||'0',
            engmin: row.cells[C.EMIN]?.textContent?.trim()||'0',
            engmax: row.cells[C.EMAX]?.textContent?.trim()||'0',
            unit:   row.cells[C.UNIT]?.textContent?.trim()||'',
            format: row.cells[C.FMT]?.textContent?.trim()||'0',
        };
    }
    function fullSnap(row) {
        return {
            io:     row.cells[C.IO]?.textContent?.trim()||'',
            addr:   row.cells[C.ADDR]?.textContent?.trim()||'',
            dtype:  row.cells[C.DTYPE]?.textContent?.trim()||'',
            rawmin: row.cells[C.RMIN]?.textContent?.trim()||'0',
            rawmax: row.cells[C.RMAX]?.textContent?.trim()||'0',
            engmin: row.cells[C.EMIN]?.textContent?.trim()||'0',
            engmax: row.cells[C.EMAX]?.textContent?.trim()||'0',
            unit:   row.cells[C.UNIT]?.textContent?.trim()||'',
            format: row.cells[C.FMT]?.textContent?.trim()||'0',
            desc:   row.cells[C.DESC]?.textContent?.trim()||'',
        };
    }
    function snapDiff(a, b) {
        const fields=[];
        if(a.io!==b.io) fields.push('IO-Enhet');
        if(a.addr!==b.addr) fields.push('Adress');
        if(a.dtype!==b.dtype) fields.push('Datatyp');
        if(a.rawmin!==b.rawmin||a.rawmax!==b.rawmax) fields.push('Rå-skalning');
        if(a.engmin!==b.engmin||a.engmax!==b.engmax) fields.push('Vy-skalning');
        if(a.unit!==b.unit) fields.push('Enhet');
        if(a.format!==b.format) fields.push('Format');
        if(a.desc!==b.desc) fields.push('Beskrivning');
        return fields;
    }
    function unconf(row) {
        const dt = row.cells[C.DTYPE]?.textContent?.trim();
        if (dt === 'DIGITAL') return false;
        const s = scl(row);
        return s.rawmin==='0'&&s.rawmax==='0'&&s.engmin==='0'&&s.engmax==='0';
    }
    function updCells(row, p) {
        if(row.cells[C.RMIN]) row.cells[C.RMIN].textContent=p.rawmin;
        if(row.cells[C.RMAX]) row.cells[C.RMAX].textContent=p.rawmax;
        if(row.cells[C.EMIN]) row.cells[C.EMIN].textContent=p.engmin;
        if(row.cells[C.EMAX]) row.cells[C.EMAX].textContent=p.engmax;
        if(row.cells[C.UNIT]) row.cells[C.UNIT].textContent=p.unit;
        if(row.cells[C.FMT])  row.cells[C.FMT].textContent=p.format;
    }
    function colorRow(r) { if(unconf(r)) r.classList.add('ru'); else r.classList.remove('ru'); }

    // ============================================================
    // API LAYER
    // ============================================================
    function encodeTag(tagName) { return tagName.replace(/_/g, '-5F-'); }

    function serializeForm(form) {
        const fd = new FormData();
        for (const el of form.querySelectorAll('input,textarea,select')) {
            if (!el.name || el.type === 'file') continue;
            if (el.type === 'checkbox') { if (el.checked) fd.append(el.name, el.value || 'on'); continue; }
            if (el.type === 'radio') { if (el.checked) fd.append(el.name, el.value); continue; }
            fd.append(el.name, el.value);
        }
        return fd;
    }

    async function fetchFormAndSave(tagName, mutator) {
        const r = await fetch('/tag/ActionEdit?show=1&type=tag&tag=' + encodeTag(tagName));
        if (!r.ok) throw new Error(`Load "${tagName}": HTTP ${r.status}`);
        const doc = new DOMParser().parseFromString(await r.text(), 'text/html');
        const form = doc.querySelector('form#frmtag');
        if (!form) throw new Error(`No form for "${tagName}"`);
        const fd = serializeForm(form);
        mutator(fd, doc);
        const s = await fetch('/tag/actionedit', { method: 'POST', body: fd });
        if (!s.ok) throw new Error(`Save "${tagName}": HTTP ${s.status}`);
        return doc;
    }

    async function apiSave(tagName, preset) {
        await fetchFormAndSave(tagName, fd => {
            fd.set('rawmin', preset.rawmin); fd.set('rawmax', preset.rawmax);
            fd.set('engmin', preset.engmin); fd.set('engmax', preset.engmax);
            fd.set('unit', preset.unit);     fd.set('format', preset.format);
        });
    }

    // Larmklass + isalarm + istrend cache
    const larmCache = {}, alarmCache = {}, trendCache = {};

    async function loadLarmklass(tagName, sel, chk, tChk, row) {
        function syncIcon(cls, on) {
            if (!row) return;
            const img = row.querySelector('img.' + cls);
            if (img) img.style.display = on ? '' : 'none';
        }
        if (larmCache[tagName] !== undefined) {
            sel.value = larmCache[tagName];
            if (sel.value !== larmCache[tagName]) { const o = document.createElement('option'); o.value = larmCache[tagName]; o.textContent = larmCache[tagName]; sel.insertBefore(o, sel.firstChild); sel.value = larmCache[tagName]; }
            if (chk) { chk.checked = !!alarmCache[tagName]; syncIcon('alarmlink', chk.checked); }
            if (tChk) { tChk.checked = !!trendCache[tagName]; syncIcon('trendlink', tChk.checked); }
            return;
        }
        try {
            const r = await fetch('/tag/ActionEdit?show=1&type=tag&tag=' + encodeTag(tagName));
            if (!r.ok) return;
            const doc = new DOMParser().parseFromString(await r.text(), 'text/html');
            const field = doc.querySelector('textarea[name="p"],input[name="p"],select[name="p"]');
            if (field) {
                const v = field.value || '';
                larmCache[tagName] = v;
                if (![...sel.options].some(o => o.value === v)) { const o = document.createElement('option'); o.value = v; o.textContent = v; sel.insertBefore(o, sel.firstChild); }
                sel.value = v;
            }
            const aChk = doc.querySelector('input[name="isalarm"][type="checkbox"]');
            alarmCache[tagName] = aChk ? aChk.checked : false;
            if (chk) { chk.checked = alarmCache[tagName]; syncIcon('alarmlink', chk.checked); }
            const tCk = doc.querySelector('input[name="istrend"][type="checkbox"]');
            trendCache[tagName] = tCk ? tCk.checked : false;
            if (tChk) { tChk.checked = trendCache[tagName]; syncIcon('trendlink', tChk.checked); }
        } catch (e) { console.warn('[INU]', 'loadLarmklass failed for', tagName, e); }
    }

    async function saveLarmklass(tagName, value) {
        try {
            await fetchFormAndSave(tagName, fd => fd.set('p', value));
            larmCache[tagName] = value;
            toast(`Larmklass ${value || '—'} → ${tagName}`);
        } catch (e) { toastErr(`Larmklass fel: ${e.message}`); }
    }

    function setCheckboxField(fd, name, enabled) {
        fd.delete(name);
        if (enabled) fd.append(name, '1');
        fd.append(name, '0');
    }

    async function saveIsAlarm(tagName, enabled) {
        try {
            await fetchFormAndSave(tagName, fd => setCheckboxField(fd, 'isalarm', enabled));
            alarmCache[tagName] = enabled;
            toast(`Larm ${enabled ? 'PÅ' : 'AV'} → ${tagName}`);
        } catch (e) { toastErr(`Larm fel: ${e.message}`); }
    }

    async function saveIsTrend(tagName, enabled) {
        try {
            await fetchFormAndSave(tagName, fd => setCheckboxField(fd, 'istrend', enabled));
            trendCache[tagName] = enabled;
            toast(`Trend ${enabled ? 'PÅ' : 'AV'} → ${tagName}`);
        } catch (e) { toastErr(`Trend fel: ${e.message}`); }
    }

    async function undoRow(row) {
        const tag = row.cells[C.NAME].textContent.trim();
        const old = getUndo(tag); if(!old) return;
        try {
            await apiSave(tag, old);
            updCells(row, old); colorRow(row); delUndo(tag); updUndo(row); updSummary();
            toast(`↩ Ångrade ${tag}`);
        } catch(e) { toastErr(e.message); }
    }

    function updUndo(row) {
        const tag = row.cells[C.NAME].textContent.trim();
        let ub = row.querySelector('.ub');
        if (getUndo(tag)) {
            if(!ub) {
                ub=document.createElement('button'); ub.className='ub'; ub.title='Ångra'; ub.innerHTML='<i class="fa fa-undo"></i>';
                ub.addEventListener('click',e=>{e.stopPropagation();undoRow(row);});
                row.querySelector('.p-chk')?.insertBefore(ub, row.querySelector('.p-chk')?.firstChild);
            }
        } else { if(ub) ub.remove(); }
    }

    // ============================================================
    // SESSION CHANGE TRACKING
    // ============================================================
    const sessionChanges = {};
    const initialSnapshot = {};

    // ============================================================
    // SELECTION
    // ============================================================
    const selNames = new Set();
    const sel = { get size() { return selNames.size; }, has(r) { return selNames.has(r.cells[C.NAME]?.textContent?.trim()); } };
    function selAll(on) {
        if (!on) selNames.clear(); // clear everything including off-page selections
        document.querySelectorAll('#tagtable tbody tr.tag').forEach(r => {
            const name = r.cells[C.NAME]?.textContent?.trim();
            if (!name) return;
            const vis = on && !r.classList.contains('p-hidden');
            if (vis) selNames.add(name);
            const cb = r.querySelector('.p-chk input'); if (cb) cb.checked = vis;
            r.classList.toggle('inu-sel', vis);
        });
        updToolbar();
    }
    function getSelectedRows() {
        const rows = [];
        document.querySelectorAll('#tagtable tbody tr.tag').forEach(r => {
            const name = r.cells[C.NAME]?.textContent?.trim();
            if (name && selNames.has(name)) rows.push(r);
        });
        return rows;
    }
    function syncSelCheckboxes() {
        document.querySelectorAll('#tagtable tbody tr.tag').forEach(r => {
            const name = r.cells[C.NAME]?.textContent?.trim();
            if (!name) return;
            const on = selNames.has(name);
            const cb = r.querySelector('.p-chk input');
            if (cb) cb.checked = on;
            r.classList.toggle('inu-sel', on);
        });
    }

    // ============================================================
    // DRAG SELECT
    // ============================================================
    let dragSelecting = false, dragMode = true;
    function initDragSelect() {
        const table = document.getElementById('tagtable');
        if (!table) return;
        table.addEventListener('mousedown', e => {
            const chk = e.target.closest('.p-chk');
            if (!chk) return;
            const row = chk.closest('tr.tag');
            if (!row) return;
            e.preventDefault();
            const name = row.cells[C.NAME]?.textContent?.trim();
            dragMode = !selNames.has(name);
            dragSelecting = true;
            applyDragToRow(row);
        });
        table.addEventListener('mouseover', e => {
            if (!dragSelecting) return;
            const row = e.target.closest('tr.tag');
            if (row) applyDragToRow(row);
        });
        document.addEventListener('mouseup', () => { dragSelecting = false; });
    }
    function applyDragToRow(row) {
        const name = row.cells[C.NAME]?.textContent?.trim();
        if (!name) return;
        if (dragMode) selNames.add(name); else selNames.delete(name);
        const cb = row.querySelector('.p-chk input');
        if (cb) cb.checked = dragMode;
        row.classList.toggle('inu-sel', dragMode);
        updToolbar();
    }

    // ============================================================
    // FILTER
    // ============================================================
    let filterMode = 'all'; // 'all', 'unconf', 'conf', 'digital', 'dupe'
    let filterColVals = {}; // { [colIdx]: Set<string> }
    const COL_FILTER_LABELS = () => ({ [C.IO]: 'IO-Enhet', [C.DTYPE]: 'Datatyp', [C.UNIT]: 'Enhet' });
    function wildcardToPattern(text) {
        // Escape regex special chars, then convert * to .* wildcard
        return text.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*');
    }
    function applyFilter() {
        document.querySelectorAll('#tagtable tbody tr.tag').forEach(r => {
            const dt = r.cells[C.DTYPE]?.textContent?.trim();
            const isDig = dt === 'DIGITAL';
            const isUnc = !isDig && unconf(r);
            const isConf = !isDig && !isUnc;
            const isDupe = r.classList.contains('inu-dupe');
            let show = true;
            if (filterMode === 'unconf') show = isUnc;
            else if (filterMode === 'conf') show = isConf;
            else if (filterMode === 'digital') show = isDig;
            else if (filterMode === 'dupe') show = isDupe;
            if (show) {
                for (const [ci, vals] of Object.entries(filterColVals)) {
                    if (!vals || !vals.size) continue;
                    const val = r.cells[+ci]?.textContent?.trim() || '';
                    if (!vals.has(val)) { show = false; break; }
                }
            }
            if (show) r.classList.remove('p-hidden');
            else r.classList.add('p-hidden');
        });
        if (bDupeInfo) bDupeInfo.classList.toggle('active', filterMode === 'dupe');
    }

    // ============================================================
    // PENDING DELETES
    // ============================================================
    const pendingDeletes = new Set();

    function rowTagName(row) {
        // Read tag name from data attribute (set before button insertion) or from pure text nodes
        if (row.dataset.inuDelTag) return row.dataset.inuDelTag;
        const cell = row.cells[C.NAME];
        if (!cell) return '';
        for (const node of cell.childNodes) {
            if (node.nodeType === Node.TEXT_NODE) {
                const t = node.textContent.trim();
                if (t) return t;
            }
        }
        return cell.textContent.trim();
    }

    function applyDelVisual(row, tag) {
        row.dataset.inuDelTag = tag;
        row.classList.add('inu-del');
        const nameCell = row.cells[C.NAME];
        if (nameCell && !nameCell.querySelector('.inu-del-undo')) {
            const btn = document.createElement('button');
            btn.className = 'inu-del-undo';
            btn.innerHTML = '<i class="fa fa-undo"></i> Ångra';
            btn.title = 'Ångra borttagning';
            btn.addEventListener('click', e => { e.stopPropagation(); unmarkForDelete(row); });
            nameCell.insertBefore(btn, nameCell.firstChild);
        }
    }

    function showSaveButton() {
        document.getElementById('li_wp_mnu_wp_tb_save')?.removeAttribute('style');
    }

    function markForDelete(row) {
        const tag = rowTagName(row);
        if (!tag || pendingDeletes.has(tag)) return;
        pendingDeletes.add(tag);
        applyDelVisual(row, tag);
        showSaveButton();
        updSummary();
    }

    function reapplyPendingDeletes() {
        if (!pendingDeletes.size) return;
        document.querySelectorAll('#tagtable tbody tr.tag').forEach(row => {
            const tag = rowTagName(row);
            if (tag && pendingDeletes.has(tag)) applyDelVisual(row, tag);
        });
    }

    function unmarkForDelete(row) {
        const tag = rowTagName(row);
        if (!tag) return;
        pendingDeletes.delete(tag);
        delete row.dataset.inuDelTag;
        row.classList.remove('inu-del');
        row.querySelector('.inu-del-undo')?.remove();
        updSummary();
    }

    async function deleteTag(tagName) {
        // Step 1: GET ActionDelete to set server-side delete state
        const r = await fetch('/tag/ActionDelete?show=1&type=tag&tag=' + encodeTag(tagName) + '&_=' + Date.now());
        if (!r.ok) throw new Error('HTTP ' + r.status);
        // Step 2: POST to confirm — form sends both type=tag and tag=tagname
        const fd = new FormData();
        fd.append('type', 'tag');
        fd.append('tag', tagName);
        const res = await fetch('/tag/actiondelete', { method: 'POST', body: fd });
        if (!res.ok) throw new Error('HTTP ' + res.status);
    }

    async function fetchCopyBaseParams(sourceName) {
        const r = await fetch('/tag/ActionCopy?show=1&type=tag&tag=' + encodeURIComponent(sourceName) + '&_=' + Date.now());
        if (!r.ok) throw new Error('HTTP ' + r.status);
        const html = await r.text();
        const doc = new DOMParser().parseFromString(html, 'text/html');
        const form = doc.querySelector('form');
        if (!form) throw new Error('Formulär ej hittat i ActionCopy-svar');
        const params = new URLSearchParams();
        form.querySelectorAll('input:not([type="submit"]):not([type="button"]), select, textarea').forEach(el => {
            if (!el.name) return;
            if (el.type === 'checkbox' && !el.checked) return;
            if (el.type === 'radio' && !el.checked) return;
            params.append(el.name, el.value || '');
        });
        return params;
    }

    async function duplicateTag(sourceName, newName, newAddr) {
        const params = await fetchCopyBaseParams(sourceName);
        params.set('name', newName);
        if (newAddr !== undefined && newAddr !== null) params.set('address', newAddr);
        const res = await fetch('/tag/actioncopy', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: params.toString()
        });
        if (!res.ok) throw new Error('HTTP ' + res.status);
    }

    async function duplicateTagBatch(sourceName, copies) {
        // Fetch form once, reuse base params for all copies
        const baseParams = await fetchCopyBaseParams(sourceName);
        let ok = 0, fail = 0;
        for (const { name, addr } of copies) {
            const p = new URLSearchParams(baseParams);
            p.set('name', name);
            if (addr !== undefined && addr !== null) p.set('address', addr);
            try {
                const res = await fetch('/tag/actioncopy', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: p.toString()
                });
                if (!res.ok) throw new Error('HTTP ' + res.status);
                ok++;
            } catch(e) {
                console.error(CFG.logPrefix, 'Batch duplicate failed:', name, e);
                fail++;
            }
        }
        return { ok, fail };
    }

    function openDuplicateDialog(sourceName, sourceAddr) {
        const esc = escHtml;
        function applyHash(pattern, n) {
            return pattern.replace(/#+/g, m => String(n).padStart(m.length, '0'));
        }
        function generateCopies(namePat, addrPat, count, startNum) {
            const hasHash = /#+/.test(namePat);
            if (count > 1 && !hasHash) return null; // require ## for batch
            return Array.from({ length: count }, (_, i) => ({
                name: hasHash ? applyHash(namePat, startNum + i) : namePat,
                addr: /#+/.test(addrPat) ? applyHash(addrPat, startNum + i) : addrPat
            }));
        }

        const m = modal(`
<h3><i class="fa fa-clone"></i> Duplicera tagg</h3>
<p style="font-size:12px;color:#666;margin:0 0 10px;">Källa: <code style="font-size:11px;background:#f0f0f0;padding:1px 4px;border-radius:3px;">${esc(sourceName)}</code></p>
<label>Namnmönster <span style="font-size:10px;color:#999;font-weight:400;">## = sekvensnummer</span></label>
<input id="dup-name" value="${esc(sourceName + '_KOPIA')}" style="width:100%;box-sizing:border-box;font-family:monospace;font-size:12px;margin-bottom:8px;">
<label>Adress <span style="font-size:10px;color:#999;font-weight:400;">## = sekvensnummer</span></label>
<input id="dup-addr" value="${esc(sourceAddr || '')}" style="width:100%;box-sizing:border-box;font-family:monospace;font-size:12px;margin-bottom:8px;">
<div style="display:flex;gap:8px;margin-bottom:8px;">
  <div style="flex:1;"><label>Antal kopior</label><input id="dup-count" type="number" min="1" max="500" value="1" style="width:100%;box-sizing:border-box;font-size:12px;"></div>
  <div style="flex:1;"><label>Startnummer</label><input id="dup-start" value="01" style="width:100%;box-sizing:border-box;font-family:monospace;font-size:12px;"></div>
</div>
<div id="dup-preview" style="background:#f8f8f8;border:1px solid #e0e0e0;border-radius:4px;padding:6px 8px;font-size:11px;font-family:monospace;min-height:32px;margin-bottom:10px;max-height:100px;overflow-y:auto;color:#333;"></div>
<div class="bt"><button class="bx" id="dup-cancel">Avbryt</button><button class="bok" id="dup-ok"><i class="fa fa-clone"></i> <span id="dup-ok-lbl">Duplicera</span></button></div>`);

        const inp     = m.querySelector('#dup-name');
        const addrInp = m.querySelector('#dup-addr');
        const cntInp  = m.querySelector('#dup-count');
        const stInp   = m.querySelector('#dup-start');
        const preview = m.querySelector('#dup-preview');
        const okBtn   = m.querySelector('#dup-ok');
        const okLbl   = m.querySelector('#dup-ok-lbl');

        inp.focus(); inp.select();

        function updatePreview() {
            const namePat = inp.value.trim();
            const addrPat = addrInp.value.trim();
            const count   = Math.max(1, Math.min(500, parseInt(cntInp.value) || 1));
            const startNum = parseInt(stInp.value) || 1;
            const hasHash  = /#+/.test(namePat);
            const hasAddrHash = /#+/.test(addrPat);

            if (count > 1 && !hasHash) {
                preview.innerHTML = '<span style="color:#e53935;"><i class="fa fa-exclamation-triangle"></i> Använd ## i namnmönstret för sekventiell namngivning (t.ex. STV2##_OP)</span>';
                okBtn.disabled = true;
                return;
            }
            okBtn.disabled = false;
            okLbl.textContent = count === 1 ? 'Duplicera' : `Duplicera ${count} st`;

            const copies = generateCopies(namePat, addrPat, count, startNum);
            if (!copies) return;
            const show = copies.slice(0, 4);
            let html = show.map(c => {
                const addrPart = hasAddrHash ? ` <span style="color:#888;">(${esc(c.addr)})</span>` : '';
                return `<div>↳ ${esc(c.name)}${addrPart}</div>`;
            }).join('');
            if (count > 4) html += `<div style="color:#999;">... och ${count - 4} till</div>`;
            preview.innerHTML = html;
        }

        inp.addEventListener('input', updatePreview);
        addrInp.addEventListener('input', updatePreview);
        cntInp.addEventListener('input', updatePreview);
        stInp.addEventListener('input', updatePreview);
        updatePreview();

        m.querySelector('#dup-cancel').addEventListener('click', () => m.remove());
        okBtn.addEventListener('click', async () => {
            const namePat  = inp.value.trim();
            const addrPat  = addrInp.value.trim();
            const count    = Math.max(1, Math.min(500, parseInt(cntInp.value) || 1));
            const startNum = parseInt(stInp.value) || 1;
            if (!namePat) { toastErr('Ange namnmönster'); return; }
            const copies = generateCopies(namePat, addrPat, count, startNum);
            if (!copies) return;
            if (copies.some(c => c.name === sourceName)) { toastErr('Nytt namn måste skilja sig från källan'); return; }
            m.remove();

            if (copies.length === 1) {
                try {
                    await duplicateTag(sourceName, copies[0].name, copies[0].addr || undefined);
                    toastOk('Duplicerad → ' + copies[0].name);
                    const sc = document.createElement('script');
                    sc.textContent = 'try{oTable.fnDraw();}catch(e){}';
                    document.head.appendChild(sc); sc.remove();
                } catch(e) {
                    toastErr('Duplicering misslyckades: ' + e.message);
                    console.error(CFG.logPrefix, 'Duplicate failed:', sourceName, e);
                }
            } else {
                toast(`Skapar ${copies.length} kopior...`);
                try {
                    const { ok, fail } = await duplicateTagBatch(sourceName, copies);
                    if (fail) toastErr(`${ok} skapade, ${fail} misslyckades`);
                    else toastOk(`${ok} taggar skapade`);
                    const sc = document.createElement('script');
                    sc.textContent = 'try{oTable.fnDraw();}catch(e){}';
                    document.head.appendChild(sc); sc.remove();
                } catch(e) {
                    toastErr('Batch-duplicering misslyckades: ' + e.message);
                    console.error(CFG.logPrefix, 'Batch duplicate failed:', e);
                }
            }
        });
        inp.addEventListener('keydown', e => {
            if (e.key === 'Enter') okBtn.click();
            if (e.key === 'Escape') m.remove();
        });
    }
    // ============================================================
    // KONFIG-RAPPORT
    // ============================================================
    async function collectReportData() {
        const sid = (location.search.match(/sid=([^&#]+)/) || [])[1] || '';
        const r = await fetch('/tag/GetTagList?sid=' + sid + '&draw=1&limit=9999&offset=0&sortcol=0&sortdir=asc&search=');
        if (!r.ok) throw new Error('HTTP ' + r.status);
        const json = await r.json();
        const dec = h => { const d = document.createElement('div'); d.innerHTML = h; return d.textContent.trim(); };
        return (json.data || []).map(row => {
            const dtype = dec(row['3'] || '');
            const emin  = dec(row['6'] || '0');
            const emax  = dec(row['7'] || '0');
            const isDigital = dtype === 'DIGITAL' || dtype === 'BOOL';
            const allZero   = ['4','5','6','7'].every(k => parseFloat(dec(row[k]||'0')) === 0);
            return {
                name:     dec(row['0'] || ''),
                io:       dec(row['1'] || ''),
                dtype,
                addr:     dec(row['2'] || ''),
                unit:     dec(row['8'] || ''),
                desc:     dec(row['10'] || ''),
                engmin:   parseFloat(emin) || 0,
                engmax:   parseFloat(emax) || 0,
                isDigital,
                isUnconf: !isDigital && allZero,
                isConf:   !isDigital && !allZero,
            };
        });
    }

    async function openKonfigRapport() {
        let tags;
        try {
            toastr.info('Hämtar taggar…', '', { timeOut: 0, extendedTimeOut: 0 });
            tags = await collectReportData();
            toastr.clear();
        } catch(e) {
            toastr.clear();
            toastErr('Kunde inte hämta taggar: ' + e.message);
            return;
        }
        if (!tags.length) { toastErr('Inga taggar hittade'); return; }

        const total   = tags.length;
        const conf    = tags.filter(t => t.isConf).length;
        const unconf  = tags.filter(t => t.isUnconf).length;
        const digital = tags.filter(t => t.isDigital).length;
        const noDesc  = tags.filter(t => !t.desc).length;

        const pct      = (n,d) => d ? Math.round(n/d*100) : 0;
        const esc      = escHtml;
        const countMap = arr => arr.reduce((m,v) => { m[v]=(m[v]||0)+1; return m; }, {});

        // ── Duplicate addresses ───────────────────────────────────
        const addrBuckets = {};
        tags.forEach(t => {
            if (!t.addr || t.addr==='0' || !t.io) return;
            const k = t.io+'|'+t.addr;
            (addrBuckets[k]||(addrBuckets[k]=[])).push(t.name);
        });
        const dupeTags = Object.values(addrBuckets).filter(a=>a.length>1).flat();

        // ── Address gap analysis ──────────────────────────────────
        const findGaps = addrs => {
            const parsed = addrs.map(a => {
                const m = a.match(/^([A-Za-z]*)(\d+)$/);
                return m ? {pfx:m[1], num:parseInt(m[2]), width:m[2].length} : null;
            }).filter(Boolean);
            const byPfx = {};
            parsed.forEach(p => (byPfx[p.pfx]||(byPfx[p.pfx]=[])).push(p));
            const result = [];
            for (const items of Object.values(byPfx)) {
                const nums = [...new Set(items.map(i=>i.num))].sort((a,b)=>a-b);
                if (nums.length < 3) continue;
                const range = nums[nums.length-1] - nums[0];
                if (range > 400 || range <= 0) continue;
                if (nums.length / (range+1) < 0.35) continue;
                const w = items[0].width, pfx = items[0].pfx;
                const fmt = n => pfx + String(n).padStart(w, '0');
                // Collect gap numbers then compress into ranges
                const gapNums = [];
                for (let i=nums[0]; i<=nums[nums.length-1]; i++) {
                    if (!nums.includes(i)) gapNums.push(i);
                }
                let i = 0;
                while (i < gapNums.length) {
                    let j = i;
                    while (j+1 < gapNums.length && gapNums[j+1] === gapNums[j]+1) j++;
                    result.push(j > i ? `${fmt(gapNums[i])}–${fmt(gapNums[j])}` : fmt(gapNums[i]));
                    i = j + 1;
                }
            }
            return result.slice(0, 50);
        };
        const addrByDev = {};
        tags.forEach(t => {
            if (!t.addr || t.addr==='0' || !t.io) return;
            (addrByDev[t.io]||(addrByDev[t.io]=[])).push(t.addr);
        });
        const addrGaps = Object.entries(addrByDev)
            .map(([dev,addrs]) => ({dev, gaps:findGaps(addrs)}))
            .filter(x=>x.gaps.length>0);

        // ── Scaling & unit deviants ───────────────────────────────
        // Group by sensorType + value suffix (GT_PV, GT_AD, GT_P separately)
        // so PV process values don't mix with AD/P/I control signals
        const stKey = name => {
            const p = name.split('_');
            if (p.length < 2) return null;
            const m = p[p.length-2].match(/^([A-Z]{1,4})\d/i);
            if (!m) return null;
            return m[1].toUpperCase() + '_' + p[p.length-1];
        };
        const stMap = {};
        tags.filter(t=>t.isConf).forEach(t => {
            const k = stKey(t.name);
            if (k) (stMap[k]||(stMap[k]=[])).push(t);
        });
        const deviants = [];
        for (const [type, group] of Object.entries(stMap)) {
            if (group.length < 3) continue;
            const [[topUnit]]    = Object.entries(countMap(group.map(t=>t.unit||''))).sort((a,b)=>b[1]-a[1]);
            const scalingKey     = t => `${Math.round(t.engmin*10)/10}|${Math.round(t.engmax*10)/10}`;
            const [[topScaling]] = Object.entries(countMap(group.map(scalingKey))).sort((a,b)=>b[1]-a[1]);
            const [topEmin, topEmax] = topScaling.split('|').map(Number);
            const span = Math.abs(topEmax - topEmin) || 1;
            const odd  = group.filter(t => {
                const badUnit    = (t.unit||'') !== topUnit;
                const badScaling = Math.abs(t.engmin-topEmin)/span > 0.15 || Math.abs(t.engmax-topEmax)/span > 0.15;
                return badUnit || badScaling;
            });
            if (odd.length > 0 && odd.length < group.length/2) {
                deviants.push({ type, total:group.length, consensus:`${topEmin}…${topEmax} ${topUnit}`, odd });
            }
        }
        deviants.sort((a,b)=>b.odd.length-a.odd.length);

        // ── Render helpers ────────────────────────────────────────
        const togSection = (title, badge, bCls, body) => `<div class="kr-hyg-sec">
  <div class="kr-hyg-hdr" onclick="var b=this.nextElementSibling,o=b.classList.toggle('open');this.querySelector('.kr-hyg-ico').style.transform=o?'rotate(90deg)':''">
    <i class="fa fa-chevron-right kr-hyg-ico"></i>
    <span class="kr-hyg-ttl">${title}</span>
    <span class="kr-hyg-badge ${bCls}">${badge}</span>
  </div>
  <div class="kr-hyg-body">${body}</div>
</div>`;
        const hygSection = (title, list, level) => {
            const empty = !list.length;
            const body = empty
                ? '<span style="color:#4caf50;font-size:11px;"><i class="fa fa-check"></i> Inga avvikelser</span>'
                : list.map(n=>`<span class="kr-tag">${esc(n)}</span>`).join('');
            return togSection(esc(title), empty?'✓':list.length, empty?'ok':level, body);
        };

        // ── Address gap sections ──────────────────────────────────
        const gapHtml = addrGaps.length === 0
            ? '<span style="color:#4caf50;font-size:11px;"><i class="fa fa-check"></i> Inga luckor hittade</span>'
            : addrGaps.map(({dev,gaps}) => hygSection(`${dev} — ${gaps.length} adressluckor`, gaps, 'warn')).join('');

        // ── Deviant sections ──────────────────────────────────────
        const devHtml = deviants.length === 0
            ? '<span style="color:#4caf50;font-size:11px;"><i class="fa fa-check"></i> Inga avvikelser hittade</span>'
            : deviants.map(d => {
                const items = d.odd.map(t =>
                    `<div style="padding:2px 0;"><span class="kr-tag">${esc(t.name)}</span> <span style="color:#888;font-size:11px;">→ ${esc(`${t.engmin}…${t.engmax} ${t.unit||'?'}`)}</span> <span style="color:#aaa;font-size:10px;">(förväntat: ${esc(d.consensus)})</span></div>`
                ).join('');
                const title = `${esc(d.type)}-taggar <span style="color:#aaa;font-weight:400;font-size:11px;">(${d.total} st, konsensus: ${esc(d.consensus)})</span>`;
                return togSection(title, `${d.odd.length} avvikare`, 'warn', items);
            }).join('');

        const ov = document.createElement('div');
        ov.className = 'kr-ov';
        ov.innerHTML = `
<div class="kr-hdr">
  <i class="fa fa-bar-chart" style="font-size:17px;color:#5b6abf;"></i>
  <h2>Konfig-Rapport</h2>
  <span class="kr-sub">${total} taggar inlästa</span>
  <button class="kr-close" id="kr-close" title="Stäng (Esc)"><i class="fa fa-times"></i></button>
</div>
<div class="kr-body">

  <div class="kr-stats">
    <div class="kr-stat"><div class="kn info">${total}</div><div class="kl">Taggar totalt</div></div>
    <div class="kr-stat"><div class="kn">${conf}</div><div class="kl">Konfigurerade</div><div class="ks">${pct(conf,total)}%</div></div>
    <div class="kr-stat"><div class="kn warn">${unconf}</div><div class="kl">Okonfigurerade</div><div class="ks">${pct(unconf,total)}%</div></div>
    <div class="kr-stat"><div class="kn gray">${digital}</div><div class="kl">Digitala</div><div class="ks">${pct(digital,total)}%</div></div>
    <div class="kr-stat"><div class="kn warn">${noDesc}</div><div class="kl">Saknar beskrivning</div><div class="ks">${pct(noDesc,total)}%</div></div>
  </div>

  <div class="kr-hyg" style="margin-bottom:14px;">
    <h4><i class="fa fa-map-marker"></i> Adressluckor <span style="font-weight:400;color:#aaa;font-size:10px;">— saknade nummer i sekventiella adressrader per IO-enhet</span></h4>
    ${gapHtml}
  </div>

  <div class="kr-hyg" style="margin-bottom:14px;">
    <h4><i class="fa fa-exclamation-triangle"></i> Skalnings- &amp; enhetsavvikelser <span style="font-weight:400;color:#aaa;font-size:10px;">— taggar vars enhet/skalning skiljer sig från konsensus för samma taggtyp + suffix</span></h4>
    ${devHtml}
  </div>

  <div class="kr-hyg">
    <h4><i class="fa fa-stethoscope"></i> Hygienrapport</h4>
    ${hygSection('Saknar beskrivning', tags.filter(t=>!t.desc).map(t=>t.name), 'warn')}
    ${hygSection('Duplicerade adresser', dupeTags, 'warn')}
  </div>

</div>`;

        document.body.appendChild(ov);
        ov.querySelector('#kr-close').addEventListener('click', () => ov.remove());
        const escHandler = e => { if (e.key==='Escape') { ov.remove(); document.removeEventListener('keydown', escHandler); } };
        document.addEventListener('keydown', escHandler);
    }

    function openColFilterDropdown(colIdx, anchor) {
        document.getElementById('inu-col-filter-dd')?.remove();
        const vals = new Set();
        document.querySelectorAll('#tagtable tbody tr.tag').forEach(r => {
            const v = r.cells[colIdx]?.textContent?.trim();
            if (v) vals.add(v);
        });
        const active = filterColVals[colIdx];
        const allActive = !active || active.size === 0;
        const dd = document.createElement('div');
        dd.id = 'inu-col-filter-dd';
        dd.className = 'inu-col-filter-dd';
        const sortedVals = [...vals].sort();
        dd.innerHTML =
            '<div class="inu-cfd-search"><input type="text" placeholder="Sök värde..." id="inu-cfd-sinp"></div>' +
            '<div class="inu-cfd-list">' +
              '<label class="inu-cfd-item"><input type="checkbox" value="__all__"' + (allActive?' checked':'') + '> <span>(Alla)</span></label>' +
              sortedVals.map(v => '<label class="inu-cfd-item"><input type="checkbox" value="' + escHtml(v) + '"' + (allActive || active?.has(v) ? ' checked' : '') + '> <span>' + escHtml(v) + '</span></label>').join('') +
            '</div>' +
            '<div class="inu-cfd-footer"><button class="inu-cfd-cancel">Avbryt</button><button class="inu-cfd-ok">OK</button></div>';
        const rect = anchor.getBoundingClientRect();
        dd.style.top = (rect.bottom + 4) + 'px';
        dd.style.left = Math.min(rect.left, window.innerWidth - 270) + 'px';
        document.body.appendChild(dd);
        dd.querySelector('#inu-cfd-sinp').addEventListener('input', e => {
            const q = e.target.value.toLowerCase();
            dd.querySelectorAll('.inu-cfd-list .inu-cfd-item').forEach((item, i) => {
                if (i === 0) return; // skip "(Alla)"
                item.style.display = item.querySelector('span').textContent.toLowerCase().includes(q) ? '' : 'none';
            });
        });
        const allCb = dd.querySelector('[value="__all__"]');
        allCb.addEventListener('change', () => {
            dd.querySelectorAll('.inu-cfd-list input:not([value="__all__"])').forEach(cb => { cb.checked = allCb.checked; });
        });
        dd.querySelectorAll('.inu-cfd-list input:not([value="__all__"])').forEach(cb => {
            cb.addEventListener('change', () => {
                const allChecked = [...dd.querySelectorAll('.inu-cfd-list input:not([value="__all__"])')].every(c => c.checked);
                allCb.checked = allChecked;
            });
        });
        dd.querySelector('.inu-cfd-ok').addEventListener('click', () => {
            const itemCbs = [...dd.querySelectorAll('.inu-cfd-list input:not([value="__all__"])')];
            const checked = itemCbs.filter(cb => cb.checked).map(cb => cb.value);
            if (checked.length === 0 || checked.length === itemCbs.length) {
                delete filterColVals[colIdx];
            } else {
                filterColVals[colIdx] = new Set(checked);
            }
            updateColFilterIndicators();
            applyFilter();
            updSummary();
            dd.remove();
        });
        dd.querySelector('.inu-cfd-cancel').addEventListener('click', () => dd.remove());
        setTimeout(() => {
            const close = e => { if (!dd.contains(e.target) && e.target !== anchor) { dd.remove(); document.removeEventListener('mousedown', close); } };
            document.addEventListener('mousedown', close);
        }, 0);
    }
    function updateColFilterIndicators() {
        document.querySelectorAll('.inu-col-filter-btn').forEach(btn => {
            const active = filterColVals[+btn.dataset.col]?.size > 0;
            btn.classList.toggle('active', !!active);
        });
    }
    function addColumnFilters() {
        const ths = document.querySelectorAll('#tagtable thead tr:first-child th');
        [{ idx: C.IO }, { idx: C.DTYPE }, { idx: C.UNIT }].forEach(({ idx }) => {
            const th = ths[idx];
            if (!th || th.querySelector('.inu-col-filter-btn')) return;
            th.style.overflow = 'visible';
            const btn = document.createElement('i');
            btn.className = 'fa fa-filter inu-col-filter-btn';
            btn.dataset.col = idx;
            btn.title = 'Filtrera kolumn';
            th.appendChild(btn);
            btn.addEventListener('click', e => { e.stopPropagation(); openColFilterDropdown(idx, btn); });
        });
    }

    // ============================================================
    // MODALS
    // ============================================================
    function modal(html) {
        const o=document.createElement('div'); o.className='mo';
        o.innerHTML=`<div class="mb">${html}</div>`;
        document.body.appendChild(o);
        o.addEventListener('click',e=>{if(e.target===o)o.remove();});
        return o;
    }

    function modalAdd(pre) {
        const p=pre||{rawmin:'0',rawmax:'10000',engmin:'0',engmax:'100',unit:'',format:'0.0'};
        const m=modal(`
<h3><i class="fa fa-plus"></i> Ny preset</h3>
<label>Namn</label><input id="pn" placeholder="T.ex. Min sensor (0…500)">
<div class="fr"><div><label>Rå-min</label><input id="p1" value="${p.rawmin}"></div><div><label>Rå-max</label><input id="p2" value="${p.rawmax}"></div></div>
<div class="fr"><div><label>Vy-min</label><input id="p3" value="${p.engmin}"></div><div><label>Vy-max</label><input id="p4" value="${p.engmax}"></div></div>
<div class="fr"><div><label>Enhet</label><input id="p5" value="${p.unit}"></div><div><label>Format</label><input id="p6" value="${p.format}"></div></div>
<div class="bt"><button class="bx" id="pc">Avbryt</button><button class="bok" id="ps">Spara</button></div>`);
        m.querySelector('#pc').addEventListener('click',()=>m.remove());
        m.querySelector('#ps').addEventListener('click',()=>{
            const n=m.querySelector('#pn').value.trim();
            if(!n){toastErr('Ange namn');return;}
            const c=loadCustom();
            c[n]={rawmin:m.querySelector('#p1').value.trim(),rawmax:m.querySelector('#p2').value.trim(),
                  engmin:m.querySelector('#p3').value.trim(),engmax:m.querySelector('#p4').value.trim(),
                  unit:m.querySelector('#p5').value.trim(),format:m.querySelector('#p6').value.trim()};
            saveCustom(c); m.remove(); refreshBulkOpts(); toastOk(`"${n}" sparad`);
        });
    }

    // ============================================================
    // TEMPLATE LIBRARY — Add from Template
    // ============================================================
    // Inline index of available device templates. Per-device JSON files are
    // lazy-fetched from TEMPLATE_BASE_URL and cached in GM storage.
    const TEMPLATE_BASE_URL = 'https://phogel1.github.io/static-assets/';
    const TEMPLATE_INDEX = {
        version: '2026-04-15.6',
        manufacturers: [
            {
                id: 'ivprodukt', name: 'IVProdukt',
                models: [
                    { id: 'climatix-ahu', name: 'Climatix AHU (v4.34)', file: 'ivprodukt/climatix-ahu.json', category: 'AHU' }
                ]
            },
            {
                id: 'nibe', name: 'NIBE',
                models: [
                    { id: 'smo-s40', name: 'SMO S40 (orkestrator)', file: 'nibe/smo-s40.json', category: 'Värmepump' }
                ]
            },
            {
                id: 'thermia', name: 'Thermia',
                models: [
                    { id: 'mega-genesis', name: 'Mega / Genesis platform (v12.00)', file: 'thermia/mega-genesis.json', category: 'Värmepump' }
                ]
            }
        ]
    };

    function _tplCacheKey(file) { return 'inu_tpl_' + file; }
    function _tplCacheGet(file, idxVersion) {
        try {
            const raw = GM_getValue(_tplCacheKey(file), null);
            if (!raw) return null;
            const obj = JSON.parse(raw);
            if (obj._idx !== idxVersion) return null; // invalidate when index bumps
            return obj.data;
        } catch (e) { return null; }
    }
    function _tplCachePut(file, idxVersion, data) {
        try { GM_setValue(_tplCacheKey(file), JSON.stringify({ _idx: idxVersion, data })); }
        catch (e) { console.warn(CFG.logPrefix, 'template cache write failed', e); }
    }

    // Fetch a template file via GM_xmlhttpRequest (bypasses CORS) with cache.
    // Returns a Promise<templateObject>. Throws on network/parse errors.
    function fetchTemplate(file) {
        const cached = _tplCacheGet(file, TEMPLATE_INDEX.version);
        if (cached) return Promise.resolve(cached);
        return new Promise((resolve, reject) => {
            if (typeof GM_xmlhttpRequest !== 'function') {
                reject(new Error('GM_xmlhttpRequest saknas — uppdatera Tampermonkey-stub med @grant GM_xmlhttpRequest och @connect phogel1.github.io'));
                return;
            }
            GM_xmlhttpRequest({
                method: 'GET',
                url: TEMPLATE_BASE_URL + file + '?v=' + encodeURIComponent(TEMPLATE_INDEX.version),
                timeout: 15000,
                onload: (res) => {
                    if (res.status < 200 || res.status >= 300) {
                        reject(new Error('HTTP ' + res.status + ' för ' + file));
                        return;
                    }
                    try {
                        const data = JSON.parse(res.responseText);
                        _tplCachePut(file, TEMPLATE_INDEX.version, data);
                        resolve(data);
                    } catch (e) {
                        reject(new Error('Ogiltig JSON: ' + e.message));
                    }
                },
                onerror: () => reject(new Error('Nätverksfel — ' + file)),
                ontimeout: () => reject(new Error('Timeout — ' + file))
            });
        });
    }

    // Pure resolver: given a template + user's config answers + rename prefix,
    // returns an ordered array of fully-resolved tag objects ready for insert.
    // `answers` shape: { <sectionId>: <optionId string for radio> | <optionId[] for multiselect> }
    function resolveTemplate(tpl, answers, prefix) {
        if (!tpl || !tpl.tags) return [];
        const chosen = new Set(Array.isArray(tpl.base) ? tpl.base : []);
        for (const sec of (tpl.config || [])) {
            const ans = answers[sec.id];
            if (ans == null) continue;
            const picks = Array.isArray(ans) ? ans : [ans];
            for (const opt of (sec.options || [])) {
                if (picks.indexOf(opt.id) === -1) continue;
                for (const tid of (opt.tags || [])) chosen.add(tid);
            }
        }
        const out = [];
        const effPrefix = (prefix && prefix.trim()) || tpl.defaultDevicePrefix || 'DEV1';
        for (const tid of chosen) {
            const raw = tpl.tags[tid];
            if (!raw) { console.warn(CFG.logPrefix, 'resolveTemplate: unknown tag id', tid); continue; }
            const resolved = {};
            for (const k of Object.keys(raw)) {
                const v = raw[k];
                resolved[k] = (typeof v === 'string') ? v.replace(/\{prefix\}/g, effPrefix) : v;
            }
            resolved._id = tid;
            out.push(resolved);
        }
        return out;
    }

    // ----- Tag creation via copy+edit seed-tag approach -----
    // Since WebPort's native "new tag" endpoint is unknown, we reuse the existing
    // copy path: clone a user-picked seed tag to the target name, then immediately
    // POST over its fields via /tag/actionedit. Two requests per created tag.

    async function createTagFromTemplate(tag, seedTagName) {
        // Step 1: copy seed → new name
        const copyParams = await fetchCopyBaseParams(seedTagName);
        copyParams.set('name', tag.name);
        if (tag.address != null) copyParams.set('address', tag.address);
        const copyRes = await fetch('/tag/actioncopy', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: copyParams.toString()
        });
        if (!copyRes.ok) throw new Error('copy HTTP ' + copyRes.status);

        // Step 2: overwrite fields on the new tag with template data
        await fetchFormAndSave(tag.name, (fd) => {
            const fields = ['device', 'address', 'datatype', 'rawmin', 'rawmax',
                            'engmin', 'engmax', 'unit', 'format', 'description',
                            'alarmoptions', 'trendoptions'];
            for (const k of fields) {
                if (tag[k] != null && tag[k] !== '') fd.set(k, String(tag[k]));
            }
        });
    }

    // Batch-create with sequential dispatch + progress callback.
    // progressCb(done, total, failing[]) fires after each tag.
    async function createTagBatch(tags, seedTagName, progressCb) {
        const failing = [];
        for (let i = 0; i < tags.length; i++) {
            const t = tags[i];
            try {
                await createTagFromTemplate(t, seedTagName);
                logAppend('success', `Skapade tagg: ${t.name}`, 'inu');
            } catch (e) {
                console.warn(CFG.logPrefix, 'create failed', t.name, e);
                logAppend('error', `Misslyckades: ${t.name} — ${e.message}`, 'inu');
                failing.push({ name: t.name, error: e.message });
            }
            if (progressCb) progressCb(i + 1, tags.length, failing);
            // Small delay to avoid hammering the server
            await new Promise(r => setTimeout(r, 80));
        }
        return failing;
    }

    // ----- UI -----
    function _tplEsc(s) {
        return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
    }

    function _tplExistingTags() {
        const out = [];
        document.querySelectorAll('#tagtable tbody tr.tag').forEach(r => {
            const name = r.cells[C.NAME]?.textContent?.trim();
            if (name) out.push(name);
        });
        return out;
    }

    // Collect existing addresses for duplicate detection. Modbus uniqueness is
    // per (device, address) — the same address on different devices is fine.
    function _tplExistingAddressesByDevice() {
        const map = new Map(); // key: `${device}|${address}` → tag name
        document.querySelectorAll('#tagtable tbody tr.tag').forEach(r => {
            const name = r.cells[C.NAME]?.textContent?.trim();
            const device = r.cells[C.IO]?.textContent?.trim();
            const addr = r.cells[C.ADDR]?.textContent?.trim();
            if (name && addr) map.set(`${device || ''}|${addr}`, name);
        });
        return map;
    }

    // Types treated as "pick exactly one" (renders as <select>):
    //   radio, dropdown
    // Types treated as "pick zero or more" (renders as checkbox list):
    //   multiselect, (and anything else)
    function _tplIsSingleSelect(sec) { return sec.type === 'radio' || sec.type === 'dropdown'; }

    // Compact "type/unit" label for the info popup.
    // Prefer the engineering unit when present (°C, %, Pa, kWh, …).
    // Otherwise infer from datatype + name suffix:
    //   DIGITAL with _AL/_FAULT/_HAL/_LAL → "Larm"
    //   DIGITAL otherwise                  → "Digital"
    //   INT/UINT/LONG/ULONG/STRING without unit → datatype
    function _tplTagTypeLabel(tag) {
        if (tag.unit && String(tag.unit).trim()) return String(tag.unit).trim();
        if (tag.datatype === 'DIGITAL') {
            const n = String(tag.name || '').toUpperCase();
            if (/_(AL\d*|FAULT|HAL|LAL)$/.test(n)) return 'Larm';
            return 'Digital';
        }
        return tag.datatype || '';
    }

    // Resolve the tags that a single configurator option will pull in, with
    // {prefix} substituted. Returns an array of {name, address, type, description}.
    function _tplOptionTags(tpl, opt, prefix) {
        if (!opt || !opt.tags || !opt.tags.length) return [];
        const effPrefix = (prefix && String(prefix).trim()) || tpl.defaultDevicePrefix || 'DEV1';
        const out = [];
        for (const tid of opt.tags) {
            const raw = tpl.tags && tpl.tags[tid];
            if (!raw) continue;
            const resolvedName = String(raw.name || tid).replace(/\{prefix\}/g, effPrefix);
            out.push({
                name: resolvedName,
                address: raw.address || '',
                type: _tplTagTypeLabel({ name: resolvedName, unit: raw.unit, datatype: raw.datatype }),
                description: raw.description || ''
            });
        }
        return out;
    }

    // Build a small "i" button that shows a popup of the tags returned by
    // getTags() (called fresh each time the popup opens, so the displayed
    // names reflect the current rename-prefix). Hover or click to toggle.
    function _tplCreateInfoBtn(getTags) {
        const btn = document.createElement('span');
        btn.className = 'tpl-info-btn';
        btn.textContent = 'i';
        btn.title = 'Visa taggar som denna option lägger till';
        let popup = null;
        function hide() {
            if (popup) { popup.remove(); popup = null; }
            btn.classList.remove('open');
        }
        function show() {
            hide();
            const tags = getTags() || [];
            popup = document.createElement('div');
            popup.className = 'tpl-info-popup';
            let html;
            if (!tags.length) {
                html = '<div class="tpl-info-hdr">Inga taggar</div>'
                     + '<div class="tpl-info-empty">Det valda alternativet lägger inte till några extra taggar.</div>';
            } else {
                html = '<div class="tpl-info-hdr">' + tags.length + ' tagg' + (tags.length === 1 ? '' : 'ar') + '</div>';
                html += '<table class="tpl-info-tbl"><thead><tr><th>Namn</th><th>Adress</th><th>Typ / Enhet</th><th>Beskrivning</th></tr></thead><tbody>';
                for (const t of tags) {
                    html += '<tr><td><code>' + _tplEsc(t.name) + '</code></td><td><code>' + _tplEsc(t.address) + '</code></td><td><span class="tpl-info-type">' + _tplEsc(t.type) + '</span></td><td>' + _tplEsc(t.description) + '</td></tr>';
                }
                html += '</tbody></table>';
            }
            popup.innerHTML = html;
            document.body.appendChild(popup);
            // Position below the button, kept inside the viewport.
            const r = btn.getBoundingClientRect();
            const pw = popup.offsetWidth;
            const ph = popup.offsetHeight;
            const vw = document.documentElement.clientWidth;
            const vh = document.documentElement.clientHeight;
            let left = r.left + window.scrollX;
            let top = r.bottom + window.scrollY + 6;
            if (left + pw > vw + window.scrollX - 8) left = Math.max(8 + window.scrollX, vw + window.scrollX - pw - 8);
            if (r.bottom + ph + 6 > vh) top = r.top + window.scrollY - ph - 6; // flip above if no room below
            popup.style.left = left + 'px';
            popup.style.top = top + 'px';
            btn.classList.add('open');
        }
        btn.addEventListener('mouseenter', show);
        btn.addEventListener('mouseleave', hide);
        // Block clicks from propagating into the parent <label> (which would
        // toggle the checkbox or focus the select).
        btn.addEventListener('click', e => { e.preventDefault(); e.stopPropagation(); });
        btn.addEventListener('mousedown', e => { e.preventDefault(); e.stopPropagation(); });
        return btn;
    }

    function _tplRenderConfig(tpl, state, container, onChange, getPrefix) {
        container.innerHTML = '';
        if (!tpl.config || !tpl.config.length) {
            const msg = document.createElement('div');
            msg.style.cssText = 'font-size:12px;color:#666;padding:8px;background:#f8f8f8;border-radius:4px;';
            msg.textContent = 'Den här mallen har inga konfigurerbara sektioner — alla taggar i base läggs till.';
            container.appendChild(msg);
            return;
        }

        // Group single-select sections into a compact grid at the top; stack
        // multiselects vertically below. This keeps the "structural" questions
        // tight (one row of dropdowns) and gives multiselects the full width
        // they need for readable checkbox lists.
        const singles = tpl.config.filter(_tplIsSingleSelect);
        const multis = tpl.config.filter(s => !_tplIsSingleSelect(s));

        if (singles.length) {
            const grid = document.createElement('div');
            grid.className = 'tpl-dropdown-grid';
            for (const sec of singles) {
                const cell = document.createElement('div');
                cell.className = 'tpl-dd-cell';
                const lbl = document.createElement('label');
                lbl.className = 'tpl-dd-lbl';
                lbl.textContent = sec.label || sec.id;
                cell.appendChild(lbl);
                const row = document.createElement('div');
                row.className = 'tpl-dd-row';
                const sel = document.createElement('select');
                sel.className = 'tpl-dd';
                for (const opt of (sec.options || [])) {
                    const o = document.createElement('option');
                    o.value = opt.id;
                    o.textContent = opt.label || opt.id;
                    if (state[sec.id] === opt.id) o.selected = true;
                    sel.appendChild(o);
                }
                sel.addEventListener('change', () => {
                    state[sec.id] = sel.value;
                    onChange();
                });
                row.appendChild(sel);
                // Info button shows tags from the currently-selected option;
                // recomputed each time the popup opens, so it always reflects
                // the current dropdown value and prefix.
                const infoBtn = _tplCreateInfoBtn(() => {
                    const currentId = state[sec.id];
                    const currentOpt = (sec.options || []).find(o => o.id === currentId);
                    return _tplOptionTags(tpl, currentOpt, getPrefix && getPrefix());
                });
                row.appendChild(infoBtn);
                cell.appendChild(row);
                if (sec.help) {
                    const h = document.createElement('div');
                    h.className = 'tpl-dd-help';
                    h.textContent = sec.help;
                    cell.appendChild(h);
                }
                grid.appendChild(cell);
            }
            container.appendChild(grid);
        }

        for (const sec of multis) {
            const sEl = document.createElement('div');
            sEl.className = 'tpl-sec';
            const lbl = document.createElement('div');
            lbl.className = 'tpl-sec-lbl';
            lbl.textContent = sec.label || sec.id;
            sEl.appendChild(lbl);
            if (sec.help) {
                const h = document.createElement('div');
                h.className = 'tpl-sec-help';
                h.textContent = sec.help;
                sEl.appendChild(h);
            }
            const opts = document.createElement('div');
            opts.className = 'tpl-opts';
            for (const opt of (sec.options || [])) {
                const lab = document.createElement('label');
                lab.className = 'tpl-opt';
                const inp = document.createElement('input');
                inp.type = 'checkbox';
                inp.name = 'tpl_' + sec.id;
                inp.value = opt.id;
                const curr = state[sec.id];
                inp.checked = Array.isArray(curr) && curr.indexOf(opt.id) !== -1;
                if (inp.checked) lab.classList.add('tpl-chk-on');
                inp.addEventListener('change', () => {
                    if (!Array.isArray(state[sec.id])) state[sec.id] = [];
                    const arr = state[sec.id];
                    const idx = arr.indexOf(opt.id);
                    if (inp.checked) { if (idx === -1) arr.push(opt.id); lab.classList.add('tpl-chk-on'); }
                    else { if (idx !== -1) arr.splice(idx, 1); lab.classList.remove('tpl-chk-on'); }
                    onChange();
                });
                lab.appendChild(inp);
                const txt = document.createElement('span');
                txt.textContent = opt.label || opt.id;
                lab.appendChild(txt);
                const infoBtn = _tplCreateInfoBtn(() => _tplOptionTags(tpl, opt, getPrefix && getPrefix()));
                lab.appendChild(infoBtn);
                opts.appendChild(lab);
            }
            sEl.appendChild(opts);
            container.appendChild(sEl);
        }
    }

    function _tplInitialAnswers(tpl) {
        const a = {};
        for (const sec of (tpl.config || [])) {
            if (_tplIsSingleSelect(sec)) {
                a[sec.id] = sec.default || (sec.options && sec.options[0] && sec.options[0].id);
            } else {
                a[sec.id] = (sec.options || []).filter(o => o.default).map(o => o.id);
            }
        }
        return a;
    }

    function _tplRenderResolvedList(tags, unchecked, container, onToggle) {
        container.innerHTML = '';
        const search = document.createElement('div');
        search.className = 'tpl-res-search';
        const sInp = document.createElement('input');
        sInp.type = 'text';
        sInp.placeholder = 'Filtrera namn eller beskrivning...';
        search.appendChild(sInp);
        container.appendChild(search);
        const body = document.createElement('div');
        container.appendChild(body);

        function draw(filter) {
            body.innerHTML = '';
            const q = (filter || '').toLowerCase();
            for (const t of tags) {
                if (q) {
                    const hay = ((t.name || '') + ' ' + (t.description || '')).toLowerCase();
                    if (hay.indexOf(q) === -1) continue;
                }
                const row = document.createElement('label');
                row.className = 'tpl-res-row';
                if (unchecked.has(t._id)) row.classList.add('tpl-r-off');
                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.checked = !unchecked.has(t._id);
                cb.addEventListener('change', () => {
                    if (cb.checked) unchecked.delete(t._id);
                    else unchecked.add(t._id);
                    row.classList.toggle('tpl-r-off', !cb.checked);
                    if (onToggle) onToggle();
                });
                row.appendChild(cb);
                const n = document.createElement('span'); n.className = 'tpl-res-name'; n.textContent = t.name; row.appendChild(n);
                const a = document.createElement('span'); a.className = 'tpl-res-addr'; a.textContent = t.address || ''; row.appendChild(a);
                const u = document.createElement('span'); u.className = 'tpl-res-unit'; u.textContent = t.unit || ''; row.appendChild(u);
                const d = document.createElement('span'); d.className = 'tpl-res-desc'; d.textContent = t.description || ''; row.appendChild(d);
                body.appendChild(row);
            }
        }
        sInp.addEventListener('input', () => draw(sInp.value));
        draw('');
    }

    async function openTemplateModal() {
        const m = modal('');
        m.classList.add('tpl-modal');
        const mb = m.querySelector('.mb');
        mb.innerHTML = `
<div class="tpl-hdr">
  <h3><i class="fa fa-cubes"></i> Lägg till från mall</h3>
  <button class="tpl-close" title="Stäng">✕</button>
</div>
<div class="tpl-pickers">
  <div>
    <label>Tillverkare</label>
    <select id="tpl-mfr"></select>
  </div>
  <div>
    <label>Modell</label>
    <select id="tpl-model"></select>
  </div>
  <div>
    <label>Enhets-prefix</label>
    <input id="tpl-prefix" placeholder="AHU1 / HP1 / …">
  </div>
</div>
<div class="tpl-status" id="tpl-status"></div>
<div class="tpl-cfg" id="tpl-cfg"><div style="font-size:12px;color:#888;padding:12px;">Välj tillverkare och modell för att ladda mallen...</div></div>
<div class="tpl-resolved" style="display:none;" id="tpl-res-wrap">
  <div class="tpl-res-hdr" id="tpl-res-hdr">
    <span class="tpl-res-tri">▶</span>
    <span>Taggar som kommer att skapas:</span>
    <span class="tpl-res-count" id="tpl-res-count">0</span>
    <span style="font-size:10px;color:#888;font-weight:400;">(klicka för att förhandsgranska / avmarkera enskilda)</span>
  </div>
  <div class="tpl-res-body" id="tpl-res-body"></div>
</div>
<div class="tpl-foot">
  <div class="tpl-prog" id="tpl-prog">
    <div id="tpl-prog-text">Väntar...</div>
    <div class="tpl-prog-bar"><div class="tpl-prog-fill" id="tpl-prog-fill"></div></div>
  </div>
  <button class="tpl-cancel" id="tpl-cancel">Avbryt</button>
  <button class="tpl-add" id="tpl-add" disabled>Lägg till</button>
</div>`;

        const q = s => mb.querySelector(s);
        const mfrSel = q('#tpl-mfr'), modelSel = q('#tpl-model');
        const prefixInp = q('#tpl-prefix'), statusEl = q('#tpl-status'), cfgEl = q('#tpl-cfg');
        const resWrap = q('#tpl-res-wrap'), resHdr = q('#tpl-res-hdr'), resBody = q('#tpl-res-body'), resCount = q('#tpl-res-count');
        const addBtn = q('#tpl-add'), cancelBtn = q('#tpl-cancel');
        const progEl = q('#tpl-prog'), progText = q('#tpl-prog-text'), progFill = q('#tpl-prog-fill');

        // Existing tags in the current WebPort table. The first one is auto-used as
        // the seed tag for the copy+edit creation path. Empty table → blocking error.
        const existing = _tplExistingTags();
        const seedTag = existing[0] || null;

        // Populate manufacturer dropdown
        mfrSel.innerHTML = '';
        for (const mfr of TEMPLATE_INDEX.manufacturers) {
            const o = document.createElement('option');
            o.value = mfr.id; o.textContent = mfr.name;
            mfrSel.appendChild(o);
        }

        let currentTpl = null;
        let currentAnswers = {};
        const uncheckedIds = new Set();

        function populateModelDropdown() {
            const mfrId = mfrSel.value;
            const mfr = TEMPLATE_INDEX.manufacturers.find(m => m.id === mfrId);
            modelSel.innerHTML = '';
            if (!mfr) return;
            for (const md of mfr.models) {
                const o = document.createElement('option');
                o.value = md.id; o.textContent = md.name;
                o.dataset.file = md.file;
                modelSel.appendChild(o);
            }
        }

        function setStatus(msg, isErr) {
            if (!msg) { statusEl.style.display = 'none'; return; }
            statusEl.textContent = msg;
            statusEl.classList.toggle('tpl-err', !!isErr);
            statusEl.style.display = 'block';
        }

        function updatePreview() {
            if (!currentTpl) { resWrap.style.display = 'none'; addBtn.disabled = true; return; }
            const resolved = resolveTemplate(currentTpl, currentAnswers, prefixInp.value);
            const effective = resolved.filter(t => !uncheckedIds.has(t._id));
            resCount.textContent = effective.length + ' taggar';
            resWrap.style.display = 'block';
            _tplRenderResolvedList(resolved, uncheckedIds, resBody, updatePreview);
            addBtn.textContent = effective.length ? `Lägg till ${effective.length} taggar` : 'Lägg till';
            addBtn.disabled = effective.length === 0 || !seedTag;
        }

        // If there are no existing tags in the table, surface a blocking warning
        // up front. The seed-tag mechanism needs at least one existing tag to copy
        // from, so creation is impossible until the commissioner adds one.
        if (!seedTag) {
            setStatus('Tabellen är tom — skapa minst en tagg manuellt först. Mall-funktionen kopierar fält från en befintlig tagg, så listan får inte vara tom.', true);
        }

        async function loadModel() {
            const opt = modelSel.options[modelSel.selectedIndex];
            if (!opt || !opt.dataset.file) return;
            if (seedTag) setStatus('Laddar mall...', false);
            cfgEl.innerHTML = '<div style="font-size:12px;color:#888;padding:12px;">Hämtar...</div>';
            try {
                const tpl = await fetchTemplate(opt.dataset.file);
                currentTpl = tpl;
                currentAnswers = _tplInitialAnswers(tpl);
                uncheckedIds.clear();
                if (tpl.defaultDevicePrefix && !prefixInp.value) prefixInp.value = tpl.defaultDevicePrefix;
                _tplRenderConfig(tpl, currentAnswers, cfgEl, updatePreview, () => prefixInp.value);
                // Don't wipe the empty-table warning when the load succeeds — it stays
                // up until the commissioner closes the modal and creates a seed tag.
                if (seedTag) setStatus('', false);
                updatePreview();
            } catch (e) {
                currentTpl = null;
                cfgEl.innerHTML = '';
                setStatus('Kunde inte ladda mall: ' + e.message, true);
                updatePreview();
            }
        }

        mfrSel.addEventListener('change', () => { populateModelDropdown(); loadModel(); });
        modelSel.addEventListener('change', loadModel);
        prefixInp.addEventListener('input', updatePreview);
        resHdr.addEventListener('click', () => {
            resHdr.classList.toggle('open');
            resBody.classList.toggle('open');
        });
        q('.tpl-close').addEventListener('click', () => m.remove());
        cancelBtn.addEventListener('click', () => m.remove());

        addBtn.addEventListener('click', async () => {
            if (!currentTpl || !seedTag) return;
            const resolved = resolveTemplate(currentTpl, currentAnswers, prefixInp.value);
            const toCreate = resolved.filter(t => !uncheckedIds.has(t._id));
            if (!toCreate.length) return;

            // Collision detection — by (device, address), not by name. Modbus
            // uniqueness is per register address; a name collision would be caught
            // by WebPort anyway, but address collisions silently break existing
            // tags so they're the real problem to flag up-front.
            const addrMap = _tplExistingAddressesByDevice();
            const collisions = [];
            for (const t of toCreate) {
                const key = `${t.device || ''}|${t.address}`;
                if (addrMap.has(key)) collisions.push({ tag: t, existingName: addrMap.get(key) });
            }
            if (collisions.length) {
                const sample = collisions.slice(0, 5)
                    .map(c => `${c.tag.address} → redan ${c.existingName}`)
                    .join('\n  ');
                const more = collisions.length > 5 ? `\n  … (+${collisions.length - 5} till)` : '';
                const ok = confirm(
                    `${collisions.length} Modbus-adress(er) används redan av befintliga taggar:\n\n  ${sample}${more}\n\nFortsätt och hoppa över dem?`
                );
                if (!ok) return;
                const skipAddrs = new Set(collisions.map(c => `${c.tag.device || ''}|${c.tag.address}`));
                for (let i = toCreate.length - 1; i >= 0; i--) {
                    const key = `${toCreate[i].device || ''}|${toCreate[i].address}`;
                    if (skipAddrs.has(key)) toCreate.splice(i, 1);
                }
                if (!toCreate.length) return;
            }

            addBtn.disabled = true;
            cancelBtn.disabled = true;
            progEl.classList.add('on');
            progText.textContent = `Skapar 0 / ${toCreate.length}...`;
            progFill.style.width = '0%';

            const failing = await createTagBatch(toCreate, seedTag, (done, total, fails) => {
                progText.textContent = `Skapar ${done} / ${total}...` + (fails.length ? ` (${fails.length} fel)` : '');
                progFill.style.width = (done / total * 100).toFixed(1) + '%';
            });

            if (failing.length === toCreate.length) {
                setStatus('Alla taggar misslyckades. Kontrollera konsolen och seed-taggen.', true);
                addBtn.disabled = false; cancelBtn.disabled = false;
                progEl.classList.remove('on');
                return;
            }

            toastOk(`${toCreate.length - failing.length} taggar skapade` + (failing.length ? `, ${failing.length} fel (se aktivitetsloggen)` : ''));
            m.remove();
            // Refresh the page to show new tags (same mechanism as the bulk flow)
            const sc = document.createElement('script');
            sc.textContent = 'location.reload()';
            document.head.appendChild(sc);
        });

        // Initial load
        populateModelDropdown();
        loadModel();
    }



    // ============================================================
    // TOOLBAR + SUMMARY
    // ============================================================
    let tbEl=null, sumEl=null, bSel=null, bPSel=null, bApply=null, bInfo=null, bDirty=null, bClear=null, bDupeInfo=null;

    function createToolbar() {
        tbEl=document.createElement('div');tbEl.className='itb';
        sumEl=document.createElement('div');sumEl.className='ism';
        const page=document.querySelector('.page.full.fulltable');
        if(page&&page.parentElement) {
            page.parentElement.insertBefore(tbEl,page);
            page.parentElement.insertBefore(sumEl,page);
        } else {
            const wr=document.querySelector('#tagtable_wrapper')||document.getElementById('tagtable')?.parentElement;
            if(wr){wr.insertBefore(sumEl,wr.firstChild);wr.insertBefore(tbEl,wr.firstChild);}
        }

        sumEl.addEventListener('click', e => {
            const pill = e.target.closest('.inu-fpill');
            if (!pill) return;
            delete filterColVals[pill.dataset.col];
            updateColFilterIndicators();
            applyFilter();
            updSummary();
        });

        bSel=document.createElement('input');bSel.type='checkbox';bSel.title='Markera alla';
        bSel.style.cssText='width:13px;height:13px;cursor:pointer;';
        bSel.addEventListener('change',()=>selAll(bSel.checked));
        tbEl.appendChild(bSel);

        bInfo=document.createElement('span');bInfo.className='inf';bInfo.textContent='0 valda';
        tbEl.appendChild(bInfo);
        bClear=document.createElement('button');bClear.className='sec';bClear.innerHTML='<i class="fa fa-times"></i>';bClear.title='Avmarkera alla';bClear.style.cssText='display:none;padding:2px 6px;font-size:10px;';
        bClear.addEventListener('click',()=>{if(bSel)bSel.checked=false;selAll(false);});
        tbEl.appendChild(bClear);

        bDupeInfo=document.createElement('span');bDupeInfo.className='inu-dupe-info';bDupeInfo.style.display='none';bDupeInfo.title='Klicka för att filtrera/visa alla';
        bDupeInfo.addEventListener('click',()=>{
            filterMode = filterMode === 'dupe' ? 'all' : 'dupe';
            applyFilter();
        });
        tbEl.appendChild(bDupeInfo);

        tbEl.appendChild(Object.assign(document.createElement('span'),{className:'td'}));

        bPSel=document.createElement('select');
        tbEl.appendChild(bPSel);
        refreshBulkOpts();

        bApply=document.createElement('button');bApply.textContent='Applicera på valda';bApply.disabled=true;
        bApply.addEventListener('click',bulkApply);
        tbEl.appendChild(bApply);

        tbEl.appendChild(Object.assign(document.createElement('span'),{className:'td'}));

        // Batch field editor
        const DROPDOWN_FIELDS = ['device','datatype'];
        const TOGGLE_FIELDS = ['isalarm_on','isalarm_off','istrend_on','istrend_off'];
        const bfSel=document.createElement('select');bfSel.className='fil';
        [{v:'',l:'Redigera fält...'},{v:'isalarm_on',l:'Larm PÅ'},{v:'isalarm_off',l:'Larm AV'},{v:'istrend_on',l:'Trend PÅ'},{v:'istrend_off',l:'Trend AV'},
         {v:'device',l:'IO-Enhet'},{v:'datatype',l:'Datatyp'},
         {v:'unit',l:'Enhet'},{v:'format',l:'Format'},{v:'description',l:'Beskrivning'},
         {v:'p',l:'Larmklass'},{v:'rawmin',l:'Rå-min'},{v:'rawmax',l:'Rå-max'},{v:'engmin',l:'Vy-min'},{v:'engmax',l:'Vy-max'}]
            .forEach(f=>{const o=document.createElement('option');o.value=f.v;o.textContent=f.l;bfSel.appendChild(o);});
        tbEl.appendChild(bfSel);
        const bfValWrap=document.createElement('span');bfValWrap.style.display='none';
        tbEl.appendChild(bfValWrap);
        const bfApply=document.createElement('button');bfApply.className='sec';bfApply.textContent='Sätt';bfApply.style.display='none';
        tbEl.appendChild(bfApply);

        bfSel.addEventListener('change', async ()=>{
            const field=bfSel.value;
            bfValWrap.innerHTML='';
            if(!field){ bfValWrap.style.display='none'; bfApply.style.display='none'; return; }
            bfApply.style.display='';
            if(TOGGLE_FIELDS.includes(field)){
                bfValWrap.style.display='none';
            } else if(DROPDOWN_FIELDS.includes(field)){
                bfValWrap.style.display='';
                const opts=await getSelectOpts();
                const sel2=document.createElement('select');sel2.className='fil';sel2.id='bf-val';
                for(const o of (opts[field]||[])){
                    const opt=document.createElement('option');opt.value=o.value;opt.textContent=o.text;sel2.appendChild(opt);
                }
                bfValWrap.appendChild(sel2);
            } else {
                const inp=document.createElement('input');inp.id='bf-val';
                inp.style.cssText='width:80px;padding:2px 5px;border:1px solid #ccc;border-radius:4px;font-size:11px;';
                inp.placeholder='Värde';
                bfValWrap.appendChild(inp);
            }
        });

        bfApply.addEventListener('click', async ()=>{
            const field=bfSel.value;
            const rows=getSelectedRows();
            if(!field||!rows.length) return;
            bfApply.disabled=true; bfApply.textContent='Sparar...';
            let ok=0,fail=0;

            // Handle toggle fields (alarm/trend on/off)
            if(TOGGLE_FIELDS.includes(field)){
                const isAlarm=field.startsWith('isalarm');
                const enable=field.endsWith('_on');
                const formField=isAlarm?'isalarm':'istrend';
                for(const row of rows){
                    const tag=row.cells[C.NAME]?.textContent?.trim();
                    if(!tag) continue;
                    row.classList.add('inu-saving');
                    try {
                        await fetchFormAndSave(tag, fd => setCheckboxField(fd, formField, enable));
                        // Reflect in the UI checkbox + cache (without dispatching change to avoid re-save)
                        const chkSel=isAlarm?'.larm-tog:not(.trend) input':'.larm-tog.trend input';
                        const chk=row.querySelector(chkSel);
                        if(chk) chk.checked=enable;
                        // Update the visual group state
                        const grp=row.querySelector(isAlarm?'.p-larm .tog-grp':'.p-trend .tog-grp');
                        if(grp) grp.classList.toggle('on', enable);
                        if(isAlarm){ row.querySelector('.larm-wrap')?.classList.toggle('off', !enable); }
                        if(isAlarm) alarmCache[tag]=enable; else trendCache[tag]=enable;
                        // Sync bell/trend icon in Typ column
                        const typCell=row.cells[C.DESC+1];
                        if(typCell){
                            const imgSel=isAlarm?'img.alarmlink':'img.trendlink';
                            let img=row.querySelector(imgSel);
                            if(enable && !img){
                                img=document.createElement('img');
                                img.className=isAlarm?'alarmlink':'trendlink';
                                img.src=isAlarm?'/images/ico16/bell.svg':'/images/ico16/chart_curve.svg';
                                typCell.insertBefore(img, typCell.firstChild);
                            }
                            if(img) img.style.display=enable?'':'none';
                        }
                        ok++;
                    } catch(e){ console.warn(CFG.logPrefix,'Batch toggle failed',tag,e); fail++; }
                    row.classList.remove('inu-saving');
                }
            } else {
                const valEl=bfValWrap.querySelector('#bf-val');
                if(!valEl){ bfApply.disabled=false; bfApply.textContent='Sätt'; return; }
                const isDropdown=valEl.tagName==='SELECT';
                const val=isDropdown?valEl.value:valEl.value.trim();
                const displayVal=isDropdown?(valEl.options[valEl.selectedIndex]?.text||val):val;
                const colMap={device:C.IO,datatype:C.DTYPE,rawmin:C.RMIN,rawmax:C.RMAX,engmin:C.EMIN,engmax:C.EMAX,unit:C.UNIT,format:C.FMT,description:C.DESC};
                for(const row of rows){
                    const tag=row.cells[C.NAME]?.textContent?.trim();
                    if(!tag) continue;
                    row.classList.add('inu-saving');
                    try {
                        await fetchFormAndSave(tag, fd=>fd.set(field,val));
                        if(colMap[field]&&row.cells[colMap[field]]) row.cells[colMap[field]].textContent=displayVal;
                        colorRow(row); ok++;
                    } catch(e){ console.warn(CFG.logPrefix,'Batch edit failed',tag,e); fail++; }
                    row.classList.remove('inu-saving');
                }
            }
            updSummary(); updDirty();
            if (field === 'address' || field === 'device') detectDuplicates();
            toastOk(`${ok} uppdaterade`+(fail?`, ${fail} misslyckades`:''));
            bfApply.disabled=false; bfApply.textContent='Sätt';
        });

        tbEl.appendChild(Object.assign(document.createElement('span'),{className:'td'}));

        const bFil=document.createElement('select');bFil.className='fil';
        [{v:'all',l:'Visa alla'},{v:'unconf',l:'Okonfigurerade'},{v:'conf',l:'Konfigurerade'},{v:'digital',l:'Digitala'}]
            .forEach(f=>{const o=document.createElement('option');o.value=f.v;o.textContent=f.l;bFil.appendChild(o);});
        bFil.addEventListener('change',()=>{filterMode=bFil.value;applyFilter();});
        tbEl.appendChild(bFil);

        tbEl.appendChild(Object.assign(document.createElement('span'),{className:'td'}));

        const bPg=document.createElement('select');bPg.className='fil';
        [{v:'50',l:'50 per sida'},{v:'100',l:'100 per sida'},{v:'250',l:'250 per sida'},{v:'-1',l:'Visa alla'}]
            .forEach(f=>{const o=document.createElement('option');o.value=f.v;o.textContent=f.l;bPg.appendChild(o);});
        bPg.addEventListener('change',()=>{
            const len=bPg.value;
            const sc=document.createElement('script');
            sc.textContent=`(function(){try{var t=$('#tagtable').dataTable();t.fnSettings()._iDisplayLength=${len};t.fnDraw();}catch(e){console.error('[INU WP+]',e);}})();`;
            document.head.appendChild(sc);
            sc.remove();
        });
        tbEl.appendChild(bPg);

        tbEl.appendChild(Object.assign(document.createElement('span'),{className:'td'}));

        const mb=document.createElement('button');mb.innerHTML='<i class="fa fa-television"></i> Övervaka';
        mb.addEventListener('click',()=>monitorPrompt());
        tbEl.appendChild(mb);

        const krBtn=document.createElement('button');krBtn.className='sec';krBtn.innerHTML='<i class="fa fa-bar-chart"></i> Konfig-Rapport';
        krBtn.addEventListener('click',()=>openKonfigRapport());
        tbEl.appendChild(krBtn);

        const srBtn=document.createElement('button');srBtn.className='sec';srBtn.innerHTML='<i class="fa fa-exchange"></i> Sök/Ersätt';
        srBtn.addEventListener('click',()=>searchReplaceModal());
        tbEl.appendChild(srBtn);

        const pidBtn=document.createElement('button');pidBtn.className='sec';pidBtn.innerHTML='<i class="fa fa-line-chart"></i> PID-tuner (Experimentell)';
        pidBtn.title='Analysera PID-reglerkrets utifrån historisk trenddata';
        pidBtn.addEventListener('click',()=>showPidAdvisor());
        tbEl.appendChild(pidBtn);

        const tplBtn=document.createElement('button');tplBtn.innerHTML='<i class="fa fa-cubes"></i> Lägg till från mall';
        tplBtn.title='Skapa taggar från en fördefinierad enhetsmall (IVProdukt AHU, NIBE SMO S40, …)';
        tplBtn.addEventListener('click',()=>openTemplateModal());
        tbEl.appendChild(tplBtn);

        tbEl.appendChild(Object.assign(document.createElement('span'),{className:'td'}));

        const delSelBtn=document.createElement('button');delSelBtn.className='danger';delSelBtn.innerHTML='<i class="fa fa-trash"></i> Ta bort valda';
        delSelBtn.addEventListener('click',()=>{
            const sel=[...document.querySelectorAll('#tagtable tbody tr.tag.inu-sel:not(.inu-del)')];
            if(!sel.length){toastErr('Inga rader markerade');return;}
            sel.forEach(r=>markForDelete(r));
            toastOk(sel.length+' tagg(ar) markerade för borttagning');
        });
        tbEl.appendChild(delSelBtn);

        const undoSelBtn=document.createElement('button');undoSelBtn.className='sec';undoSelBtn.innerHTML='<i class="fa fa-undo"></i> Ångra valda';
        undoSelBtn.title='Ångra alla ändringar (skalning + borttagning) på markerade rader';
        undoSelBtn.addEventListener('click',()=>{
            const sel=[...document.querySelectorAll('#tagtable tbody tr.tag.inu-sel')];
            if(!sel.length){toastErr('Inga rader markerade');return;}
            sel.forEach(r=>{
                if(r.classList.contains('inu-del')) unmarkForDelete(r);
                undoRow(r);
            });
            toastOk('Ångrat ändringar på '+sel.length+' tagg(ar)');
        });
        tbEl.appendChild(undoSelBtn);

        bDirty=document.createElement('span');bDirty.className='inu-dirty';bDirty.style.display='none';

        // Help tooltip
        const help=document.createElement('span');help.className='p-help';help.textContent='?';
        help.innerHTML=`?<div class="p-tip">
<h4>INU WebPort-Plus v${CFG.version}</h4>
Verktyg för snabb skalningskonfigurering av taggar i INU WebPort.

<h5>Funktioner</h5>
<table>
<tr><td><b>Bulk-applicering</b></td><td>Markera rader med kryssrutor, välj preset i verktygsfältet och applicera på alla valda. Auto-förslag baserat på taggnamn.</td></tr>
<tr><td><b>Ångra</b></td><td>↩-knappen på raden återställer till tidigare skalning.</td></tr>
<tr><td><b>Kopiera/klistra</b></td><td>Kopiera skalning från en rad och klistra in på andra via preset-menyn.</td></tr>
<tr><td><b>Filter</b></td><td>Filtrera tabellen på okonfigurerade, konfigurerade eller digitala taggar.</td></tr>
<tr><td><b>Larm/Trend</b></td><td>Ändra larmklass (OKLASSAD/A/B/C/D), slå på/av larm och trend per tagg.</td></tr>
<tr><td><b>Egna presets</b></td><td>Skapa och ta bort presets via verktygsfältet. Sparas persistent.</td></tr>
<tr><td><b>Spara-kontroll</b></td><td>Sammanfattning av alla ändringar visas innan WebPort-sparning.</td></tr>
</table>

<h5>Tangentbordsgenvägar</h5>
<table>
<tr><td><span class="pk">Ctrl</span>+<span class="pk">S</span></td><td>Spara (visar sammanfattning)</td></tr>
<tr><td><span class="pk">Ctrl</span>+<span class="pk">A</span></td><td>Markera alla synliga rader</td></tr>
<tr><td><span class="pk">Esc</span></td><td>Stäng meny / avmarkera alla</td></tr>
</table>

<h5>Statusfärger</h5>
<table>
<tr><td><span class="dt g" style="display:inline-block;width:9px;height:9px;border-radius:50%;background:#4caf50;"></span> Grön</td><td>Konfigurerad (har skalning)</td></tr>
<tr><td><span class="dt w" style="display:inline-block;width:9px;height:9px;border-radius:50%;background:#ff9800;"></span> Orange</td><td>Okonfigurerad (alla värden 0)</td></tr>
<tr><td><span class="dt x" style="display:inline-block;width:9px;height:9px;border-radius:50%;background:#90a4ae;"></span> Grå</td><td>Digital (kräver ej skalning)</td></tr>
</table>
</div>`;
        tbEl.appendChild(bDirty);
        tbEl.appendChild(help);
    }

    function refreshBulkOpts() {
        if(!bPSel) return;
        bPSel.innerHTML='';
        for(const n of Object.keys(allPresets())) {
            const o=document.createElement('option');o.value=n;o.textContent=n;bPSel.appendChild(o);
        }
    }

    function updToolbar() {
        if(bInfo) bInfo.textContent=sel.size+' valda';
        if(bApply) bApply.disabled=sel.size===0;
        if(bClear) bClear.style.display=sel.size>0?'':'none';
    }

    function updDirty() {
        const n = Object.keys(sessionChanges).length;
        if (!bDirty) return;
        if (n > 0) {
            bDirty.innerHTML = `<i class="fa fa-exclamation-circle"></i> ${n} osparad(e)`;
            bDirty.style.display = '';
        } else {
            bDirty.style.display = 'none';
        }
    }

    function bulkApply() {
        const name=bPSel?.value; if(!name) return;
        const rows=getSelectedRows(); if(!rows.length) return;
        const p=allPresets()[name]; if(!p) return;
        const m=modal(`
<h3>Bekräfta bulk-applicering</h3>
<p style="font-size:12px;color:#333;margin:8px 0;">Applicera <b>${name}</b> på <b>${rows.length}</b> valda tagg(ar)?</p>
<div class="bt">
<button class="bx" id="pc">Avbryt</button>
<button class="bok" id="ps">Applicera</button>
</div>`);
        m.querySelector('#pc').addEventListener('click',()=>m.remove());
        m.querySelector('#ps').addEventListener('click',async()=>{
            m.remove();
            bApply.disabled=true; bApply.textContent=`Sparar 0/${rows.length}...`;
            let ok=0,fail=0;
            for(let i=0;i<rows.length;i++) {
                const row=rows[i], tag=row.cells[C.NAME].textContent.trim();
                bApply.textContent=`Sparar ${i+1}/${rows.length}...`;
                try {
                    const old=scl(row);
                    saveUndo(tag,old);
                    if(!sessionChanges[tag]) sessionChanges[tag]={old,presetName:name};
                    await apiSave(tag,p);
                    updCells(row,p);colorRow(row);updUndo(row);
                    ok++;
                } catch(e){ console.warn(CFG.logPrefix, 'Bulk save failed for', tag, e); fail++; }
            }
            bApply.textContent='Applicera på valda';bApply.disabled=false;
            updSummary(); updDirty();
            toastOk(`${ok} sparade` + (fail?`, ${fail} misslyckades`:''));
        });
    }

    function updSummary() {
        if(!sumEl) return;
        // Current page counts
        const rows=document.querySelectorAll('#tagtable tbody tr.tag');
        let tot=0,conf=0,unc=0,dig=0;
        rows.forEach(r=>{
            tot++;
            const dt=r.cells[C.DTYPE]?.textContent?.trim();
            if(dt==='DIGITAL') dig++;
            else if(unconf(r)) unc++;
            else conf++;
        });

        // Total counts via DataTables server info
        let totalHtml='';
        try {
            const sc=document.createElement('script');
            sc.textContent=`document.getElementById('_inu_dt_info').textContent=JSON.stringify($('#tagtable').dataTable().fnSettings()._iRecordsTotal||0);`;
            const infoEl=document.createElement('span');infoEl.id='_inu_dt_info';infoEl.style.display='none';
            document.body.appendChild(infoEl);
            document.head.appendChild(sc);sc.remove();
            const totalRows=JSON.parse(infoEl.textContent||'0');
            infoEl.remove();
            if(totalRows>0 && totalRows!==tot) {
                totalHtml=`<span class="inf">| Totalt ${totalRows} taggar</span>`;
            }
        } catch(e){ console.debug(CFG.logPrefix, 'DataTables info unavailable'); }

        const labels = COL_FILTER_LABELS();
        const pillsHtml = Object.entries(filterColVals)
            .filter(([,vals]) => vals && vals.size > 0)
            .map(([ci, vals]) => {
                const label = labels[+ci] || `Kol ${ci}`;
                const valStr = [...vals].sort().join(', ');
                return `<span class="inu-fpill" data-col="${ci}"><i class="fa fa-times"></i> ${escHtml(label)}: ${escHtml(valStr)}</span>`;
            }).join('');
        sumEl.innerHTML=`
<span><span class="dt g"></span> ${conf} konfig.</span>
<span><span class="dt w"></span> ${unc} okonf.</span>
<span><span class="dt x"></span> ${dig} dig.</span>
<span class="inf">| Sida: ${tot}</span>
${totalHtml}${pillsHtml ? `<span class="inf">| Aktiva filter:</span>${pillsHtml}` : ''}`;
    }

    // ============================================================
    // COLUMN SORT (for our prepended columns)
    // ============================================================
    let sortState = {}; // colIdx → 'asc'|'desc'
    function sortByColumn(colIdx, type) {
        const tbody = document.querySelector('#tagtable tbody');
        if (!tbody) return;
        const dir = sortState[colIdx] === 'asc' ? 'desc' : 'asc';
        sortState = {}; sortState[colIdx] = dir;
        const rows = Array.from(tbody.querySelectorAll('tr.tag'));
        rows.sort((a, b) => {
            let va, vb;
            if (type === 'larm') {
                // Sort by: enabled first, then class A>B>C>D>OKLASSAD
                const classRank = {A:1,B:2,C:3,D:4,OKLASSAD:5,'':5,'-':5};
                const aOn = a.querySelector('.larm-tog:not(.trend) input')?.checked ? 0 : 1;
                const bOn = b.querySelector('.larm-tog:not(.trend) input')?.checked ? 0 : 1;
                const aCls = a.querySelector('.larm-sel')?.value?.toUpperCase() || '';
                const bCls = b.querySelector('.larm-sel')?.value?.toUpperCase() || '';
                va = aOn * 10 + (classRank[aCls] || 5);
                vb = bOn * 10 + (classRank[bCls] || 5);
            } else if (type === 'trend') {
                va = a.querySelector('.larm-tog.trend input')?.checked ? 0 : 1;
                vb = b.querySelector('.larm-tog.trend input')?.checked ? 0 : 1;
            }
            return dir === 'asc' ? va - vb : vb - va;
        });
        for (const r of rows) tbody.appendChild(r);
        // Update sort icons
        document.querySelectorAll('.p-sort i.fa').forEach(i => { i.className = 'fa fa-sort'; i.style.opacity = '.3'; });
        const th = document.querySelectorAll('#tagtable thead .p-sort')[type === 'larm' ? 0 : 1];
        if (th) { const i = th.querySelector('i.fa'); if (i) { i.className = 'fa fa-sort-' + (dir === 'asc' ? 'up' : 'down'); i.style.opacity = '1'; } }
    }

    // ============================================================
    // BUILD COLUMNS
    // ============================================================
    function addColumns() {
        const table=document.getElementById('tagtable');if(!table)return;

        // Headers (prepend — reverse order so checkbox ends up first)
        table.querySelectorAll('thead tr').forEach(tr=>{
            if(tr.querySelector('.p-hdr'))return;
            const th0=document.createElement('th');th0.className='p-hdr p-chk';th0.textContent='';th0.style.cssText='width:20px;';
            const thL=document.createElement('th');thL.className='p-hdr p-sort';thL.innerHTML='Larm <i class="fa fa-sort" style="opacity:.3"></i>';thL.style.cssText='width:90px;cursor:pointer;';
            const thT=document.createElement('th');thT.className='p-hdr p-sort';thT.innerHTML='Trend <i class="fa fa-sort" style="opacity:.3"></i>';thT.style.cssText='width:50px;cursor:pointer;padding-right:10px !important;';
            [th0,thL,thT].forEach(th=>th.addEventListener('click',e=>e.stopPropagation()));
            thL.addEventListener('click',()=>sortByColumn(1,'larm'));
            thT.addEventListener('click',()=>sortByColumn(2,'trend'));
            tr.insertBefore(thT, tr.firstChild);
            tr.insertBefore(thL, tr.firstChild);
            tr.insertBefore(th0, tr.firstChild);
        });

        // Rows
        table.querySelectorAll('tbody tr.tag').forEach(row=>{
            if(row.querySelector('.p-larm'))return;

            // Checkbox
            const td0=document.createElement('td');td0.className='p-chk';
            const cb=document.createElement('input');cb.type='checkbox';
            cb.addEventListener('click',e=>{e.stopPropagation();e.preventDefault();});
            td0.appendChild(cb);

            // Larm column: alarm toggle + class badge
            const tdL=document.createElement('td');tdL.className='p-larm';tdL.style.cssText='white-space:nowrap;';
            tdL.addEventListener('click',e=>e.stopPropagation());
            const wrap=document.createElement('div');wrap.className='larm-wrap off';
            const aGrp=document.createElement('span');aGrp.className='tog-grp';aGrp.title='Larm på/av';
            const aIco=document.createElement('i');aIco.className='fa fa-bell tog-ico';aIco.title='Larm på/av';
            aIco.addEventListener('click',()=>{aChk.checked=!aChk.checked;aChk.dispatchEvent(new Event('change'));});
            const tog=document.createElement('label');tog.className='larm-tog';
            const aChk=document.createElement('input');aChk.type='checkbox';
            aChk.addEventListener('click',e=>e.stopPropagation());
            const sl=document.createElement('span');sl.className='sl';
            tog.appendChild(aChk);tog.appendChild(sl);
            aGrp.appendChild(aIco);aGrp.appendChild(tog);
            wrap.appendChild(aGrp);
            const lSel=document.createElement('select');lSel.className='larm-sel';
            lSel.innerHTML='<option value="OKLASSAD">-</option><option value="A">A</option><option value="B">B</option><option value="C">C</option><option value="D">D</option>';
            wrap.appendChild(lSel);
            tdL.appendChild(wrap);

            // Trend column: trend toggle
            const tdT=document.createElement('td');tdT.className='p-trend';tdT.style.cssText='white-space:nowrap;text-align:center;';
            tdT.addEventListener('click',e=>e.stopPropagation());
            const tGrp=document.createElement('span');tGrp.className='tog-grp';tGrp.title='Trend på/av';
            const tTog=document.createElement('label');tTog.className='larm-tog trend';
            const tChk=document.createElement('input');tChk.type='checkbox';
            tChk.addEventListener('click',e=>e.stopPropagation());
            const tSl=document.createElement('span');tSl.className='sl';
            tTog.appendChild(tChk);tTog.appendChild(tSl);
            const tIco=document.createElement('i');tIco.className='fa fa-line-chart tog-ico';tIco.title='Trend på/av';
            tIco.addEventListener('click',()=>{tChk.checked=!tChk.checked;tChk.dispatchEvent(new Event('change'));});
            tGrp.appendChild(tTog);tGrp.appendChild(tIco);
            tdT.appendChild(tGrp);

            function syncLarmStyle(){
                const v=lSel.value.toUpperCase();
                lSel.className='larm-sel'+(v==='A'?' lk-a':v==='B'?' lk-b':v==='C'?' lk-c':v==='D'?' lk-d':'');
                wrap.classList.toggle('off',!aChk.checked);
                aGrp.classList.toggle('on',aChk.checked);
                tGrp.classList.toggle('on',tChk.checked);
            }
            lSel.addEventListener('change',syncLarmStyle);
            aChk.addEventListener('change',syncLarmStyle);
            tChk.addEventListener('change',syncLarmStyle);

            // Prepend columns (reverse order so td0=checkbox ends up first)
            row.insertBefore(tdT, row.firstChild);
            row.insertBefore(tdL, row.firstChild);
            row.insertBefore(td0, row.firstChild);

            // Read tag name AFTER prepending (C indices are now shifted +3)
            const tag=row.cells[C.NAME]?.textContent?.trim();

            // Snapshot after columns added so C indices are correct
            if(tag && !initialSnapshot[tag]) initialSnapshot[tag]=fullSnap(row);

            // Wire up larm + trend handlers (need tag)
            const typCell=row.cells[14]; // Typ column
            aChk.addEventListener('change',()=>{
                if(!tag) return;
                let bell=row.querySelector('img.alarmlink');
                if(aChk.checked&&!bell&&typCell){
                    bell=document.createElement('img');bell.className='alarmlink';bell.src='/images/ico16/bell.svg';
                    typCell.insertBefore(bell,typCell.firstChild);
                }
                if(bell) bell.style.display=aChk.checked?'':'none';
                row.classList.add('inu-saving'); aIco.classList.add('inu-spin');
                saveIsAlarm(tag,aChk.checked).finally(()=>{ row.classList.remove('inu-saving'); aIco.classList.remove('inu-spin'); });
            });
            tChk.addEventListener('change',()=>{
                if(!tag) return;
                let icon=row.querySelector('img.trendlink');
                if(tChk.checked&&!icon&&typCell){
                    icon=document.createElement('img');icon.className='trendlink';icon.src='/images/ico16/chart_curve.svg';
                    typCell.appendChild(icon);
                }
                if(icon) icon.style.display=tChk.checked?'':'none';
                row.classList.add('inu-saving'); tIco.classList.add('inu-spin');
                saveIsTrend(tag,tChk.checked).finally(()=>{ row.classList.remove('inu-saving'); tIco.classList.remove('inu-spin'); });
            });
            lSel.addEventListener('change',()=>{
                if(!tag) return;
                row.classList.add('inu-saving');
                saveLarmklass(tag,lSel.value).finally(()=>row.classList.remove('inu-saving'));
            });
            if(tag) loadLarmklass(tag,lSel,aChk,tChk,row).then(syncLarmStyle);

            colorRow(row);
            updUndo(row);
        });
        detectDuplicates();
    }

    // ============================================================
    // DUPLICATE ADDRESS DETECTION
    // ============================================================
    function detectDuplicates() {
        const map = {}; // "device|address" → [rows]
        const rows = document.querySelectorAll('#tagtable tbody tr.tag');
        rows.forEach(r => {
            const dev = r.cells[C.IO]?.textContent?.trim() || '';
            const addr = r.cells[C.ADDR]?.textContent?.trim() || '';
            if (!dev || !addr || addr === '0') return; // skip unconfigured
            const key = dev + '|' + addr;
            if (!map[key]) map[key] = [];
            map[key].push(r);
        });
        let dupeCount = 0;
        rows.forEach(r => r.classList.remove('inu-dupe'));
        for (const key in map) {
            if (map[key].length > 1) {
                dupeCount += map[key].length;
                map[key].forEach(r => {
                    r.classList.add('inu-dupe');
                    const addrCell = r.cells[C.ADDR];
                    if (addrCell) addrCell.title = `Duplicerad adress: ${map[key].length} taggar delar ${key.replace('|', ' @ ')}`;
                });
            }
        }
        if (bDupeInfo) {
            if (dupeCount > 0) { bDupeInfo.textContent = `⚠ ${dupeCount} dup. adresser`; bDupeInfo.style.display = ''; }
            else bDupeInfo.style.display = 'none';
        }
    }

    // ============================================================
    // TAG SOURCES OUT-OF-SYNC BANNER
    // ============================================================
    async function checkSources() {
        try {
            const r = await fetch('/tag/GetSourceList?draw=1&limit=500&offset=0&sortcol=0&sortdir=asc&search=');
            const data = await r.json();
            const dirty = [];
            const decode = h => { const d = document.createElement('div'); d.innerHTML = h; return d.textContent; };
            (data.aaData || []).forEach(row => {
                const state = String(row[2] || '').replace(/<[^>]+>/g, '').trim();
                if (state && state !== 'OK' && state !== 'Ändrades') {
                    const name = decode(String(row[0] || '')).trim();
                    dirty.push({ name, state });
                }
            });
            if (!dirty.length) return;
            showSourceBanner(dirty);
        } catch (e) {
            console.warn(CFG.logPrefix, 'checkSources failed', e);
        }
    }

    function showSourceBanner(dirty) {
        if (document.getElementById('inu-src-pill')) return;
        // Inject pill styles if injectStyles() hasn't run (non-tag/device pages)
        if (!document.getElementById('inu-src-pill-style')) {
            const s = document.createElement('style');
            s.id = 'inu-src-pill-style';
            s.textContent = '.inu-src-pill{padding:3px 9px;border-radius:3px;font-size:10px;font-weight:600;color:#fff;background:#b84700;cursor:pointer;align-self:center;display:inline-flex;align-items:center;gap:5px;text-decoration:none;white-space:nowrap;margin-right:8px;}.inu-src-pill:hover{background:#e65100;color:#fff;}';
            document.head.appendChild(s);
        }
        const nav = document.getElementById('top-menu')?.parentElement;
        if (!nav) return;
        const pill = document.createElement('a');
        pill.id = 'inu-src-pill';
        pill.className = 'inu-src-pill';
        pill.href = '/tag/sources';
        const tip = dirty.map(d => d.name + ' (' + d.state + ')').join('\n');
        pill.title = tip;
        pill.innerHTML = '<i class="fa fa-exclamation-triangle"></i> ' + dirty.length + ' st osparade tagglistor';
        nav.insertBefore(pill, nav.firstChild);
    }

    // ============================================================
    // SAVE INTERCEPT
    // ============================================================
    function hijackSave() {
        const btn=document.getElementById('wp_mnu_wp_tb_save');
        if(!btn) return;
        const origOnclick=btn.getAttribute('onclick');
        btn.removeAttribute('onclick');
        btn.addEventListener('click',e=>{
            e.stopPropagation();e.preventDefault();
            // Build current state map
            const currentRows=document.querySelectorAll('#tagtable tbody tr.tag');
            const currentMap={};
            currentRows.forEach(r=>{
                const name=r.cells[C.NAME]?.textContent?.trim();
                if(name) currentMap[name]=fullSnap(r);
            });

            let rows='';
            let editCount=0;
            const td='padding:3px 8px;font-size:11px;border-bottom:1px solid #eee;';

            // Detect modified tags (only tags visible on current page)
            for(const tag of Object.keys(currentMap)) {
                const snap=initialSnapshot[tag];
                const cur=currentMap[tag];
                if(!snap) continue; // New to this page view (pagination), skip
                const diffs=snapDiff(snap,cur);
                if(!diffs.length) continue;
                const hasScaling=diffs.some(d=>d==='Rå-skalning'||d==='Vy-skalning'||d==='Enhet'||d==='Format');
                const nonScaling=diffs.filter(d=>d!=='Rå-skalning'&&d!=='Vy-skalning'&&d!=='Enhet'&&d!=='Format');
                let detailRows='';
                if(hasScaling) {
                    detailRows+=`<div style="display:flex;gap:6px;align-items:center;padding:2px 0;">
                        <span style="font-size:10px;font-weight:600;color:#888;min-width:60px;">Skalning</span>
                        <span style="color:#999;">${escHtml(snap.engmin)}...${escHtml(snap.engmax)} ${escHtml(snap.unit)} [${escHtml(snap.format)}]</span>
                        <span style="color:#aaa;">→</span>
                        <span style="color:#2e7d32;font-weight:600;">${escHtml(cur.engmin)}...${escHtml(cur.engmax)} ${escHtml(cur.unit)} [${escHtml(cur.format)}]</span>
                    </div>`;
                }
                nonScaling.forEach(f=>{
                    const key=f==='IO-Enhet'?'io':f==='Adress'?'addr':f==='Datatyp'?'dtype':f==='Beskrivning'?'desc':'';
                    if(key) detailRows+=`<div style="display:flex;gap:6px;align-items:center;padding:2px 0;">
                        <span style="font-size:10px;font-weight:600;color:#888;min-width:60px;">${escHtml(f)}</span>
                        <span style="color:#999;">${escHtml(snap[key]||'(tom)')}</span>
                        <span style="color:#aaa;">→</span>
                        <span style="color:#2e7d32;font-weight:600;">${escHtml(cur[key]||'(tom)')}</span>
                    </div>`;
                });
                rows+=`<tr><td style="${td}" colspan="3">
                    <div style="font-weight:600;font-size:12px;">${escHtml(tag)}</div>
                    <div style="padding:2px 0 2px 8px;font-size:11px;">${detailRows}</div>
                </td></tr>`;
                editCount++;
            }

            // Include preset changes applied on other pages
            for(const tag of Object.keys(sessionChanges)) {
                if(!currentMap[tag]) {
                    const ch=sessionChanges[tag];
                    const oldStr=`${escHtml(ch.old.engmin)}...${escHtml(ch.old.engmax)} ${escHtml(ch.old.unit)} [${escHtml(ch.old.format)}]`;
                    rows+=`<tr><td style="${td}">${escHtml(tag)}<span style="color:#5b6abf;font-size:10px;"> ${escHtml(ch.presetName)}</span></td>`+
                          `<td style="${td}" colspan="2"><span style="color:#999;">${oldStr}</span> → <span style="color:#2e7d32;font-weight:600;">(annan sida)</span></td></tr>`;
                    editCount++;
                }
            }

            // Pending deletes section
            const tdDel='padding:3px 8px;font-size:11px;border-bottom:1px solid #fee2e2;';
            let delRows='';
            for(const tag of pendingDeletes) {
                delRows+=`<tr><td style="${tdDel}" colspan="3"><i class="fa fa-trash" style="color:#ef4444;margin-right:6px;"></i><span style="color:#ef4444;font-weight:600;">${escHtml(tag)}</span></td></tr>`;
            }
            const delSection = pendingDeletes.size ? `
<div style="margin-top:10px;">
<table style="width:100%;border-collapse:collapse;">
<thead><tr><th style="text-align:left;padding:4px 8px;font-size:10px;color:#ef4444;border-bottom:2px solid #fee2e2;" colspan="3">Taggar att ta bort (${pendingDeletes.size})</th></tr></thead>
<tbody>${delRows}</tbody>
</table></div>` : '';

            if(!editCount && !pendingDeletes.size) {
                const sc=document.createElement('script');
                sc.textContent=origOnclick||'SaveChanges();';
                document.head.appendChild(sc);sc.remove();
                return;
            }

            const m=modal(`
<h3>Sammanfattning av ändringar</h3>
<div style="max-height:320px;overflow-y:auto;margin-bottom:12px;">
${editCount?`<table style="width:100%;border-collapse:collapse;">
<thead><tr>
<th style="text-align:left;padding:4px 8px;font-size:10px;color:#999;border-bottom:2px solid #ddd;" colspan="3">Ändrade taggar</th>
</tr></thead>
<tbody>${rows}</tbody>
</table>`:''}
${delSection}
</div>
<p style="font-size:11px;color:#666;margin:0 0 8px;">${editCount} tagg(ar) ändrade${pendingDeletes.size?`, ${pendingDeletes.size} tas bort`:''}</p>
<div class="bt">
<button class="bx" id="pc">Avbryt</button>
<button class="bok" id="ps">Spara</button>
</div>`);
            m.querySelector('#pc').addEventListener('click',()=>m.remove());
            m.querySelector('#ps').addEventListener('click',async ()=>{
                m.remove();
                // Delete pending tags
                if(pendingDeletes.size) {
                    const toDelete=[...pendingDeletes];
                    let ok=0,fail=0;
                    for(const tag of toDelete){
                        try{
                            await deleteTag(tag);
                            pendingDeletes.delete(tag);
                            ok++;
                            console.log(CFG.logPrefix,'Deleted:',tag);
                        } catch(e){
                            fail++;
                            console.error(CFG.logPrefix,'Delete failed:',tag,e);
                        }
                    }
                    if(fail) toastErr(fail+' borttagning(ar) misslyckades');
                    else toastOk(ok+' tagg(ar) borttagna');
                }
                Object.keys(sessionChanges).forEach(k=>delete sessionChanges[k]);
                if(editCount>0){
                    // Field changes: let SaveChanges() handle the reload
                    Object.keys(initialSnapshot).forEach(k=>delete initialSnapshot[k]);
                    document.querySelectorAll('#tagtable tbody tr.tag').forEach(r=>{
                        const t=rowTagName(r);
                        if(t) initialSnapshot[t]=fullSnap(r);
                    });
                    const sc=document.createElement('script');
                    sc.textContent=origOnclick||'SaveChanges();';
                    document.head.appendChild(sc);sc.remove();
                    toast('Sparar...'); updDirty();
                } else {
                    // Deletions only: refresh the table
                    const sc=document.createElement('script');
                    sc.textContent='try{oTable.fnDraw();}catch(e){}';
                    document.head.appendChild(sc);sc.remove();
                    updDirty();
                }
            });
        },true);
    }

    // ============================================================
    // KEYBOARD SHORTCUTS
    // ============================================================
    window.addEventListener('beforeunload', e => {
        if (Object.keys(sessionChanges).length > 0 || pendingDeletes.size > 0) { e.preventDefault(); e.returnValue = ''; }
    });

    document.addEventListener('keydown',e=>{
        // Don't trigger when typing in inputs
        if(e.target.tagName==='INPUT'||e.target.tagName==='TEXTAREA'||e.target.tagName==='SELECT') return;
        // Ctrl+S — trigger save
        if((e.ctrlKey||e.metaKey)&&e.key==='s'){
            e.preventDefault();
            document.getElementById('wp_mnu_wp_tb_save')?.click();
        }
        // Ctrl+A — select all visible
        if((e.ctrlKey||e.metaKey)&&e.key==='a'){
            e.preventDefault();
            if(bSel){bSel.checked=true;selAll(true);}
        }
        // Escape — deselect all
        if(e.key==='Escape'){
            if(sel.size>0){if(bSel)bSel.checked=false;selAll(false);}
        }
    });

    // ============================================================
    // CONTEXT MENU
    // ============================================================
    let scalingClip = null;

    // Map cell index → { field: form field name, type: 'text'|'select', selectName: form select name }
    const EDITABLE_CELLS = {
        [C.NAME]:  { field:'name', type:'text' },
        [C.IO]:    { field:'device', type:'select' },
        [C.ADDR]:  { field:'address', type:'text' },
        [C.DTYPE]: { field:'datatype', type:'select' },
        [C.RMIN]:  { field:'rawmin', type:'text' },
        [C.RMAX]:  { field:'rawmax', type:'text' },
        [C.EMIN]:  { field:'engmin', type:'text' },
        [C.EMAX]:  { field:'engmax', type:'text' },
        [C.UNIT]:  { field:'unit', type:'text' },
        [C.FMT]:   { field:'format', type:'text' },
        [C.DESC]:  { field:'description', type:'text' },
    };

    // Cache for dropdown options (fetched once from a tag form)
    let selectOptsCache = null;
    async function getSelectOpts() {
        if (selectOptsCache) return selectOptsCache;
        // Fetch any tag form to extract dropdown options
        const firstRow = document.querySelector('#tagtable tbody tr.tag');
        if (!firstRow) return {};
        const tag = firstRow.cells[C.NAME]?.textContent?.trim();
        if (!tag) return {};
        const r = await fetch('/tag/ActionEdit?show=1&type=tag&tag=' + encodeTag(tag));
        const doc = new DOMParser().parseFromString(await r.text(), 'text/html');
        selectOptsCache = {};
        for (const sel of doc.querySelectorAll('select')) {
            if (!sel.name) continue;
            selectOptsCache[sel.name] = [...sel.options].map(o => ({ value: o.value, text: o.textContent.trim() }));
        }
        return selectOptsCache;
    }

    async function startCellEdit(cell, row, cellDef) {
        if (cell.querySelector('input,select')) return;
        const tag = row.cells[C.NAME]?.textContent?.trim();
        if (!tag) return;
        const cur = cell.textContent.trim();

        let el;
        if (cellDef.type === 'select') {
            const opts = await getSelectOpts();
            const fieldOpts = opts[cellDef.field] || [];
            el = document.createElement('select');
            el.style.cssText = 'width:100%;font-size:11px;padding:1px 3px;border:1px solid #5b6abf;border-radius:2px;';
            for (const o of fieldOpts) {
                const opt = document.createElement('option');
                opt.value = o.value; opt.textContent = o.text;
                if (o.text === cur) opt.selected = true;
                el.appendChild(opt);
            }
        } else {
            el = document.createElement('input');
            el.value = cur;
            el.style.cssText = 'width:100%;font-size:11px;padding:1px 3px;border:1px solid #5b6abf;border-radius:2px;box-sizing:border-box;';
        }

        const blocker = e => { e.stopPropagation(); e.stopImmediatePropagation(); };
        cell.addEventListener('click', blocker, true);
        cell.addEventListener('mousedown', blocker, true);
        cell.textContent = '';
        cell.appendChild(el);
        el.focus();
        if (el.select) el.select();

        const save = async () => {
            cell.removeEventListener('click', blocker, true);
            cell.removeEventListener('mousedown', blocker, true);
            const nv = cellDef.type === 'select' ? el.options[el.selectedIndex]?.value : el.value.trim();
            const displayVal = cellDef.type === 'select' ? el.options[el.selectedIndex]?.text : nv;
            if (nv !== undefined && displayVal !== cur) {
                const oldSnap = scl(row); // capture BEFORE mutation
                cell.textContent = displayVal;
                row.classList.add('inu-saving');
                try {
                    await fetchFormAndSave(tag, fd => fd.set(cellDef.field, nv));
                    colorRow(row); updSummary();
                    if (!sessionChanges[tag]) sessionChanges[tag] = { old: oldSnap, presetName: '(redigerad)' };
                    updDirty();
                    if (cellDef.field === 'address' || cellDef.field === 'device') detectDuplicates();
                } catch (err) { cell.textContent = cur; toastErr(err.message); }
                row.classList.remove('inu-saving');
            } else {
                cell.textContent = displayVal;
            }
        };
        el.addEventListener('blur', save);
        el.addEventListener('keydown', ev => {
            if (ev.key === 'Enter') { ev.preventDefault(); el.blur(); }
            if (ev.key === 'Escape') {
                el.removeEventListener('blur', save);
                cell.removeEventListener('click', blocker, true);
                cell.removeEventListener('mousedown', blocker, true);
                cell.textContent = cur;
            }
        });
        if (cellDef.type === 'select') el.addEventListener('change', () => el.blur());
    }

    function initContextMenu() {
        // Track which cell was right-clicked
        let ctxCellIdx = -1;
        document.addEventListener('contextmenu', e => {
            const td = e.target.closest('#tagtable tbody td');
            ctxCellIdx = td ? td.cellIndex : -1;
        }, true);

        document.addEventListener('inu-ctx', e => {
            const { action, rowId, cellIndex } = e.detail;
            const row = document.getElementById(rowId) || document.querySelector(`#tagtable tbody tr[id="${rowId}"]`);
            if (!row) return;
            const tag = row.cells[C.NAME]?.textContent?.trim();
            if (!tag) return;

            if (action === 'edit') {
                const cellDef = EDITABLE_CELLS[cellIndex];
                const cell = row.cells[cellIndex];
                if (cellDef && cell) startCellEdit(cell, row, cellDef);
            } else if (action === 'copy') {
                scalingClip = scl(row);
                toastOk('Skalning kopierad');
            } else if (action === 'paste' && scalingClip) {
                const oldSnap = scl(row); // capture BEFORE mutation
                row.classList.add('inu-saving');
                apiSave(tag, scalingClip).then(() => {
                    updCells(row, scalingClip); colorRow(row); updSummary();
                    if (!sessionChanges[tag]) sessionChanges[tag] = { old: oldSnap, presetName: '(inklistrad)' };
                    updDirty(); toastOk('Skalning inklistrad → ' + tag);
                }).catch(e => toastErr(e.message)).finally(() => row.classList.remove('inu-saving'));
            } else if (action === 'alarm-on' || action === 'alarm-off') {
                const chk = row.querySelector('.larm-tog input[type="checkbox"]');
                if (chk) { chk.checked = action === 'alarm-on'; chk.dispatchEvent(new Event('change')); }
            } else if (action === 'trend-on' || action === 'trend-off') {
                const chk = row.querySelector('.larm-tog.trend input[type="checkbox"]');
                if (chk) { chk.checked = action === 'trend-on'; chk.dispatchEvent(new Event('change')); }
            } else if (action === 'undo') {
                undoRow(row);
            } else if (action === 'dup') {
                const sourceAddr = row.cells[C.ADDR]?.textContent?.trim() || '';
                openDuplicateDialog(tag, sourceAddr);
            } else if (action === 'delete') {
                markForDelete(row);
            } else if (action === 'undo-sel') {
                const sel = [...document.querySelectorAll('#tagtable tbody tr.tag.inu-sel')];
                const targets = sel.length ? sel : [row];
                targets.forEach(r => {
                    if (r.classList.contains('inu-del')) unmarkForDelete(r);
                    undoRow(r);
                });
                if (targets.length > 1) toastOk('Ångrat ändringar på ' + targets.length + ' tagg(ar)');
            }
        });

        const sc = document.createElement('script');
        sc.textContent = `(function(){
            if(!$.contextMenu) return;
            var OFF=${CFG.colOffset};
            var editableCols={};
            for(var i=OFF;i<=OFF+10;i++) editableCols[i]=true;
            var _ctxCell=-1;
            document.addEventListener('contextmenu',function(e){
                var td=e.target.closest('#tagtable tbody td');
                _ctxCell=td?td.cellIndex:-1;
            },true);
            $.contextMenu({
                selector:'#tagtable tbody tr.tag',
                build:function($trigger,e){
                    var items={};
                    if(editableCols[_ctxCell]){
                        items.edit={name:'Redigera fält',icon:'fa-pencil'};
                        items.sep0='---';
                    }
                    items.copy={name:'Kopiera skalning',icon:'fa-copy'};
                    items.paste={name:'Klistra in skalning',icon:'fa-paste'};
                    items.sep1='---';
                    items.alarmOn={name:'Larm PÅ',icon:'fa-bell'};
                    items.alarmOff={name:'Larm AV',icon:'fa-bell-slash'};
                    items.trendOn={name:'Trend PÅ',icon:'fa-line-chart'};
                    items.trendOff={name:'Trend AV',icon:'fa-ban'};
                    items.sep2='---';
                    items.undo={name:'Ångra',icon:'fa-undo'};
                    items.undoSel={name:'Ångra markerade',icon:'fa-undo'};
                    items.sep3='---';
                    items.dup={name:'Duplicera tagg',icon:'fa-clone'};
                    items.del={name:'Ta bort tagg',icon:'fa-trash'};
                    return {
                        items:items,
                        callback:function(key,opt){
                            var rowId=opt.\$trigger.attr('id');
                            var map={edit:'edit',copy:'copy',paste:'paste',alarmOn:'alarm-on',alarmOff:'alarm-off',trendOn:'trend-on',trendOff:'trend-off',undo:'undo',undoSel:'undo-sel',dup:'dup',del:'delete'};
                            document.dispatchEvent(new CustomEvent('inu-ctx',{detail:{action:map[key],rowId:rowId,cellIndex:_ctxCell}}));
                        }
                    };
                }
            });
        })();`;
        document.head.appendChild(sc); sc.remove();
    }

    // ============================================================
    // SEARCH INTEGRATION
    // ============================================================
    function hookSearch() {
        const searchInput=document.getElementById('wp_general_search_input');
        if(!searchInput) return;
        searchInput.addEventListener('input',()=>{
            setTimeout(()=>{addColumns();updSummary();applyFilter();syncSelCheckboxes();},100);
        });
        // Intercept keyup (capture phase, before DataTables) to support * wildcard
        searchInput.addEventListener('keyup', e => {
            const raw = searchInput.value;
            if (!raw.includes('*')) return; // no wildcard — let DataTables handle normally
            e.stopImmediatePropagation(); // prevent DataTables' plain-text handler from overriding
            const pattern = wildcardToPattern(raw);
            const sc = document.createElement('script');
            sc.textContent = `(function(){try{$('#tagtable').dataTable().fnFilter(${JSON.stringify(pattern)},null,true,false);}catch(e){console.error('[INU WP+]',e);}})();`;
            document.head.appendChild(sc); sc.remove();
            setTimeout(()=>{addColumns();updSummary();applyFilter();syncSelCheckboxes();},100);
        }, true); // capture = runs before DataTables' bubble-phase handler
    }

    // ============================================================
    // CELL INDEX PATCH — fix WebPort's click handlers after prepending columns
    // ============================================================
    function patchClickIndices() {
        const sc=document.createElement('script');
        sc.textContent=`(function(){
            var OFF=${CFG.colOffset},table=document.getElementById('tagtable');
            if(!table)return;
            table.addEventListener('click',function(e){
                var sel=window.getSelection();
                if(sel&&sel.toString().length>0){e.stopImmediatePropagation();e.preventDefault();return;}
                var td=e.target.closest?e.target.closest('td,th'):null;
                if(!td||!td.closest('#tagtable'))return;
                var idx=td.cellIndex;
                if(idx>=OFF){
                    Object.defineProperty(td,'cellIndex',{value:idx-OFF,configurable:true,writable:true});
                    setTimeout(function(){try{delete td.cellIndex;}catch(x){}},0);
                }
            },true);
            if(window.jQuery&&jQuery.fn){
                var _oi=jQuery.fn.index;
                jQuery.fn.index=function(){
                    if(arguments.length===0&&this.length===1&&this[0]){
                        var el=this[0];
                        if((el.tagName==='TD'||el.tagName==='TH')&&el.closest&&el.closest('#tagtable')){
                            var r=_oi.call(this);if(r>=OFF)return r-OFF;
                        }
                    }
                    return _oi.apply(this,arguments);
                };
            }
        })();`;
        document.head.appendChild(sc);sc.remove();
    }

    // ============================================================
    // PATCH VALUE UPDATE — fix nth-child(13) → nth-child(16) for prepended columns
    // ============================================================
    function patchValueUpdate() {
        const sc=document.createElement('script');
        sc.textContent=`(function(){
            var OFF=${CFG.colOffset},NTH=13+OFF;
            if(typeof UpdateAll==='function'){
                UpdateAll=function(enable){
                    var tlist=$('#taglist tr');
                    if(!enable) watchList=[];
                    tlist.each(function(){
                        var tid=$(this).attr('id');
                        if(enable){
                            if(watchList.indexOf(tid)===-1) watchList.push(tid);
                            $(this).find('td:nth-child('+NTH+')').addClass('green');
                        } else {
                            $(this).find('td:nth-child('+NTH+')').removeClass('green');
                        }
                    });
                };
            }
            if(typeof UpdateWatchList==='function'){
                UpdateWatchList=function(){
                    var tags='';
                    $.each(watchList,function(i,v){tags+='&tag='+v;});
                    if(tags.length>0){
                        $.getJSON('/tag/read?usepid=1&formated=1&old=1&prio=1&json=1'+tags,function(data){
                            $.each(data,function(key,value){
                                $('#'+key+' td:nth-child('+NTH+')').html(value);
                            });
                        }).always(function(){setTimeout(UpdateWatchList,1000);});
                    } else {
                        setTimeout(UpdateWatchList,1000);
                    }
                    if(typeof HasUnsavedChanges==='function') HasUnsavedChanges();
                };
            }
        })();`;
        document.head.appendChild(sc);sc.remove();
    }

    // ============================================================
    // SEARCH & REPLACE
    // ============================================================
    function searchReplaceModal() {
        const m = modal(`
<h3><i class="fa fa-exchange"></i> Sök och ersätt</h3>
<div style="display:flex;gap:4px;margin-bottom:10px;">
    <button class="sr-tab active" data-val="name" style="flex:1;padding:5px 8px;font-size:11px;font-weight:600;border:1px solid #ccc;border-radius:4px;cursor:pointer;background:#5b6abf;color:#fff;">Namn</button>
    <button class="sr-tab" data-val="description" style="flex:1;padding:5px 8px;font-size:11px;font-weight:600;border:1px solid #ccc;border-radius:4px;cursor:pointer;background:#fff;color:#333;">Beskrivning</button>
    <button class="sr-tab" data-val="prefix" style="flex:1;padding:5px 8px;font-size:11px;font-weight:600;border:1px solid #ccc;border-radius:4px;cursor:pointer;background:#fff;color:#333;">Prefix (Namn)</button>
</div>
<input type="hidden" name="sr-field" value="name">
<div id="sr-inputs">
    <div id="sr-search-row"><label style="font-size:11px;font-weight:600;">Sök</label><input id="sr-search" placeholder="Text att söka efter"></div>
    <label style="font-size:11px;font-weight:600;margin-top:6px;">Ersätt med / Prefix</label>
    <input id="sr-replace" placeholder="Ersättningstext">
</div>
<div style="margin-top:10px;"><button class="sec" id="sr-preview" style="padding:5px 14px;"><i class="fa fa-eye"></i> Förhandsgranska</button></div>
<div id="sr-results" style="max-height:250px;overflow-y:auto;margin-top:10px;font-size:11px;"></div>
<div class="bt" style="margin-top:10px;">
    <button class="bx" id="sr-cancel">Avbryt</button>
    <button class="bok" id="sr-apply" disabled><i class="fa fa-check"></i> Applicera</button>
</div>`);

        const fieldHidden = m.querySelector('input[name="sr-field"]');
        const tabs = m.querySelectorAll('.sr-tab');
        const searchRow = m.querySelector('#sr-search-row');
        const searchInput = m.querySelector('#sr-search');
        const replaceInput = m.querySelector('#sr-replace');
        const resultsDiv = m.querySelector('#sr-results');
        const applyBtn = m.querySelector('#sr-apply');
        let previewData = [];

        tabs.forEach(tab => tab.addEventListener('click', () => {
            tabs.forEach(t => { t.style.background = '#fff'; t.style.color = '#333'; t.classList.remove('active'); });
            tab.style.background = '#5b6abf'; tab.style.color = '#fff'; tab.classList.add('active');
            fieldHidden.value = tab.dataset.val;
            const isPrefix = tab.dataset.val === 'prefix';
            searchRow.style.display = isPrefix ? 'none' : '';
            replaceInput.placeholder = isPrefix ? 'Prefix att lägga till' : 'Ersättningstext';
            resultsDiv.innerHTML = ''; applyBtn.disabled = true; previewData = [];
        }));

        m.querySelector('#sr-cancel').addEventListener('click', () => m.remove());

        m.querySelector('#sr-preview').addEventListener('click', async () => {
            const field = fieldHidden.value;
            const search = searchInput.value;
            const replace = replaceInput.value;
            if (field !== 'prefix' && !search) { toastErr('Ange söktext'); return; }
            if (!replace && field === 'prefix') { toastErr('Ange prefix'); return; }

            // Fetch all tags from API
            const sid = (location.search.match(/sid=([^&#]+)/) || [])[1] || '';
            const btn = m.querySelector('#sr-preview');
            btn.disabled = true; btn.innerHTML = '<i class="fa fa-spinner fa-spin"></i> Söker...';

            try {
                const r = await fetch('/tag/GetTagList?sid=' + sid + '&draw=1&limit=9999&offset=0&sortcol=0&sortdir=asc&search=');
                const json = await r.json();
                const decode = h => { const d = document.createElement('div'); d.innerHTML = h; return d.textContent; };
                previewData = [];
                const colIdx = field === 'description' ? '10' : '0';
                const formField = field === 'description' ? 'description' : 'name';

                for (const row of (json.data || [])) {
                    const tagName = row['0'] || '';
                    const curVal = field === 'description' ? decode(row['10'] || '') : tagName;
                    let newVal;
                    if (field === 'prefix') {
                        newVal = replace + curVal;
                    } else {
                        if (!curVal.includes(search)) continue;
                        newVal = curVal.split(search).join(replace);
                    }
                    if (newVal === curVal) continue;
                    previewData.push({ tagName, tagId: row.DT_RowId || encodeTag(tagName), curVal, newVal, formField });
                }

                if (!previewData.length) {
                    resultsDiv.innerHTML = '<div style="color:#999;padding:8px;">Inga träffar</div>';
                    applyBtn.disabled = true;
                } else {
                    let html = `<div style="font-weight:600;margin-bottom:4px;">${previewData.length} ändringar:</div><table style="width:100%;border-collapse:collapse;">`;
                    html += '<tr style="font-size:10px;color:#999;border-bottom:1px solid #ddd;"><th style="text-align:left;padding:2px 4px;">Tagg</th><th style="text-align:left;padding:2px 4px;">Nuvarande</th><th style="text-align:left;padding:2px 4px;">Nytt</th></tr>';
                    for (const p of previewData.slice(0, 100)) {
                        html += `<tr style="border-bottom:1px solid #eee;">
                            <td style="padding:2px 4px;font-size:10px;opacity:.6;">${escHtml(p.tagName)}</td>
                            <td style="padding:2px 4px;color:#999;">${escHtml(p.curVal)}</td>
                            <td style="padding:2px 4px;color:#2e7d32;font-weight:600;">${escHtml(p.newVal)}</td>
                        </tr>`;
                    }
                    if (previewData.length > 100) html += `<tr><td colspan="3" style="padding:4px;color:#999;font-size:10px;">...och ${previewData.length - 100} fler</td></tr>`;
                    html += '</table>';
                    resultsDiv.innerHTML = html;
                    applyBtn.disabled = false;
                }
            } catch (e) { toastErr('Kunde inte hämta taggar: ' + e.message); }
            btn.disabled = false; btn.innerHTML = '<i class="fa fa-eye"></i> Förhandsgranska';
        });

        applyBtn.addEventListener('click', async () => {
            if (!previewData.length) return;
            const total = previewData.length;
            applyBtn.disabled = true; applyBtn.innerHTML = `<i class="fa fa-spinner fa-spin"></i> 0/${total}...`;
            let ok = 0, fail = 0;
            for (let i = 0; i < previewData.length; i++) {
                const p = previewData[i];
                applyBtn.innerHTML = `<i class="fa fa-spinner fa-spin"></i> ${i + 1}/${total}...`;
                try {
                    await fetchFormAndSave(p.tagName, fd => fd.set(p.formField, p.newVal));
                    ok++;
                } catch (e) { console.warn(CFG.logPrefix, 'S&R failed', p.tagName, e); fail++; }
            }
            m.remove();
            toastOk(`${ok} uppdaterade` + (fail ? `, ${fail} misslyckades` : ''));
            // Refresh the DataTables view
            const sc = document.createElement('script');
            sc.textContent = `try{$('#tagtable').dataTable()._fnAjaxUpdate();}catch(e){}`;
            document.head.appendChild(sc); sc.remove();
        });
    }

    // ============================================================
    // LIVE MONITOR — Commissioning split-screen tool
    // ============================================================
    // Sensor type categories for monitor grouping
    const SENSOR_TYPES = {
        GT: 'Temperatur', GP: 'Tryck', GF: 'Flöde', GM: 'Fukt', GQ: 'CO₂',
    };

    function sensorCategory(device) {
        const m = device.match(/^([A-Z]+)/i);
        return (m && SENSOR_TYPES[m[1].toUpperCase()]) || 'Övrigt';
    }

    function groupTagsBySensor(tags) {
        // Tags are pre-filtered to end with _PV, _FAULT, or _V, so always split on last underscore
        const groups = {};
        for (const t of tags) {
            const sep = t.name.lastIndexOf('_');
            const key = sep > 0 ? t.name.substring(0, sep) : t.name;
            const suffix = sep > 0 ? t.name.substring(sep + 1).toUpperCase() : 'VAL';
            const device = key.split('_').pop();
            const noSpotlight = /^(CLOCK|DATE|TIME|TID|DATUM|DAT|KL|UR|KLOCKA)$/i.test(device);
            if (!groups[key]) groups[key] = { key, device, category: sensorCategory(device), noSpotlight, tags: [] };
            groups[key].tags.push({ ...t, suffix });
        }
        // Sort groups by category then device
        const sorted = Object.values(groups);
        sorted.sort((a, b) => a.category.localeCompare(b.category) || a.device.localeCompare(b.device));
        return sorted;
    }

    function _buildPrefixSuggestions(rows) {
        const decode = h => { const d = document.createElement('div'); d.innerHTML = h; return d.textContent; };
        const counts = {};
        const nextSegs = {};      // prefix → Set of all unique next-level segments
        const sensorSegs = {};    // prefix → Set of next-level segments that have a _PV descendant
        const sourceSets = {};    // prefix → Set of unique IO-Enhet source names
        rows.forEach(row => {
            const name = String(row['0'] || '');
            if (!name) return;
            const src = decode(String(row['1'] || '')).trim();
            const parts = name.split('_');
            const isPv = /_(PV)$/i.test(name);
            for (let n = 1; n <= Math.min(4, parts.length - 1); n++) {
                const prefix = parts.slice(0, n).join('_');
                counts[prefix] = (counts[prefix] || 0) + 1;
                if (src) {
                    if (!sourceSets[prefix]) sourceSets[prefix] = new Set();
                    sourceSets[prefix].add(src);
                }
                if (parts.length > n) {
                    if (!nextSegs[prefix]) nextSegs[prefix] = new Set();
                    nextSegs[prefix].add(parts[n]);
                    if (isPv) {
                        if (!sensorSegs[prefix]) sensorSegs[prefix] = new Set();
                        sensorSegs[prefix].add(parts[n]); // if this is 'PV', prefix is a direct sensor (1 level from PV)
                    }
                }
            }
        });
        const candidates = Object.entries(counts).filter(([, c]) => c >= 2).sort((a, b) => b[1] - a[1]);
        const result = [];
        for (const [prefix, count] of candidates) {
            const hasMoreSpecific = candidates.some(([p, c]) => p !== prefix && p.startsWith(prefix + '_') && c === count);
            const srcs = sourceSets[prefix];
            if (!hasMoreSpecific) result.push({
                prefix, count,
                sensors: sensorSegs[prefix]?.size || 0,
                direct: sensorSegs[prefix]?.has('PV') ?? false,
                source: srcs?.size === 1 ? [...srcs][0] : null, // only show when unambiguous
            });
            if (result.length >= 20) break;
        }
        return result.sort((a, b) => a.prefix.localeCompare(b.prefix));
    }

    function monitorPrompt() {
        const prefs = JSON.parse(GM_getValue(CFG.storageKeys.monitor, '{}'));
        const m = modal(`
<h3><i class="fa fa-television"></i> Övervaka system</h3>
<label>Systemprefix (t.ex. VS21, LB01, AS01_KVP)</label>
<input id="mon-prefix" value="${escHtml(prefs.prefix || '')}" placeholder="VS21">
<p style="font-size:11px;color:#888;margin:6px 0 0;">Taggar grupperas per sensor. Värden utanför rå-skalning markeras.</p>
<div style="display:flex;justify-content:space-between;align-items:center;margin:14px 0 6px;">
<h5 style="margin:0;font-size:11px;color:#888;font-weight:700;text-transform:uppercase;letter-spacing:.4px;">Systemförslag</h5>
<label style="display:flex;align-items:center;gap:6px;font-size:11px;color:#888;cursor:pointer;">
<label class="larm-tog" style="margin:0;"><input type="checkbox" id="mon-show-empty"><span class="sl"></span></label>
Visa allt
</label>
</div>
<div id="mon-sugg" style="display:flex;flex-direction:column;gap:4px;"><span style="font-size:11px;color:#aaa;padding:4px 0;"><i class="fa fa-spinner fa-spin"></i> Hämtar...</span></div>
<div class="bt">
<button class="bx" id="pc">Avbryt</button>
<button class="bok" id="ps"><i class="fa fa-play"></i> Starta</button>
</div>`);
        const suggEl = m.querySelector('#mon-sugg');
        const inp = m.querySelector('#mon-prefix');
        // Fetch ALL tags (no sid filter) to build suggestions from complete tag list
        fetch('/tag/GetTagList?draw=1&limit=9999&offset=0&sortcol=0&sortdir=asc&search=')
            .then(r => r.json())
            .then(json => {
                const allSuggestions = _buildPrefixSuggestions(json.data || []);
                const showEmptyChk = m.querySelector('#mon-show-empty');
                const renderSuggestions = () => {
                const showEmpty = showEmptyChk.checked;
                const suggestions = showEmpty ? allSuggestions : (() => {
                    return allSuggestions.filter(s => {
                        // Always show structural containers (parent of other suggestions)
                        const isParent = allSuggestions.some(o => o.prefix.startsWith(s.prefix + '_'));
                        if (isParent) return true;
                        // Hide if no launchable sensors
                        if (s.sensors === 0) return false;
                        // Hide only if this IS a direct sensor (PV tag 1 level deep)
                        // and a parent sub-system with sensors already covers it
                        if (s.direct) {
                            return !allSuggestions.some(o => s.prefix.startsWith(o.prefix + '_') && o.sensors > 0 && !o.direct);
                        }
                        return true;
                    });
                })();
                suggEl.innerHTML = '';
                if (!suggestions.length) { suggEl.innerHTML = '<span style="font-size:11px;color:#aaa;">Inga förslag</span>'; return; }
                // Group by root (first segment) for visual spacing
                let lastRoot = null;
                suggestions.forEach(({ prefix, count, sensors, source }) => {
                    const depth = prefix.split('_').length; // 1=HUS03, 2=HUS03_AS01, 3=HUS03_AS01_KB21 …
                    const root = prefix.split('_')[0];
                    // Spacing gap between different root groups
                    if (root !== lastRoot) {
                        if (lastRoot !== null) {
                            const gap = document.createElement('div');
                            gap.style.cssText = 'height:6px;';
                            suggEl.appendChild(gap);
                        }
                        lastRoot = root;
                    }
                    const btn = document.createElement('button');
                    const indent = Math.max(0, depth - 2) * 16; // depth 1-2: 0px, 3: 16px, 4: 32px
                    const tagCount = '<span style="opacity:.55;font-weight:400;font-size:11px;flex-shrink:0;margin-left:6px;">' + count + ' taggar</span>';
                    const srcPill = source ? '<span style="background:rgba(0,0,0,0.12);border-radius:10px;padding:1px 7px;font-size:10px;font-weight:500;margin-right:4px;opacity:.8;">' + escHtml(source) + '</span>' : '';
                    if (depth === 1) {
                        // Section label — uppercase header, no button look
                        btn.style.cssText = 'display:flex;justify-content:space-between;align-items:center;width:100%;text-align:left;background:none;border:none;border-radius:4px;padding:5px 14px 3px;cursor:pointer;font-size:10px;font-weight:700;letter-spacing:0.07em;text-transform:uppercase;color:#9ca3af;';
                        btn.onmouseover = () => btn.style.color = '#6b7280';
                        btn.onmouseout  = () => btn.style.color = '#9ca3af';
                        btn.innerHTML = '<span>' + escHtml(prefix) + '</span>' + tagCount;
                    } else if (depth === 2) {
                        // Parent system — soft filled, clearly a container not a leaf
                        btn.style.cssText = 'display:flex;justify-content:space-between;align-items:center;width:100%;text-align:left;background:#eef0f8;color:#4a5580;border:1px solid #d0d4e8;border-radius:4px;padding:5px 14px;cursor:pointer;font-size:12px;font-weight:600;';
                        btn.onmouseover = () => btn.style.background = '#e2e6f3';
                        btn.onmouseout  = () => btn.style.background = '#eef0f8';
                        btn.innerHTML = '<span>' + escHtml(prefix) + '</span><span style="display:flex;align-items:center;flex-shrink:0;">' + srcPill + tagCount + '</span>';
                    } else if (depth === 3) {
                        // Subsystem — primary action button
                        btn.className = 'bok';
                        btn.style.cssText = 'display:flex;justify-content:space-between;align-items:center;text-align:left;margin-left:' + indent + 'px;width:calc(100% - ' + indent + 'px);';
                        const last = escHtml(prefix.split('_').slice(-1)[0]);
                        const full = escHtml(prefix);
                        const pill3 = sensors > 0 ? '<span style="background:rgba(255,255,255,0.25);border-radius:10px;padding:1px 7px;font-size:10px;font-weight:600;margin-right:4px;">' + sensors + ' sensorer</span>' : '';
                        btn.innerHTML = '<span>' + last + '<span style="font-weight:400;opacity:.6;font-size:10px;margin-left:6px;">' + full + '</span></span><span style="display:flex;align-items:center;flex-shrink:0;">' + srcPill + pill3 + tagCount + '</span>';
                    } else {
                        // Depth 4+ — deeper component, darker with left-border accent
                        btn.style.cssText = 'display:flex;justify-content:space-between;align-items:center;text-align:left;margin-left:' + indent + 'px;width:calc(100% - ' + indent + 'px);background:#3d4f96;color:#fff;border:none;border-left:3px solid #8899d8;border-radius:4px;padding:6px 14px;cursor:pointer;font-size:12px;font-weight:600;';
                        btn.onmouseover = () => btn.style.background = '#2e3d7a';
                        btn.onmouseout  = () => btn.style.background = '#3d4f96';
                        const last = escHtml(prefix.split('_').slice(-1)[0]);
                        const full = escHtml(prefix);
                        btn.innerHTML = '<span>' + last + '<span style="font-weight:400;opacity:.6;font-size:10px;margin-left:6px;">' + full + '</span></span><span style="display:flex;align-items:center;flex-shrink:0;">' + srcPill + tagCount + '</span>';
                    }
                    btn.addEventListener('click', () => { inp.value = prefix; m.querySelector('#ps').click(); });
                    suggEl.appendChild(btn);
                });
                }; // end renderSuggestions
                showEmptyChk.addEventListener('change', renderSuggestions);
                renderSuggestions();
            })
            .catch(() => { suggEl.innerHTML = '<span style="font-size:11px;color:#aaa;">Kunde inte hämta förslag</span>'; });
        m.querySelector('#pc').addEventListener('click', () => m.remove());
        m.querySelector('#ps').addEventListener('click', async () => {
            const inp = m.querySelector('#mon-prefix');
            const prefix = inp.value.trim() || inp.placeholder;
            if (!prefix) { toastErr('Ange ett prefix'); return; }
            const btn = m.querySelector('#ps');
            btn.disabled = true; btn.innerHTML = '<i class="fa fa-spinner fa-spin"></i> Laddar...';
            try {
                // Fetch ALL tags from server API (no sid filter — monitor works across all sources)
                const r = await fetch('/tag/GetTagList?draw=1&limit=9999&offset=0&sortcol=0&sortdir=asc&search=');
                const json = await r.json();
                const decode = h => { const d = document.createElement('div'); d.innerHTML = h; return d.textContent; };
                const tags = [];
                for (const row of (json.data || [])) {
                    const name = row['0'] || '';
                    if (name.toUpperCase().includes(prefix.toUpperCase()) && /_(PV|FAULT|V)$/i.test(name)) {
                        tags.push({
                            id: row.DT_RowId || encodeTag(name), name,
                            unit: decode(row['8'] || ''),
                            rawmin: parseFloat(row['4']) || 0, rawmax: parseFloat(row['5']) || 0,
                            engmin: parseFloat(row['6']) || 0, engmax: parseFloat(row['7']) || 0,
                            dtype: row['3'] || '',
                        });
                    }
                }
                const resetBtn = () => { btn.disabled = false; btn.innerHTML = '<i class="fa fa-play"></i> Starta'; };
                if (!tags.length) { toastErr('Inga taggar matchar "' + prefix + '"'); resetBtn(); return; }
                const grouped = groupTagsBySensor(tags);
                const validTags = [];
                for (const g of grouped) {
                    if (g.tags.some(t => t.suffix === 'PV' || t.suffix === 'V')) g.tags.forEach(t => validTags.push(t));
                }
                if (!validTags.length) { toastErr('Inga sensorer med PV eller V hittades'); resetBtn(); return; }
                if (validTags.length > 200) { toastErr(`För många taggar (${validTags.length}). Använd ett mer specifikt prefix.`); resetBtn(); return; }
                m.remove();
                GM_setValue(CFG.storageKeys.monitor, JSON.stringify({ prefix }));
                new LiveMonitor(validTags, prefix).open();
            } catch (e) {
                toastErr('Kunde inte hämta taggar: ' + e.message);
                btn.disabled = false; btn.innerHTML = '<i class="fa fa-play"></i> Starta';
            }
        });
        inp?.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); m.querySelector('#ps')?.click(); } });
        setTimeout(() => inp?.focus(), 50);
    }

    class LiveMonitor {
        constructor(tags, prefix, labelMap, controls, faultDescMap) {
            this.prefix = prefix;
            this.labelMap = labelMap || {};
            this.allTags = tags;
            this.controls = controls || [];
            this.faultDescMap = faultDescMap || {};
            this.ctrlLocks = {}; // prefix → timestamp until which poll updates are suppressed
            this.spotOOR = true; // spotlight on out-of-range values (togglable)
            // Raw tag names for control polling (_M, _OPM/_MCMD)
            this.ctrlTagNames = [];
            for (const c of this.controls) {
                this.ctrlTagNames.push(c.prefix + '_M');
                if (c.type === 'valve') this.ctrlTagNames.push(c.prefix + '_OPM');
                if (c.type === 'pump') {
                    this.ctrlTagNames.push(c.prefix + '_MCMD');
                    this.ctrlTagNames.push(c.prefix + '_V');
                    this.ctrlTagNames.push(c.prefix + '_FAULT');
                }
            }
            this.groups = groupTagsBySensor(tags);
            this.values = {};
            this.prevValues = {};
            this.verified = new Set();
            this.spotCooldown = {};
            this.closed = false;
            this.timeline = {}; // key → [{time, type, val}]
            this.pollTimer = null;
            this.clockTimer = null;
            this.dark = false;
            this.sound = false;
            this.el = null;
            this.audioCtx = null;
        }

        open() {
            this.injectCSS();
            this.el = document.createElement('div');
            this.el.id = 'inu-monitor';
            this.el.className = 'inu-mon light';
            const total = this.allTags.length;
            const sensors = this.groups.length;
            this.el.innerHTML = `
<div class="inu-mon-hdr">
    <span class="inu-mon-title">${INU_LOGO_SVG} <span>WP+ Monitor — ${sensors} sensorer (${total} taggar)</span></span>
    <span class="inu-mon-status" id="inu-mon-status">0 / ${sensors} verifierade</span>
    <span class="inu-mon-prefix">${this.esc(this.prefix)}</span>
    <span class="inu-mon-clock"></span>
    <span class="inu-mon-controls">
        <button title="Spotlight vid utanför skalning (O)" class="inu-mon-btn inu-mon-btn-active" data-act="oor"><i class="fa fa-exclamation-triangle"></i></button>
        <button title="Ljud på/av (S)" class="inu-mon-btn" data-act="sound"><i class="fa fa-volume-off"></i></button>
        <button title="Mörkt/ljust (D)" class="inu-mon-btn" data-act="theme"><i class="fa fa-adjust"></i></button>
        <button title="Fullskärm (F)" class="inu-mon-btn" data-act="full"><i class="fa fa-expand"></i></button>
        <button title="Stäng (Esc)" class="inu-mon-btn" data-act="close"><i class="fa fa-times"></i></button>
    </span>
</div>
<div class="inu-mon-body">
    <div class="inu-mon-main">
        <div class="inu-mon-spot" id="inu-mon-spot">
            <div class="inu-mon-spot-empty"><i class="fa fa-plug"></i> Väntar på förändring...<br><span style="font-size:11px;opacity:.5;">Sensorer utanför rå-skalning markeras röda · Klicka kort för att verifiera</span></div>
        </div>
        <div class="inu-mon-grid-wrap">
            <div class="inu-mon-grid" id="inu-mon-grid"></div>
            <div class="inu-mon-scroll-hint" id="inu-mon-scroll-hint" style="display:none;"><i class="fa fa-chevron-down"></i> <span id="inu-mon-scroll-count"></span></div>
        </div>
    </div>
    ${this.controls.length ? '<div class="inu-mon-ctrl-bar" id="inu-mon-ctrl-bar"></div>' : ''}
</div>`;
            document.body.appendChild(this.el);
            this.buildGrid();
            if (this.controls.length) this.renderCtrlBar();
            this.bindKeys();
            this.bindButtons();
            this.updateClock();
            this.clockTimer = setInterval(() => this.updateClock(), 1000);
            this.poll();
        }

        close() {
            this.closed = true;
            clearTimeout(this.pollTimer);
            clearInterval(this.clockTimer);
            if (this.el) this.el.remove();
            document.removeEventListener('keydown', this._keyHandler);
            if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
        }

        buildGrid() {
            const grid = this.el.querySelector('#inu-mon-grid');
            let lastCat = '', catGroup = null;
            for (const g of this.groups) {
                if (g.category !== lastCat) {
                    lastCat = g.category;
                    catGroup = document.createElement('div');
                    catGroup.className = 'inu-mon-catgrp';
                    const hdr = document.createElement('div');
                    hdr.className = 'inu-mon-catgrp-hdr';
                    hdr.textContent = g.category;
                    catGroup.appendChild(hdr);
                    grid.appendChild(catGroup);
                }
                const card = document.createElement('div');
                card.className = 'inu-mon-card';
                card.dataset.sensor = g.key;
                const pv = g.tags.find(t => t.suffix === 'PV') || g.tags[0];
                const fault = g.tags.find(t => t.suffix === 'FAULT');
                card.innerHTML = `
<div class="inu-mon-card-hdr">${this.displayName(g.key)}</div>
<div class="inu-mon-card-pv" data-tag="${pv.id}">--</div>
${fault ? `<div class="inu-mon-card-fault" data-tag="${fault.id}"><i class="fa fa-exclamation-circle"></i> <span>OK</span></div>` : ''}
<div class="inu-mon-card-check"><i class="fa fa-check"></i></div>`;
                card.addEventListener('click', () => this.toggleVerified(g.key, card));
                catGroup.appendChild(card);
            }

            // Scroll indicator — shows how many cards are below the fold
            const hint = this.el.querySelector('#inu-mon-scroll-hint');
            const countEl = this.el.querySelector('#inu-mon-scroll-count');
            const updateHint = () => {
                const atBottom = grid.scrollHeight - grid.scrollTop <= grid.clientHeight + 20;
                if (atBottom) { hint.style.display = 'none'; return; }
                const cards = [...grid.querySelectorAll('.inu-mon-card')];
                const hidden = cards.filter(c => c.getBoundingClientRect().top >= grid.getBoundingClientRect().bottom - 10).length;
                if (hidden > 0) {
                    countEl.textContent = hidden + ' sensor' + (hidden > 1 ? 'er' : '') + ' nedanför';
                    hint.style.display = '';
                } else {
                    hint.style.display = 'none';
                }
            };
            grid.addEventListener('scroll', updateHint);
            setTimeout(updateHint, 100);
        }

        toggleVerified(key, card) {
            if (this.verified.has(key)) {
                this.verified.delete(key);
                card.classList.remove('verified');
            } else {
                this.verified.add(key);
                card.classList.add('verified');
            }
            this.updateStatus();
        }

        updateStatus() {
            const el = this.el.querySelector('#inu-mon-status');
            if (el) el.textContent = `${this.verified.size} / ${this.groups.length} verifierade`;
        }

        ctrlIcon(prefix, defaultIcon) {
            const code = prefix.split('_').pop().replace(/\d.*$/, '').toUpperCase();
            if (/^(FF|TF|KF|AF|FT)/.test(code))          return 'fa-refresh';       // fans
            if (/^(ST|SD|KS|LS|BS)/.test(code))           return 'fa-align-justify'; // dampers
            if (/^(STV|SV|VV|TV|KV|HV|MV|RV)/.test(code)) return 'fa-tint';         // valves
            if (/^(PV|PK|PP|CP|KP|VP)/.test(code))        return 'fa-circle-o-notch'; // pumps
            if (/^P$/.test(code))                          return 'fa-circle-o-notch'; // generic pump
            return defaultIcon;
        }

        renderCtrlBar() {
            const bar = this.el.querySelector('#inu-mon-ctrl-bar');
            if (!bar) return;
            bar.innerHTML = this.controls.map(c => {
                const p    = this.esc(c.prefix);
                const icon = this.ctrlIcon(c.prefix, c.type === 'valve' ? 'fa-sliders' : 'fa-circle-o-notch');
                if (c.type === 'valve') {
                    return `<div class="inu-mon-ctrl-device" data-prefix="${p}" data-type="valve">
  <div class="inu-mon-ctrl-name"><i class="fa ${icon}"></i> ${this.esc(c.label)}</div>
  <button class="inu-mon-ctrl-btn" data-ctrl="valve-auto"   data-prefix="${p}">Auto</button>
  <button class="inu-mon-ctrl-btn" data-ctrl="valve-manual" data-prefix="${p}">Manuell</button>
  <div class="inu-mon-ctrl-opm" data-prefix="${p}" style="display:none">
    <div class="inu-mon-ctrl-slider-wrap">
      <input type="range" class="inu-mon-ctrl-slider" min="0" max="100" step="1" value="0" data-prefix="${p}">
    </div>
    <span class="inu-mon-ctrl-val">0 %</span>
  </div>
</div>`;
                } else {
                    return `<div class="inu-mon-ctrl-device" data-prefix="${p}" data-type="pump">
  <div class="inu-mon-ctrl-name"><i class="fa ${icon}"></i> ${this.esc(c.label)}</div>
  <div class="inu-mon-ctrl-running" data-prefix="${p}"><i class="fa fa-circle"></i> <span>–</span></div>
  <div class="inu-mon-ctrl-fault-ind" data-prefix="${p}" style="display:none"><i class="fa fa-exclamation-circle"></i> ${this.esc(c.faultDesc || 'Driftfel')}</div>
  <button class="inu-mon-ctrl-btn" data-ctrl="pump-auto" data-prefix="${p}">Auto</button>
  <button class="inu-mon-ctrl-btn" data-ctrl="pump-off"  data-prefix="${p}">Från</button>
  <button class="inu-mon-ctrl-btn" data-ctrl="pump-on"   data-prefix="${p}">Till</button>
</div>`;
                }
            }).join('');

            // Optimistic button clicks
            bar.addEventListener('click', e => {
                const btn = e.target.closest('[data-ctrl]');
                if (!btn) return;
                const prefix = btn.dataset.prefix;
                const c = this.controls.find(x => x.prefix === prefix);
                if (!c) return;
                const action = btn.dataset.ctrl;
                const bodyMap = {
                    'valve-auto':   `${prefix}_003select=M=0&log=Fr%C3%A5n`,
                    'valve-manual': `${prefix}_003select=M=1&log=Till`,
                    'pump-auto':    `${prefix}_010select=M=0&log=Auto`,
                    'pump-off':     `${prefix}_010select=M=1,MCMD=0&log=Fr%C3%A5n`,
                    'pump-on':      `${prefix}_010select=M=1,MCMD=1&log=Till`,
                };
                const body = bodyMap[action];
                if (!body) return;
                // Snapshot for revert
                const snapshot = { ...this.values };
                // Apply visuals immediately and suppress poll overwrites for 2.5s
                this.applyCtrlVisual(c, action);
                this.ctrlLocks[prefix] = Date.now() + 2500;
                // Write in background, revert on failure
                this.sendCtrlMode(c, body).catch(e => {
                    delete this.ctrlLocks[prefix];
                    this.updateCtrlBar(snapshot);
                    toastErr('Handkörning misslyckades: ' + e.message);
                });
            });

            // Slider: live display on input, throttled write every 500ms, also write on release
            const sliderTimers = {};
            const writeSlider = (slider) => {
                const prefix = slider.dataset.prefix;
                const c = this.controls.find(x => x.prefix === prefix);
                if (!c) return;
                const newVal = parseFloat(slider.value).toFixed(1) + ' %';
                const curOpm = this.values[prefix + '_OPM'];
                const orgVal = curOpm !== undefined ? parseFloat(curOpm).toFixed(1) + ' %' : newVal;
                const p = new URLSearchParams();
                p.set('pageid', c.pageid);
                p.set('poid', c.poid);
                p.set(prefix + '_OPM_value', newVal);
                p.set(prefix + '_OPM_org', orgVal);
                this.writeCtrl('', p.toString());
            };
            bar.addEventListener('input', e => {
                if (!e.target.matches('.inu-mon-ctrl-slider')) return;
                const valEl = e.target.closest('.inu-mon-ctrl-opm')?.querySelector('.inu-mon-ctrl-val');
                if (valEl) valEl.textContent = e.target.value + ' %';
                const prefix = e.target.dataset.prefix;
                clearTimeout(sliderTimers[prefix]);
                this.ctrlLocks[prefix] = Date.now() + 2500;
                sliderTimers[prefix] = setTimeout(() => writeSlider(e.target), 500);
            });
            bar.addEventListener('change', e => {
                if (!e.target.matches('.inu-mon-ctrl-slider')) return;
                clearTimeout(sliderTimers[e.target.dataset.prefix]);
                writeSlider(e.target);
            });
        }

        applyCtrlVisual(c, action) {
            const bar = this.el.querySelector('#inu-mon-ctrl-bar');
            if (!bar) return;
            const prefix = c.prefix;
            if (c.type === 'valve') {
                const isManual = action === 'valve-manual';
                bar.querySelector(`[data-ctrl="valve-auto"][data-prefix="${prefix}"]`)?.classList.toggle('active', !isManual);
                bar.querySelector(`[data-ctrl="valve-manual"][data-prefix="${prefix}"]`)?.classList.toggle('active', isManual);
                const opmDiv = bar.querySelector(`.inu-mon-ctrl-opm[data-prefix="${prefix}"]`);
                if (opmDiv) opmDiv.style.display = isManual ? 'flex' : 'none';
            } else {
                const isAuto = action === 'pump-auto';
                const isOff  = action === 'pump-off';
                const isOn   = action === 'pump-on';
                bar.querySelector(`[data-ctrl="pump-auto"][data-prefix="${prefix}"]`)?.classList.toggle('active', isAuto);
                bar.querySelector(`[data-ctrl="pump-off"][data-prefix="${prefix}"]`)?.classList.toggle('active', isOff);
                bar.querySelector(`[data-ctrl="pump-on"][data-prefix="${prefix}"]`)?.classList.toggle('active', isOn);
                bar.querySelector(`[data-ctrl="pump-off"][data-prefix="${prefix}"]`)?.classList.toggle('pump-off-active', isOff);
                bar.querySelector(`[data-ctrl="pump-on"][data-prefix="${prefix}"]`)?.classList.toggle('pump-on-active', isOn);
            }
        }

        updateCtrlBar(values) {
            const bar = this.el.querySelector('#inu-mon-ctrl-bar');
            if (!bar) return;
            const now = Date.now();
            for (const c of this.controls) {
                if (this.ctrlLocks[c.prefix] && now < this.ctrlLocks[c.prefix]) continue;
                delete this.ctrlLocks[c.prefix];
                const isManual = (values[c.prefix + '_M'] ?? 0) == 1;
                if (c.type === 'valve') {
                    bar.querySelector(`[data-ctrl="valve-auto"][data-prefix="${c.prefix}"]`)?.classList.toggle('active', !isManual);
                    bar.querySelector(`[data-ctrl="valve-manual"][data-prefix="${c.prefix}"]`)?.classList.toggle('active', isManual);
                    const opmDiv = bar.querySelector(`.inu-mon-ctrl-opm[data-prefix="${c.prefix}"]`);
                    if (opmDiv) {
                        opmDiv.style.display = isManual ? 'flex' : 'none';
                        const slider = opmDiv.querySelector('.inu-mon-ctrl-slider');
                        const opm = parseFloat(values[c.prefix + '_OPM'] ?? 0);
                        if (slider && document.activeElement !== slider) {
                            slider.value = opm;
                            const valEl = opmDiv.querySelector('.inu-mon-ctrl-val');
                            if (valEl) valEl.textContent = opm.toFixed(0) + ' %';
                        }
                    }
                } else {
                    const isOn = (values[c.prefix + '_MCMD'] ?? 0) == 1;
                    bar.querySelector(`[data-ctrl="pump-auto"][data-prefix="${c.prefix}"]`)?.classList.toggle('active', !isManual);
                    bar.querySelector(`[data-ctrl="pump-off"][data-prefix="${c.prefix}"]`)?.classList.toggle('active', isManual && !isOn);
                    bar.querySelector(`[data-ctrl="pump-on"][data-prefix="${c.prefix}"]`)?.classList.toggle('active', isManual && isOn);
                    bar.querySelector(`[data-ctrl="pump-off"][data-prefix="${c.prefix}"]`)?.classList.toggle('pump-off-active', isManual && !isOn);
                    bar.querySelector(`[data-ctrl="pump-on"][data-prefix="${c.prefix}"]`)?.classList.toggle('pump-on-active', isManual && isOn);
                    // Driftindikering
                    const isRunning = (values[c.prefix + '_V'] ?? 0) == 1;
                    const hasFault  = (values[c.prefix + '_FAULT'] ?? 0) == 1;
                    const runEl = bar.querySelector(`.inu-mon-ctrl-running[data-prefix="${c.prefix}"]`);
                    if (runEl) {
                        runEl.querySelector('span').textContent = isRunning ? 'Drift' : 'Stoppad';
                        runEl.classList.toggle('running', isRunning);
                    }
                    const faultEl = bar.querySelector(`.inu-mon-ctrl-fault-ind[data-prefix="${c.prefix}"]`);
                    if (faultEl) faultEl.style.display = hasFault ? '' : 'none';
                }
            }
        }

        async sendCtrlMode(ctrl, body) {
            await this.writeCtrl(`pageid=${ctrl.pageid}&poid=${ctrl.poid}`, body);
        }

        async writeCtrl(urlParams, body) {
            const url = '/page/UpdatePageObjectSettings' + (urlParams ? '?' + urlParams : '');
            const res = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body,
            });
            if (!res.ok) throw new Error('HTTP ' + res.status);
        }

        async poll() {
            if (this.closed) return;
            try {
                const params = this.allTags.map(t => 'tag=' + t.id).join('&');
                const ctrlParams = this.ctrlTagNames.map(n => 'tag=' + encodeURIComponent(n)).join('&');
                const allParams = params + (ctrlParams ? '&' + ctrlParams : '');
                const r = await fetch(CFG.endpoints.tagRead + '?usepid=1&formated=1&old=1&prio=1&json=1&' + allParams);
                const data = await r.json();
                if (this.closed) return;
                this.prevValues = { ...this.values };
                this.values = data;
                this.render();
                if (this.controls.length) this.updateCtrlBar(data);
            } catch (e) {
                if (this.closed) return;
                console.warn(CFG.logPrefix, 'Monitor poll error', e);
            }
            if (this.closed) return;
            this.pollTimer = setTimeout(() => this.poll(), CFG.pollMs);
        }

        parseNum(val) { return parseFloat(String(val).replace(/[()]/g, '')); }
        esc(val) { const d = document.createElement('div'); d.textContent = String(val ?? '--'); return d.innerHTML; }
        displayName(key) {
            if (this.labelMap[key]) return this.labelMap[key];
            const strip = this.prefix + '_';
            return key.startsWith(strip) ? key.slice(strip.length) : key.split('_').pop();
        }

        // Check if a value change exceeds the 10% deadband based on engineering range
        isSignificant(tag, oldVal, newVal) {
            if (tag.dtype === 'DIGITAL') return oldVal !== newVal;
            const o = this.parseNum(oldVal), n = this.parseNum(newVal);
            if (isNaN(o) || isNaN(n)) return oldVal !== newVal;
            const range = Math.abs(tag.engmax - tag.engmin);
            if (range === 0) return oldVal !== newVal;
            return Math.abs(n - o) >= range * 0.10;
        }

        // Check if value is outside engineering range
        isOutOfRange(tag, val) {
            if (tag.dtype === 'DIGITAL') return false;
            if (tag.engmin === 0 && tag.engmax === 0) return false;
            const num = this.parseNum(val);
            if (isNaN(num)) return false;
            return num < tag.engmin || num > tag.engmax;
        }

        render() {
            for (const g of this.groups) {
                const card = this.el.querySelector(`.inu-mon-card[data-sensor="${g.key}"]`);
                if (!card) continue;
                const pv = g.tags.find(t => t.suffix === 'PV') || g.tags[0];
                const fault = g.tags.find(t => t.suffix === 'FAULT');

                // Update PV display
                const pvVal = this.values[pv.id] ?? '--';
                const pvEl = card.querySelector(`.inu-mon-card-pv[data-tag="${pv.id}"]`);
                if (pvEl) {
                    if (pv.suffix === 'V') {
                        pvEl.textContent = pvVal === '1' || pvVal === '(1)' ? 'Till' : pvVal === '0' || pvVal === '(0)' ? 'Från' : pvVal;
                        pvEl.className = 'inu-mon-card-pv' + (pvVal === '1' || pvVal === '(1)' ? ' digital-on' : pvVal === '0' || pvVal === '(0)' ? ' digital-off' : '');
                    } else {
                        pvEl.textContent = pvVal;
                        pvEl.className = 'inu-mon-card-pv';
                    }
                }

                // Update fault display
                if (fault) {
                    const fVal = this.values[fault.id] ?? '--';
                    const fEl = card.querySelector(`.inu-mon-card-fault[data-tag="${fault.id}"] span`);
                    const faultOn = fVal !== '0' && fVal !== '(0)' && fVal !== '--';
                    if (fEl) { fEl.textContent = faultOn ? 'GIVARFEL' : 'OK'; card.querySelector('.inu-mon-card-fault')?.classList.toggle('active', faultOn); }
                }

                const hasOOR = this.isOutOfRange(pv, pvVal);

                card.classList.toggle('oor', hasOOR);

                // Spotlight: trigger on OOR (if enabled), fault activation, or 10% PV swing
                const pvPrev = this.prevValues[pv.id];
                const faultVal = fault ? (this.values[fault.id] ?? '--') : null;
                const faultPrev = fault ? (this.prevValues[fault.id] ?? '--') : null;
                const faultActive = faultVal && faultVal !== '0' && faultVal !== '(0)' && faultVal !== '--';
                const faultChanged = fault && faultPrev !== undefined && faultVal !== faultPrev && faultActive;
                const shouldSpot = (this.spotOOR && hasOOR) || faultChanged ||
                    (pvPrev !== undefined && this.isSignificant(pv, pvPrev, pvVal) && (this.spotOOR || !hasOOR));

                if (shouldSpot && !this.spotCooldown[g.key] && !g.noSpotlight) {
                    card.classList.add('changed');
                    setTimeout(() => card.classList.remove('changed'), 3000);
                    this.spotCooldown[g.key] = true;
                    setTimeout(() => delete this.spotCooldown[g.key], 5000);
                    try { this.promoteToSpotlight(g, pv, pvPrev, pvVal, faultVal, hasOOR); }
                    catch (e) { console.warn(CFG.logPrefix, 'Spotlight error for', g.key, e); }
                }

                // Live-update existing spotlight cards for this sensor
                const spotCard = this.el.querySelector(`.inu-mon-spot-card[data-sensor="${g.key}"]`);
                if (spotCard) {
                    const sv = spotCard.querySelector('.inu-mon-spot-val');
                    if (sv) {
                        const dispVal = pv.suffix === 'V'
                            ? (pvVal === '1' || pvVal === '(1)' ? 'Till' : pvVal === '0' || pvVal === '(0)' ? 'Från' : pvVal)
                            : pvVal;
                        sv.textContent = this.esc(dispVal);
                        this.fitText(sv, spotCard);
                    }
                    const sf = spotCard.querySelector('.inu-mon-spot-fault');
                    const faultCleared = sf && !faultActive;
                    if (faultCleared) { sf.remove(); this.addTimelineEvent(g.key, 'ok', 'OK'); }
                    if (!sf && faultActive) { const f = document.createElement('div'); f.className = 'inu-mon-spot-fault'; f.innerHTML = '<i class="fa fa-exclamation-circle"></i> GIVARFEL'; spotCard.querySelector('.inu-mon-spot-val').after(f); }
                    spotCard.classList.toggle('oor', hasOOR);
                    // Refresh timeline
                    const oldTl = spotCard.querySelector('.inu-tl');
                    const newTlHtml = this.buildTimelineHtml(g.key);
                    if (newTlHtml) {
                        if (oldTl) oldTl.outerHTML = newTlHtml;
                        else spotCard.insertAdjacentHTML('beforeend', newTlHtml);
                    }
                }
            }
        }

        addTimelineEvent(key, type, val) {
            if (!this.timeline[key]) this.timeline[key] = [];
            this.timeline[key].push({ time: new Date().toLocaleTimeString('sv-SE'), type, val });
            if (this.timeline[key].length > 10) this.timeline[key].shift();
        }

        buildTimelineHtml(key) {
            const events = this.timeline[key] || [];
            if (!events.length) return '';
            let html = '';
            for (let i = 0; i < events.length; i++) {
                if (i > 0) html += '<div class="inu-tl-arrow"><i class="fa fa-long-arrow-right"></i></div>';
                const ev = events[i];
                const cls = ev.type === 'ok' ? 'ok' : 'nok';
                const icon = ev.type === 'fault' ? 'fa-exclamation-circle' : ev.type === 'oor' ? 'fa-exclamation-triangle' : ev.type === 'ok' ? 'fa-check' : 'fa-arrow-right';
                html += `<div class="inu-tl-ev ${cls}"><div class="inu-tl-time">${ev.time}</div><div class="inu-tl-box"><i class="fa ${icon}"></i> ${this.esc(ev.val)}</div></div>`;
            }
            return `<div class="inu-tl">${html}</div>`;
        }

        clearOORSpotCards() {
            const spot = this.el.querySelector('#inu-mon-spot');
            if (!spot) return;
            spot.querySelectorAll('.inu-mon-spot-card.oor').forEach(card => {
                if (!card.querySelector('.inu-mon-spot-fault')) {
                    card.remove();
                } else {
                    card.classList.remove('oor');
                    card.querySelector('.inu-mon-spot-oor')?.remove();
                }
            });
            const remaining = spot.querySelectorAll('.inu-mon-spot-card');
            if (!remaining.length) {
                spot.innerHTML = '<div class="inu-mon-spot-empty"><i class="fa fa-plug"></i><br>Väntar på förändring...</div>';
            } else {
                remaining.forEach((c, i, a) => c.classList.toggle('secondary', i < a.length - 1));
                requestAnimationFrame(() => remaining.forEach(c => { const v = c.querySelector('.inu-mon-spot-val'); if (v && c.clientWidth > 0) this.fitText(v, c); }));
            }
        }

        promoteToSpotlight(group, pvTag, oldVal, newVal, faultVal, outOfRange) {
            const spot = this.el.querySelector('#inu-mon-spot');
            const faultOn = faultVal && faultVal !== '0' && faultVal !== '(0)' && faultVal !== '--';

            // Log timeline event
            if (faultOn) this.addTimelineEvent(group.key, 'fault', 'GIVARFEL');
            else if (outOfRange) this.addTimelineEvent(group.key, 'oor', newVal);
            else this.addTimelineEvent(group.key, 'swing', newVal);

            // Reuse existing card for same sensor
            const existingCard = spot.querySelector(`.inu-mon-spot-card[data-sensor="${group.key}"]`);
            if (existingCard) {
                existingCard.classList.add('flash');
                existingCard.classList.toggle('oor', outOfRange);
                setTimeout(() => existingCard.classList.remove('flash'), 3000);
                // Update content in live-update section instead
                return;
            }

            // Keep max 2 cards — remove oldest if full
            const existing = spot.querySelectorAll('.inu-mon-spot-card');
            if (existing.length >= 2) existing[0].remove();
            spot.querySelectorAll('.inu-mon-spot-card').forEach(c => c.classList.add('secondary'));
            const empty = spot.querySelector('.inu-mon-spot-empty');
            if (empty) empty.remove();

            const card = document.createElement('div');
            card.className = 'inu-mon-spot-card flash' + (outOfRange ? ' oor' : '');
            card.dataset.sensor = group.key;
            const eNew = this.esc(newVal);
            const rangeNote = outOfRange ? `<div class="inu-mon-spot-oor"><i class="fa fa-exclamation-triangle"></i> Utanför skalning (${pvTag.engmin}…${pvTag.engmax} ${this.esc(pvTag.unit)})</div>` : '';
            const faultLabel = this.faultDescMap[group.key] || 'Driftfel';
            const faultNote = faultOn ? `<div class="inu-mon-spot-fault"><i class="fa fa-exclamation-circle"></i> ${this.esc(faultLabel)}</div>` : '';
            card.innerHTML = `
<div class="inu-mon-spot-sensor">${this.esc(this.displayName(group.key))}</div>
<div class="inu-mon-spot-val">${eNew}</div>
${faultNote}
${rangeNote}
${this.buildTimelineHtml(group.key)}`;
            spot.appendChild(card);
            setTimeout(() => card.classList.remove('flash'), 3000);
            // Fit text for all visible cards (refit after layout change)
            // Defer fitText to next frame so the DOM has rendered dimensions
            requestAnimationFrame(() => {
                spot.querySelectorAll('.inu-mon-spot-card').forEach(c => {
                    const v = c.querySelector('.inu-mon-spot-val');
                    if (v && c.clientWidth > 0) this.fitText(v, c);
                });
            });
            if (this.sound) this.beep();
        }

        fitText(el, container) {
            if (!container.clientWidth || !container.clientHeight) return;
            const len = (el.textContent || '').length || 1;
            const maxW = container.clientWidth - 32;
            const valSize = Math.max(28, Math.min(Math.floor(maxW / (len * 0.6)), 220, container.clientHeight * 0.45));
            el.style.fontSize = valSize + 'px';
            // Scale sensor name to 90% and labels to 50% of value size
            const sensor = container.querySelector('.inu-mon-spot-sensor');
            if (sensor) sensor.style.fontSize = Math.round(valSize * 0.35) + 'px';
            container.querySelectorAll('.inu-mon-spot-fault, .inu-mon-spot-oor').forEach(l => {
                l.style.fontSize = Math.max(10, Math.round(valSize * 0.14)) + 'px';
            });
        }

        beep() {
            try {
                if (!this.audioCtx) this.audioCtx = new (window.AudioContext || window.webkitAudioContext)();
                const osc = this.audioCtx.createOscillator();
                const gain = this.audioCtx.createGain();
                osc.frequency.value = 880;
                gain.gain.value = 0.3;
                osc.connect(gain);
                gain.connect(this.audioCtx.destination);
                osc.start();
                osc.stop(this.audioCtx.currentTime + 0.12);
            } catch (e) {}
        }

        updateClock() {
            const cl = this.el?.querySelector('.inu-mon-clock');
            if (cl) cl.textContent = new Date().toLocaleTimeString('sv-SE');
        }

        bindButtons() {
            this.el.addEventListener('click', e => {
                const btn = e.target.closest('[data-act]');
                if (!btn) return;
                const act = btn.dataset.act;
                if (act === 'close') this.close();
                if (act === 'theme') { this.dark = !this.dark; this.el.classList.toggle('dark', this.dark); this.el.classList.toggle('light', !this.dark); }
                if (act === 'oor') { this.spotOOR = !this.spotOOR; btn.classList.toggle('inu-mon-btn-active', this.spotOOR); if (!this.spotOOR) this.clearOORSpotCards(); }
                if (act === 'sound') { this.sound = !this.sound; btn.querySelector('i').className = 'fa fa-volume-' + (this.sound ? 'up' : 'off'); }
                if (act === 'full') {
                    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
                    else this.el.requestFullscreen().catch(() => {});
                }
            });
        }

        bindKeys() {
            this._keyHandler = e => {
                if (e.key === 'Escape') this.close();
                if (e.key.toLowerCase() === 'f') { if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); else this.el.requestFullscreen().catch(() => {}); }
                if (e.key.toLowerCase() === 'd') { this.dark = !this.dark; this.el.classList.toggle('dark', this.dark); this.el.classList.toggle('light', !this.dark); }
                if (e.key.toLowerCase() === 'o') { this.spotOOR = !this.spotOOR; const b = this.el.querySelector('[data-act="oor"]'); if (b) b.classList.toggle('inu-mon-btn-active', this.spotOOR); if (!this.spotOOR) this.clearOORSpotCards(); }
                if (e.key.toLowerCase() === 's') { this.sound = !this.sound; const i = this.el.querySelector('[data-act="sound"] i'); if (i) i.className = 'fa fa-volume-' + (this.sound ? 'up' : 'off'); }
                if (e.key.toLowerCase() === 'c') { const s = this.el.querySelector('#inu-mon-spot'); if (s) s.innerHTML = '<div class="inu-mon-spot-empty"><i class="fa fa-plug"></i><br>Väntar på förändring...</div>'; }
            };
            document.addEventListener('keydown', this._keyHandler);
        }

        injectCSS() {
            if (document.getElementById('inu-mon-css')) return;
            const s = document.createElement('style');
            s.id = 'inu-mon-css';
            s.textContent = `
#inu-monitor { position:fixed; inset:0; z-index:200000; display:flex; flex-direction:column; font-family:'Open Sans',sans-serif; }
#inu-monitor.dark { background:#1a1a2e; color:#e0e0e0; }
#inu-monitor.light { background:#f0f2f5; color:#333; }
.inu-mon-hdr { display:flex; align-items:center; padding:8px 16px; gap:16px; flex-shrink:0; }
.dark .inu-mon-hdr { background:#16213e; }
.light .inu-mon-hdr { background:#fff; border-bottom:1px solid #ddd; }
.inu-mon-title { font-size:14px; font-weight:600; display:inline-flex; align-items:center; gap:8px; }
.inu-mon-title svg { height:16px !important; }
.inu-mon-status { font-size:12px; padding:3px 10px; border-radius:10px; font-weight:600; }
.dark .inu-mon-status { background:#2a4a2a; color:#8f8; }
.light .inu-mon-status { background:#e8f5e9; color:#2e7d32; }
.inu-mon-prefix { font-size:14px; font-weight:600; padding:2px 10px; border-radius:4px; letter-spacing:.5px; }
.dark .inu-mon-prefix { background:#2a2a4a; }
.light .inu-mon-prefix { background:#e0e0e0; }
.inu-mon-clock { font-size:20px; font-weight:300; font-variant-numeric:tabular-nums; }
.inu-mon-controls { margin-left:auto; display:flex; gap:6px; }
.inu-mon-btn { background:none; border:1px solid rgba(255,255,255,.2); color:inherit; border-radius:4px; padding:4px 8px; cursor:pointer; font-size:13px; }
.light .inu-mon-btn { border-color:rgba(0,0,0,.15); }
.inu-mon-btn:hover { background:rgba(255,255,255,.1); }
.light .inu-mon-btn:hover { background:rgba(0,0,0,.05); }
.inu-mon-btn.inu-mon-btn-active { background:rgba(255,255,255,.15); border-color:rgba(255,255,255,.5); }
.light .inu-mon-btn.inu-mon-btn-active { background:rgba(0,0,0,.1); border-color:rgba(0,0,0,.35); }
.inu-mon-body { display:flex; flex-direction:row; flex:1; overflow:hidden; }
.inu-mon-main { display:flex; flex-direction:column; flex:1; min-width:0; overflow:hidden; }
.inu-mon-spot { flex:1 1 50%; display:flex; flex-direction:row; align-items:stretch; justify-content:center; gap:12px; padding:12px 24px; position:relative; overflow:hidden; min-height:0; }
.dark .inu-mon-spot { border-bottom:1px solid #2a2a4a; }
.light .inu-mon-spot { border-bottom:1px solid #ddd; }
.inu-mon-grid-wrap { flex:1; position:relative; overflow:hidden; display:flex; flex-direction:column; min-height:0; }
.inu-mon-grid { flex:1; display:flex; flex-wrap:wrap; gap:6px 16px; padding:8px 12px; overflow-y:auto; align-items:flex-start; align-content:start; }
.inu-mon-scroll-hint { position:absolute; bottom:0; left:0; right:0; text-align:center; padding:8px 12px 10px; font-size:12px; font-weight:600; pointer-events:none; }
.dark .inu-mon-scroll-hint { background:linear-gradient(transparent, #1a1a2e 60%); color:rgba(255,255,255,.6); }
.light .inu-mon-scroll-hint { background:linear-gradient(transparent, #f0f0f0 60%); color:rgba(0,0,0,.45); }
.inu-mon-card { border-radius:8px; padding:8px 14px; text-align:center; display:flex; flex-direction:column; align-items:center; transition:border-color .3s, transform .2s; border:2px solid transparent; cursor:pointer; position:relative; user-select:none; min-width:90px; }
.dark .inu-mon-card { background:#16213e; }
.light .inu-mon-card { background:#fff; box-shadow:0 1px 3px rgba(0,0,0,.1); }
.inu-mon-card:hover { transform:scale(1.02); }
.inu-mon-card.changed { border-color:#fdd835; transform:scale(1.04); }
.inu-mon-card.oor { border-color:#e53935; }
.inu-mon-card.oor .inu-mon-card-pv { color:#e53935; }
.inu-mon-card.verified { border-color:#4caf50; }
.inu-mon-card.verified .inu-mon-card-check { display:flex; }
.inu-mon-card-check { display:none; position:absolute; top:4px; right:6px; color:#4caf50; font-size:14px; width:20px; height:20px; align-items:center; justify-content:center; }
.inu-mon-card-hdr { font-size:15px; font-weight:700; margin-bottom:2px; }
.inu-mon-card-pv { font-size:22px; font-weight:800; font-variant-numeric:tabular-nums; transition:color .3s; margin:4px 0; }
.inu-mon-catgrp { display:flex; flex-wrap:wrap; gap:6px; align-items:flex-start; }
.inu-mon-catgrp-hdr { width:100%; font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:1px; padding:0 2px 2px; opacity:.35; }

.inu-mon-card-fault { font-size:9px; margin-top:4px; padding:2px 6px; border-radius:3px; opacity:.4; }
.inu-mon-card-fault.active { opacity:1; background:rgba(229,57,53,.15); color:#e53935; font-weight:600; }
.inu-mon-spot-empty { font-size:16px; opacity:.25; text-align:center; line-height:1.8; width:100%; }
.inu-mon-spot-card { flex:1; min-width:0; text-align:center; display:flex; flex-direction:column; align-items:center; justify-content:center; border:3px solid transparent; border-radius:12px; padding:8px 16px; transition:border-color 1s, flex .3s, opacity .3s; overflow:hidden; }
.inu-mon-spot-card.secondary { flex:0.6; opacity:.6; }
.inu-mon-spot-card.flash { border-color:#fdd835; }
.inu-mon-spot-card.oor { border-color:#e53935; }
.dark .inu-mon-spot-card { background:#16213e; }
.light .inu-mon-spot-card { background:#fff; box-shadow:0 2px 8px rgba(0,0,0,.1); }
.inu-mon-spot-sensor { font-weight:800; }
.inu-mon-spot-val { font-weight:800; line-height:1; font-variant-numeric:tabular-nums; transition:font-size .2s; white-space:nowrap; }
.inu-mon-spot-card.oor .inu-mon-spot-val { color:#e53935; }
.inu-mon-spot-oor, .inu-mon-spot-fault { padding:4px 12px; border-radius:6px; background:rgba(229,57,53,.15); color:#e53935; font-weight:600; }
.inu-tl { display:flex; gap:0; align-items:flex-end; margin-top:auto; padding-top:8px; overflow-x:auto; width:100%; justify-content:center; flex-wrap:nowrap; }
.inu-tl-ev { display:flex; flex-direction:column; align-items:center; flex-shrink:0; }
.inu-tl-time { font-size:9px; opacity:.35; margin-bottom:3px; white-space:nowrap; font-variant-numeric:tabular-nums; }
.inu-tl-box { padding:8px 12px; border-radius:6px; font-size:11px; font-weight:600; white-space:nowrap; border:2px solid transparent; min-width:50px; text-align:center; }
.inu-tl-ev.nok .inu-tl-box { background:rgba(229,57,53,.15); color:#e53935; border-color:#e53935; }
.inu-tl-ev.ok .inu-tl-box { background:rgba(76,175,80,.15); color:#4caf50; border-color:#4caf50; }
.inu-tl-arrow { display:flex; align-items:flex-end; padding:0 4px 10px; opacity:.25; font-size:14px; }
/* Control panel — right-side vertical */
.inu-mon-ctrl-bar { display:flex; flex-direction:column; gap:10px; padding:12px 10px; width:160px; flex-shrink:0; overflow-y:auto; }
.dark  .inu-mon-ctrl-bar { background:rgba(0,0,0,.2); border-left:1px solid rgba(255,255,255,.07); }
.light .inu-mon-ctrl-bar { background:rgba(0,0,0,.04); border-left:1px solid rgba(0,0,0,.1); }
.inu-mon-ctrl-device { display:flex; flex-direction:column; gap:5px; padding:10px; border-radius:10px; }
.dark  .inu-mon-ctrl-device { background:#16213e; }
.light .inu-mon-ctrl-device { background:#fff; box-shadow:0 1px 3px rgba(0,0,0,.1); }
.inu-mon-ctrl-name { font-size:12px; font-weight:700; opacity:.7; margin-bottom:2px; }
.inu-mon-ctrl-btn { width:100%; padding:0; height:40px; font-size:13px; font-weight:700; border:2px solid transparent; border-radius:7px; cursor:pointer; transition:background .1s, border-color .1s; }
.dark  .inu-mon-ctrl-btn { background:rgba(255,255,255,.08); color:#e0e0e0; }
.light .inu-mon-ctrl-btn { background:rgba(0,0,0,.07); color:#333; }
.dark  .inu-mon-ctrl-btn:hover { background:rgba(255,255,255,.14); }
.light .inu-mon-ctrl-btn:hover { background:rgba(0,0,0,.13); }
.inu-mon-ctrl-btn.active { background:#1e3a5f; color:#fff; border-color:#2d5a9e; }
.inu-mon-ctrl-btn.pump-on-active  { background:#1b5e20 !important; border-color:#4caf50 !important; color:#fff; }
.inu-mon-ctrl-btn.pump-off-active { background:#7f0000 !important; border-color:#e53935 !important; color:#fff; }
.inu-mon-ctrl-opm { display:none; flex-direction:column; align-items:center; gap:6px; margin-top:4px; }
.inu-mon-ctrl-slider-wrap { width:72px; height:156px; display:flex; align-items:center; justify-content:center; overflow:visible; }
.inu-mon-ctrl-slider { -webkit-appearance:none; appearance:none; writing-mode:vertical-lr; direction:rtl; width:4px; height:156px; outline:none; cursor:pointer; padding:0; border:none; flex-shrink:0; }
.dark  .inu-mon-ctrl-slider { background:rgba(255,255,255,.15); border-radius:2px; }
.light .inu-mon-ctrl-slider { background:rgba(0,0,0,.18); border-radius:2px; }
.inu-mon-ctrl-slider::-webkit-slider-thumb { -webkit-appearance:none; width:64px; height:20px; border-radius:4px; background:#2d5a9e; cursor:grab; border:2px solid #1a3f73; box-shadow:0 2px 6px rgba(0,0,0,.45); }
.inu-mon-ctrl-slider::-webkit-slider-thumb:active { cursor:grabbing; background:#3a6ec4; }
.inu-mon-ctrl-slider::-moz-range-thumb { width:64px; height:20px; border-radius:4px; background:#2d5a9e; cursor:grab; border:2px solid #1a3f73; box-shadow:0 2px 6px rgba(0,0,0,.45); }
.inu-mon-ctrl-val { font-size:12px; font-weight:700; text-align:center; font-variant-numeric:tabular-nums; opacity:.8; }
.inu-mon-ctrl-running { font-size:11px; font-weight:600; opacity:.5; display:flex; align-items:center; gap:5px; }
.inu-mon-ctrl-running.running { opacity:1; color:#4caf50; }
.inu-mon-card-pv.digital-on  { color:#4caf50; font-weight:700; }
.inu-mon-card-pv.digital-off { opacity:.5; }
.inu-mon-ctrl-fault-ind { font-size:11px; font-weight:700; color:#e53935; display:flex; align-items:center; gap:4px; }
`;
            document.head.appendChild(s);
        }
    }

    // ============================================================
    // IO-ENHETER PAGE — IP address column
    // ============================================================
    function isDevicePage() {
        return !!(document.getElementById('devicetable') && document.querySelector('#devicetable tbody tr.devicerow'));
    }

    const deviceCache = {};

    async function saveDeviceField(guid, mutator) {
        const r = await fetch('/device/ActionEdit?show=1&type=device&guid=' + guid);
        if (!r.ok) throw new Error('Load device: ' + r.status);
        const doc = new DOMParser().parseFromString(await r.text(), 'text/html');
        const form = doc.querySelector('form');
        if (!form) throw new Error('No device form');
        const fd = serializeForm(form);
        mutator(fd, doc);
        const s = await fetch('/device/actionedit', { method: 'POST', body: fd });
        if (!s.ok) throw new Error('Save device: ' + s.status);
    }

    function refreshDeviceTable() {
        // Cache already updated optimistically by the toggle callback.
        // Trigger a quick table refresh so status text/row class updates.
        const sc = document.createElement('script');
        sc.textContent = 'if(typeof GetDeviceList==="function")GetDeviceList(false);';
        document.head.appendChild(sc); sc.remove();
    }

    function setDeviceCheckbox(fd, name, enabled) {
        fd.delete(name);
        if (enabled) fd.append(name, '1');
        fd.append(name, '0');
    }

    function buildDeviceToggle(label, icon, cssClass, checked, onChange) {
        const grp = document.createElement('span'); grp.className = 'tog-grp' + (checked ? ' on' : ''); grp.title = label;
        const ico = document.createElement('i'); ico.className = 'fa ' + icon + ' tog-ico';
        const tog = document.createElement('label'); tog.className = 'larm-tog' + (cssClass ? ' ' + cssClass : '');
        const chk = document.createElement('input'); chk.type = 'checkbox'; chk.checked = checked;
        chk.addEventListener('click', e => e.stopPropagation());
        const sl = document.createElement('span'); sl.className = 'sl';
        tog.appendChild(chk); tog.appendChild(sl);
        ico.addEventListener('click', () => { chk.checked = !chk.checked; chk.dispatchEvent(new Event('change')); });
        grp.appendChild(ico); grp.appendChild(tog);
        chk.addEventListener('change', () => { grp.classList.toggle('on', chk.checked); onChange(chk.checked); });
        return { grp, chk };
    }

    async function pingHost(ip) {
        const r = await fetch('/Tag/Debug?type=ping&ping=' + encodeURIComponent(ip));
        if (!r.ok) throw new Error('HTTP ' + r.status);
        const text = await r.text();
        // Response is plain text, strip HTML tags if any
        return text.replace(/<[^>]+>/g, '').trim();
    }

    function setIpCell(td, ip) {
        td.innerHTML = '';
        td.style.cssText = 'white-space:nowrap;';
        const span = document.createElement('span'); span.textContent = ip;
        td.appendChild(span);
        if (ip && ip !== '-' && ip !== '...' && ip !== 'fel') {
            // Copy button
            const copyBtn = document.createElement('i');
            copyBtn.className = 'fa fa-copy';
            copyBtn.title = 'Kopiera';
            copyBtn.style.cssText = 'margin-left:6px;cursor:pointer;opacity:.3;font-size:11px;';
            copyBtn.addEventListener('mouseenter', () => copyBtn.style.opacity = '1');
            copyBtn.addEventListener('mouseleave', () => copyBtn.style.opacity = '.3');
            copyBtn.addEventListener('click', e => {
                e.stopPropagation();
                navigator.clipboard.writeText(ip).then(() => toastOk('Kopierat: ' + ip));
            });
            td.appendChild(copyBtn);

            // Ping button
            const pingBtn = document.createElement('i');
            pingBtn.className = 'fa fa-signal';
            pingBtn.title = 'Ping';
            pingBtn.style.cssText = 'margin-left:6px;cursor:pointer;opacity:.3;font-size:11px;';
            pingBtn.addEventListener('mouseenter', () => { if (!pingBtn.classList.contains('fa-spinner')) pingBtn.style.opacity = '1'; });
            pingBtn.addEventListener('mouseleave', () => { if (!pingBtn.classList.contains('fa-spinner')) pingBtn.style.opacity = '.3'; });
            pingBtn.addEventListener('click', async e => {
                e.stopPropagation();
                pingBtn.className = 'fa fa-spinner fa-spin';
                pingBtn.style.opacity = '1';
                try {
                    const result = await pingHost(ip);
                    const ok = /status:\s*success|svar från|reply from|bytes=/i.test(result);
                    const rttMatch = result.match(/RoundTrip time:\s*(\d+)/i) || result.match(/tid[<=](\d+)/i);
                    const ttlMatch = result.match(/Time to live:\s*(\d+)/i);
                    const lines = [`<b>${escHtml(ip)}</b>`];
                    lines.push('Status: ' + (ok ? 'OK' : 'Misslyckades'));
                    if (rttMatch) lines.push('RTT: ' + rttMatch[1] + ' ms');
                    if (ttlMatch) lines.push('TTL: ' + ttlMatch[1]);
                    const html = lines.join('<br>');
                    const tr = unsafeWindow.toastr || window.toastr;
                    if (tr) tr[ok ? 'success' : 'error'](html, '', { timeOut: ok ? 3000 : 5000, positionClass: 'toast-bottom-right', escapeHtml: false });
                } catch (err) {
                    toastErr('Ping fel: ' + err.message);
                }
                pingBtn.className = 'fa fa-signal';
                pingBtn.style.opacity = '.3';
            });
            td.appendChild(pingBtn);
        }
    }

    function addDeviceHeaders() {
        const thead = document.querySelector('#devicetable thead tr');
        if (thead && !thead.querySelector('.p-hdr')) {
            const thIp = document.createElement('th'); thIp.className = 'p-hdr'; thIp.textContent = 'IP-adress'; thIp.style.cssText = 'width:120px;';
            const thPort = document.createElement('th'); thPort.className = 'p-hdr'; thPort.textContent = 'Port'; thPort.style.cssText = 'width:50px;';
            const thSlave = document.createElement('th'); thSlave.className = 'p-hdr'; thSlave.textContent = 'Stations-ID'; thSlave.style.cssText = 'width:70px;';
            thead.appendChild(thIp);
            thead.appendChild(thPort);
            thead.appendChild(thSlave);
        }
    }

    function makeDeviceToggles(guid, togWrap, active, alarm, trend) {
        togWrap.innerHTML = '';
        const rollback = (cacheKey, prev, tog) => {
            deviceCache[guid][cacheKey] = prev;
            tog.chk.checked = prev;
            tog.grp.classList.toggle('on', prev);
        };
        const activeTog = buildDeviceToggle('Aktiv', 'fa-power-off', '', active, on => {
            const prev = deviceCache[guid].active;
            deviceCache[guid].active = on;
            saveDeviceField(guid, fd => setDeviceCheckbox(fd, 'Active', on))
                .then(refreshDeviceTable)
                .catch(e => { rollback('active', prev, activeTog); toastErr(e.message); });
        });
        togWrap.appendChild(activeTog.grp);
        const alarmTog = buildDeviceToggle('Larm', 'fa-bell', '', alarm, on => {
            const prev = deviceCache[guid].alarm;
            deviceCache[guid].alarm = on;
            saveDeviceField(guid, fd => setDeviceCheckbox(fd, 'ActiveAlarm', on))
                .then(refreshDeviceTable)
                .catch(e => { rollback('alarm', prev, alarmTog); toastErr(e.message); });
        });
        togWrap.appendChild(alarmTog.grp);
        const trendTog = buildDeviceToggle('Trend', 'fa-line-chart', 'trend', trend, on => {
            const prev = deviceCache[guid].trend;
            deviceCache[guid].trend = on;
            saveDeviceField(guid, fd => setDeviceCheckbox(fd, 'ActiveTrend', on))
                .then(refreshDeviceTable)
                .catch(e => { rollback('trend', prev, trendTog); toastErr(e.message); });
        });
        togWrap.appendChild(trendTog.grp);
    }

    async function addDeviceColumns() {
        const rows = document.querySelectorAll('#devicetable tbody tr.devicerow');
        const toFetch = [];

        // Pass 1: synchronously add all cells from cache (no flicker)
        for (const row of rows) {
            if (row.querySelector('.inu-ip')) continue;
            const guid = row.id;
            const c = deviceCache[guid];

            const tdIp = document.createElement('td'); tdIp.className = 'inu-ip';
            const tdPort = document.createElement('td'); tdPort.className = 'inu-port';
            const tdSlave = document.createElement('td'); tdSlave.className = 'inu-slave';
            setIpCell(tdIp, c?.host || '...'); tdPort.textContent = c?.port || ''; tdSlave.textContent = c?.slave || '';
            row.appendChild(tdIp); row.appendChild(tdPort); row.appendChild(tdSlave);

            const tdStatus = row.cells[2];
            if (tdStatus && !tdStatus.querySelector('.inu-dev-toggles')) {
                const statusText = tdStatus.textContent.trim();
                tdStatus.innerHTML = '';
                tdStatus.style.cssText = 'white-space:nowrap;';
                const wrap = document.createElement('div'); wrap.className = 'inu-dev-toggles';
                wrap.style.cssText = 'display:flex;align-items:center;gap:4px;justify-content:space-between;width:100%;';
                const badge = document.createElement('span');
                const isOk = row.className.includes('OK');
                badge.textContent = statusText;
                badge.style.cssText = `font-size:10px;font-weight:600;padding:2px 8px;border-radius:3px;color:#fff;background:${isOk ? '#1b5e20' : '#b71c1c'};`;
                wrap.appendChild(badge);
                const togWrap = document.createElement('span');
                togWrap.style.cssText = 'display:flex;align-items:center;gap:4px;';
                togWrap.addEventListener('click', e => e.stopPropagation());
                if (c?.active !== undefined) {
                    makeDeviceToggles(guid, togWrap, c.active, c.alarm, c.trend);
                } else {
                    toFetch.push({ guid, tdIp, tdPort, tdSlave, togWrap });
                }
                wrap.appendChild(togWrap);
                tdStatus.appendChild(wrap);
            }
        }

        // Pass 2: async fetch for rows without cached toggle state
        for (const item of toFetch) {
            try {
                const r = await fetch('/device/ActionEdit?show=1&type=device&guid=' + item.guid);
                const doc = new DOMParser().parseFromString(await r.text(), 'text/html');
                const host = doc.querySelector('textarea[name="Host"],input[name="Host"]')?.value || '-';
                const port = doc.querySelector('textarea[name="Port"],input[name="Port"]')?.value || '-';
                const slaveEl = doc.querySelector('[name="Slave Address"],[name="Station"],[name="Unit ID"],[name="Node ID"],[name="Device Address"]');
                const slave = slaveEl ? slaveEl.value : '-';
                const active = doc.querySelector('input[name="Active"][type="checkbox"]')?.checked ?? false;
                const alarm = doc.querySelector('input[name="ActiveAlarm"][type="checkbox"]')?.checked ?? false;
                const trend = doc.querySelector('input[name="ActiveTrend"][type="checkbox"]')?.checked ?? false;
                deviceCache[item.guid] = { host, port, slave, active, alarm, trend };
                setIpCell(item.tdIp, host); item.tdPort.textContent = port; item.tdSlave.textContent = slave;
                makeDeviceToggles(item.guid, item.togWrap, active, alarm, trend);
            } catch (e) {
                setIpCell(item.tdIp, 'fel');
                console.warn(CFG.logPrefix, 'Device fetch failed', item.guid, e);
            }
        }
    }

    // Device cell → form field mapping (cell indices are: 0=Namn, 1=Typ, 2=Tillstånd, 3=Beskrivning, 4=IP, 5=Port, 6=Slave)
    const DEV_EDITABLE = {
        0: { field: 'name', label: 'Namn' },
        3: { field: 'desc', label: 'Beskrivning' },
        4: { field: 'Host', label: 'IP-adress' },
        5: { field: 'Port', label: 'Port' },
        6: { field: '_slave', label: 'Stations-ID' },
    };

    function startDeviceCellEdit(cell, row, fieldName) {
        if (cell.querySelector('input')) return;
        const guid = row.id;
        if (!guid) return;
        const cur = cell.textContent.trim();
        const blocker = e => { e.stopPropagation(); e.stopImmediatePropagation(); };
        cell.addEventListener('click', blocker, true);
        cell.addEventListener('mousedown', blocker, true);
        const input = document.createElement('input');
        input.value = cur;
        input.style.cssText = 'width:100%;font-size:11px;padding:1px 3px;border:1px solid #5b6abf;border-radius:2px;box-sizing:border-box;';
        cell.textContent = '';
        cell.appendChild(input);
        input.focus(); input.select();
        const save = async () => {
            cell.removeEventListener('click', blocker, true);
            cell.removeEventListener('mousedown', blocker, true);
            const nv = input.value.trim();
            cell.textContent = nv;
            if (nv !== cur) {
                try {
                    await saveDeviceField(guid, (fd, doc) => {
                        if (fieldName === '_slave') {
                            const el = doc.querySelector('[name="Slave Address"],[name="Station"],[name="Unit ID"],[name="Node ID"],[name="Device Address"]');
                            if (el) fd.set(el.name, nv);
                        } else {
                            fd.set(fieldName, nv);
                        }
                    });
                    // Update cache
                    if (deviceCache[guid]) {
                        if (fieldName === 'Host') deviceCache[guid].host = nv;
                        else if (fieldName === 'Port') deviceCache[guid].port = nv;
                        else if (fieldName === '_slave') deviceCache[guid].slave = nv;
                    }
                    toast(`${nv} → sparad`);
                } catch (err) { cell.textContent = cur; toastErr(err.message); }
            }
        };
        input.addEventListener('blur', save);
        input.addEventListener('keydown', ev => {
            if (ev.key === 'Enter') { ev.preventDefault(); input.blur(); }
            if (ev.key === 'Escape') {
                input.removeEventListener('blur', save);
                cell.removeEventListener('click', blocker, true);
                cell.removeEventListener('mousedown', blocker, true);
                cell.textContent = cur;
            }
        });
    }

    function initDeviceContextMenu() {
        let devCtxCell = -1;
        document.addEventListener('contextmenu', e => {
            const td = e.target.closest('#devicetable tbody td');
            devCtxCell = td ? td.cellIndex : -1;
        }, true);

        document.addEventListener('inu-dev-ctx', e => {
            const { action, rowId, cellIndex } = e.detail;
            const row = document.getElementById(rowId);
            if (!row) return;
            if (action === 'edit') {
                const def = DEV_EDITABLE[cellIndex];
                const cell = row.cells[cellIndex];
                if (def && cell) startDeviceCellEdit(cell, row, def.field);
            }
        });

        const sc = document.createElement('script');
        sc.textContent = `(function(){
            if(!$.contextMenu) return;
            var editable={0:true,3:true,4:true,5:true,6:true};
            var _cell=-1;
            document.addEventListener('contextmenu',function(e){
                var td=e.target.closest('#devicetable tbody td');
                _cell=td?td.cellIndex:-1;
            },true);
            $.contextMenu({
                selector:'#devicetable tbody tr.devicerow',
                build:function($trigger,e){
                    var items={};
                    if(editable[_cell]){
                        items.edit={name:'Redigera fält',icon:'fa-pencil'};
                    }
                    return {
                        items:items,
                        callback:function(key,opt){
                            var rowId=opt.\$trigger.attr('id');
                            document.dispatchEvent(new CustomEvent('inu-dev-ctx',{detail:{action:key,rowId:rowId,cellIndex:_cell}}));
                        }
                    };
                }
            });
        })();`;
        document.head.appendChild(sc); sc.remove();
    }

    // ============================================================
    // PAGE EDITOR — grid snap, alignment, undo, keyboard, z-order
    // ============================================================
    function isPageEditorPage() {
        const url = new URL(window.location.href);
        return url.pathname.startsWith('/page') && url.searchParams.get('route') === 'edit';
    }

    const EDITOR_GRID_SIZES = [4, 8, 16, 24, 32];
    // serverValue matches WebPort's /Page/PageObjectProperties textposition field:
    // 0=Vänster, 1=Över L, 2=Höger, 3=Under L, 4=Under C, 5=Över C, 6=Under R, 7=Över R
    const TEXT_POSITIONS = [
        { id: 'top-left',      serverValue: 1, label: 'Över vänster',    cls: ['wpCompTextPositionTop',    'wpCompTextLeft']  },
        { id: 'top-right',     serverValue: 7, label: 'Över höger',      cls: ['wpCompTextPositionTop',    'wpCompTextRight'] },
        { id: 'top-center',    serverValue: 5, label: 'Över centrerat',  cls: ['wpCompTextPositionTop']                       },
        { id: 'right',         serverValue: 2, label: 'Höger',           cls: ['wpCompTextPositionRight']                     },
        { id: 'bottom-left',   serverValue: 3, label: 'Under vänster',   cls: ['wpCompTextPositionBottom', 'wpCompTextLeft']  },
        { id: 'bottom-right',  serverValue: 6, label: 'Under höger',     cls: ['wpCompTextPositionBottom', 'wpCompTextRight'] },
        { id: 'bottom-center', serverValue: 4, label: 'Under centrerat', cls: ['wpCompTextPositionBottom']                    },
        { id: 'left',          serverValue: 0, label: 'Vänster',         cls: ['wpCompTextPositionLeft']                      },
    ];
    let editorGridEnabled = GM_getValue('inu_editor_grid_enabled', false);
    let editorGridSize    = GM_getValue('inu_editor_grid_size', 8);
    let editorUndoStack = [];
    let editorRedoStack = [];
    const EDITOR_UNDO_MAX = 20;
    let _editorIframe  = null;
    let _editorSelCount = 0;
    let _editorPosClipboard = null; // {x, y} for copy/paste position
    let _editorLocked = {};         // {elementId: true} persisted via GM_setValue

    function editorGetIframe() { return document.querySelector('#content iframe'); }

    function editorGetSelected() {
        if (!_editorIframe) return [];
        return [..._editorIframe.contentDocument.querySelectorAll('.wpCompObject.ui-selected')];
    }

    // ── Grid ──────────────────────────────────────────────────────
    function editorApplyGrid(el) {
        if (!_editorIframe) return;
        const $ = _editorIframe.contentWindow.jQuery;
        const targets = el ? [el] : [..._editorIframe.contentDocument.querySelectorAll('.wpCompObject')];
        targets.forEach(t => {
            if (!$(t).data('ui-draggable')) return;
            $(t).draggable('option', 'grid', editorGridEnabled ? [editorGridSize, editorGridSize] : false);
        });
    }

    function editorSnapAll() {
        if (!_editorIframe) return;
        const sel = editorGetSelected();
        if (!sel.length) return;
        editorPushUndo();
        sel.forEach(el => {
            el.style.left = Math.round(parseFloat(el.style.left) / editorGridSize) * editorGridSize + 'px';
            el.style.top  = Math.round(parseFloat(el.style.top)  / editorGridSize) * editorGridSize + 'px';
            el.style.transform = '';
        });
        editorSavePositions(sel);
    }

    // ── Alignment ─────────────────────────────────────────────────
    function editorAlign(axis, mode) {
        const sel = editorGetSelected();
        if (sel.length < 2) return;
        editorPushUndo();
        const pos = sel.map(el => ({ el, x: parseFloat(el.style.left)||0, y: parseFloat(el.style.top)||0, w: el.offsetWidth, h: el.offsetHeight }));
        pos.forEach(p => {
            let v;
            if (axis === 'x') {
                v = mode === 'min' ? Math.min(...pos.map(q => q.x))
                  : mode === 'max' ? Math.max(...pos.map(q => q.x + q.w)) - p.w
                  : Math.round(pos.reduce((s,q) => s + q.x + q.w/2, 0) / pos.length - p.w/2);
                if (editorGridEnabled) v = Math.round(v / editorGridSize) * editorGridSize;
                p.el.style.left = v + 'px';
            } else {
                v = mode === 'min' ? Math.min(...pos.map(q => q.y))
                  : mode === 'max' ? Math.max(...pos.map(q => q.y + q.h)) - p.h
                  : Math.round(pos.reduce((s,q) => s + q.y + q.h/2, 0) / pos.length - p.h/2);
                if (editorGridEnabled) v = Math.round(v / editorGridSize) * editorGridSize;
                p.el.style.top = v + 'px';
            }
            p.el.style.transform = '';
        });
        editorSavePositions(pos.map(p => p.el));
    }

    function editorDistribute(axis) {
        const sel = editorGetSelected();
        if (sel.length < 3) return;
        editorPushUndo();
        const sorted = [...sel].sort((a,b) =>
            (parseFloat(axis==='x' ? a.style.left : a.style.top)||0) -
            (parseFloat(axis==='x' ? b.style.left : b.style.top)||0)
        );
        // Equal-gap distribute: equal whitespace between elements (not equal centerpoints)
        const firstPos  = parseFloat(axis==='x' ? sorted[0].style.left  : sorted[0].style.top)  || 0;
        const lastPos   = parseFloat(axis==='x' ? sorted[sorted.length-1].style.left : sorted[sorted.length-1].style.top) || 0;
        const lastSize  = axis==='x' ? sorted[sorted.length-1].offsetWidth : sorted[sorted.length-1].offsetHeight;
        const totalSpan = (lastPos + lastSize) - firstPos;
        const totalSize = sorted.reduce((s, el) => s + (axis==='x' ? el.offsetWidth : el.offsetHeight), 0);
        const gap = (totalSpan - totalSize) / (sorted.length - 1);
        let cursor = firstPos;
        sorted.forEach(el => {
            const v = Math.round(cursor);
            if (axis === 'x') el.style.left = v + 'px'; else el.style.top = v + 'px';
            el.style.transform = '';
            cursor += (axis==='x' ? el.offsetWidth : el.offsetHeight) + gap;
        });
        editorSavePositions(sorted);
    }

    // ── Sort by name ─────────────────────────────────────────────
    function editorSortByName(dir = 'asc') {
        const sel = editorGetSelected();
        if (sel.length < 2) return;
        editorPushUndo();

        // Collect positions sorted in reading order (top→bottom, left→right)
        const positions = sel
            .map(el => ({ left: parseFloat(el.style.left) || 0, top: parseFloat(el.style.top) || 0 }))
            .sort((a, b) => a.top !== b.top ? a.top - b.top : a.left - b.left);

        // Sort elements alphabetically by prefix → text content → id
        const getLabel = el => {
            const prefix = el.getAttribute('data-prefix');
            if (prefix) return prefix.toLowerCase();
            const txt = el.querySelector('.wpCompText');
            if (txt?.textContent?.trim()) return txt.textContent.trim().toLowerCase();
            return (el.id || '').toLowerCase();
        };
        const sorted = [...sel].sort((a, b) => {
            const cmp = getLabel(a).localeCompare(getLabel(b), 'sv');
            return dir === 'asc' ? cmp : -cmp;
        });

        // Assign sorted elements to reading-order positions
        sorted.forEach((el, i) => {
            el.style.left = positions[i].left + 'px';
            el.style.top  = positions[i].top  + 'px';
        });
        editorSavePositions(sorted);
        editorUpdateToolbarContext();
    }


    // ── Position inspector ────────────────────────────────────────
    function editorUpdatePositionInspector(el) {
        const iDoc = _editorIframe?.contentDocument;
        const xi = iDoc?.getElementById('inu-et-pos-x');
        const yi = iDoc?.getElementById('inu-et-pos-y');
        const wi = iDoc?.getElementById('inu-et-pos-w');
        const hi = iDoc?.getElementById('inu-et-pos-h');
        if (!xi) return;
        xi.value = parseFloat(el.style.left) || 0;
        yi.value = parseFloat(el.style.top)  || 0;
        if (wi) wi.textContent = el.offsetWidth;
        if (hi) hi.textContent = el.offsetHeight;
    }

    // ── Select All / None ─────────────────────────────────────────
    function editorSelectAll() {
        if (!_editorIframe) return;
        _editorIframe.contentDocument.querySelectorAll('.wpCompObject').forEach(el => {
            if (!_editorLocked[el.id]) el.classList.add('ui-selected');
        });
        editorUpdateToolbarContext();
    }

    function editorDeselectAll() {
        if (!_editorIframe) return;
        _editorIframe.contentDocument.querySelectorAll('.wpCompObject.ui-selected').forEach(el => el.classList.remove('ui-selected'));
        editorUpdateToolbarContext();
    }

    // ── Undo stack ────────────────────────────────────────────────
    // Rotation is applied only to .wpCompImage (the SVG container) so text labels are unaffected.
    // The transform string is stored directly as a snapshot field.
    function editorGetImgTransform(el) {
        return el.querySelector('.wpCompImage')?.style.transform || '';
    }

    function editorSnapshot() {
        if (!_editorIframe) return [];
        return [..._editorIframe.contentDocument.querySelectorAll('.wpCompObject')].map(el => ({
            id: el.id, left: el.style.left, top: el.style.top,
            zIndex: el.style.zIndex, dataZIndex: el.dataset.zindex,
            imgTransform: editorGetImgTransform(el),
        }));
    }

    function editorSnapshotsEqual(a, b) {
        if (!a || !b || a.length !== b.length) return false;
        return a.every((s, i) => s.left === b[i].left && s.top === b[i].top && s.zIndex === b[i].zIndex && s.imgTransform === b[i].imgTransform);
    }

    function editorPushUndo() {
        const snap = editorSnapshot();
        if (editorSnapshotsEqual(snap, editorUndoStack[editorUndoStack.length - 1])) return;
        editorUndoStack.push(snap);
        if (editorUndoStack.length > EDITOR_UNDO_MAX) editorUndoStack.shift();
        editorRedoStack = [];
        editorUpdateUndoButtons();
    }

    function editorRestoreSnapshot(snap) {
        if (!_editorIframe) return;
        const iDoc = _editorIframe.contentDocument;
        snap.forEach(s => {
            const el = iDoc.getElementById(s.id);
            if (!el) return;
            el.style.left = s.left; el.style.top = s.top;
            el.style.zIndex = s.zIndex; el.dataset.zindex = s.dataZIndex;
            el.style.transform = '';
            if (s.imgTransform !== undefined) {
                const img = el.querySelector('.wpCompImage');
                if (img) img.style.transform = s.imgTransform;
            }
        });
        editorUpdateToolbarContext();
    }

    function editorUndo() {
        if (!editorUndoStack.length) return;
        editorRedoStack.push(editorSnapshot());
        editorRestoreSnapshot(editorUndoStack.pop());
        editorUpdateUndoButtons();
    }

    function editorRedo() {
        if (!editorRedoStack.length) return;
        editorUndoStack.push(editorSnapshot());
        editorRestoreSnapshot(editorRedoStack.pop());
        editorUpdateUndoButtons();
    }

    function editorUpdateUndoButtons() {
        const iDoc = _editorIframe?.contentDocument;
        const bu = iDoc?.getElementById('inu-et-undo');
        const br = iDoc?.getElementById('inu-et-redo');
        if (bu) bu.disabled = !editorUndoStack.length;
        if (br) br.disabled = !editorRedoStack.length;
    }

    // ── Toolbar context (selection → show/hide groups) ────────────
    function editorUpdateToolbarContext() {
        const iDoc = _editorIframe?.contentDocument;
        const sel = editorGetSelected();
        _editorSelCount = sel.length;

        const counter = iDoc?.getElementById('inu-et-sel-count');
        if (counter) counter.textContent = sel.length + ' markerade';

        const gAlign     = iDoc?.getElementById('inu-et-align-group');
        const gSort      = iDoc?.getElementById('inu-et-sort-group');
        const gSize      = iDoc?.getElementById('inu-et-step-dist-group');
        const gPos       = iDoc?.getElementById('inu-et-pos-group');
        const gTextPos   = iDoc?.getElementById('inu-et-textpos-group');
        const gLock      = iDoc?.getElementById('inu-et-lock-group');
        const gUnlockAll = iDoc?.getElementById('inu-et-unlock-all-group');
        if (gAlign)      gAlign.style.display      = sel.length >= 2 ? '' : 'none';
        if (gSort)       gSort.style.display       = sel.length >= 2 ? '' : 'none';
        if (gSize)       gSize.style.display       = sel.length >= 2 ? '' : 'none';
        if (gPos)        gPos.style.display        = sel.length === 1 ? '' : 'none';
        if (gTextPos)    gTextPos.style.display    = sel.length >= 1 ? '' : 'none';
        if (gLock)       gLock.style.display       = sel.length >= 1 ? '' : 'none';
        if (gUnlockAll)  gUnlockAll.style.display  = Object.keys(_editorLocked).length > 0 ? '' : 'none';

        if (sel.length >= 1) {
            const lockBtn = iDoc?.getElementById('inu-et-lock-btn');
            if (lockBtn) {
                const anyLocked = sel.some(el => _editorLocked[el.id]);
                lockBtn.title = anyLocked ? 'Lås upp (Ctrl+L)' : 'Lås (Ctrl+L)';
                lockBtn.textContent = anyLocked ? '🔓' : '🔒';
            }
        }
        if (sel.length === 1) {
            editorUpdatePositionInspector(sel[0]);
        }
        if (sel.length >= 1) editorUpdateTextPosDropdown();

        editorUpdateStatusBar();
    }

    function editorUpdateStatusBar() {
        if (!_editorIframe) return;
        const iDoc = _editorIframe.contentDocument;
        const c = iDoc.getElementById('inu-sb-count');
        const s = iDoc.getElementById('inu-sb-sel');
        const p = iDoc.getElementById('inu-sb-pos-clip');
        if (c) c.textContent = iDoc.querySelectorAll('.wpCompObject').length + ' komponenter';
        if (s) s.textContent = _editorSelCount + ' markerade';
        if (p) p.style.display = _editorPosClipboard ? '' : 'none';
        editorUpdateZoom();
    }

    // ── Text position ─────────────────────────────────────────────
    function editorGetTextPosition(el) {
        const d = el.querySelector('.wpCompText');
        if (!d) return null;
        const h = c => d.classList.contains(c);
        if (h('wpCompTextPositionTop')    && h('wpCompTextLeft'))  return 'top-left';
        if (h('wpCompTextPositionTop')    && h('wpCompTextRight')) return 'top-right';
        if (h('wpCompTextPositionTop'))                            return 'top-center';
        if (h('wpCompTextPositionRight'))                          return 'right';
        if (h('wpCompTextPositionBottom') && h('wpCompTextLeft'))  return 'bottom-left';
        if (h('wpCompTextPositionBottom') && h('wpCompTextRight')) return 'bottom-right';
        if (h('wpCompTextPositionBottom'))                         return 'bottom-center';
        if (h('wpCompTextPositionLeft'))                           return 'left';
        return null;
    }

    function editorSetTextPosition(posId) {
        const sel = editorGetSelected();
        if (!sel.length) return;
        const pos = TEXT_POSITIONS.find(p => p.id === posId);
        if (!pos) return;
        editorPushUndo();
        const iwin = _editorIframe.contentWindow;
        const $    = iwin.jQuery;
        // Deselect all before any renderComponentV2 call to prevent WebPort's
        // snap-all-selected side-effect (was triggered by Ctrl+Arrow; Alt+Arrow avoids
        // WebPort's native handler, but guard remains for safety).
        sel.forEach(el => el.classList.remove('ui-selected'));
        sel.forEach(el => {
            const d = el.querySelector('.wpCompText');
            if (!d) return;
            [...d.classList].filter(c => /^wpCompTextPosition|^wpCompTextLeft$|^wpCompTextRight$/.test(c))
                .forEach(c => d.classList.remove(c));
            pos.cls.forEach(c => d.classList.add(c));
            iwin.renderComponentV2?.($(el));
            // renderComponentV2 right-aligns text when no horizontal class is present
            // (treats missing class as right-aligned). Override left for centered variants only.
            if (posId === 'top-center' || posId === 'bottom-center')
                d.style.left = ((el.offsetWidth - d.offsetWidth) / 2) + 'px';
            editorSaveTextPosition(el, pos.serverValue);
        });
        sel.forEach(el => el.classList.add('ui-selected'));
        editorUpdateTextPosDropdown();
        editorUpdateToolbarContext();
    }

    async function editorSaveTextPosition(el, serverValue) {
        try {
            // pageid from URL (plain underscores, e.g. "VENTILATION_FL01_02_...")
            const pageId = new URLSearchParams(location.search).get('pageid');
            // id is the full element id (e.g. "VENTILATION-5F-...-2E-FL01-5F-KF101")
            const poid = el.id;
            if (!pageId || !poid) return;
            // Fetch current component form (contains all settings with current values)
            const getResp = await fetch(`/Page/PageObjectProperties?pageid=${pageId}&id=${poid}`);
            const html    = await getResp.text();
            const doc     = new DOMParser().parseFromString(html, 'text/html');
            const form    = doc.getElementById('poproperties');
            if (!form) return;
            // Serialize form fields manually to handle checkbox+hidden pairs correctly
            const fd = new FormData();
            for (const f of form.elements) {
                if (!f.name) continue;
                if ((f.type === 'checkbox' || f.type === 'radio') && !f.checked) continue;
                fd.append(f.name, f.value);
            }
            fd.set('textposition', String(serverValue));
            await fetch('/page/UpdatePageObjectProperties', { method: 'POST', body: fd });
        } catch (e) {
            console.warn('INU: failed to sync text position to server', e);
        }
    }

    // Save posx/posy for one or more elements to the server.
    // Pass an array of .wpCompObject elements; runs all GETs in parallel then all POSTs in parallel.
    async function editorSavePositions(elements) {
        const pageId = new URLSearchParams(location.search).get('pageid');
        if (!pageId || !elements?.length) return;
        try {
            await Promise.all(elements.map(async el => {
                const poid = el.id;
                if (!poid) return;
                const x = parseFloat(el.style.left) || 0;
                const y = parseFloat(el.style.top)  || 0;
                const getResp = await fetch(`/Page/PageObjectProperties?pageid=${pageId}&id=${poid}`);
                const html    = await getResp.text();
                const doc     = new DOMParser().parseFromString(html, 'text/html');
                const form    = doc.getElementById('poproperties');
                if (!form) return;
                const fd = new FormData();
                for (const f of form.elements) {
                    if (!f.name) continue;
                    if ((f.type === 'checkbox' || f.type === 'radio') && !f.checked) continue;
                    fd.append(f.name, f.value);
                }
                fd.set('posx', String(x));
                fd.set('posy', String(y));
                await fetch('/page/UpdatePageObjectProperties', { method: 'POST', body: fd });
            }));
        } catch (e) {
            console.warn('INU: failed to sync positions to server', e);
        }
    }


    function editorUpdateTextPosDropdown() {
        const iDoc = _editorIframe?.contentDocument;
        const dd = iDoc?.getElementById('inu-et-textpos-select');
        if (!dd) return;
        const sel = editorGetSelected();
        if (!sel.length) return;
        const pos = editorGetTextPosition(sel[0]);
        dd.value = pos || '';
    }

    // ── Step distribution ─────────────────────────────────────────
    // Keeps the topmost selected element fixed, places the rest at base + n*step px
    function editorDistributeByStep(step) {
        const sel = editorGetSelected();
        if (sel.length < 2) return;
        editorPushUndo();
        const sorted = [...sel].sort((a, b) => (parseFloat(a.style.top)||0) - (parseFloat(b.style.top)||0));
        const baseTop = parseFloat(sorted[0].style.top) || 0;
        sorted.forEach((el, i) => {
            el.style.top = Math.round(baseTop + i * step) + 'px';
            el.style.transform = '';
        });
        editorSavePositions(sorted);
        editorUpdateToolbarContext();
    }

    // ── Copy / Paste position ─────────────────────────────────────
    function editorCopyPosition() {
        const sel = editorGetSelected();
        if (sel.length !== 1) return;
        _editorPosClipboard = { x: parseFloat(sel[0].style.left) || 0, y: parseFloat(sel[0].style.top) || 0 };
        toastr.info('Position kopierad (Ctrl+Shift+V för att klistra in)');
        editorUpdateStatusBar();
    }

    function editorPastePosition() {
        if (!_editorPosClipboard) return;
        const sel = editorGetSelected();
        if (!sel.length) return;
        editorPushUndo();
        sel.forEach(el => {
            el.style.left = _editorPosClipboard.x + 'px';
            el.style.top  = _editorPosClipboard.y + 'px';
            el.style.transform = '';
        });
        editorSavePositions(sel);
        editorUpdateToolbarContext();
    }

    // ── Lock ──────────────────────────────────────────────────────
    function editorLoadLocks() {
        try { _editorLocked = JSON.parse(GM_getValue('inu_locked_ids', '{}')); } catch { _editorLocked = {}; }
    }
    function editorSaveLocks() {
        GM_setValue('inu_locked_ids', JSON.stringify(_editorLocked));
    }
    function editorApplyLock(el, locked) {
        if (!_editorIframe) return;
        const $ = _editorIframe.contentWindow.jQuery;
        el.classList.toggle('inu-locked', locked);
        if ($(el).data('ui-draggable')) $(el).draggable(locked ? 'disable' : 'enable');
        let badge = el.querySelector('.inu-lock-badge');
        if (locked && !badge) {
            badge = _editorIframe.contentDocument.createElement('span');
            badge.className = 'inu-lock-badge';
            badge.textContent = '🔒';
            el.appendChild(badge);
        } else if (!locked && badge) {
            badge.remove();
        }
    }
    function editorToggleLock() {
        const sel = editorGetSelected();
        if (!sel.length) return;
        const anyUnlocked = sel.some(el => !_editorLocked[el.id]);
        sel.forEach(el => {
            if (anyUnlocked) _editorLocked[el.id] = true;
            else delete _editorLocked[el.id];
            editorApplyLock(el, anyUnlocked);
        });
        editorSaveLocks();
        editorUpdateToolbarContext();
    }
    function editorUnlockAll() {
        if (!_editorIframe) return;
        _editorIframe.contentDocument.querySelectorAll('.wpCompObject').forEach(el => {
            delete _editorLocked[el.id];
            editorApplyLock(el, false);
        });
        editorSaveLocks();
        editorUpdateToolbarContext();
    }

    function editorApplyAllLocks() {
        if (!_editorIframe) return;
        _editorIframe.contentDocument.querySelectorAll('.wpCompObject').forEach(el => {
            if (_editorLocked[el.id]) editorApplyLock(el, true);
        });
    }

    // ── Nudge hint ────────────────────────────────────────────────
    function editorUpdateNudgeHint() {
        const iDoc = _editorIframe?.contentDocument;
        const el = iDoc?.getElementById('inu-et-nudge-hint');
        if (el) el.textContent = 'Nudge: ' + (editorGridEnabled ? editorGridSize : 1) + ' px';
    }

    // ── Zoom ──────────────────────────────────────────────────────
    function editorUpdateZoom() {
        const iDoc = _editorIframe?.contentDocument;
        const wpp = iDoc?.getElementById('wpp');
        if (!wpp) return;
        const t = wpp.style.transform || getComputedStyle(wpp).transform || '';
        let pct = 100;
        // "translate(Xpx, Ypx) scale(s, s)" — WebPort's inline style format
        const sc = t.match(/scale\(\s*([\d.]+)/);
        if (sc) { pct = Math.round(parseFloat(sc[1]) * 100); }
        // Fallback: "matrix(a, b, c, d, tx, ty)"
        else { const mx = t.match(/matrix\(\s*([\d.]+)/); if (mx) pct = Math.round(parseFloat(mx[1]) * 100); }
        const el = iDoc.getElementById('inu-sb-zoom');
        if (el) el.textContent = 'Zoom: ' + pct + '%';
    }

    // ── Keyboard shortcuts ────────────────────────────────────────
    function editorBindKeyboard(iDoc) {
        // Listen on the iframe's window (not document) in capture phase so we fire
        // before any document-level or window-level capture handlers WebPort registers
        // (e.g. Ctrl+Up/Down Z-order stacking, native 1px nudge).
        iDoc.defaultView.addEventListener('keydown', e => {
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'SELECT') return;
            const ctrl = e.ctrlKey || e.metaKey;
            if (ctrl && !e.shiftKey && e.key.toLowerCase() === 'z') { e.preventDefault(); e.stopImmediatePropagation(); editorUndo(); return; }
            if (ctrl && (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z'))) { e.preventDefault(); e.stopImmediatePropagation(); editorRedo(); return; }
            if (ctrl && e.key.toLowerCase() === 'a') { e.preventDefault(); e.stopImmediatePropagation(); editorSelectAll(); return; }
            if (ctrl && e.shiftKey && e.key.toLowerCase() === 'c') { e.preventDefault(); e.stopImmediatePropagation(); editorCopyPosition(); return; }
            if (ctrl && e.shiftKey && e.key.toLowerCase() === 'v') { e.preventDefault(); e.stopImmediatePropagation(); editorPastePosition(); return; }
            if (ctrl && e.key.toLowerCase() === 'l') { e.preventDefault(); e.stopImmediatePropagation(); editorToggleLock(); return; }
            // Tool shortcuts (no modifier) — V = select/arrow, H = hand/pan, ? = shortcut help
            if (!ctrl && !e.shiftKey && !e.altKey) {
                if (e.key.toLowerCase() === 'v') { e.preventDefault(); e.stopImmediatePropagation(); _editorIframe.contentWindow.setMouseNormal?.(); return; }
                if (e.key.toLowerCase() === 'h') { e.preventDefault(); e.stopImmediatePropagation(); _editorIframe.contentWindow.setMouseMove?.(); return; }
                if (e.key === '?') { e.preventDefault(); e.stopImmediatePropagation(); editorToggleShortcutHelp(); return; }
            }
            if (e.key === 'Escape') {
                const _ov = _editorIframe?.contentDocument?.getElementById('inu-shortcut-overlay');
                if (_ov) { _ov.remove(); e.stopImmediatePropagation(); return; }
                editorDeselectAll(); e.stopImmediatePropagation(); return;
            }
            // Alt+T/B/L/R → text position shortcuts (arrow keys conflict with WebPort's native handlers)
            if (e.altKey && !ctrl) {
                const TP = { t:'top-center', b:'bottom-center', l:'left', r:'right' };
                const p = TP[e.key.toLowerCase()];
                if (p) { e.preventDefault(); e.stopImmediatePropagation(); editorSetTextPosition(p); return; }
            }
            const DIRS = { ArrowLeft:[-1,0], ArrowRight:[1,0], ArrowUp:[0,-1], ArrowDown:[0,1] };
            if (!DIRS[e.key]) return;
            e.preventDefault(); e.stopImmediatePropagation();
            const sel = editorGetSelected();
            if (!sel.length) return;
            // Plain/Shift+Arrow → nudge
            const step = e.shiftKey ? 10 : (editorGridEnabled ? editorGridSize : 1);
            const [dx, dy] = DIRS[e.key].map(v => v * step);
            editorPushUndo();
            sel.forEach(el => {
                el.style.left = (parseFloat(el.style.left) || 0) + dx + 'px';
                el.style.top  = (parseFloat(el.style.top)  || 0) + dy + 'px';
                el.style.transform = '';
            });
            // Debounce nudge saves — don't POST on every keypress while holding arrow
            clearTimeout(editorBindKeyboard._nudgeTimer);
            editorBindKeyboard._nudgeTimer = setTimeout(() => editorSavePositions(sel), 600);
            editorUpdateToolbarContext();
        }, { capture: true });
    }

    // ── Inject styles + overlays into iframe ──────────────────────
    function injectEditorStyles(iframeDoc) {
        if (iframeDoc.getElementById('inu-editor-styles')) return;
        const s = iframeDoc.createElement('style');
        s.id = 'inu-editor-styles';
        s.textContent = `
            #inu-editor-status {
                position: absolute; bottom: 0; left: 0; right: 0; height: 20px;
                display: flex; align-items: center; gap: 8px; padding: 0 10px;
                background: rgba(30,30,46,0.88); color: #888; font-size: 11px;
                font-family: monospace; pointer-events: none; z-index: 9999; box-sizing: border-box;
            }
            .inu-sb-sep { color: #444; }
            .inu-locked { outline: 2px dashed #e53935 !important; }
            .inu-lock-badge { position: absolute; top: 2px; right: 2px; font-size: 10px; pointer-events: none; line-height: 1; z-index: 1; }
            /* ── Hover tooltip ── */
            #inu-sym-tooltip {
                display: none; position: fixed; z-index: 99997;
                background: #1e1e2e; color: #cdd6f4; border: 1px solid #45475a;
                border-radius: 6px; padding: 7px 10px; font-size: 11px;
                font-family: monospace; pointer-events: none; line-height: 1.6;
                box-shadow: 0 4px 18px rgba(0,0,0,0.55); max-width: 320px;
            }
            .inu-tip-pre { color: #89b4fa; font-weight: bold; margin-bottom: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
            .inu-tip-row { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
            .inu-tip-row span { color: #a6e3a1; }
            /* ── Batch generate overlay ── */
            #inu-batch-overlay {
                position: fixed; inset: 0; z-index: 99998;
                background: rgba(0,0,0,0.35);
                display: flex; align-items: center; justify-content: center;
            }
            #inu-batch-modal {
                background: #fff; color: #24292f; border-radius: 8px;
                padding: 20px 24px; min-width: 420px; max-width: 520px;
                box-shadow: 0 4px 24px rgba(0,0,0,0.18); font-family: sans-serif; font-size: 12px;
            }
            #inu-batch-modal h3 { margin: 0 0 14px; font-size: 13px; color: #0969da; border-bottom: 1px solid #d0d7de; padding-bottom: 8px; }
            .inu-bg-row { display: flex; flex-direction: column; gap: 3px; margin-bottom: 10px; }
            .inu-bg-inline { display: flex; gap: 12px; }
            .inu-bg-inline .inu-bg-row { flex: 1; }
            .inu-bg-row label { font-size: 10px; color: #57606a; text-transform: uppercase; letter-spacing: 0.5px; }
            .inu-bg-row input { background: #fff; color: #24292f; border: 1px solid #d0d7de; border-radius: 4px; padding: 4px 7px; font-size: 12px; outline: none; width: 100%; box-sizing: border-box; }
            .inu-bg-row input:focus { border-color: #0969da; box-shadow: 0 0 0 2px rgba(9,105,218,0.15); }
            .inu-bg-hint { font-size: 10px; color: #57606a; }
            .inu-bg-preview-list { background: #f6f8fa; border: 1px solid #d0d7de; border-radius: 4px; padding: 7px 10px; max-height: 130px; overflow-y: auto; display: flex; flex-wrap: wrap; gap: 4px; margin-bottom: 12px; }
            .inu-bg-preview-list span { background: #ddf4ff; border: 1px solid #54aeff; border-radius: 3px; padding: 2px 6px; font-size: 10px; font-family: monospace; color: #0550ae; }
            .inu-bg-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 4px; }
            .inu-bg-actions button { padding: 5px 14px; border-radius: 4px; border: 1px solid #d0d7de; background: #f6f8fa; color: #24292f; cursor: pointer; font-size: 12px; }
            .inu-bg-actions button:hover { background: #eaeef2; }
            .inu-bg-primary { background: #0969da !important; color: #fff !important; border-color: #0969da !important; font-weight: bold; }
            .inu-bg-primary:hover { background: #0860ca !important; }
            .inu-bg-primary:disabled { background: #d0d7de !important; color: #8c959f !important; border-color: #d0d7de !important; cursor: not-allowed !important; }
            #inu-bg-status { margin-top: 10px; font-size: 11px; color: #1a7f37; font-family: monospace; }
            .inu-bg-type-row { display: flex; gap: 6px; }
            .inu-bg-type-row input { flex: 1; }
            .inu-bg-type-row button { padding: 4px 10px; border-radius: 4px; border: 1px solid #d0d7de; background: #f6f8fa; color: #24292f; cursor: pointer; white-space: nowrap; }
            .inu-bg-type-row button:hover { background: #eaeef2; }
            .inu-bg-type-preview { min-height: 24px; display: flex; align-items: center; gap: 8px; margin-top: 4px; }
            .inu-bg-type-preview svg, .inu-bg-type-preview img { width: 32px; height: 32px; }
            .inu-bg-type-name { font-size: 11px; color: #0969da; font-family: monospace; }
            /* ── Symbol picker ── */
            #inu-sym-picker {
                position: fixed; inset: 0; z-index: 99999;
                background: rgba(0,0,0,0.4);
                display: flex; align-items: center; justify-content: center;
            }
            #inu-sp-modal {
                background: #fff; color: #24292f; border-radius: 8px;
                width: min(720px, 94vw); max-height: 82vh;
                display: flex; flex-direction: column;
                box-shadow: 0 8px 32px rgba(0,0,0,0.22); overflow: hidden;
            }
            #inu-sp-header {
                display: flex; align-items: center; gap: 8px;
                padding: 10px 14px; border-bottom: 1px solid #d0d7de;
                flex-shrink: 0; background: #f6f8fa;
            }
            #inu-sp-header span { font-size: 13px; color: #24292f; font-weight: 600; white-space: nowrap; }
            #inu-sp-lib { background: #fff; color: #24292f; border: 1px solid #d0d7de; border-radius: 4px; padding: 3px 7px; font-size: 12px; outline: none; cursor: pointer; }
            #inu-sp-lib:focus { border-color: #0969da; }
            #inu-sp-search { flex: 1; background: #fff; color: #24292f; border: 1px solid #d0d7de; border-radius: 4px; padding: 4px 8px; font-size: 12px; outline: none; }
            #inu-sp-search:focus { border-color: #0969da; box-shadow: 0 0 0 2px rgba(9,105,218,0.15); }
            #inu-sp-close { background: none; border: none; color: #57606a; font-size: 18px; cursor: pointer; padding: 0 2px; line-height: 1; }
            #inu-sp-close:hover { color: #24292f; }
            #inu-sp-body { overflow-y: auto; padding: 12px 14px; flex: 1; }
            .inu-sp-group-label { font-size: 10px; color: #57606a; text-transform: uppercase; letter-spacing: 0.8px; margin: 14px 0 6px; border-bottom: 1px solid #eaeef2; padding-bottom: 4px; }
            .inu-sp-grid { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 4px; }
            .inu-sp-item {
                display: flex; flex-direction: column; align-items: center; justify-content: flex-start;
                width: 76px; padding: 8px 4px 6px; border-radius: 6px; cursor: pointer;
                border: 1px solid #d0d7de; background: #f6f8fa; gap: 5px;
                transition: background 0.1s, border-color 0.1s;
            }
            .inu-sp-item:hover { background: #ddf4ff; border-color: #54aeff; }
            .inu-sp-item.inu-sp-selected { background: #ddf4ff; border-color: #0969da; }
            .inu-sp-icon { width: 36px; height: 36px; display: flex; align-items: center; justify-content: center; color: #24292f; }
            .inu-sp-icon svg, .inu-sp-icon img { width: 36px; height: 36px; }
            .inu-sp-label { font-size: 9px; color: #57606a; text-align: center; word-break: break-word; line-height: 1.3; max-width: 70px; }
            /* ── Shortcut overlay ── */
            #inu-shortcut-overlay {
                position: fixed; inset: 0; z-index: 99998;
                background: rgba(0,0,0,0.45);
                display: flex; align-items: center; justify-content: center;
            }
            #inu-shortcut-modal {
                background: #1e1e2e; color: #cdd6f4; border-radius: 8px;
                padding: 20px 28px; min-width: 380px;
                box-shadow: 0 8px 40px rgba(0,0,0,0.7); font-family: monospace;
            }
            #inu-shortcut-modal h3 { margin: 0 0 12px; font-size: 13px; color: #89b4fa; border-bottom: 1px solid #313244; padding-bottom: 8px; }
            #inu-shortcut-modal table { border-collapse: collapse; width: 100%; }
            #inu-shortcut-modal td { padding: 4px 6px; font-size: 11px; vertical-align: middle; border-bottom: 1px solid #1e1e2e; }
            #inu-shortcut-modal td:first-child { white-space: nowrap; padding-right: 16px; }
            #inu-shortcut-modal kbd { display: inline-block; background: #313244; color: #cdd6f4; border: 1px solid #45475a; border-radius: 3px; padding: 1px 5px; font-size: 10px; font-family: monospace; }
            .inu-sc-hint { margin: 10px 0 0; font-size: 10px; color: #6c7086; text-align: right; }
        `;
        iframeDoc.head.appendChild(s);

        const wpp = iframeDoc.getElementById('wpp');
        const ep = iframeDoc.getElementById('wp_editpanel');
        if (ep && !iframeDoc.getElementById('inu-editor-status')) {
            const sb = iframeDoc.createElement('div');
            sb.id = 'inu-editor-status';
            const cw = (wpp && wpp.offsetWidth)  || 1520;
            const ch = (wpp && wpp.offsetHeight) || 850;
            sb.innerHTML = `<span id="inu-sb-size">${cw} × ${ch} px</span><span class="inu-sb-sep">|</span><span id="inu-sb-count">— komponenter</span><span class="inu-sb-sep">|</span><span id="inu-sb-sel">0 markerade</span><span class="inu-sb-sep">|</span><span id="inu-sb-zoom">Zoom: 100%</span><span class="inu-sb-sep" id="inu-sb-pos-sep" style="display:none">|</span><span id="inu-sb-pos-clip" style="display:none">📋 Ctrl+Shift+V klistrar in position</span>`;
            ep.appendChild(sb);
        }
    }

    // ── Inject toolbar into iframe after WebPort's native toolbar ─
    function injectEditorToolbar(iframe) {
        const iDoc = iframe.contentDocument;
        if (iDoc.getElementById('inu-editor-toolbar')) return;

        const ALIGN_BTNS = [
            { id:'al', title:'Vänsterjustera',      fn:()=>editorAlign('x','min') },
            { id:'ac', title:'Centrera horisontalt', fn:()=>editorAlign('x','center') },
            { id:'ar', title:'Högerjustera',         fn:()=>editorAlign('x','max') },
            { id:'at', title:'Toppjustera',          fn:()=>editorAlign('y','min') },
            { id:'am', title:'Centrera vertikalt',   fn:()=>editorAlign('y','center') },
            { id:'ab', title:'Bottenjustera',        fn:()=>editorAlign('y','max') },
            { id:'dh', title:'Fördela horisontalt',  fn:()=>editorDistribute('x') },
            { id:'dv', title:'Fördela vertikalt',    fn:()=>editorDistribute('y') },
        ];
        const SVG = {
            al: '<rect x="3" y="4" width="12" height="3"/><rect x="3" y="9" width="8" height="3"/><rect x="3" y="14" width="10" height="3"/><line x1="3" y1="3" x2="3" y2="18" stroke="currentColor" stroke-width="1.5"/>',
            ac: '<rect x="5" y="4" width="10" height="3"/><rect x="7" y="9" width="6" height="3"/><rect x="4" y="14" width="12" height="3"/><line x1="10" y1="2" x2="10" y2="19" stroke="currentColor" stroke-width="1.5"/>',
            ar: '<rect x="5" y="4" width="12" height="3"/><rect x="9" y="9" width="8" height="3"/><rect x="7" y="14" width="10" height="3"/><line x1="17" y1="3" x2="17" y2="18" stroke="currentColor" stroke-width="1.5"/>',
            at: '<rect x="4" y="5" width="3" height="12"/><rect x="9" y="5" width="3" height="8"/><rect x="14" y="5" width="3" height="10"/><line x1="3" y1="5" x2="18" y2="5" stroke="currentColor" stroke-width="1.5"/>',
            am: '<rect x="4" y="5" width="3" height="10"/><rect x="9" y="7" width="3" height="6"/><rect x="14" y="4" width="3" height="12"/><line x1="3" y1="10" x2="18" y2="10" stroke="currentColor" stroke-width="1.5"/>',
            ab: '<rect x="4" y="4" width="3" height="12"/><rect x="9" y="8" width="3" height="8"/><rect x="14" y="6" width="3" height="10"/><line x1="3" y1="16" x2="18" y2="16" stroke="currentColor" stroke-width="1.5"/>',
            dh: '<rect x="2" y="4" width="3" height="12"/><rect x="17" y="4" width="3" height="12"/><rect x="8" y="7" width="4" height="6"/><line x1="5" y1="10" x2="8" y2="10" stroke="currentColor" stroke-width="1.5"/><line x1="12" y1="10" x2="15" y2="10" stroke="currentColor" stroke-width="1.5"/>',
            dv: '<rect x="4" y="2" width="12" height="3"/><rect x="4" y="17" width="12" height="3"/><rect x="7" y="8" width="6" height="4"/><line x1="10" y1="5" x2="10" y2="8" stroke="currentColor" stroke-width="1.5"/><line x1="10" y1="12" x2="10" y2="15" stroke="currentColor" stroke-width="1.5"/>',
            undo: '<text x="10" y="14" text-anchor="middle" font-size="16" fill="currentColor" style="font-family:sans-serif">↺</text>',
            redo: '<text x="10" y="14" text-anchor="middle" font-size="16" fill="currentColor" style="font-family:sans-serif">↻</text>',
            front: '<polyline points="5,10 10,5 15,10" stroke="currentColor" stroke-width="1.8" fill="none"/><polyline points="5,14 10,9 15,14" stroke="currentColor" stroke-width="1.8" fill="none"/><line x1="3" y1="17" x2="17" y2="17" stroke="currentColor" stroke-width="1.5"/>',
            fwd:   '<polyline points="5,12 10,7 15,12" stroke="currentColor" stroke-width="1.8" fill="none"/><line x1="3" y1="15" x2="17" y2="15" stroke="currentColor" stroke-width="1.5"/>',
            bwd:   '<polyline points="5,8 10,13 15,8" stroke="currentColor" stroke-width="1.8" fill="none"/><line x1="3" y1="5" x2="17" y2="5" stroke="currentColor" stroke-width="1.5"/>',
            back:  '<polyline points="5,10 10,15 15,10" stroke="currentColor" stroke-width="1.8" fill="none"/><polyline points="5,6 10,11 15,6" stroke="currentColor" stroke-width="1.8" fill="none"/><line x1="3" y1="3" x2="17" y2="3" stroke="currentColor" stroke-width="1.5"/>',
            mw: '<rect x="2" y="7" width="16" height="6" rx="1" fill="none" stroke="currentColor" stroke-width="1.5"/><rect x="5" y="4" width="10" height="12" rx="1" fill="none" stroke="currentColor" stroke-width="1.5" stroke-dasharray="2,1"/><line x1="2" y1="2" x2="2" y2="18" stroke="currentColor" stroke-width="1.5"/><line x1="18" y1="2" x2="18" y2="18" stroke="currentColor" stroke-width="1.5"/>',
            mh: '<rect x="7" y="2" width="6" height="16" rx="1" fill="none" stroke="currentColor" stroke-width="1.5"/><rect x="4" y="5" width="12" height="10" rx="1" fill="none" stroke="currentColor" stroke-width="1.5" stroke-dasharray="2,1"/><line x1="2" y1="2" x2="18" y2="2" stroke="currentColor" stroke-width="1.5"/><line x1="2" y1="18" x2="18" y2="18" stroke="currentColor" stroke-width="1.5"/>',
            'sort-az': '<text x="1" y="9" font-size="8" font-family="sans-serif" fill="currentColor" font-weight="bold">A</text><text x="1" y="18" font-size="8" font-family="sans-serif" fill="currentColor" font-weight="bold">Z</text><line x1="13" y1="3" x2="13" y2="15" stroke="currentColor" stroke-width="1.5"/><polyline points="10,12 13,15 16,12" stroke="currentColor" stroke-width="1.5" fill="none"/>',
            'sort-za': '<text x="1" y="9" font-size="8" font-family="sans-serif" fill="currentColor" font-weight="bold">Z</text><text x="1" y="18" font-size="8" font-family="sans-serif" fill="currentColor" font-weight="bold">A</text><line x1="13" y1="5" x2="13" y2="17" stroke="currentColor" stroke-width="1.5"/><polyline points="10,8 13,5 16,8" stroke="currentColor" stroke-width="1.5" fill="none"/>',
            'gen-sym': '<rect x="2" y="2" width="7" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="1.5"/><rect x="11" y="2" width="7" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="1.5"/><rect x="2" y="11" width="7" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="1.5"/><rect x="11" y="11" width="7" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="1.5" stroke-dasharray="2,1.5"/><line x1="14.5" y1="8.5" x2="14.5" y2="10.5" stroke="currentColor" stroke-width="1.5"/><line x1="13.5" y1="9.5" x2="15.5" y2="9.5" stroke="currentColor" stroke-width="1.5"/>',
        };
        const mk = k => `<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20" fill="currentColor">${SVG[k]}</svg>`;
        const bar = document.createElement('div');
        bar.id = 'inu-editor-toolbar';
        bar.innerHTML = `
            <span class="inu-et-group">
                <button id="inu-et-undo" title="Ångra (Ctrl+Z)" disabled>&lt;</button>
                <button id="inu-et-redo" title="Gör om (Ctrl+Y)" disabled>&gt;</button>
            </span>
            <span class="inu-et-group">
                <button class="inu-et-ic" id="inu-et-gen-btn" title="Generera symboler…">${mk('gen-sym')}</button>
            </span>
            <span class="inu-et-group">
                <label class="inu-et-label"><input type="checkbox" id="inu-grid-snap"${editorGridEnabled?' checked':''}><span>Rutnät</span></label>
                <select id="inu-grid-size">${EDITOR_GRID_SIZES.map(s=>`<option value="${s}"${s===editorGridSize?' selected':''}>${s} px</option>`).join('')}</select>
                <span class="inu-et-muted" id="inu-et-nudge-hint">Nudge: ${editorGridEnabled ? editorGridSize : 1} px</span>
                <button id="inu-grid-snap-all" title="Fäst markerade till rutnät"${editorGridEnabled?'':' style="display:none"'}>Fäst markerade</button>
            </span>
            <span class="inu-et-group">
                <span id="inu-et-sel-count" class="inu-et-muted">0 markerade</span>
            </span>
            <span class="inu-et-group" id="inu-et-align-group" style="display:none">
                ${ALIGN_BTNS.map(b=>`<button class="inu-et-ic" data-align="${b.id}" title="${b.title}">${mk(b.id)}</button>`).join('')}
            </span>
            <span class="inu-et-group" id="inu-et-sort-group" style="display:none">
                <button class="inu-et-ic" id="inu-et-sort-az-btn" title="Sortera markerade A–Ö (behåller positioner)">${mk('sort-az')}</button>
                <button class="inu-et-ic" id="inu-et-sort-za-btn" title="Sortera markerade Ö–A (behåller positioner)">${mk('sort-za')}</button>
            </span>
            <span class="inu-et-group" id="inu-et-step-dist-group" style="display:none">
                <span class="inu-et-muted">↕</span>
                <input type="number" id="inu-et-step-val" value="20" min="0" title="Steg i px" style="width:42px;padding:1px 3px;border:1px solid #bbb;border-radius:3px;font-size:11px;text-align:right">
                <span class="inu-et-muted" style="margin-left:1px">px</span>
                <button id="inu-et-step-dist-btn" title="Fördela markerade med fast vertikalt steg (toppobjektet fixerat)">Fördela</button>
            </span>
            <span class="inu-et-group" id="inu-et-pos-group" style="display:none">
                <span class="inu-et-muted">X</span><input type="number" id="inu-et-pos-x" class="inu-et-pos-input">
                <span class="inu-et-muted">Y</span><input type="number" id="inu-et-pos-y" class="inu-et-pos-input">
                <span class="inu-et-muted">B <span id="inu-et-pos-w">—</span></span>
                <span class="inu-et-muted">H <span id="inu-et-pos-h">—</span></span>
            </span>
            <span class="inu-et-group" id="inu-et-textpos-group" style="display:none">
                <span class="inu-et-muted">Text</span>
                <select id="inu-et-textpos-select" title="Textposition (Alt+T/B/L/R för snabbval)">
                    ${TEXT_POSITIONS.map(p => `<option value="${p.id}">${p.label}</option>`).join('')}
                </select>
            </span>
            <span class="inu-et-group" id="inu-et-lock-group" style="display:none">
                <button id="inu-et-lock-btn" title="Lås/lås upp markerade (Ctrl+L)">🔒</button>
            </span>
            <span class="inu-et-group" id="inu-et-unlock-all-group" style="display:none">
                <button id="inu-et-unlock-all-btn" title="Lås upp alla låsta objekt">🔓 Lås upp alla</button>
            </span>
        `;
        const iContent = iDoc.getElementById('content');
        const nativeNav = iContent?.querySelector(':scope > nav, :scope > .toolbar-container');
        if (nativeNav) nativeNav.after(bar);
        else if (iContent) iContent.prepend(bar);
        else iframe.parentElement.insertBefore(bar, iframe);

        iDoc.getElementById('inu-grid-snap').addEventListener('change', e => {
            editorGridEnabled = e.target.checked;
            GM_setValue('inu_editor_grid_enabled', editorGridEnabled);
            iDoc.getElementById('inu-grid-snap-all').style.display = editorGridEnabled ? '' : 'none';
            editorApplyGrid();
            editorUpdateNudgeHint();
        });
        iDoc.getElementById('inu-grid-size').addEventListener('change', e => {
            editorGridSize = parseInt(e.target.value);
            GM_setValue('inu_editor_grid_size', editorGridSize);
            if (editorGridEnabled) editorApplyGrid();
            editorUpdateNudgeHint();
        });
        iDoc.getElementById('inu-grid-snap-all').addEventListener('click', editorSnapAll);
        bar.querySelectorAll('[data-align]').forEach(btn => {
            btn.addEventListener('click', () => ALIGN_BTNS.find(x => x.id === btn.dataset.align)?.fn());
        });
        iDoc.getElementById('inu-et-undo').addEventListener('click', editorUndo);
        iDoc.getElementById('inu-et-redo').addEventListener('click', editorRedo);
        ['inu-et-pos-x', 'inu-et-pos-y'].forEach(id => {
            iDoc.getElementById(id).addEventListener('keydown', e => {
                if (e.key !== 'Enter') return;
                const sel = editorGetSelected();
                if (sel.length !== 1) return;
                const v = parseFloat(e.target.value);
                if (isNaN(v)) return;
                editorPushUndo();
                if (id.endsWith('-x')) sel[0].style.left = v + 'px';
                else                    sel[0].style.top  = v + 'px';
                sel[0].style.transform = '';
                editorSavePositions(sel);
                editorUpdatePositionInspector(sel[0]);
                _editorIframe.contentWindow.focus();
            });
            iDoc.getElementById(id).addEventListener('blur', () => {
                setTimeout(() => _editorIframe && _editorIframe.contentWindow.focus(), 50);
            });
        });

        // Step distribution
        const stepDistBtn = iDoc.getElementById('inu-et-step-dist-btn');
        const stepValInp  = iDoc.getElementById('inu-et-step-val');
        stepDistBtn.addEventListener('click', () => {
            editorDistributeByStep(parseFloat(stepValInp.value) || 20);
            _editorIframe.contentWindow.focus();
        });
        stepValInp.addEventListener('keydown', e => {
            if (e.key === 'Enter') { editorDistributeByStep(parseFloat(stepValInp.value) || 20); _editorIframe.contentWindow.focus(); }
        });

        // Text position dropdown
        iDoc.getElementById('inu-et-textpos-select').addEventListener('change', e => {
            editorSetTextPosition(e.target.value);
            _editorIframe.contentWindow.focus();
        });

        // Generate button
        iDoc.getElementById('inu-et-gen-btn').addEventListener('click', editorShowBatchGenerate);

        // Sort buttons
        iDoc.getElementById('inu-et-sort-az-btn').addEventListener('click', () => editorSortByName('asc'));
        iDoc.getElementById('inu-et-sort-za-btn').addEventListener('click', () => editorSortByName('desc'));

        // Lock button
        iDoc.getElementById('inu-et-lock-btn').addEventListener('click', editorToggleLock);
        iDoc.getElementById('inu-et-unlock-all-btn').addEventListener('click', editorUnlockAll);

        const s = iDoc.createElement('style');
        s.textContent = `
            #inu-editor-toolbar {
                display: flex; align-items: center; gap: 0; flex-wrap: wrap;
                padding: 4px 10px; background: #f5f5f8; border-bottom: 1px solid #ddd;
                font-size: 11px; color: #333; min-height: 30px;
            }
            .inu-et-group { display: inline-flex; align-items: center; gap: 4px; flex-wrap: nowrap; border-right: 1px solid #ccc; padding-right: 10px; margin-right: 6px; }
            .inu-et-muted { font-size: 11px; color: #666; white-space: nowrap; }
            #inu-editor-toolbar select, #inu-editor-toolbar button {
                font-size: 11px; padding: 2px 7px; border-radius: 4px;
                border: none; background: #5b6abf; color: #fff; cursor: pointer; font-weight: 600; line-height: 1.4;
            }
            #inu-editor-toolbar button:hover:not(:disabled) { background: #4a58a8; }
            #inu-editor-toolbar button:disabled { opacity: 0.35; cursor: default; }
            #inu-editor-toolbar select { background: #fff; color: #333; border: 1px solid #ccc; font-weight: normal; padding: 2px 5px; }
            .inu-et-label { display: inline-flex; align-items: center; gap: 3px; cursor: pointer; white-space: nowrap; color: #333; }
            #inu-editor-toolbar input[type=checkbox] { width: 13px; height: 13px; margin: 0; cursor: pointer; accent-color: #5b6abf; flex-shrink: 0; }
            .inu-et-ic { padding: 2px 4px !important; }
            .inu-et-pos-input {
                width: 52px; font-variant-numeric: tabular-nums; font-size: 11px;
                padding: 1px 4px; border-radius: 4px; border: 1px solid #ccc;
                background: #fff; color: #333;
            }
        `;
        iDoc.head.appendChild(s);
    }

    // ── Symbol type picker ────────────────────────────────────────
    const _SP_NS_LABELS = { 'inu-svg': 'INU SVG', 'tools': 'Verktyg', 'portlet-lib': 'Portlets', 'map-icons': 'Kartikoner' };
    const _spNsLabel = ns => _SP_NS_LABELS[ns] || ns.split('-').map(w => w[0].toUpperCase() + w.slice(1)).join(' ');

    function editorUpdateTypePreview(iDoc, objType) {
        const prev = iDoc.getElementById('inu-bg-type-preview');
        if (!prev) return;
        prev.innerHTML = '';
        if (!objType) return;
        const libEl = iDoc.getElementById(objType);
        const icon  = iDoc.createElement('div');
        icon.className = 'inu-sp-icon';
        const svg = libEl?.querySelector('svg')?.cloneNode(true);
        const img = libEl?.querySelector('img')?.cloneNode(true);
        if (svg) { svg.setAttribute('width', '32'); svg.setAttribute('height', '32'); icon.appendChild(svg); }
        else if (img) { img.style.cssText = 'width:32px;height:32px;object-fit:contain'; icon.appendChild(img); }
        if (svg || img) prev.appendChild(icon);
        const ns   = objType.includes('_') ? objType.substring(0, objType.indexOf('_')) : '';
        const name = ns ? objType.slice(ns.length + 1) : objType;
        const span = iDoc.createElement('span');
        span.className = 'inu-bg-type-name';
        span.textContent = name;
        prev.appendChild(span);

        // Update step Y default from the libitem's rendered height (2x gives safe non-overlapping spacing)
        const stepEl = iDoc.getElementById('inu-bg-step');
        if (stepEl && libEl) {
            const svgH = parseFloat(libEl.querySelector('svg')?.getAttribute('height')) || 0;
            const h = svgH || libEl.offsetHeight || 0;
            if (h > 0) stepEl.value = Math.ceil(h) + 5;
        }
    }

    function editorShowSymbolPicker(iDoc, onSelect) {
        if (iDoc.getElementById('inu-sym-picker')) return;

        // Discover all namespaces dynamically
        const allItems = [...iDoc.querySelectorAll('.libitem[id]')];
        const nsMap = {}; // ns → [{id, name, label}]
        allItems.forEach(el => {
            const id = el.id;
            const uscore = id.indexOf('_');
            if (uscore < 0) return;
            const ns   = id.substring(0, uscore);
            const name = id.slice(uscore + 1);
            if (!nsMap[ns]) nsMap[ns] = [];
            nsMap[ns].push({ id, name });
        });

        const namespaces = Object.keys(nsMap).sort((a, b) => {
            // INU SVG first, rest alphabetically
            if (a === 'inu-svg') return -1;
            if (b === 'inu-svg') return 1;
            return a.localeCompare(b);
        });
        if (!namespaces.length) return;

        const getGroup  = name => name.match(/^([A-Z][a-z]+)/)?.[1] ?? name.substring(0, 4);
        const toLabel   = name => name.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2');

        // Build picker element
        const picker = iDoc.createElement('div');
        picker.id = 'inu-sym-picker';

        const modal = iDoc.createElement('div');
        modal.id = 'inu-sp-modal';

        // Header
        const header = iDoc.createElement('div');
        header.id = 'inu-sp-header';

        const title = iDoc.createElement('span');
        title.textContent = 'Välj symboltyp';

        const libSel = iDoc.createElement('select');
        libSel.id = 'inu-sp-lib';
        namespaces.forEach(ns => {
            const opt = iDoc.createElement('option');
            opt.value = ns; opt.textContent = _spNsLabel(ns);
            libSel.appendChild(opt);
        });

        const search = iDoc.createElement('input');
        search.type = 'text'; search.id = 'inu-sp-search'; search.placeholder = 'Sök…';

        const closeBtn = iDoc.createElement('button');
        closeBtn.id = 'inu-sp-close'; closeBtn.textContent = '✕';

        header.append(title, libSel, search, closeBtn);

        // Body
        const bodyDiv = iDoc.createElement('div');
        bodyDiv.id = 'inu-sp-body';

        const buildGrid = ns => {
            bodyDiv.innerHTML = '';
            const items = nsMap[ns] || [];
            const groups = {};
            items.forEach(sym => {
                const grp = getGroup(sym.name);
                if (!groups[grp]) groups[grp] = [];
                groups[grp].push(sym);
            });

            Object.keys(groups).sort().forEach(grp => {
                const groupDiv = iDoc.createElement('div');
                groupDiv.className = 'inu-sp-group';

                const hdr = iDoc.createElement('div');
                hdr.className = 'inu-sp-group-label';
                hdr.textContent = grp;

                const grid = iDoc.createElement('div');
                grid.className = 'inu-sp-grid';

                groups[grp].forEach(sym => {
                    const item = iDoc.createElement('div');
                    item.className = 'inu-sp-item';
                    item.dataset.id   = sym.id;
                    item.dataset.name = sym.name.toLowerCase();
                    item.title = sym.id;

                    const iconDiv = iDoc.createElement('div');
                    iconDiv.className = 'inu-sp-icon';
                    const srcEl = iDoc.getElementById(sym.id);
                    const svg   = srcEl?.querySelector('svg')?.cloneNode(true);
                    const img   = srcEl?.querySelector('img')?.cloneNode(true);
                    if (svg) { svg.setAttribute('width', '36'); svg.setAttribute('height', '36'); iconDiv.appendChild(svg); }
                    else if (img) { img.style.cssText = 'width:36px;height:36px;object-fit:contain'; iconDiv.appendChild(img); }
                    else { iconDiv.style.cssText = 'font-size:20px;color:#8c959f'; iconDiv.textContent = '□'; }

                    const labelDiv = iDoc.createElement('div');
                    labelDiv.className = 'inu-sp-label';
                    labelDiv.textContent = toLabel(sym.name).replace(grp + ' ', '');

                    item.append(iconDiv, labelDiv);
                    grid.appendChild(item);
                });

                groupDiv.append(hdr, grid);
                bodyDiv.appendChild(groupDiv);
            });
            search.value = '';
        };

        buildGrid(namespaces[0]);
        modal.append(header, bodyDiv);
        picker.appendChild(modal);
        iDoc.body.appendChild(picker);

        // Library switch
        libSel.addEventListener('change', () => buildGrid(libSel.value));

        // Search filter
        search.addEventListener('input', () => {
            const q = search.value.toLowerCase();
            bodyDiv.querySelectorAll('.inu-sp-item').forEach(el => {
                el.style.display = !q || el.dataset.name.includes(q) ? '' : 'none';
            });
            bodyDiv.querySelectorAll('.inu-sp-group').forEach(g => {
                g.style.display = [...g.querySelectorAll('.inu-sp-item')].some(i => i.style.display !== 'none') ? '' : 'none';
            });
        });

        // Select
        bodyDiv.addEventListener('click', e => {
            const item = e.target.closest('.inu-sp-item');
            if (!item) return;
            onSelect(item.dataset.id);
            picker.remove();
        });

        closeBtn.addEventListener('click', () => picker.remove());
        picker.addEventListener('click', e => { if (e.target === picker) picker.remove(); });
        search.focus();
    }

    // ── Batch symbol generation ───────────────────────────────────
    function _expandPattern(pattern, n) {
        return pattern.replace(/#+/, m => String(n).padStart(m.length, '0'));
    }

    async function editorShowBatchGenerate() {
        const iDoc = _editorIframe?.contentDocument;
        if (!iDoc) return;
        // Toggle off if already open
        if (iDoc.getElementById('inu-batch-overlay')) { iDoc.getElementById('inu-batch-overlay').remove(); return; }

        const pageid = new URLSearchParams(location.search).get('pageid');
        const sel = editorGetSelected();
        let objType = '', startX = 50, startY = 50, stepDefault = 30, patternDefault = '';

        if (sel.length >= 1) {
            const ref = sel[0];
            startX = parseFloat(ref.style.left) || 50;
            startY = parseFloat(ref.style.top)  || 50;
            const selectBox = ref.querySelector('.wpCompSelectBox');
            const h = Math.ceil(parseFloat(selectBox?.style.height) || selectBox?.offsetHeight || 0);
            if (h > 0) stepDefault = h + 4; // symbol height + small gap
            try {
                const res = await fetch(`/Page/PageObjectProperties?pageid=${encodeURIComponent(pageid)}&id=${encodeURIComponent(ref.id)}`);
                const html = await res.text();
                const tmp = document.createElement('div');
                tmp.innerHTML = html;
                const form = tmp.querySelector('#poproperties');
                if (form) {
                    objType = form.elements['libobjectpath']?.value || '';
                    const existId = form.elements['id']?.value || '';
                    // Strip trailing _NNN suffix to suggest a pattern base
                    patternDefault = existId.replace(/_\d+$/, '_##');
                    if (patternDefault === existId) patternDefault = existId + '_##';
                }
            } catch(e) {}
        }

        const overlay = iDoc.createElement('div');
        overlay.id = 'inu-batch-overlay';
        overlay.innerHTML = `
            <div id="inu-batch-modal">
                <h3>Generera symboler</h3>
                <div class="inu-bg-row">
                    <label>Symboltyp</label>
                    <div class="inu-bg-type-row">
                        <input type="text" id="inu-bg-type" value="${objType}" placeholder="Välj symbol eller skriv typ…">
                        <button id="inu-bg-pick-btn">Välj…</button>
                    </div>
                    <div id="inu-bg-type-preview" class="inu-bg-type-preview">${objType ? `<span class="inu-bg-type-name">${objType.replace('inu-svg_','')}</span>` : ''}</div>
                </div>
                <div class="inu-bg-row">
                    <label>Namnmönster — använd # för siffror</label>
                    <input type="text" id="inu-bg-pattern" value="${patternDefault}" placeholder="t.ex. HUS01_AS01_LB01_BGS101_###">
                    <span class="inu-bg-hint">Exempel: BGS101_### → BGS101_001, BGS101_002 …</span>
                </div>
                <div class="inu-bg-row">
                    <label>Symboltext (valfri, ## = sekvensnummer)</label>
                    <input type="text" id="inu-bg-text" placeholder="t.ex. BGS101-## (lämna tomt = inget textlabel)">
                </div>
                <div class="inu-bg-inline">
                    <div class="inu-bg-row">
                        <label>Starttal</label>
                        <input type="number" id="inu-bg-start" value="1" min="1">
                    </div>
                    <div class="inu-bg-row">
                        <label>Antal (max 50)</label>
                        <input type="number" id="inu-bg-count" value="10" min="1" max="50">
                    </div>
                </div>
                <div class="inu-bg-inline">
                    <div class="inu-bg-row">
                        <label>Startpos X</label>
                        <input type="number" id="inu-bg-x" value="${startX}">
                    </div>
                    <div class="inu-bg-row">
                        <label>Startpos Y</label>
                        <input type="number" id="inu-bg-y" value="${startY}">
                    </div>
                    <div class="inu-bg-row">
                        <label>Steg Y (px)</label>
                        <input type="number" id="inu-bg-step" value="${stepDefault}">
                    </div>
                </div>
                <div class="inu-bg-hint" style="margin:8px 0 4px">Förhandsvisning:</div>
                <div id="inu-bg-preview" class="inu-bg-preview-list"></div>
                <div class="inu-bg-actions">
                    <button id="inu-bg-cancel">Avbryt</button>
                    <button id="inu-bg-generate" class="inu-bg-primary">Generera</button>
                </div>
                <div id="inu-bg-status" style="display:none"></div>
            </div>`;

        const updatePreview = () => {
            const pat   = iDoc.getElementById('inu-bg-pattern').value;
            const start = parseInt(iDoc.getElementById('inu-bg-start').value) || 1;
            const count = Math.min(parseInt(iDoc.getElementById('inu-bg-count').value) || 10, 50);
            const prev  = iDoc.getElementById('inu-bg-preview');
            if (!pat.includes('#')) { prev.innerHTML = '<span style="color:#f38ba8">Mönstret måste innehålla minst ett #</span>'; return; }
            prev.innerHTML = Array.from({length: count}, (_, i) => `<span>${_expandPattern(pat, start + i)}</span>`).join('');
        };

        overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
        iDoc.body.appendChild(overlay);

        ['inu-bg-pattern','inu-bg-start','inu-bg-count'].forEach(id =>
            iDoc.getElementById(id).addEventListener('input', updatePreview));
        updatePreview();

        // Keep type preview in sync with manual edits
        iDoc.getElementById('inu-bg-type').addEventListener('input', e =>
            editorUpdateTypePreview(iDoc, e.target.value.trim()));
        if (objType) editorUpdateTypePreview(iDoc, objType);

        iDoc.getElementById('inu-bg-cancel').addEventListener('click', () => overlay.remove());
        iDoc.getElementById('inu-bg-generate').addEventListener('click', () => editorRunBatchGenerate(overlay));
        iDoc.getElementById('inu-bg-pick-btn').addEventListener('click', () => {
            editorShowSymbolPicker(iDoc, (id) => {
                iDoc.getElementById('inu-bg-type').value = id;
                editorUpdateTypePreview(iDoc, id);
            });
        });

        iDoc.getElementById('inu-bg-pattern').focus();
        iDoc.getElementById('inu-bg-pattern').select();
    }

    async function editorRunBatchGenerate(overlay) {
        const iDoc   = _editorIframe?.contentDocument;
        const pageid = new URLSearchParams(location.search).get('pageid');

        const objType = iDoc.getElementById('inu-bg-type').value.trim();
        const pattern = iDoc.getElementById('inu-bg-pattern').value.trim();
        const start   = parseInt(iDoc.getElementById('inu-bg-start').value) || 1;
        const count   = Math.min(parseInt(iDoc.getElementById('inu-bg-count').value) || 10, 50);
        const px      = parseInt(iDoc.getElementById('inu-bg-x').value) || 50;
        const py      = parseInt(iDoc.getElementById('inu-bg-y').value) || 50;
        const step    = parseInt(iDoc.getElementById('inu-bg-step').value) ?? 30;

        if (!objType || !pattern || !pattern.includes('#')) return;

        const maxZ = Math.max(0, ...[...iDoc.querySelectorAll('.wpCompObject')]
            .map(e => parseInt(e.style.zIndex) || 0));

        const status = iDoc.getElementById('inu-bg-status');
        const genBtn = iDoc.getElementById('inu-bg-generate');
        genBtn.disabled = true;
        status.style.display = 'block';

        const textPattern = iDoc.getElementById('inu-bg-text')?.value?.trim() || '';
        const encodeWPId = s => s.replace(/\./g, '-2E-').replace(/_/g, '-5F-');

        let created = 0;
        for (let i = 0; i < count; i++) {
            const name = _expandPattern(pattern, start + i);
            status.textContent = `Skapar ${i + 1}/${count}: ${name}…`;
            try {
                const actionType = objType.endsWith('_portlet') ? 'portlet' : 'pageobject';
                const getRes = await fetch(
                    `/page/ActionAdd?show=0&type=${actionType}` +
                    `&objecttype=${encodeURIComponent(objType)}` +
                    `&pageid=${encodeURIComponent(pageid)}` +
                    `&posx=${px}&posy=${py + i * step}&posz=${maxZ + 1 + i}`
                );
                const html = await getRes.text();
                const tmp  = document.createElement('div');
                tmp.innerHTML = html;
                const form = tmp.querySelector('#frmpage');
                if (!form) throw new Error('Inget formulär i svar');

                const fd = new FormData();
                for (const f of form.elements) {
                    if (f.name) fd.set(f.name, f.value);
                }
                fd.set('name', name);
                await fetch('/page/actionadd', { method: 'POST', body: fd });

                // Fix prefix: ActionAdd sets prefix=_name which causes WebPort to prepend
                // the page station prefix. We must update properties immediately after
                // creation to store the full BMS path directly (no leading underscore).
                const elementId = encodeWPId(pageid) + '-2E-' + encodeWPId(name);
                const propsRes = await fetch(
                    `/Page/PageObjectProperties?pageid=${encodeURIComponent(pageid)}&id=${encodeURIComponent(elementId)}`
                );
                const propsHtml = await propsRes.text();
                const propsDiv = document.createElement('div');
                propsDiv.innerHTML = propsHtml;
                const propsForm = propsDiv.querySelector('#poproperties');
                if (propsForm) {
                    const fd2 = new FormData();
                    for (const f of propsForm.elements) {
                        if (!f.name) continue;
                        if ((f.type === 'checkbox' || f.type === 'radio') && !f.checked) continue;
                        fd2.append(f.name, f.value);
                    }
                    // Set prefix to full BMS path without leading underscore
                    fd2.set('prefix', name);
                    // Set display text if a text pattern was provided
                    if (textPattern) fd2.set('name', _expandPattern(textPattern, start + i));
                    await fetch('/page/UpdatePageObjectProperties', { method: 'POST', body: fd2 });
                }

                created++;
            } catch(e) {
                status.textContent = `Fel vid ${name}: ${e.message}`;
                await new Promise(r => setTimeout(r, 2000));
            }
        }

        status.textContent = `✓ ${created} av ${count} symboler skapade. Sidan laddas om…`;
        setTimeout(() => { overlay.remove(); _editorIframe.contentWindow.location.reload(); }, 1200);
    }

    function editorToggleShortcutHelp() {
        const iDoc = _editorIframe?.contentDocument;
        if (!iDoc) return;
        const existing = iDoc.getElementById('inu-shortcut-overlay');
        if (existing) { existing.remove(); return; }
        const overlay = iDoc.createElement('div');
        overlay.id = 'inu-shortcut-overlay';
        overlay.innerHTML = `
            <div id="inu-shortcut-modal">
                <h3>Tangentbordsgenvägar</h3>
                <table>
                    <tr><td><kbd>&lt;</kbd> / <kbd>&gt;</kbd></td><td>Ångra / Gör om</td></tr>
                    <tr><td><kbd>Ctrl+A</kbd></td><td>Markera alla</td></tr>
                    <tr><td><kbd>Escape</kbd></td><td>Avmarkera alla</td></tr>
                    <tr><td><kbd>Ctrl+L</kbd></td><td>Lås / lås upp markerade</td></tr>
                    <tr><td><kbd>Ctrl+Shift+C</kbd> / <kbd>V</kbd></td><td>Kopiera / klistra in position</td></tr>
                    <tr><td><kbd>Alt+T/B/L/R</kbd></td><td>Textposition topp / botten / vänster / höger</td></tr>
                    <tr><td><kbd>↑↓←→</kbd></td><td>Flytta 1 px (eller rastersteg)</td></tr>
                    <tr><td><kbd>Shift+↑↓←→</kbd></td><td>Flytta 10 px</td></tr>
                    <tr><td><kbd>V</kbd></td><td>Markeringsverktyg</td></tr>
                    <tr><td><kbd>H</kbd></td><td>Panoreringsverktyg</td></tr>
                    <tr><td>Mittknapp + dra</td><td>Panorera canvas</td></tr>
                </table>
                <p class="inu-sc-hint">Stäng: <kbd>?</kbd> eller <kbd>Escape</kbd></p>
            </div>`;
        overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
        iDoc.body.appendChild(overlay);
    }

    // ── Hover tooltip ─────────────────────────────────────────────
    function editorInitTooltip(iDoc, panel) {
        if (iDoc.getElementById('inu-sym-tooltip')) return;
        const tip = iDoc.createElement('div');
        tip.id = 'inu-sym-tooltip';
        iDoc.body.appendChild(tip);
        const $ = _editorIframe.contentWindow.jQuery;
        const show = (el, cx, cy) => {
            const prefix = el.dataset.prefix || el.id || '—';
            const label  = el.querySelector('.wpCompText')?.textContent?.trim() || '—';
            const tpId   = editorGetTextPosition(el);
            const tpLbl  = tpId ? (TEXT_POSITIONS.find(p => p.id === tpId)?.label ?? tpId) : '—';
            tip.innerHTML = `<div class="inu-tip-pre">${prefix}</div>` +
                `<div class="inu-tip-row">Text: <span>${label}</span></div>` +
                `<div class="inu-tip-row">Textpos: <span>${tpLbl}</span></div>`;
            tip.style.display = 'block';
            const vw = iDoc.defaultView.innerWidth, vh = iDoc.defaultView.innerHeight;
            const tw = tip.offsetWidth, th = tip.offsetHeight;
            tip.style.left = (cx + 14 + tw > vw - 8 ? cx - tw - 10 : cx + 14) + 'px';
            tip.style.top  = (cy + 14 + th > vh - 8 ? cy - th - 10 : cy + 14) + 'px';
        };
        const hide = () => tip.style.display = 'none';
        $(panel).on('mouseenter.inuTip', '.wpCompObject', function(e) { show(this, e.clientX, e.clientY); });
        $(panel).on('mousemove.inuTip',  '.wpCompObject', function(e) { if (tip.style.display !== 'none') show(this, e.clientX, e.clientY); });
        $(panel).on('mouseleave.inuTip dragstart.inuTip', '.wpCompObject', hide);
    }

    function initPageEditor() {
        console.log(CFG.logPrefix, 'v' + CFG.version, 'Activating (Page Editor)');
        const iframe = editorGetIframe();
        if (!iframe) return;
        _editorIframe = iframe;

        const setup = () => {
            const iDoc = iframe.contentDocument;
            if (!iDoc || !iDoc.querySelector('.wpCompObject')) return;
            // Guard: bail if we already injected into this document instance
            if (iDoc.getElementById('inu-editor-toolbar')) return;
            const $ = iframe.contentWindow.jQuery;

            injectEditorStyles(iDoc);
            injectEditorToolbar(iframe);

            editorApplyGrid();
            editorLoadLocks();
            editorApplyAllLocks();
            editorUpdateUndoButtons();
            editorUpdateZoom();

            const wpp = iDoc.getElementById('wpp');
            new MutationObserver(muts => {
                for (const m of muts) {
                    if (m.type === 'attributes' && m.attributeName === 'class') {
                        const el = m.target;
                        if (!el.classList.contains('wpCompObject') || el.classList.contains('ui-draggable-disabled')) continue;
                        if (editorGridEnabled && $(el).data('ui-draggable')) $(el).draggable('option', 'grid', [editorGridSize, editorGridSize]);
                        if (_editorLocked[el.id]) editorApplyLock(el, true);
                    }
                }
            }).observe(wpp, { subtree: true, attributes: true, attributeFilter: ['class'] });
            // Zoom observer — watch #wpp style for transform changes
            new MutationObserver(() => editorUpdateZoom())
                .observe(wpp, { attributes: true, attributeFilter: ['style'] });

            // Middle-mouse-button pan.
            // WebPort stores pan as "translate(Xpx, Ypx) scale(s, s)" on #wpp inline style.
            // We parse the translate values and write them back, leaving scale untouched.
            const getWppTranslate = () => {
                const t = wpp.style.transform || '';
                const m = t.match(/translate\(\s*([-\d.]+)px,\s*([-\d.]+)px\s*\)/);
                if (m) return [parseFloat(m[1]), parseFloat(m[2])];
                // Fallback: matrix format
                const mx = t.match(/matrix\(([^)]+)\)/);
                if (mx) { const p = mx[1].split(',').map(Number); return [p[4], p[5]]; }
                return [0, 0];
            };
            const setWppTranslate = (tx, ty) => {
                const t = wpp.style.transform || '';
                if (t.includes('translate(')) {
                    wpp.style.transform = t.replace(/translate\([^)]+\)/, `translate(${tx}px, ${ty}px)`);
                } else if (t.includes('matrix(')) {
                    wpp.style.transform = t.replace(/matrix\(([^)]+)\)/, (_, p) => {
                        const parts = p.split(',').map(s => s.trim());
                        parts[4] = tx; parts[5] = ty;
                        return `matrix(${parts.join(', ')})`;
                    });
                } else {
                    wpp.style.transform = `translate(${tx}px, ${ty}px)`;
                }
            };
            iDoc.addEventListener('mousedown', e => {
                if (e.button !== 1) return;
                e.preventDefault();
                e.stopImmediatePropagation();
                const [origTx, origTy] = getWppTranslate();
                const startX = e.clientX;
                const startY = e.clientY;
                iDoc.body.style.cursor = 'grabbing';
                const onMove = mv => {
                    setWppTranslate(origTx + (mv.clientX - startX), origTy + (mv.clientY - startY));
                };
                const onUp = up => {
                    if (up.button !== 1) return;
                    iDoc.removeEventListener('mousemove', onMove, true);
                    iDoc.removeEventListener('mouseup',   onUp,   true);
                    iDoc.body.style.cursor = '';
                };
                iDoc.addEventListener('mousemove', onMove, true);
                iDoc.addEventListener('mouseup',   onUp,   true);
            }, true);
            // Suppress browser autoscroll overlay on middle click
            iDoc.addEventListener('auxclick', e => { if (e.button === 1) e.preventDefault(); }, true);

            const panel = iDoc.getElementById('wp_editpanel');
            $(panel).on('selectablestop', () => {
                // Strip locked elements from rubber-band selection
                iDoc.querySelectorAll('.wpCompObject.ui-selected').forEach(el => {
                    if (_editorLocked[el.id]) el.classList.remove('ui-selected');
                });
                editorUpdateToolbarContext();
            });
            // selectablestop only fires for rubber-band selections.
            // Single clicks use WebPort's own click handler, so we sync after each click too.
            $(panel).on('click', '.wpCompObject', () => setTimeout(editorUpdateToolbarContext, 30));
            $(panel).on('dragstart', '.wpCompObject', editorPushUndo);
            $(panel).on('drag', '.wpCompObject', function(_e, ui) {
                if (editorGridEnabled) {
                    // Snap to canvas origin, not drag-start position
                    ui.position.left = Math.round(ui.position.left / editorGridSize) * editorGridSize;
                    ui.position.top  = Math.round(ui.position.top  / editorGridSize) * editorGridSize;
                }
                const sel = editorGetSelected();
                if (sel.length === 1) editorUpdatePositionInspector(sel[0]);
            });
            let _dragSaveTimer = null;
            $(panel).on('dragstop', '.wpCompObject', function() {
                editorUpdateToolbarContext();
                // Batch saves: collect all moved elements, save after short delay
                // (multiple elements can fire dragstop nearly simultaneously)
                clearTimeout(_dragSaveTimer);
                _dragSaveTimer = setTimeout(() => {
                    const sel = editorGetSelected();
                    editorSavePositions(sel.length ? sel : [this]);
                }, 300);
            });

            if (!iDoc.body.hasAttribute('tabindex')) iDoc.body.setAttribute('tabindex', '-1');
            editorBindKeyboard(iDoc);
            // editorInitTooltip disabled — replaced by diagram tooltip (initDiagramTooltip)
            // which shows the same label + all related live tag values in one tooltip.
            // Call here because the init polling loop is already cleared by the time
            // the iframe content is ready.
            initDiagramTooltip();
            iframe.contentWindow.focus();

            editorUpdateToolbarContext();
        };

        // Poll until .wpCompObject elements exist, then run setup once.
        // Called on initial load AND after every iframe reload (e.g. WebPort
        // rebuilds the iframe when symbols are added from the portlet panel).
        const waitForSymbols = () => {
            if (iframe.contentDocument?.getElementById('inu-editor-toolbar')) return; // already set up
            if (iframe.contentDocument?.querySelector('.wpCompObject')) { setup(); return; }
            let att = 0;
            const t = setInterval(() => {
                att++;
                if (iframe.contentDocument?.querySelector('.wpCompObject')) { clearInterval(t); setup(); }
                else if (att > 60) clearInterval(t); // 12 s safety timeout
            }, 200);
        };

        // Re-run after every iframe reload so toolbar survives WebPort rebuilds
        iframe.addEventListener('load', waitForSymbols);
        waitForSymbols();
    }

    let devMutBusy = false;
    async function initDevicePage() {
        console.log(CFG.logPrefix, 'v' + CFG.version, 'Activating (IO-Enheter)');
        injectStyles();
        addDeviceHeaders();
        await addDeviceColumns();
        initDeviceContextMenu();
        const tb = document.querySelector('#devicetable tbody');
        if (tb) new MutationObserver(() => {
            if (devMutBusy) return;
            devMutBusy = true;
            addDeviceHeaders();
            addDeviceColumns().finally(() => { devMutBusy = false; });
        }).observe(tb, { childList: true });
    }

    // ============================================================
    // CONTENT PAGE — floating Live Monitor button
    // ============================================================
    function isContentPage() {
        if (location.pathname !== '/page') return false;
        if (!/route=view/.test(location.search)) return false;
        // Wait until iframe content has loaded with symbol elements
        try {
            const doc = document.querySelector('iframe')?.contentDocument;
            return !!(doc && doc.querySelector('.komponent[data-prefix]'));
        } catch(e) { return false; }
    }

    function getContentPrefixes() {
        try {
            const doc = document.querySelector('iframe')?.contentDocument;
            if (!doc) return [];
            const seen = new Set();
            const result = [];
            for (const el of doc.querySelectorAll('.komponent[data-prefix]')) {
                const prefix = el.getAttribute('data-prefix');
                if (!prefix || /\s/.test(prefix) || seen.has(prefix)) continue;
                seen.add(prefix);
                // On view pages WebPort renders live values inside the symbol element, so innerText
                // contains e.g. "GT102 (Reglerar)19.0 °C0.0 °C20.5 °C" — strip all number+unit substrings
                const rawLabel = (el.innerText || '').split(/[\n\r]/)[0] || '';
                const label = rawLabel.replace(/-?\d+(?:[,.]?\d+)?\s*[°%℃℉]\s*/g, ' ').replace(/\s{2,}/g, ' ').trim()
                           || prefix.split('_').pop();
                const poid   = el.id;
                const pageid = poid.split('-2E-')[0];
                result.push({ prefix, label, poid, pageid });
            }
            return result;
        } catch(e) { return []; }
    }

    async function launchContentMonitor() {
        const prefixEntries = getContentPrefixes();
        if (!prefixEntries.length) { toastErr('Inga symboler med tagprefix hittades på sidan'); return; }

        const btn = document.getElementById('inu-content-mon-btn');
        if (btn) { btn.disabled = true; btn.innerHTML = '<i class="fa fa-spinner fa-spin"></i>'; }

        try {
            const r = await fetch('/tag/GetTagList?draw=1&limit=9999&offset=0&sortcol=0&sortdir=asc&search=');
            if (!r.ok) throw new Error('HTTP ' + r.status);
            const json = await r.json();
            const decode = h => { const d = document.createElement('div'); d.innerHTML = h; return d.textContent; };

            // Map prefix → visual label (e.g. "AS01-GT301")
            const labelMap = {};
            prefixEntries.forEach(({ prefix, label }) => { labelMap[prefix] = label; });
            const prefixes = prefixEntries.map(e => e.prefix);

            const faultDescMap = {}; // prefix → _FAULT tag description
            const ctrlTagSet  = {}; // prefix → Set of control suffixes found ('M','OPM','MCMD')
            const tags = [];
            for (const row of (json.data || [])) {
                const name = decode(row['0'] || '');
                if (!name) continue;
                // Sensor tags for the monitor grid
                if (prefixes.some(p => name.startsWith(p + '_') && /_(PV|FAULT|V)$/.test(name))) {
                    tags.push({
                        id: row.DT_RowId || encodeTag(name), name,
                        unit: decode(row['8'] || ''),
                        rawmin: parseFloat(row['4']) || 0, rawmax: parseFloat(row['5']) || 0,
                        engmin: parseFloat(row['6']) || 0, engmax: parseFloat(row['7']) || 0,
                        dtype: row['3'] || '',
                    });
                }
                // Capture _FAULT description
                if (name.endsWith('_FAULT')) {
                    const fp = prefixes.find(p => name === p + '_FAULT');
                    if (fp) { const desc = decode(row['10'] || '').trim(); if (desc) faultDescMap[fp] = desc; }
                }
                // Detect control capabilities: _M, _OPM, _MCMD
                const ctrlSuffix = ['_M','_OPM','_MCMD'].find(s => name.endsWith(s));
                if (ctrlSuffix) {
                    const cp = prefixes.find(p => name === p + ctrlSuffix);
                    if (cp) {
                        if (!ctrlTagSet[cp]) ctrlTagSet[cp] = new Set();
                        ctrlTagSet[cp].add(ctrlSuffix.slice(1)); // 'M','OPM','MCMD'
                    }
                }
            }

            if (!tags.length) { toastErr('Inga sensorer hittades för symbolen på denna sida'); return; }

            const grouped = groupTagsBySensor(tags);
            const validTags = [];
            for (const g of grouped) {
                if (g.tags.some(t => t.suffix === 'PV' || t.suffix === 'V')) g.tags.forEach(t => validTags.push(t));
            }
            if (!validTags.length) { toastErr('Inga sensorer med PV eller V hittades'); return; }

            // Build labelMap keyed by sensor group key (= prefix of the PV tag)
            // groupTagsBySensor uses the tag name minus suffix as the key — map those to visual labels
            const sensorLabelMap = {};
            for (const g of grouped) {
                // g.key is e.g. "HUS03_AS01_GT301"; find which page prefix it belongs to
                const match = prefixes.find(p => g.key === p || g.key.startsWith(p + '_') || p.startsWith(g.key + '_') || p === g.key);
                if (match && labelMap[match]) sensorLabelMap[g.key] = labelMap[match];
            }

            // Build controls list: detect type from available tags (_OPM+_M → valve, _MCMD+_M → pump/fan)
            const controls = [];
            for (const e of prefixEntries) {
                const ct = ctrlTagSet[e.prefix];
                if (!ct) continue;
                if (ct.has('OPM') && ct.has('M')) {
                    controls.push({ type: 'valve', selectKey: '003select', poid: e.poid, pageid: e.pageid, prefix: e.prefix, label: e.label, faultDesc: faultDescMap[e.prefix] || '' });
                } else if (ct.has('MCMD') && ct.has('M')) {
                    controls.push({ type: 'pump',  selectKey: '010select', poid: e.poid, pageid: e.pageid, prefix: e.prefix, label: e.label, faultDesc: faultDescMap[e.prefix] || '' });
                }
            }

            const pageLabel = document.title.replace(/^Web Port\s*[-–]\s*/i, '').trim() || prefixes[0];
            new LiveMonitor(validTags, pageLabel, sensorLabelMap, controls, faultDescMap).open();
        } catch(e) {
            toastErr('Kunde inte hämta taggar: ' + e.message);
        } finally {
            if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fa fa-television"></i>'; }
        }
    }

    function initContentPage() {
        const btn = document.createElement('button');
        btn.id = 'inu-content-mon-btn';
        btn.title = 'Live Monitor för denna sida (WP+)';
        btn.innerHTML = '<i class="fa fa-television"></i>';
        btn.style.cssText = [
            'position:fixed', 'bottom:20px', 'right:20px', 'z-index:99999',
            'width:46px', 'height:46px', 'border-radius:50%',
            'background:#1e3a5f', 'color:#fff', 'border:none',
            'font-size:18px', 'cursor:pointer',
            'box-shadow:0 2px 10px rgba(0,0,0,.4)',
            'display:flex', 'align-items:center', 'justify-content:center',
            'transition:background .15s',
        ].join(';');
        btn.addEventListener('mouseenter', () => { btn.style.background = '#2d5a9e'; });
        btn.addEventListener('mouseleave', () => { btn.style.background = '#1e3a5f'; });
        btn.addEventListener('click', launchContentMonitor);
        document.body.appendChild(btn);
    }

    // ============================================================
    // DIAGRAM TOOLTIP — hover a symbol to see all related tag values
    // ============================================================
    let _diagramTooltipDoc = null; // tracks which iframe document we initialized for
    let _diagramTooltipEnabled = true;

    function initDiagramTooltip() {
        try {
        const iframe = document.querySelector('iframe');
        if (!iframe) return;
        const iDoc = iframe.contentDocument || iframe.contentWindow?.document;
        if (!iDoc) return;
        const wpp = iDoc.getElementById('wpp');
        if (!wpp) return;
        const pageId = new URLSearchParams(location.search).get('pageid');
        if (!pageId) return;

        // If the iframe document is the same one we already initialized, skip.
        // If it changed (SPA navigation), re-initialize for the new page.
        if (_diagramTooltipDoc === iDoc) return;
        _diagramTooltipDoc = iDoc;

        // Restore preference (default OFF)
        try { _diagramTooltipEnabled = GM_getValue('inu_diagram_tooltip', false); } catch (e) {}

        // Toggle button next to brand pill
        const pill = document.getElementById('inu-wp-pill');
        if (pill && !document.getElementById('inu-dt-toggle')) {
            const btn = document.createElement('span');
            btn.id = 'inu-dt-toggle';
            btn.style.cssText = 'padding:3px 8px;border-radius:3px;font-size:10px;font-weight:600;cursor:pointer;align-self:center;display:inline-flex;align-items:center;gap:4px;margin-left:4px;user-select:none;';
            function updateBtn() {
                btn.style.background = _diagramTooltipEnabled ? '#1565c0' : '#555';
                btn.style.color = '#fff';
                btn.textContent = _diagramTooltipEnabled ? '🏷 Tooltips PÅ' : '🏷 Tooltips AV';
                btn.title = _diagramTooltipEnabled ? 'Klicka för att dölja diagram-tooltips' : 'Klicka för att visa diagram-tooltips';
            }
            btn.addEventListener('click', () => {
                _diagramTooltipEnabled = !_diagramTooltipEnabled;
                try { GM_setValue('inu_diagram_tooltip', _diagramTooltipEnabled); } catch (e) {}
                updateBtn();
                if (!_diagramTooltipEnabled) hideTooltip();
            });
            updateBtn();
            pill.parentElement.insertBefore(btn, pill.nextSibling);
        }

        // Cache for refreshvalues response
        let cachedData = null;

        function fetchValues() {
            const xhr = new XMLHttpRequest();
            xhr.open('GET', '/page/refreshvalues?pageid=' + encodeURIComponent(pageId) + '&_=' + Date.now(), true);
            xhr.onload = function () {
                if (xhr.status === 200) {
                    try { cachedData = JSON.parse(xhr.responseText); } catch (e) { /* ignore */ }
                }
            };
            xhr.send();
        }

        fetchValues();
        setInterval(fetchValues, 2000);

        function extractPoid(el) {
            const div = el.closest('.wpCompObject');
            if (!div || !div.id) return null;
            const parts = div.id.split('-2E-');
            return parts.length > 1 ? parts[parts.length - 1] : null;
        }

        function getFunctionsForPoid(poid) {
            if (!cachedData || !cachedData.functions) return [];
            const prefix = poid + '-5F-';
            return cachedData.functions.filter(f => f.id && f.id.startsWith(prefix)).map(f => {
                const suffix = f.id.substring(prefix.length).replace(/-5F-/g, '_');
                const fullTag = f.id.replace(/-5F-/g, '_');
                return { suffix: '_' + suffix, fullTag: fullTag, value: f.value || '' };
            });
        }

        function getObjectForPoid(poid) {
            if (!cachedData || !cachedData.objects) return null;
            return cachedData.objects.find(o => o.poid === poid);
        }

        let tooltip = null;
        let currentPoid = null;

        // Prefix cache: poid → real tag prefix (e.g. "03_AS01_KVP001_DHW_LOWER").
        // Fetched once per component via /Page/PageObjectProperties.
        // When the fetch completes and a tooltip is still showing for the same
        // poid, the tooltip content is re-rendered in place.
        const prefixCache = {};
        function fetchPrefix(poid) {
            if (poid in prefixCache) return;
            prefixCache[poid] = null; // mark pending
            const obj = getObjectForPoid(poid);
            if (!obj) return;
            const xhr = new XMLHttpRequest();
            xhr.open('GET', '/Page/PageObjectProperties?pageid=' + encodeURIComponent(pageId) + '&id=' + encodeURIComponent(obj.fullId), true);
            xhr.onload = function () {
                if (xhr.status === 200) {
                    const doc = new DOMParser().parseFromString(xhr.responseText, 'text/html');
                    const pfx = doc.querySelector('textarea[name="prefix"], input[name="prefix"]');
                    if (pfx && pfx.value) prefixCache[poid] = pfx.value;
                    else prefixCache[poid] = '';
                }
                // Re-render tooltip if still showing for this poid
                if (tooltip && currentPoid === poid) {
                    tooltip.innerHTML = renderContentForPoid(poid);
                }
            };
            xhr.send();
        }

        // Render tooltip HTML for a given poid — called from showTooltip
        // and again from the fetchPrefix callback when the real name arrives.
        function renderContentForPoid(poid) {
            const obj = getObjectForPoid(poid);
            const funcs = getFunctionsForPoid(poid);
            const label = obj ? (obj.text || poid.replace(/-5F-/g, '_')) : poid.replace(/-5F-/g, '_');
            const realPrefix = prefixCache[poid];

            let html = '<div class="inu-dt-header">' + _tplEsc(label) + '</div>';
            if (realPrefix) {
                html += '<div class="inu-dt-stem">' + _tplEsc(realPrefix) + '</div>';
            } else if (realPrefix === null) {
                html += '<div class="inu-dt-stem inu-dt-loading"></div>';
            }
            if (funcs.length === 0) {
                html += '<div class="inu-dt-empty">Inga taggvärden</div>';
            } else {
                html += '<table class="inu-dt-table"><tbody>';
                for (const f of funcs) {
                    const isAlarm = /_(AL\d*|FAULT|HAL|LAL)$/i.test(f.suffix);
                    const isActive = isAlarm && f.value !== '0' && f.value !== '';
                    const rowCls = isActive ? ' class="inu-dt-alarm"' : '';
                    html += '<tr' + rowCls + '><td class="inu-dt-suffix">' + _tplEsc(f.suffix) + '</td><td class="inu-dt-value">' + _tplEsc(f.value) + '</td></tr>';
                }
                html += '</tbody></table>';
            }
            return html;
        }

        function showTooltip(compEl, poid) {
            hideTooltip();
            currentPoid = poid;

            // Kick off prefix fetch — re-renders tooltip in place when it completes
            fetchPrefix(poid);

            tooltip = iDoc.createElement('div');
            tooltip.className = 'inu-diagram-tooltip';
            tooltip.innerHTML = renderContentForPoid(poid);
            iDoc.body.appendChild(tooltip);

            const rect = compEl.getBoundingClientRect();
            let left = rect.right + 8;
            let top = rect.top;
            const tw = tooltip.offsetWidth;
            const th = tooltip.offsetHeight;
            const vw = iDoc.defaultView.innerWidth;
            const vh = iDoc.defaultView.innerHeight;
            if (left + tw > vw - 8) left = rect.left - tw - 8;
            if (top + th > vh - 8) top = Math.max(8, vh - th - 8);
            tooltip.style.left = left + 'px';
            tooltip.style.top = top + 'px';
        }

        function hideTooltip() {
            if (tooltip) { tooltip.remove(); tooltip = null; }
            currentPoid = null;
        }

        // Inject tooltip CSS into iframe
        const style = iDoc.createElement('style');
        style.textContent = `
.inu-diagram-tooltip { position:fixed; z-index:99999; background:#1e293b; color:#e2e8f0; border-radius:6px; box-shadow:0 6px 24px rgba(0,0,0,.4); padding:0; min-width:200px; max-width:420px; font-family:system-ui,-apple-system,sans-serif; font-size:12px; pointer-events:none; }
.inu-dt-header { padding:8px 12px 2px; font-weight:700; font-size:13px; color:#fff; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.inu-dt-stem { padding:0 12px 6px; font-family:monospace; font-size:11px; color:#64748b; border-bottom:1px solid #334155; }
.inu-dt-loading::after { content:'.'; animation:inu-dt-dots 1.2s steps(3,end) infinite; }
@keyframes inu-dt-dots { 33% { content:'..'; } 66% { content:'...'; } 100% { content:'.'; } }
.inu-dt-empty { padding:8px 12px; font-size:11px; color:#94a3b8; font-style:italic; }
.inu-dt-table { width:100%; border-collapse:collapse; }
.inu-dt-table td { padding:3px 10px; border-bottom:1px solid #262f3d; }
.inu-dt-table tr:last-child td { border-bottom:none; }
.inu-dt-suffix { font-family:monospace; font-weight:600; color:#94a3b8; font-size:11px; white-space:nowrap; }
.inu-dt-addr { font-family:monospace; font-size:10px; color:#475569; white-space:nowrap; }
.inu-dt-value { text-align:right; font-family:monospace; font-weight:600; color:#e2e8f0; font-size:12px; white-space:nowrap; }
.inu-dt-alarm .inu-dt-value { color:#f87171; }
.inu-dt-alarm .inu-dt-suffix { color:#f87171; }
`;
        iDoc.head.appendChild(style);

        // Suppress the old editor tooltip if it exists, and strip native
        // title attributes that cause a duplicate browser tooltip.
        const oldTip = iDoc.getElementById('inu-sym-tooltip');
        if (oldTip) oldTip.remove();
        wpp.querySelectorAll('.wpCompObject[title]').forEach(el => el.removeAttribute('title'));

        // Event delegation: .wpCompObject has pointer-events:none so direct
        // mouseenter doesn't fire. Instead listen on iDoc.body for mouseover
        // events that bubble up from child elements inside the component divs.
        iDoc.body.addEventListener('mouseover', function (e) {
            if (!_diagramTooltipEnabled) return;
            const comp = e.target.closest('.wpCompObject');
            if (!comp) { if (currentPoid) hideTooltip(); return; }
            const poid = extractPoid(comp);
            if (!poid || poid === currentPoid) return;
            showTooltip(comp, poid);
        }, true);

        iDoc.body.addEventListener('mouseout', function (e) {
            const comp = e.target.closest('.wpCompObject');
            const related = e.relatedTarget?.closest?.('.wpCompObject');
            if (comp && (!related || related !== comp)) {
                hideTooltip();
            }
        }, true);

        // Fix missing rotations — WebPort sometimes fails to apply them during
        // initial rendering (race condition with our script's DOM modifications).
        // Parse r0/r90/r180/r270 from the SVG class and apply the transform.
        wpp.querySelectorAll('.wpCompObject').forEach(comp => {
            const img = comp.querySelector('.wpCompImage');
            const svg = img?.querySelector('svg');
            if (!img || !svg) return;
            const cls = svg.getAttribute('class') || '';
            const rotMatch = cls.match(/\br(\d+)\b/);
            if (!rotMatch) return;
            const deg = parseInt(rotMatch[1]);
            if (deg === 0) return;
            if (img.style.transform && img.style.transform.includes('rotate')) return;
            const existing = img.style.transform || '';
            img.style.transform = existing + (existing ? ' ' : '') + 'rotate(' + deg + 'deg)';
        });

        console.log(CFG.logPrefix, 'Diagram tooltip active for', pageId, '(' + wpp.querySelectorAll('.wpCompObject').length + ' components)');

        // Watch for SPA navigation: if the iframe document changes, reset and
        // re-initialize. Check every 2 seconds alongside the value refresh.
        setInterval(() => {
            const curDoc = iframe.contentDocument || iframe.contentWindow?.document;
            if (curDoc && curDoc !== _diagramTooltipDoc) {
                _diagramTooltipDoc = null; // reset so next call re-initializes
                initDiagramTooltip();
            }
        }, 2000);

        } catch (e) { console.error(CFG.logPrefix, 'Diagram tooltip init failed:', e); }
    }

    // ============================================================
    // INIT
    // ============================================================
    function init() {
        let att=0;
        const wait=setInterval(()=>{
            att++;
            // Bail if not a WebPort page
            if (att > 5 && !isWebPort()) { clearInterval(wait); return; }
            // Brand pill, source check, and log panel on any WebPort page (once)
            if (isWebPort() && document.getElementById('top-menu') && !document.getElementById('inu-wp-pill')) { injectBrandPill(); checkSources(); hookToastr(); initLogPanel(); }
            // Diagram tooltip: retry each tick until the iframe's wpp container exists.
            // Safe to call repeatedly — returns immediately if already active.
            if (isWebPort()) initDiagramTooltip();
            if(isInuTagPage()){
                clearInterval(wait);
                console.log(CFG.logPrefix, 'v' + CFG.version, 'Activating');
                injectStyles();
                addColumns();
                patchClickIndices();
                patchValueUpdate();
                initContextMenu();
                createToolbar();
                updSummary();
                hijackSave();
                hookSearch();
                initDragSelect();
                addColumnFilters();
                const tb=document.querySelector('#tagtable tbody');
                if(tb) new MutationObserver(()=>setTimeout(()=>{addColumns();updSummary();applyFilter();syncSelCheckboxes();addColumnFilters();reapplyPendingDeletes();},50)).observe(tb,{childList:true});
                // Keep save button visible while pending deletes exist
                const liSave = document.getElementById('li_wp_mnu_wp_tb_save');
                if (liSave) new MutationObserver(() => {
                    if (pendingDeletes.size > 0 && liSave.style.display === 'none') liSave.removeAttribute('style');
                }).observe(liSave, {attributes: true, attributeFilter: ['style']});
            } else if(isDevicePage()){
                clearInterval(wait);
                initDevicePage();
            } else if(isContentPage()){
                clearInterval(wait);
                initContentPage();
            } else if(isPageEditorPage()){
                clearInterval(wait);
                initPageEditor();
            } else if(att>=100) {
                clearInterval(wait);
            }
        }, 300);
    }

    init();

    function cleanupPageEditor() {
        const iDoc = _editorIframe?.contentDocument;
        if (iDoc) {
            // Remove injected toolbar and styles
            iDoc.getElementById('inu-editor-toolbar')?.remove();
            iDoc.getElementById('inu-editor-styles')?.remove();
            // Remove lock badges and classes from all symbols
            iDoc.querySelectorAll('.inu-lock-badge').forEach(el => el.remove());
            iDoc.querySelectorAll('.inu-locked').forEach(el => el.classList.remove('inu-locked'));
            iDoc.getElementById('inu-shortcut-overlay')?.remove();
            iDoc.getElementById('inu-batch-overlay')?.remove();
            iDoc.getElementById('inu-sym-picker')?.remove();
            iDoc.getElementById('inu-sym-tooltip')?.remove();
        }
        // Also remove toolbar from outer page (fallback injection path)
        document.getElementById('inu-editor-toolbar')?.remove();
        _editorIframe = null;
    }

    // ============================================================
    // PID ADVISOR
    // ============================================================
    let _pidStylesInjected = false;

    const PID_CIRCUIT_TYPES = {
        supply_temp:   { label: 'Tillufttemperatur',           tiRange: [100, 400],  kpRange: [0.5, 3],   responseTime: [30,  120], defaultWindow: 3600000  },
        room_temp:     { label: 'Rumstemperatur',              tiRange: [300, 900],  kpRange: [0.2, 1.5], responseTime: [120, 600], defaultWindow: 21600000 },
        fan_pressure:  { label: 'Fläkttryck / Kanalstatik',   tiRange: [20,  120],  kpRange: [0.5, 5],   responseTime: [5,   30],  defaultWindow: 600000   },
        pump_pressure: { label: 'Differenstryck (pump)',       tiRange: [20,  100],  kpRange: [0.5, 5],   responseTime: [5,   20],  defaultWindow: 600000   },
        district_heat: { label: 'Fjärrvärme (primär)',         tiRange: [200, 600],  kpRange: [0.3, 2],   responseTime: [60,  300], defaultWindow: 3600000  },
        district_cool: { label: 'Fjärrkyla (primär)',          tiRange: [200, 600],  kpRange: [0.3, 2],   responseTime: [60,  300], defaultWindow: 3600000  },
        humidity:      { label: 'Luftfuktighet',               tiRange: [400, 1200], kpRange: [0.3, 1.5], responseTime: [120, 600], defaultWindow: 21600000 },
        split_range:   { label: 'Splitreglering (värme/kyla)', tiRange: [100, 400],  kpRange: [0.5, 3],   responseTime: [30,  120], defaultWindow: 3600000  },
    };

    // Infer circuit type from tag prefix — uses last sensor segment before suffix
    // GT* = temperature, GP* = pressure, FJV* prefix = district heat, FJK* = district cool
    function detectCircuitType(sensorPrefix) {
        if (!sensorPrefix) return null;
        const parts = sensorPrefix.split('_');
        const sensor = parts[parts.length - 1].toUpperCase();
        const prefixUpper = sensorPrefix.toUpperCase();
        // District energy: FJK = cooling, FJV = heating — check anywhere in prefix
        if (/^FJK/.test(prefixUpper) || parts.some(p => /^FJK/i.test(p))) return 'district_cool';
        if (/^FJV/.test(prefixUpper) || parts.some(p => /^FJV/i.test(p))) return 'district_heat';
        // Pressure
        if (/^GP\d/.test(sensor)) return 'fan_pressure';
        // Humidity — GM = givare, moisture
        if (/^GM\d/.test(sensor)) return 'humidity';
        // Temperature
        if (/^GT\d/.test(sensor)) {
            if (parts.some(p => /^RU\d*$/i.test(p) || /^RUM$/i.test(p))) return 'room_temp';
            return 'supply_temp';
        }
        return null;
    }

    function injectPidStyles() {
        if (_pidStylesInjected) return; _pidStylesInjected = true;
        const s = document.createElement('style');
        s.textContent = `
.inu-pid-modal { min-width:560px; max-width:800px; width:90vw; }
.inu-pid-modal h3 { margin:0 0 12px; font-size:14px; display:flex; align-items:center; gap:8px; }
.inu-pid-section { margin-bottom:12px; }
.inu-pid-row { display:flex; gap:10px; align-items:center; margin-bottom:8px; flex-wrap:wrap; }
.inu-pid-row label { font-size:11px; color:#888; min-width:60px; }
.inu-pid-row input[type=text] { flex:1; min-width:120px; }
.inu-pid-row select { flex:1; }
.inu-pid-tags { background:#f5f5f5; border:1px solid #ddd; border-radius:4px; padding:8px 10px; font-size:11px; line-height:1.8; }
body.dark .inu-pid-tags { background:#2a2a2a; border-color:#444; }
.inu-pid-tag-ok   { color:#388e3c; }
.inu-pid-tag-warn { color:#f57c00; }
.inu-pid-tag-miss { color:#999; }
.inu-pid-chart-wrap { position:relative; width:100%; margin-bottom:12px; }
.inu-pid-chart-wrap svg { width:100%; display:block; }
.inu-pid-tooltip { position:absolute; pointer-events:none; background:rgba(30,30,30,.88); color:#fff;
    font-size:10px; padding:4px 7px; border-radius:3px; white-space:nowrap; display:none; z-index:10; }
.inu-pid-results { display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-bottom:12px; }
.inu-pid-results-box { background:#f5f5f5; border:1px solid #ddd; border-radius:4px; padding:8px 10px; font-size:11px; }
body.dark .inu-pid-results-box { background:#2a2a2a; border-color:#444; }
.inu-pid-results-box h5 { margin:0 0 6px; font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:#888; }
.inu-pid-metric-row { display:flex; justify-content:space-between; padding:2px 0; border-bottom:1px solid #eee; }
body.dark .inu-pid-metric-row { border-color:#333; }
.inu-pid-metric-row:last-child { border:none; }
.inu-pid-metric-val { font-weight:600; }
.inu-pid-advice { margin-bottom:12px; }
.inu-pid-advice h5 { margin:0 0 6px; font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:#888; }
.inu-pid-adv-item { display:flex; gap:8px; align-items:flex-start; padding:5px 8px; border-radius:3px; margin-bottom:4px; font-size:11px; }
.inu-pid-adv-item.ok   { background:#e8f5e9; color:#1b5e20; }
.inu-pid-adv-item.warn { background:#fff3e0; color:#e65100; }
.inu-pid-adv-item.info { background:#e3f2fd; color:#0d47a1; }
body.dark .inu-pid-adv-item.ok   { background:#1b3a1f; color:#a5d6a7; }
body.dark .inu-pid-adv-item.warn { background:#3e2000; color:#ffcc80; }
body.dark .inu-pid-adv-item.info { background:#0d2137; color:#90caf9; }
.inu-pid-live-bar { background:#e3f2fd; border:1px solid #90caf9; border-radius:4px; padding:10px 12px; font-size:11px; margin-bottom:12px; }
body.dark .inu-pid-live-bar { background:#0d2137; border-color:#1565c0; }
.inu-pid-progress { height:6px; background:#ddd; border-radius:3px; margin-top:6px; overflow:hidden; }
body.dark .inu-pid-progress { background:#333; }
.inu-pid-progress-fill { height:100%; background:#2d5a9e; border-radius:3px; transition:width .3s; }
.inu-pid-btn-row { display:flex; gap:8px; justify-content:flex-end; margin-top:4px; }
.inu-pid-status { font-size:10px; color:#888; margin-top:4px; }
.inu-pid-tag-grid { display:table; width:100%; border-spacing:0 3px; }
.inu-pid-tag-row  { display:table-row; }
.inu-pid-tag-lbl  { display:table-cell; white-space:nowrap; padding:0 8px 0 0; font-size:10px; font-weight:700; color:#666; vertical-align:middle; width:52px; }
.inu-pid-tag-inp  { display:table-cell; width:100%; }
.inu-pid-tag-inp input { width:100%; box-sizing:border-box; font-size:11px; padding:3px 6px; border:1px solid #ccc; border-radius:3px; background:#fff; }
.inu-pid-tag-inp input.has-val { border-color:#388e3c; background:#f4faf4; }
.inu-pid-tag-inp input.is-req  { }
body.dark .inu-pid-tag-inp input { background:#1e1e1e; border-color:#555; color:#ddd; }
body.dark .inu-pid-tag-inp input.has-val { background:#1b3a1f; border-color:#388e3c; }
.inu-pid-tag-stat { display:table-cell; padding:0 0 0 5px; font-size:12px; vertical-align:middle; white-space:nowrap; }
.inu-pid-tune-table { width:100%; border-collapse:collapse; font-size:11px; }
.inu-pid-tune-table th { text-align:left; color:#888; font-weight:600; padding:2px 6px 4px 0; border-bottom:1px solid #ddd; }
.inu-pid-tune-table td { padding:3px 6px 3px 0; border-bottom:1px solid #eee; }
.inu-pid-tune-table tr:last-child td { border:none; }
body.dark .inu-pid-tune-table th { border-color:#444; }
body.dark .inu-pid-tune-table td { border-color:#333; }
.inu-pid-group-hdr { font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:.05em; color:#aaa; margin:8px 0 4px; padding-bottom:2px; border-bottom:1px solid #e0e0e0; }
body.dark .inu-pid-group-hdr { border-color:#333; color:#666; }
.inu-pid-live-strip { display:flex; gap:16px; align-items:center; padding:5px 8px; background:#f0f4ff; border:1px solid #c5d5f0; border-radius:3px; font-size:11px; margin-top:6px; min-height:24px; }
body.dark .inu-pid-live-strip { background:#0d1a2e; border-color:#1a3a6e; }
.inu-pid-live-strip span { color:#555; }
body.dark .inu-pid-live-strip span { color:#aaa; }
.inu-pid-live-strip b { color:#2d5a9e; margin-right:3px; }
body.dark .inu-pid-live-strip b { color:#90caf9; }
.inu-pid-hint { font-size:11px; color:#555; margin-top:8px; border:1px solid #e0e0e0; border-radius:3px; padding:5px 8px; background:#fafafa; }
body.dark .inu-pid-hint { background:#1e1e1e; border-color:#333; color:#aaa; }
.inu-pid-hint summary { cursor:pointer; color:#2d5a9e; font-weight:600; list-style:none; outline:none; }
body.dark .inu-pid-hint summary { color:#90caf9; }
.inu-pid-hint ol { margin:6px 0 4px 16px; padding:0; line-height:1.8; }
.inu-pid-hint p  { margin:4px 0 0; color:#888; font-style:italic; }
.inu-pid-params-details { font-size:11px; color:#555; margin-top:6px; border:1px solid #e0e0e0; border-radius:3px; padding:4px 8px; background:#fafafa; }
body.dark .inu-pid-params-details { background:#1e1e1e; border-color:#333; color:#aaa; }
.inu-pid-params-details summary { cursor:pointer; color:#2d5a9e; font-weight:600; list-style:none; outline:none; user-select:none; }
body.dark .inu-pid-params-details summary { color:#90caf9; }
.inu-pid-live-body { display:flex; gap:10px; align-items:flex-start; }
.inu-pid-live-body .inu-pid-chart-wrap { flex:1; min-width:0; }
.inu-pid-live-stats { width:110px; flex-shrink:0; font-size:11px; display:flex; flex-direction:column; gap:8px; }
.inu-pid-stat-group { background:#f5f5f5; border:1px solid #e0e0e0; border-radius:4px; padding:4px 6px; }
body.dark .inu-pid-stat-group { background:#222; border-color:#333; }
.inu-pid-stat-hdr { font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:.05em; color:#888; margin-bottom:3px; border-bottom:1px solid #e0e0e0; padding-bottom:2px; }
body.dark .inu-pid-stat-hdr { border-color:#333; }
.inu-pid-stat-row { display:flex; justify-content:space-between; gap:4px; line-height:1.7; color:#555; }
body.dark .inu-pid-stat-row { color:#bbb; }
.inu-pid-stat-row span:first-child { color:#888; font-size:10px; }
.inu-pid-stat-row span:last-child { font-weight:600; text-align:right; }
.inu-pid-stat-hi span:first-child { color:#555; font-weight:600; }
body.dark .inu-pid-stat-hi span:first-child { color:#ddd; }
[data-tip] { position:relative; cursor:help; }
[data-tip]::after {
    content: attr(data-tip);
    position: absolute;
    left: 50%; bottom: calc(100% + 6px);
    transform: translateX(-50%);
    background: #333; color: #fff;
    font-size: 10px; font-weight: normal;
    padding: 5px 8px; border-radius: 4px;
    white-space: normal; width: 220px;
    pointer-events: none; opacity: 0;
    transition: opacity 0.15s;
    z-index: 9999;
    line-height: 1.4;
    text-align: left;
}
[data-tip]:hover::after { opacity: 1; }
`;
        document.head.appendChild(s);
    }

    async function fetchTrend(tagName, fromMs, toMs) {
        const enc = t => encodeURIComponent(t);
        // WebPort confirmed endpoint: /trend/gettrenddata
        // Date format: "YYYY-MM-DD HH:MM" (local time, no seconds)
        // Response: { trend: [{ data: [[ts_ms, val], ...] }] }
        const fmtDate = ms => {
            const d = new Date(ms);
            const pad = n => String(n).padStart(2, '0');
            return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
        };
        try {
            const url = `/trend/gettrenddata?from=${enc(fmtDate(fromMs))}&to=${enc(fmtDate(toMs))}&view=0&cbxselectall=&tags=${enc(tagName)}`;
            const r = await fetch(url, { credentials: 'include' });
            if (r.ok) {
                const d = await r.json();
                // Response shape: { trend: [{ data: [[ts_ms, val], ...] }] }
                if (d && d.trend && Array.isArray(d.trend) && d.trend[0]?.data?.length >= 2) {
                    return d.trend[0].data.map(p => ({ ts: p[0], val: parseFloat(p[1]) })).filter(p => !isNaN(p.val));
                }
            }
        } catch { /* fall through to probe */ }

        return null;
    }

    async function discoverPidTags(sensorPrefix, devicePrefix) {
        // Uses fnGetNodes() via getAllTagNames() to scan ALL tags regardless of pagination.
        // sensorPrefix → PV, SP/CSP, P, I, D
        // devicePrefix → OP, OPM  (falls back to sensorPrefix if not provided)
        const names = await getAllTagNames();
        const buckets = { PV: [], SP: [], OP: [], P: [], I: [], D: [], DB: [] };

        function scanPrefix(prefix, roles) {
            if (!prefix) return;
            const pfx = prefix.endsWith('_') ? prefix : prefix + '_';
            names.forEach(name => {
                if (!name.startsWith(pfx) && name !== prefix) return;
                const lastSeg = name.split('_').pop().toUpperCase();
                if (roles.includes('PV')  && lastSeg === 'PV')                         buckets.PV.push(name);
                if (roles.includes('SP')  && (lastSeg === 'CSP' || lastSeg === 'SP'))  buckets.SP.push(name);
                if (roles.includes('OP')  && (lastSeg === 'OP'  || lastSeg === 'OPM')) buckets.OP.push(name);
                if (roles.includes('P')   && lastSeg === 'P')                          buckets.P.push(name);
                if (roles.includes('I')   && lastSeg === 'I')                          buckets.I.push(name);
                if (roles.includes('D')   && lastSeg === 'D')                          buckets.D.push(name);
                if (roles.includes('DB')  && (lastSeg === 'DB' || lastSeg === 'DZ'))   buckets.DB.push(name);
            });
        }

        scanPrefix(sensorPrefix, ['PV', 'SP', 'P', 'I', 'D', 'DB']);
        scanPrefix(devicePrefix || sensorPrefix, ['OP']);

        return {
            pv:  buckets.PV[0] || null,
            sp:  buckets.SP[0] || null,
            op:  buckets.OP[0] || null,
            p:   buckets.P[0]  || null,
            i:   buckets.I[0]  || null,
            d:   buckets.D[0]  || null,
            db:  buckets.DB[0] || null,
            buckets,
        };
    }

    // LTTB downsampling: reduces point array to ~targetN points while preserving visual shape.
    // Each point must have { ts, val }. Returns a new array (or original if already small enough).
    function lttbDownsample(pts, targetN) {
        if (pts.length <= targetN) return pts;
        const out = [pts[0]]; // always keep first
        const bucketSize = (pts.length - 2) / (targetN - 2);
        let prevIdx = 0;
        for (let i = 1; i < targetN - 1; i++) {
            const bStart = Math.floor((i - 1) * bucketSize) + 1;
            const bEnd   = Math.min(Math.floor(i * bucketSize) + 1, pts.length - 1);
            // Average of next bucket (for triangle area calculation)
            const nStart = Math.floor(i * bucketSize) + 1;
            const nEnd   = Math.min(Math.floor((i + 1) * bucketSize) + 1, pts.length - 1);
            let avgTs = 0, avgVal = 0, nCount = 0;
            for (let j = nStart; j < nEnd; j++) { avgTs += pts[j].ts; avgVal += pts[j].val; nCount++; }
            if (nCount === 0) { avgTs = pts[pts.length - 1].ts; avgVal = pts[pts.length - 1].val; }
            else { avgTs /= nCount; avgVal /= nCount; }
            // Pick point in current bucket with largest triangle area
            let maxArea = -1, bestIdx = bStart;
            const pTs = pts[prevIdx].ts, pVal = pts[prevIdx].val;
            for (let j = bStart; j < bEnd; j++) {
                const area = Math.abs((pTs - avgTs) * (pts[j].val - pVal) - (pTs - pts[j].ts) * (avgVal - pVal));
                if (area > maxArea) { maxArea = area; bestIdx = j; }
            }
            out.push(pts[bestIdx]);
            prevIdx = bestIdx;
        }
        out.push(pts[pts.length - 1]); // always keep last
        return out;
    }

    function renderPidChart(container, pvPts, spPts, opPts) {
        const W = 500, H = 200, PAD = { t: 12, r: 48, b: 28, l: 48 };
        const inner = { w: W - PAD.l - PAD.r, h: H - PAD.t - PAD.b };

        // Downsample for SVG rendering performance (~800 points per series)
        const CHART_MAX_PTS = 800;
        const pvPlot = lttbDownsample(pvPts, CHART_MAX_PTS);
        const spPlot = lttbDownsample(spPts, CHART_MAX_PTS);
        const opPlot = lttbDownsample(opPts, CHART_MAX_PTS);

        // Time range — use reduce to avoid stack overflow on large arrays (>100k pts)
        const allTs = [...pvPts, ...spPts, ...opPts].map(p => p.ts);
        const tMin = allTs.reduce((a, b) => a < b ? a : b, Infinity);
        const tMax = allTs.reduce((a, b) => a > b ? a : b, -Infinity);
        const tRange = tMax - tMin || 1;

        // Left Y range: PV + SP
        const pvspVals = [...pvPts, ...spPts].map(p => p.val);
        let yMin = pvspVals.reduce((a, b) => a < b ? a : b, Infinity);
        let yMax = pvspVals.reduce((a, b) => a > b ? a : b, -Infinity);
        const pad = (yMax - yMin) * 0.1 || 1;
        yMin -= pad; yMax += pad;
        const yRange = yMax - yMin;

        const tx = ts => PAD.l + (ts - tMin) / tRange * inner.w;
        const tyL = v  => PAD.t + (1 - (v - yMin) / yRange) * inner.h;
        const tyR = v  => PAD.t + (1 - v / 100) * inner.h;

        function polyline(pts, yFn, color, dash) {
            if (!pts.length) return '';
            const d = pts.map((p, i) => `${i ? 'L' : 'M'}${tx(p.ts).toFixed(1)},${yFn(p.val).toFixed(1)}`).join('');
            return `<path d="${d}" fill="none" stroke="${color}" stroke-width="1.5"${dash ? ` stroke-dasharray="${dash}"` : ''}/>`;
        }

        // X axis ticks
        const durMs = tRange;
        const tickCount = 5;
        const tickStep = durMs / tickCount;
        let xTicks = '';
        for (let i = 0; i <= tickCount; i++) {
            const ts = tMin + i * tickStep;
            const x = tx(ts).toFixed(1);
            const elapsed = (ts - tMin) / 1000;
            const lbl = elapsed < 60 ? `${elapsed.toFixed(0)}s`
                      : elapsed < 3600 ? `${(elapsed/60).toFixed(0)}m`
                      : `${(elapsed/3600).toFixed(1)}h`;
            xTicks += `<line x1="${x}" y1="${(PAD.t+inner.h).toFixed(1)}" x2="${x}" y2="${(PAD.t+inner.h+4).toFixed(1)}" stroke="#999" stroke-width="1"/>`;
            xTicks += `<text x="${x}" y="${(PAD.t+inner.h+13).toFixed(1)}" text-anchor="middle" font-size="8" fill="#999">${lbl}</text>`;
        }

        // Left Y ticks
        let yTicksL = '';
        const yTickCount = 4;
        for (let i = 0; i <= yTickCount; i++) {
            const v = yMin + (yRange * i / yTickCount);
            const y = tyL(v).toFixed(1);
            const lbl = v.toFixed(1);
            yTicksL += `<line x1="${(PAD.l-4).toFixed(1)}" y1="${y}" x2="${PAD.l.toFixed(1)}" y2="${y}" stroke="#999" stroke-width="1"/>`;
            yTicksL += `<text x="${(PAD.l-6).toFixed(1)}" y="${y}" text-anchor="end" dominant-baseline="middle" font-size="8" fill="#999">${lbl}</text>`;
        }

        // Right Y ticks (OP %)
        let yTicksR = '';
        if (opPts.length) {
            for (let i = 0; i <= 4; i++) {
                const v = i * 25;
                const y = tyR(v).toFixed(1);
                const rx = (PAD.l + inner.w + 4).toFixed(1);
                yTicksR += `<line x1="${(PAD.l+inner.w).toFixed(1)}" y1="${y}" x2="${rx}" y2="${y}" stroke="#999" stroke-width="1"/>`;
                yTicksR += `<text x="${(PAD.l+inner.w+6).toFixed(1)}" y="${y}" text-anchor="start" dominant-baseline="middle" font-size="8" fill="#999">${v}%</text>`;
            }
        }

        // Grid lines
        let grid = '';
        for (let i = 0; i <= yTickCount; i++) {
            const v = yMin + (yRange * i / yTickCount);
            const y = tyL(v).toFixed(1);
            grid += `<line x1="${PAD.l.toFixed(1)}" y1="${y}" x2="${(PAD.l+inner.w).toFixed(1)}" y2="${y}" stroke="#ddd" stroke-width="0.5"/>`;
        }

        // Legend
        const legend = [
            opPts.length ? `<line x1="0" y1="5" x2="16" y2="5" stroke="#388e3c" stroke-width="1.5" stroke-dasharray="2,2"/><text x="20" y="9" font-size="8" fill="#388e3c">OP</text>` : '',
            spPts.length ? `<line x1="${opPts.length ? 50 : 0}" y1="5" x2="${opPts.length ? 66 : 16}" y2="5" stroke="#f57c00" stroke-width="1.5" stroke-dasharray="4,2"/><text x="${opPts.length ? 70 : 20}" y="9" font-size="8" fill="#f57c00">SP</text>` : '',
            pvPts.length ? `<line x1="${opPts.length && spPts.length ? 100 : spPts.length ? 50 : 0}" y1="5" x2="${opPts.length && spPts.length ? 116 : spPts.length ? 66 : 16}" y2="5" stroke="#2d5a9e" stroke-width="1.5"/><text x="${opPts.length && spPts.length ? 120 : spPts.length ? 70 : 20}" y="9" font-size="8" fill="#2d5a9e">PV</text>` : '',
        ].filter(Boolean).join('');

        // Crosshair overlay (interactive, managed via JS after render)
        const overlayId = 'inu-pid-overlay-' + Date.now();
        const crossId   = 'inu-pid-cross-' + Date.now();

        const svg = `<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg" id="${overlayId}">
  <rect x="${PAD.l}" y="${PAD.t}" width="${inner.w}" height="${inner.h}" fill="none" stroke="#ccc" stroke-width="0.5"/>
  ${grid}
  ${xTicks}${yTicksL}${yTicksR}
  ${polyline(opPlot, tyR, '#388e3c', '2,2')}
  ${polyline(spPlot, tyL, '#f57c00', '4,2')}
  ${polyline(pvPlot, tyL, '#2d5a9e', '')}
  <g transform="translate(${(PAD.l+inner.w-160).toFixed(0)},${PAD.t})">${legend}</g>
  <line id="${crossId}" x1="0" y1="${PAD.t}" x2="0" y2="${PAD.t+inner.h}" stroke="rgba(0,0,0,.4)" stroke-width="1" display="none"/>
  <rect x="${PAD.l}" y="${PAD.t}" width="${inner.w}" height="${inner.h}" fill="transparent" class="inu-pid-hit"/>
</svg>`;
        container.innerHTML = svg;

        // Tooltip + crosshair interactivity
        const tooltip = document.createElement('div');
        tooltip.className = 'inu-pid-tooltip';
        container.style.position = 'relative';
        container.appendChild(tooltip);

        const svgEl    = container.querySelector(`#${overlayId}`);
        const crossEl  = container.querySelector(`#${crossId}`);
        const hitEl    = container.querySelector('.inu-pid-hit');

        function interp(pts, ts) {
            if (!pts.length) return null;
            let lo = pts[0], hi = pts[pts.length - 1];
            for (let i = 0; i < pts.length - 1; i++) {
                if (pts[i].ts <= ts && pts[i+1].ts >= ts) { lo = pts[i]; hi = pts[i+1]; break; }
            }
            if (lo === hi) return lo.val;
            const t = (ts - lo.ts) / (hi.ts - lo.ts);
            return lo.val + t * (hi.val - lo.val);
        }

        hitEl.addEventListener('mousemove', e => {
            const rect = svgEl.getBoundingClientRect();
            const xRel = e.clientX - rect.left;
            const xFrac = (xRel - (rect.width * PAD.l / W)) / (rect.width * inner.w / W);
            if (xFrac < 0 || xFrac > 1) { crossEl.setAttribute('display','none'); tooltip.style.display='none'; return; }
            const ts = tMin + xFrac * tRange;
            const xSvg = (PAD.l + xFrac * inner.w).toFixed(1);
            crossEl.setAttribute('x1', xSvg); crossEl.setAttribute('x2', xSvg); crossEl.setAttribute('display','');
            const pvV = interp(pvPts, ts), spV = interp(spPts, ts), opV = interp(opPts, ts);
            const elapsed = (ts - tMin) / 1000;
            const timeLbl = elapsed < 60 ? `${elapsed.toFixed(0)}s` : elapsed < 3600 ? `${(elapsed/60).toFixed(1)}m` : `${(elapsed/3600).toFixed(2)}h`;
            let lines = `t: ${timeLbl}`;
            if (pvV !== null) lines += `  PV: ${pvV.toFixed(2)}`;
            if (spV !== null) lines += `  SP: ${spV.toFixed(2)}`;
            if (opV !== null) lines += `  OP: ${opV.toFixed(1)}%`;
            tooltip.textContent = lines;
            tooltip.style.display = 'block';
            const tipLeft = Math.min(e.clientX - rect.left + 8, rect.width - tooltip.offsetWidth - 4);
            tooltip.style.left = tipLeft + 'px';
            tooltip.style.top  = '4px';
        });
        hitEl.addEventListener('mouseleave', () => {
            crossEl.setAttribute('display','none');
            tooltip.style.display = 'none';
        });
    }

    // Nelder-Mead simplex minimiser — minimises fn over x0.length dimensions
    function nelderMead(fn, x0, maxIter, tol) {
        const n = x0.length;
        const alpha = 1.0, gamma = 2.0, rho = 0.5, sigma = 0.5;
        const simplex = [{ x: x0.slice(), f: fn(x0) }];
        for (let i = 0; i < n; i++) {
            const xi = x0.slice();
            xi[i] += (Math.abs(xi[i]) < 1e-8 ? 0.05 : xi[i] * 0.05);
            simplex.push({ x: xi, f: fn(xi) });
        }
        for (let iter = 0; iter < maxIter; iter++) {
            simplex.sort((a, b) => a.f - b.f);
            if (Math.abs(simplex[n].f - simplex[0].f) < tol) break;
            const centroid = new Array(n).fill(0);
            for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) centroid[j] += simplex[i].x[j];
            for (let j = 0; j < n; j++) centroid[j] /= n;
            const xr = centroid.map((c, j) => c + alpha * (c - simplex[n].x[j]));
            const fr = fn(xr);
            if (fr < simplex[0].f) {
                const xe = centroid.map((c, j) => c + gamma * (xr[j] - c));
                const fe = fn(xe);
                simplex[n] = fe < fr ? { x: xe, f: fe } : { x: xr, f: fr };
            } else if (fr < simplex[n - 1].f) {
                simplex[n] = { x: xr, f: fr };
            } else {
                const xc = centroid.map((c, j) => c + rho * (simplex[n].x[j] - c));
                const fc = fn(xc);
                if (fc < simplex[n].f) {
                    simplex[n] = { x: xc, f: fc };
                } else {
                    const best = simplex[0].x;
                    for (let i = 1; i <= n; i++) {
                        for (let j = 0; j < n; j++) simplex[i].x[j] = best[j] + sigma * (simplex[i].x[j] - best[j]);
                        simplex[i].f = fn(simplex[i].x);
                    }
                }
            }
        }
        simplex.sort((a, b) => a.f - b.f);
        return simplex[0].x;
    }

    // Least-squares FOPDT fit via coarse grid search + Nelder-Mead refinement.
    // Model: pv(t) = baseline + kp*opStepSize*(1-exp(-(t-L)/T)) for t>=L, else baseline.
    // Returns {kp, L, T} or null if fit is infeasible.
    function fitFopdt(pvPts, opStepTs, opStepSize, baseline) {
        const pts = pvPts.filter(p => p.ts >= opStepTs).map(p => ({ t: (p.ts - opStepTs) / 1000, val: p.val }));
        if (pts.length < 5) return null;
        const timeWindow = pts[pts.length - 1].t;
        if (timeWindow < 5) return null;
        // Determine Kp sign from last 20% of post-step data
        const tail = pts.slice(Math.floor(pts.length * 0.8));
        const pvFinal = tail.reduce((s, p) => s + p.val, 0) / tail.length;
        const kpSign = Math.sign((pvFinal - baseline) / opStepSize) || 1;
        // Parameter bounds
        const Lmin = 0, Lmax = 0.25 * timeWindow;
        const Tmin = 0.5, Tmax = 2 * timeWindow;
        const kpMagMin = 0.001;
        // Dynamic upper bound: observed |ΔPV / ΔOP| × 3, floored at 10
        const pvRange = pts.reduce((mx, p) => Math.max(mx, Math.abs(p.val - baseline)), 0);
        const kpMagMax = Math.max(10, 3 * pvRange / Math.abs(opStepSize));
        function sse(params) {
            const [L, T, kpMag] = params;
            if (L < Lmin || L > Lmax || T < Tmin || T > Tmax || kpMag < kpMagMin || kpMag > kpMagMax) return 1e30;
            const gain = kpSign * kpMag * opStepSize;
            let sum = 0;
            for (const p of pts) {
                const sim = p.t < L ? baseline : baseline + gain * (1 - Math.exp(-(p.t - L) / T));
                const r = p.val - sim;
                sum += r * r;
            }
            return sum;
        }
        // Coarse grid search 10×10×5 = 500 evaluations
        let bestSSE = Infinity, bestParams = null;
        for (let iL = 0; iL < 10; iL++) {
            const L = Lmin + (Lmax - Lmin) * (iL + 0.5) / 10;
            for (let iT = 0; iT < 10; iT++) {
                const T = Tmin + (Tmax - Tmin) * (iT + 0.5) / 10;
                for (let iK = 0; iK < 5; iK++) {
                    const kpMag = kpMagMin + (kpMagMax - kpMagMin) * (iK + 0.5) / 5;
                    const s = sse([L, T, kpMag]);
                    if (s < bestSSE) { bestSSE = s; bestParams = [L, T, kpMag]; }
                }
            }
        }
        if (!bestParams) return null;
        // Nelder-Mead refinement
        const refined = nelderMead(sse, bestParams, 50, 1e-6);
        return {
            kp: kpSign * Math.max(kpMagMin, Math.min(kpMagMax, refined[2])),
            L:  Math.max(Lmin, Math.min(Lmax, refined[0])),
            T:  Math.max(Tmin, Math.min(Tmax, refined[1])),
        };
    }

    function computePidMetrics(pvPts, spPts, opPts, deadbandEu) {
        opPts = opPts || [];
        if (pvPts.length < 2) return null;
        const m = { stepDetected: false, opStepDetected: false, stepSize: 0, stepTime: 0, overshoot: null,
                    riseTime: null, settlingTime: null, oscillationCount: null, steadyStateError: null,
                    steadyStateUnreliable: false, pvUnresponsive: false, processModel: null,
                    dataGaps: 0, hasLargeGap: false };

        // Data gap detection — flag gaps > 3× median sampling interval
        if (pvPts.length >= 3) {
            const intervals = pvPts.slice(1).map((p, i) => p.ts - pvPts[i].ts);
            const sorted = intervals.slice().sort((a, b) => a - b);
            const medianInterval = sorted[Math.floor(sorted.length / 2)];
            const gapThreshold = medianInterval * 3;
            m.dataGaps = intervals.filter(d => d > gapThreshold).length;
            m.hasLargeGap = m.dataGaps > 0;
            m.medianIntervalMs = medianInterval;
        }

        // Interpolate SP at each PV timestamp — Zero-Order Hold (step-wise)
        // BMS setpoints are step changes, not ramps. Linear interpolation between two logged
        // SP values would fabricate phantom ramps that never existed, corrupting error metrics.
        function interpSp(ts) {
            if (ts <= spPts[0].ts) return spPts[0].val;
            if (ts >= spPts[spPts.length-1].ts) return spPts[spPts.length-1].val;
            for (let i = 0; i < spPts.length - 1; i++) {
                if (spPts[i].ts <= ts && spPts[i+1].ts > ts) return spPts[i].val; // hold until next step
            }
            return spPts[spPts.length-1].val;
        }

        // SP range for thresholds — reduce avoids stack overflow on large arrays
        const spVals = spPts.map(p => p.val);
        const spMin  = spVals.reduce((a, b) => a < b ? a : b, Infinity);
        const spMax  = spVals.reduce((a, b) => a > b ? a : b, -Infinity);
        const spRange = spMax - spMin || 1;

        // Step detection: largest absolute SP change
        let maxDelta = 0, stepIdx = -1;
        for (let i = 1; i < spPts.length; i++) {
            const delta = Math.abs(spPts[i].val - spPts[i-1].val);
            if (delta > maxDelta) { maxDelta = delta; stepIdx = i; }
        }
        if (maxDelta > spRange * 0.05 && stepIdx > 0) {
            m.stepDetected = true;
            m.stepTime = spPts[stepIdx].ts;
            m.stepSize = spPts[stepIdx].val - spPts[stepIdx-1].val;
            const spFinal = spPts[spPts.length-1].val;
            const stepDir = m.stepSize > 0 ? 1 : -1;

            // Post-step PV samples
            const postPv = pvPts.filter(p => p.ts >= m.stepTime);
            if (postPv.length >= 2) {
                // Overshoot: max deviation in direction of step beyond spFinal
                const deviations = postPv.map(p => stepDir * (p.val - spFinal));
                const maxDev = deviations.reduce((a, b) => a > b ? a : b, -Infinity);
                m.overshoot = Math.max(0, maxDev) / Math.abs(m.stepSize) * 100;

                // Rise time: first time PV crosses 90% of step target
                const target90 = (stepIdx > 0 ? spPts[stepIdx-1].val : spFinal - m.stepSize) + m.stepSize * 0.9;
                const risen = postPv.find(p => stepDir * (p.val - target90) >= 0);
                if (risen) m.riseTime = (risen.ts - m.stepTime) / 1000;

                // Settling time: last time |PV - spFinal| exceeds settling band
                // 2% of spRange is below sensor noise for constant SP (range defaults to 1 → 0.02°).
                // Use deadband if known, else 5% of step size minimum 0.5 EU.
                const band = deadbandEu != null ? deadbandEu
                           : Math.max(0.5, Math.abs(m.stepSize) * 0.05);
                let lastOutside = null;
                postPv.forEach(p => { if (Math.abs(p.val - spFinal) > band) lastOutside = p.ts; });
                if (lastOutside !== null) m.settlingTime = (lastOutside - m.stepTime) / 1000;
                else if (postPv.length) m.settlingTime = (postPv[0].ts - m.stepTime) / 1000; // already settled

                // Oscillation: zero-crossings of (PV - spFinal) after step
                let crossings = 0, prevSign = null;
                postPv.forEach(p => {
                    const sign = Math.sign(p.val - spFinal);
                    if (prevSign !== null && sign !== 0 && sign !== prevSign) crossings++;
                    if (sign !== 0) prevSign = sign;
                });
                m.oscillationCount = Math.floor(crossings / 2);
            }
        }

        // Steady-state error: last 20% of window regardless of step
        const winEnd   = pvPts[pvPts.length-1].ts;
        const winBegin = pvPts[0].ts;
        const windowStart = winBegin + (winEnd - winBegin) * 0.8;
        const tail = pvPts.filter(p => p.ts >= windowStart);
        // Check if OP changed significantly within the last 20% — if so, process is mid-transient
        // and steady-state error is meaningless
        let opChangedInTail = false;
        if (opPts.length >= 2) {
            const tailOp = opPts.filter(p => p.ts >= windowStart);
            if (tailOp.length >= 2) {
                const tailOpVals = tailOp.map(p => p.val);
                const opRange = tailOpVals.reduce((a, b) => a > b ? a : b, -Infinity)
                              - tailOpVals.reduce((a, b) => a < b ? a : b, Infinity);
                if (opRange >= 5) opChangedInTail = true;
            }
        }
        if (opChangedInTail) {
            m.steadyStateUnreliable = true;
        } else if (tail.length >= 2 && spPts.length >= 2) {
            const errors = tail.map(p => p.val - interpSp(p.ts));
            m.steadyStateError = errors.reduce((a, b) => a + b, 0) / errors.length;
            m.spRef = spPts[spPts.length - 1].val; // store for advice threshold calculations
        }

        // FOPDT process model from OP step (for open-loop ZN / Cohen-Coon identification)
        if (opPts.length >= 4) {
            let maxOpDelta = 0, opStepIdx = -1;
            for (let i = 1; i < opPts.length; i++) {
                const delta = Math.abs(opPts[i].val - opPts[i-1].val);
                if (delta > maxOpDelta) { maxOpDelta = delta; opStepIdx = i; }
            }
            if (maxOpDelta >= 5 && opStepIdx > 0) { // require ≥5% OP step
                const opStepSize = opPts[opStepIdx].val - opPts[opStepIdx-1].val;
                const opStepTs   = opPts[opStepIdx].ts;
                // Validate step is sustained: OP must stay near new level for ≥3 subsequent samples.
                // A single-sample spike (glitch, mis-read) would otherwise anchor the entire FOPDT model.
                const opNewLevel  = opPts[opStepIdx].val;
                const postOpCheck = opPts.slice(opStepIdx + 1, opStepIdx + 4);
                const stepSustained = postOpCheck.length >= 2 &&
                    postOpCheck.every(p => Math.abs(p.val - opNewLevel) < maxOpDelta * 0.5);
                if (stepSustained) {
                    const pvBefore = pvPts.filter(p => p.ts < opStepTs);
                    const pvAfter  = pvPts.filter(p => p.ts >= opStepTs);
                    if (pvBefore.length >= 2 && pvAfter.length >= 5) {
                        // Baseline: use samples within 60s before the OP step (time-based window)
                        const baselineWin = pvBefore.filter(p => p.ts >= opStepTs - 60000);
                        const basePts  = baselineWin.length >= 2 ? baselineWin : pvBefore;
                        const baseline = basePts.reduce((s, p) => s + p.val, 0) / basePts.length;

                        // Least-squares FOPDT fit (replaces graphical 63.2% method)
                        const fit = fitFopdt(pvPts, opStepTs, opStepSize, baseline);
                        if (fit) {
                            const processKp = fit.kp;   // EU per %OP
                            const deadTime  = fit.L;    // seconds
                            const timeConst = fit.T;    // seconds
                            m.processModel = {
                                kp: +processKp.toFixed(4),
                                L:  +deadTime.toFixed(2),
                                T:  +timeConst.toFixed(1),
                                opStep: opStepSize,
                            };
                            m.opStepDetected = true;

                            // R² over 3× time-constant window using the fitted model
                            const r2Window = pvAfter.filter(p => p.ts <= opStepTs + 3 * timeConst * 1000);
                            if (r2Window.length >= 2) {
                                const meanActual = r2Window.reduce((s, p) => s + p.val, 0) / r2Window.length;
                                const gain = processKp * opStepSize;
                                let ssRes = 0, ssTot = 0;
                                for (const p of r2Window) {
                                    const t = (p.ts - opStepTs) / 1000;
                                    const pvSim = t <= deadTime
                                        ? baseline
                                        : baseline + gain * (1 - Math.exp(-(t - deadTime) / timeConst));
                                    ssRes += (p.val - pvSim) ** 2;
                                    ssTot += (p.val - meanActual) ** 2;
                                }
                                m.processModel.r2 = ssTot < 0.001 ? null : +(1 - ssRes / ssTot).toFixed(3);
                            }

                            // Process direction validation
                            if (processKp < 0) m.processModel.reverseActing = true;
                        }
                    }
                }
            }
        }

        // Oscillation detection — runs on full window around mean SP (no step required)
        // Also runs post-step if a step was detected
        if (spPts.length >= 2) {
            const analysePts = m.stepDetected ? pvPts.filter(p => p.ts >= m.stepTime) : pvPts;
            const spRef = m.stepDetected
                ? spPts[spPts.length - 1].val
                : spPts.reduce((s, p) => s + p.val, 0) / spPts.length; // mean SP

            // Always compute PV peak-to-peak swing and mean for display (independent of oscillation detection)
            const pvVals = analysePts.map(p => p.val);
            const pvMax = pvVals.reduce((a, b) => a > b ? a : b, -Infinity);
            const pvMin = pvVals.reduce((a, b) => a < b ? a : b, Infinity);
            const pvAmplitude = (pvMax - pvMin) / 2;
            m.pvSwingPP = +(pvMax - pvMin).toFixed(3);
            m.pvMean    = +(pvVals.reduce((a, b) => a + b, 0) / pvVals.length).toFixed(3);
            const spVal = Math.abs(spRef) || 1;

            // Only flag oscillation if amplitude exceeds deadband (if known) or minimum threshold
            const minAmplitude = deadbandEu != null ? deadbandEu : Math.max(0.1, spVal * 0.005);

            // When the PV mean is significantly offset from SP (steady-state error), zero-crossings
            // against spRef will be infrequent and oscillations go undetected. Use pvMean as the
            // crossing reference in that case so we detect oscillations around the actual operating point.
            const pvMeanVal = m.pvMean;
            const ssOffset  = Math.abs(pvMeanVal - spRef);
            const oscRef    = ssOffset > minAmplitude * 1.5 ? pvMeanVal : spRef;

            const crossTs = [];
            let prev = null;
            analysePts.forEach(p => {
                const dev = p.val - oscRef;
                if (prev !== null && prev.dev !== 0 && Math.sign(dev) !== Math.sign(prev.dev)) {
                    const frac = Math.abs(prev.dev) / (Math.abs(prev.dev) + Math.abs(dev));
                    crossTs.push(prev.ts + frac * (p.ts - prev.ts));
                }
                if (dev !== 0) prev = { dev, ts: p.ts };
            });

            if (crossTs.length >= 4 && pvAmplitude > minAmplitude) {
                const halfPeriods = crossTs.slice(1).map((t, i) => t - crossTs[i]);
                // Filter out wildly inconsistent half-periods (outliers from noise)
                const medHalf = halfPeriods.slice().sort((a,b)=>a-b)[Math.floor(halfPeriods.length/2)];
                const consistent = halfPeriods.filter(h => h > medHalf * 0.3 && h < medHalf * 3);
                if (consistent.length >= 3) {
                    const avgHalf = consistent.reduce((a, b) => a + b, 0) / consistent.length;
                    // CV computed from full halfPeriods (not filtered) — gives honest regularity measure.
                    // Using filtered set would hide irregularity by discarding the worst offenders.
                    const meanAll = halfPeriods.reduce((a, b) => a + b, 0) / halfPeriods.length;
                    const varAll  = halfPeriods.reduce((s, h) => s + (h - meanAll) ** 2, 0) / halfPeriods.length;
                    const cvHalf  = Math.sqrt(varAll) / meanAll;
                    m.ultimatePeriod = (2 * avgHalf) / 1000; // seconds — uses filtered set for robustness
                    m.oscillationCount = Math.floor(crossTs.length / 2);
                    m.oscillationAmplitude = +pvAmplitude.toFixed(3);
                    m.periodCV = +cvHalf.toFixed(3); // >0.35 = irregular, likely process disturbance

                    // C1: Amplitude trend via windowed RMS regression — replaces thirds peak-to-peak.
                    // Thirds method is unreliable on noisy data (one outlier dominates pp within a chunk).
                    // Instead: 6 overlapping windows, compute RMS of (PV - oscRef) in each, fit linear
                    // regression on window index. Classify by normalized slope.
                    const numWindows = Math.min(6, Math.floor(analysePts.length / 4));
                    if (numWindows >= 4) {
                        const winSize = Math.floor(analysePts.length / numWindows);
                        const rmsVals = [];
                        for (let w = 0; w < numWindows; w++) {
                            const start = Math.floor((w / numWindows) * (analysePts.length - winSize));
                            const win   = analysePts.slice(start, start + winSize);
                            const rms   = Math.sqrt(win.reduce((s, p) => s + (p.val - oscRef) ** 2, 0) / win.length);
                            rmsVals.push(rms);
                        }
                        const n = rmsVals.length;
                        const meanX = (n - 1) / 2;
                        const meanY = rmsVals.reduce((a, b) => a + b, 0) / n;
                        let num = 0, den = 0;
                        for (let i = 0; i < n; i++) { num += (i - meanX) * (rmsVals[i] - meanY); den += (i - meanX) ** 2; }
                        const slope     = den > 0 ? num / den : 0;
                        const relSlope  = meanY > 0 ? slope / meanY : 0;
                        if      (relSlope >  0.15) m.amplitudeTrend = 'growing';
                        else if (relSlope < -0.15) m.amplitudeTrend = 'decaying';
                        else                       m.amplitudeTrend = 'constant';
                    }
                }
            }
        }

        // Actuator saturation detection — if OP is pinned at 0% or 100% for >30% of the window,
        // the loop is saturated and FOPDT / metrics analysis is unreliable.
        if (opPts.length >= 4) {
            const satHigh = opPts.filter(p => p.val >= 99).length / opPts.length;
            const satLow  = opPts.filter(p => p.val <= 1).length  / opPts.length;
            if (satHigh > 0.3) m.actuatorSaturated = 'high';
            else if (satLow > 0.3) m.actuatorSaturated = 'low';
        }

        // OP statistics — mean and peak-to-peak, used for controller responsiveness analysis
        if (opPts.length >= 2) {
            const opVals = opPts.map(p => p.val);
            const opMax  = opVals.reduce((a, b) => a > b ? a : b, -Infinity);
            const opMin  = opVals.reduce((a, b) => a < b ? a : b, Infinity);
            m.opMean    = +(opVals.reduce((a, b) => a + b, 0) / opVals.length).toFixed(1);
            m.opSwingPP = +(opMax - opMin).toFixed(1);
        }

        return m;
    }

    // Compute Ziegler-Nichols and Cohen-Coon tuning parameters from process model
    // Returns { openLoop: { zn, cc } } and/or { closedLoop: { zn } } when data is available
    function computeTuning(metrics, pidValues, pconv) {
        const result = {};
        const pm = metrics && metrics.processModel;

        // C2: Actuator saturated → all analysis is unreliable, suppress tuning entirely
        if (metrics?.actuatorSaturated) {
            result.saturated = true;
            return result;
        }

        if (pm && pm.T > 0 && pm.kp !== 0) {
            const Kp = Math.abs(pm.kp); // use absolute gain so reverse-acting processes get valid positive tuning
            const { L, T } = pm;
            const r = L / (T || 0.001); // L/T ratio

            if (pm.r2 !== null && pm.r2 !== undefined && pm.r2 < 0.70) {
                // FOPDT model fit too poor to trust — suppress tuning suggestions
                result.poorFit = true;
                result.r2 = pm.r2;
            } else {
                // Lambda/IMC: Kc = T/(Kp*(lambda+L)), Ti = T+L/2 (disturbance-rejection form)
                // IMC does NOT divide by L so it works even when L ≈ 0.
                const lambda = Math.max(2 * L, T / 3);
                const imc = {
                    pi: { kc: +(T / (Kp * (lambda + L))).toFixed(3), ti: +(T + L / 2).toFixed(1), td: null, lambda: +lambda.toFixed(1) },
                };

                if (L < 1) {
                    // C4: Dead time too small to measure — ZN/CC divide by L and produce garbage.
                    // IMC is still valid (lambda absorbs the near-zero L). Offer IMC only.
                    result.noDeadTime = true;
                    result.openLoop = { imc, model: { kp: +Kp.toFixed(4), L: +L.toFixed(1), T: +T.toFixed(1), ratio: +r.toFixed(3) } };
                    if (pm.reverseActing) result.openLoop.reverseActing = true;
                } else {
                    // Full ZN + CC + IMC
                    // Ziegler-Nichols open-loop (step response)
                    const zn = {
                        pid: { kc: +(1.2 * T / (Kp * L)).toFixed(3), ti: +(2 * L).toFixed(1),   td: +(0.5 * L).toFixed(1) },
                        pi:  { kc: +(0.9 * T / (Kp * L)).toFixed(3), ti: +(L / 0.3).toFixed(1), td: null },
                        p:   { kc: +(T      / (Kp * L)).toFixed(3),  ti: null,                   td: null },
                    };

                    // Cohen-Coon open-loop
                    const cc = {
                        pid: { kc: +((T / (Kp * L)) * (4/3 + r/4)).toFixed(3),
                               ti: +(L * (32 + 6*r) / (13 + 8*r)).toFixed(1),
                               td: +(4 * L / (11 + 2*r)).toFixed(1) },
                        pi:  { kc: +((T / (Kp * L)) * (9/10 + r/12)).toFixed(3),
                               ti: +(L * (30 + 3*r) / (9 + 20*r)).toFixed(1),
                               td: null },
                        p:   { kc: +((T / (Kp * L)) * (1 + r/3)).toFixed(3), ti: null, td: null },
                    };

                    result.openLoop = { zn, cc, imc, model: { kp: +Kp.toFixed(4), L: +L.toFixed(1), T: +T.toFixed(1), ratio: +r.toFixed(3) } };
                    if (pm.reverseActing) result.openLoop.reverseActing = true;
                }
            }
        }

        // Closed-loop ZN from ultimate period + estimated Ku
        // ONLY valid when oscillations are consistent (low period CV) — irregular periods indicate
        // process disturbances, not a true limit cycle, so Ku ≈ Kp assumption breaks down.
        if (metrics && metrics.ultimatePeriod && pidValues && pidValues.p !== null) {
            const Pu = metrics.ultimatePeriod;
            // Convert P-band (%) to gain before using as Ku — P-band convention stores 100/Kp
            const Ku = (pconv === 'pband' && pidValues.p > 0) ? 100 / pidValues.p : pidValues.p;
            const cv = metrics.periodCV ?? 0;
            // Mark as disturbance-driven if period is irregular (CV > 35%) or amplitude is not constant — ZN formula unreliable
            // Grey zone 0.25–0.35: oscillation regularity borderline, ZN should be treated with caution
            const disturbanceDriven = cv > 0.35 ||
                (metrics.amplitudeTrend && metrics.amplitudeTrend !== 'constant');
            const greyZone = !disturbanceDriven && cv > 0.25;
            const clZn = {
                pid: { kc: +(0.6  * Ku).toFixed(3), ti: +(Pu / 2).toFixed(1),   td: +(Pu / 8).toFixed(1) },
                pi:  { kc: +(0.45 * Ku).toFixed(3), ti: +(Pu / 1.2).toFixed(1), td: null },
                p:   { kc: +(0.5  * Ku).toFixed(3), ti: null,                    td: null },
            };
            result.closedLoop = { zn: clZn, Ku, Pu: +Pu.toFixed(1), cv: +cv.toFixed(3), disturbanceDriven, greyZone };
        }

        return result;
    }

    function getPidAdvice(metrics, circuitType, pidValues, pconv, tuning, deadbandEu) {
        const ct = PID_CIRCUIT_TYPES[circuitType];
        const advice = [];
        const isPband = pconv === 'pband';

        // Helper: format a Kp value in the user's convention for advice text
        function fmtP(kc) {
            if (kc == null) return '?';
            if (isPband) return `P% ≈ ${(100 / kc).toFixed(1)} (Kp ≈ ${kc})`;
            return `Kp ≈ ${kc}`;
        }
        // ZN PID suggestion shorthand if available
        const znPid = tuning?.openLoop?.zn?.pid;

        if (!metrics) return [{ severity: 'info', text: 'Inga mätvärden tillgängliga.' }];

        // Actuator saturation — analysis is unreliable when OP is pinned
        if (metrics.actuatorSaturated) {
            const dir = metrics.actuatorSaturated === 'high' ? 'maxläge (100%)' : 'minläge (0%)';
            advice.push({ severity: 'warn', text: `Aktuatorn verkar vara i ${dir} under mer än 30% av mätfönstret — regulatorn är sannolikt mättad. FOPDT-modell och mätvärden är opålitliga under mättning. Kontrollera börvärde, dimensionering och driftläge.` });
            advice.push({ severity: 'info', text: 'Kontrollera att regulatorn har anti-windup aktiverat (back-calculation eller clamping) — annars riskerar I-termen att vinda upp kraftigt under mättning.' });
        }

        // Process direction validation (Phase 2.6)
        if (metrics.processModel?.reverseActing) {
            advice.push({ severity: 'warn', text: 'Negativ processgain detekterad — regulatorn verkar vara omvänd-verkande (OP ökar → PV sjunker). Kontrollera aktionsriktning och regulatorns verkningsriktning innan du tillämpar inställningsförslag.' });
        }

        // Amplitude envelope warnings (Phase 3.1)
        if (metrics.amplitudeTrend === 'decaying') {
            advice.push({ severity: 'info', text: 'Dämpade oscillationer detekterade — svängtiden avtar, systemet stabiliserar sig. Inga justeringar krävs om felet är litet.' });
        } else if (metrics.amplitudeTrend === 'growing') {
            advice.push({ severity: 'warn', text: 'Växande oscillationer detekterade — systemet är instabilt. Sänk Kp omedelbart med 30–50%.' });
        }

        // Oscillation check — runs regardless of step detection
        if (metrics.oscillationCount !== null && metrics.oscillationCount >= 2) {
            const cl = tuning?.closedLoop;
            const clZn = cl?.zn?.pi;
            const ampStr = metrics.oscillationAmplitude != null ? ` (amplitud ±${metrics.oscillationAmplitude})` : '';
            const puStr  = metrics.ultimatePeriod != null ? `, Pu = ${metrics.ultimatePeriod.toFixed(1)} s` : '';
            if (cl?.disturbanceDriven) {
                // Irregular period — oscillations are likely external disturbances, not a limit cycle.
                // ZN closed-loop would just keep suggesting lower and lower values as you reduce Kp.
                // When deadband unknown, use 5% of SP as fallback threshold for "acceptable" SSE
                const ssThreshold = deadbandEu != null ? deadbandEu * 1.5
                    : (metrics.spRef ? Math.abs(metrics.spRef) * 0.05 : 2);
                const ssOk = metrics.steadyStateError != null && Math.abs(metrics.steadyStateError) < ssThreshold;
                const ssStr = metrics.steadyStateError != null ? ` Kvarstående fel: ${metrics.steadyStateError.toFixed(2)}.` : '';
                if (ssOk || metrics.steadyStateError === null) {
                    advice.push({ severity: 'info', text: `Oregelbunden oscillation detekterad${ampStr}${puStr} (periodens CV = ${(cl.cv*100).toFixed(0)}%) — troligen processstörning, inte regulatorinducerad.${ssStr} Nuvarande inställningar verkar acceptabla — undvik att sänka P ytterligare.` });
                } else {
                    advice.push({ severity: 'warn', text: `Oregelbunden oscillation detekterad${ampStr}${puStr} (periodens CV = ${(cl.cv*100).toFixed(0)}%) — troligen processstörning.${ssStr} Kvarstående fel är ändå högt — kontrollera I-tid och kalibrering. ZN sluten krets ej tillförlitlig.` });
                }
            } else if (clZn) {
                const greyNote = cl?.greyZone ? ` (periodens CV = ${(cl.cv*100).toFixed(0)}% — gränsfall, behandla med försiktighet)` : '';
                advice.push({ severity: 'warn', text: `Regelbunden oscillation detekterad${ampStr}${puStr}${greyNote} — ZN sluten krets rekommenderar ${fmtP(clZn.kc)}, Ti = ${clZn.ti} s. Se inställningsförslaget.` });
            } else {
                const dir = isPband ? 'öka P% och/eller öka I-tid' : 'sänk Kp och/eller öka I-tid';
                advice.push({ severity: 'warn', text: `Oscillation detekterad${ampStr}${puStr} — ${dir}.` });
            }
        }

        // Controller responsiveness check: if PV is swinging noticeably but OP is barely moving,
        // the controller is not responding — likely wrong mode, extreme deadband, or very conservative tuning.
        if (metrics.pvSwingPP != null && metrics.opSwingPP != null && metrics.pvSwingPP > 0) {
            const spRef = metrics.spRef != null ? Math.abs(metrics.spRef) || 1 : 1;
            const pvPct = metrics.pvSwingPP / spRef * 100;
            if (pvPct > 2 && metrics.opSwingPP < 3 && !metrics.actuatorSaturated) {
                advice.push({ severity: 'warn', text: `PV svänger ${metrics.pvSwingPP.toFixed(2)} enheter men OP varierar bara ${metrics.opSwingPP.toFixed(1)}% — regulatorn svarar inte på avvikelsen. Kontrollera att regulatorn är i Auto-läge, att inget handdriftsläge eller dödzon blockerar utgången, och att Kp/P-bandet inte är extremt konservativt.` });
            }
        }

        if (!metrics.stepDetected && !metrics.opStepDetected) {
            if (!metrics.oscillationCount) {
                advice.push({ severity: 'info', text: 'Inget stegrespons detekterat. För korrekt FOPDT-modell: sätt regulatorn i handdrift, stega OP manuellt 10–15% och låt PV stabilisera sig. Börvärdessteg ger slutet kretssvar och försämrar modellidentifieringen.' });
            }
        } else {
            if (metrics.overshoot !== null) {
                if (metrics.overshoot > 20) {
                    if (znPid) {
                        advice.push({ severity: 'warn', text: `Överskjutning hög (${metrics.overshoot.toFixed(0)}%) — ZN rekommenderar ${fmtP(znPid.kc)}, Ti = ${znPid.ti} s. Se inställningsförslaget.` });
                    } else {
                        const dir = isPband ? 'öka P% med ~15–25% (minskar förstärkning)' : 'sänk Kp med ~15–25%';
                        advice.push({ severity: 'warn', text: `Överskjutning hög (${metrics.overshoot.toFixed(0)}%) — ${dir}.` });
                    }
                } else if (metrics.overshoot < 3 && metrics.riseTime !== null && ct && metrics.riseTime > ct.responseTime[1]) {
                    const dir = isPband ? 'minska P% (ökar förstärkning) eller minska I-tid' : 'öka Kp eller minska I-tid';
                    advice.push({ severity: 'warn', text: `Trög insvängning (${metrics.riseTime.toFixed(0)} s, förväntat <${ct.responseTime[1]} s) — ${dir}.` });
                } else {
                    advice.push({ severity: 'ok', text: `Överskjutning acceptabel (${metrics.overshoot.toFixed(0)}%).` });
                }
            }
            if (metrics.settlingTime !== null && ct && metrics.settlingTime > ct.responseTime[1] * 3) {
                advice.push({ severity: 'warn', text: `Lång insvängningstid (${metrics.settlingTime.toFixed(0)} s) — kontrollera processfördröjning och I-tid.` });
            }
        }

        if (metrics.steadyStateUnreliable) {
            advice.push({ severity: 'info', text: 'Kvarstående fel kan ej bedömas — OP ändrades nära slutet av mätfönstret och processen har inte hunnit stabilisera sig.' });
        } else if (metrics.steadyStateError !== null) {
            const absErr = Math.abs(metrics.steadyStateError);
            const errThreshold = deadbandEu != null ? deadbandEu : 1;
            const dbNote = deadbandEu != null ? ` (dödzon: ±${deadbandEu.toFixed(2)})` : '';
            if (absErr > errThreshold) {
                advice.push({ severity: 'warn', text: `Kvarstående fel ${metrics.steadyStateError.toFixed(2)}${dbNote} — kontrollera I-tid (för lång?) eller kalibrering.` });
            } else {
                advice.push({ severity: 'ok', text: `Kvarstående fel litet (${metrics.steadyStateError.toFixed(2)})${dbNote}.` });
            }
        }

        if (ct && pidValues) {
            if (pidValues.i !== null && pidValues.i > ct.tiRange[1]) {
                advice.push({ severity: 'info', text: `I-tid ${pidValues.i} s är längre än typvärdet för ${ct.label} (${ct.tiRange[0]}–${ct.tiRange[1]} s).` });
            } else if (pidValues.i !== null && pidValues.i < ct.tiRange[0]) {
                advice.push({ severity: 'info', text: `I-tid ${pidValues.i} s är kortare än typvärdet för ${ct.label} (${ct.tiRange[0]}–${ct.tiRange[1]} s).` });
                // Anti-windup advisory when Ti is shorter than typical (higher risk of wind-up)
                advice.push({ severity: 'info', text: 'Kort I-tid ökar risken för integratormättning (windup) vid aktuatorbegränsning. Kontrollera att regulatorn har anti-windup (back-calculation eller clamping) aktiverat.' });
            }
            // P value check — always compare using Kp internally (pidValues.p is already normalised to Kp)
            if (pidValues.p !== null && pidValues.p > ct.kpRange[1] * 2) {
                const typPband = `P%: ${(100/ct.kpRange[1]).toFixed(0)}–${(100/ct.kpRange[0]).toFixed(0)}`;
                const typKp    = `Kp: ${ct.kpRange[0]}–${ct.kpRange[1]}`;
                const hasTuning = znPid != null;
                if (isPband) {
                    advice.push({ severity: 'warn', text: `P% = ${(100/pidValues.p).toFixed(2)} (Kp = ${pidValues.p.toFixed(1)}) är aggressivt för ${ct.label} (typvärde ${typPband}).${hasTuning ? ' Se inställningsförslaget.' : ''}` });
                } else {
                    advice.push({ severity: 'warn', text: `Kp = ${pidValues.p.toFixed(1)} är hög för ${ct.label} (typvärde ${typKp}).${hasTuning ? ' Se inställningsförslaget.' : ''}` });
                }
            }
        }

        // Circuit-type specific advice
        if (circuitType === 'humidity') {
            advice.push({ severity: 'info', text: 'Fuktighetskretsar är tröga och känsliga för mätarbrus — undvik Ti under 400 s. Fuktighetssensorer kan driva med tiden; verifiera kalibrering mot referensmätare minst en gång per år.' });
            if (pidValues?.i != null && pidValues.i < 400) {
                advice.push({ severity: 'warn', text: `Ti = ${pidValues.i} s är aggressivt för fuktighetsstyrning (typvärde 400–1200 s). Kort Ti riskerar att skapa oscillationer kring börvärdet. Öka Ti till minst 400 s.` });
            }
            advice.push({ severity: 'info', text: 'Vid splitreglering (separat befuktare och avfuktare med samma sensor): säkerställ att det finns ett dödband på minst ±3–5 %RH mellan befuktarens och avfuktarens börvärden, annars motarbetar regulatorerna varandra. Aktivera anti-windup (back-calculation eller clamping) på båda regulatorerna — den inaktiva parten vinar annars upp och ger kraftiga överskott vid aktivering.' });
        }

        if (circuitType === 'split_range') {
            advice.push({ severity: 'info', text: 'Splitreglering: OP 0–50% styr typiskt kylsida, OP 50–100% värmesida (eller omvänt — verifiera med driftdokumentation). FOPDT-modellen identifieras på det aktiva OP-området. Kör separata stegresponstester för kyl- och värmesidan om processen beter sig asymmetriskt.' });
            advice.push({ severity: 'info', text: 'Säkerställ att det finns anti-windup på regulatorn — vid splittning kan regulatorn fasta nära 50%-övergången och vina upp om deadband saknas. Lägg till ett litet deadband (1–3%) runt övergångspunkten för att undvika jakt.' });
            if (metrics?.processModel?.reverseActing == null) {
                advice.push({ severity: 'info', text: 'Kontrollera aktionsriktning: i splitreglering kan processförstärkning vara positiv på ena sidan och negativ på den andra. Analysera varje sida separat vid behov.' });
            }
        }

        if (circuitType === 'pump_pressure') {
            advice.push({ severity: 'info', text: 'Differenstrycksreglering för pump är en snabb process — Ti under 30 s kan ge oscillationer om rörledningssystemet har resonans. Börja med Ti = 60 s och minska om insvängningstiden är acceptabel.' });
        }

        // Sampling interval adequacy check
        // FOPDT identification requires several samples during the process transient.
        // Rule of thumb: interval must be ≤ responseTime[0] / 3 for reliable identification.
        const medIntervalS = metrics.medianIntervalMs ? metrics.medianIntervalMs / 1000 : null;
        if (medIntervalS != null && ct) {
            const fineThreshold = ct.responseTime[0] / 3;
            const coarseThreshold = ct.responseTime[0];
            if (medIntervalS > coarseThreshold) {
                advice.push({ severity: 'warn', text: `Loggningsintervall ~${Math.round(medIntervalS)} s är för grovt för ${ct.label} (typisk svarstid ${ct.responseTime[0]}–${ct.responseTime[1]} s) — FOPDT-modell och insvängningstider är opålitliga. Processmodellen kan inte identifieras korrekt när intervallet är längre än svarstiden. Hämta data med tätare loggning (idealiskt under ${Math.round(fineThreshold)} s) för tillförlitliga inställningsförslag.` });
            } else if (medIntervalS > fineThreshold) {
                advice.push({ severity: 'info', text: `Loggningsintervall ~${Math.round(medIntervalS)} s är marginellt för ${ct.label} (idealt under ${Math.round(fineThreshold)} s). FOPDT-modellen kan ha reducerad noggrannhet — tolka Kp, L och T som uppskattningar snarare än exakta värden.` });
            }
        }

        if (!advice.length) advice.push({ severity: 'ok', text: 'Reglerkretsen verkar stabil — inga direkta åtgärder behövs.' });
        return advice;
    }

    // Get all tag names — uses GetTagList API (server-side, no pagination limit)
    // Also caches tag metadata (engmin/engmax) keyed by name for deadband calculations
    let _allTagNamesCache = null;
    let _allTagMetaCache  = {}; // name → { engmin, engmax }
    async function getAllTagNames() {
        if (_allTagNamesCache) return _allTagNamesCache;
        try {
            const sid = (location.search.match(/sid=([^&#]+)/) || [])[1] || '';
            const r = await fetch('/tag/GetTagList?sid=' + sid + '&draw=1&limit=9999&offset=0&sortcol=0&sortdir=asc&search=');
            if (r.ok) {
                const json = await r.json();
                const dec = h => { const d = document.createElement('div'); d.innerHTML = h; return d.textContent.trim(); };
                const names = [];
                for (const row of (json.data || [])) {
                    const name = dec(row['0'] || '');
                    if (!name) continue;
                    names.push(name);
                    _allTagMetaCache[name] = {
                        engmin: parseFloat(row['6']) || 0,
                        engmax: parseFloat(row['7']) || 0,
                    };
                }
                if (names.length) { _allTagNamesCache = names; return names; }
            }
        } catch { /* fall through */ }
        // Fallback: visible rows only
        const fallback = [...document.querySelectorAll('#tagtable tbody tr.tag')]
            .map(r => r.cells[CFG.colOffset]?.textContent?.trim()).filter(Boolean);
        if (fallback.length) _allTagNamesCache = fallback;
        return fallback;
    }

    function showPidAdvisor() {
        injectPidStyles();
        const lastSensorPfx = GM_getValue('pid_sensor_prefix', GM_getValue('pid_prefix', ''));
        const lastDevicePfx = GM_getValue('pid_device_prefix', '');
        const lastCircuit = GM_getValue('pid_circuit', 'supply_temp');
        const lastWindow  = GM_getValue('pid_window', '3600000');
        const lastPConv   = GM_getValue('pid_pconv', 'pband'); // 'pband' = P% | 'gain' = Kp

        const circuitOpts = Object.entries(PID_CIRCUIT_TYPES)
            .map(([k, v]) => `<option value="${k}"${k === lastCircuit ? ' selected' : ''}>${v.label}</option>`).join('');
        const windowOpts = [
            { v: '300000',    l: '5 min'     },
            { v: '600000',    l: '10 min'    },
            { v: '900000',    l: '15 min'    },
            { v: '1800000',   l: '30 min'    },
            { v: '3600000',   l: '1 timme'   },
            { v: '21600000',  l: '6 timmar'  },
            { v: '86400000',  l: '24 timmar' },
            { v: '604800000', l: '7 dagar'   },
        ].map(o => `<option value="${o.v}"${o.v === lastWindow ? ' selected' : ''}>${o.l}</option>`).join('');

        const mo = modal(`<div class="inu-pid-modal">
<h3><i class="fa fa-line-chart"></i> PID-tuner (Experimentell)</h3>
<div id="inu-pid-setup">
  <div class="inu-pid-row">
    <label>Sensor</label>
    <input type="text" id="inu-pid-sensor-prefix" value="${escHtml(lastSensorPfx)}" placeholder="T.ex. LB05_GP101 (ger PV, SP, P, I, D)" style="flex:2">
    <button id="inu-pid-search" style="flex-shrink:0"><i class="fa fa-search"></i> Sök</button>
  </div>
  <div class="inu-pid-row" id="inu-pid-device-row" style="display:${lastDevicePfx ? '' : 'none'}">
    <label>Enhet</label>
    <input type="text" id="inu-pid-device-prefix" value="${escHtml(lastDevicePfx)}" placeholder="T.ex. LB05_FT101 (ger OP) — lämna tomt om samma" style="flex:2">
    <button id="inu-pid-device-search" style="flex-shrink:0"><i class="fa fa-search"></i> Sök</button>
  </div>
  <div class="inu-pid-row">
    <label>Kretstyp</label>
    <select id="inu-pid-circuit">${circuitOpts}</select>
    <label style="min-width:50px">Fönster</label>
    <select id="inu-pid-window">${windowOpts}</select>
  </div>
  <div id="inu-pid-pconv-notice" style="display:none;font-size:10px;color:#888;margin:-4px 0 4px 0"></div>
  <input type="hidden" id="inu-pid-pconv" value="${escHtml(lastPConv)}">
  <div class="inu-pid-tags" id="inu-pid-taglist">Ange ett sensorprefix och klicka Sök.</div>
  <div id="inu-pid-win-warn" style="display:none;color:#e67e22;font-size:10px;margin-top:4px;padding:4px 6px;background:#fff3e0;border-radius:4px;border:1px solid #ffe082"></div>
  <div class="inu-pid-btn-row" style="margin-top:10px">
    <button class="bx" id="inu-pid-cancel">Stäng</button>
    <button class="bok" id="inu-pid-fetch" disabled><i class="fa fa-download"></i> Hämta data</button>
  </div>
  <details class="inu-pid-hint">
    <summary><i class="fa fa-info-circle"></i> Hur gör jag en korrekt mätning?</summary>
    <ol>
      <li>Sätt regulatorn i <b>handdrift</b> (manuellt läge).</li>
      <li>Vänta 20–30 s tills PV stabiliserar sig — se livevärden ovan.</li>
      <li>Klicka <b>Hämta data</b> för att starta datainsamlingen.</li>
      <li>Stega OP manuellt med <b>10–15%</b> (t.ex. 60% → 75%).</li>
      <li>Vänta tills PV har stabiliserat sig — klicka sedan <b>Analysera nu</b>.</li>
    </ol>
    <p>Om historisk trenddata finns tillgänglig hämtas den direkt och steg 3–5 behövs inte — analysen körs automatiskt. Börvärdessteg med regulatorn i auto ger sämre modellidentifiering.</p>
  </details>
  <details class="inu-pid-hint" style="margin-top:4px">
    <summary><i class="fa fa-cog"></i> P-konvention</summary>
    <div style="padding:6px 0 2px">
      <label style="font-size:11px;color:#888;display:flex;gap:8px;align-items:center">
        P-enhet:
        <select id="inu-pid-pconv-sel">
          <option value="gain"${lastPConv === 'gain'  ? ' selected' : ''}>Kp (förstärkning)</option>
          <option value="pband"${lastPConv === 'pband' ? ' selected' : ''}>P% (band)</option>
        </select>
        <span style="color:#aaa;font-weight:normal">(detekteras automatiskt från P-taggen)</span>
      </label>
    </div>
  </details>
</div>
<div id="inu-pid-results" style="display:none"></div>
</div>`);

        // Cleanup registry: all timers/intervals register here; fired on modal dismiss (any path)
        const _cleanups = [];
        function onModalDismiss(fn) { _cleanups.push(fn); }
        function runCleanups() { while (_cleanups.length) { try { _cleanups.pop()(); } catch {} } }
        // Watch for modal removal from DOM (covers background click, .remove(), etc.)
        const _moObserver = new MutationObserver(() => {
            if (!mo.parentNode) { _moObserver.disconnect(); runCleanups(); }
        });
        _moObserver.observe(mo.parentNode || document.body, { childList: true });

        mo.querySelector('#inu-pid-cancel').addEventListener('click', () => { stopLiveMetrics(); mo.remove(); });

        // Sync pconv hidden input with the override selector
        const _pconvHidden = mo.querySelector('#inu-pid-pconv');
        const _pconvSel    = mo.querySelector('#inu-pid-pconv-sel');
        const _pconvNotice = mo.querySelector('#inu-pid-pconv-notice');
        function setPconv(val, label) {
            _pconvHidden.value = val;
            if (_pconvSel) _pconvSel.value = val;
            GM_setValue('pid_pconv', val);
            if (label && _pconvNotice) {
                _pconvNotice.textContent = label;
                _pconvNotice.style.display = '';
            }
        }
        if (_pconvSel) _pconvSel.addEventListener('change', () => setPconv(_pconvSel.value, null));

        // Sensor roles (discovered from sensor prefix)
        const SENSOR_ROLES = [
            { key: 'pv', label: 'PV',     placeholder: 'Mätvärde (_PV)',       req: true  },
            { key: 'sp', label: 'SP/CSP', placeholder: 'Börvärde (_CSP/_SP)',  req: false },
            { key: 'p',  label: 'P',      placeholder: 'P-parameter',           req: false },
            { key: 'i',  label: 'I',      placeholder: 'I-tid',                 req: false },
            { key: 'd',  label: 'D',      placeholder: 'D-parameter',           req: false },
            { key: 'db', label: 'Dödzon', placeholder: 'Dödzon % (_DB/_DZ)',    req: false },
        ];
        // Device roles (discovered from device prefix)
        const DEVICE_ROLES = [
            { key: 'op', label: 'OP', placeholder: 'Utsignal % (_OP/_OPM)', req: false },
        ];

        // Read current tag value from a role input
        function getTag(role) {
            const inp = mo.querySelector(`#inu-pid-taglist input[data-role="${role}"]`);
            return inp ? (inp.value.trim() || null) : null;
        }

        // Live metrics polling
        let liveMetricsTimer = null;
        function stopLiveMetrics() {
            if (liveMetricsTimer) { clearInterval(liveMetricsTimer); liveMetricsTimer = null; }
        }
        async function updateLiveMetrics() {
            const pvTag = getTag('pv'), spTag = getTag('sp'), opTag = getTag('op');
            const tagList = [pvTag, spTag, opTag].filter(Boolean);
            const el = mo.querySelector('#inu-pid-live-metrics');
            if (!tagList.length || !el) return;
            try {
                const params = tagList.map(t => 'tag=' + encodeTag(t)).join('&');
                const r = await fetch(CFG.endpoints.tagRead + '?usepid=1&formated=1&old=0&prio=1&json=1&' + params, { credentials: 'include' });
                if (!r.ok) return;
                const data = await r.json();
                const rv = t => t ? String(data[encodeTag(t)] ?? '—').replace(/[()]/g, '') : null;
                const parts = [];
                if (pvTag) parts.push(`<b>PV</b> ${rv(pvTag)}`);
                if (spTag) parts.push(`<b>SP</b> ${rv(spTag)}`);
                if (opTag) parts.push(`<b>OP</b> ${rv(opTag)}`);
                el.innerHTML = parts.map(p => `<span>${p}</span>`).join('');
            } catch { /* non-critical */ }
        }
        function startLiveMetrics() {
            stopLiveMetrics();
            updateLiveMetrics();
            liveMetricsTimer = setInterval(updateLiveMetrics, 2000);
        }
        onModalDismiss(stopLiveMetrics);

        function buildTagGrid(discovered) {
            const tl = mo.querySelector('#inu-pid-taglist');
            function roleRow(r) {
                const val = discovered[r.key] || '';
                return `<div class="inu-pid-tag-row">
  <span class="inu-pid-tag-lbl">${r.label}</span>
  <span class="inu-pid-tag-inp"><input type="text" list="inu-pid-dl" data-role="${r.key}" value="${escHtml(val)}" placeholder="${r.placeholder}"${r.req ? ' class="is-req"' : ''}></span>
  <span class="inu-pid-tag-stat" id="inu-pid-stat-${r.key}">${val ? '<span class="inu-pid-tag-ok">✓</span>' : '<span class="inu-pid-tag-miss">○</span>'}</span>
</div>`;
            }
            // Primary signal roles shown by default; parameter tags (P/I/D/DB) collapsed
            const primarySensorRoles = SENSOR_ROLES.filter(r => ['pv', 'sp'].includes(r.key));
            const paramRoles         = SENSOR_ROLES.filter(r => !['pv', 'sp'].includes(r.key));
            const hasParams = paramRoles.some(r => discovered[r.key]);
            tl.innerHTML = `
<div class="inu-pid-group-hdr">Sensor</div>
<div class="inu-pid-tag-grid">${primarySensorRoles.map(roleRow).join('')}</div>
<div class="inu-pid-group-hdr" style="margin-top:8px">Enhet (aktuator)</div>
<div class="inu-pid-tag-grid">${DEVICE_ROLES.map(roleRow).join('')}</div>
<details class="inu-pid-params-details"${hasParams ? ' open' : ''}>
  <summary><i class="fa fa-sliders"></i> Parametertagger (P, I, D, Dödzon)</summary>
  <div class="inu-pid-tag-grid" style="margin-top:4px">${paramRoles.map(roleRow).join('')}</div>
</details>
<div class="inu-pid-live-strip" id="inu-pid-live-metrics"><span style="color:#aaa"><i class="fa fa-circle-o-notch fa-spin"></i> Väntar på livevärden...</span></div>
<datalist id="inu-pid-dl"></datalist>`;

            // Update status dot + fetch button on any input change
            tl.querySelectorAll('input[data-role]').forEach(inp => {
                const update = () => {
                    const stat = tl.querySelector(`#inu-pid-stat-${inp.dataset.role}`);
                    const has = !!inp.value.trim();
                    inp.classList.toggle('has-val', has);
                    if (stat) stat.innerHTML = has ? '<span class="inu-pid-tag-ok">✓</span>' : '<span class="inu-pid-tag-miss">○</span>';
                    const hasPv = !!getTag('pv');
                    mo.querySelector('#inu-pid-fetch').disabled = !hasPv;
                };
                if (inp.value) inp.classList.add('has-val');
                inp.addEventListener('input', update);
            });

            mo.querySelector('#inu-pid-fetch').disabled = !discovered.pv;
            startLiveMetrics();

            // Populate datalist async
            getAllTagNames().then(names => {
                const dl = tl.querySelector('#inu-pid-dl');
                if (!dl || !names.length) return;
                const frag = document.createDocumentFragment();
                names.forEach(n => { const o = document.createElement('option'); o.value = n; frag.appendChild(o); });
                dl.appendChild(frag);
            });
        }

        async function doSearch() {
            const sensorPfx = mo.querySelector('#inu-pid-sensor-prefix').value.trim();
            const devicePfxEl = mo.querySelector('#inu-pid-device-prefix');
            const devicePfx = devicePfxEl.value.trim();
            GM_setValue('pid_sensor_prefix', sensorPfx);
            GM_setValue('pid_device_prefix', devicePfx);
            const discovered = sensorPfx
                ? await discoverPidTags(sensorPfx, devicePfx || null)
                : { pv:null, sp:null, op:null, p:null, i:null, d:null, buckets:{} };
            buildTagGrid(discovered);

            // Show device prefix row only when OP was not found under sensor prefix
            const devRow = mo.querySelector('#inu-pid-device-row');
            if (devRow) devRow.style.display = (discovered.op || !sensorPfx) ? 'none' : '';

            // Auto-detect circuit type from sensor prefix
            const detected = detectCircuitType(sensorPfx);
            if (detected) {
                mo.querySelector('#inu-pid-circuit').value = detected;
                // Auto-select sensible default window for the detected circuit
                const winSel = mo.querySelector('#inu-pid-window');
                const ct = PID_CIRCUIT_TYPES[detected];
                if (ct && winSel) winSel.value = String(ct.defaultWindow);
                updateWindowWarning();
            }

            // Auto-detect P convention from live P tag value
            // P% (band) controllers typically use values > 10 (e.g. 200% band)
            // Kp (gain) controllers use values < 10 (e.g. 0.5–5 gain)
            if (discovered.p) {
                try {
                    const r = await fetch(CFG.endpoints.tagRead + '?usepid=1&formated=0&old=1&prio=1&json=1&tag=' + encodeTag(discovered.p), { credentials: 'include' });
                    if (r.ok) {
                        const data = await r.json();
                        const raw = parseFloat(String(data[encodeTag(discovered.p)] ?? '').replace(/[()]/g, ''));
                        if (!isNaN(raw) && raw > 0) {
                            // raw ≥ 30: almost certainly P-band (Kp equivalent ≤ 3.3, unrealistic as a gain)
                            // raw ≤ 3:  almost certainly gain (P-band 3% = Kp 33, absurdly aggressive)
                            // 3–30:    ambiguous — e.g. Kp 12 for pressure vs P% 12 for temperature.
                            //          Don't auto-detect in this range to avoid false identifications.
                            const conv = raw >= 30 ? 'pband' : raw <= 3 ? 'gain' : null;
                            if (conv) {
                                setPconv(conv, `P-konvention detekterad automatiskt: ${conv === 'pband' ? 'P% (band)' : 'Kp (förstärkning)'} — P-taggen = ${raw}`);
                            } else {
                                const noticeEl = mo.querySelector('#inu-pid-pconv-notice');
                                if (noticeEl) { noticeEl.textContent = `P-taggen = ${raw} — oklart om P% eller Kp (gråzon 3–30). Välj manuellt i P-konvention-menyn ovan.`; noticeEl.style.display = ''; }
                            }
                        }
                    }
                } catch { /* non-critical */ }
            }
        }

        // Update short-window warning whenever circuit or window selection changes
        function updateWindowWarning() {
            const circuit = mo.querySelector('#inu-pid-circuit').value;
            const winMs   = parseInt(mo.querySelector('#inu-pid-window').value, 10);
            const ctCheck = PID_CIRCUIT_TYPES[circuit];
            const winWarnEl = mo.querySelector('#inu-pid-win-warn');
            if (!winWarnEl || !ctCheck) return;
            const minMs = ctCheck.responseTime[1] * 3 * 1000;
            if (winMs < minMs) {
                const minStr = minMs >= 3600000 ? `${(minMs/3600000).toFixed(1)} h` : `${Math.round(minMs/60000)} min`;
                winWarnEl.textContent = `⚠ Fönstret kan vara för kort för ${ctCheck.label} (typisk insvängningstid ${ctCheck.responseTime[0]}–${ctCheck.responseTime[1]} s) — rekommenderat minimum: ${minStr}.`;
                winWarnEl.style.display = '';
            } else {
                winWarnEl.style.display = 'none';
            }
        }
        mo.querySelector('#inu-pid-window').addEventListener('change', updateWindowWarning);
        mo.querySelector('#inu-pid-circuit').addEventListener('change', updateWindowWarning);
        updateWindowWarning(); // run on open

        mo.querySelector('#inu-pid-search').addEventListener('click', doSearch);
        mo.querySelector('#inu-pid-device-search')?.addEventListener('click', doSearch);
        mo.querySelector('#inu-pid-sensor-prefix').addEventListener('keydown', e => { if (e.key === 'Enter') doSearch(); });
        mo.querySelector('#inu-pid-device-prefix').addEventListener('keydown', e => { if (e.key === 'Enter') doSearch(); });
        // Always show the grid; auto-search if prefix remembered
        buildTagGrid({ pv:null, sp:null, op:null, p:null, i:null, d:null, buckets:{} });
        if (lastSensorPfx) setTimeout(doSearch, 50);

        mo.querySelector('#inu-pid-fetch').addEventListener('click', async () => {
            const pvTag = getTag('pv');
            if (!pvTag) return;
            const spTag = getTag('sp'), opTag = getTag('op');
            const pTag  = getTag('p'),  iTag  = getTag('i'), dTag = getTag('d'), dbTag = getTag('db');

            const circuit = mo.querySelector('#inu-pid-circuit').value;
            const winMs   = parseInt(mo.querySelector('#inu-pid-window').value, 10);
            // Short window warning for slow circuit types
            const ctCheck = PID_CIRCUIT_TYPES[circuit];
            const minRecommendedMs = ctCheck ? ctCheck.responseTime[1] * 3 * 1000 : 0;
            const winWarnEl = mo.querySelector('#inu-pid-win-warn');
            if (winWarnEl) {
                if (winMs < minRecommendedMs) {
                    const minStr = minRecommendedMs >= 3600000 ? `${(minRecommendedMs/3600000).toFixed(1)} h`
                                 : `${Math.round(minRecommendedMs/60000)} min`;
                    winWarnEl.textContent = `⚠ Fönstret kan vara för kort för ${ctCheck.label} (typisk insvängningstid ${ctCheck.responseTime[0]}–${ctCheck.responseTime[1]} s) — rekommenderat minimum: ${minStr}.`;
                    winWarnEl.style.display = '';
                } else {
                    winWarnEl.style.display = 'none';
                }
            }
            const pconv   = mo.querySelector('#inu-pid-pconv').value;
            GM_setValue('pid_circuit', circuit);
            GM_setValue('pid_window', String(winMs));
            GM_setValue('pid_pconv', pconv);

            stopLiveMetrics();
            const fetchBtn = mo.querySelector('#inu-pid-fetch');
            fetchBtn.disabled = true; fetchBtn.innerHTML = '<i class="fa fa-spinner fa-spin"></i> Hämtar...';

            const now = Date.now(), fromMs = now - winMs;

            // Snapshot of current tag selections (passed to results/fallback)
            const tagSel = { pv: pvTag, sp: spTag, op: opTag, p: pTag, i: iTag, d: dTag, db: dbTag };

            // Try trend API for PV, SP, OP
            const [pvPts, spPtsRaw, opPtsRaw] = await Promise.all([
                fetchTrend(pvTag, fromMs, now),
                spTag ? fetchTrend(spTag, fromMs, now) : Promise.resolve([]),
                opTag ? fetchTrend(opTag, fromMs, now) : Promise.resolve([]),
            ]);
            const spPts = spPtsRaw || [], opPts = opPtsRaw || [];

            fetchBtn.disabled = false; fetchBtn.innerHTML = '<i class="fa fa-download"></i> Hämta data';

            if (!pvPts || pvPts.length < 3) {
                startLiveFallback(mo, tagSel, circuit, winMs, pconv, onModalDismiss,
                    !pvPts ? { reason: 'api_unavailable' }
                           : { reason: 'too_few_pts', trendPts: pvPts.length });
                return;
            }

            // Check if trend data is too coarse for reliable FOPDT on this circuit type.
            // For fast circuits (responseTime[0] ≤ 60 s): auto-redirect to live when interval
            // exceeds the minimum response time — trend data can't capture the transient at all.
            // For slow circuits: proceed to results (advice warning will explain the limitation).
            {
                const ctCheck2 = PID_CIRCUIT_TYPES[circuit];
                if (ctCheck2 && ctCheck2.responseTime[0] <= 60) {
                    const ivals = pvPts.slice(1).map((p, i) => p.ts - pvPts[i].ts);
                    const sorted = ivals.slice().sort((a, b) => a - b);
                    const medS = sorted[Math.floor(sorted.length / 2)] / 1000;
                    if (medS > ctCheck2.responseTime[0]) {
                        startLiveFallback(mo, tagSel, circuit, winMs, pconv, onModalDismiss,
                            { reason: 'data_coarse', intervalS: Math.round(medS) });
                        return;
                    }
                }
            }

            // Read live PID parameter values + deadband using same format as live monitor
            const pidValues = { p: null, i: null, d: null };
            let deadbandPct = null; // deadband as % (0–100), as read from controller tag
            const pidTagList = [pTag, iTag, dTag, dbTag].filter(Boolean);
            if (pidTagList.length) {
                try {
                    const params = pidTagList.map(t => 'tag=' + encodeTag(t)).join('&');
                    const r = await fetch(CFG.endpoints.tagRead + '?usepid=1&formated=0&old=1&prio=1&json=1&' + params, { credentials: 'include' });
                    if (r.ok) {
                        const data = await r.json();
                        const rv = t => { const v = parseFloat(String(data[encodeTag(t)] ?? '').replace(/[()]/g, '')); return isNaN(v) ? null : v; };
                        if (pTag)  pidValues.p = rv(pTag);
                        if (iTag)  pidValues.i = rv(iTag);
                        if (dTag)  pidValues.d = rv(dTag);
                        if (dbTag) deadbandPct = rv(dbTag);
                    }
                } catch { /* non-critical */ }
            }

            // Convert deadband % to engineering units using PV tag scaling
            let deadbandEu = null;
            if (deadbandPct !== null) {
                const meta = _allTagMetaCache[pvTag];
                if (meta && meta.engmax !== meta.engmin) {
                    deadbandEu = Math.abs(meta.engmax - meta.engmin) * deadbandPct / 100;
                }
            }

            // Normalize OP to 0–100% if tag has non-standard engineering range
            // (e.g. engmin=0, engmax=10 means 0–10V → normalize to 0–100%)
            let opPtsNorm = opPts;
            if (opTag) {
                const opMeta = _allTagMetaCache[opTag];
                if (opMeta && opMeta.engmax > opMeta.engmin &&
                    (Math.abs(opMeta.engmin) > 1 || Math.abs(opMeta.engmax - 100) > 1)) {
                    const span = opMeta.engmax - opMeta.engmin;
                    opPtsNorm = opPts.map(p => ({ ts: p.ts, val: (p.val - opMeta.engmin) / span * 100 }));
                }
            }

            renderResults(mo, tagSel, pvPts, spPts, opPtsNorm, circuit, pidValues, winMs, pconv, deadbandEu, 'trend');
        });
    }

    // opts: { reason: 'api_unavailable'|'data_coarse'|'too_few_pts', intervalS, trendPts }
    function startLiveFallback(mo, discovered, circuit, winMs, pconv, onModalDismiss, opts) {
        const resultsEl = mo.querySelector('#inu-pid-results');
        const setupEl   = mo.querySelector('#inu-pid-setup');
        setupEl.style.display = 'none';
        resultsEl.style.display = '';

        const reason = opts?.reason || 'api_unavailable';
        const ct = PID_CIRCUIT_TYPES[circuit];

        // Build context-specific reason message
        let liveReason;
        if (reason === 'data_coarse') {
            const intS = opts.intervalS;
            liveReason = `Trenddata för grov för processidentifiering — loggningsintervall ~${intS} s överstiger svarstiden för ${ct?.label || 'kretsen'} (${ct?.responseTime[0]}–${ct?.responseTime[1]} s). Live-insamling ger bättre modellidentifiering.`;
        } else if (reason === 'too_few_pts') {
            liveReason = `Trenddata innehåller för få punkter (${opts.trendPts ?? '?'}) för tillförlitlig analys. Live-insamling samlar ny data direkt.`;
        } else {
            liveReason = `Trenddataendpunkten svarade inte eller returnerade tomma data.`;
        }

        const bufferMax  = 1800; // max 60 min at 2s
        const needed     = 120;  // 2 min — full confidence
        const earlyAllow = 30;   // 30s — allow early analysis with a warning
        const pvBuf = [], spBuf = [], opBuf = [];
        let elapsed = 0;
        let startTs = null;
        let intervalId = null;
        let stopped = false;
        let _pollInflight = false; // prevent concurrent fetches if one hangs

        resultsEl.innerHTML = `
<div class="inu-pid-live-bar" id="inu-pid-live-info">
  <b><i class="fa fa-circle" style="color:#f57c00"></i> Live-insamling</b> — ${liveReason}<br>
  <span id="inu-pid-live-status">Samlar data... 0 s / ~${needed} s minimum</span>
  <div class="inu-pid-progress"><div class="inu-pid-progress-fill" id="inu-pid-progress-fill" style="width:0%"></div></div>
</div>
<div class="inu-pid-live-body">
  <div id="inu-pid-live-chart" class="inu-pid-chart-wrap"></div>
  <div id="inu-pid-live-stats" class="inu-pid-live-stats">
    <div class="inu-pid-stat-group">
      <div class="inu-pid-stat-hdr">PV</div>
      <div class="inu-pid-stat-row"><span>Nu</span><span id="lv-pv-now">—</span></div>
      <div class="inu-pid-stat-row inu-pid-stat-hi"><span>Svängning</span><span id="lv-pv-range">—</span></div>
      <div class="inu-pid-stat-row"><span>Min</span><span id="lv-pv-min">—</span></div>
      <div class="inu-pid-stat-row"><span>Max</span><span id="lv-pv-max">—</span></div>
      <div class="inu-pid-stat-row"><span>Medel</span><span id="lv-pv-mean">—</span></div>
      <div class="inu-pid-stat-row"><span>StdAv</span><span id="lv-pv-std">—</span></div>
    </div>
    <div class="inu-pid-stat-group" id="lv-sp-group" style="display:none">
      <div class="inu-pid-stat-hdr">SP</div>
      <div class="inu-pid-stat-row"><span>Nu</span><span id="lv-sp-now">—</span></div>
      <div class="inu-pid-stat-row inu-pid-stat-hi"><span>Fel (e)</span><span id="lv-sp-err">—</span></div>
    </div>
    <div class="inu-pid-stat-group" id="lv-op-group" style="display:none">
      <div class="inu-pid-stat-hdr">OP</div>
      <div class="inu-pid-stat-row"><span>Nu</span><span id="lv-op-now">—</span></div>
      <div class="inu-pid-stat-row"><span>Min</span><span id="lv-op-min">—</span></div>
      <div class="inu-pid-stat-row"><span>Max</span><span id="lv-op-max">—</span></div>
    </div>
  </div>
</div>
<div class="inu-pid-btn-row">
  <button class="bx" id="inu-pid-live-stop"><i class="fa fa-stop"></i> Stoppa</button>
  <button class="bok" id="inu-pid-live-now" disabled><i class="fa fa-bar-chart"></i> Analysera nu</button>
</div>`;

        function stop() {
            stopped = true;
            if (intervalId) { clearInterval(intervalId); intervalId = null; }
        }
        if (onModalDismiss) onModalDismiss(stop);

        mo.querySelector('#inu-pid-live-stop').addEventListener('click', () => { stop(); mo.remove(); });
        mo.querySelector('#inu-pid-live-now').addEventListener('click', async () => {
            stop();
            if (pvBuf.length < 2) { toastErr('Inte tillräckligt med data — kontrollera att taggnamnet stämmer.'); return; }
            // Fetch P/I/D current values before rendering (same as trend path)
            const pidValues = { p: null, i: null, d: null };
            const pidTagList = [discovered.p, discovered.i, discovered.d].filter(Boolean);
            if (pidTagList.length) {
                try {
                    const params = pidTagList.map(t => 'tag=' + encodeTag(t)).join('&');
                    const r = await fetch(CFG.endpoints.tagRead + '?usepid=1&formated=0&old=1&prio=1&json=1&' + params, { credentials: 'include' });
                    if (r.ok) {
                        const data = await r.json();
                        const rv = t => { const v = parseFloat(String(data[encodeTag(t)] ?? '').replace(/[()]/g, '')); return isNaN(v) ? null : v; };
                        if (discovered.p) pidValues.p = rv(discovered.p);
                        if (discovered.i) pidValues.i = rv(discovered.i);
                        if (discovered.d) pidValues.d = rv(discovered.d);
                    }
                } catch { /* non-critical */ }
            }
            renderResults(mo, discovered, [...pvBuf], [...spBuf], [...opBuf], circuit, pidValues, winMs, pconv);
        });

        function updateLiveStats() {
            const s = n => n != null ? n.toFixed(2) : '—';
            function stats(buf) {
                if (!buf.length) return null;
                const vals = buf.map(p => p.val);
                const min = vals.reduce((a, b) => Math.min(a, b), Infinity);
                const max = vals.reduce((a, b) => Math.max(a, b), -Infinity);
                const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
                const variance = vals.reduce((a, b) => a + (b - mean) ** 2, 0) / vals.length;
                return { now: vals[vals.length - 1], min, max, range: max - min, mean, std: Math.sqrt(variance) };
            }
            const pv = stats(pvBuf), sp = stats(spBuf), op = stats(opBuf);
            if (!pv) return;
            const set = (id, v) => { const el = mo.querySelector('#' + id); if (el) el.textContent = v; };
            set('lv-pv-now',   s(pv.now));
            set('lv-pv-range', s(pv.range));
            set('lv-pv-min',   s(pv.min));
            set('lv-pv-max',   s(pv.max));
            set('lv-pv-mean',  s(pv.mean));
            set('lv-pv-std',   s(pv.std));
            // Colour-code fluctuation: red if range > 1× stdDev threshold for circuit
            const rangeEl = mo.querySelector('#lv-pv-range');
            if (rangeEl) {
                const ct = PID_CIRCUIT_TYPES[circuit];
                const threshold = ct ? (ct.tiRange[0] > 200 ? 2 : 1) : 1; // looser for slow loops
                rangeEl.style.color = pv.range > threshold ? '#e53935' : '#43a047';
            }
            if (sp) {
                const spGrp = mo.querySelector('#lv-sp-group');
                if (spGrp) spGrp.style.display = '';
                set('lv-sp-now', s(sp.now));
                const err = pv.now - sp.now;
                const errEl = mo.querySelector('#lv-sp-err');
                if (errEl) {
                    errEl.textContent = (err >= 0 ? '+' : '') + err.toFixed(2);
                    errEl.style.color = Math.abs(err) > pv.std * 2 ? '#e53935' : '#555';
                }
            }
            if (op) {
                const opGrp = mo.querySelector('#lv-op-group');
                if (opGrp) opGrp.style.display = '';
                set('lv-op-now', s(op.now) + ' %');
                set('lv-op-min', s(op.min) + ' %');
                set('lv-op-max', s(op.max) + ' %');
            }
        }

        async function poll() {
            if (stopped || _pollInflight) return;
            _pollInflight = true;
            try {
                // WebPort /tag/read: request with tag=<-5F-encoded-id>, response is a flat object {encodedId: "value"}
                const tagList = [discovered.pv, discovered.sp, discovered.op].filter(Boolean);
                const params = tagList.map(t => 'tag=' + encodeTag(t)).join('&');
                const r = await fetch(CFG.endpoints.tagRead + '?usepid=1&formated=0&old=1&prio=1&json=1&' + params, { credentials: 'include' });
                if (!r.ok) return;
                const data = await r.json();
                function readTag(name) {
                    if (!name) return NaN;
                    // Response keyed by encoded name; strip formatting parens (old values shown as "(x.x)")
                    const raw = data[encodeTag(name)];
                    return parseFloat(String(raw ?? '').replace(/[()]/g, ''));
                }
                const ts = Date.now();
                const pvVal = readTag(discovered.pv);
                const spVal = readTag(discovered.sp);
                const opVal = readTag(discovered.op);
                if (!isNaN(pvVal)) pvBuf.push({ ts, val: pvVal });
                if (!isNaN(spVal)) spBuf.push({ ts, val: spVal });
                if (!isNaN(opVal)) opBuf.push({ ts, val: opVal });
                // Trim to bufferMax
                if (pvBuf.length > bufferMax) pvBuf.shift();
                if (spBuf.length > bufferMax) spBuf.shift();
                if (opBuf.length > bufferMax) opBuf.shift();
            } catch { /* ignore poll error */ } finally { _pollInflight = false; }

            elapsed = Math.floor((Date.now() - startTs) / 1000);
            const pct = Math.min(100, (elapsed / needed) * 100);
            const fillEl = mo.querySelector('#inu-pid-progress-fill');
            const statusEl = mo.querySelector('#inu-pid-live-status');
            const nowBtn = mo.querySelector('#inu-pid-live-now');
            if (fillEl) fillEl.style.width = pct + '%';
            // Detect SP step
            const hasStep = spBuf.length >= 4 && Math.abs(spBuf[spBuf.length-1].val - spBuf[0].val) > 0.5;
            if (statusEl) {
                statusEl.textContent = pvBuf.length === 0 && elapsed >= 10
                    ? `Inga data mottagna (${elapsed} s) — kontrollera taggnamnet eller att WebPort svarar.`
                    : hasStep
                    ? `Stegrespons detekterat — ${pvBuf.length} sampel (${elapsed} s). Klicka "Analysera nu" för resultat.`
                    : elapsed >= needed
                    ? `Samlar data... ${elapsed} s (${pvBuf.length} sampel) — klicka "Analysera nu" när du är klar.`
                    : `Samlar data... ${elapsed} s / ~${needed} s (${pvBuf.length} sampel)`;
            }
            if (nowBtn) {
                nowBtn.disabled = !(elapsed >= earlyAllow || hasStep);
                nowBtn.title = elapsed < needed && !hasStep
                    ? `Tillräckligt data för tidig analys — bättre resultat efter ${needed} s`
                    : '';
            }

            // Update chart and stats every poll
            if (pvBuf.length >= 2) {
                const chartEl = mo.querySelector('#inu-pid-live-chart');
                if (chartEl) renderPidChart(chartEl, [...pvBuf], [...spBuf], [...opBuf]);
                updateLiveStats();
            }
        }

        startTs = Date.now();
        intervalId = setInterval(poll, 2000);
        poll(); // immediate first poll
    }

    function renderResults(mo, discovered, pvPts, spPts, opPts, circuit, pidValues, winMs, pconv, deadbandEu, source) {
        pconv = pconv || GM_getValue('pid_pconv', 'pband');
        const setupEl   = mo.querySelector('#inu-pid-setup');
        const resultsEl = mo.querySelector('#inu-pid-results');
        setupEl.style.display = 'none';
        resultsEl.style.display = '';

        const isPband = pconv === 'pband';
        const metrics = computePidMetrics(pvPts, spPts, opPts, deadbandEu);
        const tuning  = computeTuning(metrics, pidValues, pconv);
        // Normalise P to Kp for advice comparisons (ZN/CC formulas always use gain)
        const pidForAdvice = pidValues ? {
            p: (pidValues.p != null && pidValues.p !== 0 && isPband) ? 100 / pidValues.p : pidValues.p,
            i: pidValues.i,
            d: pidValues.d,
        } : null;
        const advice  = getPidAdvice(metrics, circuit, pidForAdvice, pconv, tuning, deadbandEu);
        const ct      = PID_CIRCUIT_TYPES[circuit];
        const prefix  = discovered.pv ? discovered.pv.replace(/_PV$/i, '') : '';

        const winLabel = winMs >= 86400000 ? `${winMs/86400000}d`
                       : winMs >= 3600000  ? `${winMs/3600000}h`
                       : `${winMs/60000}m`;

        function metricRow(label, val, unit, tip) {
            const labelHtml = tip ? `<span data-tip="${tip}">${label}</span>` : `<span>${label}</span>`;
            if (val === null || val === undefined) return `<div class="inu-pid-metric-row">${labelHtml}<span class="inu-pid-metric-val">—</span></div>`;
            return `<div class="inu-pid-metric-row">${labelHtml}<span class="inu-pid-metric-val">${typeof val === 'number' ? val.toFixed(1) : val}${unit || ''}</span></div>`;
        }
        // P-convention helpers: pconv='pband' means controller uses P% (band), 'gain' means Kp
        // Convert Kp (gain) → display value in chosen convention
        function toDisplayP(kc) {
            if (kc == null || kc === 0) return null; // guard division by zero
            return isPband ? +(100 / kc).toFixed(2) : +kc;
        }
        // Convert stored P-tag value → Kp gain (for advice comparisons)
        function toKp(pVal) {
            if (pVal == null || pVal === 0) return null; // guard division by zero
            return isPband ? 100 / pVal : pVal;
        }
        const pLabel = isPband ? 'P%' : 'Kp';

        function pidRow(label, val) {
            if (val == null) return `<div class="inu-pid-metric-row"><span>${label}</span><span class="inu-pid-metric-val">—</span></div>`;
            return `<div class="inu-pid-metric-row"><span>${label}</span><span class="inu-pid-metric-val">${val.toFixed(2)}</span></div>`;
        }

        // Show current P value with both conventions if pband
        function pValueDisplay(rawP) {
            if (rawP == null) return pidRow(isPband ? 'P%' : 'Kp', null);
            if (isPband) {
                const kpEquiv = toKp(rawP);
                return `<div class="inu-pid-metric-row"><span>P%</span><span class="inu-pid-metric-val">${rawP.toFixed(2)} <span style="font-weight:normal;color:#888">(Kp≈${kpEquiv?.toFixed(3) ?? '—'})</span></span></div>`;
            }
            return pidRow('Kp', rawP);
        }

        // Hard limits per circuit type — clamp suggested Kp and Ti to 3× typical range
        // Returns { kc, ti, td, _clampedKc, _clampedTi } with boolean flags when clamped
        function clampTuning(t) {
            if (!ct) return t;
            const out = Object.assign({}, t);
            if (t.kc != null) {
                const maxKc = ct.kpRange[1] * 3;
                const minKc = ct.kpRange[0] / 3;
                if (t.kc > maxKc) { out.kc = maxKc; out._clampedKc = true; }
                else if (t.kc < minKc && t.kc > 0) { out.kc = minKc; out._clampedKc = true; }
            }
            if (t.ti != null) {
                const maxTi = ct.tiRange[1] * 3;
                const minTi = ct.tiRange[0] / 3;
                if (t.ti > maxTi) { out.ti = maxTi; out._clampedTi = true; }
                else if (t.ti < minTi && t.ti > 0) { out.ti = minTi; out._clampedTi = true; }
            }
            return out;
        }

        const adviceHtml = advice.map(a => {
            const icon = a.severity === 'ok' ? 'check-circle' : a.severity === 'warn' ? 'exclamation-triangle' : 'info-circle';
            return `<div class="inu-pid-adv-item ${a.severity}"><i class="fa fa-${icon}"></i><span>${a.text}</span></div>`;
        }).join('');

        // Tuning table — convert Kp output to chosen convention, show delta from current
        const currentKp = pidForAdvice?.p; // already normalised to Kp in caller
        const currentTi = pidValues?.i;

        function deltaArrow(suggested, current) {
            if (current == null || current === 0 || suggested == null) return '';
            const ratio = suggested / current;
            const pct   = Math.round((ratio - 1) * 100);
            const absRatio = Math.abs(ratio);
            const col = absRatio < 0.7 || absRatio > 1.43  // >30% change
                      ? (absRatio < 0.5 || absRatio > 2 ? '#e53935' : '#f57c00')  // >2× = red
                      : '#43a047';
            const arrow = ratio > 1 ? '▲' : '▼';
            const label = `${arrow} ${Math.abs(pct)}%`;
            return ` <span style="color:${col};font-size:10px;font-weight:600">${label}</span>`;
        }

        function sanityCheck(kcRaw) {
            if (currentKp == null || currentKp === 0 || kcRaw == null) return '';
            const ratio = kcRaw / currentKp;
            if (ratio > 3 || ratio < 0.2) {
                return `<div style="color:#e53935;font-size:10px;margin-top:2px">⚠ ${(ratio).toFixed(1)}× nuvarande — tillämpa stegvis</div>`;
            }
            return '';
        }

        function tuningRow(method, tRaw) {
            const t = clampTuning(tRaw);
            const clampWarn = (t._clampedKc || t._clampedTi)
                ? `<div style="color:#e67e22;font-size:10px;margin-top:1px">⚠ Begränsat till typintervall för ${ct?.label || 'kretsen'}</div>` : '';
            const kcRaw = t.kc !== undefined ? t.kc : null;
            const pDisp = kcRaw !== null ? (toDisplayP(kcRaw) ?? '—') : '—';
            const kcDisp = kcRaw !== null ? kcRaw : '—';
            const ti = t.ti !== null && t.ti !== undefined ? t.ti + ' s' : '—';
            const td = t.td !== null && t.td !== undefined ? t.td + ' s' : '—';
            // Sanity warning if suggestion is >3× or <0.2× current
            const sanity = sanityCheck(kcRaw);
            // Before → After format: show current value → suggested value when current is known
            function beforeAfterP() {
                if (currentKp == null || kcRaw == null) {
                    return isPband
                        ? `${pDisp}% <span style="color:#aaa;font-size:10px">(Kp=${kcDisp})</span>`
                        : String(kcDisp);
                }
                const curDisp = isPband ? (toDisplayP(currentKp) ?? '—') + '%' : String(+currentKp.toFixed(2));
                const sugDisp = isPband ? pDisp + '%' : String(kcDisp);
                const arrow = deltaArrow(kcRaw, currentKp);
                return `<span style="color:#aaa">${curDisp}</span> → <b>${sugDisp}</b>${arrow}`;
            }
            function beforeAfterTi() {
                if (currentTi == null || t.ti == null) return ti;
                const arrow = deltaArrow(t.ti, currentTi);
                return `<span style="color:#aaa">${currentTi} s</span> → <b>${t.ti} s</b>${arrow}`;
            }
            return `<tr><td>${method}${sanity}${clampWarn}</td><td>${beforeAfterP()}</td><td>${beforeAfterTi()}</td><td>${td}</td></tr>`;
        }

        let tuningHtml = '';
        if (tuning.saturated) {
            tuningHtml += `
<div class="inu-pid-results-box" style="grid-column:1/-1;margin-top:0">
  <h5><i class="fa fa-sliders"></i> Inställningsförslag</h5>
  <div style="color:#e53935;font-size:11px;">
    <i class="fa fa-times-circle"></i>
    Aktuatorn är mättad (&gt;30% av fönstret vid 0% eller 100%) — FOPDT-modell och alla beräknade mätvärden är opålitliga under mättning.
    Inställningsförslag undertrycks. Åtgärda mättningsorsaken (börvärde, dimensionering, driftläge) innan analys.
  </div>
</div>`;
        } else if (tuning.poorFit) {
            tuningHtml += `
<div class="inu-pid-results-box" style="grid-column:1/-1;margin-top:0">
  <h5><i class="fa fa-sliders"></i> Inställningsförslag</h5>
  <div style="color:#e53935;font-size:11px;">
    <i class="fa fa-times-circle"></i>
    Processmodellen passar dåligt mot mätdata (R² = ${tuning.r2}) — troligen brusig signal, glapp i datan eller inget tydligt stegrespons.
    Inställningsförslag undertrycks för att undvika felaktiga värden. Prova ett tydligare OP-steg (10–20%) i handdrift med stabil signal.
  </div>
</div>`;
        } else if (tuning.noDeadTime && !tuning.openLoop) {
            // Only show "no tuning" when IMC also couldn't be computed (no model at all)
            tuningHtml += `
<div class="inu-pid-results-box" style="grid-column:1/-1;margin-top:0">
  <h5><i class="fa fa-sliders"></i> Inställningsförslag</h5>
  <div style="color:#e67e22;font-size:11px;">
    <i class="fa fa-exclamation-triangle"></i>
    Dödtiden är för kort (L &lt; 1 s) och ingen processmodell kunde beräknas — kan ej ge inställningsförslag.
    Prova stängd-krets ZN: öka Kp tills processen oscillerar, notera perioden och kör analysen igen.
  </div>
</div>`;
        }
        if (tuning.openLoop) {
            const ol = tuning.openLoop;
            const m  = ol.model;
            const pm = metrics?.processModel;
            const pvChange = (m.kp * 10).toFixed(2);
            const dir = m.kp >= 0 ? 'ökar' : 'minskar';
            const r2str = pm?.r2 != null
                ? ` &ensp; <span data-tip="Processmodellens (FOPDT) anpassning mot mätdata (0–1). Över 0.85 = tillförlitlig modell. 0.70–0.85 = acceptabel. Under 0.70 = dålig anpassning, inställningsförslag undertrycks." style="color:${pm.r2 >= 0.85 ? '#43a047' : pm.r2 >= 0.70 ? '#f57c00' : '#e53935'}">R² = ${pm.r2}</span>`
                : '';
            const noDeadTimeNote = tuning.noDeadTime
                ? `<div style="color:#e67e22;font-size:10px;margin-bottom:6px"><i class="fa fa-exclamation-triangle"></i> L &lt; 1 s — dödtiden är för kort för ZN/CC (de dividerar med L). Endast IMC visas.</div>`
                : '';
            // ZN/CC only available when L ≥ 1
            const znCcSection = (ol.zn && ol.cc) ? (() => {
                const ccRows = m.ratio >= 0.1 && m.ratio <= 1.0
                    ? tuningRow('CC &ndash; PI', ol.cc.pi) + tuningRow('CC &ndash; PID', ol.cc.pid)
                    : `<tr><td colspan="4" style="color:#e67e22;font-size:10px;padding:4px 6px"><i class="fa fa-exclamation-triangle"></i> CC ej tillämplig (L/T = ${m.ratio})</td></tr>`;
                return `
  <details style="margin-top:8px">
    <summary style="font-size:11px;cursor:pointer;list-style:none;color:#888"><i class="fa fa-table"></i> Visa fler metoder (ZN, CC)</summary>
    <table class="inu-pid-tune-table" style="margin-top:6px">
      <thead><tr><th>Metod</th><th data-tip="${isPband ? 'Proportionellt band (P%). Lägre värde = mer aggressiv förstärkning. P% = 100/Kp.' : 'Proportionell förstärkning. Högre värde = mer aggressiv. Kp = 100/P%.'}">${pLabel}</th><th data-tip="Integraltid i sekunder. Styr hur snabbt kvarstående fel elimineras. Kort Ti = aggressiv korrigering, risk för oscillationer. Typvärde för ${ct?.label}: ${ct?.tiRange[0]}–${ct?.tiRange[1]} s.">Ti</th><th data-tip="Derivatatid i sekunder. Förstärker förändringstakten i felet — dämpar överskjutning men förstärker också brus. Undvik på temperaturkretsar med brusiga givare.">Td</th></tr></thead>
      <tbody>
        ${tuningRow('ZN &ndash; PI',  ol.zn.pi)}
        ${tuningRow('ZN &ndash; PID', ol.zn.pid)}
        ${ccRows}
      </tbody>
    </table>
    <div style="font-size:10px;color:#e67e22;margin-top:4px;"><i class="fa fa-exclamation-triangle"></i> <b>D-del (Td):</b> förstärker brus — undvik på temperaturkretsar utan D-filter.</div>
    <div style="font-size:10px;color:#888;margin-top:2px;">ZN = Ziegler-Nichols (aggressiv) &ensp; CC = Cohen-Coon (giltig 0.1 &lt; L/T &lt; 1.0)</div>
  </details>`;
            })() : '';
            tuningHtml += `
<div class="inu-pid-results-box" style="grid-column:1/-1;margin-top:0">
  <h5><i class="fa fa-check-circle" style="color:#43a047"></i> Rekommenderat inställningsförslag — IMC/Lambda</h5>
  ${noDeadTimeNote}
  <table class="inu-pid-tune-table">
    <thead><tr><th>Metod</th><th data-tip="${isPband ? 'Proportionellt band (P%). Lägre värde = mer aggressiv förstärkning. P% = 100/Kp.' : 'Proportionell förstärkning. Högre värde = mer aggressiv. Kp = 100/P%.'}">${pLabel}</th><th data-tip="Integraltid i sekunder. Styr hur snabbt kvarstående fel elimineras. Kort Ti = aggressiv korrigering, risk för oscillationer. Typvärde för ${ct?.label}: ${ct?.tiRange[0]}–${ct?.tiRange[1]} s.">Ti</th><th data-tip="Derivatatid i sekunder. Förstärker förändringstakten i felet — dämpar överskjutning men förstärker också brus. Undvik på temperaturkretsar med brusiga givare.">Td</th></tr></thead>
    <tbody>${tuningRow('IMC &ndash; PI', ol.imc.pi)}</tbody>
  </table>
  <div style="font-size:10px;color:#555;margin-top:4px;font-style:italic">
    En 10%-ig ökning av OP ${dir} PV med ca ${Math.abs(pvChange)} enheter — reaktionen märks efter ~${m.L} s och stabiliseras efter ca ${m.T} s.${r2str}
  </div>
  <details style="margin-top:8px">
    <summary style="font-size:11px;font-weight:600;cursor:pointer;list-style:none" data-tip="Lambda (λ) är den önskade tidskonstanten för den slutna kretsen. Lägre λ = snabbare respons men högre risk för oscillationer. Högre λ = långsammare men robustare reglering. Standard: max(2L, T/3)."><i class="fa fa-sliders"></i> Avancerat: justera Lambda</summary>
    <div style="margin-top:6px">
      <div style="display:flex;align-items:center;gap:8px;font-size:11px">
        <span style="color:#999">Snabb</span>
        <input type="range" id="inu-pid-lambda-slider" min="${(ol.model.L).toFixed(2)}" max="${(ol.model.L * 10).toFixed(2)}"
               step="${(ol.model.L * 0.1).toFixed(2)}" value="${ol.imc.pi.lambda}" style="flex:1;accent-color:#2d5a9e">
        <span style="color:#999">Robust</span>
        <span style="min-width:40px;font-size:10px;color:#666">λ=<b id="inu-pid-lambda-val">${ol.imc.pi.lambda}</b>s</span>
      </div>
      <div style="font-size:11px;margin-top:4px;color:#555">
        IMC PI: ${isPband ? 'P%' : 'Kp'} = <b id="inu-pid-imc-kc">${isPband ? (toDisplayP(ol.imc.pi.kc) ?? '—') : ol.imc.pi.kc}</b>
        &ensp; Ti = <b id="inu-pid-imc-ti">${ol.imc.pi.ti}</b> s
      </div>
      <div style="font-size:10px;color:#888;margin-top:4px;">Processmodell: Kp = ${m.kp} &ensp; L = ${m.L} s &ensp; T = ${m.T} s &ensp; L/T = ${m.ratio}</div>
    </div>
  </details>
  ${znCcSection}
</div>`;
        }
        if (tuning.closedLoop && !tuning.closedLoop.disturbanceDriven) {
            const cl = tuning.closedLoop;
            const kuDisp = isPband ? `Kp=${cl.Ku} (P%≈${(100/cl.Ku).toFixed(1)})` : cl.Ku;
            tuningHtml += `
<div class="inu-pid-results-box" style="grid-column:1/-1;margin-top:0">
  <h5><i class="fa fa-sliders"></i> Inställningsförslag (sluten krets — ZN)</h5>
  <div style="font-size:10px;color:#888;margin-bottom:6px;">
    Ku ≈ ${kuDisp} (nuvarande P-värde) &ensp; Pu = ${cl.Pu} s (oscillationsperiod) — <i>konservativt</i>
  </div>
  <table class="inu-pid-tune-table">
    <thead><tr><th>Metod</th><th data-tip="${isPband ? 'Proportionellt band (P%). Lägre värde = mer aggressiv förstärkning. P% = 100/Kp.' : 'Proportionell förstärkning. Högre värde = mer aggressiv. Kp = 100/P%.'}">${pLabel}</th><th data-tip="Integraltid i sekunder. Styr hur snabbt kvarstående fel elimineras. Kort Ti = aggressiv korrigering, risk för oscillationer. Typvärde för ${ct?.label}: ${ct?.tiRange[0]}–${ct?.tiRange[1]} s.">Ti</th><th data-tip="Derivatatid i sekunder. Förstärker förändringstakten i felet — dämpar överskjutning men förstärker också brus. Undvik på temperaturkretsar med brusiga givare.">Td</th></tr></thead>
    <tbody>
      ${tuningRow('ZN &ndash; PID', cl.zn.pid)}
      ${tuningRow('ZN &ndash; PI',  cl.zn.pi)}
    </tbody>
  </table>
</div>`;
        } else if (tuning.closedLoop?.disturbanceDriven) {
            const cl = tuning.closedLoop;
            tuningHtml += `
<div class="inu-pid-results-box" style="grid-column:1/-1;margin-top:0">
  <h5><i class="fa fa-info-circle"></i> Sluten krets — ZN ej tillförlitlig</h5>
  <div style="font-size:11px;color:#888;">
    Oscillationsperioden är oregelbunden (CV = ${(cl.cv*100).toFixed(0)}%) — oscillationerna drivs troligen av processstörningar snarare än regulatorförstärkningen.
    ZN sluten krets förutsätter att nuvarande Kp är kritisk förstärkning Ku, vilket inte stämmer om I-verkan eller yttre störningar driver oscillationerna.
    Använd ett OP-steg i handdrift för att få en tillförlitlig FOPDT-modell.
  </div>
</div>`;
        }

        const medIntervalS = metrics?.medianIntervalMs ? metrics.medianIntervalMs / 1000 : null;
        const intervalLabel = medIntervalS != null ? ` · ~${medIntervalS < 60 ? Math.round(medIntervalS) + ' s' : Math.round(medIntervalS / 60) + ' min'} intervall` : '';
        const samplesLabel = `${pvPts.length} PV${spPts.length ? ` · ${spPts.length} SP` : ''}${opPts.length ? ` · ${opPts.length} OP` : ''} sampel${intervalLabel}`;

        // Confidence indicator
        const r2      = metrics?.processModel?.r2;
        const hasStep = metrics?.opStepDetected || metrics?.stepDetected;
        const nPts    = pvPts.length;
        const isTrend = source === 'trend';
        const samplingTooCoarse = medIntervalS != null && ct && medIntervalS > ct.responseTime[0];
        const samplingMarginal  = medIntervalS != null && ct && medIntervalS > ct.responseTime[0] / 3 && !samplingTooCoarse;
        let confLevel, confColor, confIcon;
        if (samplingTooCoarse) {
            confLevel = 'Låg'; confColor = '#e53935'; confIcon = 'times-circle';
        } else if (isTrend && hasStep && r2 != null && r2 >= 0.85 && nPts >= 100 && !metrics?.actuatorSaturated && !samplingMarginal) {
            confLevel = 'Hög'; confColor = '#43a047'; confIcon = 'check-circle';
        } else if (isTrend && nPts >= 50 && !metrics?.actuatorSaturated) {
            confLevel = 'Medel'; confColor = '#f57c00'; confIcon = 'exclamation-circle';
        } else {
            confLevel = 'Låg'; confColor = '#e53935'; confIcon = 'times-circle';
        }
        const intervalTip = medIntervalS != null ? ` Loggningsintervall: ~${medIntervalS < 60 ? Math.round(medIntervalS) + ' s' : Math.round(medIntervalS / 60) + ' min'}${samplingTooCoarse ? ' — för grovt för FOPDT.' : samplingMarginal ? ' — marginellt för FOPDT.' : '.'}` : '';
        const confTip = `Tillförlitlighet baseras på: datakälla (trend=bättre), antal sampel (${nPts}), om ett tydligt OP-steg detekterades, R²-anpassning av processmodellen, och loggningsintervallets lämplighet.${intervalTip} Hög = alla kriterier OK. Låg = live-data, grovt loggintervall eller kort mätfönster — tolka inställningsförslag med försiktighet.`;
        const confidencePill = `<span data-tip="${confTip}" style="display:inline-block;padding:1px 8px;border-radius:10px;background:${confColor}20;color:${confColor};font-size:10px;font-weight:600;border:1px solid ${confColor}40"><i class="fa fa-${confIcon}"></i> ${confLevel}${r2 != null ? ` · R²=${r2}` : ''}</span>`;

        // Traffic-light status
        const hasWarnAdvice = advice.some(a => a.severity === 'warn');
        const hasCritical   = metrics?.actuatorSaturated || metrics?.processModel?.reverseActing ||
            metrics?.amplitudeTrend === 'growing' || tuning?.saturated;
        let statusBg, statusBorder, statusIcon, statusText;
        if (hasCritical || (confLevel === 'Låg' && hasWarnAdvice)) {
            statusBg = '#ffebee'; statusBorder = '#e53935'; statusIcon = 'exclamation-triangle'; statusText = 'Åtgärd kan krävas — se rekommendationer nedan';
        } else if (hasWarnAdvice) {
            statusBg = '#fff3e0'; statusBorder = '#f57c00'; statusIcon = 'exclamation-circle'; statusText = 'Loop behöver justeras';
        } else {
            statusBg = '#e8f5e9'; statusBorder = '#43a047'; statusIcon = 'check-circle'; statusText = 'Loop presterar bra';
        }

        // Primary recommendation — most important single action
        let primaryRec = null;
        const firstWarn = advice.find(a => a.severity === 'warn');
        if (firstWarn) {
            primaryRec = `<i class="fa fa-exclamation-triangle" style="color:#e65100"></i> ${firstWarn.text}`;
        } else if (tuning.openLoop?.imc?.pi) {
            const imc = tuning.openLoop.imc.pi;
            const pStr = isPband ? `P% = ${toDisplayP(imc.kc) ?? '—'}` : `Kp = ${imc.kc}`;
            primaryRec = `<i class="fa fa-check-circle" style="color:#43a047"></i> Inget akut — IMC/Lambda föreslår: <b>${pStr}</b>, Ti = <b>${imc.ti} s</b>`;
        }

        // Data gap warning (inline)
        const gapWarn = metrics?.hasLargeGap
            ? `<div style="color:#e67e22;font-size:10px;margin-top:4px"><i class="fa fa-exclamation-triangle"></i> ${metrics.dataGaps} datahål detekterade (>${(metrics.medianIntervalMs * 3 / 1000).toFixed(0)}s) — FOPDT-modell kan vara opålitlig.</div>`
            : '';

        // Trend data note
        const trendNote = source === 'trend'
            ? `<div style="font-size:10px;color:#888;margin-top:4px"><i class="fa fa-info-circle"></i> Historisk trenddata — inställningar kan ha ändrats under fönstret.</div>`
            : '';

        const hasNonOkAdvice = advice.some(a => a.severity !== 'ok');

        resultsEl.innerHTML = `
<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;flex-wrap:wrap;gap:4px">
  <span style="font-size:11px;color:#888"><b>${prefix}</b> &middot; ${ct.label} &middot; ${winLabel} &middot; ${samplesLabel} &ensp;${confidencePill}</span>
  <button class="sec" id="inu-pid-back" style="font-size:10px;padding:2px 8px"><i class="fa fa-arrow-left"></i> Tillbaka</button>
</div>
<div style="background:${statusBg};border:1px solid ${statusBorder};border-radius:4px;padding:8px 12px;font-size:12px;font-weight:600;margin-bottom:8px;color:${statusBorder}">
  <i class="fa fa-${statusIcon}"></i> ${statusText}
  ${primaryRec ? `<div style="font-size:11px;font-weight:normal;margin-top:4px;color:#333">${primaryRec}</div>` : ''}
  ${gapWarn}${trendNote}
</div>
<div class="inu-pid-chart-wrap" id="inu-pid-chart-area"></div>
${tuningHtml}
<details${hasNonOkAdvice ? ' open' : ''} style="margin-bottom:8px">
  <summary style="cursor:pointer;font-size:11px;font-weight:600;list-style:none;padding:4px 0"><i class="fa fa-lightbulb-o"></i> Rekommendationer ${hasNonOkAdvice ? '' : '<span style="font-weight:normal;color:#888">(inga åtgärder)</span>'}</summary>
  <div style="margin-top:4px">${adviceHtml}</div>
</details>
<details style="margin-bottom:8px">
  <summary style="cursor:pointer;font-size:11px;font-weight:600;list-style:none;padding:4px 0"><i class="fa fa-bar-chart"></i> Analysdetaljer &amp; aktuella inställningar</summary>
  <div class="inu-pid-results" style="margin-top:6px">
    <div class="inu-pid-results-box">
      <h5>Analys</h5>
      ${metricRow('Kvarstående fel', metrics?.steadyStateError, '', 'Medelvärdet av (PV − SP) under de sista 20% av mätfönstret. Noll = ingen systematisk avvikelse. Högt värde = regulatorn hittar inte till börvärdet (kontrollera I-tid).')}
      ${metricRow('Svängning (pp)', metrics?.pvSwingPP ?? metrics?.oscillationAmplitude, '', 'Peak-to-peak variation av PV under analysfönstret. Stor svängning med litet kvarstående fel = processen oscillerar runt sitt medelvärde — typiskt processstörning. Stor svängning med högt kvarstående fel = regulatorn hänger inte med.')}
      ${metricRow('Överskjutning', metrics?.overshoot, '%', 'Hur långt PV passerar SP efter ett steg, i % av stegstorlek. Under 10% är normalt för HVAC. Högt värde → sänk Kp eller öka Ti.')}
      ${metricRow('Insvängningstid', metrics?.riseTime, ' s', 'Tid för PV att nå 90% av stegets slutvärde efter OP-steget. Lång tid → försök öka Kp försiktigt.')}
      ${metricRow('Insvängning (2%)', metrics?.settlingTime, ' s', 'Tid tills PV håller sig stabilt inom ±2% av börvärdet. Anger total insvängningstid inklusive eventuell överskjutning.')}
      ${metricRow('Oscillationer', metrics?.oscillationCount, '', 'Antal oscillationscykler (nollgenomgångar i felet PV−SP). Fler än 3–4 i ett normalt mätfönster tyder på instabilitet eller störning.')}
      ${metrics?.ultimatePeriod != null ? `<div class="inu-pid-metric-row"><span data-tip="Estimerad oscillationsperiod (sekunder). Används för ZN sluten krets om oscillationerna är regelbundna.">Period (Pu)</span><span class="inu-pid-metric-val">${metrics.ultimatePeriod.toFixed(1)} s</span></div>` : ''}
      ${metrics?.opMean != null ? `<div class="inu-pid-metric-row"><span data-tip="Medelvärde av utgångssignalen (OP/styrdon) under analysfönstret. Värdet nära 0% eller 100% kan tyda på under- eller överdimensionering.">OP medel</span><span class="inu-pid-metric-val">${metrics.opMean} %</span></div>` : ''}
      ${metrics?.opSwingPP != null ? `<div class="inu-pid-metric-row"><span data-tip="Peak-to-peak variation av utgångssignalen. Litet värde (< 3%) vid stor PV-svängning indikerar att regulatorn inte svarar på avvikelsen — kontrollera driftsläge och parametrar.">OP variation (pp)</span><span class="inu-pid-metric-val" style="${metrics.opSwingPP < 3 && (metrics.pvSwingPP ?? 0) > 0.2 ? 'color:#e67e22;font-weight:600' : ''}">${metrics.opSwingPP} %</span></div>` : ''}
      ${metrics?.stepDetected ? `<div class="inu-pid-metric-row"><span>Stegstorlek</span><span class="inu-pid-metric-val">${metrics.stepSize?.toFixed(2)}</span></div>` : '<div class="inu-pid-metric-row"><span style="color:#999">Inget stegrespons</span></div>'}
    </div>
    <div class="inu-pid-results-box">
      <h5>PID-värden (aktuella)</h5>
      ${pValueDisplay(pidValues?.p)}
      ${pidRow('Ti', pidValues?.i)}
      ${pidRow('Td', pidValues?.d)}
      <div style="margin-top:8px;font-size:10px;color:#999">Typvärden (${ct.label})<br>${isPband ? `P%: ${(100/ct.kpRange[1]).toFixed(0)}–${(100/ct.kpRange[0]).toFixed(0)}` : `Kp: ${ct.kpRange[0]}–${ct.kpRange[1]}`}<br>Ti: ${ct.tiRange[0]}–${ct.tiRange[1]} s</div>
    </div>
  </div>
</details>
<div class="inu-pid-btn-row">
  <button class="bx" id="inu-pid-close"><i class="fa fa-times"></i> Stäng</button>
</div>`;

        renderPidChart(resultsEl.querySelector('#inu-pid-chart-area'), pvPts, spPts, opPts);

        // Lambda slider — live recalculation of IMC Kc and Ti
        const lambdaSlider = resultsEl.querySelector('#inu-pid-lambda-slider');
        if (lambdaSlider && tuning.openLoop) {
            const ol = tuning.openLoop;
            lambdaSlider.addEventListener('input', () => {
                const lam = parseFloat(lambdaSlider.value);
                const L = ol.model.L, T = ol.model.T, Kp = Math.abs(ol.model.kp);
                const newKc = +(T / (Kp * (lam + L))).toFixed(3);
                const newTi = +(T + L / 2).toFixed(1); // Ti is independent of lambda
                const kcDisp = isPband ? (toDisplayP(newKc) ?? '—') + '%' : String(newKc);
                resultsEl.querySelector('#inu-pid-lambda-val').textContent = lam.toFixed(1);
                resultsEl.querySelector('#inu-pid-imc-kc').textContent = kcDisp;
                resultsEl.querySelector('#inu-pid-imc-ti').textContent = newTi;
            });
        }

        mo.querySelector('#inu-pid-back').addEventListener('click', () => {
            resultsEl.style.display = 'none';
            mo.querySelector('#inu-pid-setup').style.display = '';
        });
        mo.querySelector('#inu-pid-close').addEventListener('click', () => mo.remove());
    }

    // SPA navigation — re-detect page type when WebPort changes route without a full reload
    let _lastHref = location.href;
    setInterval(() => {
        if (location.href === _lastHref) return;
        _lastHref = location.href;
        // Always clean up editor toolbar when navigating away — it belongs only in edit mode
        if (!isPageEditorPage()) cleanupPageEditor();
        let spa_att = 0;
        const spa_wait = setInterval(() => {
            spa_att++;
            if (isPageEditorPage()) { clearInterval(spa_wait); initPageEditor(); }
            else if (isDevicePage())     { clearInterval(spa_wait); initDevicePage(); }
            else if (spa_att >= 30)      { clearInterval(spa_wait); }
        }, 200);
    }, 300);
})();