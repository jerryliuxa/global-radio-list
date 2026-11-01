/* 环球收听时钟 app.js
   思路：节目表统一以 UTC 存储；渲染时把"当前时刻"换算到所选城市时区。
   在播判定完全基于 UTC 时段 + 起始日（1=周日…7=周六），因此天然兼容跨日时段；
   显示层用 Intl.DateTimeFormat 按 IANA 时区格式化，自动处理夏令时。 */
(function () {
'use strict';

var DATA = window.RADIO_DATA, CITIES = window.RADIO_CITIES;
var E = DATA.entries, META = DATA.meta;

/* ---------- 界面语言（zh / en） ---------- */
var UI = 'zh';
try { UI = localStorage.getItem('rg_lang') || ((navigator.language || 'zh').toLowerCase().indexOf('en') === 0 ? 'en' : 'zh'); } catch (e) {}
var I18N = {
  zh: {
    docTitle: '环球收听时钟 · Global Radio Guide — 长波/中波/短波(DRM)/FM 在播节目表',
    h1: '环球收听时钟', h1sub: 'Global Radio Guide',
    tagline: '不同国家城市的听友 · 按当地时间看此刻在播的长波 / 中波 / 短波（DRM）/ FM 节目与频率',
    myLoc: '📍 我在', sdr: '🎧 SDR 收听',
    customPh: '✏️ 任意收听地点：城市名 / IANA 时区 / UTC±N',
    auto: '跟随本机时区（自动）', deviceTz: '本机时区', customMark: '（自定义）',
    all: '全部',
    band: { SW: '短波 SW', MW: '中波 MW', LW: '长波 LW', time: '授时台', FM: 'FM 调频', other: '其他' },
    src: { aoki: '青木 B26', fm: 'FM 城市' },
    allSrc: '全部来源', srcLbl: '数据源',
    onair: '正在播出', next: '即将播出', allv: '全部检索',
    searchPh: '🔍 搜电台 / 频率 / 节目名…', allLang: '全部语言',
    live: '● 在播', drm: 'DRM 数字',
    nextday: '（次日）', full24: '全天（24 小时）', utcFull: 'UTC 全天',
    startsIn: '约 {x}后开始', played: '已播 {p} · 剩 {l}', h: ' 时 ', m: ' 分',
    az: '方位', tgt: '目标区',
    txCoord: '发射台坐标', distT: '大圆距离；方位为从你的接收位置指向发射台的方向（天线对准参考）',
    fromYou: '在你 {b}° {d}',
    mapBtn: '🔍 地图', mapT: '在 OpenStreetMap 上查看两点连线与直线距离',
    listenT: '用 {s} 打开并调谐到 {f} kHz',
    backup: '⋯ 备用 {n} 台', backupT: '用其他接收机收听 {f} kHz',
    stats: '📍 {c} · 当地 {t} · 全球此刻共有 {f} 个频率 / {p} 路节目正在播出（当前筛选显示 {n} 个频率）',
    headItems: '{a} 个频率 · {b} 个节目', headAll: '全部条目：共 {n} 条，显示前 {m} 条',
    headNext: '未来 3 小时内开播 {n} 条',
    more: '▾ 还有 {n} 路同频节目…',
    cap: '仅显示前 {n} 个频率 — 换个筛选条件缩小范围', loadMore: '加载更多（还有 {n} 条）',
    emptyLW: '当前数据源（HFCC B26 青木格式频率表）为短波广播表（2485~26040 kHz），<br>不含长波 / 中波 / 授时台条目。',
    emptyFM: '「{c}」暂无 FM 数据。<br>FM 数据现覆盖 {n} 个城市（radiomap.eu，以欧洲为主）——可在上方输入框切换到这些城市，或用搜索在其他城市中查找电台。',
    emptyNoFM: '当前数据集中没有 FM 条目。<br>接入 FM 数据后，该频段会自动显示在这里。',
    emptyNext: '未来 3 小时内没有符合筛选条件的开播节目。<br>试试放宽筛选，或切换到「正在播出」。',
    emptyGeneric: '没有符合当前筛选条件的条目。<br>试试更换语言 / 数据源 / 频段，或清空搜索词。',
    f1: '数据来源：<b>HFCC B26 频率表（青木格式，用户提供 xlsx，2026-10-30 版）</b>——短波广播 2485~26040 kHz 共 3115 条（DRM 由 HFCC 原始表 MOD="N" 判定，已按 FDATE/TDATE 做播出季有效性过滤），当前季 B26（2026-10-25 ~ 2027-03-27）· <b>radiomap.eu 各城市 FM 电台（54 城，全天播出）</b>。',
    f2: '时段均为广播时刻表，不包含电离层传播预测；同一频率可能有多地发射机同时播出。FM 条目按所选城市显示（已覆盖 54 个城市，以欧洲为主，其他城市可搜索）。本表为短波表，不含长波 / 中波 / 授时台。',
    f3: '生成于 {g}（数据截至 {u}） · 本地刷新频率 20 秒 · 数据工具：<code>tools/build_data.py</code>',
    sdrT: '管理 WebSDR / KiwiSDR / OpenWebRX 接收机',
    sdrHint: '点击频率卡片上的 🎧 按钮，会用<b>默认接收机</b>打开并自动调谐到该频率。支持 KiwiSDR、OpenWebRX、WebSDR（如 <code>http://jjm0311.proxy.kiwisdr.com:8073/</code>）。',
    sdrEmpty: '还没有接收机，请在下方添加地址。',
    sdrUrlPh: '接收机地址，如 http://…:8073/ 或 https://…/', sdrNamePh: '名称（留空自动取域名）',
    sdrAuto: '自动识别类型', sdrAdd: '＋ 添加', sdrDel: '删除', sdrDefT: '设为默认接收机',
    sdrBadUrl: '地址无效：请粘贴完整 URL（含 http:// 或 https://）', sdrProto: '仅支持 http/https 地址',
    sdrDup: '该接收机已在列表中', sdrAdded: '✓ 已添加 {n}（{t}）',
    sdrReset: '↺ 恢复默认', sdrResetT: '放弃本地修改，重新采用 js/sdrs.js 配置文件中的列表'
  },
  en: {
    docTitle: 'Global Radio Guide — LW/MW/SW (DRM)/FM schedules on air now, in your local time',
    h1: 'Global Radio Guide', h1sub: 'International Edition',
    tagline: 'Listeners worldwide · see what is on air right now on LW / MW / SW (DRM) / FM, in your local time',
    myLoc: '📍 My location', sdr: '🎧 SDR Listen',
    customPh: '✏️ Any location: city / IANA timezone / UTC±N',
    auto: 'Follow device timezone (auto)', deviceTz: 'Device timezone', customMark: ' (custom)',
    all: 'All',
    band: { SW: 'Shortwave SW', MW: 'Mediumwave MW', LW: 'Longwave LW', time: 'Time Signal', FM: 'FM', other: 'Other' },
    src: { aoki: 'Aoki B26', fm: 'FM Cities' },
    allSrc: 'All sources', srcLbl: 'Source',
    onair: 'On Air', next: 'Coming Up', allv: 'All Entries',
    searchPh: '🔍 Search stations / frequencies / programs…', allLang: 'All languages',
    live: '● On Air', drm: 'DRM Digital',
    nextday: ' (next day)', full24: 'Around the clock (24 h)', utcFull: 'UTC all day',
    startsIn: 'starts in {x}', played: 'played {p} · left {l}', h: ' h ', m: ' min',
    az: 'Azimuth', tgt: 'Target zones',
    txCoord: 'Transmitter coordinates', distT: 'Great-circle distance; the bearing is from your receiver toward the transmitter (antenna-aiming reference)',
    fromYou: 'bearing {b}° {d} from you',
    mapBtn: '🔍 Map', mapT: 'View the two-point line and straight-line distance on OpenStreetMap',
    listenT: 'Open {s} tuned to {f} kHz',
    backup: '⋯ {n} backups', backupT: 'Listen to {f} kHz on another receiver',
    stats: '📍 {c} · local {t} · {f} frequencies / {p} programs on air worldwide ({n} shown)',
    headItems: '{a} frequencies · {b} programs', headAll: 'All entries: {n} total, showing first {m}',
    headNext: '{n} programs starting within 3 hours',
    more: '▾ {n} more programs on this frequency…',
    cap: 'Showing first {n} frequencies — narrow the filters', loadMore: 'Load more ({n} more)',
    emptyLW: 'The current source (HFCC B26 Aoki-format schedule) covers shortwave only (2485–26040 kHz);<br>no longwave / mediumwave / time-signal entries.',
    emptyFM: 'No FM data for "{c}" yet.<br>FM data covers {n} cities (radiomap.eu, mostly Europe) — switch cities above, or use search to find stations in other cities.',
    emptyNoFM: 'No FM entries in the current dataset.<br>They will appear here once FM data is added.',
    emptyNext: 'Nothing starts within the next 3 hours for these filters.<br>Try widening the filters or switch to "On Air".',
    emptyGeneric: 'No entries match the current filters.<br>Try changing language / source / band, or clear the search.',
    f1: 'Data sources: <b>HFCC B26 schedule in Aoki format (user-provided xlsx, 2026-10-30 edition)</b> — shortwave broadcasting 2485–26040 kHz, 3115 entries (DRM flagged from MOD="N" in the raw HFCC table; season validity filtered by FDATE/TDATE), season B26 (2026-10-25 ~ 2027-03-27) · <b>radiomap.eu city FM stations (54 cities, 24 h)</b>.',
    f2: 'Times are broadcast schedules and do not include ionospheric propagation prediction; one frequency may be transmitted from several sites. FM entries are shown for the selected city (54 cities covered, mostly Europe; searchable across cities). This is a shortwave table — no longwave / mediumwave / time signals.',
    f3: 'Generated {g} (data as of {u}) · refreshes every 20 s · data tool: <code>tools/build_data.py</code>',
    sdrT: 'Manage WebSDR / KiwiSDR / OpenWebRX receivers',
    sdrHint: 'Click the 🎧 button on any frequency card to open your <b>default receiver</b> tuned to that frequency. Supports KiwiSDR, OpenWebRX and WebSDR (e.g. <code>http://jjm0311.proxy.kiwisdr.com:8073/</code>).',
    sdrEmpty: 'No receivers yet — add one below.',
    sdrUrlPh: 'Receiver URL, e.g. http://…:8073/ or https://…/', sdrNamePh: 'Name (blank = hostname)',
    sdrAuto: 'Detect type automatically', sdrAdd: '＋ Add', sdrDel: 'Delete', sdrDefT: 'Set as default receiver',
    sdrBadUrl: 'Invalid URL: paste a full URL including http:// or https://', sdrProto: 'Only http/https URLs are supported',
    sdrDup: 'This receiver is already in the list', sdrAdded: '✓ Added {n} ({t})',
    sdrReset: '↺ Reset to defaults', sdrResetT: 'Discard local changes and re-apply the list from js/sdrs.js'
  }
};
function t(k, vars) {
  var s = I18N[UI] && I18N[UI][k] !== undefined ? I18N[UI][k] : I18N.zh[k];
  if (s === undefined) return k;
  if (vars) for (var v in vars) s = s.split('{' + v + '}').join(vars[v]);
  return s;
}
function enPart(n) {
  var i = n.indexOf(' ');
  var e = i > 0 ? n.slice(i + 1) : n;
  e = e.replace(/（[^）]*）/g, '').replace(/\([^)]*[\u4e00-\u9fff][^)]*\)/g, '').trim();
  return e || n;
}
function dispLabel(n) { return UI === 'en' ? enPart(n) : n; }
var BANDS = ['SW', 'MW', 'LW', 'time', 'FM'];
function bandName(b) { return I18N[UI].band[b] || I18N.zh.band[b] || b; }
function srcName(s) { return I18N[UI].src[s] || I18N.zh.src[s] || s; }
var DAY_MS = 86400000, CAP_NOW = 500, CAP_NEXT = 60, CAP_ALL = 300;

var state = {
  tz: null, cityLabel: '', band: 'all', src: 'all', lang: 'all',
  q: '', view: 'now', expand: {}, limitAll: CAP_ALL
};

/* ---------- 小工具 ---------- */
function $(id) { return document.getElementById(id); }
function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
function flag(iso) {
  if (!iso) return '';
  return String(iso).toUpperCase().replace(/./g, function (c) {
    return String.fromCodePoint(127397 + c.charCodeAt(0));
  });
}
function pad2(n) { return (n < 10 ? '0' : '') + n; }
function meterBand(f) { return f > 0 && f < 30000 ? Math.round(300000 / f) + 'm' : ''; }
function fmtDur(mins) {
  mins = Math.max(0, Math.round(mins));
  if (mins >= 60) return Math.floor(mins / 60) + t('h') + (mins % 60) + t('m');
  return mins + t('m');
}

/* ---------- 位置：坐标格式化与大圆距离/方位 ---------- */
function compass8() {
  return UI === 'en' ? ['N', 'NE', 'E', 'SE', 'S', 'SW', 'W', 'NW']
    : ['北', '东北', '东', '东南', '南', '西南', '西', '西北'];
}
function fmtLat(v) {
  var h = v >= 0 ? 'N' : 'S'; v = Math.abs(v);
  var d = Math.floor(v), m = Math.round((v - d) * 60);
  if (m === 60) { d++; m = 0; }
  return d + '°' + pad2(m) + '′' + h;
}
function fmtLon(v) {
  var h = v >= 0 ? 'E' : 'W'; v = Math.abs(v);
  var d = Math.floor(v), m = Math.round((v - d) * 60);
  if (m === 60) { d++; m = 0; }
  return d + '°' + pad2(m) + '′' + h;
}
function fmtLatLon(ll) { return fmtLat(ll[0]) + ' ' + fmtLon(ll[1]); }
function distBearing(from, to) {
  var rad = Math.PI / 180, R = 6371;
  var p1 = from[0] * rad, p2 = to[0] * rad;
  var dp = (to[0] - from[0]) * rad, dl = (to[1] - from[1]) * rad;
  var a = Math.sin(dp / 2) * Math.sin(dp / 2) + Math.cos(p1) * Math.cos(p2) * Math.sin(dl / 2) * Math.sin(dl / 2);
  var km = Math.round(2 * R * Math.asin(Math.sqrt(a)) / 10) * 10;
  var brg = (Math.atan2(Math.sin(dl) * Math.cos(p2),
    Math.cos(p1) * Math.sin(p2) - Math.sin(p1) * Math.cos(p2) * Math.cos(dl)) / rad + 360) % 360;
  return { km: km, brg: Math.round(brg), dir: compass8()[Math.round(brg / 45) % 8] };
}

/* ---------- 时区格式化（缓存） ---------- */
var fmtCache = {};
function fmt(tz) {
  var key = UI + '|' + tz;
  if (!fmtCache[key]) {
    var loc = UI === 'en' ? 'en-GB' : 'zh-CN';
    var off;
    try {
      off = new Intl.DateTimeFormat(loc, { timeZone: tz, timeZoneName: 'shortOffset' })
        .formatToParts(new Date()).find(function (p) { return p.type === 'timeZoneName'; }).value;
    } catch (e) {
      off = '';
    }
    fmtCache[key] = {
      time: new Intl.DateTimeFormat(loc, { timeZone: tz, hour: '2-digit', minute: '2-digit', hour12: false }),
      sec: new Intl.DateTimeFormat(loc, { timeZone: tz, hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }),
      date: new Intl.DateTimeFormat(loc, { timeZone: tz, month: '2-digit', day: '2-digit' }),
      dateLong: new Intl.DateTimeFormat(loc, { timeZone: tz, year: 'numeric', month: 'long', day: 'numeric', weekday: 'short' }),
      off: String(off).replace('GMT', 'UTC')
    };
  }
  return fmtCache[key];
}
function hm(tz, ms) { return fmt(tz).time.format(ms); }
function dateKey(tz, ms) { return fmt(tz).date.format(ms); }

function utcMin(ms) { var d = new Date(ms); return d.getUTCHours() * 60 + d.getUTCMinutes(); }
function dowUTC(ms) { return new Date(ms).getUTCDay(); }           // 0=周日
function dayNum(dow) { return dow + 1; }                            // 1=周日 … 7=周六
function utcMidnight(ms) { return Math.floor(ms / DAY_MS) * DAY_MS; }
function dowOfMs(ms) { return new Date(ms).getUTCDay(); }

/* ---------- 在播判定（全 UTC） ---------- */
function inDays(it, dn) { return it.d.indexOf(String(dn)) >= 0; }
function durOf(it) { return ((it.e - it.s + 1440) % 1440) || 1440; }

function onAir(it, now) {
  var m = utcMin(now), dow = dowUTC(now), s = it.s, e = it.e;
  if (s < e) return inDays(it, dayNum(dow)) && m >= s && m < e;
  // 跨日（含全天）
  if (m >= s) return inDays(it, dayNum(dow));
  if (m < e) return inDays(it, dayNum((dow + 6) % 7));
  return false;
}

/* 当前/下一次播出区间的 UTC 绝对时刻 */
function occurrence(it, now) {
  var day0 = utcMidnight(now), m = utcMin(now), dow = dowUTC(now);
  var s = it.s, dur = durOf(it);
  if (onAir(it, now)) {
    var st = (m >= s ? day0 : day0 - DAY_MS) + s * 60000;
    return { start: st, end: st + dur * 60000, live: true };
  }
  for (var k = 0; k < 8; k++) {
    var d0 = day0 + k * DAY_MS;
    if (inDays(it, dayNum(dowOfMs(d0)))) {
      var s0 = d0 + s * 60000;
      if (s0 > now) return { start: s0, end: s0 + dur * 60000, live: false };
    }
  }
  return null;
}

/* FM 城市归属：选中城市若本身有 FM 数据则精确匹配，否则显示同时区全部 FM */
var fmCityLabels = {};
E.forEach(function (it) { if (it.src === 'fm' && it.city) fmCityLabels[it.city] = it.tz; });
function fmCityVisible(it) {
  return it.city === state.cityLabel ||
    (it.tz === state.tz && !fmCityLabels.hasOwnProperty(state.cityLabel));
}

/* ---------- 筛选 ---------- */
function passFilter(it) {
  if (state.band === 'drm') { if (!it.drm) return false; }
  else if (state.band !== 'all' && it.band !== state.band) return false;
  if (state.src !== 'all' && it.src !== state.src) return false;
  var lgKey = UI === 'en' ? (it.lge || 'Unspecified') : (it.lgn || '未指定');
  if (state.lang !== 'all' && lgKey !== state.lang) return false;
  if (it.src === 'fm' && !state.q && !fmCityVisible(it)) return false;
  if (state.q) {
    var hay = (it.st + ' ' + (it.nm || '') + ' ' + (it.lg || '') + ' ' + (it.lgn || '') + ' ' +
      it.f + ' ' + (it.si || '') + ' ' + (it.cn || '') + ' ' + (it.nt || '')).toLowerCase();
    if (hay.indexOf(state.q.toLowerCase()) < 0) return false;
  }
  return true;
}

/* ---------- 卡片渲染 ---------- */
function badges(it, live) {
  var h = '<span class="badge b-band">' + bandName(it.band) + '</span>';
  if (it.drm) h += '<span class="badge b-drm">' + t('drm') + '</span>';
  if (live) h += '<span class="badge b-live">' + t('live') + '</span>';
  h += '<span class="badge b-src">' + srcName(it.src) + '</span>';
  return h;
}
function mapHref(it) {
  if (!it.ll || !state.cityCoord) return '';
  var azv = String(it.az || '').trim();
  var az = /^\d+(\.\d+)?$/.test(azv) ? azv : '';
  return 'map.html?a=' + state.cityCoord[0] + ',' + state.cityCoord[1] +
    '&b=' + it.ll[0] + ',' + it.ll[1] +
    '&na=' + encodeURIComponent(dispLabel(state.cityLabel)) +
    '&nb=' + encodeURIComponent(((it.st || '') + (it.si ? ' @ ' + it.si : '')).slice(0, 70)) +
    (az ? '&az=' + az : '') +
    '&lang=' + UI;
}
function metaLine(it) {
  var bits = [];
  var lgd = UI === 'en' ? (it.lge || it.lg || 'Unspecified') : (it.lgn || it.lg || '未指定');
  if (lgd) bits.push('<span class="m">' + esc(lgd) + '</span>');
  var place = (UI === 'en' ? (it.ce || it.cn) : (it.cn || it.ce)) || '';
  var cty = UI === 'en' ? (it.cty || (it.city ? enPart(it.city) : '')) : it.city;
  if (cty && cty !== place) place += (place ? ' · ' : '') + cty;
  if (it.si && it.si !== it.city && it.si !== cty) place += (place ? ' · ' : '') + it.si;
  if (place) bits.push('<span class="m">' + flag(it.iso) + ' ' + esc(place) + '</span>');
  if (it.ll) bits.push('<span class="m" title="' + t('txCoord') + '">📡 ' + fmtLatLon(it.ll) + '</span>');
  if (it.ll && state.cityCoord) {
    var db = distBearing(state.cityCoord, it.ll);
    bits.push('<span class="m" title="' + t('distT') + '">↔ ' +
      db.km.toLocaleString('zh-CN') + ' km · ' + t('fromYou', { b: db.brg, d: db.dir }) + '</span>');
    var mh = mapHref(it);
    if (mh) bits.push('<a class="maplink" href="' + mh + '" target="_blank" rel="noopener noreferrer" ' +
      'title="' + t('mapT') + '">' + t('mapBtn') + '</a>');
  }
  if (it.pw) bits.push('<span class="m">' + esc(it.pw) + ' kW</span>');
  if (it.az) bits.push('<span class="m">' + t('az') + ' ' + esc(it.az) + '</span>');
  if (it.tgt) bits.push('<span class="m">' + t('tgt') + ' ' + esc(it.tgt) + '</span>');
  return bits.length ? '<div class="meta">' + bits.join('') + '</div>' : '';
}
function timeLine(occ, now) {
  var tz = state.tz, f = fmt(tz);
  var fullDay = occ.end - occ.start >= 23.9 * 3600000;
  var local = fullDay ? t('full24') : hm(tz, occ.start) + ' – ' + hm(tz, occ.end) +
    (dateKey(tz, occ.start) !== dateKey(tz, occ.end) ? '<small>' + t('nextday') + '</small>' : '');
  var utc = fullDay ? t('utcFull') : 'UTC ' + pad2(Math.floor(occ.start / 3600000 % 24)) + ':' + pad2(Math.floor(occ.start / 60000 % 60)) +
    '–' + pad2(Math.floor(occ.end / 3600000 % 24)) + ':' + pad2(Math.floor(occ.end / 60000 % 60));
  if (!occ.live) {
    var mins = (occ.start - now) / 60000;
    return '<div class="timerow"><span class="t-local" style="color:var(--blue)">' + local +
      '</span><span class="remain">' + t('startsIn', { x: fmtDur(mins) }) + '</span><span class="t-utc">' + utc + '</span></div>';
  }
  var durMs = occ.end - occ.start, pct = Math.max(0, Math.min(100, (now - occ.start) / durMs * 100));
  if (fullDay) {
    return '<div class="timerow"><span class="t-local">' + local +
      '</span><span class="progbar"><i style="width:' + pct.toFixed(1) + '%"></i></span><span class="t-utc">' + utc + '</span></div>';
  }
  var played = (now - occ.start) / 60000, left = (occ.end - now) / 60000;
  return '<div class="timerow"><span class="t-local">' + local +
    '</span><span class="progbar"><i style="width:' + pct.toFixed(1) + '%"></i></span>' +
    '<span class="remain">' + t('played', { p: fmtDur(played), l: fmtDur(left) }) + '</span><span class="t-utc">' + utc + '</span></div>';
}
function cardInner(it, occ, now) {
  var mb = meterBand(it.f);
  var freqDisp = it.band === 'FM'
    ? (it.f / 1000).toFixed(1) + '<span class="unit"> MHz</span>'
    : it.f + '<span class="unit"> kHz</span>';
  var h = '<div class="freqrow"><span class="freq">' + freqDisp + '</span>' +
    (mb ? '<span class="badge b-band">' + mb + '</span>' : '') + badges(it, occ.live) + '</div>';
  h += '<div class="stname">' + esc(it.st || '—') + '</div>';
  if (it.nm && it.nm !== it.st) h += '<div class="prog">' + esc(it.nm) + '</div>';
  h += metaLine(it);
  h += timeLine(occ, now);
  h += listenRow(it.f);
  if (it.nt) h += '<div class="note">📌 ' + esc(it.nt) + '</div>';
  return h;
}

/* ---------- 视图：正在播出（按频率分组） ---------- */
function renderNow(now) {
  var list = E.filter(function (it) { return passFilter(it) && onAir(it, now); });
  var byFreq = {};
  list.forEach(function (it) { (byFreq[it.f] = byFreq[it.f] || { band: it.band, items: [] }).items.push(it); });
  var freqs = Object.keys(byFreq).map(Number).sort(function (a, b) { return a - b; });

  var html = '', shown = 0, truncated = false;
  BANDS.forEach(function (band) {
    var fs = freqs.filter(function (f) { return byFreq[f].band === band; });
    if (!fs.length) return;
    var nItems = fs.reduce(function (a, f) { return a + byFreq[f].items.length; }, 0);
    html += '<div class="bandhead">' + bandName(band) + ' <span class="n">' + t('headItems', { a: fs.length, b: nItems }) + '</span></div><div class="grid">';
    fs.forEach(function (f) {
      if (shown >= CAP_NOW) { truncated = true; return; }
      shown++;
      var grp = byFreq[f];
      grp.items.sort(function (a, b) { return (b.drm || 0) - (a.drm || 0) || a.s - b.s; });
      var main = grp.items[0];
      var occ = occurrence(main, now);
      var open = state.expand['f' + f];
      var body = cardInner(main, occ, now);
      if (grp.items.length > 1) {
        var rest = grp.items.slice(1);
        var visible = open ? rest : rest.slice(0, 2);
        body += '<div class="sublist">';
        visible.forEach(function (it2) {
          var occ2 = occurrence(it2, now);
          body += '<div class="subrow"><div class="s1">' + esc(it2.st || '—') +
            (it2.nm && it2.nm !== it2.st ? ' — ' + esc(it2.nm) : '') + '</div>' +
            '<div class="s2">' + timeLine(occ2, now) + metaLine(it2) + '</div></div>';
        });
        if (!open && rest.length > 2) body += '<span class="more" data-expand="' + f + '">' + t('more', { n: rest.length - 2 }) + '</span>';
        body += '</div>';
      }
      html += '<div class="card onair">' + body + '</div>';
    });
    html += '</div>';
  });
  if (!shown) html = emptyHtml();
  if (truncated) html += '<button class="loadmore" id="btnMore">' + t('cap', { n: CAP_NOW }) + '</button>';
  return { html: html, nFreq: freqs.length, nItems: list.length };
}

/* ---------- 视图：即将播出 ---------- */
function renderNext(now) {
  var list = [];
  E.forEach(function (it) {
    if (!passFilter(it) || onAir(it, now)) return;
    var occ = occurrence(it, now);
    if (occ && occ.start - now <= 3 * 3600000) list.push({ it: it, occ: occ });
  });
  list.sort(function (a, b) { return a.occ.start - b.occ.start; });
  var html = '<div class="bandhead">' + t('headNext', { n: Math.min(list.length, CAP_NEXT) }) + '</div><div class="grid">';
  list.slice(0, CAP_NEXT).forEach(function (x) {
    html += '<div class="card next">' + cardInner(x.it, x.occ, now) + '</div>';
  });
  html += '</div>';
  if (!list.length) html = emptyHtml();
  return { html: html, nFreq: list.length, nItems: list.length };
}

/* ---------- 视图：全部检索 ---------- */
function renderAll(now) {
  var list = E.filter(passFilter);
  list.sort(function (a, b) {
    var ba = BANDS.indexOf(a.band) < 0 ? 9 : BANDS.indexOf(a.band);
    var bb = BANDS.indexOf(b.band) < 0 ? 9 : BANDS.indexOf(b.band);
    return ba - bb || a.f - b.f || a.s - b.s;
  });
  var cap = state.limitAll;
  var html = '<div class="bandhead">' + t('headAll', { n: list.length, m: Math.min(cap, list.length) }) + '</div><div class="grid">';
  list.slice(0, cap).forEach(function (it) {
    var occ = occurrence(it, now);
    html += '<div class="card' + (occ && occ.live ? ' onair' : '') + '">' + cardInner(it, occ, now) + '</div>';
  });
  html += '</div>';
  if (list.length > cap) html += '<button class="loadmore" id="btnMoreAll">' + t('loadMore', { n: list.length - cap }) + '</button>';
  if (!list.length) html = emptyHtml();
  return { html: html, nFreq: list.length, nItems: list.length };
}

function emptyHtml() {
  var msg;
  var nFm = fmCityLabels ? Object.keys(fmCityLabels).length : 0;
  if (state.band === 'FM') {
    if (nFm) {
      msg = t('emptyFM', { c: dispLabel(state.cityLabel), n: nFm });
    } else {
      msg = t('emptyNoFM');
    }
  } else if (state.view === 'next') {
    msg = t('emptyNext');
  } else if (state.band === 'LW' || state.band === 'MW' || state.band === 'time') {
    msg = t('emptyLW');
  } else {
    msg = t('emptyGeneric');
  }
  return '<div class="empty"><span class="big">📡</span>' + msg + '</div>';
}

/* ---------- 渲染入口 ---------- */
function render() {
  var now = Date.now();
  var r = state.view === 'next' ? renderNext(now) : state.view === 'all' ? renderAll(now) : renderNow(now);
  $('list').innerHTML = r.html;
  var liveCount = 0, liveItems = 0;
  if (state.view === 'now') { liveCount = r.nFreq; liveItems = r.nItems; }
  $('stats').innerHTML = t('stats', {
    c: '<b>' + esc(dispLabel(state.cityLabel)) + '</b>',
    t: '<b>' + hm(state.tz, now) + '</b>',
    f: '<b>' + countLive()[0] + '</b>',
    p: '<b>' + countLive()[1] + '</b>',
    n: r.nFreq
  });
  bindMore();
}

function bindMore() {
  var m = $('btnMoreAll');
  if (m) m.onclick = function () { state.limitAll += CAP_ALL; render(); };
  document.querySelectorAll('[data-expand]').forEach(function (el) {
    el.onclick = function () { state.expand['f' + el.getAttribute('data-expand')] = true; render(); };
  });
}

var liveCache = { t: 0, v: [0, 0] };
function countLive() {
  var now = Date.now();
  if (now - liveCache.t > 30000) {
    var freqs = {}, n = 0;
    E.forEach(function (it) {
      if (onAir(it, now) && (it.src !== 'fm' || fmCityVisible(it))) { n++; freqs[it.f] = 1; }
    });
    liveCache = { t: now, v: [Object.keys(freqs).length, n] };
  }
  return liveCache.v;
}

/* ---------- SDR 收音机（KiwiSDR / OpenWebRX / WebSDR 联动） ---------- */
var SDR_TYPES = { kiwi: 'KiwiSDR', openwebrx: 'OpenWebRX', websdr: 'WebSDR' };
/* SDR 列表：js/sdrs.js 为默认配置（rev 变化时全员更新），localStorage 保存访客自定义 */
var SDR_CFG = window.RADIO_SDRS_CFG || { rev: 0, sdrs: [] };
var SDRS = (function () {
  try {
    var saved = JSON.parse(localStorage.getItem('rg_sdrs') || 'null');
    var savedRev = localStorage.getItem('rg_sdrs_rev');
    if (Array.isArray(saved) && saved.length && (savedRev === String(SDR_CFG.rev) || !savedRev)) {
      return saved;   // 访客自定义（或旧数据）优先生效
    }
  } catch (e) {}
  var list = Array.isArray(SDR_CFG.sdrs) && SDR_CFG.sdrs.length
    ? JSON.parse(JSON.stringify(SDR_CFG.sdrs))
    : [{ name: 'KiwiSDR', url: 'http://jjm0311.proxy.kiwisdr.com:8073/', type: 'kiwi', def: true }];
  try {
    localStorage.setItem('rg_sdrs', JSON.stringify(list));
    localStorage.setItem('rg_sdrs_rev', String(SDR_CFG.rev));
  } catch (e) {}
  return list;
})();
function resetSdrsToConfig() {
  SDRS = JSON.parse(JSON.stringify(Array.isArray(SDR_CFG.sdrs) && SDR_CFG.sdrs.length
    ? SDR_CFG.sdrs
    : [{ name: 'KiwiSDR', url: 'http://jjm0311.proxy.kiwisdr.com:8073/', type: 'kiwi', def: true }]));
  saveSdrs(); renderSdrModal(); render();
}
function saveSdrs() {
  try {
    localStorage.setItem('rg_sdrs', JSON.stringify(SDRS));
    localStorage.setItem('rg_sdrs_rev', String(SDR_CFG.rev));
  } catch (e) {}
}
function defaultSdr() {
  return SDRS.filter(function (s) { return s.def; })[0] || SDRS[0];
}
function detectSdrType(url) {
  var u = url.toLowerCase();
  if (u.indexOf('kiwi') >= 0) return 'kiwi';
  if (u.indexOf('websdr') >= 0) return 'websdr';
  if (u.indexOf('webrx') >= 0) return 'openwebrx';
  return 'kiwi';
}
/* 生成调谐链接：
   KiwiSDR   http://host:port/?f=9570am-z8   （kHz + 模式 + 变焦）
   OpenWebRX http://host/#freq=9570000,mod=am （Hz，hash 参数）
   WebSDR    http://host/?tune=9570am         （Twente 系统格式） */
function sdrHref(sdr, fKhz) {
  var f = (Math.round(fKhz * 100) / 100).toString();
  var mode = 'am';
  try {
    if (sdr.type === 'openwebrx') {
      var u = new URL(sdr.url);
      u.hash = '#freq=' + Math.round(fKhz * 1000) + ',mod=' + mode;
      return u.href;
    }
    var base = sdr.url.split('#')[0];
    var sep = base.indexOf('?') >= 0 ? '&' : '?';
    if (sdr.type === 'websdr') return base + sep + 'tune=' + f + mode;
    return base + sep + 'f=' + f + mode + '-z8';
  } catch (e) {
    return sdr.url;
  }
}
function listenRow(fKhz) {
  var d = defaultSdr();
  if (!d) return '';
  var h = '<div class="listenrow"><a class="listen" href="' + esc(sdrHref(d, fKhz)) +
    '" target="_blank" rel="noopener noreferrer" title="' + t('listenT', { s: d.name, f: fKhz }) + '">🎧 ' +
    esc(d.name) + '</a>';
  if (SDRS.length > 1) {
    h += '<span class="sdrmenu-wrap"><button type="button" class="sdrmore" data-sdrmenu title="' +
      t('backupT', { f: fKhz }) + '">' + t('backup', { n: SDRS.length - 1 }) + '</button><span class="sdrmenu hidden">';
    SDRS.forEach(function (s) {
      if (s === d) return;
      h += '<a class="listen" href="' + esc(sdrHref(s, fKhz)) + '" target="_blank" rel="noopener noreferrer">' +
        '▸ ' + esc(s.name) + ' <span class="t">' + (SDR_TYPES[s.type] || s.type) + '</span></a>';
    });
    h += '</span></span>';
  }
  h += '</div>';
  return h;
}
function closeSdrMenus() {
  document.querySelectorAll('.sdrmenu').forEach(function (m) { m.classList.add('hidden'); });
}
function renderSdrModal() {
  $('sdrList').innerHTML = SDRS.length ? SDRS.map(function (s, i) {
    return '<div class="sdr-item">' +
      '<input type="radio" name="sdrdef" data-i="' + i + '" title="' + t('sdrDefT') + '"' + (s.def ? ' checked' : '') + '>' +
      '<div class="grow"><div class="nm">' + esc(s.name) +
      ' <span class="badge t-type">' + (SDR_TYPES[s.type] || s.type) + '</span></div>' +
      '<div class="url">' + esc(s.url) + '</div></div>' +
      '<button class="del" data-del="' + i + '">' + t('sdrDel') + '</button></div>';
  }).join('') : '<p class="hint">' + t('sdrEmpty') + '</p>';
  $('sdrCount').textContent = SDRS.length ? '×' + SDRS.length : '';
  var rb = $('sdrReset');
  if (rb) rb.style.display = (Array.isArray(SDR_CFG.sdrs) && SDR_CFG.sdrs.length) ? '' : 'none';
  $('sdrList').querySelectorAll('input[name=sdrdef]').forEach(function (r) {
    r.onchange = function () {
      var i = +r.getAttribute('data-i');
      SDRS.forEach(function (s, j) { s.def = j === i; });
      saveSdrs(); renderSdrModal(); render();
    };
  });
  $('sdrList').querySelectorAll('[data-del]').forEach(function (b) {
    b.onclick = function () {
      SDRS.splice(+b.getAttribute('data-del'), 1);
      if (SDRS.length && !defaultSdr()) SDRS[0].def = true;
      saveSdrs(); renderSdrModal(); render();
    };
  });
}
function addSdr() {
  var raw = $('sdrUrl').value.trim();
  var msg = $('sdrMsg');
  msg.textContent = ''; msg.classList.remove('ok');
  var u;
  try { u = new URL(raw); } catch (e) { msg.textContent = t('sdrBadUrl'); return; }
  if (u.protocol !== 'http:' && u.protocol !== 'https:') { msg.textContent = t('sdrProto'); return; }
  var type = $('sdrType').value;
  if (type === 'auto') type = detectSdrType(raw);
  var name = $('sdrName').value.trim() || u.hostname;
  if (SDRS.some(function (s) { return s.url.replace(/\/+$/, '') === u.href.replace(/\/+$/, ''); })) {
    msg.textContent = t('sdrDup');
    return;
  }
  SDRS.push({ name: name, url: u.href, type: type, def: SDRS.length === 0 });
  saveSdrs(); renderSdrModal(); render();
  $('sdrUrl').value = ''; $('sdrName').value = '';
  msg.textContent = t('sdrAdded', { n: name, t: SDR_TYPES[type] || type });
  msg.classList.add('ok');
}

/* ---------- 时钟 ---------- */
function tickClock() {
  var now = Date.now(), f = fmt(state.tz);
  $('localClock').textContent = f.sec.format(now);
  $('localDate').textContent = f.dateLong.format(now);
  $('tzOffset').textContent = f.off;
  $('utcClock').textContent = pad2(new Date(now).getUTCHours()) + ':' + pad2(new Date(now).getUTCMinutes());
}

/* ---------- 控件构建 ---------- */
function buildChips() {
  var counts = {};
  E.forEach(function (it) {
    if (it.drm) counts.drm = (counts.drm || 0) + 1;
    counts[it.band] = (counts[it.band] || 0) + 1;
  });
  var defs = [['all', t('all')], ['SW', bandName('SW')], ['MW', bandName('MW')], ['LW', bandName('LW')],
    ['drm', t('drm')], ['time', bandName('time')], ['FM', bandName('FM')]];
  $('bandChips').innerHTML = defs.map(function (d) {
    return '<button class="chip' + (state.band === d[0] ? ' active' : '') + '" data-band="' + d[0] + '">' +
      d[1] + (counts[d[0]] ? '<span class="cnt">' + counts[d[0]] + '</span>' : '') + '</button>';
  }).join('');
  document.querySelectorAll('#bandChips .chip').forEach(function (el) {
    el.onclick = function () { state.band = el.getAttribute('data-band'); buildChips(); render(); };
  });

  var sdefs = [['all', t('allSrc')], ['aoki', srcName('aoki')], ['fm', srcName('fm')]];
  $('srcChips').innerHTML = '<span class="lbl">' + t('srcLbl') + '</span>' + sdefs.map(function (d) {
    return '<button class="chip' + (state.src === d[0] ? ' active' : '') + '" data-src="' + d[0] + '">' + d[1] + '</button>';
  }).join('');
  document.querySelectorAll('#srcChips .chip').forEach(function (el) {
    el.onclick = function () { state.src = el.getAttribute('data-src'); buildChips(); render(); };
  });
}

function buildLangs() {
  var cnt = {};
  var fk = UI === 'en' ? 'lge' : 'lgn';
  var dflt = UI === 'en' ? 'Unspecified' : '未指定';
  E.forEach(function (it) { var l = it[fk] || dflt; cnt[l] = (cnt[l] || 0) + 1; });
  var keys = Object.keys(cnt).filter(function (k) { return cnt[k] >= 8; })
    .sort(function (a, b) { return cnt[b] - cnt[a]; });
  $('langSelect').innerHTML = '<option value="all">' + t('allLang') + '</option>' + keys.map(function (k) {
    return '<option value="' + esc(k) + '">' + esc(k) + ' (' + cnt[k] + ')</option>';
  }).join('');
  $('langSelect').onchange = function () { state.lang = this.value; render(); };
}

function validTz(tz) {
  try { new Intl.DateTimeFormat('en-US', { timeZone: tz }).format(0); return true; } catch (e) { return false; }
}

function ensureCustomOption(label) {
  var sel = $('citySelect'), val = 'custom|' + state.tz + '|' + label;
  var opt = sel.querySelector('option[data-custom]');
  if (!opt) {
    opt = document.createElement('option');
    opt.setAttribute('data-custom', '1');
    sel.appendChild(opt);
  }
  opt.value = val;
  opt.textContent = '✏️ ' + dispLabel(label) + t('customMark');
  sel.value = val;
}

function setCustomLocation(tz, label) {
  if (!validTz(tz)) return false;
  state.tz = tz; state.cityLabel = label;
  try { localStorage.setItem('rg_city', 'custom|' + tz + '|' + label); } catch (e) {}
  ensureCustomOption(label);
  return true;
}

function buildTzList() {
  var dl = $('tzList'), frag = document.createDocumentFragment();
  CITIES.forEach(function (c) {
    var o = document.createElement('option'); o.value = c.n; frag.appendChild(o);
  });
  var zones = [];
  try { zones = Intl.supportedValuesOf ? Intl.supportedValuesOf('timeZone') : []; } catch (e) {}
  zones.forEach(function (z) {
    var o = document.createElement('option'); o.value = z; frag.appendChild(o);
  });
  for (var h = -12; h <= 14; h++) {
    var o = document.createElement('option'); o.value = h === 0 ? 'UTC' : 'UTC' + (h > 0 ? '+' : '') + h; frag.appendChild(o);
  }
  dl.appendChild(frag);
}

function applyCustomInput() {
  var inp = $('customCity'), v = inp.value.trim();
  if (!v) return;
  var lv = v.toLowerCase();
  var hit = CITIES.filter(function (c) {
    return c.n.toLowerCase().indexOf(lv) >= 0 || c.tz.toLowerCase() === lv;
  })[0];
  if (hit) { inp.classList.remove('err'); inp.value = ''; setCustomLocation(hit.tz, hit.n); render(); return; }
  if (validTz(v)) {
    inp.classList.remove('err'); inp.value = '';
    setCustomLocation(v, v.split('/').pop().replace(/_/g, ' '));
    render(); return;
  }
  var m = /^utc\s*([+-]?\d{1,2})$/i.exec(v);
  if (m) {
    var off = parseInt(m[1], 10);
    var z = off === 0 ? 'UTC' : 'Etc/GMT' + (off > 0 ? '-' + off : '+' + (-off));
    if (validTz(z)) {
      inp.classList.remove('err'); inp.value = '';
      setCustomLocation(z, off === 0 ? 'UTC' : 'UTC' + (off > 0 ? '+' : '') + off);
      render(); return;
    }
  }
  inp.classList.add('err');
  setTimeout(function () { inp.classList.remove('err'); }, 1800);
}

function buildCities() {
  var autoTz;
  try { autoTz = Intl.DateTimeFormat().resolvedOptions().timeZone; } catch (e) { autoTz = 'UTC'; }
  var opts = ['<option value="auto|跟随本机时区 (自动)">' + t('auto') + '</option>'];
  CITIES.forEach(function (c) {
    opts.push('<option value="' + c.tz + '|' + esc(c.n) + '">' + flag(c.c) + ' ' + esc(dispLabel(c.n)) + '</option>');
  });
  if (autoTz && !CITIES.some(function (c) { return c.tz === autoTz; })) {
    opts.push('<option value="' + autoTz + '|本机时区 ' + autoTz + '">' + t('deviceTz') + ' ' + esc(autoTz) + '</option>');
  }
  $('citySelect').innerHTML = opts.join('');
  var saved = null;
  try { saved = localStorage.getItem('rg_city'); } catch (e) {}
  if (saved && saved.indexOf('|') > 0) {
    if (saved.indexOf('custom|') === 0) {
      var p = saved.split('|');
      state.tz = p[1]; state.cityLabel = p[2];
      ensureCustomOption(p[2]);
    } else { $('citySelect').value = saved; }
  }
  applyCity();
  buildTzList();
  $('citySelect').onchange = function () {
    $('customCity').value = '';
    try { localStorage.setItem('rg_city', this.value); } catch (e) {}
    applyCity(); render();
  };
  $('customCity').addEventListener('change', applyCustomInput);
  $('customCity').addEventListener('keydown', function (e) {
    if (e.key === 'Enter') { e.preventDefault(); applyCustomInput(); }
  });
}

function applyCity() {
  var v = $('citySelect').value.split('|');
  if (v[0] === 'custom') {
    state.tz = v[1]; state.cityLabel = v[2] || v[1];
  } else if (v[0] === 'auto') {
    var tz;
    try { tz = Intl.DateTimeFormat().resolvedOptions().timeZone; } catch (e) { tz = 'UTC'; }
    state.tz = tz;
    state.cityLabel = v[1].indexOf('本机') === 0 ? v[1] : '本机时区';
    if (v[1] === '跟随本机时区 (自动)') {
      var hit = CITIES.filter(function (c) { return c.tz === tz; })[0];
      state.cityLabel = hit ? hit.n : (UI === 'en' ? 'Device timezone ' + tz : '本机时区 ' + tz);
    }
  } else {
    state.tz = v[0]; state.cityLabel = v[1];
  }
  // 接收机位置坐标（用于卡片上计算距离/方位）
  var me = CITIES.filter(function (c) { return c.n === state.cityLabel && c.x; })[0];
  state.cityCoord = me ? me.x : null;
  var rc = $('rxCoord');
  if (rc) rc.textContent = state.cityCoord ? fmtLatLon(state.cityCoord) : '—';
}

/* ---------- 界面语言应用 ---------- */
function applyI18n() {
  document.documentElement.lang = UI === 'en' ? 'en' : 'zh-CN';
  document.title = t('docTitle');
  $('h1text').textContent = t('h1');
  $('h1sub').textContent = t('h1sub');
  $('tagline').textContent = t('tagline');
  $('locLabel').textContent = t('myLoc');
  $('customCity').placeholder = t('customPh');
  $('search').placeholder = t('searchPh');
  var sb = $('sdrBtn');
  if (sb && sb.childNodes[0]) sb.childNodes[0].nodeValue = t('sdr') + ' ';
  $('viewSeg').querySelectorAll('button').forEach(function (b) {
    b.textContent = { now: t('onair'), next: t('next'), all: t('allv') }[b.getAttribute('data-view')];
  });
  $('foot1').innerHTML = t('f1');
  $('foot2').innerHTML = t('f2');
  $('foot3').innerHTML = t('f3', { g: META.generated, u: META.updated });
  $('sdrBtn').title = t('sdrT');
  $('sdrHint').innerHTML = t('sdrHint');
  $('sdrUrl').placeholder = t('sdrUrlPh');
  $('sdrName').placeholder = t('sdrNamePh');
  $('sdrType').options[0].textContent = t('sdrAuto');
  $('sdrAdd').textContent = t('sdrAdd');
  var rb = $('sdrReset');
  if (rb) { rb.textContent = t('sdrReset'); rb.title = t('sdrResetT'); }
  $('langBtn').textContent = UI === 'en' ? '中文' : 'EN';
}

/* ---------- 启动 ---------- */
function init() {
  applyI18n();
  buildCities(); buildChips(); buildLangs();
  // 语言切换
  $('langBtn').onclick = function () {
    UI = UI === 'en' ? 'zh' : 'en';
    try { localStorage.setItem('rg_lang', UI); } catch (e) {}
    fmtCache = {};
    applyI18n(); buildCities(); buildChips(); buildLangs(); render(); tickClock();
  };
  // SDR 联动
  $('sdrBtn').onclick = function () { $('sdrModal').classList.remove('hidden'); renderSdrModal(); };
  $('sdrClose').onclick = function () { $('sdrModal').classList.add('hidden'); };
  $('sdrModal').addEventListener('click', function (e) { if (e.target === this) this.classList.add('hidden'); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') { $('sdrModal').classList.add('hidden'); closeSdrMenus(); } });
  $('sdrAdd').onclick = addSdr;
  $('sdrReset').onclick = resetSdrsToConfig;
  $('sdrUrl').addEventListener('keydown', function (e) { if (e.key === 'Enter') addSdr(); });
  // 卡片"备用接收机"菜单：点按钮开合，点外面关闭，菜单内是真实链接直接跳转
  document.addEventListener('click', function (e) {
    var tog = e.target && e.target.closest ? e.target.closest('[data-sdrmenu]') : null;
    if (tog) {
      var menu = tog.parentElement.querySelector('.sdrmenu');
      if (menu) {
        var wasHidden = menu.classList.contains('hidden');
        closeSdrMenus();
        if (wasHidden) menu.classList.remove('hidden');
      }
      return;
    }
    if (!e.target.closest || !e.target.closest('.sdrmenu-wrap')) closeSdrMenus();
  });
  $('sdrCount').textContent = SDRS.length ? '×' + SDRS.length : '';
  $('viewSeg').querySelectorAll('button').forEach(function (el) {
    el.onclick = function () {
      $('viewSeg').querySelectorAll('button').forEach(function (b) { b.classList.remove('active'); });
      el.classList.add('active');
      state.view = el.getAttribute('data-view');
      state.limitAll = CAP_ALL;
      render();
    };
  });
  var st;
  st = $('search'); st.oninput = function () { state.q = this.value.trim(); render(); };
  tickClock(); render();
  setInterval(tickClock, 1000);
  setInterval(render, 20000);
}
init();

})();
