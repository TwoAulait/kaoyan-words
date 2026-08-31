'use strict';

/* 考研英语背单词（安卓端）— 逻辑与 Windows 版 app.py 保持一致：
   - progress = 当前单词序号（0 起），浏览即自动保存
   - 认识：生词本有则删去。都跳下一个
   - 不认识：生词本有则跳过，否则记入生词本。都跳下一个
   - 无「已认识」记录、无时间戳
   - 同步：坚果云是唯一同步方式（设置页填写），与电脑共用同一 WebDAV 文件，全手动——
     打开 app 时弹窗询问是否从云端获取；主页面「☁ 上传」「☁ 从云端获取」按钮手动同步并显示上次同步时间；
     退出时若有未上传改动，由 Java 原生弹窗三选（先上传再退出 / 直接退出 / 继续学习）。
     其余操作不联网，每次操作时自动把进度和生词做一份本地备份（保留最近 10 条）。
   - 备份：JSON {"version":1,"progress":N,"unknown":[{word,phonetic,meaning}...]}
   - 本地备份：每次操作时把当前数据存为本地备份（不联网），保留最近 10 条，设置页可恢复
   - 切后台/退出时的本地备份、退出确认弹窗与「先上传再退出」的上传都由 Java 原生
     在 onPause/onBackPressed 里用 JS 平时推来的最新快照直接执行（WebView 后台冻结/
     删进程时 JS 不执行，必须走原生）。
 */

var WORDS = window.WORDS || [];
var TOTAL = WORDS.length;
var LS_PROG = 'kaoyan_progress';
var LS_UNK = 'kaoyan_unknown';
var LS_STORAGE_VERSION = 'kaoyan_storage_version';
var LS_CLOUD_URL = 'kaoyan_cloud_url';
var LS_CLOUD_USER = 'kaoyan_cloud_user';
var LS_CLOUD_PASS = 'kaoyan_cloud_pass';
var LS_BACKUPS = 'kaoyan_local_backups';
var LS_LAST_SYNC = 'kaoyan_last_sync';       // 上次成功同步（获取或上传）时间，显示在主页面同步按钮旁
var LS_LAST_SYNCED = 'kaoyan_last_synced';   // 基线：上次与云端一致的 {progress, unknown}，用于判断是否有未上传改动
var STORAGE_VERSION = 2;   // 2 = 去掉已认识、改为「手机覆盖/镜像」同步规则后的结构
var KEY_SEP = '';    // word+注音 拼接分隔符，避免单词本身含分隔符

var state = {
  current: 0,
  meaningShown: false,
  unknown: [],           // [{word, phonetic, meaning}]
  unkSet: new Set()      // word+KEY_SEP+phonetic
};

function el(id) { return document.getElementById(id); }
function clamp(n, lo, hi) { return Math.max(lo, Math.min(n, hi)); }
function wordKey(row) { return row.word + KEY_SEP + (row.phonetic || ''); }

// ---------------------------------------------------------- 本地存储
function loadNum(key, def) {
  try { var v = parseInt(localStorage.getItem(key), 10); return isNaN(v) ? def : v; } catch (e) { return def; }
}
function loadUnknown() {
  try { var a = JSON.parse(localStorage.getItem(LS_UNK) || '[]'); return Array.isArray(a) ? a : []; }
  catch (e) { return []; }
}
var _suppressAuto = false;   // 从云端/备份拉取写回本地时不触发自动本地备份
function saveProgress() {
  try { localStorage.setItem(LS_PROG, String(state.current)); } catch (e) {}
  pushStateToNative();   // 每次保存都把最新快照推给原生，供切后台/退出时原生直接备份+上传
  if (!_suppressAuto) addLocalBackup();   // 每次操作只做本地备份，不联网
}
function saveUnknown() {
  try { localStorage.setItem(LS_UNK, JSON.stringify(state.unknown)); } catch (e) {}
  pushStateToNative();
  if (!_suppressAuto) addLocalBackup();
}
// 把最新进度/生词推给原生（volatile 快照）；onPause/退出时 Java 用该快照直接写文件、发请求。
// 同时把「是否有未上传改动」也推给原生，供退出弹窗判断是否提示。
function pushStateToNative() {
  if (window.AndroidBridge && window.AndroidBridge.updateState) {
    try { window.AndroidBridge.updateState(String(state.current), JSON.stringify(state.unknown)); } catch (e) {}
  }
  pushDirtyToNative();
}
// 把云端配置推给原生，让 Java 在切后台/退出时能直接发上传请求。
function syncConfigToNative() {
  if (window.AndroidBridge && window.AndroidBridge.setCloudConfig) {
    try {
      window.AndroidBridge.setCloudConfig(
        (localStorage.getItem(LS_CLOUD_URL) || '').trim(),
        (localStorage.getItem(LS_CLOUD_USER) || '').trim(),
        (localStorage.getItem(LS_CLOUD_PASS) || '').trim());
    } catch (e) {}
  }
}
function rebuildUnkSet() {
  state.unkSet = new Set(state.unknown.map(function (r) { return wordKey(r); }));
}
function hasUnk(key) { return state.unkSet.has(key); }

// 一次性初始化（非常态化规则）：本版本首次运行清空手机端旧数据，
// 之后第一次「打开电脑端同步」就会用电脑端当前数据初始化手机端。
// 返回 true 表示本次是首次安装/升级（刚清空过数据），用于弹出首次使用提示。
function migrateStorage() {
  try {
    if (parseInt(localStorage.getItem(LS_STORAGE_VERSION), 10) === STORAGE_VERSION) return false;
    localStorage.removeItem(LS_PROG);
    localStorage.removeItem(LS_UNK);
    localStorage.removeItem('kaoyan_known');   // 清理旧版「已认识」数据
    localStorage.setItem(LS_STORAGE_VERSION, String(STORAGE_VERSION));
    return true;
  } catch (e) { return false; }
}

// 首次使用提示：先操作电脑端完成一次同步，不要在手机上学习，避免空数据覆盖电脑端。
function showFirstRunNotice() {
  el('firstRunOverlay').style.display = 'flex';
}
function dismissFirstRunNotice() {
  el('firstRunOverlay').style.display = 'none';
}

// ---------------------------------------------------------- 渲染
function updateHeaderUnknown() {
  el('unkCount').textContent = '生词本 ' + state.unknown.length + ' 词';
  if (el('screenUnknown').style.display !== 'none') renderUnknownList();
}
function currentKey() {
  var w = WORDS[state.current];
  return wordKey({ word: w[0], phonetic: w[1] || '' });
}
function renderWord() {
  var w = WORDS[state.current];
  el('word').textContent = w[0];
  el('phonetic').textContent = w[1] || '';
  el('meaning').textContent = w[2] || '';
  el('pos').textContent = '第 ' + (state.current + 1) + ' / ' + TOTAL + ' 个单词';
  el('barFill').style.width = (((state.current + 1) / TOTAL) * 100) + '%';
  el('badge').style.display = hasUnk(currentKey()) ? 'inline-block' : 'none';
  el('prevBtn').disabled = state.current <= 0;
  el('nextBtn').disabled = state.current >= TOTAL - 1;
  setMeaningShown(false);               // 翻页后默认隐藏释义（与 Windows 一致）
  updateHeaderUnknown();
}
function setMeaningShown(shown) {
  state.meaningShown = shown;
  el('meaning').style.display = shown ? 'block' : 'none';
  el('viewLabel').textContent = shown ? '收起' : '查看';
}
function setStatus(text, kind) {
  var s = el('status');
  s.textContent = text;
  s.className = 'status' + (kind ? ' ' + kind : '');
  if (s._t) clearTimeout(s._t);
  s._t = setTimeout(function () { s.textContent = ''; }, 2000);
}
function setSyncStatus(text, isErr) {
  var s = el('syncStatus');
  s.textContent = text;
  s.className = 'status' + (isErr ? ' err' : ' ok');
}
function setCloudStatus(text, isErr) {
  var s = el('cloudStatus');
  if (!s) return;
  s.textContent = text;
  s.className = 'status' + (isErr ? ' err' : ' ok');
}
function toast(msg) {
  if (window.AndroidBridge && window.AndroidBridge.toast) {
    try { window.AndroidBridge.toast(String(msg)); } catch (e) {}
  }
}
function showScreen(name) {
  var map = { study: 'screenStudy', unknown: 'screenUnknown', settings: 'screenSettings' };
  for (var k in map) el(map[k]).style.display = (k === name) ? 'block' : 'none';
  if (name === 'settings') {
    el('cloudUrl').value = localStorage.getItem(LS_CLOUD_URL) || '';
    el('cloudUser').value = localStorage.getItem(LS_CLOUD_USER) || '';
    el('cloudPass').value = localStorage.getItem(LS_CLOUD_PASS) || '';
    refreshCloudStatus();
    renderLocalBackups();
  }
  if (name === 'unknown') renderUnknownList();
}
function renderUnknownList() {
  var list = el('unkList'), empty = el('unkEmpty');
  list.innerHTML = '';
  if (!state.unknown.length) { empty.style.display = 'block'; return; }
  empty.style.display = 'none';
  state.unknown.forEach(function (r, i) {
    var item = document.createElement('div');
    item.className = 'unk-item';
    var del = document.createElement('button');
    del.className = 'del';
    del.textContent = '删除';
    del.onclick = function () { removeUnknownAt(i); };
    var w = document.createElement('div'); w.className = 'w'; w.textContent = r.word;
    var p = document.createElement('div'); p.className = 'p'; p.textContent = r.phonetic || '';
    var m = document.createElement('div'); m.className = 'm'; m.textContent = r.meaning || '';
    item.appendChild(del);
    item.appendChild(w);
    item.appendChild(p);
    item.appendChild(m);
    list.appendChild(item);
  });
}
function removeUnknownAt(i) {
  var r = state.unknown[i];
  if (!r) return;
  state.unknown.splice(i, 1);
  rebuildUnkSet();
  saveUnknown();
  renderUnknownList();
  renderWord();
}
function clearAllUnknown() {
  if (!state.unknown.length) { toast('生词本已是空的'); return; }
  askConfirm('确定清空生词本（' + state.unknown.length + ' 个词）？', function () {
    state.unknown = [];
    rebuildUnkSet();
    saveUnknown();
    renderUnknownList();
    renderWord();
    setStatus('生词本已清空');
  });
}

// 应用内确认弹层（WebView 不支持原生 confirm，改用它，桌面浏览器同样可用）
var _confirmCb = null;
function askConfirm(text, cb) {
  el('confirmText').textContent = text;
  _confirmCb = cb;
  el('confirmOverlay').style.display = 'flex';
}
function confirmOk() {
  var cb = _confirmCb;
  _confirmCb = null;
  el('confirmOverlay').style.display = 'none';
  if (cb) cb();
}
function confirmCancel() {
  _confirmCb = null;
  el('confirmOverlay').style.display = 'none';
}

// ---------------------------------------------------------- 学习交互
function goNext() { if (state.current < TOTAL - 1) showWord(state.current + 1); }
function goPrev() { if (state.current > 0) showWord(state.current - 1); }
function showWord(pos) {
  state.current = clamp(pos, 0, TOTAL - 1);
  renderWord();
  saveProgress();
}
function toggleMeaning() { setMeaningShown(!state.meaningShown); }

function markKnow() {
  var w = WORDS[state.current];
  var key = wordKey({ word: w[0], phonetic: w[1] || '' });
  var msg = '已标记认识';
  if (hasUnk(key)) {
    state.unknown = state.unknown.filter(function (r) { return wordKey(r) !== key; });
    rebuildUnkSet();
    msg = '已从生词本删除';
  }
  saveUnknown();
  updateHeaderUnknown();
  setStatus(msg);
  goNext();
}

function markUnknown() {
  var w = WORDS[state.current];
  var key = wordKey({ word: w[0], phonetic: w[1] || '' });
  var msg = '该词已在生词本';
  if (!hasUnk(key)) {
    state.unknown.push({ word: w[0], phonetic: w[1] || '', meaning: w[2] || '' });
    rebuildUnkSet();
    msg = '已记入生词本';
  }
  saveUnknown();
  updateHeaderUnknown();
  setStatus(msg);
  goNext();
}

// ---------------------------------------------------------- 坚果云网盘同步（唯一同步方式）
// 规则（全手动）：打开 app 时询问是否从云端获取；主页面「☁ 上传」「☁ 从云端获取」手动同步；
// 退出时若有未上传改动，原生弹窗三选（先上传再退出/直接退出/继续学习）。其余操作不联网（仅本地备份，保留最近 10 条）。
// 云端只放一份 JSON（备份格式），与电脑端共用同一 WebDAV 地址。
// 网络请求优先走安卓原生桥（AndroidBridge.netCloudGet/Put，HttpURLConnection，绕开 CORS）；
// 桌面浏览器调试时回退为 fetch（坚果云若未开 CORS 会失败，手机端不受影响）。
var _cloudSeq = 0, _cloudPending = {};
function isCloudConfigured() {
  return !!(localStorage.getItem(LS_CLOUD_URL) && localStorage.getItem(LS_CLOUD_USER) && localStorage.getItem(LS_CLOUD_PASS));
}
function refreshCloudStatus() {
  setCloudStatus(isCloudConfigured() ? '坚果云同步已开启（手动：主页面「☁ 上传」「☁ 从云端获取」同步，打开时询问是否获取）' : '坚果云同步未配置', false);
}

// —— 手动同步：主页面「☁ 上传」「☁ 从云端获取」按钮、上次同步时间、脏标记
function stateDesc(progress, unknownCount) {
  // 与界面一致的进度描述：第 X / 5493 词，生词 N 个（progress 为 0 基）
  return '第 ' + (progress + 1) + '/' + TOTAL + ' 词，生词 ' + unknownCount + ' 个';
}
function _stateFingerprint() {
  return JSON.stringify({ progress: state.current, unknown: state.unknown });
}
function isDirty() {
  // 本地数据相对上次成功同步（获取或上传）是否有未上传的改动；从未同步过也算有。
  if (!isCloudConfigured()) return false;
  var base = null;
  try { base = JSON.parse(localStorage.getItem(LS_LAST_SYNCED) || 'null'); } catch (e) { base = null; }
  if (!base) return true;
  return _stateFingerprint() !== JSON.stringify({ progress: base.progress, unknown: base.unknown || [] });
}
function setBaseline() {
  // 成功从云端获取或上传后：记录「与云端一致的基线」，并刷新原生的脏标记
  try { localStorage.setItem(LS_LAST_SYNCED, JSON.stringify({ progress: state.current, unknown: state.unknown })); } catch (e) {}
  pushDirtyToNative();
}
function setLastSync() {
  try { localStorage.setItem(LS_LAST_SYNC, formatBackupTime(Date.now())); } catch (e) {}
}
function getLastSync() { return localStorage.getItem(LS_LAST_SYNC) || ''; }
function pushDirtyToNative() {
  if (window.AndroidBridge && window.AndroidBridge.setDirty) {
    try { window.AndroidBridge.setDirty(isDirty()); } catch (e) {}
  }
  renderSyncButton();   // 脏标记变化时同步刷新主页面红字警示
}
function renderSyncButton() {
  var b = el('syncBtn'), l = el('lastSyncLabel');
  if (!b || !l) return;
  if (!isCloudConfigured()) {
    l.textContent = '未配置坚果云同步（设置页 ⚙ 可配置）';
    l.className = 'last-sync';
    return;
  }
  var ls = getLastSync();
  if (isDirty()) {
    // 有改动未上传到云端 → 红字警示
    l.textContent = '⚠ 有改动未上传云端' + (ls ? '（上次同步：' + ls + '）' : '');
    l.className = 'last-sync dirty';
  } else {
    l.textContent = ls ? '上次同步：' + ls : '尚未同步过云端';
    l.className = 'last-sync';
  }
}
function cloudHttp(method, body, done, cfg) {
  // cfg 可选：测试连接时传当前输入框里的值（未保存也能测）
  var url = cfg ? cfg.url : (localStorage.getItem(LS_CLOUD_URL) || '').trim();
  var user = cfg ? cfg.user : (localStorage.getItem(LS_CLOUD_USER) || '').trim();
  var pass = cfg ? cfg.pass : (localStorage.getItem(LS_CLOUD_PASS) || '').trim();
  if (!url || !user || !pass) { done({ ok: false, err: '未填写网盘配置' }); return; }
  if (window.AndroidBridge && window.AndroidBridge.cloudGet) {
    var id = 'c' + (++_cloudSeq);
    _cloudPending[id] = done;
    if (method === 'PUT') window.AndroidBridge.cloudPut(url, user, pass, body, id);
    else window.AndroidBridge.cloudGet(url, user, pass, id);
    return;
  }
  try {
    var h = new Headers({ 'Authorization': 'Basic ' + btoa(user + ':' + pass) });
    fetch(url, { method: method, headers: h, body: method === 'PUT' ? body : undefined })
      .then(function (r) { return r.text().then(function (t) { done({ ok: true, status: r.status, body: t }); }); })
      .catch(function (e) { done({ ok: false, err: String(e) }); });
  } catch (e) { done({ ok: false, err: String(e) }); }
}
// 主页面「从云端获取」：先读云端数据二次确认，确认后应用
function fetchNow() {
  if (!isCloudConfigured()) { setStatus('未配置坚果云同步，请到设置页 ⚙ 填写', 'err'); toast('未配置坚果云同步'); return; }
  addLocalBackup();            // 获取前先本地备份，覆盖错了也能用备份找回
  setStatus('正在读取云端数据…', 'ok');
  cloudHttp('GET', null, function (r) {
    if (!r.ok) { setStatus('云端获取失败：' + (r.err || ''), 'err'); toast('云端获取失败，请检查网络'); return; }
    if (r.status === 404) { setStatus('云端暂无数据，保持本地进度和生词', 'ok'); toast('云端暂无数据，保持本地进度和生词'); return; }
    try {
      var d = JSON.parse(r.body);
      var prog = parseInt(d.progress, 10);
      if (isNaN(prog)) throw new Error('进度无效');
      var rows = Array.isArray(d.unknown) ? d.unknown : [];
      if (prog <= 0 && !rows.length) { setStatus('云端为空，保持本地进度和生词', 'ok'); toast('云端为空，保持本地进度和生词'); return; }
      // 二次确认：显示云端与本机数字，确认才覆盖本机
      askConfirm('云端当前：' + stateDesc(prog, rows.length) +
                 '\n本机当前：' + stateDesc(state.current, state.unknown.length) +
                 '\n\n确认用云端数据覆盖本机吗？', function () {
        _suppressAuto = true;
        state.current = clamp(prog, 0, TOTAL - 1);
        state.unknown = rows;
        rebuildUnkSet();
        saveProgress();
        saveUnknown();
        _suppressAuto = false;
        renderWord();
        setLastSync();      // 获取成功 → 更新上次同步时间
        setBaseline();      // 更新基线，本机与云端一致
        renderSyncButton();
        setStatus('已从云端获取：' + stateDesc(state.current, state.unknown.length), 'ok');
        toast('已从云端获取最新进度和生词');
      });
    } catch (e) { setStatus('云端数据无效：' + e.message, 'err'); toast('云端数据无效'); }
  });
}
// 测试连接：用当前输入框的值连一次坚果云，报告云端当前数据
function testCloudConnection() {
  var url = (el('cloudUrl').value || '').trim();
  var user = (el('cloudUser').value || '').trim();
  var pass = (el('cloudPass').value || '').trim();
  if (!url || !user || !pass) { setCloudStatus('请先填写完整的地址/用户名/应用密码', true); toast('配置不完整'); return; }
  setCloudStatus('正在测试连接…', false);
  cloudHttp('GET', null, function (r) {
    if (!r.ok) { setCloudStatus('连接失败：' + (r.err || ''), true); toast('连接失败'); return; }
    if (r.status === 404) { setCloudStatus('连接成功：云端还没有数据文件', false); toast('连接成功'); return; }
    if (r.status === 200) {
      try {
        var d = JSON.parse(r.body);
        var prog = parseInt(d.progress, 10);
        var rows = Array.isArray(d.unknown) ? d.unknown : [];
        var desc = (!isNaN(prog) && (prog > 0 || rows.length)) ? stateDesc(prog, rows.length) : '暂无数据';
        setCloudStatus('连接成功：云端当前 ' + desc, false);
        toast('连接成功');
      } catch (e) { setCloudStatus('连接成功（云端数据格式异常）', false); toast('连接成功'); }
    } else setCloudStatus('连接失败：HTTP ' + r.status, true);
  }, { url: url, user: user, pass: pass });
}
// 上传前：先读云端当前数据，弹「上传确认」显示云端与本机数字；确认后调 onUpload()
function confirmUpload(onUpload) {
  addLocalBackup();            // 本地备份同步落盘，不依赖网络回调
  cloudHttp('GET', null, function (r) {
    var cloudDesc;
    if (!r.ok) cloudDesc = '云端读取失败，未知';
    else if (r.status === 404) cloudDesc = '暂无数据';
    else if (r.status === 200) {
      try {
        var d = JSON.parse(r.body);
        var prog = parseInt(d.progress, 10);
        var rows = Array.isArray(d.unknown) ? d.unknown : [];
        cloudDesc = (!isNaN(prog) && (prog > 0 || rows.length)) ? stateDesc(prog, rows.length) : '暂无数据';
      } catch (e) { cloudDesc = '云端数据无效'; }
    } else cloudDesc = '云端读取失败（HTTP ' + r.status + '）';
    askConfirm('云端当前：' + cloudDesc +
               '\n本机当前：' + stateDesc(state.current, state.unknown.length) +
               '\n\n确认把本机数据上传覆盖云端吗？', onUpload);
  });
}
function syncUploadNow() {
  // 主页面「☁ 上传」：先读云端二次确认，确认后上传（手动，不下载）
  if (!isCloudConfigured()) { setStatus('未配置坚果云同步，请到设置页 ⚙ 填写', 'err'); toast('未配置坚果云同步'); return; }
  confirmUpload(function () {
    setStatus('正在上传到云端…', 'ok');
    cloudHttp('PUT', buildBackupJson(), function (r) {
      if (!r.ok) { setStatus('云端上传失败：' + (r.err || ''), 'err'); toast('云端上传失败'); return; }
      if (r.status >= 200 && r.status < 300) {
        setLastSync(); setBaseline(); renderSyncButton();
        setStatus('已上传云端：' + stateDesc(state.current, state.unknown.length) + '（' + getLastSync() + '）', 'ok');
        toast('已上传到云端');
      }
      else { setStatus('云端上传失败：HTTP ' + r.status, 'err'); toast('云端上传失败'); }
    });
  });
}
function saveCloudConfig() {
  try {
    localStorage.setItem(LS_CLOUD_URL, el('cloudUrl').value.trim());
    localStorage.setItem(LS_CLOUD_USER, el('cloudUser').value.trim());
    localStorage.setItem(LS_CLOUD_PASS, el('cloudPass').value.trim());
  } catch (e) {}
  syncConfigToNative();
  pushDirtyToNative();
  renderSyncButton();
  setCloudStatus('配置已保存' + (isCloudConfigured() ? '：可点主页面「☁ 上传」上传、「☁ 从云端获取」拉取' : '，请填写完整'), false);
}
function clearCloudConfig() {
  try {
    localStorage.removeItem(LS_CLOUD_URL);
    localStorage.removeItem(LS_CLOUD_USER);
    localStorage.removeItem(LS_CLOUD_PASS);
  } catch (e) {}
  el('cloudUrl').value = ''; el('cloudUser').value = ''; el('cloudPass').value = '';
  syncConfigToNative();
  pushDirtyToNative();
  renderSyncButton();
  setCloudStatus('已清除网盘配置', false);
}

// ---------------------------------------------------------- 本地备份（自动保留最近 10 条）
// 每次操作/切后台/退出时把当前数据存为本地备份，设置页可一键恢复（与云端互为双保险）。
// 优先走原生：备份由 Java 直接写文件（WebView 销毁/进程被杀后 JS 不执行，原生文件持久且可靠）；
// 桌面浏览器调试时回退为 localStorage。
function loadLocalBackups() {
  try { var a = JSON.parse(localStorage.getItem(LS_BACKUPS) || '[]'); return Array.isArray(a) ? a : []; }
  catch (e) { return []; }
}
function addLocalBackup() {
  if (window.AndroidBridge && window.AndroidBridge.saveLocalBackup) {
    try { window.AndroidBridge.saveLocalBackup(buildBackupJson()); return; } catch (e) {}
  }
  try {
    var list = loadLocalBackups();
    list.push({ t: Date.now(), data: { progress: state.current, unknown: state.unknown } });
    while (list.length > 10) list.shift();
    localStorage.setItem(LS_BACKUPS, JSON.stringify(list));
    renderLocalBackups();
  } catch (e) {}
}
function formatBackupTime(t) {
  var d = new Date(t);
  function p(n) { return (n < 10 ? '0' : '') + n; }
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' '
       + p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
}
// 原生备份列表缓存（newest-first，与 Java 端 backupFiles() 一致）；null = 尚未加载
var _nativeBackups = null;
function renderLocalBackups() {
  var list = el('backupList');
  if (!list) return;
  if (window.AndroidBridge && window.AndroidBridge.listLocalBackups) {
    // 每次进设置页都重新向原生读取，保证列表最新
    _nativeBackups = null;
    list.innerHTML = '<div class="empty-tip" style="margin-top:10px;font-size:12px">加载中…</div>';
    window.AndroidBridge.listLocalBackups();
    return;
  }
  renderFallbackBackups();
}
function renderBackupList() {
  var list = el('backupList');
  if (!list) return;
  list.innerHTML = '';
  if (!_nativeBackups || !_nativeBackups.length) {
    list.innerHTML = '<div class="empty-tip" style="margin-top:10px;font-size:12px">还没有本地备份，操作并同步后会在这里出现</div>';
    return;
  }
  _nativeBackups.forEach(function (b, i) {
    var item = document.createElement('div');
    item.className = 'unk-item';
    var btn = document.createElement('button');
    btn.className = 'del';
    btn.textContent = '恢复';
    btn.onclick = function () { restoreNativeBackup(i); };
    var t = document.createElement('div'); t.className = 'w'; t.textContent = formatBackupTime(b.t);
    var prog = (typeof b.progress === 'number') ? (b.progress + 1) : 1;
    var m = document.createElement('div'); m.className = 'm';
    m.textContent = '进度 ' + prog + ' / ' + TOTAL + '，生词 ' + (b.unknown_total || 0) + ' 个';
    item.appendChild(btn);
    item.appendChild(t);
    item.appendChild(m);
    list.appendChild(item);
  });
}
function restoreNativeBackup(i) {
  if (!_nativeBackups || !_nativeBackups[i]) { setSyncStatus('备份不存在', true); return; }
  askConfirm('确定恢复到 ' + formatBackupTime(_nativeBackups[i].t) + ' 的备份？\n当前进度和生词本将被覆盖。', function () {
    if (window.AndroidBridge && window.AndroidBridge.readLocalBackup) {
      window.AndroidBridge.readLocalBackup(i);
    }
  });
}
// 桌面浏览器调试回退：localStorage 里存的就是 旧→新，这里倒序显示新在前
function renderFallbackBackups() {
  var list = el('backupList');
  if (!list) return;
  list.innerHTML = '';
  var arr = loadLocalBackups();
  if (!arr.length) {
    list.innerHTML = '<div class="empty-tip" style="margin-top:10px;font-size:12px">还没有本地备份，操作并同步后会在这里出现</div>';
    return;
  }
  for (var idx = arr.length - 1; idx >= 0; idx--) {
    (function (i) {
      var b = arr[i];
      var item = document.createElement('div');
      item.className = 'unk-item';
      var btn = document.createElement('button');
      btn.className = 'del';
      btn.textContent = '恢复';
      btn.onclick = function () { restoreLocalBackup(i); };
      var t = document.createElement('div'); t.className = 'w'; t.textContent = formatBackupTime(b.t);
      var prog = (b.data && typeof b.data.progress === 'number') ? (b.data.progress + 1) : 1;
      var unk = (b.data && Array.isArray(b.data.unknown)) ? b.data.unknown.length : 0;
      var m = document.createElement('div'); m.className = 'm';
      m.textContent = '进度 ' + prog + ' / ' + TOTAL + '，生词 ' + unk + ' 个';
      item.appendChild(btn);
      item.appendChild(t);
      item.appendChild(m);
      list.appendChild(item);
    })(idx);
  }
}
function restoreLocalBackup(i) {
  var b = loadLocalBackups()[i];
  if (!b || !b.data) { setSyncStatus('备份不存在', true); return; }
  askConfirm('确定恢复到 ' + formatBackupTime(b.t) + ' 的备份？\n当前进度和生词本将被覆盖。', function () {
    applyBackup(JSON.stringify({ progress: b.data.progress, unknown: Array.isArray(b.data.unknown) ? b.data.unknown : [] }));
  });
}

// ---------------------------------------------------------- 备份导出/导入
function buildBackupJson() {
  return JSON.stringify({ version: 1, progress: state.current, unknown: state.unknown }, null, 1);
}
function exportBackup() {
  if (window.AndroidBridge && window.AndroidBridge.exportBackup) {
    window.AndroidBridge.exportBackup(buildBackupJson());
  } else {
    // 桌面浏览器调试回退：直接下载文件
    try {
      var blob = new Blob([buildBackupJson()], { type: 'application/json' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = '考研英语背单词备份.json';
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
      setSyncStatus('已导出备份文件', false);
    } catch (e) { setSyncStatus('导出失败：' + e.message, true); }
  }
}
function importBackup() {
  if (window.AndroidBridge && window.AndroidBridge.importBackup) {
    window.AndroidBridge.importBackup();   // 结果经 onImportDone 回调返回
  } else {
    // 桌面浏览器调试回退：文件选择
    var inp = document.createElement('input');
    inp.type = 'file';
    inp.accept = '.json,application/json';
    inp.onchange = function () {
      var f = inp.files && inp.files[0];
      if (!f) return;
      var rd = new FileReader();
      rd.onload = function () { applyBackup(String(rd.result)); };
      rd.readAsText(f, 'utf-8');
    };
    inp.click();
  }
}
function applyBackup(jsonStr) {
  try {
    var d = JSON.parse(jsonStr);
    var prog = parseInt(d.progress, 10);
    if (isNaN(prog)) throw new Error('备份中进度无效');
    if (d.unknown != null && !Array.isArray(d.unknown)) throw new Error('备份中生词数据无效');
    state.unknown = Array.isArray(d.unknown) ? d.unknown : [];
    rebuildUnkSet();
    state.current = clamp(prog, 0, TOTAL - 1);
    saveProgress();
    saveUnknown();
    renderWord();
    setSyncStatus('备份已导入：进度 ' + (state.current + 1) + '/' + TOTAL + '，生词本 ' + state.unknown.length + ' 个', false);
    toast('备份导入成功');
  } catch (e) {
    setSyncStatus('备份导入失败：' + e.message, true);
    toast('导入失败');
  }
}

// 安卓桥回调（由 Java 端 evaluateJavascript 注入调用）
window.BridgeCallbacks = {
  onExportDone: function (msg) { setSyncStatus(msg || '备份已导出', false); toast('导出成功'); },
  onExportError: function (msg) { setSyncStatus(msg || '导出失败', true); toast('导出失败'); },
  onImportDone: function (jsonStr) { applyBackup(jsonStr); },
  onImportError: function (msg) { setSyncStatus(msg || '导入失败', true); toast('导入失败'); },
  // 坚果云 WebDAV 原生请求的回调：resultJson = {"ok":..,"status":..,"body":..} 或 {"ok":false,"err":..}
  onCloud: function (id, resultJson) {
    var done = _cloudPending[id];
    if (!done) return;
    delete _cloudPending[id];
    try { done(JSON.parse(resultJson)); } catch (e) { done({ ok: false, err: '解析失败' }); }
  },
  // 原生本地备份回调：onLocalBackups = 备份列表 JSON 字符串（newest-first），
  // onLocalBackupRead = 备份文件内容（JSON 字符串），失败/越界返回 "null"
  onLocalBackups: function (json) {
    try { _nativeBackups = JSON.parse(json) || []; } catch (e) { _nativeBackups = []; }
    renderBackupList();
  },
  onLocalBackupRead: function (json) {
    if (!json || json === 'null') { setSyncStatus('备份读取失败', true); return; }
    applyBackup(json);
  }
};

// 打开 app 时（已配置坚果云）：询问是否从云端获取最新数据（手动同步，不自动取）。
function askOpenFetch() {
  askConfirm(
    '是否从云端获取最新数据？\n\n获取会用云端最新进度和生词覆盖本机；\n本机尚未上传的改动会被云端覆盖。\n\n（点「取消」暂不获取，之后可点主页面「☁ 上传」上传本机数据）',
    fetchNow);
}

// ---------------------------------------------------------- 启动
if (migrateStorage()) showFirstRunNotice();   // 首次安装/升级：提示先配置坚果云并获取云端数据
state.current = clamp(loadNum(LS_PROG, 0), 0, TOTAL - 1);
state.unknown = loadUnknown();
rebuildUnkSet();
renderWord();
// 先把当前快照和云端配置推给原生（此后每次保存都会更新快照），
// 这样即使一次操作都没有，退出时原生也能用刚进 app 的状态做备份。
pushStateToNative();
syncConfigToNative();
renderSyncButton();
// 手动同步：打开时若已配置坚果云，询问是否从云端获取最新数据；
// 退出时的确认弹窗与「先上传再退出」的上传由 Java 原生在 onBackPressed 处理。
if (isCloudConfigured()) askOpenFetch();
