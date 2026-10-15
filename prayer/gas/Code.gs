// 오늘의 기도 영상 앱 저장소 (GitHub Pages의 prayer/index.html에서 fetch로 사용)
// 영상 파일은 내 구글 드라이브 '오늘의 기도 영상' 폴더에, 목록은 스크립트 속성에 저장
// GET  ?a=list                       → 영상 목록 + 오늘 날짜(한국 시간)
// GET  ?a=video&id=..                → 영상 하나를 base64로 (드라이브 공유가 막혀 있을 때 재생용)
// POST {a:'mine',family,pin}         → 우리 가정 영상 (가정 암호 확인, 처음이면 통과)
// POST {a:'upload',family,pin,kind,name,mime,data}
//      kind: 'child'(아이 기도) | 'parent'(부모 기도). 가정마다 하나씩.
//      새 영상은 오늘부터 비어 있는 가장 빠른 날짜에 배정, 같은 가정·같은 종류면 영상만 바꾸고 날짜는 그대로
// POST {a:'del',id,family,pin}       → 우리 가정 영상 지우기 (관리자는 pw로)
// POST {a:'setpw'|'check'|'chpw', pw, newPw} → 관리자 암호
// POST {a:'move',pw,id,date}         → 날짜 바꾸기 (그 날짜에 다른 영상이 있으면 서로 맞바꿈)
// POST {a:'resetpin',pw,family}      → 가정 암호 초기화
var CHUNK = 2500, TZ = 'Asia/Seoul', MAX_B64 = 36 * 1024 * 1024;
var KINDS = { child: '아이 기도', parent: '부모 기도' };

function props() { return PropertiesService.getScriptProperties(); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hash(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'prayer:' + s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function readKey(k) {
  var p = props(), n = Number(p.getProperty('k:' + k + ':n') || 0);
  if (!n) return null;
  var s = '';
  for (var i = 0; i < n; i++) s += p.getProperty('k:' + k + ':' + i) || '';
  try { return JSON.parse(s); } catch (e) { return null; }
}
function writeKey(k, v) {
  var p = props(), s = JSON.stringify(v), old = Number(p.getProperty('k:' + k + ':n') || 0);
  var obj = {}, n = Math.max(1, Math.ceil(s.length / CHUNK));
  for (var i = 0; i < n; i++) obj['k:' + k + ':' + i] = s.slice(i * CHUNK, (i + 1) * CHUNK);
  obj['k:' + k + ':n'] = String(n);
  p.setProperties(obj);
  for (var j = n; j < old; j++) p.deleteProperty('k:' + k + ':' + j);
}

function today() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function addDays(d, n) { var t = new Date(d + 'T00:00:00Z'); t.setUTCDate(t.getUTCDate() + n); return t.toISOString().slice(0, 10); }
function famKey(s) { return String(s || '').replace(/\s+/g, '').toLowerCase(); }
function clean(s, n) { return String(s || '').replace(/\s+/g, ' ').trim().slice(0, n); }
function entries() { return readKey('entries') || []; }
// upTo가 있으면 그 날짜 이후(아직 공개 전) 영상의 파일 주소는 숨김
function pub(e, upTo) {
  var open = e.shared && (!upTo || e.date <= upTo);
  return { id: e.id, family: e.family, kind: e.kind, name: e.name, date: e.date, mime: e.mime,
    fileId: open ? e.fileId : '', shared: !!e.shared, created: e.created };
}
function freeDate(list, from) {
  var used = {};
  list.forEach(function (e) { used[e.date] = 1; });
  var d = from;
  while (used[d]) d = addDays(d, 1);
  return d;
}
function folder() {
  var p = props(), id = p.getProperty('folder');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var nf = DriveApp.createFolder('오늘의 기도 영상');
  p.setProperty('folder', nf.getId());
  return nf;
}
function trash(fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {} }
// 가정 암호: 처음 쓰면 그 암호로 정해지고, 다음부터는 같아야 함
function famOk(family, pin) {
  var fk = famKey(family);
  if (!fk) return '가정 이름을 써 주세요.';
  if (!/^\d{4}$/.test(String(pin || ''))) return '가정 암호는 숫자 4자리예요.';
  var saved = props().getProperty('pin:' + fk);
  if (saved && saved !== hash(fk + ':' + pin)) return '가정 암호가 틀려요.';
  return '';
}
function adminOk(pw) { var s = props().getProperty('pw'); return !!s && hash(pw || '') === s; }

function doGet(e) {
  var a = (e && e.parameter && e.parameter.a) || 'list';
  if (a === 'video') {
    var it = entries().filter(function (x) { return x.id === e.parameter.id; })[0];
    if (!it) return json({ ok: false, msg: '영상을 찾을 수 없어요.' });
    if (it.date > today()) return json({ ok: false, msg: '아직 공개 전이에요.' });
    var blob = DriveApp.getFileById(it.fileId).getBlob();
    return json({ ok: true, mime: it.mime, data: Utilities.base64Encode(blob.getBytes()) });
  }
  var t = today();
  return json({ ok: true, today: t, hasPw: !!props().getProperty('pw'),
    entries: entries().map(function (x) { return pub(x, t); }) });
}

function doPost(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, msg: '잘못된 요청' }); }
  var p = props();

  if (b.a === 'mine') {
    var m = famOk(b.family, b.pin);
    if (m) return json({ ok: false, msg: m });
    var fk = famKey(b.family);
    return json({ ok: true, isNew: !p.getProperty('pin:' + fk), today: today(),
      entries: entries().filter(function (x) { return famKey(x.family) === fk; }).map(function (x) { return pub(x); }) });
  }

  if (b.a === 'upload') {
    var msg = famOk(b.family, b.pin);
    if (msg) return json({ ok: false, msg: msg });
    if (!KINDS[b.kind]) return json({ ok: false, msg: '아이 기도인지 부모 기도인지 골라 주세요.' });
    var name = clean(b.name, 20), family = clean(b.family, 20);
    if (!name) return json({ ok: false, msg: '기도하는 사람 이름을 써 주세요.' });
    if (!b.data) return json({ ok: false, msg: '영상이 없어요.' });
    if (String(b.data).length > MAX_B64) return json({ ok: false, msg: '영상이 너무 커요.' });
    var mime = /^video\/[\w.+-]+$/.test(String(b.mime || '').split(';')[0]) ? String(b.mime).split(';')[0] : 'video/mp4';
    var ext = mime.indexOf('webm') >= 0 ? 'webm' : mime.indexOf('quicktime') >= 0 ? 'mov' : 'mp4';
    var file = folder().createFile(Utilities.newBlob(Utilities.base64Decode(b.data), mime,
      family + '_' + KINDS[b.kind] + '_' + name + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss') + '.' + ext));
    var shared = false;
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); shared = true; } catch (err) {}

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var fk2 = famKey(family), list = entries(), now = new Date().toISOString();
      if (!p.getProperty('pin:' + fk2)) p.setProperty('pin:' + fk2, hash(fk2 + ':' + b.pin));
      var old = list.filter(function (x) { return famKey(x.family) === fk2 && x.kind === b.kind; })[0], it2;
      if (old) {
        trash(old.fileId);
        old.family = family; old.name = name; old.fileId = file.getId(); old.mime = mime; old.shared = shared; old.created = now;
        it2 = old;
      } else {
        it2 = { id: Utilities.getUuid().slice(0, 8), family: family, kind: b.kind, name: name,
          date: freeDate(list, today()), fileId: file.getId(), mime: mime, shared: shared, created: now };
        list.push(it2);
      }
      list.sort(function (x, y) { return x.date < y.date ? -1 : x.date > y.date ? 1 : 0; });
      writeKey('entries', list);
      return json({ ok: true, replaced: !!old, entry: pub(it2) });
    } finally { lock.releaseLock(); }
  }

  if (b.a === 'del') {
    var admin = adminOk(b.pw);
    if (!admin) { var m2 = famOk(b.family, b.pin); if (m2) return json({ ok: false, msg: m2 }); }
    var lock2 = LockService.getScriptLock();
    lock2.waitLock(20000);
    try {
      var list2 = entries(), it3 = list2.filter(function (x) { return x.id === b.id; })[0];
      if (!it3) return json({ ok: false, msg: '영상을 찾을 수 없어요.' });
      if (!admin && famKey(it3.family) !== famKey(b.family)) return json({ ok: false, msg: '우리 가정 영상만 지울 수 있어요.' });
      trash(it3.fileId);
      writeKey('entries', list2.filter(function (x) { return x.id !== b.id; }));
      return json({ ok: true });
    } finally { lock2.releaseLock(); }
  }

  // ----- 관리자 -----
  var pw = p.getProperty('pw');
  if (b.a === 'setpw') {
    if (pw) return json({ ok: false, msg: '이미 관리자 암호가 있어요.' });
    if (!b.pw || String(b.pw).length < 4) return json({ ok: false, msg: '암호는 4자 이상으로 해 주세요.' });
    p.setProperty('pw', hash(b.pw));
    return json({ ok: true });
  }
  if (!adminOk(b.pw)) return json({ ok: false, msg: '관리자 암호가 틀려요.' });
  if (b.a === 'check') return json({ ok: true });
  if (b.a === 'chpw') {
    if (!b.newPw || String(b.newPw).length < 4) return json({ ok: false, msg: '암호는 4자 이상으로 해 주세요.' });
    p.setProperty('pw', hash(b.newPw));
    return json({ ok: true });
  }
  if (b.a === 'resetpin') {
    p.deleteProperty('pin:' + famKey(b.family));
    return json({ ok: true });
  }
  if (b.a === 'move') {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(b.date || ''))) return json({ ok: false, msg: '날짜가 이상해요.' });
    var lock3 = LockService.getScriptLock();
    lock3.waitLock(20000);
    try {
      var list3 = entries(), it4 = list3.filter(function (x) { return x.id === b.id; })[0];
      if (!it4) return json({ ok: false, msg: '영상을 찾을 수 없어요.' });
      var other = list3.filter(function (x) { return x.date === b.date && x.id !== b.id; })[0];
      if (other) other.date = it4.date;
      it4.date = b.date;
      list3.sort(function (x, y) { return x.date < y.date ? -1 : x.date > y.date ? 1 : 0; });
      writeKey('entries', list3);
      return json({ ok: true, swapped: other ? other.name : '' });
    } finally { lock3.releaseLock(); }
  }
  return json({ ok: false, msg: '알 수 없는 요청' });
}

// 처음 한 번 편집기에서 실행해서 드라이브 권한을 허락해 두세요.
function setup() { Logger.log(folder().getUrl()); }
