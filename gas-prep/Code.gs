// 전치사 3D 방 사진 저장소 (preposition-room.html의 '📸 Snap 찰칵'이 fetch로 사용)
// POST {a:'snap', grade, cls, num, name, sentences:[...], img:'<base64 JPEG>'}
//   → 내 드라이브 '전치사 방 앨범/2학년 3반/' 폴더에 사진 저장 (설명란에 학생이 쓴 문장)
//   → 같은 폴더의 '전치사 방 제출 목록' 시트에 한 줄 기록
// 처음 한 번: 편집기에서 setup() 실행 → 권한 허용 → 실행 로그에 나온 주소가 반별 앨범 폴더
var ROOT_NAME = '전치사 방 앨범';
var SHEET_NAME = '전치사 방 제출 목록';

function props() { return PropertiesService.getScriptProperties(); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clampInt(v, lo, hi) { var n = parseInt(v, 10); return n >= lo && n <= hi ? n : 0; }

function rootFolder() {
  var id = props().getProperty('root');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder(ROOT_NAME);
  props().setProperty('root', f.getId());
  return f;
}
function subFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function sheet() {
  var id = props().getProperty('sheet');
  if (id) { try { return SpreadsheetApp.openById(id).getSheets()[0]; } catch (e) {} }
  var ss = SpreadsheetApp.create(SHEET_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(rootFolder());
  var sh = ss.getSheets()[0];
  sh.appendRow(['보낸 시각', '학년', '반', '번호', '이름', '문장 수', '문장', '사진']);
  sh.setFrozenRows(1);
  props().setProperty('sheet', ss.getId());
  return sh;
}

function setup() {
  var root = rootFolder();
  sheet();
  Logger.log('반별 앨범 폴더: ' + root.getUrl());
}

function doGet() { return json({ ok: true, app: 'prep-room' }); }

function doPost(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return json({ ok: false, err: 'bad_request' }); }
  if (!b || b.a !== 'snap') return json({ ok: false, err: 'unknown_action' });
  var grade = clampInt(b.grade, 1, 3), cls = clampInt(b.cls, 1, 20), num = clampInt(b.num, 1, 60);
  var name = String(b.name || '').replace(/[\\\/:*?"<>|\r\n\t]/g, '').trim().slice(0, 20);
  if (!grade || !cls || !num || !name) return json({ ok: false, err: 'who' });
  var img = String(b.img || '');
  if (!img || img.length > 3000000) return json({ ok: false, err: 'img' });
  var sents = (Array.isArray(b.sentences) ? b.sentences : []).slice(0, 40).map(function (s) { return String(s).slice(0, 200); });
  var bytes;
  try { bytes = Utilities.base64Decode(img); } catch (x) { return json({ ok: false, err: 'img' }); }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var folder = subFolder(rootFolder(), grade + '학년 ' + cls + '반');
    var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd HHmm');
    var file = folder.createFile(Utilities.newBlob(bytes, 'image/jpeg', num + '번 ' + name + ' ' + stamp + '.jpg'));
    file.setDescription(sents.join('\n'));
    sheet().appendRow([new Date(), grade, cls, num, name, sents.length, sents.join('\n'), file.getUrl()]);
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true });
}
