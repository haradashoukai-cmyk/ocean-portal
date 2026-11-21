// ============================================================
// saily 投稿システム用 Google Apps Script
// 投稿 + 編集 + 削除 + コメント + 画像/短尺動画対応
// シート名:
//   posts    : 投稿データ
//   comments : コメントデータ
//   設定     : ログイン必須（はい／いいえ）と合言葉。初回アクセス時に自動で作られる
//   メンバー : 投稿できる人の名簿。合言葉を入れた人が自動で載る。「状態」を「停止」にするとその人は投稿できなくなる
//
// ログインのしくみ:
//   アプリが Google ログイン（Firebase）で受け取った本人確認の札（idToken）を一緒に送ってくる。
//   ここで Google に問い合わせて本物か確かめ、メンバー名簿に載っている人だけ書き込みを通す。
//   設定シートの「ログイン必須」が「いいえ」の間は、これまでどおり誰でも書き込める（移行期間用）。
// ============================================================

const SPREADSHEET_ID = '1plvsfmgJ-_4UJmiNuJq_plswejZoBT6P26ZHw3LxsL8';
const PHOTO_FOLDER_ID = '1ZgI8oopwAZpM6pvKV1Vu1zK4Upf-sWV1';
// Firebase（saily-movie）のウェブ用キー。アプリの HTML に載っているものと同じで、秘密ではない
const FIREBASE_API_KEY = 'AIzaSyAVXLdzvDH0F4ja5r95kI_7BmP9nKSEbTY';
const SETTINGS_SHEET_NAME = '設定';
const MEMBERS_SHEET_NAME = 'メンバー';
const SETTING_REQUIRE_LOGIN = 'ログイン必須';
const SETTING_PASSCODE = '合言葉';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'create';

    if (action === 'session') return sessionInfo(data);
    if (action === 'join') return joinMember(data);

    // ここから下は書き込み。ログイン必須のときはメンバーだけ通す
    const auth = authorize(data);
    if (!auth.ok) return jsonResponse({ success: false, error: auth.error });

    if (action === 'edit') return editPost(data);
    if (action === 'delete') return deletePost(data);
    if (action === 'addComment') return addComment(data, auth);
    if (action === 'deleteComment') return deleteComment(data);

    return createPost(data, auth);
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'get';
    if (action === 'getComments') return getComments(e.parameter.postId);
    if (action === 'getCommentCounts') return getCommentCounts();
    return getPosts();
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString(), posts: [], comments: [] });
  }
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getPostsSheet() {
  const ss = getSpreadsheet();
  const named = ss.getSheetByName('posts') || ss.getSheetByName('saily_投稿データ') || ss.getSheetByName('投稿データ');
  if (named && sheetLooksLikePosts(named)) return named;

  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheetLooksLikePosts(sheets[i])) return sheets[i];
  }
  return named || sheets[0];
}

function sheetLooksLikePosts(sheet) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 9) return false;
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 9)).getValues()[0].map(String);
  return headers[0] === 'ID'
    && headers[5].indexOf('緯度') !== -1
    && headers[6].indexOf('経度') !== -1
    && (headers[7].indexOf('写真') !== -1 || headers[10] === 'mediaUrl');
}

function getCommentsSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('comments');
  if (!sheet) {
    sheet = ss.insertSheet('comments');
    setupCommentsSheet(sheet);
  }
  return sheet;
}

// ---------- ログインとメンバー名簿 ----------

function getSettingsSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    sheet.appendRow(['項目', '値', '説明']);
    sheet.appendRow([SETTING_REQUIRE_LOGIN, 'いいえ', '「はい」にすると、ログインしてメンバー名簿に載っている人だけが投稿・編集・削除・コメントできます']);
    sheet.appendRow([SETTING_PASSCODE, '', '初めて投稿する人が1回だけ入力する言葉。空欄の間は誰も新しく登録できません']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#0e7490').setFontColor('#ffffff');
  }
  return sheet;
}

function getSetting(key) {
  const rows = getSettingsSheet().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) return rows[i][1];
  }
  return '';
}

function isLoginRequired() {
  const value = String(getSetting(SETTING_REQUIRE_LOGIN)).trim().toLowerCase();
  return ['はい', 'true', 'yes', '1', 'on'].indexOf(value) !== -1;
}

function getMembersSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MEMBERS_SHEET_NAME);
    sheet.appendRow(['メールアドレス', '名前', '登録日時', '状態', 'メモ']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#0e7490').setFontColor('#ffffff');
  }
  return sheet;
}

function findMember(email) {
  const target = String(email || '').trim().toLowerCase();
  if (!target) return null;
  const rows = getMembersSheet().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === target) {
      return { row: i + 1, email: target, name: String(rows[i][1] || ''), active: String(rows[i][3]).trim() !== '停止' };
    }
  }
  return null;
}

function cacheKey(prefix, value) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8);
  return prefix + Utilities.base64EncodeWebSafe(digest);
}

// アプリから届いた本人確認の札を Google に問い合わせて確かめる。本物ならメールアドレスと名前を返す
function verifyIdToken(idToken) {
  if (!idToken) return null;
  const cache = CacheService.getScriptCache();
  const key = cacheKey('tok_', idToken);
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const res = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ idToken: idToken }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return null;
  const user = (JSON.parse(res.getContentText()).users || [])[0];
  if (!user || !user.email || user.emailVerified === false) return null;

  const info = { email: String(user.email).toLowerCase(), name: String(user.displayName || '') };
  cache.put(key, JSON.stringify(info), 300);
  return info;
}

// 外部への問い合わせ（Google への本人確認）が許可されているかを確かめる。エディタから手動で実行する。
// 実行ログに「400 … INVALID_ID_TOKEN」と出れば正常（でたらめな札を送っているので断られるのが正しい）
function checkExternalAccess() {
  const res = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ idToken: 'check' }),
    muteHttpExceptions: true
  });
  Logger.log(res.getResponseCode() + ' ' + res.getContentText().replace(/\s+/g, ' ').slice(0, 200));
}

// 書き込みを通してよいかを決める
function authorize(data) {
  const user = verifyIdToken(data.idToken);
  const member = user ? findMember(user.email) : null;
  const isMember = !!(member && member.active);
  if (!isLoginRequired()) return { ok: true, user: user, member: isMember };
  if (!user) return { ok: false, error: 'login_required' };
  if (!isMember) return { ok: false, error: 'not_member' };
  return { ok: true, user: user, member: true };
}

// アプリが「今ログインしているのは誰か、投稿できるか」を聞くための窓口
function sessionInfo(data) {
  const user = verifyIdToken(data.idToken);
  const member = user ? findMember(user.email) : null;
  return jsonResponse({
    success: true,
    requireLogin: isLoginRequired(),
    loggedIn: !!user,
    member: !!(member && member.active),
    suspended: !!(member && !member.active),
    name: member ? member.name : (user ? user.name : '')
  });
}

function normalizePasscode(value) {
  return String(value === null || value === undefined ? '' : value).normalize('NFKC').trim().toLowerCase();
}

// 合言葉が合っていれば、ログイン中の人をメンバー名簿に載せる
function joinMember(data) {
  const user = verifyIdToken(data.idToken);
  if (!user) return jsonResponse({ success: false, error: 'login_required' });

  const existing = findMember(user.email);
  if (existing) return jsonResponse(existing.active ? { success: true, member: true, name: existing.name } : { success: false, error: 'suspended' });

  const cache = CacheService.getScriptCache();
  const tryKey = cacheKey('try_', user.email);
  const tries = Number(cache.get(tryKey) || 0);
  if (tries >= 5) return jsonResponse({ success: false, error: 'too_many_tries' });

  const passcode = normalizePasscode(getSetting(SETTING_PASSCODE));
  if (!passcode) return jsonResponse({ success: false, error: 'passcode_not_set' });
  if (normalizePasscode(data.passcode) !== passcode) {
    cache.put(tryKey, String(tries + 1), 600);
    return jsonResponse({ success: false, error: 'wrong_passcode' });
  }

  const name = safeCell(String(data.name || user.name || '').trim().slice(0, 40));
  getMembersSheet().appendRow([user.email, name, formatJstTimestamp(), '有効', '']);
  return jsonResponse({ success: true, member: true, name: name });
}

// セルに入れたとき数式として動かないようにする
function safeCell(value) {
  const text = String(value === null || value === undefined ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

// ---------- 投稿 ----------

function createPost(data, auth) {
  const mediaType = data.mediaType === 'video' ? 'video' : 'image';
  const mediaBase64 = data.mediaBase64 || data.photoBase64;
  let fileId = '';
  let photoUrl = '';
  let mediaUrl = '';
  let fileIds = [];
  let mediaUrls = [];

  if (mediaType === 'video') {
    mediaUrl = data.mediaUrl || '';
    if (!mediaUrl) return jsonResponse({ success: false, error: 'video mediaUrl is required' });
    mediaUrls = [mediaUrl];
  } else {
    const mediaItems = Array.isArray(data.mediaItems) && data.mediaItems.length
      ? data.mediaItems.slice(0, 3)
      : [{ base64: mediaBase64, name: data.mediaName || data.photoName || 'photo.jpg' }];
    if (!mediaItems[0] || !mediaItems[0].base64) return jsonResponse({ success: false, error: 'image mediaBase64 is required' });
    const folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    mediaItems.forEach((item, index) => {
      const fileName = `${Date.now()}_${index + 1}_${item.name || data.mediaName || data.photoName || 'photo.jpg'}`;
      const blob = Utilities.newBlob(Utilities.base64Decode(item.base64), item.mimeType || 'image/jpeg', fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const createdFileId = file.getId();
      fileIds.push(createdFileId);
      mediaUrls.push(`https://drive.google.com/thumbnail?id=${createdFileId}&sz=w800`);
    });
    fileId = fileIds[0] || '';
    photoUrl = mediaUrls[0] || '';
    mediaUrl = photoUrl;
  }

  const sheet = getPostsSheet();
  ensurePostHeaders(sheet);
  const id = /^p\d{8,}$/.test(String(data.id || '')) ? String(data.id) : 'p' + Date.now();
  sheet.appendRow([
    id,
    data.timestamp,
    data.author,
    data.title,
    data.comment,
    data.lat,
    data.lng,
    photoUrl,
    fileId,
    mediaType,
    mediaUrl,
    JSON.stringify(mediaUrls),
    JSON.stringify(fileIds),
    auth && auth.user ? auth.user.email : ''
  ]);

  return jsonResponse({ success: true, id: id, photoUrl: photoUrl, mediaType: mediaType, mediaUrl: mediaUrl, mediaUrls: mediaUrls });
}

function inferVideoMimeType(name) {
  const lower = String(name || '').toLowerCase();
  if (lower.endsWith('.mov')) return 'video/quicktime';
  return 'video/mp4';
}

function ensurePostHeaders(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 14)).getValues()[0];
  if (!headers[9]) sheet.getRange(1, 10).setValue('mediaType');
  if (!headers[10]) sheet.getRange(1, 11).setValue('mediaUrl');
  if (!headers[11]) sheet.getRange(1, 12).setValue('mediaUrls');
  if (!headers[12]) sheet.getRange(1, 13).setValue('fileIds');
  // 14列目：投稿した人のメールアドレス（ログインして投稿したときだけ入る。アプリには返さない）
  if (!headers[13]) sheet.getRange(1, 14).setValue('投稿者メール');
}

function editPost(data) {
  const sheet = getPostsSheet();
  const allData = sheet.getDataRange().getValues();

  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.id) {
      const row = i + 1;
      if (data.author !== undefined) sheet.getRange(row, 3).setValue(data.author);
      if (data.title !== undefined) sheet.getRange(row, 4).setValue(data.title);
      if (data.comment !== undefined) sheet.getRange(row, 5).setValue(data.comment);
      return jsonResponse({ success: true, id: data.id, message: '編集しました' });
    }
  }

  return jsonResponse({ success: false, error: '投稿が見つかりません' });
}

function deletePost(data) {
  const sheet = getPostsSheet();
  const allData = sheet.getDataRange().getValues();

  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.id) {
      const fileId = allData[i][8];
      const fileIds = parseJsonArray(allData[i][12]);
      const idsToDelete = fileIds.length ? fileIds : [fileId].filter(Boolean);

      idsToDelete.forEach(id => {
        try {
          DriveApp.getFileById(id).setTrashed(true);
        } catch (e) {
          Logger.log('media delete skipped: ' + e.toString());
        }
      });

      sheet.deleteRow(i + 1);
      deleteCommentsForPost(data.id);
      return jsonResponse({ success: true, id: data.id, message: 'deleted' });
    }
  }

  return jsonResponse({ success: false, error: 'post not found' });
}

function parseJsonArray(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value.filter(Boolean);
  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed.filter(Boolean) : [];
  } catch (e) {
    return [];
  }
}

function getPosts() {
  const sheet = getPostsSheet();
  ensurePostHeaders(sheet);
  const data = sheet.getDataRange().getValues();

  const posts = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const mediaType = row[9] || 'image';
    const mediaUrl = row[10] || row[7] || '';
    const mediaUrls = parseJsonArray(row[11]);
    const fileIds = parseJsonArray(row[12]);
    posts.push({
      id: row[0],
      timestamp: row[1],
      author: row[2],
      title: row[3],
      comment: row[4],
      lat: parseFloat(row[5]),
      lng: parseFloat(row[6]),
      photoUrl: row[7],
      fileId: row[8],
      mediaType: mediaType,
      mediaUrl: mediaUrl,
      mediaUrls: mediaUrls.length ? mediaUrls : (mediaUrl ? [mediaUrl] : []),
      fileIds: fileIds
    });
  }

  posts.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  return jsonResponse({ success: true, posts: posts, requireLogin: isLoginRequired() });
}

function addComment(data, auth) {
  if (!data.postId) return jsonResponse({ success: false, error: '投稿IDがありません' });
  if (!data.comment) return jsonResponse({ success: false, error: 'コメントがありません' });

  const sheet = getCommentsSheet();
  const id = 'c' + Date.now();
  sheet.appendRow([
    id,
    data.postId,
    data.timestamp || formatJstTimestamp(),
    data.author || '匿名',
    data.comment,
    auth && auth.user ? auth.user.email : ''
  ]);

  return jsonResponse({ success: true, id: id });
}

// 全投稿のコメント件数を1回でまとめて返す（投稿ごとに問い合わせると Google の上限に当たるため）
function getCommentCounts() {
  const data = getCommentsSheet().getDataRange().getValues();
  const counts = {};
  for (let i = 1; i < data.length; i++) {
    const postId = data[i][1];
    if (!data[i][0] || !postId) continue;
    counts[postId] = (counts[postId] || 0) + 1;
  }
  return jsonResponse({ success: true, counts: counts });
}

function getComments(postId) {
  if (!postId) return jsonResponse({ success: true, comments: [] });

  const sheet = getCommentsSheet();
  const data = sheet.getDataRange().getValues();
  const comments = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || row[1] !== postId) continue;
    comments.push({
      id: row[0],
      postId: row[1],
      timestamp: row[2],
      author: row[3],
      comment: row[4]
    });
  }

  comments.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
  return jsonResponse({ success: true, comments: comments });
}

function deleteComment(data) {
  const sheet = getCommentsSheet();
  const allData = sheet.getDataRange().getValues();

  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.commentId) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ success: true, id: data.commentId });
    }
  }

  return jsonResponse({ success: false, error: 'コメントが見つかりません' });
}

function deleteCommentsForPost(postId) {
  const sheet = getCommentsSheet();
  const allData = sheet.getDataRange().getValues();
  for (let i = allData.length - 1; i >= 1; i--) {
    if (allData[i][1] === postId) sheet.deleteRow(i + 1);
  }
}

function formatJstTimestamp() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupSpreadsheet() {
  const sheet = getPostsSheet();
  sheet.clear();
  sheet.appendRow([
    'ID', 'タイムスタンプ', '投稿者', 'タイトル', 'コメント',
    '緯度', '経度', '写真URL', 'ファイルID', 'mediaType', 'mediaUrl', 'mediaUrls', 'fileIds'
  ]);
  sheet.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#0e7490').setFontColor('#ffffff');
  Logger.log('posts セットアップ完了しました');
}

function setupCommentsSheet(sheet) {
  sheet.clear();
  sheet.appendRow(['コメントID', '投稿ID', 'タイムスタンプ', '投稿者', 'コメント']);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#0e7490').setFontColor('#ffffff');
}
