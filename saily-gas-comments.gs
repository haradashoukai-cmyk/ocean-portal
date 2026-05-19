// ============================================================
// saily 投稿システム用 Google Apps Script
// 投稿 + 編集 + 削除 + コメント + 画像/短尺動画対応
// シート名:
//   posts    : 投稿データ
//   comments : コメントデータ
// ============================================================

const SPREADSHEET_ID = '1plvsfmgJ-_4UJmiNuJq_plswejZoBT6P26ZHw3LxsL8';
const PHOTO_FOLDER_ID = '1ZgI8oopwAZpM6pvKV1Vu1zK4Upf-sWV1';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'create';

    if (action === 'edit') return editPost(data);
    if (action === 'delete') return deletePost(data);
    if (action === 'addComment') return addComment(data);
    if (action === 'deleteComment') return deleteComment(data);

    return createPost(data);
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'get';
    if (action === 'getComments') return getComments(e.parameter.postId);
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

function createPost(data) {
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
    JSON.stringify(fileIds)
  ]);

  return jsonResponse({ success: true, id: id, photoUrl: photoUrl, mediaType: mediaType, mediaUrl: mediaUrl, mediaUrls: mediaUrls });
}

function inferVideoMimeType(name) {
  const lower = String(name || '').toLowerCase();
  if (lower.endsWith('.mov')) return 'video/quicktime';
  return 'video/mp4';
}

function ensurePostHeaders(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 13)).getValues()[0];
  if (!headers[9]) sheet.getRange(1, 10).setValue('mediaType');
  if (!headers[10]) sheet.getRange(1, 11).setValue('mediaUrl');
  if (!headers[11]) sheet.getRange(1, 12).setValue('mediaUrls');
  if (!headers[12]) sheet.getRange(1, 13).setValue('fileIds');
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
  return jsonResponse({ success: true, posts: posts });
}

function addComment(data) {
  if (!data.postId) return jsonResponse({ success: false, error: '投稿IDがありません' });
  if (!data.comment) return jsonResponse({ success: false, error: 'コメントがありません' });

  const sheet = getCommentsSheet();
  const id = 'c' + Date.now();
  sheet.appendRow([
    id,
    data.postId,
    data.timestamp || formatJstTimestamp(),
    data.author || '匿名',
    data.comment
  ]);

  return jsonResponse({ success: true, id: id });
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
