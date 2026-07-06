/**
 * GO2 Triage — Google Sheets 백엔드 v5 (놓친 멘션 백필)
 *
 * v4 대비 추가:
 *   - backfillMentions: 봇이 있는 모든 채널에서 최근 24시간 멘션 일괄 수집
 *
 * v3 대비 추가:
 *   - P1 멘션 수신 시 이메일 알림 (MailApp)
 *   - 팀원 Slack ID ↔ 이름 매핑 저장 (Members 시트)
 *   - 테스트 이메일 발송
 *
 * 스크립트 속성 (프로젝트 설정 → 스크립트 속성):
 *   SLACK_BOT_TOKEN  : xoxb- 로 시작하는 봇 토큰
 *   MY_SLACK_USER_ID : 본인 Slack 멤버 ID (U로 시작)
 *   EMAIL_NOTIFY     : 알림 받을 이메일 (대시보드 설정에서 자동 등록)
 */

const SHEET_NAME = 'Triage';
const INBOX_SHEET = 'Inbox';
const MEMBERS_SHEET = 'Members';
const HEADERS = ['ID', '등록시각', 'Slack시각', '발신자', '우선순위', '유형', '상태', '메시지', '링크', '동기화시각'];
const INBOX_HEADERS = ['이벤트ID', '수신시각', '발신자', '메시지', '링크', '가져옴'];
const MEMBERS_HEADERS = ['이름', 'Slack ID', '이메일'];

/* ───────── 공통 유틸 ───────── */
function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  return sheet;
}
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ───────── 수신 진입점 ───────── */
function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: 'invalid json' });
  }

  // ① Slack 이벤트 구독 URL 검증 (앱 설정 시 1회)
  if (data.type === 'url_verification') {
    return ContentService.createTextOutput(data.challenge);
  }

  // ② Slack 멘션 이벤트 → Inbox에 기록
  if (data.type === 'event_callback') {
    try { handleSlackEvent(data); } catch (err) { /* Slack 재시도 대비 무시 */ }
    return ContentService.createTextOutput('ok');
  }

  // ③ 대시보드 → Inbox 가져옴 표시
  if (data.action === 'markImported') {
    markImported(data.ids || []);
    return jsonOut({ ok: true });
  }

  // ④ 팀원 목록 업데이트
  if (data.action === 'updateMembers') {
    updateMembers(data.members || []);
    return jsonOut({ ok: true });
  }

  // ⑤ 이메일 알림 설정
  if (data.action === 'setEmailNotify') {
    const props = PropertiesService.getScriptProperties();
    if (data.email) props.setProperty('EMAIL_NOTIFY', data.email);
    return jsonOut({ ok: true });
  }

  // ⑥ 테스트 이메일 발송
  if (data.action === 'testEmail') {
    try {
      MailApp.sendEmail({
        to: data.email,
        subject: '[GO2 Triage] 이메일 알림 테스트',
        htmlBody: '<div style="font-family:sans-serif;padding:20px">'
          + '<h2 style="color:#1F5E7A">GO2 Triage 이메일 알림</h2>'
          + '<p>이 메일이 보이면 이메일 알림이 정상 동작해요.</p>'
          + '<p style="color:#61707D;font-size:13px">— GO2 Triage Dashboard</p>'
          + '</div>'
      });
      return jsonOut({ ok: true });
    } catch (err) {
      return jsonOut({ ok: false, error: String(err) });
    }
  }

  // ⑦ (백필은 doGet으로 이동)

  // ⑧ 대시보드 → Triage 시트 동기화 (기존 기능)
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const sheet = getSheet(SHEET_NAME, HEADERS);
    const now = new Date();
    if (data.action === 'replace' && sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).clearContent();
    }
    const rows = (data.items || []).map(function (it) {
      return [it.id, it.date, it.slackTime, it.sender, it.priority, it.msgType || '', it.status, it.message, it.link || '', now];
    });
    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    }
    return jsonOut({ ok: true, count: rows.length });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ───────── Slack 이벤트 처리 ───────── */
function handleSlackEvent(data) {
  const ev = data.event || {};
  if (ev.type !== 'message' || ev.subtype || ev.bot_id) return;

  const props = PropertiesService.getScriptProperties();
  const myId = props.getProperty('MY_SLACK_USER_ID');
  if (!myId || !ev.text) return;
  if (ev.text.indexOf('<@' + myId + '>') === -1) return;
  if (ev.user === myId) return;

  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const sheet = getSheet(INBOX_SHEET, INBOX_HEADERS);
    const evId = ev.client_msg_id || (ev.channel + '-' + ev.ts);

    const last = sheet.getLastRow();
    if (last > 1) {
      const ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
      if (ids.indexOf(evId) !== -1) return;
    }

    const token = props.getProperty('SLACK_BOT_TOKEN');
    const sender = resolveSenderName(token, ev.user);
    const permalink = getPermalink(token, ev.channel, ev.ts);
    const when = new Date(parseFloat(ev.ts) * 1000);
    const text = cleanSlackText(ev.text, myId);

    sheet.appendRow([evId, when, sender, text, permalink, false]);

    // P1 키워드 감지 시 이메일 알림
    sendP1EmailIfNeeded(props, sender, text, permalink);
  } finally {
    lock.releaseLock();
  }
}

/* ───────── 이메일 알림 ───────── */
function sendP1EmailIfNeeded(props, sender, text, link) {
  const email = props.getProperty('EMAIL_NOTIFY');
  if (!email) return;

  var p1Keywords = ['긴급', 'urgent', 'ASAP', '장애', 'VND'];
  var isP1 = false;
  var lower = text.toLowerCase();
  for (var i = 0; i < p1Keywords.length; i++) {
    if (lower.indexOf(p1Keywords[i].toLowerCase()) !== -1) { isP1 = true; break; }
  }
  if (!isP1) return;

  try {
    var linkHtml = link ? '<p><a href="' + link + '" style="color:#1F5E7A">Slack에서 보기</a></p>' : '';
    MailApp.sendEmail({
      to: email,
      subject: '[GO2 Triage] P1 멘션: ' + sender,
      htmlBody: '<div style="font-family:sans-serif;padding:20px">'
        + '<h2 style="color:#E5484D">P1 긴급 멘션</h2>'
        + '<p><strong>발신자:</strong> ' + sender + '</p>'
        + '<p><strong>메시지:</strong> ' + text.substring(0, 500) + '</p>'
        + linkHtml
        + '<p style="color:#61707D;font-size:13px;margin-top:16px">— GO2 Triage Dashboard</p>'
        + '</div>'
    });
  } catch (err) {
    console.log('이메일 발송 실패: ' + err);
  }
}

/* ───────── 팀원 관리 ───────── */
function updateMembers(members) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const sheet = getSheet(MEMBERS_SHEET, MEMBERS_HEADERS);
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, MEMBERS_HEADERS.length).clearContent();
    }
    if (members.length) {
      var rows = members.map(function (m) {
        return [m.name || '', m.slackId || '', m.email || ''];
      });
      sheet.getRange(2, 1, rows.length, MEMBERS_HEADERS.length).setValues(rows);
    }
  } finally {
    lock.releaseLock();
  }
}

function resolveSenderName(token, userId) {
  // Members 시트에서 먼저 찾기
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBERS_SHEET);
    if (sheet && sheet.getLastRow() > 1) {
      var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][1] === userId) return rows[i][0];
      }
    }
  } catch (e) {}

  // 못 찾으면 Slack API로 조회
  if (!token || !userId) return userId || '미상';
  try {
    var res = UrlFetchApp.fetch('https://slack.com/api/users.info?user=' + encodeURIComponent(userId), {
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
    });
    var j = JSON.parse(res.getContentText());
    if (j.ok) return (j.user.profile && j.user.profile.display_name) || j.user.real_name || userId;
  } catch (e) {}
  return userId;
}

function getPermalink(token, channel, ts) {
  if (!token) return '';
  try {
    var res = UrlFetchApp.fetch('https://slack.com/api/chat.getPermalink?channel=' + encodeURIComponent(channel) + '&message_ts=' + encodeURIComponent(ts), {
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
    });
    var j = JSON.parse(res.getContentText());
    if (j.ok) return j.permalink;
  } catch (e) {}
  return '';
}

function cleanSlackText(text, myId) {
  return text
    .replace(new RegExp('<@' + myId + '>', 'g'), '@나')
    .replace(/<@([A-Z0-9]+)>/g, '@$1')
    .replace(/<(https?:[^|>]+)\|([^>]+)>/g, '$2 ($1)')
    .replace(/<(https?:[^>]+)>/g, '$1')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

/* ───────── 대시보드가 Inbox를 끌어갈 때 ───────── */
function doGet(e) {
  var mode = e && e.parameter && e.parameter.mode;
  if (mode === 'inbox') {
    var sheet = getSheet(INBOX_SHEET, INBOX_HEADERS);
    var last = sheet.getLastRow();
    var items = [];
    if (last > 1) {
      var rows = sheet.getRange(2, 1, last - 1, INBOX_HEADERS.length).getValues();
      rows.forEach(function (r) {
        if (r[5] === true) return;
        items.push({
          id: String(r[0]),
          date: (r[1] instanceof Date) ? r[1].toISOString() : String(r[1]),
          sender: String(r[2] || ''),
          message: String(r[3] || ''),
          link: String(r[4] || '')
        });
      });
    }
    return jsonOut({ ok: true, items: items });
  }
  if (mode === 'members') {
    var sheet = getSheet(MEMBERS_SHEET, MEMBERS_HEADERS);
    var last = sheet.getLastRow();
    var members = [];
    if (last > 1) {
      var rows = sheet.getRange(2, 1, last - 1, MEMBERS_HEADERS.length).getValues();
      rows.forEach(function (r) {
        members.push({ name: String(r[0]), slackId: String(r[1]), email: String(r[2] || '') });
      });
    }
    return jsonOut({ ok: true, members: members });
  }
  if (mode === 'backfill') {
    try {
      var hours = parseInt(e.parameter.hours) || 24;
      var count = backfillMentions(hours);
      return jsonOut({ ok: true, count: count });
    } catch (err) {
      return jsonOut({ ok: false, error: String(err) });
    }
  }
  return ContentService.createTextOutput('GO2 Triage 백엔드 v5 정상 동작 중이에요.');
}

/* ───────── 놓친 멘션 백필 ───────── */
function backfillMentions(hours) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('SLACK_BOT_TOKEN');
  var myId = props.getProperty('MY_SLACK_USER_ID');
  if (!token || !myId) throw new Error('SLACK_BOT_TOKEN 또는 MY_SLACK_USER_ID가 설정되지 않았어요');

  var oldest = String(Math.floor((Date.now() - (hours || 24) * 3600 * 1000) / 1000));
  var channels = getBotChannels(token);
  if (!channels.length) throw new Error('봇이 참여한 채널이 없어요. 채널에 /invite @GO2 Triage 로 초대해주세요');

  var sheet = getSheet(INBOX_SHEET, INBOX_HEADERS);
  var existingIds = new Set();
  var last = sheet.getLastRow();
  if (last > 1) {
    sheet.getRange(2, 1, last - 1, 1).getValues().flat().forEach(function(id) {
      existingIds.add(String(id));
    });
  }

  var added = 0;
  var lock = LockService.getScriptLock();

  for (var ci = 0; ci < channels.length; ci++) {
    var chId = channels[ci];
    var cursor = null;
    var pages = 0;

    do {
      var url = 'https://slack.com/api/conversations.history?channel=' + encodeURIComponent(chId)
        + '&oldest=' + oldest + '&limit=200';
      if (cursor) url += '&cursor=' + encodeURIComponent(cursor);

      var res = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
      });
      var data = JSON.parse(res.getContentText());
      if (!data.ok) break;

      var messages = data.messages || [];
      for (var mi = 0; mi < messages.length; mi++) {
        var msg = messages[mi];
        if (msg.subtype || msg.bot_id) continue;
        if (!msg.text || msg.text.indexOf('<@' + myId + '>') === -1) continue;
        if (msg.user === myId) continue;

        var evId = msg.client_msg_id || (chId + '-' + msg.ts);
        if (existingIds.has(evId)) continue;

        var sender = resolveSenderName(token, msg.user);
        var permalink = getPermalink(token, chId, msg.ts);
        var when = new Date(parseFloat(msg.ts) * 1000);
        var text = cleanSlackText(msg.text, myId);

        lock.tryLock(10000);
        try {
          sheet.appendRow([evId, when, sender, text, permalink, false]);
        } finally {
          lock.releaseLock();
        }
        existingIds.add(evId);
        added++;
      }

      cursor = (data.response_metadata && data.response_metadata.next_cursor) || null;
      pages++;
    } while (cursor && pages < 5);
  }

  return added;
}

function getBotChannels(token) {
  var channels = [];
  var cursor = null;
  var pages = 0;
  do {
    var url = 'https://slack.com/api/conversations.list?types=public_channel,private_channel&exclude_archived=true&limit=200';
    if (cursor) url += '&cursor=' + encodeURIComponent(cursor);

    var res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
    });
    var data = JSON.parse(res.getContentText());
    if (!data.ok) break;

    (data.channels || []).forEach(function(ch) {
      if (ch.is_member) channels.push(ch.id);
    });

    cursor = (data.response_metadata && data.response_metadata.next_cursor) || null;
    pages++;
  } while (cursor && pages < 10);

  return channels;
}

function markImported(ids) {
  if (!ids.length) return;
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var sheet = getSheet(INBOX_SHEET, INBOX_HEADERS);
    var last = sheet.getLastRow();
    if (last < 2) return;
    var range = sheet.getRange(2, 1, last - 1, INBOX_HEADERS.length);
    var rows = range.getValues();
    rows.forEach(function (r) {
      if (ids.indexOf(String(r[0])) !== -1) r[5] = true;
    });
    range.setValues(rows);
  } finally {
    lock.releaseLock();
  }
}
