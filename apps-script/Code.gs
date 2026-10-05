/**
 * OXX 장부 자동화
 *
 * - 신한은행 입출금 문자, 우리카드 승인 문자를 웹앱(doPost)으로 받아 통장내역/카드내역에 기록
 * - 신한 거래내역 엑셀, 우리카드 이용대금명세서 엑셀을 '가져오기' 탭에 붙여넣으면 반영
 * - 분류규칙으로 지출을 자동 분류하고, 매출/매입/현금출납 탭과 대조해 '장부 확인' 칸을 채움
 * - 통장에서 현금을 뽑으면 현금출납 탭에 자동으로 'ATM 인출' 기록
 *
 * 열은 머리글 이름으로 찾으므로 열 순서를 바꿔도 동작한다.
 */

var SHEETS = {
  sales: '매출',
  purchase: '매입',
  cash: '현금출납',
  bank: '통장내역',
  card: '카드내역',
  import: '가져오기',
  sms: '문자수신',
  settings: '설정',
};

var MATCH_DAYS = 3;          // 카드/계좌 지출과 매입 기록의 날짜 허용 오차
var SALES_BEFORE_DAYS = 60;  // 입금일 기준, 주문일이 이만큼 이전이어도 같은 매출로 봄
var SALES_AFTER_DAYS = 30;   // 입금일 기준, 주문일이 이만큼 이후여도 같은 매출로 봄 (선입금)

// ---------------------------------------------------------------------------
// 메뉴
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('OXX 장부')
    .addItem('가져오기 탭 반영 (통장/카드 엑셀)', 'importPasted')
    .addItem('분류 다시 적용 + 장부 대조', 'refreshAll')
    .addItem('선택한 입금을 매출에 추가', 'addSelectedDepositsToSales')
    .addSeparator()
    .addItem('문자 연동 설정 보기', 'showSmsSetup')
    .addToUi();
}

function refreshAll() {
  var ss = SpreadsheetApp.getActive();
  applyRules_(ss);
  syncCashWithdrawals_(ss);
  reconcile_(ss);
  toast_('분류와 대조를 마쳤습니다.');
}

// ---------------------------------------------------------------------------
// 문자 웹앱
// ---------------------------------------------------------------------------

/**
 * 아이폰 단축어가 보내는 JSON: {"token": "...", "text": "문자 원문"}
 */
function doPost(e) {
  var body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    body = { text: e.parameter && e.parameter.text, token: e.parameter && e.parameter.token };
  }
  var token = PropertiesService.getScriptProperties().getProperty('SMS_TOKEN');
  if (!token || body.token !== token) {
    return json_({ ok: false, error: 'token' });
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var result = handleSms_(SpreadsheetApp.getActive(), String(body.text || ''), new Date());
    return json_({ ok: true, result: result });
  } finally {
    lock.releaseLock();
  }
}

function handleSms_(ss, text, now) {
  var smsSheet = ss.getSheetByName(SHEETS.sms);
  var bank = parseShinhanSms(text, now);
  var card = bank ? null : parseWooriSms(text, now);
  var kind = bank ? '통장' : card ? '카드' : '인식 못함';
  if (smsSheet) smsSheet.appendRow([now, kind, text]);

  if (bank) {
    upsertBank_(ss, [bank], '문자');
  } else if (card) {
    upsertCard_(ss, [card], '문자');
  } else {
    return 'unrecognized';
  }
  applyRules_(ss);
  syncCashWithdrawals_(ss);
  reconcile_(ss);
  return kind;
}

function showSmsSetup() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('SMS_TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('SMS_TOKEN', token);
  }
  var url = ScriptApp.getService().getUrl() || '(아직 웹앱으로 배포하지 않았습니다)';
  SpreadsheetApp.getUi().alert(
    '문자 연동 설정',
    '웹앱 주소:\n' + url + '\n\n토큰:\n' + token +
      '\n\n아이폰 단축어의 요청 본문(JSON)에 token 과 text 를 넣으세요.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ---------------------------------------------------------------------------
// 파서 (순수 함수, tests/parsers.test.js 에서 검사)
// ---------------------------------------------------------------------------

/** 연도가 없는 월/일에 연도를 붙인다. 지금보다 한 달 넘게 미래면 작년으로 본다. */
function inferDate(month, day, now) {
  var year = now.getFullYear();
  var d = new Date(year, month - 1, day);
  var limit = new Date(now.getFullYear(), now.getMonth() + 1, now.getDate());
  if (d > limit) d = new Date(year - 1, month - 1, day);
  return d;
}

function toNumber(v) {
  if (typeof v === 'number') return v;
  var s = String(v == null ? '' : v).replace(/[,\s원]/g, '');
  if (s === '' || s === '-') return 0;
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}

function ymd(d) {
  return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
}

function pad2(n) {
  return (n < 10 ? '0' : '') + n;
}

/**
 * 신한 입출금 문자
 *   [Web발신]
 *   신한10/03 19:26
 *   110-***-****** (계좌, 무시)
 *   입금     250,000
 *   정단비(오후웍스
 */
function parseShinhanSms(text, now) {
  var head = text.match(/신한\s*(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})/);
  var amt = text.match(/(입금|출금)\s+([\d,]+)/);
  if (!head || !amt) return null;
  var lines = text.split(/\r?\n/).map(function (l) { return l.trim(); });
  var idx = -1;
  for (var i = 0; i < lines.length; i++) {
    if (/^(입금|출금)\s+[\d,]+/.test(lines[i])) { idx = i; break; }
  }
  var memo = '';
  for (var j = idx + 1; idx >= 0 && j < lines.length; j++) {
    if (lines[j] && !/^잔액/.test(lines[j])) { memo = lines[j]; break; }
  }
  var amount = toNumber(amt[2]);
  return {
    date: inferDate(+head[1], +head[2], now),
    time: pad2(+head[3]) + ':' + head[4],
    summary: '',
    out: amt[1] === '출금' ? amount : 0,
    in: amt[1] === '입금' ? amount : 0,
    memo: memo,
    branch: '',
  };
}

/**
 * 우리카드 승인 문자
 *   [Web발신]
 *   우리(5272)승인
 *   이*진님
 *   3,000원 일시불
 *   10/05 11:47
 *   서울고속버스터미널(
 *   누적 ...원
 */
function parseWooriSms(text, now) {
  var head = text.match(/우리\((\d{4})\)\s*(승인|취소)/);
  if (!head) return null;
  var amt = text.match(/([\d,]+)원\s*(일시불|\d+개월)?/);
  var when = text.match(/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})/);
  if (!amt || !when) return null;
  var lines = text.split(/\r?\n/).map(function (l) { return l.trim(); });
  var merchant = '';
  for (var i = 0; i < lines.length; i++) {
    if (/^\d{1,2}\/\d{1,2}\s+\d{1,2}:\d{2}/.test(lines[i])) {
      merchant = (lines[i + 1] || '').replace(/\($/, '');
      break;
    }
  }
  var sign = head[2] === '취소' ? -1 : 1;
  return {
    date: inferDate(+when[1], +when[2], now),
    time: pad2(+when[3]) + ':' + when[4],
    merchant: merchant,
    amount: sign * toNumber(amt[1]),
    kind: head[2] + (amt[2] ? ' ' + amt[2] : ''),
  };
}

/** 머리글 행을 찾아 {이름: 열번호} 를 돌려준다. */
function findHeader(rows, required) {
  for (var r = 0; r < Math.min(rows.length, 15); r++) {
    var map = {};
    rows[r].forEach(function (v, c) {
      var name = String(v).replace(/\s+/g, '');
      if (name) map[name] = c;
    });
    var ok = required.every(function (k) { return k in map; });
    if (ok) return { row: r, map: map };
  }
  return null;
}

function parseExportDate(v, now) {
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  var s = String(v).trim();
  var m = s.match(/^(\d{4})[.\-\/]?(\d{2})[.\-\/]?(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})[.\-\/](\d{1,2})$/);
  if (m) return inferDate(+m[1], +m[2], now);
  return null;
}

/** 신한 거래내역 엑셀 (거래일자, 거래시간, 적요, 출금(원), 입금(원), 내용, 잔액(원), 거래점) */
function parseShinhanExport(rows, now) {
  var h = findHeader(rows, ['거래일자', '출금(원)', '입금(원)']);
  if (!h) return null;
  var m = h.map, out = [];
  for (var r = h.row + 1; r < rows.length; r++) {
    var row = rows[r];
    var date = parseExportDate(row[m['거래일자']], now);
    if (!date) continue;
    var time = String(row[m['거래시간']] || '');
    if (Object.prototype.toString.call(row[m['거래시간']]) === '[object Date]') {
      time = pad2(row[m['거래시간']].getHours()) + ':' + pad2(row[m['거래시간']].getMinutes());
    }
    out.push({
      date: date,
      time: time.slice(0, 5),
      summary: String(row[m['적요']] || '').replace(/\s+/g, ' ').trim(),
      out: toNumber(row[m['출금(원)']]),
      in: toNumber(row[m['입금(원)']]),
      memo: String(row[m['내용']] || '').replace(/\s+/g, ' ').trim(),
      branch: String(row[m['거래점']] || '').trim(),
    });
  }
  return out;
}

/** 우리카드 이용대금명세서 엑셀 (이용일자, 이용가맹점(은행)명, 이용금액 ...) */
function parseWooriStatement(rows, now) {
  var h = findHeader(rows, ['이용일자', '이용가맹점(은행)명']);
  if (!h) return null;
  var m = h.map, amtCol = -1;
  Object.keys(m).forEach(function (k) { if (k.indexOf('이용금액') === 0) amtCol = m[k]; });
  var out = [];
  for (var r = h.row + 1; r < rows.length; r++) {
    var row = rows[r];
    var date = parseExportDate(row[m['이용일자']], now);
    if (!date) continue;
    var amount = toNumber(row[amtCol]);
    if (amount === 0) continue;
    out.push({
      date: date,
      time: '',
      merchant: String(row[m['이용가맹점(은행)명']] || '').trim(),
      amount: amount,
      kind: String(row[m['매출구분']] || '').trim(),
    });
  }
  return out;
}

function bankKey(t) {
  return ['B', ymd(t.date), t.time.slice(0, 5), t.out, t.in].join('|');
}

function cardKey(t) {
  return ['C', ymd(t.date), t.amount].join('|');
}

/** 이미 있는 키 목록과 새 항목을 비교해, 새로 추가할 것과 기존 행 중 확인된 것을 나눈다. */
function matchByKey(existingKeys, items, keyFn) {
  var pool = {};
  existingKeys.forEach(function (k, i) {
    if (!k) return;
    (pool[k] = pool[k] || []).push(i);
  });
  var add = [], confirmed = [];
  items.forEach(function (it) {
    var k = keyFn(it);
    if (pool[k] && pool[k].length) confirmed.push({ index: pool[k].shift(), item: it });
    else add.push(it);
  });
  return { add: add, confirmed: confirmed };
}

/** 분류규칙: [{keyword, category}] 중 처음 맞는 것 */
function classify(text, rules) {
  var s = String(text || '');
  for (var i = 0; i < rules.length; i++) {
    if (rules[i].keyword && s.indexOf(rules[i].keyword) >= 0) return rules[i].category;
  }
  return '';
}

function daysBetween(a, b) {
  return Math.round((b - a) / 86400000);
}

// ---------------------------------------------------------------------------
// 시트 입출력
// ---------------------------------------------------------------------------

function table_(sheet) {
  var values = sheet.getDataRange().getValues();
  var header = values[0].map(function (v) { return String(v).trim(); });
  var col = {};
  header.forEach(function (h, i) { if (h) col[h] = i; });
  return { sheet: sheet, header: header, col: col, rows: values.slice(1) };
}

function need_(t, names) {
  names.forEach(function (n) {
    if (!(n in t.col)) throw new Error("'" + t.sheet.getName() + "' 탭에 '" + n + "' 머리글이 없습니다.");
  });
}

function lastDataRow_(t, colName) {
  var c = t.col[colName];
  for (var i = t.rows.length - 1; i >= 0; i--) {
    if (t.rows[i][c] !== '' && t.rows[i][c] !== null) return i + 2;
  }
  return 1;
}

function appendRows_(t, objs, keyCol) {
  if (!objs.length) return;
  var start = lastDataRow_(t, keyCol) + 1;
  // 쓰는 열까지만 덮어써서 오른쪽의 요약 칸(현금출납 G:H 등)을 건드리지 않는다.
  var width = 0;
  objs.forEach(function (o) {
    Object.keys(o).forEach(function (k) { if (k in t.col) width = Math.max(width, t.col[k] + 1); });
  });
  var data = objs.map(function (o) {
    var row = new Array(width).fill('');
    Object.keys(o).forEach(function (k) { if (k in t.col) row[t.col[k]] = o[k]; });
    return row;
  });
  t.sheet.getRange(start, 1, data.length, width).setValues(data);
}

function setColumn_(t, name, values) {
  if (!values.length) return;
  t.sheet.getRange(2, t.col[name] + 1, values.length, 1)
    .setValues(values.map(function (v) { return [v]; }));
}

function upsertBank_(ss, items, source) {
  var t = table_(ss.getSheetByName(SHEETS.bank));
  need_(t, ['거래일자', '거래시간', '적요', '출금', '입금', '내용', '출처', '키']);
  var keys = t.rows.map(function (r) { return String(r[t.col['키']]); });
  var res = matchByKey(keys, items, bankKey);
  res.confirmed.forEach(function (c) {
    var r = c.index + 2, row = t.rows[c.index];
    var src = String(row[t.col['출처']]);
    if (src.indexOf(source) < 0) t.sheet.getRange(r, t.col['출처'] + 1).setValue(src ? src + '+' + source : source);
    if (c.item.summary && !row[t.col['적요']]) t.sheet.getRange(r, t.col['적요'] + 1).setValue(c.item.summary);
    if (c.item.memo && String(c.item.memo).length > String(row[t.col['내용']]).length) {
      t.sheet.getRange(r, t.col['내용'] + 1).setValue(c.item.memo);
    }
    if (c.item.branch && '거래점' in t.col) t.sheet.getRange(r, t.col['거래점'] + 1).setValue(c.item.branch);
  });
  appendRows_(t, res.add.map(function (it) {
    return {
      '거래일자': it.date, '거래시간': "'" + it.time, '적요': it.summary,
      '출금': it.out || '', '입금': it.in || '', '내용': it.memo, '거래점': it.branch,
      '출처': source, '키': bankKey(it),
    };
  }), '키');
  return { added: res.add.length, confirmed: res.confirmed.length };
}

function upsertCard_(ss, items, source) {
  var t = table_(ss.getSheetByName(SHEETS.card));
  need_(t, ['이용일', '가맹점명', '이용금액', '출처', '키']);
  var keys = t.rows.map(function (r) { return String(r[t.col['키']]); });
  var res = matchByKey(keys, items, cardKey);
  res.confirmed.forEach(function (c) {
    var r = c.index + 2, row = t.rows[c.index];
    var src = String(row[t.col['출처']]);
    if (src.indexOf(source) < 0) t.sheet.getRange(r, t.col['출처'] + 1).setValue(src ? src + '+' + source : source);
    if (String(c.item.merchant).length > String(row[t.col['가맹점명']]).length) {
      t.sheet.getRange(r, t.col['가맹점명'] + 1).setValue(c.item.merchant);
    }
  });
  appendRows_(t, res.add.map(function (it) {
    return {
      '이용일': it.date, '이용시간': it.time ? "'" + it.time : '', '가맹점명': it.merchant,
      '이용금액': it.amount, '승인구분': it.kind, '출처': source, '키': cardKey(it),
    };
  }), '키');
  return { added: res.add.length, confirmed: res.confirmed.length };
}

// ---------------------------------------------------------------------------
// 가져오기
// ---------------------------------------------------------------------------

function importPasted() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(SHEETS.import);
  var rows = sheet.getDataRange().getValues();
  var now = new Date();
  var bank = parseShinhanExport(rows, now);
  var msg;
  if (bank) {
    var r1 = upsertBank_(ss, bank, '엑셀');
    msg = '통장 내역: 새로 ' + r1.added + '건, 이미 있던 ' + r1.confirmed + '건 확인';
  } else {
    var card = parseWooriStatement(rows, now);
    if (!card) {
      SpreadsheetApp.getUi().alert('가져오기 탭에서 신한 거래내역이나 우리카드 명세서 머리글을 찾지 못했습니다.');
      return;
    }
    var r2 = upsertCard_(ss, card, '명세서');
    msg = '카드 내역: 새로 ' + r2.added + '건, 이미 있던 ' + r2.confirmed + '건 확인';
  }
  sheet.clearContents();
  applyRules_(ss);
  syncCashWithdrawals_(ss);
  reconcile_(ss);
  SpreadsheetApp.getUi().alert(msg + '\n가져오기 탭은 비웠습니다.');
}

// ---------------------------------------------------------------------------
// 분류
// ---------------------------------------------------------------------------

function readRules_(ss) {
  var values = ss.getSheetByName(SHEETS.settings).getDataRange().getValues();
  var start = -1, kc = -1;
  for (var r = 0; r < values.length && start < 0; r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (String(values[r][c]).trim() === '키워드' && String(values[r][c + 1]).trim() === '분류') {
        start = r + 1; kc = c; break;
      }
    }
  }
  var rules = [];
  for (var i = start; start >= 0 && i < values.length; i++) {
    var k = String(values[i][kc]).trim(), cat = String(values[i][kc + 1]).trim();
    if (k && cat) rules.push({ keyword: k, category: cat });
  }
  return rules;
}

/** '분류'가 비어 있는 행만 채운다. 직접 고친 분류는 덮어쓰지 않는다. */
function applyRules_(ss) {
  var rules = readRules_(ss);
  [[SHEETS.bank, ['적요', '내용']], [SHEETS.card, ['가맹점명']]].forEach(function (spec) {
    var t = table_(ss.getSheetByName(spec[0]));
    if (!('분류' in t.col)) return;
    var changed = false;
    var vals = t.rows.map(function (row) {
      var cur = row[t.col['분류']];
      if (cur) return cur;
      var hasData = spec[1].some(function (n) { return row[t.col[n]]; });
      if (!hasData) return '';
      var text = spec[1].map(function (n) { return row[t.col[n]]; }).join(' ');
      var cat = classify(text, rules);
      if (!cat && spec[0] === SHEETS.bank && toNumber(row[t.col['입금']]) > 0) cat = '매출입금';
      if (cat) changed = true;
      return cat;
    });
    if (changed) setColumn_(t, '분류', vals);
  });
}

/** 통장내역에서 분류가 '현금인출'인 출금을 현금출납 탭에 'ATM 인출'로 옮긴다. */
function syncCashWithdrawals_(ss) {
  var bank = table_(ss.getSheetByName(SHEETS.bank));
  var cash = table_(ss.getSheetByName(SHEETS.cash));
  if (!('키' in cash.col)) return;
  var have = {};
  cash.rows.forEach(function (r) { if (r[cash.col['키']]) have[r[cash.col['키']]] = true; });
  var add = [];
  bank.rows.forEach(function (r) {
    if (r[bank.col['분류']] !== '현금인출') return;
    var key = r[bank.col['키']];
    var amount = toNumber(r[bank.col['출금']]);
    if (!key || !amount || have[key]) return;
    add.push({
      '날짜': r[bank.col['거래일자']], '구분': 'ATM 인출', '들어온 금액': amount,
      '메모': '통장내역 자동 (' + (r[bank.col['적요']] || r[bank.col['내용']]) + ')', '키': key,
    });
  });
  appendRows_(cash, add, '날짜');
}

// ---------------------------------------------------------------------------
// 대조
// ---------------------------------------------------------------------------

function asDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  return parseExportDate(v, new Date());
}

function reconcile_(ss) {
  var sales = table_(ss.getSheetByName(SHEETS.sales));
  var pur = table_(ss.getSheetByName(SHEETS.purchase));
  var bank = table_(ss.getSheetByName(SHEETS.bank));
  var card = table_(ss.getSheetByName(SHEETS.card));

  // 매입: 날짜별 카드/계좌이체 결제 금액
  var purCard = [], purTransfer = [], purCardByDay = {};
  pur.rows.forEach(function (r) {
    var d = asDate_(r[pur.col['날짜']]);
    if (!d) return;
    var c = toNumber(r[pur.col['카드']]), e = toNumber(r[pur.col['계좌이체']]);
    if (c) {
      purCard.push({ date: d, amount: c, used: false });
      purCardByDay[ymd(d)] = (purCardByDay[ymd(d)] || 0) + c;
    }
    if (e) purTransfer.push({ date: d, amount: e, used: false });
  });

  function take(list, date, amount) {
    for (var i = 0; i < list.length; i++) {
      var x = list[i];
      if (!x.used && x.amount === amount && Math.abs(daysBetween(x.date, date)) <= MATCH_DAYS) {
        x.used = true;
        return true;
      }
    }
    return false;
  }

  // 카드
  var cardDay = {};
  card.rows.forEach(function (r) {
    var d = asDate_(r[card.col['이용일']]);
    if (d && r[card.col['분류']] === '매입') cardDay[ymd(d)] = (cardDay[ymd(d)] || 0) + toNumber(r[card.col['이용금액']]);
  });
  var cardStatus = card.rows.map(function (r) {
    var d = asDate_(r[card.col['이용일']]);
    if (!d) return '';
    var cat = r[card.col['분류']];
    var amount = toNumber(r[card.col['이용금액']]);
    if (!cat) return '분류 필요';
    if (cat !== '매입') return '경비';
    if (take(purCard, d, amount)) return '매입 있음';
    if (cardDay[ymd(d)] && cardDay[ymd(d)] === purCardByDay[ymd(d)]) return '일 합계 일치';
    return '매입 확인 필요';
  });
  setColumn_(card, '장부 확인', cardStatus);

  // 매출: 합계 금액, 결제일 또는 주문일
  var saleList = sales.rows.map(function (r) {
    return {
      order: asDate_(r[sales.col['주문일']]),
      paid: asDate_(r[sales.col['결제일']]),
      amount: toNumber(r[sales.col['합계']]),
      used: false,
    };
  }).filter(function (s) { return s.amount && (s.order || s.paid); });

  function takeSale(date, amount) {
    for (var i = 0; i < saleList.length; i++) {
      var s = saleList[i];
      if (s.used || s.amount !== amount) continue;
      var ok = s.paid ? Math.abs(daysBetween(s.paid, date)) <= MATCH_DAYS
        : (daysBetween(s.order, date) <= SALES_BEFORE_DAYS && daysBetween(date, s.order) <= SALES_AFTER_DAYS);
      if (ok) { s.used = true; return true; }
    }
    return false;
  }

  var bankStatus = bank.rows.map(function (r) {
    var d = asDate_(r[bank.col['거래일자']]);
    if (!d) return '';
    var cat = r[bank.col['분류']];
    var inAmt = toNumber(r[bank.col['입금']]), outAmt = toNumber(r[bank.col['출금']]);
    if (inAmt) {
      if (cat && cat !== '매출입금') return '-';
      return takeSale(d, inAmt) ? '매출 있음' : '매출 확인 필요';
    }
    if (!cat) return '분류 필요';
    if (cat === '현금인출') return '현금출납 반영';
    if (cat === '매입') return take(purTransfer, d, outAmt) ? '매입 있음' : '매입 확인 필요';
    return '경비';
  });
  setColumn_(bank, '장부 확인', bankStatus);
}

// ---------------------------------------------------------------------------
// 입금 → 매출
// ---------------------------------------------------------------------------

function addSelectedDepositsToSales() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  if (sheet.getName() !== SHEETS.bank) {
    SpreadsheetApp.getUi().alert('통장내역 탭에서 매출로 추가할 입금 행을 선택한 뒤 실행하세요.');
    return;
  }
  var t = table_(sheet);
  var range = sheet.getActiveRange();
  var add = [];
  for (var r = range.getRow(); r < range.getRow() + range.getNumRows(); r++) {
    if (r < 2) continue;
    var row = t.rows[r - 2];
    var amount = toNumber(row[t.col['입금']]);
    if (!amount) continue;
    add.push({
      '주문일': row[t.col['거래일자']], '주문자': row[t.col['내용']], '구분': 'BtoC',
      '상품금액': amount, '결제수단': '입금', '결제일': row[t.col['거래일자']],
      '입금자명': row[t.col['내용']], '메모': '통장내역에서 추가',
    });
  }
  var sales = table_(ss.getSheetByName(SHEETS.sales));
  appendRows_(sales, add, '주문자');
  reconcile_(ss);
  toast_(add.length + '건을 매출 탭에 추가했습니다. 구분(BtoC/BtoB)과 주문자를 확인하세요.');
}

// ---------------------------------------------------------------------------

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function toast_(msg) {
  SpreadsheetApp.getActive().toast(msg, 'OXX 장부', 5);
}

if (typeof module !== 'undefined') {
  module.exports = {
    inferDate: inferDate, toNumber: toNumber, ymd: ymd,
    parseShinhanSms: parseShinhanSms, parseWooriSms: parseWooriSms,
    parseShinhanExport: parseShinhanExport, parseWooriStatement: parseWooriStatement,
    bankKey: bankKey, cardKey: cardKey, matchByKey: matchByKey, classify: classify,
  };
}
