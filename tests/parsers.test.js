// 실행: node --test tests/
// 문자/엑셀 예시는 실제 형식을 따르되 이름과 계좌는 가짜 값이다.
const test = require('node:test');
const assert = require('node:assert');
const lib = require('../apps-script/Code.gs');

const NOW = new Date(2026, 9, 5, 12, 0); // 2026-10-05

test('신한 입금 문자', () => {
  const text = '[Web발신]\n신한10/03 19:26\n110-***-123456\n입금     250,000\n홍길동(테스트상';
  const t = lib.parseShinhanSms(text, NOW);
  assert.strictEqual(lib.ymd(t.date), '2026-10-03');
  assert.strictEqual(t.time, '19:26');
  assert.strictEqual(t.in, 250000);
  assert.strictEqual(t.out, 0);
  assert.strictEqual(t.memo, '홍길동(테스트상');
});

test('신한 출금 문자', () => {
  const text = '[Web발신]\n신한10/02 12:32\n110-***-123456\n출금     850,000\n김철수';
  const t = lib.parseShinhanSms(text, NOW);
  assert.strictEqual(t.out, 850000);
  assert.strictEqual(t.in, 0);
  assert.strictEqual(t.memo, '김철수');
});

test('우리카드 승인 문자', () => {
  const text = '[Web발신]\n우리(1234)승인\n홍*동님\n3,000원 일시불\n10/05 11:47\n서울고속버스터미널(\n누적1,234,567원';
  const t = lib.parseWooriSms(text, NOW);
  assert.strictEqual(lib.ymd(t.date), '2026-10-05');
  assert.strictEqual(t.time, '11:47');
  assert.strictEqual(t.amount, 3000);
  assert.strictEqual(t.merchant, '서울고속버스터미널');
  assert.strictEqual(t.kind, '승인 일시불');
});

test('우리카드 취소 문자는 음수', () => {
  const text = '[Web발신]\n우리(1234)취소\n홍*동님\n15,000원 일시불\n10/04 09:10\n테스트상점\n누적1,000원';
  assert.strictEqual(lib.parseWooriSms(text, NOW).amount, -15000);
});

test('다른 문자는 무시', () => {
  assert.strictEqual(lib.parseShinhanSms('[Web발신]\n인증번호 123456', NOW), null);
  assert.strictEqual(lib.parseWooriSms('[Web발신]\n인증번호 123456', NOW), null);
});

test('연초에 받은 12월 날짜는 작년', () => {
  const jan = new Date(2027, 0, 3);
  assert.strictEqual(lib.ymd(lib.inferDate(12, 30, jan)), '2026-12-30');
  assert.strictEqual(lib.ymd(lib.inferDate(1, 2, jan)), '2027-01-02');
});

test('신한 거래내역 엑셀', () => {
  const rows = [
    ['거래일자', '거래시간', '적요', '출금(원)', '입금(원)', '내용', '잔액(원)', '거래점'],
    ['20261003', '19:26:50', '타행모바일\n뱅킹', '0', '250,000', '홍길동(테스트\n상', '**********', '(국민)'],
    ['20261002', '12:32:04', '모바일', '850,000', '0', '김철수', '**********', '용산금'],
  ];
  const out = lib.parseShinhanExport(rows, NOW);
  assert.strictEqual(out.length, 2);
  assert.strictEqual(lib.ymd(out[0].date), '2026-10-03');
  assert.strictEqual(out[0].time, '19:26');
  assert.strictEqual(out[0].summary, '타행모바일 뱅킹');
  assert.strictEqual(out[0].in, 250000);
  assert.strictEqual(out[1].out, 850000);
  assert.strictEqual(out[1].branch, '용산금');
});

test('문자와 엑셀의 통장 키가 같다', () => {
  const sms = lib.parseShinhanSms('[Web발신]\n신한10/03 19:26\n110-***-123456\n입금     250,000\n홍길동', NOW);
  const xls = lib.parseShinhanExport([
    ['거래일자', '거래시간', '적요', '출금(원)', '입금(원)', '내용', '잔액(원)', '거래점'],
    ['20261003', '19:26:50', '타행', '0', '250,000', '홍길동', '', ''],
  ], NOW)[0];
  assert.strictEqual(lib.bankKey(sms), lib.bankKey(xls));
});

test('우리카드 명세서 엑셀 (소계, 0원 행 제외)', () => {
  const rows = [
    ['   이용대금명세서 상세 내역', '', '', '', '', '', '', ''],
    ['이용\n일자', '카드\n구분', '이용\n카드', '매출\n구분', '이용가맹점(은행)명', '이용금액\n(해외현지/\n체크카드)', '할부\n개월', '당월결제하실금액'],
    ['', '', '', '', '', '', '', '회차'],
    ['09.01', '신용/본인', '1234', '국내일시불', '테스트도자기', '41,800', '0', '0'],
    ['09.13', '신용/본인', '1234', '국내일시불', '알림서비스 면제', '0', '0', '0'],
    ['09.30', '신용/본인', '1234', '국내일시불', '전기요금(자동이체)', '121,090', '0', '0'],
    ['', '신용/본인', '1234', '카드소계', '소계', '', '0', '0'],
  ];
  const out = lib.parseWooriStatement(rows, NOW);
  assert.strictEqual(out.length, 2);
  assert.strictEqual(lib.ymd(out[0].date), '2026-09-01');
  assert.strictEqual(out[0].amount, 41800);
  assert.strictEqual(out[1].merchant, '전기요금(자동이체)');
});

test('같은 키가 여러 번이면 개수만큼만 확인 처리', () => {
  const day = new Date(2026, 8, 7);
  const items = [
    { date: day, amount: 3000 }, { date: day, amount: 3000 }, { date: day, amount: 5000 },
  ];
  const existing = [lib.cardKey({ date: day, amount: 3000 })];
  const res = lib.matchByKey(existing, items, lib.cardKey);
  assert.strictEqual(res.confirmed.length, 1);
  assert.strictEqual(res.add.length, 2);
});

test('분류규칙은 처음 맞는 것', () => {
  const rules = [{ keyword: '카카오', category: '교통·배송' }, { keyword: '데코', category: '매입' }];
  assert.strictEqual(lib.classify('카카오T_주식회사', rules), '교통·배송');
  assert.strictEqual(lib.classify('(주)현대데코', rules), '매입');
  assert.strictEqual(lib.classify('모르는곳', rules), '');
});
