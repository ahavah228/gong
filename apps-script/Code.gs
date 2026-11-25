/**
 * 공덕교회 주차관리 — Google Apps Script 백엔드
 *
 * 1) 이 파일 전체를 구글 시트의 [확장 프로그램 > Apps Script]에 붙여넣습니다.
 * 2) 아래 INITIAL_ADMIN_PASSWORD 를 바꾼 뒤, 상단에서 setup 함수를 선택해 [실행]합니다.
 * 3) [배포 > 새 배포 > 웹 앱] (실행: 나 / 액세스: 모든 사용자) 로 배포하고 URL을 복사합니다.
 */

const INITIAL_ADMIN_PASSWORD = 'gongdeok1234'; // ← setup() 실행 전에 꼭 바꾸세요

const SH = { V: '교인차량', R: '주차신청', S: '설정' };

const HEAD = {
  '교인차량': ['ID', '이름', '연락처', '차량번호', '차종', '소속', 'PIN', '등록일시', '상태', '메모'],
  '주차신청': ['ID', '신청일시', '예배일', '예배', '차량번호', '이름', '연락처', '구분', '사유', '상태', '처리자', '처리일시', '입차시각', '메모'],
  '설정': ['항목', '값', '설명'],
};

const DEFAULTS = [
  ['교회명', '공덕교회', '앱 상단에 표시됩니다'],
  ['예배목록', '1부 예배,2부 예배,3부 예배,오후 예배', '쉼표(,)로 구분'],
  ['예배별주차면수', '30', '예배 1회당 주차 가능 대수'],
  ['자동승인', 'N', 'Y: 면수 안에서 즉시 승인 / N: 관리자가 승인'],
  ['신청가능일수', '14', '오늘부터 며칠 뒤까지 신청 가능한지'],
  ['신청구분', '장애인,임산부,노약자,영유아 동반,봉사자,새가족,일반', '쉼표(,)로 구분'],
  ['공지', '주일 주차 공간이 부족합니다. 가급적 대중교통을 이용해 주세요.', '앱 상단 공지 (비우면 숨김)'],
];

const ACTIVE = ['신청', '대기', '승인', '입차'];        // 자리를 차지하는 상태
const ADMIN_STATUSES = ['신청', '대기', '승인', '반려', '취소', '입차'];

/* ───────────────────────── 초기 설정 ───────────────────────── */

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEAD).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange('A:Z').setNumberFormat('@'); // 전화번호 앞자리 0, 날짜 자동변환 방지
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]])
        .setFontWeight('bold').setBackground('#e8eef7');
      sh.setFrozenRows(1);
    }
  });
  const s = ss.getSheetByName(SH.S);
  if (s.getLastRow() === 1) s.getRange(2, 1, DEFAULTS.length, 3).setValues(DEFAULTS);

  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('ADMIN_PASSWORD')) p.setProperty('ADMIN_PASSWORD', INITIAL_ADMIN_PASSWORD);
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());

  ['시트1', 'Sheet1'].forEach(n => {
    const x = ss.getSheetByName(n);
    if (x && x.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(x);
  });
  Logger.log('설정 완료! 이제 [배포 > 새 배포 > 웹 앱]으로 배포하세요.');
}

/* ───────────────────────── 웹 앱 진입점 ───────────────────────── */

function doGet() {
  return json_({ ok: true, data: '공덕교회 주차관리 API 동작 중' });
}

function doPost(e) {
  let res;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = API[req.action];
    if (!fn) throw new Error('알 수 없는 요청입니다.');
    res = { ok: true, data: fn(req) };
  } catch (err) {
    res = { ok: false, error: err.message || String(err) };
  }
  return json_(res);
}

/* ───────────────────────── API ───────────────────────── */

const API = {
  config() {
    const s = settings_();
    return {
      church: s['교회명'] || '공덕교회',
      services: list_(s['예배목록']),
      categories: list_(s['신청구분']),
      capacity: Number(s['예배별주차면수'] || 0),
      autoApprove: s['자동승인'] === 'Y',
      daysAhead: Number(s['신청가능일수'] || 14),
      notice: s['공지'] || '',
      today: today_(),
    };
  },

  /** 신규 차량 등록 (처음 등록 시 회원 생성) */
  register(r) {
    const name = String(r.name || '').trim();
    const phone = phone_(r.phone);
    const plate = plate_(r.plate);
    const pin = String(r.pin || '');
    if (!name) throw new Error('이름을 입력하세요.');
    checkPhone_(phone); checkPlate_(plate); checkPin_(pin);
    if (!r.agree) throw new Error('개인정보 수집·이용에 동의해 주세요.');

    return lock_(() => {
      const vs = vehicles_();
      if (vs.some(v => v['차량번호'] === plate)) throw new Error('이미 등록된 차량번호입니다. 관리자에게 문의하세요.');
      const mine = vs.filter(v => v['연락처'] === phone);
      const h = hash_(pin, phone);
      if (mine.length && mine[0]['PIN'] !== h) throw new Error('이미 등록된 연락처입니다. 로그인 후 [내 정보]에서 차량을 추가하세요.');
      append_(SH.V, {
        ID: id_('V'), 이름: name, 연락처: phone, 차량번호: plate,
        차종: clip_(r.model, 30), 소속: clip_(r.group, 30), PIN: h, 등록일시: now_(), 상태: '정상',
      });
      return session_(phone);
    });
  },

  login(r) {
    const phone = phone_(r.phone);
    const mine = vehicles_().filter(v => v['연락처'] === phone);
    if (!mine.length || mine[0]['PIN'] !== hash_(String(r.pin || ''), phone)) {
      throw new Error('연락처 또는 PIN이 올바르지 않습니다.');
    }
    return session_(phone);
  },

  me(r) {
    const p = verify_(r.token, 'member');
    return profile_(p.phone);
  },

  addVehicle(r) {
    const p = verify_(r.token, 'member');
    const plate = plate_(r.plate);
    checkPlate_(plate);
    return lock_(() => {
      const vs = vehicles_();
      if (vs.some(v => v['차량번호'] === plate)) throw new Error('이미 등록된 차량번호입니다.');
      const mine = vs.filter(v => v['연락처'] === p.phone);
      if (!mine.length) throw new Error('회원 정보를 찾을 수 없습니다.');
      const base = mine[0];
      append_(SH.V, {
        ID: id_('V'), 이름: base['이름'], 연락처: p.phone, 차량번호: plate,
        차종: clip_(r.model, 30), 소속: base['소속'], PIN: base['PIN'], 등록일시: now_(), 상태: '정상',
      });
      return profile_(p.phone);
    });
  },

  removeVehicle(r) {
    const p = verify_(r.token, 'member');
    return lock_(() => {
      const mine = vehicles_().filter(v => v['연락처'] === p.phone);
      if (mine.length <= 1) throw new Error('마지막 차량은 삭제할 수 없습니다. 관리자에게 문의하세요.');
      const v = mine.find(x => x['차량번호'] === plate_(r.plate));
      if (!v) throw new Error('차량을 찾을 수 없습니다.');
      update_(SH.V, v._row, { 상태: '삭제' });
      return profile_(p.phone);
    });
  },

  /** 주차 신청 */
  apply(r) {
    const p = verify_(r.token, 'member');
    const cfg = API.config();
    const date = String(r.date || '');
    const plate = plate_(r.plate);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('예배일을 선택하세요.');
    if (date < cfg.today) throw new Error('지난 날짜는 신청할 수 없습니다.');
    if (date > addDays_(cfg.today, cfg.daysAhead)) throw new Error(`${cfg.daysAhead}일 이내 날짜만 신청할 수 있습니다.`);
    if (cfg.services.indexOf(r.service) < 0) throw new Error('예배를 선택하세요.');
    const category = cfg.categories.indexOf(r.category) >= 0 ? r.category : '일반';

    return lock_(() => {
      const v = vehicles_().find(x => x['연락처'] === p.phone && x['차량번호'] === plate);
      if (!v) throw new Error('내 차량을 선택하세요.');
      const same = rows_(SH.R).filter(x => x['예배일'] === date && x['예배'] === r.service && ACTIVE.indexOf(x['상태']) >= 0);
      if (same.some(x => x['차량번호'] === plate)) throw new Error('이미 같은 예배에 신청한 차량입니다.');

      let status = '신청';
      if (cfg.autoApprove) {
        const approved = same.filter(x => x['상태'] === '승인' || x['상태'] === '입차').length;
        status = approved < cfg.capacity ? '승인' : '대기';
      }
      append_(SH.R, {
        ID: id_('R'), 신청일시: now_(), 예배일: date, 예배: r.service, 차량번호: plate,
        이름: v['이름'], 연락처: p.phone, 구분: category, 사유: clip_(r.reason, 100), 상태: status,
        처리자: status === '신청' ? '' : '자동', 처리일시: status === '신청' ? '' : now_(),
      });
      return { status, requests: myRequests_(p.phone) };
    });
  },

  myRequests(r) {
    const p = verify_(r.token, 'member');
    return myRequests_(p.phone);
  },

  cancel(r) {
    const p = verify_(r.token, 'member');
    return lock_(() => {
      const x = rows_(SH.R).find(q => q.ID === r.id && q['연락처'] === p.phone);
      if (!x) throw new Error('신청 내역을 찾을 수 없습니다.');
      if (['신청', '대기', '승인'].indexOf(x['상태']) < 0) throw new Error('취소할 수 없는 상태입니다.');
      update_(SH.R, x._row, { 상태: '취소', 처리자: '본인', 처리일시: now_() });
      return myRequests_(p.phone);
    });
  },

  /* ───── 관리자 ───── */

  adminLogin(r) {
    const name = String(r.name || '').trim();
    if (!name) throw new Error('관리자 이름을 입력하세요.');
    const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
    if (!pw || String(r.password || '') !== pw) {
      Utilities.sleep(800);
      throw new Error('비밀번호가 올바르지 않습니다.');
    }
    return { token: sign_({ role: 'admin', name, exp: Date.now() + 12 * 3600e3 }), name };
  },

  adminList(r) {
    verify_(r.token, 'admin');
    const cfg = API.config();
    const date = String(r.date || cfg.today);
    const list = rows_(SH.R).filter(x => x['예배일'] === date).map(publicReq_);
    const summary = cfg.services.map(sv => {
      const xs = list.filter(x => x.service === sv);
      const c = st => xs.filter(x => x.status === st).length;
      return { service: sv, capacity: cfg.capacity, 신청: c('신청'), 대기: c('대기'), 승인: c('승인'), 입차: c('입차') };
    });
    return { date, summary, list };
  },

  adminSet(r) {
    const a = verify_(r.token, 'admin');
    if (ADMIN_STATUSES.indexOf(r.status) < 0) throw new Error('잘못된 상태입니다.');
    return lock_(() => {
      const x = rows_(SH.R).find(q => q.ID === r.id);
      if (!x) throw new Error('신청 내역을 찾을 수 없습니다.');
      const patch = { 상태: r.status, 처리자: a.name, 처리일시: now_() };
      if (r.status === '입차') patch['입차시각'] = Utilities.formatDate(new Date(), tz_(), 'HH:mm');
      if (r.memo != null) patch['메모'] = clip_(r.memo, 100);
      update_(SH.R, x._row, patch);
      return API.adminList({ token: r.token, date: x['예배일'] });
    });
  },

  /** 관리자 대리 신청 (전화·현장 요청) */
  adminCreate(r) {
    const a = verify_(r.token, 'admin');
    const cfg = API.config();
    const plate = plate_(r.plate);
    checkPlate_(plate);
    if (cfg.services.indexOf(r.service) < 0) throw new Error('예배를 선택하세요.');
    return lock_(() => {
      const v = vehicles_().find(x => x['차량번호'] === plate);
      append_(SH.R, {
        ID: id_('R'), 신청일시: now_(), 예배일: r.date, 예배: r.service, 차량번호: plate,
        이름: v ? v['이름'] : clip_(r.name, 20) || '(미등록)', 연락처: v ? v['연락처'] : phone_(r.phone),
        구분: r.category || '일반', 사유: clip_(r.reason, 100), 상태: r.status || '승인',
        처리자: a.name, 처리일시: now_(),
        입차시각: r.status === '입차' ? Utilities.formatDate(new Date(), tz_(), 'HH:mm') : '',
      });
      return API.adminList({ token: r.token, date: r.date });
    });
  },

  adminVehicles(r) {
    verify_(r.token, 'admin');
    const q = String(r.q || '').replace(/\s/g, '').toUpperCase();
    if (q.length < 2) throw new Error('2글자 이상 입력하세요.');
    return vehicles_()
      .filter(v => v['차량번호'].indexOf(q) >= 0 || v['이름'].indexOf(q) >= 0 || v['연락처'].indexOf(q) >= 0)
      .slice(0, 50)
      .map(v => ({ name: v['이름'], phone: v['연락처'], plate: v['차량번호'], model: v['차종'], group: v['소속'] }));
  },
};

/* ───────────────────────── 내부 함수 ───────────────────────── */

function session_(phone) {
  return { token: sign_({ role: 'member', phone, exp: Date.now() + 30 * 86400e3 }), ...profile_(phone) };
}

function profile_(phone) {
  const mine = vehicles_().filter(v => v['연락처'] === phone);
  if (!mine.length) throw new Error('회원 정보를 찾을 수 없습니다. 다시 로그인하세요.');
  return {
    name: mine[0]['이름'], phone, group: mine[0]['소속'],
    vehicles: mine.map(v => ({ plate: v['차량번호'], model: v['차종'] })),
  };
}

function myRequests_(phone) {
  return rows_(SH.R).filter(x => x['연락처'] === phone)
    .sort((a, b) => (b['예배일'] + b['신청일시']).localeCompare(a['예배일'] + a['신청일시']))
    .slice(0, 30).map(publicReq_);
}

function publicReq_(x) {
  return {
    id: x.ID, createdAt: x['신청일시'], date: x['예배일'], service: x['예배'], plate: x['차량번호'],
    name: x['이름'], phone: x['연락처'], category: x['구분'], reason: x['사유'], status: x['상태'],
    by: x['처리자'], at: x['처리일시'], inTime: x['입차시각'], memo: x['메모'],
  };
}

function vehicles_() { return rows_(SH.V).filter(v => v['상태'] !== '삭제'); }

function sheet_(n) {
  const s = SpreadsheetApp.getActive().getSheetByName(n);
  if (!s) throw new Error(`'${n}' 시트가 없습니다. Apps Script에서 setup()을 먼저 실행하세요.`);
  return s;
}

function rows_(n) {
  const v = sheet_(n).getDataRange().getValues();
  const h = v.shift() || [];
  return v.map((r, i) => {
    const o = { _row: i + 2 };
    h.forEach((k, j) => { o[k] = str_(r[j]); });
    return o;
  }).filter(o => o.ID);
}

function append_(n, obj) {
  const s = sheet_(n);
  const row = s.getLastRow() + 1;
  s.getRange(row, 1, 1, HEAD[n].length).setNumberFormat('@')
    .setValues([HEAD[n].map(k => obj[k] == null ? '' : String(obj[k]))]);
}

function update_(n, row, patch) {
  const s = sheet_(n);
  Object.keys(patch).forEach(k => {
    const c = HEAD[n].indexOf(k);
    if (c >= 0) s.getRange(row, c + 1).setNumberFormat('@').setValue(String(patch[k]));
  });
}

function settings_() {
  const o = {};
  sheet_(SH.S).getDataRange().getValues().slice(1).forEach(r => { if (r[0]) o[String(r[0]).trim()] = str_(r[1]).trim(); });
  return o;
}

function lock_(fn) {
  const l = LockService.getScriptLock();
  if (!l.tryLock(15000)) throw new Error('요청이 많습니다. 잠시 후 다시 시도하세요.');
  try { return fn(); } finally { l.releaseLock(); }
}

function secret_() {
  const s = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!s) throw new Error('setup()을 먼저 실행하세요.');
  return s;
}

function hash_(pin, phone) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, secret_() + '|' + phone + '|' + pin, Utilities.Charset.UTF_8));
}

function sign_(payload) {
  const body = Utilities.base64EncodeWebSafe(JSON.stringify(payload), Utilities.Charset.UTF_8);
  const sig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(body, secret_()));
  return body + '.' + sig;
}

function verify_(token, role) {
  if (!token) throw new Error('로그인이 필요합니다.');
  const [body, sig] = String(token).split('.');
  const expect = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(body || '', secret_()));
  if (sig !== expect) throw new Error('로그인 정보가 올바르지 않습니다. 다시 로그인하세요.');
  const p = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(body)).getDataAsString('UTF-8'));
  if (p.exp < Date.now()) throw new Error('로그인이 만료되었습니다. 다시 로그인하세요.');
  if (role && p.role !== role) throw new Error('권한이 없습니다.');
  return p;
}

function checkPhone_(p) { if (!/^01\d{8,9}$/.test(p)) throw new Error('휴대폰 번호를 정확히 입력하세요.'); }
function checkPin_(p) { if (!/^\d{4,6}$/.test(p)) throw new Error('PIN은 숫자 4~6자리입니다.'); }
function checkPlate_(p) {
  if (!/^([가-힣]{2})?\d{2,3}[가-힣]\d{4}$/.test(p)) throw new Error('차량번호 형식을 확인하세요. 예) 12가3456, 123가4567');
}

const phone_ = p => String(p || '').replace(/\D/g, '');
const plate_ = p => String(p || '').replace(/\s/g, '').toUpperCase();
const list_ = s => String(s || '').split(',').map(x => x.trim()).filter(Boolean);
const clip_ = (s, n) => String(s || '').trim().slice(0, n);
const id_ = pre => pre + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMddHHmmss') + Math.floor(Math.random() * 900 + 100);

function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Seoul'; }
function now_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'); }
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function addDays_(ymd, n) {
  const d = new Date(ymd + 'T12:00:00');
  d.setDate(d.getDate() + n);
  return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
}
function str_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return v == null ? '' : String(v);
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
