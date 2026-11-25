// 데모 모드: API_URL이 비어 있을 때 브라우저 안에서 Apps Script를 흉내 냅니다.
// 관리자 데모 비밀번호: 1234
(function () {
  const KEY = 'gdp_demo_db_v1';
  const mem = {};
  const store = {
    get() { try { return JSON.parse(localStorage.getItem(KEY)) || null; } catch (e) { return mem[KEY] || null; } },
    set(v) { try { localStorage.setItem(KEY, JSON.stringify(v)); } catch (e) { mem[KEY] = v; } },
  };
  const pad = n => String(n).padStart(2, '0');
  const ymd = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const now = () => { const d = new Date(); return `${ymd(d)} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`; };
  const hm = () => { const d = new Date(); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };
  const today = () => ymd(new Date());
  const addDays = (s, n) => { const d = new Date(s + 'T12:00:00'); d.setDate(d.getDate() + n); return ymd(d); };
  const nextSun = () => { const d = new Date(); d.setDate(d.getDate() + ((7 - d.getDay()) % 7)); return ymd(d); };
  const id = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
  const phoneN = p => String(p || '').replace(/\D/g, '');
  const plateN = p => String(p || '').replace(/\s/g, '').toUpperCase();
  const fail = m => { throw new Error(m); };

  const cfg = {
    church: '공덕교회', services: ['1부 예배', '2부 예배', '3부 예배', '오후 예배'],
    categories: ['장애인', '임산부', '노약자', '영유아 동반', '봉사자', '새가족', '일반'],
    capacity: 30, autoApprove: false, daysAhead: 14,
    notice: '데모 모드입니다. 입력한 데이터는 이 기기에만 저장됩니다. (관리자 비밀번호: 1234)',
  };

  function db() {
    let d = store.get();
    if (!d) {
      const s = nextSun();
      d = {
        V: [
          { name: '김은혜', phone: '01012345678', plate: '12가3456', model: '쏘나타', group: '1교구', pin: '1111' },
          { name: '이믿음', phone: '01023456789', plate: '345나6789', model: '카니발', group: '청년부', pin: '2222' },
          { name: '박소망', phone: '01034567890', plate: '56다7890', model: '아반떼', group: '2교구', pin: '3333' },
        ],
        R: [
          { id: id('R'), createdAt: now(), date: s, service: '2부 예배', plate: '12가3456', name: '김은혜', phone: '01012345678', category: '노약자', reason: '부모님 동행', status: '신청' },
          { id: id('R'), createdAt: now(), date: s, service: '2부 예배', plate: '345나6789', name: '이믿음', phone: '01023456789', category: '봉사자', reason: '찬양팀', status: '승인', by: '관리자' },
          { id: id('R'), createdAt: now(), date: s, service: '3부 예배', plate: '56다7890', name: '박소망', phone: '01034567890', category: '임산부', reason: '', status: '신청' },
        ],
      };
      store.set(d);
    }
    return d;
  }
  const save = d => store.set(d);
  const tok = o => btoa(unescape(encodeURIComponent(JSON.stringify(o))));
  const untok = (t, role) => {
    if (!t) fail('로그인이 필요합니다.');
    const o = JSON.parse(decodeURIComponent(escape(atob(t))));
    if (role && o.role !== role) fail('권한이 없습니다.');
    return o;
  };
  const profile = (d, phone) => {
    const mine = d.V.filter(v => v.phone === phone);
    if (!mine.length) fail('회원 정보를 찾을 수 없습니다.');
    return { name: mine[0].name, phone, group: mine[0].group, vehicles: mine.map(v => ({ plate: v.plate, model: v.model })) };
  };
  const mineReq = (d, phone) => d.R.filter(x => x.phone === phone)
    .sort((a, b) => (b.date + b.createdAt).localeCompare(a.date + a.createdAt)).slice(0, 30);
  const plateOk = p => /^([가-힣]{2})?\d{2,3}[가-힣]\d{4}$/.test(p) || fail('차량번호 형식을 확인하세요. 예) 12가3456, 123가4567');
  const ACTIVE = ['신청', '대기', '승인', '입차'];

  const API = {
    config: () => ({ ...cfg, today: today() }),
    register(r) {
      const d = db(), phone = phoneN(r.phone), plate = plateN(r.plate), pin = String(r.pin || '');
      if (!String(r.name || '').trim()) fail('이름을 입력하세요.');
      if (!/^01\d{8,9}$/.test(phone)) fail('휴대폰 번호를 정확히 입력하세요.');
      plateOk(plate);
      if (!/^\d{4,6}$/.test(pin)) fail('PIN은 숫자 4~6자리입니다.');
      if (!r.agree) fail('개인정보 수집·이용에 동의해 주세요.');
      if (d.V.some(v => v.plate === plate)) fail('이미 등록된 차량번호입니다.');
      const mine = d.V.filter(v => v.phone === phone);
      if (mine.length && mine[0].pin !== pin) fail('이미 등록된 연락처입니다. 로그인 후 [내 정보]에서 차량을 추가하세요.');
      d.V.push({ name: r.name.trim(), phone, plate, model: r.model || '', group: r.group || '', pin });
      save(d);
      return { token: tok({ role: 'member', phone }), ...profile(d, phone) };
    },
    login(r) {
      const d = db(), phone = phoneN(r.phone);
      const mine = d.V.filter(v => v.phone === phone);
      if (!mine.length || mine[0].pin !== String(r.pin)) fail('연락처 또는 PIN이 올바르지 않습니다.');
      return { token: tok({ role: 'member', phone }), ...profile(d, phone) };
    },
    me: r => profile(db(), untok(r.token, 'member').phone),
    addVehicle(r) {
      const p = untok(r.token, 'member'), d = db(), plate = plateN(r.plate);
      plateOk(plate);
      if (d.V.some(v => v.plate === plate)) fail('이미 등록된 차량번호입니다.');
      const b = d.V.find(v => v.phone === p.phone);
      d.V.push({ ...b, plate, model: r.model || '' });
      save(d); return profile(d, p.phone);
    },
    removeVehicle(r) {
      const p = untok(r.token, 'member'), d = db();
      if (d.V.filter(v => v.phone === p.phone).length <= 1) fail('마지막 차량은 삭제할 수 없습니다.');
      d.V = d.V.filter(v => !(v.phone === p.phone && v.plate === r.plate));
      save(d); return profile(d, p.phone);
    },
    apply(r) {
      const p = untok(r.token, 'member'), d = db(), t = today();
      if (!r.date) fail('예배일을 선택하세요.');
      if (r.date < t) fail('지난 날짜는 신청할 수 없습니다.');
      if (r.date > addDays(t, cfg.daysAhead)) fail(`${cfg.daysAhead}일 이내 날짜만 신청할 수 있습니다.`);
      const v = d.V.find(x => x.phone === p.phone && x.plate === r.plate) || fail('내 차량을 선택하세요.');
      const same = d.R.filter(x => x.date === r.date && x.service === r.service && ACTIVE.includes(x.status));
      if (same.some(x => x.plate === r.plate)) fail('이미 같은 예배에 신청한 차량입니다.');
      d.R.push({ id: id('R'), createdAt: now(), date: r.date, service: r.service, plate: r.plate, name: v.name, phone: p.phone, category: r.category || '일반', reason: r.reason || '', status: '신청' });
      save(d); return { status: '신청', requests: mineReq(d, p.phone) };
    },
    myRequests: r => mineReq(db(), untok(r.token, 'member').phone),
    cancel(r) {
      const p = untok(r.token, 'member'), d = db();
      const x = d.R.find(q => q.id === r.id && q.phone === p.phone) || fail('신청 내역을 찾을 수 없습니다.');
      if (!['신청', '대기', '승인'].includes(x.status)) fail('취소할 수 없는 상태입니다.');
      Object.assign(x, { status: '취소', by: '본인', at: now() });
      save(d); return mineReq(d, p.phone);
    },
    adminLogin(r) {
      if (!String(r.name || '').trim()) fail('관리자 이름을 입력하세요.');
      if (r.password !== '1234') fail('비밀번호가 올바르지 않습니다. (데모: 1234)');
      return { token: tok({ role: 'admin', name: r.name }), name: r.name };
    },
    adminList(r) {
      untok(r.token, 'admin');
      const list = db().R.filter(x => x.date === r.date);
      const summary = cfg.services.map(sv => {
        const xs = list.filter(x => x.service === sv), c = s => xs.filter(x => x.status === s).length;
        return { service: sv, capacity: cfg.capacity, 신청: c('신청'), 대기: c('대기'), 승인: c('승인'), 입차: c('입차') };
      });
      return { date: r.date, summary, list };
    },
    adminSet(r) {
      const a = untok(r.token, 'admin'), d = db();
      const x = d.R.find(q => q.id === r.id) || fail('신청 내역을 찾을 수 없습니다.');
      Object.assign(x, { status: r.status, by: a.name, at: now() });
      if (r.status === '입차') x.inTime = hm();
      save(d); return API.adminList({ token: r.token, date: x.date });
    },
    adminCreate(r) {
      const a = untok(r.token, 'admin'), d = db(), plate = plateN(r.plate);
      plateOk(plate);
      const v = d.V.find(x => x.plate === plate);
      d.R.push({ id: id('R'), createdAt: now(), date: r.date, service: r.service, plate, name: v ? v.name : (r.name || '(미등록)'), phone: v ? v.phone : phoneN(r.phone), category: r.category || '일반', reason: r.reason || '', status: r.status || '승인', by: a.name, at: now(), inTime: r.status === '입차' ? hm() : '' });
      save(d); return API.adminList({ token: r.token, date: r.date });
    },
    adminVehicles(r) {
      untok(r.token, 'admin');
      const q = plateN(r.q);
      if (q.length < 2) fail('2글자 이상 입력하세요.');
      return db().V.filter(v => v.plate.includes(q) || v.name.includes(q) || v.phone.includes(q))
        .map(({ pin, ...v }) => v);
    },
  };

  window.DemoAPI = async function (action, data) {
    await new Promise(r => setTimeout(r, 150));
    const fn = API[action] || fail('알 수 없는 요청입니다.');
    return JSON.parse(JSON.stringify(fn(data || {})));
  };
  window.DemoReset = () => { try { localStorage.removeItem(KEY); } catch (e) { delete mem[KEY]; } };
})();
