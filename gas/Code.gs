/**
 * 잇다 (it,da) — 사랑방과 목자를 잇는 주일보고 · 백엔드 (Google Apps Script + Google Sheets)
 * AI목양지원센터 제작 · 만든 사람 이세진 목사 · https://mokyangdesk.com
 * © 2026 AI목양지원센터. All rights reserved.
 * 앱 주소: https://shepherdsdesk.github.io/itda/
 *
 * 시트 구성
 *  - 설정       : 교회이름, 앱주소
 *  - 사용자     : 이름 | 역할(사랑방장/목회자/관리자) | 교구 | 사랑방 | 접속코드 | 연락처 | 토큰 | 개인링크 | 마지막접속
 *  - 성도명단   : 교구 | 마을 | 사랑방 | 이름 | 직분 | 연락처 | 비고 | 성도ID  (앱 ‘명단 관리’에서 엑셀 업로드·편성 가능)
 *  - 주일보고   : 한 주일·한 사랑방당 한 줄 (다시 제출하면 덮어씀)
 *  - 기도제목   : 기도제목 한 줄씩 + 부탁(심방전화/기도부탁/심방부탁) + 목회자 확인
 *  - 목회자응답 : 목회자 코멘트, 부탁 확인 기록 (사랑방장 화면에 보임)
 *  - 문자자료   : 상황별 문자·말씀·문자예화 (쉐퍼드 데스크 자료를 붙여 넣으면 앱 문자 작성에 바로 쓰임)
 */

const SH = { CONFIG: '설정', USERS: '사용자', MEMBERS: '성도명단', REPORTS: '주일보고', PRAYERS: '기도제목', RESP: '목회자응답', SMS: '문자자료' };
const HEAD = {
  CONFIG: ['항목', '값', '설명'],
  USERS: ['이름', '역할', '교구', '사랑방', '접속코드', '연락처', '토큰', '개인링크', '마지막접속'],
  MEMBERS: ['교구', '마을', '사랑방', '이름', '직분', '연락처', '비고', '성도ID'],
  REPORTS: ['주일', '교구', '사랑방', '작성자', '재적', '예배출석', '사랑방출석', '새가족수',
    '예배출석명단', '사랑방출석명단', '결석명단', '새가족', '특이사항', '부탁', '제출시각', '수정시각', '기록ID', '원본JSON'],
  PRAYERS: ['주일', '교구', '사랑방', '대상', '기도제목', '부탁', '작성자', '기록ID', '기도ID', '목회자확인', '처리', '처리시각'],
  RESP: ['기록ID', '주일', '교구', '사랑방', '코멘트', '코멘트작성', '코멘트시각', '확인JSON'],
  SMS: ['상황', '제목', '본문', '말씀', '예화']
};
const ROLE = { LEADER: '사랑방장', PASTOR: '목회자', ADMIN: '관리자' };
const REQ = ['심방전화', '기도부탁', '심방부탁'];
const APP_URL = 'https://shepherdsdesk.github.io/itda/';
const REQUEST_URL = 'https://mokyangdesk.com/itda.html#request';
const TRIAL_PARISH = '체험교구';
// 사용키 확인용 값 — AI목양지원센터 키 발급기와 같아야 합니다 (바꾸면 이미 발급한 키가 모두 무효)
const LICENSE_SALT = 'itda·AI목양지원센터·2026';

/* ───────────── 메뉴 · 처음 설정 ───────────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('잇다')
    .addItem('① 처음 설정 (시트 만들기)', 'setup')
    .addItem('② 접속코드·개인 링크 만들기', 'makeLinks')
    .addSeparator()
    .addItem('③ 사용키 넣기', 'enterKey')
    .addItem('④ 체험 admin 코드 새로 만들기', 'rotateTrialCode')
    .addItem('⑤ 체험 데이터(열두 제자) 지우기', 'clearTrial')
    .addSeparator()
    .addItem('⑥ 주일 미제출 알림 메일 켜기·끄기', 'toggleAlert')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  const make = (key, sample) => {
    let s = ss.getSheetByName(SH[key]);
    if (!s) s = ss.insertSheet(SH[key]);
    if (s.getLastRow() === 0) {
      s.getRange(1, 1, 1, HEAD[key].length).setValues([HEAD[key]]).setFontWeight('bold').setBackground('#eef2ff');
      s.setFrozenRows(1);
      if (sample && sample.length) s.getRange(2, 1, sample.length, sample[0].length).setValues(sample);
    }
    return s;
  };
  make('CONFIG', [
    ['교회이름', '', '앱 상단에 표시 (예: 할렐루야교회 5교구)'],
    ['앱주소', APP_URL, '잇다 앱 주소 — 바꾸지 않아도 됩니다'],
    ['서버주소', '', '배포한 웹 앱 URL — 앱에 처음 들어오면 자동으로 채워집니다']
  ]);
  const firstTime = !ss.getSheetByName(SH.USERS);
  const users = make('USERS');
  users.getRange('E:F').setNumberFormat('@');
  const mem = make('MEMBERS');
  mem.getRange('A:H').setNumberFormat('@');
  make('REPORTS').getRange('A:A').setNumberFormat('@');
  make('PRAYERS').getRange('A:A').setNumberFormat('@');
  make('RESP').getRange('B:B').setNumberFormat('@');
  const cfgRows = rows(SH.CONFIG).map(r => r['항목']);
  if (cfgRows.indexOf('사용키') < 0) setConfig('사용키', '');
  if (firstTime) seedTrial();
  make('SMS', [
    ['환우', '(예시) 쉐퍼드 데스크 문안을 이렇게 붙여 넣으세요', '치료받으시는 동안 교구가 함께 기도하고 있습니다. {이름}님 마음에 주님의 평안이 가득하기를 기도합니다.', '여호와는 너를 지키시는 이시라 여호와께서 네 오른쪽에서 네 그늘이 되시나니 (시 121:5)', '📝 (문자예화 칸) 짧은 예화를 넣으면 앱에서 ‘문자예화 첨부’로 붙일 수 있습니다.'],
  ]);
  const adm = rows(SH.USERS).find(x => x['이름'] === 'admin');
  SpreadsheetApp.getUi().alert('잇다 설정 완료!' + (adm ? '\n\n🧪 체험판이 준비됐습니다 — 체험교구 · 열두 제자 · 사랑방 3곳\n   관리자 이름: admin   접속코드: ' + adm['접속코드'] + '\n   (배포 후 메뉴 ② 를 누르면 바로 들어가는 링크가 생깁니다)' : '') + '\n\n1) [사용자] 시트에 관리자(본인)와 교구 목회자만 넣으세요.\n2) 성도 명단·사랑방 편성·사랑방장 지정은 앱의 ‘명단 관리’에서 엑셀 파일을 올려 하면 됩니다.\n3) [문자자료] 시트에 쉐퍼드 데스크 문자·예화를 붙여 넣으면 앱에서 바로 고를 수 있습니다.\n4) 배포 → 새 배포 → 웹 앱(실행: 나 · 액세스: 모든 사용자)으로 배포한 뒤, 메뉴 ② 를 누르면 개인 링크가 만들어집니다.');
}

/** 메뉴 ②: 서버 주소가 없으면 물어보고, 접속코드·개인 링크를 만든 뒤 내 링크를 보여 줌 */
function makeLinks() {
  const ui = SpreadsheetApp.getUi();
  let sid = serverIdOf(getConfig()['서버주소']);
  if (!sid) {
    const r = ui.prompt('웹 앱 주소 넣기', '배포 → 배포 관리에서 복사한 웹 앱 URL을 붙여 넣으세요.\n(https://script.google.com/macros/s/…/exec)', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    sid = serverIdOf(r.getResponseText());
    if (!sid) { ui.alert('웹 앱 URL이 아닌 것 같아요. 다시 해 주세요.'); return; }
    setConfig('서버주소', 'https://script.google.com/macros/s/' + sid + '/exec');
  }
  fillTokens();
  const me = rows(SH.USERS).find(x => x['역할'] === ROLE.ADMIN) || rows(SH.USERS).find(x => x['역할'] === ROLE.PASTOR);
  ui.alert('개인 링크를 만들었습니다.\n\n[사용자] 시트의 ‘개인링크’ 칸을 확인하세요.' + (me ? '\n\n' + me['이름'] + ' 님 링크:\n' + me['개인링크'] + '\n\n이 링크를 휴대폰으로 보내 열면 바로 들어갑니다.' : ''));
}

/** 비어 있는 접속코드(4자리)·토큰·개인 링크를 채움 */
function fillTokens() {
  const s = sh(SH.USERS);
  const last = s.getLastRow();
  if (last < 2) return;
  const cfg = getConfig();
  const base = String(cfg['앱주소'] || APP_URL).trim();
  const sid = serverIdOf(cfg['서버주소']);
  const range = s.getRange(2, 1, last - 1, HEAD.USERS.length);
  const v = range.getValues();
  v.forEach(r => {
    if (!r[0]) return;
    if (!String(r[4]).trim()) r[4] = code4();
    if (!String(r[6]).trim()) r[6] = newToken();
    if (base) r[7] = base.replace(/\/?$/, '/') + '?t=' + r[6] + (sid ? '&s=' + sid : '');
  });
  range.setValues(v);
  flush();
}

/* ───────────── 웹 API ───────────── */

function doGet() { return json({ ok: true, msg: '잇다 주일보고 API가 작동 중입니다.' }); }

function doPost(e) {
  try {
    flush();
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = API[req.action];
    if (!fn) throw new Error('알 수 없는 요청입니다.');
    return json({ ok: true, data: fn(req) });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

const API = {
  login(req) {
    const name = String(req.name || '').trim(), code = String(req.code || '').trim();
    if (!name || !code) throw new Error('이름과 접속코드를 입력해 주세요.');
    const u = rows(SH.USERS).find(x => String(x['이름']) === name && String(x['접속코드']) === code);
    if (!u) throw new Error('이름 또는 접속코드가 맞지 않습니다.');
    licenseGate(u, 'me');
    let token = String(u['토큰'] || '');
    if (!token) { token = newToken(); sh(SH.USERS).getRange(u._row, 7).setValue(token); }
    touch(u);
    return { token };
  },

  ping() { return { ok: true }; },

  /** 앱에서 사용키 넣기 (관리자) */
  setKey(req) {
    const u = auth(req);
    if (u['역할'] !== ROLE.ADMIN) throw new Error('관리자만 할 수 있습니다.');
    const c = getConfig();
    if (req.church != null && String(req.church).trim()) setConfig('교회이름', String(req.church).trim());
    const church = getConfig()['교회이름'];
    const k = String(req.key || '').trim().toUpperCase();
    const L = licenseCheck(k, church);
    if (L.state === 'none' || L.state === 'invalid') throw new Error('사용키가 맞지 않습니다. 교회이름(' + (church || '비어 있음') + ')이 사용 요청 때 적은 이름과 같은지 확인해 주세요.');
    setConfig('사용키', k);
    return { state: L.state, until: L.until || '', daysLeft: L.daysLeft == null ? null : L.daysLeft };
  },

  me(req) {
    const u = auth(req);
    touch(u);
    // 처음 들어올 때 서버주소·앱주소를 설정 시트에 자동 기록 → 이후 개인 링크에 들어감
    try {
      const cfg = getConfig(), sid = serverIdOf(req.server);
      if (sid && !serverIdOf(cfg['서버주소'])) { setConfig('서버주소', 'https://script.google.com/macros/s/' + sid + '/exec'); fillTokens(); }
      if (req.app && /^https:\/\//.test(req.app) && !cfg['앱주소']) setConfig('앱주소', String(req.app));
    } catch (e) { }
    const L = license();
    const out = { user: view(u), church: getConfig()['교회이름'] || '', license: { state: L.state, until: L.until || '', daysLeft: L.daysLeft == null ? null : L.daysLeft, trialData: u['교구'] === TRIAL_PARISH, blocked: !licenseOK(u, L) } };
    if (u['역할'] === ROLE.PASTOR) out.sarangbangs = sarangbangsOf(u['교구']);
    if (u['역할'] === ROLE.ADMIN && !isChurchAdmin(u)) out.sarangbangs = sarangbangsOf(u['교구']);
    if (isChurchAdmin(u)) out.parishes = allParishes();
    out.churchAdmin = isChurchAdmin(u);
    return out;
  },

  /** 보고 작성 화면: 명단 + 이미 낸 보고 + 목회자 응답 */
  form(req) {
    const u = auth(req);
    const t = target(u, req.parish, req.sarangbang);
    const date = checkDate(req.date);
    const members = rows(SH.MEMBERS).filter(m => m['교구'] === t.parish && m['사랑방'] === t.sarangbang)
      .map(m => ({ name: String(m['이름']), title: String(m['직분'] || '') }));
    const r = findReport(date, t.parish, t.sarangbang);
    const pd = new Date(date + 'T00:00:00'); pd.setDate(pd.getDate() - 7);
    const pr = findReport(Utilities.formatDate(pd, tz(), 'yyyy-MM-dd'), t.parish, t.sarangbang);
    const pj = pr ? safeParse(pr['원본JSON']) : null;
    return {
      prev: pj ? { worship: pj.worship || [], group: pj.group || [] } : null,
      parish: t.parish, sarangbang: t.sarangbang, date, members,
      report: r ? safeParse(r['원본JSON']) : null, updatedAt: r ? r['수정시각'] : '',
      resp: r ? respView(findResp(r['기록ID'])) : null
    };
  },

  /** 보고 제출 (같은 주일 다시 내면 수정 — 목회자 코멘트·확인은 그대로 유지) */
  save(req) {
    const u = auth(req);
    const rp = req.report || {};
    const t = target(u, rp.parish, rp.sarangbang);
    const date = checkDate(rp.date);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const members = rows(SH.MEMBERS).filter(m => m['교구'] === t.parish && m['사랑방'] === t.sarangbang).map(m => String(m['이름']));
      const worship = uniq(rp.worship || []);
      const group = uniq(rp.group || []);
      const newcomers = uniq((rp.newcomers || []).map(s => String(s).trim()).filter(Boolean));
      const absent = members.filter(n => worship.indexOf(n) < 0);
      const prayers = (rp.prayers || []).filter(p => p && String(p.text || '').trim()).map(p => ({
        id: String(p.id || Utilities.getUuid().slice(0, 6)),
        who: String(p.who || '사랑방 공동'),
        text: String(p.text).trim(),
        req: REQ.indexOf(p.req) >= 0 ? p.req : ''
      }));
      const clean = { date, parish: t.parish, sarangbang: t.sarangbang, worship, group, newcomers, prayers,
        note: String(rp.note || '').trim(), addNewToMembers: !!rp.addNewToMembers };
      const now = stamp();
      const old = findReport(date, t.parish, t.sarangbang);
      const id = old ? old['기록ID'] : Utilities.getUuid().slice(0, 8);
      const reqs = prayers.filter(p => p.req).map(p => p.who + '(' + p.req + ')');
      const row = [date, t.parish, t.sarangbang, u['이름'], members.length, worship.length, group.length, newcomers.length,
        worship.join(', '), group.join(', '), absent.join(', '), newcomers.join(', '), clean.note, reqs.join(', '),
        old ? old['제출시각'] : now, now, id, JSON.stringify(clean)];
      const rs = sh(SH.REPORTS);
      if (old) rs.getRange(old._row, 1, 1, row.length).setValues([row]); else rs.appendRow(row);
      flush();

      // 기도제목 시트 다시 쓰기 (확인 기록 유지)
      const acks = (respView(findResp(id)) || {}).acks || {};
      ensureHead('PRAYERS');
      const ps = sh(SH.PRAYERS);
      findRowNums(SH.PRAYERS, '기록ID', id).sort((a, b) => b - a).forEach(r => ps.deleteRow(r));
      if (prayers.length) {
        const pv = prayers.map(p => [date, t.parish, t.sarangbang, p.who, p.text, p.req, u['이름'], id, p.id,
          acks[p.id] ? acks[p.id].by + ' ' + acks[p.id].at : '', acks[p.id] && acks[p.id].action || '', acks[p.id] && acks[p.id].actionAt || '']);
        ps.getRange(ps.getLastRow() + 1, 1, pv.length, pv[0].length).setValues(pv);
      }

      if (clean.addNewToMembers && newcomers.length) {
        const am = rows(SH.MEMBERS).find(m => m['교구'] === t.parish && m['사랑방'] === t.sarangbang && m['마을']);
        newcomers.filter(n => members.indexOf(n) < 0).forEach(n => appendObj(SH.MEMBERS, { '교구': t.parish, '마을': am ? am['마을'] : '', '사랑방': t.sarangbang, '이름': n, '비고': '새가족 (' + date + ' 보고에서 추가)', '성도ID': mid() }));
      }
      return { savedAt: now, updated: !!old };
    } finally { lock.releaseLock(); }
  },

  /** 교구 목회자 화면 */
  dashboard(req) {
    const u = auth(req);
    if (u['역할'] === ROLE.LEADER) throw new Error('목회자만 볼 수 있습니다.');
    const parish = isChurchAdmin(u) ? String(req.parish || allParishes()[0] || '') : u['교구'];
    const date = checkDate(req.date);
    const dates = [];
    const d0 = new Date(date + 'T00:00:00');
    for (let i = 7; i >= 0; i--) { const d = new Date(d0); d.setDate(d.getDate() - i * 7); dates.push(Utilities.formatDate(d, tz(), 'yyyy-MM-dd')); }
    const all = rowsOnDates(SH.REPORTS, dates).filter(r => r['교구'] === parish);
    const leaders = rows(SH.USERS).filter(x => x['역할'] === ROLE.LEADER && x['교구'] === parish);
    const members = rows(SH.MEMBERS).filter(m => m['교구'] === parish);
    const resp = rowsOnDates(SH.RESP, [date]);
    const list = sarangbangsOf(parish).map(sb => {
      const r = all.find(x => x['주일'] === date && x['사랑방'] === sb);
      const ms = members.filter(m => m['사랑방'] === sb);
      const am = ms.find(m => m['마을']);
      return {
        sarangbang: sb,
        area: am ? String(am['마을']) : '',
        leaders: leaders.filter(l => l['사랑방'] === sb).map(l => ({ name: l['이름'], phone: String(l['연락처'] || '') })),
        people: ms.map(m => ({ name: String(m['이름']), title: String(m['직분'] || ''), phone: String(m['연락처'] || '') })),
        enrolled: ms.length,
        submitted: !!r,
        writer: r ? r['작성자'] : '',
        updatedAt: r ? r['수정시각'] : '',
        report: r ? safeParse(r['원본JSON']) : null,
        absent: r ? String(r['결석명단'] || '') : '',
        resp: r ? respView(resp.find(x => x['기록ID'] === r['기록ID'])) : null
      };
    });
    const trend = dates.map(dt => {
      const rs = all.filter(x => x['주일'] === dt);
      return { date: dt, worship: sum(rs, '예배출석'), group: sum(rs, '사랑방출석'), count: rs.length };
    });
    // 장기 결석: 최근 제출된 보고에서 3주 이상 연속으로 예배에 빠진 성도
    const longAbsent = [];
    list.forEach(x => {
      const rs = all.filter(r => r['사랑방'] === x.sarangbang && r['주일'] <= date).sort((a, b) => a['주일'] < b['주일'] ? 1 : -1);
      if (rs.length < 3) return;
      const sets = rs.map(r => String(r['예배출석명단'] || '').split(', '));
      x.people.forEach(m => {
        let k = 0;
        for (let i = 0; i < sets.length; i++) { if (sets[i].indexOf(m.name) >= 0) break; k++; }
        if (k >= 3) longAbsent.push({ name: m.name, title: m.title, phone: m.phone, sarangbang: x.sarangbang, area: x.area, weeks: k, more: k === sets.length });
      });
    });
    longAbsent.sort((a, b) => b.weeks - a.weeks);
    return { parish, date, list, trend, longAbsent, pastor: u['이름'] };
  },

  /** 연간 현황: 최근 52주 출석 추이 · 사랑방별 비교 · 새가족 · 부탁 처리 기록 (10분간 기억) */
  yearly(req) {
    const u = auth(req);
    if (u['역할'] === ROLE.LEADER) throw new Error('목회자만 볼 수 있습니다.');
    const parish = isChurchAdmin(u) ? String(req.parish || allParishes()[0] || '') : u['교구'];
    const cache = CacheService.getScriptCache(), ck = 'yr:' + parish;
    if (!req.fresh) { const hit = cache.get(ck); if (hit) return JSON.parse(hit); }
    const t0 = new Date(); t0.setDate(t0.getDate() - t0.getDay() - 7 * 51);
    const from = Utilities.formatDate(t0, tz(), 'yyyy-MM-dd');
    const s = sh(SH.REPORTS), last = s.getLastRow();
    const vals = last >= 2 ? s.getRange(2, 1, last - 1, 8).getValues() : []; // 주일·교구·사랑방·작성자·재적·예배·사랑방·새가족 (긴 칸은 읽지 않음)
    const weeks = {}, bySb = {};
    vals.forEach(r => {
      if (String(r[1]).trim() !== parish) return;
      const d = norm(r[0], '주일'); if (!d || d < from) return;
      const w = weeks[d] = weeks[d] || { date: d, worship: 0, group: 0, newcomers: 0, count: 0, enrolled: 0 };
      w.worship += Number(r[5]) || 0; w.group += Number(r[6]) || 0; w.newcomers += Number(r[7]) || 0; w.enrolled += Number(r[4]) || 0; w.count++;
      const b = bySb[String(r[2])] = bySb[String(r[2])] || { reports: 0, worship: 0, group: 0, newcomers: 0 };
      b.reports++; b.worship += Number(r[5]) || 0; b.group += Number(r[6]) || 0; b.newcomers += Number(r[7]) || 0;
    });
    const members = rows(SH.MEMBERS).filter(m => m['교구'] === parish);
    const nWeeks = Object.keys(weeks).length;
    const sbs = sarangbangsOf(parish).map(sb => {
      const ms = members.filter(m => m['사랑방'] === sb), am = ms.find(m => m['마을']), b = bySb[sb] || { reports: 0, worship: 0, group: 0, newcomers: 0 };
      return { sarangbang: sb, area: am ? String(am['마을']) : '', enrolled: ms.length, reports: b.reports,
        avgWorship: b.reports ? Math.round(b.worship / b.reports * 10) / 10 : 0, avgGroup: b.reports ? Math.round(b.group / b.reports * 10) / 10 : 0, newcomers: b.newcomers };
    });
    // 부탁 처리 기록
    let care = [];
    try {
      care = rows(SH.PRAYERS).filter(p => p['교구'] === parish && p['부탁'] && p['주일'] >= from)
        .map(p => ({ date: p['주일'], sarangbang: p['사랑방'], who: p['대상'], req: p['부탁'], text: String(p['기도제목'] || ''), ack: String(p['목회자확인'] || ''), action: String(p['처리'] || ''), actionAt: String(p['처리시각'] || '') }))
        .sort((a, b) => a.date < b.date ? 1 : -1);
    } catch (e) { }
    const out = { parish, from, nWeeks, weeks: Object.keys(weeks).sort().map(k => weeks[k]), sbs, enrolledTotal: members.length,
      careCount: { total: care.length, ack: care.filter(c => c.ack).length, call: care.filter(c => c.action === '전화함').length, visit: care.filter(c => c.action === '심방함').length, pray: care.filter(c => c.action === '기도함').length },
      care: care.slice(0, 150), at: stamp() };
    try { cache.put(ck, JSON.stringify(out), 600); } catch (e) { }
    return out;
  },

  /** 목회자 코멘트 (사랑방장 화면에 표시) */
  comment(req) {
    const u = auth(req);
    const r = reportForPastor(u, req);
    return withResp(r, o => { o['코멘트'] = String(req.text || '').trim(); o['코멘트작성'] = u['이름']; o['코멘트시각'] = stamp(); });
  },

  /** 부탁(심방전화·기도부탁·심방부탁) 확인 / 확인 취소 */
  ack(req) {
    const u = auth(req);
    const r = reportForPastor(u, req);
    const pid = String(req.prayerId || '');
    if (!pid) throw new Error('기도제목을 찾을 수 없습니다.');
    const res = withResp(r, o => {
      const acks = safeParse(o['확인JSON']) || {};
      const STEPS = ['전화함', '심방함', '기도함'];
      if (req.done === false) delete acks[pid];
      else if (req.step === 'undo') { if (acks[pid]) { delete acks[pid].action; delete acks[pid].actionAt; } }
      else if (STEPS.indexOf(req.step) >= 0) acks[pid] = Object.assign(acks[pid] || { by: u['이름'], at: stamp() }, { action: req.step, actionAt: stamp() });
      else acks[pid] = Object.assign({ by: u['이름'], at: stamp() }, acks[pid] && acks[pid].action ? { action: acks[pid].action, actionAt: acks[pid].actionAt } : {});
      o['확인JSON'] = JSON.stringify(acks);
    });
    ensureHead('PRAYERS');
    // 기도제목 시트에도 표시
    const pr = findRowNums(SH.PRAYERS, '기도ID', pid).map(n => rowObj(SH.PRAYERS, n)).find(x => x['기록ID'] === r['기록ID']);
    if (pr) { const a = res.acks[pid]; sh(SH.PRAYERS).getRange(pr._row, 10, 1, 3).setValues([[a ? a.by + ' ' + a.at : '', a && a.action || '', a && a.actionAt || '']]); flush(); }
    return res;
  },


  /* ───── 명단 관리 (교구 목회자·관리자) ───── */

  /** 교구 명단 + 사랑방장 + 링크 정보 */
  roster(req) {
    const u = auth(req);
    const parish = pastorParish(u, req.parish);
    ensureMemberIds();
    const members = rows(SH.MEMBERS).filter(m => m['교구'] === parish).map(memberView);
    const leaders = rows(SH.USERS).filter(x => x['역할'] === ROLE.LEADER && x['교구'] === parish)
      .map(x => ({ name: x['이름'], sarangbang: x['사랑방'], phone: String(x['연락처'] || ''), code: String(x['접속코드'] || ''), token: String(x['토큰'] || ''), lastSeen: x['마지막접속'] || '' }));
    return { parish, members, leaders, parishes: isChurchAdmin(u) ? allParishes() : null };
  },

  /** 성도 한 명 추가·수정·삭제 + 사랑방장 지정 */
  memberSave(req) {
    const u = auth(req);
    const parish = pastorParish(u, req.parish);
    if (parish !== TRIAL_PARISH) needLicense('명단 고치기');
    else if (license().state === 'expired') needLicense('명단 고치기');
    const m = req.member || {};
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      ensureMemberIds();
      const all = rows(SH.MEMBERS);
      const cur = m.id ? all.find(x => String(x['성도ID']) === String(m.id) && x['교구'] === parish) : null;
      if (req.remove) {
        if (cur) { sh(SH.MEMBERS).deleteRow(cur._row); flush(); syncLeader(parish, cur['이름'], null); }
        return { ok: true };
      }
      const name = String(m.name || '').trim(), sb = String(m.sarangbang || '').trim();
      if (!name || !sb) throw new Error('이름과 사랑방은 꼭 적어 주세요.');
      const obj = { '교구': parish, '마을': String(m.area != null ? m.area : (cur ? cur['마을'] : '') || '').trim(), '사랑방': sb, '이름': name, '직분': String(m.title || '').trim(), '연락처': phoneFmt(m.phone), '성도ID': cur ? cur['성도ID'] : mid() };
      if (cur) updateObj(SH.MEMBERS, cur._row, obj); else appendObj(SH.MEMBERS, obj);
      syncLeader(parish, cur ? cur['이름'] : name, m.leader ? { name, sarangbang: sb, phone: obj['연락처'] } : null);
      return { ok: true, id: obj['성도ID'] };
    } finally { lock.releaseLock(); }
  },

  /** 엑셀·CSV 명단 한꺼번에 반영
   *  mode: merge = 추가·갱신, replace = 교구 명단 교체
   *  rows: 사랑방원·사랑방장 / pastors: 그룹직책이 교구목사인 사람 → 그 교구 목회자 계정
   *  all: 관리자가 교회 전체 파일을 올릴 때 파일 안의 모든 교구를 한 번에 반영 */
  rosterImport(req) {
    const u = auth(req);
    needLicense('명단 파일 올리기');
    const all = !!req.all && isChurchAdmin(u);
    const base = all ? '' : pastorParish(u, req.parish);
    const pick = p => all ? parishNorm(p) : base;
    ensureMemberIds();
    const list = (req.rows || []).map(r => ({ parish: pick(r.parish), area: String(r.area || '').trim(), sarangbang: String(r.sarangbang || '').trim(), name: String(r.name || '').trim(), title: String(r.title || '').trim(), phone: phoneFmt(r.phone), leader: !!r.leader }))
      .filter(r => r.name && r.sarangbang && r.parish);
    const pastors = (req.pastors || []).map(p => ({ parish: pick(p.parish), name: String(p.name || '').trim(), phone: phoneFmt(p.phone) })).filter(p => p.name && p.parish);
    if (!list.length && !pastors.length) throw new Error('반영할 명단이 없습니다.');
    const lock = LockService.getScriptLock(); lock.waitLock(30000);
    try {
      const out = { added: 0, updated: 0, leaders: 0, pastors: 0, parishes: [] };
      const byParish = {};
      list.forEach(r => (byParish[r.parish] = byParish[r.parish] || []).push(r));
      Object.keys(byParish).sort(koSort).forEach(pa => {
        const r = importParish(pa, byParish[pa], req.mode === 'replace' ? 'replace' : 'merge');
        out.added += r.added; out.updated += r.updated; out.leaders += r.leaders; out.parishes.push(pa);
      });
      pastors.forEach(p => {
        const cur = rows(SH.USERS).find(x => x['이름'] === p.name && (x['역할'] === ROLE.PASTOR || x['역할'] === ROLE.ADMIN));
        if (cur && cur['역할'] === ROLE.ADMIN) return; // 관리자는 그대로 둠
        const obj = { '이름': p.name, '역할': ROLE.PASTOR, '교구': p.parish, '사랑방': '', '연락처': p.phone || (cur ? cur['연락처'] : '') };
        if (cur) { forgetToken(cur['토큰']); updateObj(SH.USERS, cur._row, obj); } else appendObj(SH.USERS, Object.assign(obj, { '접속코드': code4(), '토큰': newToken() }));
        out.pastors++;
        if (out.parishes.indexOf(p.parish) < 0) out.parishes.push(p.parish);
      });
      return out;
    } finally { lock.releaseLock(); }
  },

  /** 사랑방장 링크 새로 만들기 (예전 링크는 막힘) */
  resetLink(req) {
    const u = auth(req);
    const parish = pastorParish(u, req.parish);
    const x = rows(SH.USERS).find(r => r['교구'] === parish && r['이름'] === req.name && r['역할'] !== ROLE.ADMIN);
    if (!x) throw new Error('사용자를 찾을 수 없습니다.');
    const token = newToken();
    forgetToken(x['토큰']);
    sh(SH.USERS).getRange(x._row, 7).setValue(token);
    flush();
    return { token };
  },

  /* ───── 교역자 배정 (관리자) ───── */
  staff(req) {
    const u = auth(req);
    if (u['역할'] !== ROLE.ADMIN) throw new Error('관리자만 할 수 있습니다.');
    const only = isChurchAdmin(u) ? '' : u['교구'];
    return { lockedParish: only, list: rows(SH.USERS).filter(x => (x['역할'] === ROLE.PASTOR || x['역할'] === ROLE.ADMIN) && (!only || x['교구'] === only))
      .map(x => ({ name: x['이름'], role: x['역할'], parish: x['교구'], phone: String(x['연락처'] || ''), code: String(x['접속코드'] || ''), token: String(x['토큰'] || ''), lastSeen: x['마지막접속'] || '' })), parishes: only ? [only] : allParishes() };
  },

  staffSave(req) {
    const u = auth(req);
    if (u['역할'] !== ROLE.ADMIN) throw new Error('관리자만 할 수 있습니다.');
    needLicense('교역자 배정');
    const p = req.person || {};
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      if (!isChurchAdmin(u)) p.parish = u['교구']; // 교구 관리자는 자기 교구 교역자만
      const cur = rows(SH.USERS).find(x => x['이름'] === (p.oldName || p.name) && (x['역할'] === ROLE.PASTOR || x['역할'] === ROLE.ADMIN) && (isChurchAdmin(u) || x['교구'] === u['교구']));
      if (req.remove) { if (cur && cur['토큰'] !== u['토큰']) { forgetToken(cur['토큰']); sh(SH.USERS).deleteRow(cur._row); flush(); } return { ok: true }; }
      if (!p.name || !p.parish) throw new Error('이름과 교구를 적어 주세요.');
      const obj = { '이름': String(p.name).trim(), '역할': p.role === ROLE.ADMIN ? ROLE.ADMIN : ROLE.PASTOR, '교구': String(p.parish).trim(), '사랑방': '', '연락처': phoneFmt(p.phone) };
      if (cur) { forgetToken(cur['토큰']); updateObj(SH.USERS, cur._row, obj); }
      else appendObj(SH.USERS, Object.assign(obj, { '접속코드': code4(), '토큰': newToken() }));
      return { ok: true };
    } finally { lock.releaseLock(); }
  },

  /** 문자자료 시트 (쉐퍼드 데스크 자료 붙여 넣는 곳) */
  sms(req) {
    auth(req);
    let list = [];
    try {
      list = rows(SH.SMS).filter(r => r['본문']).map(r => ({
        situation: String(r['상황'] || '일반'), title: String(r['제목'] || ''), body: String(r['본문'] || ''),
        verse: String(r['말씀'] || ''), illus: String(r['예화'] || '')
      }));
    } catch (e) { }
    return { list };
  }
};



/* ════════════ 사용키 (AI목양지원센터 사용 허락) ════════════
 * 형식: ITDA-T260110-1A2B3C4D  (T=체험판·F=정식, 뒤 6자리=사용 마감 yymmdd, 정식 무기한은 999999)
 * 키는 [설정] 교회이름에 묶여 있어 교회이름을 바꾸면 다시 받아야 합니다.
 * 사용키가 없으면 체험 데이터로 둘러보기만 할 수 있고, 실제 명단 올리기·교역자 배정은 막힙니다. */
function licenseCheck(key, church) {
  key = String(key || '').trim().toUpperCase();
  if (!key) return { state: 'none' };
  const m = key.match(/^ITDA-([TF])(\d{6})-([0-9A-F]{8})$/);
  if (!m) return { state: 'invalid' };
  const sig = sha8(LICENSE_SALT + '|' + m[1] + '|' + m[2] + '|' + normChurch(church));
  if (sig !== m[3]) return { state: 'invalid' };
  const type = m[1] === 'F' ? 'full' : 'trial';
  if (m[2] === '999999') return { state: type, until: '', daysLeft: null };
  const until = '20' + m[2].slice(0, 2) + '-' + m[2].slice(2, 4) + '-' + m[2].slice(4, 6);
  const today = Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd');
  const days = Math.round((new Date(until + 'T00:00:00') - new Date(today + 'T00:00:00')) / 864e5);
  return { state: days < 0 ? 'expired' : type, until, daysLeft: days };
}
function license() { const c = getConfig(); return licenseCheck(c['사용키'], c['교회이름']); }
function normChurch(c) { return String(c || '').replace(/\s+/g, ''); }
function sha8(t) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t, Utilities.Charset.UTF_8)
    .slice(0, 4).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('').toUpperCase();
}
/** 사용 허락 문 — 체험판 7일이 지나거나 키가 없으면(체험교구 제외) 관리자의 '사용키 넣기' 말고는 모두 멈춤 */
function licenseOK(u, L) {
  L = L || license();
  return L.state === 'trial' || L.state === 'full' || (L.state === 'none' && u['교구'] === TRIAL_PARISH);
}
function licenseGate(u, action) {
  const L = license();
  if (licenseOK(u, L)) return;
  if (u['역할'] === ROLE.ADMIN && ['me', 'setKey', 'ping'].indexOf(action) >= 0) return;
  const why = L.state === 'expired' ? '7일 체험 기간(' + L.until + '까지)이 끝났습니다.'
    : L.state === 'invalid' ? '사용키가 교회이름과 맞지 않습니다.' : '아직 사용키가 등록되지 않았습니다.';
  throw new Error('LICENSE_STOP|' + why + (u['역할'] === ROLE.ADMIN ? ' 정식 사용키를 넣으면 바로 다시 쓸 수 있어요.' : ' 교구 관리자(목회자)에게 알려 주세요. 사용키가 등록되면 그동안의 기록 그대로 다시 열립니다.'));
}

/** 실제 명단을 다루는 일은 사용키가 있어야 함 */
function needLicense(what) {
  const L = license();
  if (L.state === 'trial' || L.state === 'full') return;
  const why = L.state === 'expired' ? '사용 기간(' + L.until + ')이 끝났습니다.' : L.state === 'invalid' ? '사용키가 [설정] 교회이름과 맞지 않습니다.' : '체험판에서는 예시 명단(열두 제자)으로만 둘러볼 수 있습니다.';
  throw new Error(why + ' 「' + what + '」 기능은 사용키를 넣은 뒤 쓸 수 있어요. (사용 요청: mokyangdesk.com/itda.html)');
}

function enterKey() {
  const ui = SpreadsheetApp.getUi();
  const c = getConfig();
  if (!String(c['교회이름'] || '').trim()) { ui.alert('먼저 [설정] 탭의 교회이름을 적어 주세요.\n(사용키는 교회이름에 맞춰 발급됩니다 — 사용 요청 때 적은 이름과 똑같이)'); return; }
  const r = ui.prompt('사용키 넣기', 'AI목양지원센터에서 받은 사용키를 붙여 넣으세요.\n(예: ITDA-T270110-1A2B3C4D)\n\n교회이름: ' + c['교회이름'], ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const k = r.getResponseText().trim().toUpperCase();
  const L = licenseCheck(k, c['교회이름']);
  if (L.state === 'invalid' || L.state === 'none') { ui.alert('사용키가 맞지 않습니다.\n[설정] 교회이름이 사용 요청 때 적은 이름과 같은지 확인해 주세요.'); return; }
  setConfig('사용키', k);
  ui.alert(L.state === 'full' ? '정식 사용키가 등록됐습니다. 감사합니다 🙏' : L.state === 'trial' ? '체험판 사용키가 등록됐습니다. ' + L.until + '까지 (' + L.daysLeft + '일 남음) 모든 기능을 쓸 수 있습니다.' : '이미 기간이 끝난 키입니다 (' + L.until + ').');
}

/* ════════════ 체험 데이터 — 체험교구 · 열두 제자 ════════════ */
function seedTrial() {
  const P = TRIAL_PARISH;
  appendObj(SH.USERS, { '이름': 'admin', '역할': ROLE.ADMIN, '교구': P, '접속코드': code6(), '토큰': newToken() });
  const D = [ // 마을, 사랑방, 이름, 직분, 사랑방장
    ['갈릴리', '갈릴리1', '베드로', '집사', 1], ['갈릴리', '갈릴리1', '안드레', '집사', 0], ['갈릴리', '갈릴리1', '야고보', '성도', 0], ['갈릴리', '갈릴리1', '요한', '성도', 0],
    ['갈릴리', '갈릴리2', '빌립', '권사', 1], ['갈릴리', '갈릴리2', '바돌로매', '집사', 0], ['갈릴리', '갈릴리2', '도마', '성도', 0], ['갈릴리', '갈릴리2', '마태', '집사', 0],
    ['예루살렘', '예루살렘1', '작은야고보', '집사', 1], ['예루살렘', '예루살렘1', '다대오', '성도', 0], ['예루살렘', '예루살렘1', '시몬', '성도', 0], ['예루살렘', '예루살렘1', '맛디아', '새가족', 0]
  ];
  ensureMemberIds();
  D.forEach(d => {
    appendObj(SH.MEMBERS, { '교구': P, '마을': d[0], '사랑방': d[1], '이름': d[2], '직분': d[3], '비고': '체험 데이터', '성도ID': mid() });
    if (d[4]) appendObj(SH.USERS, { '이름': d[2], '역할': ROLE.LEADER, '교구': P, '사랑방': d[1], '접속코드': code4(), '토큰': newToken() });
  });
  // 지난 3주 예시 보고
  const ps = [['베드로', '어머니 건강 회복을 위해', ''], ['도마', '직장 이직 문제로 기도 부탁드립니다', '기도부탁'], ['맛디아', '처음 교회에 나왔어요. 잘 정착하도록', '심방전화']];
  for (let w = 1; w <= 3; w++) {
    const d = new Date(); d.setDate(d.getDate() - d.getDay() - 7 * (w - 1));
    const date = Utilities.formatDate(d, tz(), 'yyyy-MM-dd');
    ['갈릴리1', '갈릴리2', '예루살렘1'].forEach((sb, i) => {
      if (w === 1 && i === 2) return; // 이번 주 예루살렘1은 미제출
      const ms = D.filter(x => x[1] === sb).map(x => x[2]);
      const worship = ms.slice(0, ms.length - ((w + i) % 2)), group = ms.slice(0, ms.length - 1);
      const pr = w === 1 ? ps.filter(p => ms.indexOf(p[0]) >= 0).map((p, k) => ({ id: 'tr' + i + k, who: p[0], text: p[1], req: p[2] })) : [];
      const clean = { date, parish: P, sarangbang: sb, worship, group, newcomers: [], prayers: pr, note: '', addNewToMembers: false };
      const id = Utilities.getUuid().slice(0, 8), now = date + ' 14:00';
      appendObj(SH.REPORTS, { '주일': date, '교구': P, '사랑방': sb, '작성자': D.find(x => x[1] === sb && x[4])[2], '재적': ms.length, '예배출석': worship.length, '사랑방출석': group.length, '새가족수': 0,
        '예배출석명단': worship.join(', '), '사랑방출석명단': group.join(', '), '결석명단': ms.filter(n => worship.indexOf(n) < 0).join(', '), '새가족': '', '특이사항': '', '부탁': pr.filter(p => p.req).map(p => p.who + '(' + p.req + ')').join(', '),
        '제출시각': now, '수정시각': now, '기록ID': id, '원본JSON': JSON.stringify(clean) });
      pr.forEach(p => appendObj(SH.PRAYERS, { '주일': date, '교구': P, '사랑방': sb, '대상': p.who, '기도제목': p.text, '부탁': p.req, '작성자': '', '기록ID': id, '기도ID': p.id }));
    });
  }
}

/** 체험 admin 접속코드를 새로 바꿈 (예전 코드·링크는 막힘) */
function rotateTrialCode() {
  const u = rows(SH.USERS).find(x => x['이름'] === 'admin' && x['역할'] === ROLE.ADMIN);
  if (!u) { SpreadsheetApp.getUi().alert('체험 admin 계정이 없습니다.'); return; }
  forgetToken(u['토큰']);
  updateObj(SH.USERS, u._row, { '접속코드': code6(), '토큰': newToken() });
  fillTokens();
  const n = rows(SH.USERS).find(x => x['이름'] === 'admin');
  SpreadsheetApp.getUi().alert('체험 admin 코드를 새로 만들었습니다.\n\n이름: admin   접속코드: ' + n['접속코드'] + (n['개인링크'] ? '\n링크: ' + n['개인링크'] : ''));
}

/** 체험교구 데이터를 모두 지움 — 실제 사용을 시작할 때 */
function clearTrial() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('체험 데이터 지우기', '체험교구(열두 제자)의 계정·명단·보고·기도제목을 모두 지울까요?\n실제 교구 데이터는 그대로 남습니다.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  [SH.USERS, SH.MEMBERS, SH.REPORTS, SH.PRAYERS, SH.RESP].forEach(name => {
    const s = sh(name);
    rows(name).filter(r => r['교구'] === TRIAL_PARISH).map(r => r._row).sort((a, b) => b - a).forEach(r => s.deleteRow(r));
    flush();
  });
  ui.alert('체험 데이터를 지웠습니다. [사용자] 시트에 본인을 관리자로 넣고 메뉴 ② 를 눌러 주세요.');
}
function code6() { return String(Math.floor(100000 + Math.random() * 900000)); }

/** 예전에 만든 시트에 새로 생긴 칸 제목을 끝에 붙임 */
function ensureHead(key) {
  const k = 'head#' + key; if (MEMO[k]) return; MEMO[k] = 1;
  const s = sh(SH[key]);
  const head = s.getRange(1, 1, 1, Math.max(1, s.getLastColumn())).getValues()[0];
  HEAD[key].forEach((h, i) => { if (head[i] !== h && head.indexOf(h) < 0 && !head[i]) s.getRange(1, i + 1).setValue(h).setFontWeight('bold').setBackground('#eef2ff'); });
}

/* ════════════ 주일 미제출 알림 메일 ════════════ */
function toggleAlert() {
  const ui = SpreadsheetApp.getUi();
  const old = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sendMissingMail');
  if (old.length) {
    if (ui.alert('미제출 알림', '지금 알림이 켜져 있습니다. 끌까요?\n(아니요 = 시각·메일 주소 다시 정하기)', ui.ButtonSet.YES_NO) === ui.Button.YES) { old.forEach(t => ScriptApp.deleteTrigger(t)); ui.alert('알림을 껐습니다.'); return; }
    old.forEach(t => ScriptApp.deleteTrigger(t));
  }
  const c = getConfig();
  const h = ui.prompt('미제출 알림 — 시각', '주일 몇 시에 보낼까요? (0~23, 예: 18 = 저녁 6시)', ui.ButtonSet.OK_CANCEL);
  if (h.getSelectedButton() !== ui.Button.OK) return;
  const hour = Math.min(23, Math.max(0, parseInt(h.getResponseText() || c['알림시각'] || '18', 10) || 18));
  const m = ui.prompt('미제출 알림 — 받는 메일', '받는 메일 주소 (여러 명은 쉼표로)\n비워 두면: ' + (c['알림메일'] || Session.getEffectiveUser().getEmail()), ui.ButtonSet.OK_CANCEL);
  if (m.getSelectedButton() !== ui.Button.OK) return;
  const mail = String(m.getResponseText() || c['알림메일'] || Session.getEffectiveUser().getEmail()).trim();
  setConfig('알림시각', hour); setConfig('알림메일', mail);
  ScriptApp.newTrigger('sendMissingMail').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(hour).create();
  ui.alert('알림을 켰습니다.\n매주 주일 ' + hour + '시쯤, 보고를 내지 않은 사랑방 목록을 ' + mail + ' 로 보냅니다.');
}

function sendMissingMail() {
  flush();
  const L = license();
  if (L.state !== 'full' && L.state !== 'trial') return;
  const c = getConfig(), to = String(c['알림메일'] || '').trim();
  if (!to) return;
  const d = new Date(); d.setDate(d.getDate() - d.getDay());
  const date = Utilities.formatDate(d, tz(), 'yyyy-MM-dd');
  const reps = rowsOnDates(SH.REPORTS, [date]);
  const leaders = rows(SH.USERS).filter(x => x['역할'] === ROLE.LEADER);
  const lines = []; let miss = 0, total = 0;
  allParishes().filter(p => p !== TRIAL_PARISH).forEach(p => {
    const sbs = sarangbangsOf(p); total += sbs.length;
    const no = sbs.filter(sb => !reps.some(r => r['교구'] === p && r['사랑방'] === sb));
    miss += no.length;
    lines.push('■ ' + p + ' — 제출 ' + (sbs.length - no.length) + '/' + sbs.length);
    no.forEach(sb => { const l = leaders.filter(x => x['교구'] === p && x['사랑방'] === sb).map(x => x['이름'] + (x['연락처'] ? ' ' + x['연락처'] : '')).join(', '); lines.push('  · ' + sb + (l ? ' (' + l + ')' : '')); });
  });
  if (!total) return;
  const app = String(c['앱주소'] || APP_URL), sid = serverIdOf(c['서버주소']);
  MailApp.sendEmail(to, '[잇다] ' + date + ' 주일보고 — ' + (miss ? '미제출 ' + miss + '곳' : '모두 제출 🙏'),
    (c['교회이름'] || '') + ' ' + date + ' 주일보고 현황입니다.\n\n' + lines.join('\n') + '\n\n앱에서 보기: ' + app + (sid ? '?s=' + sid : '') + '\n\n— 잇다 (AI목양지원센터)');
}

/* ───── 명단 도우미 ───── */
/** 교회 전체 관리자 = 역할 '관리자' + 교구 칸이 비어 있음. 교구가 적힌 관리자는 그 교구만 봄 */
function isChurchAdmin(u) { return u['역할'] === ROLE.ADMIN && !String(u['교구'] || '').trim(); }

/** 한 교구 명단 반영 (잠금은 호출하는 쪽에서) */
function importParish(parish, list, mode) {
  const s = sh(SH.MEMBERS);
  const head = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const all = rows(SH.MEMBERS);
  const others = all.filter(x => x['교구'] !== parish);
  const mine = mode === 'replace' ? [] : all.filter(x => x['교구'] === parish);
  let added = 0, updated = 0;
  list.forEach(r => {
    const hit = mine.find(x => x['이름'] === r.name && (!r.phone || !x['연락처'] || x['연락처'] === r.phone));
    if (hit) { Object.assign(hit, { '마을': r.area || hit['마을'] || '', '사랑방': r.sarangbang, '직분': r.title || hit['직분'], '연락처': r.phone || hit['연락처'] }); updated++; }
    else { mine.push({ '교구': parish, '마을': r.area || '', '사랑방': r.sarangbang, '이름': r.name, '직분': r.title, '연락처': r.phone, '비고': '', '성도ID': mid() }); added++; }
  });
  const out = others.concat(mine).map(o => head.map(k => o[k] == null ? '' : o[k]));
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, head.length).clearContent();
  if (out.length) s.getRange(2, 1, out.length, head.length).setValues(out);
  flush();
  const flagged = list.filter(r => r.leader);
  if (flagged.length) {
    if (mode === 'replace') {
      rows(SH.USERS).filter(x => x['역할'] === ROLE.LEADER && x['교구'] === parish && !flagged.some(f => f.name === x['이름']))
        .forEach(x => syncLeader(parish, x['이름'], null));
    }
    flagged.forEach(f => syncLeader(parish, f.name, { name: f.name, sarangbang: f.sarangbang, phone: f.phone }));
  }
  return { added, updated, leaders: flagged.length };
}

function parishNorm(p) { return String(p || '').replace(/\s+/g, ''); }

function pastorParish(u, parish) {
  if (u['역할'] === ROLE.PASTOR) return u['교구'];
  if (u['역할'] === ROLE.ADMIN) {
    if (!isChurchAdmin(u)) return u['교구'];
    if (!parish) throw new Error('교구를 선택해 주세요.'); return String(parish);
  }
  throw new Error('교구 목회자만 할 수 있습니다.');
}

/** 사랑방장 계정 맞추기: info가 있으면 지정(이름·사랑방 갱신, 기존 링크 유지), null이면 해제 */
function syncLeader(parish, name, info) {
  const cur = rows(SH.USERS).find(x => x['역할'] === ROLE.LEADER && x['교구'] === parish && x['이름'] === name);
  if (!info) { if (cur) { forgetToken(cur['토큰']); sh(SH.USERS).deleteRow(cur._row); flush(); } return; }
  const obj = { '이름': info.name, '역할': ROLE.LEADER, '교구': parish, '사랑방': info.sarangbang, '연락처': info.phone || (cur ? cur['연락처'] : '') };
  if (cur) { forgetToken(cur['토큰']); updateObj(SH.USERS, cur._row, obj); }
  else appendObj(SH.USERS, Object.assign(obj, { '접속코드': code4(), '토큰': newToken() }));
}

function memberView(m) {
  return { id: String(m['성도ID']), name: String(m['이름']), title: String(m['직분'] || ''), phone: String(m['연락처'] || ''), sarangbang: String(m['사랑방'] || ''), area: String(m['마을'] || ''), note: String(m['비고'] || '') };
}

function ensureMemberIds() {
  const s = sh(SH.MEMBERS);
  let head = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  if (head.indexOf('마을') < 0) { // 예전 시트: '마을' 칸을 맨 끝에 추가
    s.getRange(1, head.length + 1).setValue('마을').setFontWeight('bold').setBackground('#eef2ff');
    head = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    flush();
  }
  let col = head.indexOf('성도ID') + 1;
  if (!col) { col = head.length + 1; s.getRange(1, col).setValue('성도ID').setFontWeight('bold').setBackground('#eef2ff'); }
  const n = s.getLastRow() - 1;
  if (n < 1) return;
  const r = s.getRange(2, col, n, 1), v = r.getValues();
  let changed = false;
  v.forEach(x => { if (!String(x[0]).trim()) { x[0] = mid(); changed = true; } });
  if (changed) { r.setValues(v); flush(); }
}

function appendObj(name, o) {
  const s = sh(name);
  const head = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  s.appendRow(head.map(k => o[k] == null ? '' : o[k]));
  flush();
}

function updateObj(name, row, o) {
  const s = sh(name);
  const head = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const r = s.getRange(row, 1, 1, head.length), v = r.getValues()[0];
  head.forEach((k, i) => { if (k in o) v[i] = o[k]; });
  r.setValues([v]);
  flush();
}

function phoneFmt(p) {
  const d = String(p || '').replace(/[^0-9]/g, '');
  if (/^01\d{8,9}$/.test(d)) return d.length === 11 ? d.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3') : d.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
  if (/^1\d{9}$/.test(d)) return phoneFmt('0' + d); // 엑셀이 앞자리 0을 지운 경우
  return String(p || '').trim();
}
function code4() { return String(Math.floor(1000 + Math.random() * 9000)); }
function mid() { return 'm' + Utilities.getUuid().replace(/-/g, '').slice(0, 9); }

/* ───────────── 도우미 ───────────── */

/** 로그인 확인 — 10분 동안 기억해 두어 매번 사용자 시트를 읽지 않음 */
function auth(req) {
  const token = String(req.token || '').trim();
  if (!token) throw new Error('로그인이 필요합니다.');
  const cache = CacheService.getScriptCache();
  const hit = cache.get('tok:' + token);
  if (hit) { const c = JSON.parse(hit); licenseGate(c, req.action); return c; }
  const u = rows(SH.USERS).find(x => String(x['토큰']) === token);
  if (!u) throw new Error('로그인 정보가 바뀌었습니다. 다시 로그인해 주세요.');
  const o = {}; HEAD.USERS.forEach(k => o[k] = u[k] instanceof Date ? '' : u[k]);
  try { cache.put('tok:' + token, JSON.stringify(o), 600); } catch (e) { }
  licenseGate(o, req.action);
  return o;
}
function forgetToken(t) { if (t) try { CacheService.getScriptCache().remove('tok:' + t); } catch (e) { } }

function target(u, parish, sarangbang) {
  const role = u['역할'];
  if (role === ROLE.LEADER) return { parish: u['교구'], sarangbang: u['사랑방'] };
  if (role === ROLE.PASTOR) {
    if (!sarangbang) throw new Error('사랑방을 선택해 주세요.');
    return { parish: u['교구'], sarangbang: String(sarangbang) };
  }
  if (role === ROLE.ADMIN) {
    if (!sarangbang) throw new Error('사랑방을 선택해 주세요.');
    if (!isChurchAdmin(u)) return { parish: u['교구'], sarangbang: String(sarangbang) };
    if (!parish) throw new Error('교구를 선택해 주세요.');
    return { parish: String(parish), sarangbang: String(sarangbang) };
  }
  throw new Error('역할이 지정되지 않았습니다. (사랑방장/목회자/관리자)');
}

function reportForPastor(u, req) {
  if (u['역할'] === ROLE.LEADER) throw new Error('목회자만 할 수 있습니다.');
  const t = target(u, req.parish, req.sarangbang);
  const r = findReport(checkDate(req.date), t.parish, t.sarangbang);
  if (!r) throw new Error('아직 제출된 보고가 없습니다.');
  return r;
}

function withResp(r, fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sh(SH.RESP);
    const cur = findResp(r['기록ID']);
    const o = cur || { '기록ID': r['기록ID'], '주일': r['주일'], '교구': r['교구'], '사랑방': r['사랑방'], '코멘트': '', '코멘트작성': '', '코멘트시각': '', '확인JSON': '{}' };
    fn(o);
    const row = HEAD.RESP.map(k => o[k] == null ? '' : o[k]);
    if (cur) s.getRange(cur._row, 1, 1, row.length).setValues([row]); else s.appendRow(row);
    flush();
    return respView(o);
  } finally { lock.releaseLock(); }
}

function findResp(id) { const n = findRowNums(SH.RESP, '기록ID', id); return n.length ? rowObj(SH.RESP, n[0]) : null; }

function respView(o) {
  if (!o) return null;
  return { comment: String(o['코멘트'] || ''), by: String(o['코멘트작성'] || ''), at: String(o['코멘트시각'] || ''), acks: safeParse(o['확인JSON']) || {} };
}

function view(u) { return { name: u['이름'], role: u['역할'], parish: u['교구'], sarangbang: u['사랑방'] }; }

function sarangbangsOf(parish) {
  const set = {};
  rows(SH.USERS).concat(rows(SH.MEMBERS)).forEach(x => { if (x['교구'] === parish && x['사랑방']) set[x['사랑방']] = 1; });
  return Object.keys(set).sort(koSort);
}

function allParishes() {
  const set = {};
  rows(SH.USERS).concat(rows(SH.MEMBERS)).forEach(x => { if (x['교구']) set[x['교구']] = 1; });
  return Object.keys(set).sort(koSort);
}

function findReport(date, parish, sarangbang) {
  const hit = r => r['주일'] === date && r['교구'] === parish && r['사랑방'] === sarangbang;
  return rowsOnDates(SH.REPORTS, [date]).find(hit) || null;
}

function touch(u) {
  try {
    const c = CacheService.getScriptCache(), k = 'seen:' + u['토큰'];
    if (c.get(k)) return; // 30분에 한 번만 기록
    const n = findRowNums(SH.USERS, '토큰', u['토큰']);
    if (n.length) sh(SH.USERS).getRange(n[0], 9).setValue(new Date());
    c.put(k, '1', 1800);
  } catch (e) { }
}

function serverIdOf(v) {
  v = String(v || '').trim();
  const m = v.match(/\/macros\/s\/([\w-]{20,})/);
  return m ? m[1] : (/^[\w-]{20,}$/.test(v) ? v : '');
}

function setConfig(key, value) {
  const s = sh(SH.CONFIG);
  const r = rows(SH.CONFIG).find(x => x['항목'] === key);
  if (r) s.getRange(r._row, 2).setValue(value); else s.appendRow([key, value, '']);
  flush();
}

function getConfig() {
  const o = {};
  try { rows(SH.CONFIG).forEach(r => o[r['항목']] = r['값']); } catch (e) { }
  return o;
}

function sh(name) {
  const s = SpreadsheetApp.getActive().getSheetByName(name);
  if (!s) throw new Error('[' + name + '] 시트가 없습니다. 메뉴 ‘잇다 → 처음 설정’을 실행해 주세요.');
  return s;
}

/* ───── 속도: 한 번의 요청 안에서 같은 시트를 여러 번 읽지 않도록 기억해 둠 ───── */
const MEMO = {};
function flush() { Object.keys(MEMO).forEach(k => delete MEMO[k]); }

function rows(name) {
  if (MEMO[name]) return MEMO[name];
  MEMO[name] = readRows(name);
  return MEMO[name];
}

function readRows(name) {
  const v = sh(name).getDataRange().getValues();
  const h = v.shift() || [];
  const out = [];
  v.forEach((r, i) => {
    if (r.every(c => c === '' || c === null)) return;
    const o = { _row: i + 2 };
    h.forEach((k, j) => o[k] = norm(r[j], k));
    out.push(o);
  });
  return out;
}

function norm(v, key) {
  if (v instanceof Date) return Utilities.formatDate(v, tz(), key === '주일' ? 'yyyy-MM-dd' : 'yyyy-MM-dd HH:mm');
  if (typeof v === 'string') return v.trim();
  if (['접속코드', '사랑방', '교구', '연락처', '기도ID', '기록ID'].indexOf(key) >= 0) return String(v);
  return v;
}

/** 특정 주일들의 기록만 읽기 — '주일' 칸에서 날짜를 바로 찾아 그 범위만 읽음 (시트 전체를 읽지 않음) */
function rowsOnDates(name, dates) {
  dates = dates.filter(Boolean);
  if (!finderOK(name)) return rows(name).filter(o => dates.indexOf(o['주일']) >= 0);
  let nums = [];
  dates.forEach(d => { nums = nums.concat(findRowNums(name, '주일', d)); });
  if (!nums.length) return [];
  return readSpan(name, Math.min.apply(null, nums), Math.max.apply(null, nums)).filter(o => dates.indexOf(o['주일']) >= 0);
}

/** '주일' 칸이 글자로 저장돼 찾기가 되는지 한 번 확인 (날짜 서식으로 바뀌었으면 전체 읽기로) */
function finderOK(name) {
  const k = name + '#finder';
  if (k in MEMO) return MEMO[k];
  const s = sh(name), last = s.getLastRow();
  let ok = true;
  if (last >= 2) {
    const head = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    const col = head.indexOf('주일') + 1;
    const v = s.getRange(last, col).getValues()[0][0];
    const str = v instanceof Date ? Utilities.formatDate(v, tz(), 'yyyy-MM-dd') : String(v);
    ok = !(v instanceof Date) && findRowNums(name, '주일', str).length > 0;
  }
  MEMO[k] = ok;
  return ok;
}

function readSpan(name, a, b) {
  const s = sh(name), w = s.getLastColumn();
  const head = s.getRange(1, 1, 1, w).getValues()[0];
  const out = [];
  s.getRange(a, 1, b - a + 1, w).getValues().forEach((r, i) => {
    if (r.every(c => c === '' || c === null)) return;
    const o = { _row: a + i };
    head.forEach((k, j) => o[k] = norm(r[j], k));
    out.push(o);
  });
  return out;
}

/** 한 칸 값으로 줄 번호 찾기 (시트 전체를 읽지 않음) */
function findRowNums(name, header, value) {
  const s = sh(name);
  const head = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const col = head.indexOf(header) + 1;
  if (!col || s.getLastRow() < 2 || !value) return [];
  return s.getRange(2, col, s.getLastRow() - 1, 1).createTextFinder(String(value)).matchEntireCell(true).findAll().map(r => r.getRow());
}

function rowObj(name, rowNum) {
  const s = sh(name), w = s.getLastColumn();
  const head = s.getRange(1, 1, 1, w).getValues()[0];
  const v = s.getRange(rowNum, 1, 1, w).getValues()[0];
  const o = { _row: rowNum };
  head.forEach((k, j) => o[k] = norm(v[j], k));
  return o;
}

function checkDate(d) {
  d = String(d || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('날짜 형식이 올바르지 않습니다.');
  return d;
}

function stamp() { return Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd HH:mm'); }
function newToken() { return Utilities.getUuid().replace(/-/g, '').slice(0, 20); }
function uniq(a) { return a.map(String).filter((x, i, s) => x && s.indexOf(x) === i); }
function sum(rs, k) { return rs.reduce((a, r) => a + (Number(r[k]) || 0), 0); }
function koSort(a, b) { return a.localeCompare(b, 'ko', { numeric: true }); }
function safeParse(s) { try { return JSON.parse(s); } catch (e) { return null; } }
function tz() { return Session.getScriptTimeZone() || 'Asia/Seoul'; }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
