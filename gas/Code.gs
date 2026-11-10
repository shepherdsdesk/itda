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
  PRAYERS: ['주일', '교구', '사랑방', '대상', '기도제목', '부탁', '작성자', '기록ID', '기도ID', '목회자확인'],
  RESP: ['기록ID', '주일', '교구', '사랑방', '코멘트', '코멘트작성', '코멘트시각', '확인JSON'],
  SMS: ['상황', '제목', '본문', '말씀', '예화']
};
const ROLE = { LEADER: '사랑방장', PASTOR: '목회자', ADMIN: '관리자' };
const REQ = ['심방전화', '기도부탁', '심방부탁'];

/* ───────────── 메뉴 · 처음 설정 ───────────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('잇다')
    .addItem('① 처음 설정 (시트 만들기)', 'setup')
    .addItem('② 접속코드·개인 링크 만들기', 'fillTokens')
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
    ['교회이름', '할렐루야교회', '앱 상단에 표시'],
    ['앱주소', '', 'GitHub Pages 주소 (예: https://shepherdsdesk.github.io/sarangbang/) — 개인 링크를 만들 때 사용']
  ]);
  const users = make('USERS', [
    ['홍길동', ROLE.LEADER, '5교구', '이매1', '', '010-0000-0000', '', '', ''],
    ['이세진', ROLE.PASTOR, '5교구', '', '', '', '', '', '']
  ]);
  users.getRange('E:F').setNumberFormat('@');
  const mem = make('MEMBERS', [
    ['5교구', '이매', '이매1', '홍길동', '집사', '010-0000-0000', '', 'm-sample1'],
    ['5교구', '이매', '이매1', '김은혜', '권사', '', '', 'm-sample2'],
    ['5교구', '이매', '이매1', '박소망', '성도', '', '', 'm-sample3']
  ]);
  mem.getRange('E:E').setNumberFormat('@');
  make('REPORTS').getRange('A:A').setNumberFormat('@');
  make('PRAYERS').getRange('A:A').setNumberFormat('@');
  make('RESP').getRange('B:B').setNumberFormat('@');
  make('SMS', [
    ['환우', '(예시) 쉐퍼드 데스크 문안을 이렇게 붙여 넣으세요', '치료받으시는 동안 교구가 함께 기도하고 있습니다. {이름}님 마음에 주님의 평안이 가득하기를 기도합니다.', '여호와는 너를 지키시는 이시라 여호와께서 네 오른쪽에서 네 그늘이 되시나니 (시 121:5)', '📝 (문자예화 칸) 짧은 예화를 넣으면 앱에서 ‘문자예화 첨부’로 붙일 수 있습니다.'],
  ]);
  fillTokens();
  SpreadsheetApp.getUi().alert('잇다 설정 완료!\n\n1) [사용자] 시트에 관리자(본인)와 교구 목회자만 넣으세요.\n2) 성도 명단·사랑방 편성·사랑방장 지정은 앱의 ‘명단 관리’에서 엑셀 파일을 올려 하면 됩니다.\n3) [문자자료] 시트에 쉐퍼드 데스크 문자·예화를 붙여 넣으면 앱에서 바로 고를 수 있습니다.\n4) 메뉴 ② 를 누르면 접속코드와 개인 링크가 만들어집니다.');
}

/** 비어 있는 접속코드(4자리)·토큰·개인 링크를 채움 */
function fillTokens() {
  const s = sh(SH.USERS);
  const last = s.getLastRow();
  if (last < 2) return;
  const base = String(getConfig()['앱주소'] || '').trim();
  const range = s.getRange(2, 1, last - 1, HEAD.USERS.length);
  const v = range.getValues();
  v.forEach(r => {
    if (!r[0]) return;
    if (!String(r[4]).trim()) r[4] = code4();
    if (!String(r[6]).trim()) r[6] = newToken();
    if (base) r[7] = base.replace(/\/?$/, '/') + '?t=' + r[6];
  });
  range.setValues(v);
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
    let token = String(u['토큰'] || '');
    if (!token) { token = newToken(); sh(SH.USERS).getRange(u._row, 7).setValue(token); }
    touch(u);
    return { token };
  },

  me(req) {
    const u = auth(req);
    touch(u);
    const out = { user: view(u), church: getConfig()['교회이름'] || '' };
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
    return {
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
      const ps = sh(SH.PRAYERS);
      findRowNums(SH.PRAYERS, '기록ID', id).sort((a, b) => b - a).forEach(r => ps.deleteRow(r));
      if (prayers.length) {
        const pv = prayers.map(p => [date, t.parish, t.sarangbang, p.who, p.text, p.req, u['이름'], id, p.id,
          acks[p.id] ? acks[p.id].by + ' ' + acks[p.id].at : '']);
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
    return { parish, date, list, trend, pastor: u['이름'] };
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
      if (req.done === false) delete acks[pid]; else acks[pid] = { by: u['이름'], at: stamp() };
      o['확인JSON'] = JSON.stringify(acks);
    });
    // 기도제목 시트에도 표시
    const pr = findRowNums(SH.PRAYERS, '기도ID', pid).map(n => rowObj(SH.PRAYERS, n)).find(x => x['기록ID'] === r['기록ID']);
    if (pr) { const a = res.acks[pid]; sh(SH.PRAYERS).getRange(pr._row, 10).setValue(a ? a.by + ' ' + a.at : ''); }
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
  if (hit) return JSON.parse(hit);
  const u = rows(SH.USERS).find(x => String(x['토큰']) === token);
  if (!u) throw new Error('로그인 정보가 바뀌었습니다. 다시 로그인해 주세요.');
  const o = {}; HEAD.USERS.forEach(k => o[k] = u[k] instanceof Date ? '' : u[k]);
  try { cache.put('tok:' + token, JSON.stringify(o), 600); } catch (e) { }
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
