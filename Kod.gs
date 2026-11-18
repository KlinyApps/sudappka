/**
 * Obowiązki Sudackich – serwer (Google Apps Script)
 * Dane zapisują się w arkuszu, do którego podpięty jest ten skrypt.
 *
 * ZMIEŃ PIN poniżej na własny (tylko cyfry). Dzieci go nie zobaczą –
 * kod skryptu widzisz tylko Ty.
 */
const PIN_RODZICA = '1987';
const STREFA = 'Europe/Warsaw';
const ARK_POSTEPY = 'Postępy';
const ARK_KONFIG = '_konfiguracja';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('sudappka')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Bramka danych dla wersji na GitHub Pages (aplikacja na telefon).
 * Strona wysyła {fn, args} bez konta Google, więc blokada kont dzieci poniżej 13 lat jej nie dotyczy.
 */
const API_FUNKCJE = () => ({ pobierzStan, sprawdzPin, odhacz, decyzja, zapiszKonfig, odpowiedzQuiz, kup, wplac, sklepRodzic });
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    const mapa = API_FUNKCJE(), fn = Object.prototype.hasOwnProperty.call(mapa, req.fn) ? mapa[req.fn] : null;
    if (!fn) throw new Error('Nieznana funkcja.');
    out = { ok: true, data: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- pomocnicze ---------- */
function dzis_() { return Utilities.formatDate(new Date(), STREFA, 'yyyy-MM-dd'); }
function ss_() {
  if (ss_.c) return ss_.c;
  const aktywny = SpreadsheetApp.getActiveSpreadsheet();
  if (aktywny) return (ss_.c = aktywny);
  // Projekt niepodpięty do arkusza: używamy własnego arkusza, a przy pierwszym uruchomieniu go zakładamy.
  const prop = PropertiesService.getScriptProperties();
  const id = prop.getProperty('ARKUSZ_ID');
  if (id) { try { return (ss_.c = SpreadsheetApp.openById(id)); } catch (e) {} }
  const nowy = SpreadsheetApp.create('Obowiązki Sudackich – dane');
  prop.setProperty('ARKUSZ_ID', nowy.getId());
  return (ss_.c = nowy);
}

/** Uruchom w edytorze, żeby zobaczyć w dzienniku link do arkusza z danymi. */
function pokazArkusz() {
  pobierzStan();
  Logger.log('Arkusz z danymi: ' + ss_().getUrl());
}
function uid_() { return 't' + Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function dzienTyg_(k) { const a = k.split('-').map(Number); return new Date(a[0], a[1] - 1, a[2]).getDay(); }
function pinOk_(pin) { return String(pin || '') === String(PIN_RODZICA); }

const NAGLOWKI = ['Klucz', 'Data', 'Dziecko', 'Zatwierdzone', 'Razem', 'Gwiazdki', 'Komplet', 'ID zatwierdzonych', 'Zmieniono', 'ID oczekujących', 'ID odrzuconych', 'Oczekuje na zatwierdzenie'];

function arkPostepy_() {
  const ss = ss_();
  let sh = ss.getSheetByName(ARK_POSTEPY);
  if (!sh) {
    sh = ss.insertSheet(ARK_POSTEPY, 0);
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
  }
  if (sh.getRange(1, 12).getValue() !== NAGLOWKI[11]) sh.getRange(1, 1, 1, 12).setValues([NAGLOWKI]).setFontWeight('bold');
  return sh;
}

function arkKonfig_() {
  const ss = ss_();
  let sh = ss.getSheetByName(ARK_KONFIG);
  if (!sh) {
    sh = ss.insertSheet(ARK_KONFIG);
    sh.getRange('A1').setValue(JSON.stringify(domyslna_()));
    sh.getRange('A2').setValue('Ten arkusz trzyma plan zadań. Zmieniaj go w aplikacji (tryb rodzica), nie tutaj.');
    sh.hideSheet();
  }
  return sh;
}

function wczytajKonfig_() { return JSON.parse(arkKonfig_().getRange('A1').getValue()); }

function domyslna_() {
  const W = [0, 1, 2, 3, 4, 5, 6], SZK = [1, 2, 3, 4, 5], PLE = [0, 1, 2, 3, 4];
  const T = (e, t, s, d) => ({ id: uid_(), e: e, t: t, s: s, d: d });
  return {
    start: dzis_(),
    kids: [
      { id: 'krystyna', name: 'Krystyna', color: '#ff5d8f' },
      { id: 'michal', name: 'Michał', color: '#2f7cf6' },
      { id: 'anna', name: 'Anna', color: '#ffb703' }
    ],
    tasks: {
      krystyna: [T('🛏️', 'Pościel łóżko', 1, W), T('🦷', 'Umyj zęby rano i wieczorem', 1, W), T('📚', 'Odrób lekcje', 2, SZK), T('🎒', 'Spakuj plecak na jutro', 1, PLE), T('🍽️', 'Pomóż przy naczyniach', 1, W), T('📖', 'Czytanie 20 minut', 2, W), T('🧹', 'Posprzątaj swój pokój', 3, [6])],
      michal: [T('🛏️', 'Pościel łóżko', 1, W), T('🦷', 'Umyj zęby rano i wieczorem', 1, W), T('📚', 'Odrób lekcje', 2, SZK), T('🎒', 'Spakuj plecak na jutro', 1, PLE), T('🗑️', 'Wynieś śmieci', 2, [2, 5]), T('📖', 'Czytanie 20 minut', 2, W), T('🧹', 'Posprzątaj swój pokój', 3, [6])],
      anna: [T('🛏️', 'Pościel łóżko', 1, W), T('🦷', 'Umyj zęby rano i wieczorem', 1, W), T('🎒', 'Spakuj plecak na jutro', 1, PLE), T('👕', 'Odłóż ubrania na miejsce', 1, W), T('🧸', 'Posprzątaj zabawki', 1, W), T('📖', 'Czytanie 15 minut', 2, W), T('🧹', 'Posprzątaj swój pokój', 3, [6])]
    },
    rewards: [
      { id: uid_(), t: 'Wybierasz film na rodzinny wieczór', c: 30 },
      { id: uid_(), t: 'Wyjście na lody', c: 75 },
      { id: uid_(), t: 'Dodatkowa godzina gry w weekend', c: 150 },
      { id: uid_(), t: 'Wycieczka w wybrane miejsce', c: 300 }
    ]
  };
}

function json_(v) { try { const x = JSON.parse(v || '[]'); return Array.isArray(x) ? x : []; } catch (e) { return []; } }

function wczytajPostepy_() {
  const sh = arkPostepy_(), n = sh.getLastRow();
  const out = {};
  if (n < 2) return out;
  sh.getRange(2, 1, n - 1, 11).getValues().forEach(r => {
    const klucz = String(r[0]);
    if (!klucz) return;
    const i = klucz.lastIndexOf('_');
    out[klucz] = { date: klucz.slice(0, i), kid: klucz.slice(i + 1), ids: json_(r[7]), pend: json_(r[9]), odrz: json_(r[10]),
      total: Number(r[4]) || 0, stars: Number(r[5]) || 0, all: r[6] === true || r[6] === 'TAK' };
  });
  return out;
}

/** Znajduje wiersz dnia dziecka (0 = brak) i jego listy. */
function znajdz_(sh, klucz) {
  const n = sh.getLastRow();
  if (n >= 2) {
    const k = sh.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < k.length; i++) if (String(k[i][0]) === klucz) {
      const r = sh.getRange(i + 2, 1, 1, 11).getValues()[0];
      return { wiersz: i + 2, ids: json_(r[7]), pend: json_(r[9]), odrz: json_(r[10]) };
    }
  }
  return { wiersz: 0, ids: [], pend: [], odrz: [] };
}

function zadaniaDnia_(cfg, kid, day) {
  const wd = dzienTyg_(day);
  return (cfg.tasks[kid] || []).filter(x => (x.d || []).indexOf(wd) > -1);
}

/** Zapisuje stan dnia dziecka i zwraca rekord dla strony. */
function zapiszDzien_(sh, cfg, kid, day, st) {
  const kidObj = cfg.kids.find(k => k.id === kid);
  const ts = zadaniaDnia_(cfg, kid, day);
  const jest = id => ts.some(x => x.id === id);
  const ids = st.ids.filter(jest);
  const pend = st.pend.filter(id => jest(id) && ids.indexOf(id) < 0);
  const odrz = st.odrz.filter(id => jest(id) && ids.indexOf(id) < 0 && pend.indexOf(id) < 0);
  const zatw = ts.filter(x => ids.indexOf(x.id) > -1);
  const stars = zatw.reduce((a, x) => a + (Number(x.s) || 1), 0);
  const all = ts.length > 0 && ids.length === ts.length;
  const czeka = ts.filter(x => pend.indexOf(x.id) > -1).map(x => x.t).join(', ');
  const klucz = day + '_' + kid;
  const rekord = [klucz, day, kidObj ? kidObj.name : kid, zatw.map(x => x.t).join(', '), ts.length, stars, all ? 'TAK' : '',
    JSON.stringify(ids), new Date(), JSON.stringify(pend), JSON.stringify(odrz), czeka];
  if (st.wiersz) sh.getRange(st.wiersz, 1, 1, 12).setValues([rekord]);
  else { sh.appendRow(rekord); st.wiersz = sh.getLastRow(); }
  return { key: klucz, rec: { date: day, kid: kid, ids: ids, pend: pend, odrz: odrz, total: ts.length, stars: stars, all: all } };
}

function zLockiem_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- funkcje wywoływane ze strony ---------- */
function pobierzStan() {
  return { cfg: wczytajKonfig_(), done: wczytajPostepy_(), quiz: wczytajQuiz_(), sklep: wczytajSklep_(), today: dzis_() };
}

/* ---------- quiz dnia ---------- */
const ARK_QUIZ2 = 'Quiz dnia';
const QUIZ_N = 16, QUIZ_ZAL = 11; // 11/16 ≈ 65% – ukryty próg zaliczenia (rodzinny bonus)
const BONUS_RODZINNY = 6;
const NA_SIEBIE = 0.7; // najwyżej 70% zarobionych gwiazdek na nagrody osobiste
// zakresy pytań: matematyka 0–5, angielski 6–9, historia 10–12, geografia 13–15
const ZAKRESY = { m: [0, 6], e: [6, 10], h: [10, 13], g: [13, 16] };

function arkQuiz_() {
  const ss = ss_();
  let sh = ss.getSheetByName(ARK_QUIZ2);
  if (!sh) {
    sh = ss.insertSheet(ARK_QUIZ2, 1);
    sh.getRange(1, 1, 1, 13).setValues([['Klucz', 'Data', 'Dziecko', 'Matematyka (z 6)', 'Angielski (z 4)', 'Historia (z 3)', 'Geografia (z 3)', 'Razem (z 16)', 'Gwiazdki', 'Odpowiedzi', 'Poprawne', 'Czas', 'Ukończony']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
  }
  return sh;
}

function quizRek_(r) {
  const k = String(r[0]), i = k.lastIndexOf('_');
  const ans = json_(r[9]), ok = json_(r[10]), done = r[12] === 'TAK';
  return { date: k.slice(0, i), kid: k.slice(i + 1), m: Number(r[3]) || 0, e: Number(r[4]) || 0, h: Number(r[5]) || 0, g: Number(r[6]) || 0,
    s: Number(r[7]) || 0, stars: done ? Number(r[8]) || 0 : 0, ans: ans, ok: ok, done: done, pass: done && Number(r[7]) >= QUIZ_ZAL };
}

function wczytajQuiz_() {
  const sh = arkQuiz_(), n = sh.getLastRow(), out = {};
  if (n < 2) return out;
  sh.getRange(2, 1, n - 1, 13).getValues().forEach(r => { if (String(r[0])) out[String(r[0])] = quizRek_(r); });
  return out;
}

/** Zapisuje jedną odpowiedź. Odpowiedzi nie da się zmienić. Po 16. odpowiedzi liczy wynik. */
function odpowiedzQuiz(kid, day, i, ans, ok) {
  if (day !== dzis_()) throw new Error('Quiz można rozwiązać tylko w swoim dniu.');
  i = Number(i);
  if (!(i >= 0 && i < QUIZ_N)) throw new Error('Złe pytanie.');
  return zLockiem_(() => {
    const cfg = wczytajKonfig_(), kidObj = cfg.kids.find(k => k.id === kid);
    if (!kidObj) throw new Error('Nie ma takiego dziecka.');
    const sh = arkQuiz_(), klucz = day + '_' + kid, n = sh.getLastRow();
    let wiersz = 0, r = null;
    if (n >= 2) {
      const k = sh.getRange(2, 1, n - 1, 1).getValues();
      for (let j = 0; j < k.length; j++) if (String(k[j][0]) === klucz) { wiersz = j + 2; break; }
    }
    if (wiersz) r = sh.getRange(wiersz, 1, 1, 13).getValues()[0];
    const odp = r ? json_(r[9]) : [], pop = r ? json_(r[10]) : [];
    while (odp.length < QUIZ_N) { odp.push(null); pop.push(null); }
    if (odp[i] === null || odp[i] === '') {
      odp[i] = String(ans == null ? '' : ans).slice(0, 40);
      pop[i] = ok === true;
    }
    const ile = (a, b) => pop.slice(a, b).filter(x => x === true).length;
    const m = ile.apply(null, ZAKRESY.m), e = ile.apply(null, ZAKRESY.e), h = ile.apply(null, ZAKRESY.h), g = ile.apply(null, ZAKRESY.g);
    const s = m + e + h + g, done = odp.every(x => x !== null && x !== '');
    const stars = !done ? 0 : s >= 15 ? 3 : s >= 12 ? 2 : s >= 8 ? 1 : 0;
    const wiersz2 = [klucz, day, kidObj.name, m, e, h, g, s, stars, JSON.stringify(odp), JSON.stringify(pop), new Date(), done ? 'TAK' : ''];
    if (wiersz) sh.getRange(wiersz, 1, 1, 13).setValues([wiersz2]); else sh.appendRow(wiersz2);
    return { key: klucz, rec: quizRek_(wiersz2) };
  });
}

/* ---------- sklepik ---------- */
const ARK_SKLEP = 'Sklepik';
const ST = { DO_WYDANIA: 'do wydania', WYDANE: 'wydane', W_SKARBONCE: 'w skarbonce', ZREALIZOWANE: 'zrealizowane', ZWROCONE: 'zwrócone' };

function arkSklep_() {
  const ss = ss_();
  let sh = ss.getSheetByName(ARK_SKLEP);
  if (!sh) {
    sh = ss.insertSheet(ARK_SKLEP, 2);
    sh.getRange(1, 1, 1, 9).setValues([['ID', 'Czas', 'Dziecko', 'Imię', 'Rodzaj', 'ID nagrody', 'Nagroda', 'Gwiazdki', 'Status']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function wczytajSklep_() {
  const sh = arkSklep_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 9).getValues().filter(r => r[0]).map((r, i) => ({
    id: String(r[0]), ts: r[1] instanceof Date ? r[1].getTime() : 0, kid: String(r[2]), typ: String(r[4]), rid: String(r[5]), t: String(r[6]), c: Number(r[7]) || 0, st: String(r[8]), row: i + 2
  }));
}

/** Gwiazdki zarobione łącznie (zadania + quiz + rodzinny bonus) – tak samo jak liczy strona. */
function zarobione_(cfg, done, quiz) {
  const out = {};
  cfg.kids.forEach(k => out[k.id] = 0);
  Object.keys(done).forEach(k => { const r = done[k]; if (out[r.kid] != null) out[r.kid] += r.stars || 0; });
  Object.keys(quiz).forEach(k => { const r = quiz[k]; if (out[r.kid] != null) out[r.kid] += r.stars || 0; });
  const dni = {};
  Object.keys(done).forEach(k => dni[done[k].date] = 1);
  Object.keys(dni).forEach(d => {
    const ok = cfg.kids.every(k => { const r = done[d + '_' + k.id], q = quiz[d + '_' + k.id]; return r && r.all && q && q.pass; });
    if (ok) cfg.kids.forEach(k => out[k.id] += BONUS_RODZINNY);
  });
  return out;
}

/** Dwie kieszenie dziecka: own (70%, na siebie lub do skarbonek) i fam (30%, tylko do skarbonek).
 *  Wpłaty do skarbonek zużywają najpierw kieszeń rodzinną. */
function kieszenie_(cfg, kid, sklep) {
  const z = zarobione_(cfg, wczytajPostepy_(), wczytajQuiz_())[kid] || 0;
  const cap = Math.floor(z * NA_SIEBIE), famPocket = z - cap;
  const moje = sklep.filter(x => x.kid === kid && x.st !== ST.ZWROCONE);
  const zakupy = moje.filter(x => x.typ === 'Zakup').reduce((a, x) => a + x.c, 0);
  const wplaty = moje.filter(x => x.typ === 'Wpłata').reduce((a, x) => a + x.c, 0);
  const zRodz = Math.min(wplaty, famPocket);
  return { own: cap - zakupy - (wplaty - zRodz), fam: famPocket - zRodz };
}

function wpis_(kid, kidObj, rodzaj, nagroda, ile, status) {
  arkSklep_().appendRow(['s' + uid_().slice(1), new Date(), kid, kidObj.name, rodzaj, nagroda.id, nagroda.t, ile, status]);
}

/** Dziecko kupuje nagrodę osobistą – gwiazdki znikają od razu, rodzic potem wydaje nagrodę. */
function kup(kid, rid) {
  return zLockiem_(() => {
    const cfg = wczytajKonfig_(), kidObj = cfg.kids.find(k => k.id === kid);
    const n = (cfg.rewards || []).find(x => x.id === rid);
    if (!kidObj || !n || n.typ === 'rodzinna') throw new Error('Nie ma takiej nagrody.');
    if (kieszenie_(cfg, kid, wczytajSklep_()).own < n.c) throw new Error('Za mało gwiazdek w Twojej kieszeni.');
    wpis_(kid, kidObj, 'Zakup', n, n.c, ST.DO_WYDANIA);
    return wczytajSklep_();
  });
}

/** Dziecko wpłaca gwiazdki do skarbonki nagrody rodzinnej (nie więcej, niż brakuje). */
function wplac(kid, rid, ile) {
  return zLockiem_(() => {
    const cfg = wczytajKonfig_(), kidObj = cfg.kids.find(k => k.id === kid);
    const n = (cfg.rewards || []).find(x => x.id === rid);
    if (!kidObj || !n || n.typ !== 'rodzinna') throw new Error('Nie ma takiej skarbonki.');
    const sklep = wczytajSklep_();
    const zebrane = sklep.filter(x => x.rid === rid && x.st === ST.W_SKARBONCE).reduce((a, x) => a + x.c, 0);
    const kz = kieszenie_(cfg, kid, sklep);
    ile = Math.min(Math.floor(Number(ile) || 0), n.c - zebrane, Math.max(0, kz.own) + kz.fam);
    if (ile <= 0) throw new Error(zebrane >= n.c ? 'Skarbonka jest już pełna.' : 'Za mało gwiazdek.');
    wpis_(kid, kidObj, 'Wpłata', n, ile, ST.W_SKARBONCE);
    return wczytajSklep_();
  });
}

/** Rodzic: zmiana statusu. akcja: 'wydaj' (zakup) | 'zwroc' (zakup) | 'zrealizuj' (skarbonka) | 'zwrocSkarbonke'. */
function sklepRodzic(akcja, id, pin) {
  if (!pinOk_(pin)) throw new Error('Zły PIN rodzica.');
  return zLockiem_(() => {
    const sh = arkSklep_(), sklep = wczytajSklep_();
    const ustaw = (x, st) => sh.getRange(x.row, 9).setValue(st);
    if (akcja === 'wydaj' || akcja === 'zwroc') {
      const x = sklep.find(y => y.id === id && y.st === ST.DO_WYDANIA);
      if (x) ustaw(x, akcja === 'wydaj' ? ST.WYDANE : ST.ZWROCONE);
    } else if (akcja === 'zrealizuj' || akcja === 'zwrocSkarbonke') {
      sklep.filter(y => y.rid === id && y.st === ST.W_SKARBONCE).forEach(x => ustaw(x, akcja === 'zrealizuj' ? ST.ZREALIZOWANE : ST.ZWROCONE));
    }
    return wczytajSklep_();
  });
}


function sprawdzPin(pin) { return pinOk_(pin); }

/**
 * Dziecko (bez PIN-u): zgłasza dzisiejsze zadanie do zatwierdzenia albo cofa zgłoszenie.
 * Rodzic (z PIN-em): od razu zalicza albo odznacza zadanie, także z poprzednich dni.
 */
function odhacz(kid, day, taskId, pin) {
  const t = dzis_(), rodzic = pinOk_(pin);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day) || day > t) throw new Error('Tego dnia nie można jeszcze odhaczać.');
  if (day !== t && !rodzic) throw new Error('Poprzednie dni może poprawić tylko rodzic.');
  return zLockiem_(() => {
    const cfg = wczytajKonfig_();
    if (!cfg.kids.some(k => k.id === kid)) throw new Error('Nie ma takiego dziecka.');
    if (!zadaniaDnia_(cfg, kid, day).some(x => x.id === taskId)) throw new Error('Nie ma takiego zadania na ten dzień.');
    const sh = arkPostepy_(), st = znajdz_(sh, day + '_' + kid);
    const bez = (a, id) => a.filter(x => x !== id);
    st.odrz = bez(st.odrz, taskId);
    if (rodzic) {
      if (st.ids.indexOf(taskId) > -1) st.ids = bez(st.ids, taskId);
      else { st.ids.push(taskId); st.pend = bez(st.pend, taskId); }
    } else {
      if (st.ids.indexOf(taskId) > -1) throw new Error('To zadanie jest już zatwierdzone. Zmienić je może rodzic.');
      st.pend = st.pend.indexOf(taskId) > -1 ? bez(st.pend, taskId) : st.pend.concat(taskId);
    }
    return zapiszDzien_(sh, cfg, kid, day, st);
  });
}

/** Rodzic zatwierdza (ok = true) albo odrzuca zgłoszenia: [{kid, day, taskId, ok}]. */
function decyzja(lista, pin) {
  if (!pinOk_(pin)) throw new Error('Zły PIN rodzica.');
  return zLockiem_(() => {
    const cfg = wczytajKonfig_(), sh = arkPostepy_(), wynik = {}, stany = {};
    (lista || []).forEach(d => {
      const klucz = d.day + '_' + d.kid;
      const st = stany[klucz] || (stany[klucz] = znajdz_(sh, klucz));
      if (st.pend.indexOf(d.taskId) < 0) return;
      st.pend = st.pend.filter(x => x !== d.taskId);
      if (d.ok) st.ids.push(d.taskId); else st.odrz.push(d.taskId);
    });
    Object.keys(stany).forEach(klucz => {
      const i = klucz.lastIndexOf('_');
      const r = zapiszDzien_(sh, cfg, klucz.slice(i + 1), klucz.slice(0, i), stany[klucz]);
      wynik[r.key] = r.rec;
    });
    return wynik;
  });
}

function zapiszKonfig(cfg, pin) {
  if (!pinOk_(pin)) throw new Error('Zły PIN rodzica.');
  if (!cfg || !Array.isArray(cfg.kids) || !cfg.tasks) throw new Error('Niepoprawny plan.');
  const stary = wczytajKonfig_();
  cfg.start = stary.start || dzis_();
  cfg.kids = stary.kids.map(k => {
    const nowy = cfg.kids.find(x => x.id === k.id) || {};
    return { id: k.id, name: String(nowy.name || k.name).slice(0, 30), color: k.color };
  });
  arkKonfig_().getRange('A1').setValue(JSON.stringify(cfg));
  return cfg;
}
