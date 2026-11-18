/**
 * Plan zadań. Uruchom funkcję ustawPlan jeden raz z edytora.
 * Nagrody zostają bez zmian. Strona dzieci pokaże nowy plan po odświeżeniu
 * (najpóźniej po 30 s). Wdrażanie nowej wersji nie jest potrzebne.
 *
 * Gwiazdki są ułożone tak, żeby każdego dnia każde dziecko mogło zdobyć
 * tyle samo: pn 11, wt 11, śr 11, cz 14, pt 9, so 9, nd 14 (bez quizu i bonusu).
 * Wyrównujemy modlitwą i ćwiczeniami za 2★ w wybrane dni.
 */
function ustawPlan() {
  const CODZ = [0, 1, 2, 3, 4, 5, 6], ND_CZW = [0, 1, 2, 3, 4], PN_PT = [1, 2, 3, 4, 5];
  const PN = 1, WT = 2, SR = 3, CZW = 4, PT = 5, SO = 6, ND = 0;
  const T = (e, t, s, d) => ({ id: uid_(), e: e, t: t, s: s, d: d });
  // zadanie za 2★ w podane dni i za 1★ w pozostałe
  const dwa = (e, t, dni2) => [T(e, t, 2, dni2), T(e, t, 1, CODZ.filter(d => dni2.indexOf(d) < 0))].filter(x => x.d.length);
  const STOLY = 'Czyste stoły i wyspa na parterze';

  const wspolne = (modl2, cw2, porzadek) => [].concat(
    [T('🛏️', 'Pościel łóżko', 1, CODZ),
     T('🦷', 'Zęby umyte rano i wieczorem', 1, CODZ)],
    dwa('🙏', 'Modlitwa', modl2),
    [T('🧺', 'Porządek w pokoju + zabrane pranie', porzadek, CODZ)],
    dwa('💪', 'Ćwiczenia/Brzuszki/Pompki', cw2),
    [T('🎒', 'Zabierz swoje plecaki i torby z parteru', 1, PN_PT),
     T('📚', '22 minuty przygotowanie lekcji i plecaka na jutro', 2, ND_CZW),
     T('🧹', 'Sprzątanie pokoju', 3, [CZW]),
     T('📖', 'Czytanie książki 30 minut', 2, [SO, ND]),
     T('📵', 'Bez ekranów do obiadu', 2, [ND])]
  );

  const cfg = wczytajKonfig_();
  cfg.tasks = {
    krystyna: wspolne([PN, WT, CZW, PT, SO], [PN, WT, CZW, PT], 1).concat([
      T('🍽️', 'Zmywarka', 1, CODZ),
      T('🧽', STOLY, 2, [SR, ND])
    ]),
    michal: wspolne([CZW, SO, ND], [CZW, SO], 1).concat([
      T('🔢', 'Matematyka/Kangur', 1, PN_PT),
      T('🗑️', 'Wynieś śmieci', 2, [PN, PT]),
      T('🧼', 'Umyj umywalkę w łazience', 2, [SR]),
      T('🧽', STOLY, 2, [WT, ND])
    ]),
    anna: wspolne([WT, SR, PT, SO], [WT, SR, PT], 2).concat([
      T('🧽', STOLY, 2, [PN, CZW, ND])
    ])
  };
  arkKonfig_().getRange('A1').setValue(JSON.stringify(cfg));
  Logger.log('Nowy plan zapisany.');
}

/**
 * Sklepik z nagrodami. Uruchom funkcję ustawSklep jeden raz z edytora.
 * 'osobista' – dziecko kupuje od razu ze swoich gwiazdek,
 * 'rodzinna' – dzieci wspólnie zbierają do skarbonki.
 */
function ustawSklep() {
  const cfg = wczytajKonfig_();
  cfg.rewards = [
    { id: 'r_kinder', t: 'Słodycz Kinder do szkoły', e: '🍫', c: 40, typ: 'osobista' },
    { id: 'r_apka', t: 'Moja appka w Claude', e: '💻', c: 60, typ: 'osobista' },
    { id: 'r_film', t: 'Wybór filmu na rodzinny wieczór', e: '🎬', c: 90, typ: 'osobista' },
    { id: 'r_gra', t: 'Dodatkowa godzina grania', e: '🎮', c: 100, typ: 'osobista' },
    { id: 'r_lody', t: 'Wyjście na lody', e: '🍦', c: 160, typ: 'osobista' },
    { id: 'r_ciuch', t: 'Wybrany ciuch', e: '👕', c: 260, typ: 'osobista' },
    { id: 'r_pizza', t: 'Rodzinna pizza', e: '🍕', c: 200, typ: 'rodzinna' },
    { id: 'r_termy', t: 'Rodzinne wyjście na termy', e: '♨️', c: 1555, typ: 'rodzinna' },
    { id: 'r_energy', t: 'Rodzinne wyjście do Energylandii', e: '🎢', c: 2222, typ: 'rodzinna' }
  ];
  arkKonfig_().getRange('A1').setValue(JSON.stringify(cfg));
  Logger.log('Sklepik zapisany.');
}
