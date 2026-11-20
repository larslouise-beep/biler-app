/**
 * Frimann Biler – Bil- og nøgleoversigt
 * Server/API til appen: Google Sheets er databasen, appen (index.html) ligger på GitHub Pages.
 *
 * Opsætning (kort – se den fulde guide):
 *  1. Skriv dit navn og en midlertidig PIN i FOERSTE_ADMIN herunder.
 *  2. Vælg funktionen "opsaetning" i værktøjslinjen og tryk Kør (giv adgang).
 *  3. Implementér > Ny implementering > Webapp:  Udfør som: Mig  ·  Hvem har adgang: Alle
 *  4. Kopiér webapp-URL'en (slutter på /exec) ind i index.html (API_URL).
 */

const FOERSTE_ADMIN = { navn: '', pin: '' }; // udfyldes kun lokalt første gang opsaetning() køres – læg aldrig PIN i git
const SESSION_TIMER = 6;          // så længe et login højst gælder (max 6 timer)
const MAKS_FORKERTE_PIN = 5;      // derefter låses medarbejderen i 15 minutter
const RADIUS_STD = 75;            // standard-radius (m) for en lokations GPS-område (cirkel)

const ARK = {
  medarbejdere: ['id', 'navn', 'rolle', 'aktiv', 'pinHash', 'salt', 'oprettet'],
  lokationer:   ['id', 'navn', 'type', 'adresse', 'oprettet', 'lat', 'lng', 'radius', 'omraade'],
  biler:        ['id', 'vin', 'regnr', 'maerke', 'model', 'farve', 'kategori', 'status',
                 'lokation', 'lokationNavn', 'plads', 'bilTs', 'bilAf', 'bilAfNavn',
                 'noegler', 'noeglerTekst', 'note', 'oprettet', 'opdateret', 'opdateretAf', 'opdateretAfNavn',
                 'gpsLat', 'gpsLng', 'gpsNoejagtighed', 'gpsTs', 'gpsAfNavn', 'kort'],
  historik:     ['tidspunkt', 'bilId', 'regnr', 'bil', 'medarbejderId', 'medarbejder', 'handling']
};
const ARKNAVN = { medarbejdere: 'Medarbejdere', lokationer: 'Lokationer', biler: 'Biler', historik: 'Historik' };
const TEKSTKOLONNER = ['id', 'vin', 'regnr', 'maerke', 'model', 'farve', 'plads', 'note', 'navn', 'adresse', 'bilId'];
const SPECIAL = { _ude: 'Ude af huset', _hos: 'Hos medarbejder' };
const STATUS = ['Ankommet', 'Klargøring', 'Klar til salg', 'Solgt', 'Udleveret'];

// ---------------------------------------------------------------- web-API
function doGet() {
  return json_({ ok: true, besked: 'Frimann bil- og nøgle-API kører.' });
}
function doPost(e) {
  let r;
  try {
    const b = JSON.parse(e.postData.contents);
    r = api(b.action, b.token, b.payload);
  } catch (x) {
    r = { error: 'FEJL', besked: String(x && x.message || x) };
  }
  return json_(r);
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function opsaetning() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(ARK).forEach(function (k) {
    let sh = ss.getSheetByName(ARKNAVN[k]);
    if (!sh) sh = ss.insertSheet(ARKNAVN[k]);
    if (sh.getLastRow() === 0) {
      sh.appendRow(ARK[k]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, ARK[k].length).setFontWeight('bold').setBackground('#282830').setFontColor('#ffffff');
    }
    const har = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    ARK[k].filter(function (c) { return har.indexOf(c) < 0; }).forEach(function (c) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(c).setFontWeight('bold').setBackground('#282830').setFontColor('#ffffff');
    });
    ARK[k].forEach(function (c, i) {
      if (TEKSTKOLONNER.indexOf(c) >= 0) sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    });
  });
  const ark1 = ss.getSheetByName('Ark1') || ss.getSheetByName('Sheet1');
  if (ark1 && ark1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(ark1);
  if (rows_('medarbejdere').length === 0) {
    if (!FOERSTE_ADMIN.navn || !FOERSTE_ADMIN.pin) throw new Error('Udfyld FOERSTE_ADMIN (navn og PIN) før opsaetning() køres første gang.');
    const salt = Utilities.getUuid();
    writeRow_('medarbejdere', { id: Utilities.getUuid(), navn: FOERSTE_ADMIN.navn, rolle: 'admin', aktiv: true,
      pinHash: hash_(String(FOERSTE_ADMIN.pin), salt), salt: salt, oprettet: new Date() });
  }
  ss.getSheetByName('Medarbejdere').hideColumns(5, 2); // skjul PIN-hash og salt
}

// ---------------------------------------------------------------- API (kaldes fra Index.html)
function api(action, token, p) {
  p = p || {};
  try {
    if (action === 'medarbejderListe') {
      return { ok: true, liste: rows_('medarbejdere').filter(aktiv_).map(function (e) { return { id: e.id, navn: e.navn }; })
        .sort(function (a, b) { return String(a.navn).localeCompare(String(b.navn), 'da'); }) };
    }
    if (action === 'login') return login_(p);
    const me = me_(token);
    if (action === 'logout') { CacheService.getScriptCache().remove('s_' + token); return { ok: true }; }
    if (action === 'data') return data_(me);
    if (action === 'historik') return { ok: true, liste: historik_(String(p.id)) };

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      switch (action) {
        case 'opretBil': opretBil_(me, p); break;
        case 'flyt': flyt_(me, p); break;
        case 'retBil': retBil_(me, p); break;
        case 'tilfoejNoegle': tilfoejNoegle_(me, p); break;
        case 'fjernNoegle': fjernNoegle_(me, p); break;
        case 'sletBil': admin_(me); sletRow_('biler', String(p.id)); break;
        case 'gemLokation': admin_(me); gemLokation_(p); break;
        case 'sletLokation': admin_(me); sletRow_('lokationer', String(p.id)); break;
        case 'gemMedarbejder': admin_(me); gemMedarbejder_(me, p); break;
        case 'skiftPin': skiftPin_(me, p); break;
        default: throw fejl_('UKENDT', 'Ukendt handling');
      }
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return data_(me_(token));
  } catch (e) {
    return { error: e.code || 'FEJL', besked: e.besked || e.message || String(e) };
  }
}

// ---------------------------------------------------------------- login
function hash_(pin, salt) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(b);
}
function aktiv_(e) { return e.aktiv === true || String(e.aktiv).toUpperCase() === 'TRUE'; }
function fejl_(code, besked) { return { code: code, besked: besked }; }
function admin_(me) { if (me.rolle !== 'admin') throw fejl_('ADGANG', 'Kun administratorer kan gøre dette.'); }

function login_(p) {
  const cache = CacheService.getScriptCache();
  const k = 'f_' + p.id;
  const n = Number(cache.get(k) || 0);
  if (n >= MAKS_FORKERTE_PIN) throw fejl_('LAAST', 'For mange forkerte forsøg. Prøv igen om 15 minutter.');
  const e = rows_('medarbejdere').filter(aktiv_).filter(function (x) { return x.id === p.id; })[0];
  if (!e || hash_(String(p.pin), e.salt) !== e.pinHash) {
    cache.put(k, String(n + 1), 900);
    throw fejl_('PIN', 'Forkert PIN-kode' + (n + 1 >= MAKS_FORKERTE_PIN ? ' – kontoen er låst i 15 minutter.' : '.'));
  }
  cache.remove(k);
  const token = Utilities.getUuid();
  cache.put('s_' + token, e.id, Math.min(21600, SESSION_TIMER * 3600));
  return { ok: true, token: token };
}
function me_(token) {
  const id = token ? CacheService.getScriptCache().get('s_' + token) : null;
  const e = id ? rows_('medarbejdere').filter(aktiv_).filter(function (x) { return x.id === id; })[0] : null;
  if (!e) throw fejl_('AUTH', 'Log ind igen.');
  return e;
}

// ---------------------------------------------------------------- data
function data_(me) {
  const lok = rows_('lokationer').map(function (l) { return { id: l.id, navn: String(l.navn), type: l.type || 'begge', adresse: String(l.adresse || ''),
    lat: num_(l.lat), lng: num_(l.lng), radius: num_(l.radius), omraade: polygon_(l.omraade) }; });
  const biler = rows_('biler').map(function (v) {
    const o = {};
    ARK.biler.forEach(function (c) { o[c] = v[c]; });
    ['vin', 'regnr', 'maerke', 'model', 'farve', 'plads', 'note', 'lokation', 'status', 'kategori'].forEach(function (c) { o[c] = String(o[c] == null ? '' : o[c]); });
    o.noegler = parseJson_(v.noegler);
    delete o.noeglerTekst; delete o.lokationNavn;
    return o;
  });
  const med = rows_('medarbejdere').map(function (e) { return { id: e.id, navn: String(e.navn), rolle: e.rolle, aktiv: aktiv_(e) }; });
  const res = { ok: true, me: { id: me.id, navn: String(me.navn), rolle: me.rolle }, lokationer: lok, biler: biler, medarbejdere: med };
  if (me.rolle === 'admin') res.arkUrl = SpreadsheetApp.getActive().getUrl();
  return res;
}
function historik_(id) {
  return rows_('historik').filter(function (h) { return String(h.bilId) === id; })
    .map(function (h) { return { ts: h.tidspunkt, af: h.medarbejder, tekst: h.handling }; })
    .sort(function (a, b) { return b.ts - a.ts; }).slice(0, 60);
}

// ---------------------------------------------------------------- biler
function locName_(id) {
  if (!id) return 'Ikke registreret';
  if (SPECIAL[id]) return SPECIAL[id];
  const l = rows_('lokationer').filter(function (x) { return x.id === id; })[0];
  return l ? String(l.navn) : 'Slettet lokation';
}
function bilLabel_(v) { return [v.maerke, v.model].filter(String).join(' ') || 'Ukendt bil'; }
function keyName_(keys, n) {
  const cnt = Math.max(1, Object.keys(keys).length);
  if (cnt === 1) return 'Nøgler';
  return 'Nøgle ' + n + (n === '1' ? ' (daglig)' : n === '2' ? ' (reserve)' : '');
}
function findBil_(id) {
  const v = rows_('biler').filter(function (x) { return x.id === id; })[0];
  if (!v) throw fejl_('FINDES_IKKE', 'Bilen findes ikke (måske slettet).');
  v.noegler = parseJson_(v.noegler);
  return v;
}
function gemBil_(v, me, tekst) {
  const now = Date.now();
  const keys = v.noegler || {};
  v.noeglerTekst = Object.keys(keys).sort().map(function (n) {
    const k = keys[n];
    const hvor = k.lokation === '_hos' ? 'Hos ' + (k.afNavn || '') : locName_(k.lokation) + (k.plads ? ' – ' + k.plads : '');
    return keyName_(keys, n) + ': ' + hvor;
  }).join(' | ');
  v.lokationNavn = locName_(v.lokation) + (v.plads ? ' – ' + v.plads : '');
  v.noegler = JSON.stringify(keys);
  v.opdateret = now; v.opdateretAf = me.id; v.opdateretAfNavn = me.navn;
  writeRow_('biler', v);
  if (tekst) writeRow_('historik', { tidspunkt: new Date(now), bilId: v.id, regnr: v.regnr, bil: bilLabel_(v), medarbejderId: me.id, medarbejder: me.navn, handling: tekst });
}
function opretBil_(me, p) {
  const vin = String(p.vin || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const reg = String(p.regnr || '').toUpperCase().trim();
  if (vin.length !== 17 && !reg) throw fejl_('DATA', 'Angiv stelnummer (17 tegn) eller reg.nr.');
  if (!p.lokation) throw fejl_('DATA', 'Vælg hvor bilen står.');
  const id = vin.length === 17 ? vin : 'REG-' + reg.replace(/[^A-Z0-9]/g, '');
  if (rows_('biler').some(function (x) { return x.id === id; })) throw fejl_('FINDES', 'Bilen findes allerede.');
  const now = Date.now();
  const noegler = {};
  const n = Math.max(1, Math.min(3, Number(p.antalNoegler) || 1));
  for (let i = 1; i <= n; i++) {
    noegler[String(i)] = p.noeglerSammeSted
      ? { lokation: p.lokation, plads: '', ts: now, af: me.id, afNavn: me.navn }
      : { lokation: '', plads: '', ts: 0, af: '', afNavn: '' };
  }
  const kat = ['ny', 'brugt', 'demo'].indexOf(p.kategori) >= 0 ? p.kategori : 'ny';
  const v = { _ny: true, id: id, vin: vin.length === 17 ? vin : '', regnr: reg, maerke: str_(p.maerke), model: str_(p.model), farve: str_(p.farve),
    kategori: kat, status: kat === 'brugt' ? 'Klargøring' : 'Ankommet', lokation: p.lokation, plads: str_(p.plads),
    bilTs: now, bilAf: me.id, bilAfNavn: me.navn, noegler: noegler, note: str_(p.note), oprettet: now };
  delete v._ny;
  const gps = saetGps_(v, me, p.gps);
  gemBil_(v, me, 'Oprettet – ' + locName_(v.lokation) + (v.plads ? ' – ' + v.plads : '') + (gps ? ' (GPS ±' + v.gpsNoejagtighed + ' m)' + gpsAfvigelse_(v.lokation, v) : ''));
}
function flyt_(me, p) {
  const v = findBil_(String(p.id));
  const lok = String(p.lokation || '');
  if (!lok) throw fejl_('DATA', 'Vælg en lokation.');
  const plads = str_(p.plads);
  let tekst;
  if (p.target === 'bil') {
    const flyttet = lok !== v.lokation || plads !== String(v.plads || '');
    v.lokation = lok; v.plads = plads; v.bilTs = Date.now(); v.bilAf = me.id; v.bilAfNavn = me.navn;
    tekst = (flyttet ? 'Bil flyttet til ' : 'Bil registreret på ') + locName_(lok) + (plads ? ' – ' + plads : '');
    if (saetGps_(v, me, p.gps)) tekst += ' (GPS ±' + v.gpsNoejagtighed + ' m)' + gpsAfvigelse_(lok, v);
  } else {
    const n = String(p.target).replace(/^n/, '');
    if (!v.noegler[n]) throw fejl_('DATA', 'Nøglen findes ikke.');
    v.noegler[n] = { lokation: lok, plads: plads, ts: Date.now(), af: me.id, afNavn: me.navn };
    tekst = keyName_(v.noegler, n) + (lok === '_hos' ? ' taget med' : ' lagt i ' + locName_(lok) + (plads ? ' – ' + plads : ''));
  }
  if (p.status && STATUS.indexOf(p.status) >= 0 && p.status !== v.status) { v.status = p.status; tekst += ' · status: ' + p.status; }
  gemBil_(v, me, tekst);
}
// Afstand i meter mellem to GPS-punkter (haversine)
function afstand_(lat1, lng1, lat2, lng2) {
  const r = Math.PI / 180, a = Math.sin((lat2 - lat1) * r / 2), b = Math.sin((lng2 - lng1) * r / 2);
  return 2 * 6371000 * Math.asin(Math.sqrt(a * a + Math.cos(lat1 * r) * Math.cos(lat2 * r) * b * b));
}
// Område-polygon gemt som JSON [[lat,lng],...] (3–30 hjørner) – ellers null
function polygon_(s) {
  try {
    const a = typeof s === 'string' ? JSON.parse(s || 'null') : s;
    if (!Array.isArray(a) || a.length < 3 || a.length > 30) return null;
    const ud = a.map(function (h) { return [Number(h && h[0]), Number(h && h[1])]; });
    return ud.every(function (h) { return isFinite(h[0]) && isFinite(h[1]) && Math.abs(h[0]) <= 90 && Math.abs(h[1]) <= 180 && !(h[0] === 0 && h[1] === 0); })
      ? ud.map(function (h) { return [Math.round(h[0] * 1e6) / 1e6, Math.round(h[1] * 1e6) / 1e6]; }) : null;
  } catch (e) { return null; }
}
// Hvor langt (m) et punkt ligger UDEN FOR polygonen (0 = indenfor). Lokal flad projektion – fint på pladsstørrelse.
function udenforPolygon_(poly, lat, lng) {
  const k = Math.PI / 180 * 6371000, c = Math.cos(lat * Math.PI / 180);
  const p = poly.map(function (h) { return [(h[1] - lng) * k * c, (h[0] - lat) * k]; });
  let inde = false, min = Infinity;
  for (let i = 0, j = p.length - 1; i < p.length; j = i++) {
    const a = p[i], b = p[j];
    if ((a[1] > 0) !== (b[1] > 0) && 0 < (b[0] - a[0]) * (0 - a[1]) / (b[1] - a[1]) + a[0]) inde = !inde;
    const dx = b[0] - a[0], dy = b[1] - a[1], L = dx * dx + dy * dy;
    const t = L ? Math.max(0, Math.min(1, -(a[0] * dx + a[1] * dy) / L)) : 0;
    min = Math.min(min, Math.hypot(a[0] + t * dx, a[1] + t * dy));
  }
  return inde ? 0 : min;
}
// Tekst til historikken, hvis bilens GPS ligger tydeligt uden for den valgte lokations område
function gpsAfvigelse_(lokId, v) {
  if (!lokId || SPECIAL[lokId]) return '';
  const l = rows_('lokationer').filter(function (x) { return x.id === lokId; })[0];
  if (!l) return '';
  const acc = Number(v.gpsNoejagtighed) || 0, poly = polygon_(l.omraade);
  let ude;
  if (poly) ude = udenforPolygon_(poly, v.gpsLat, v.gpsLng);
  else if (num_(l.lat) !== null && num_(l.lng) !== null) ude = afstand_(num_(l.lat), num_(l.lng), v.gpsLat, v.gpsLng) - (num_(l.radius) || RADIUS_STD);
  else return '';
  return ude - acc > 0 ? ' ⚠ GPS ' + Math.round(ude) + ' m uden for ' + String(l.navn) : '';
}
function saetGps_(v, me, g) {
  if (!g) return false;
  const lat = Number(g.lat), lng = Number(g.lng), acc = Math.round(Number(g.acc) || 0);
  if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180 || (lat === 0 && lng === 0)) return false;
  v.gpsLat = Math.round(lat * 1e6) / 1e6; v.gpsLng = Math.round(lng * 1e6) / 1e6; v.gpsNoejagtighed = acc;
  v.gpsTs = Date.now(); v.gpsAfNavn = me.navn;
  v.kort = 'https://www.google.com/maps/search/?api=1&query=' + v.gpsLat + ',' + v.gpsLng;
  return true;
}
function retBil_(me, p) {
  const v = findBil_(String(p.id));
  const f = p.felter || {};
  const aendret = [];
  ['regnr', 'status', 'maerke', 'model', 'farve', 'kategori', 'note'].forEach(function (k) {
    if (f[k] === undefined) return;
    let val = str_(f[k]); if (k === 'regnr') val = val.toUpperCase();
    if (k === 'status' && STATUS.indexOf(val) < 0) return;
    if (k === 'kategori' && ['ny', 'brugt', 'demo'].indexOf(val) < 0) return;
    if (String(v[k] == null ? '' : v[k]) !== val) { v[k] = val; aendret.push(k); }
  });
  if (!aendret.length) return;
  gemBil_(v, me, aendret.indexOf('status') >= 0 ? 'Status: ' + v.status : 'Oplysninger rettet');
}
function tilfoejNoegle_(me, p) {
  const v = findBil_(String(p.id));
  const ns = Object.keys(v.noegler).map(Number);
  const n = String((ns.length ? Math.max.apply(null, ns) : 0) + 1);
  v.noegler[n] = { lokation: '', plads: '', ts: 0, af: '', afNavn: '' };
  gemBil_(v, me, 'Reservenøgle tilføjet');
}
function fjernNoegle_(me, p) {
  const v = findBil_(String(p.id));
  const ns = Object.keys(v.noegler).sort();
  if (ns.length < 2) throw fejl_('DATA', 'Bilen skal have mindst ét nøgletag.');
  delete v.noegler[ns[ns.length - 1]];
  gemBil_(v, me, 'Reservenøgle fjernet');
}

// ---------------------------------------------------------------- lokationer / medarbejdere
function gemLokation_(p) {
  const navn = str_(p.navn);
  if (!navn) throw fejl_('DATA', 'Skriv et navn.');
  const type = ['begge', 'bil', 'noegle'].indexOf(p.type) >= 0 ? p.type : 'begge';
  let l = p.id ? rows_('lokationer').filter(function (x) { return x.id === p.id; })[0] : null;
  if (!l) l = { id: Utilities.getUuid().slice(0, 8), oprettet: new Date() };
  l.navn = navn; l.type = type; l.adresse = str_(p.adresse);
  if (p.rydOmraade) l.omraade = '';
  else if (p.omraade !== undefined && p.omraade !== null && p.omraade !== '') {
    const poly = polygon_(p.omraade);
    if (!poly) throw fejl_('DATA', 'Området skal have mindst 3 gyldige hjørner.');
    l.omraade = JSON.stringify(poly);
  }
  if (p.rydPos) { l.lat = ''; l.lng = ''; l.radius = ''; }
  else {
    if (p.lat !== undefined && p.lat !== null && p.lat !== '') {
      const lat = Number(p.lat), lng = Number(p.lng);
      if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180 || (lat === 0 && lng === 0)) throw fejl_('DATA', 'Ugyldig GPS-position.');
      l.lat = Math.round(lat * 1e6) / 1e6; l.lng = Math.round(lng * 1e6) / 1e6;
    }
    if (p.radius !== undefined && p.radius !== '') l.radius = Math.max(10, Math.min(2000, Math.round(Number(p.radius) || RADIUS_STD)));
  }
  sikrKolonner_('lokationer');
  writeRow_('lokationer', l);
}
function gyldigPin_(pin) {
  if (!/^\d{4,8}$/.test(String(pin))) throw fejl_('DATA', 'PIN-koden skal være 4–8 cifre.');
}
function gemMedarbejder_(me, p) {
  const navn = str_(p.navn);
  if (!navn) throw fejl_('DATA', 'Skriv et navn.');
  let e = p.id ? rows_('medarbejdere').filter(function (x) { return x.id === p.id; })[0] : null;
  if (!e) {
    if (!p.pin) throw fejl_('DATA', 'Giv den nye medarbejder en PIN-kode.');
    e = { id: Utilities.getUuid(), oprettet: new Date() };
  }
  if (e.id === me.id && (p.rolle !== 'admin' || p.aktiv === false)) throw fejl_('DATA', 'Du kan ikke fjerne din egen admin-adgang.');
  e.navn = navn;
  e.rolle = p.rolle === 'admin' ? 'admin' : 'bruger';
  e.aktiv = p.aktiv !== false;
  if (p.pin) { gyldigPin_(p.pin); e.salt = Utilities.getUuid(); e.pinHash = hash_(String(p.pin), e.salt); }
  writeRow_('medarbejdere', e);
}
function skiftPin_(me, p) {
  if (hash_(String(p.gammel), me.salt) !== me.pinHash) throw fejl_('PIN', 'Den nuværende PIN-kode er forkert.');
  gyldigPin_(p.ny);
  me.salt = Utilities.getUuid(); me.pinHash = hash_(String(p.ny), me.salt);
  writeRow_('medarbejdere', me);
}

// ---------------------------------------------------------------- ark-hjælpere
function sheet_(k) {
  const sh = SpreadsheetApp.getActive().getSheetByName(ARKNAVN[k]);
  if (!sh) throw fejl_('OPSAETNING', 'Arket "' + ARKNAVN[k] + '" mangler. Kør funktionen opsaetning().');
  return sh;
}
function rows_(k) {
  const v = sheet_(k).getDataRange().getValues();
  if (v.length < 2) return [];
  const h = v[0];
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const o = { _row: i + 1 };
    for (let j = 0; j < h.length; j++) { let x = v[i][j]; if (x instanceof Date) x = x.getTime(); o[h[j]] = x; }
    if (k === 'historik' ? o.bilId !== '' : o.id !== '') out.push(o);
  }
  return out;
}
function writeRow_(k, o) {
  const sh = sheet_(k);
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const vals = h.map(function (c) { return o[c] === undefined || o[c] === null ? '' : o[c]; });
  if (o._row) sh.getRange(o._row, 1, 1, h.length).setValues([vals]);
  else sh.appendRow(vals);
}
function sletRow_(k, id) {
  const r = rows_(k).filter(function (x) { return x.id === id; })[0];
  if (r) sheet_(k).deleteRow(r._row);
}
function parseJson_(s) { try { const o = JSON.parse(s || '{}'); return o && typeof o === 'object' ? o : {}; } catch (e) { return {}; } }
function str_(s) { return String(s == null ? '' : s).trim(); }
function num_(x) { if (x === '' || x === null || x === undefined) return null; const n = Number(x); return isFinite(n) ? n : null; }
// Tilføjer manglende kolonner bagerst i arket (så opsaetning() ikke skal køres igen efter en opdatering)
function sikrKolonner_(k) {
  const sh = sheet_(k);
  const har = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  ARK[k].filter(function (c) { return har.indexOf(c) < 0; }).forEach(function (c) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue(c).setFontWeight('bold').setBackground('#282830').setFontColor('#ffffff');
  });
}
