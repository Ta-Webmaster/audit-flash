/**
 * AUDIT FLASH
 * Conçu par Alexia Gombert (Ta Webmaster) avec Claude Code. Licence MIT.
 * Colle des URL dans l'onglet "Audits" (colonne A), puis menu "Audit flash" > "Analyser les nouvelles URL".
 * Le script interroge PageSpeed Insights (mobile), lit le code de la page pour reconnaître
 * WordPress, le constructeur, le thème et les extensions visibles, puis rédige les symptômes
 * et un brouillon de mail. Il ne modifie jamais le site analysé.
 */

var CONFIG = {
  ONGLET: 'Audits',
  MAX_PAR_LANCEMENT: 6,          // 6 sites par lancement pour rester sous la limite de 6 min de Google
  WP_VERSION_RECENTE: '6.6',     // à relever de temps en temps : en dessous, la version est jugée ancienne
  LIEN_RDV: 'https://exemple.fr/prendre-rendez-vous',   // à remplacer par ton lien de prise de rendez-vous
  NOM: 'Prénom Nom',                                      // à remplacer
  PRESENTATION: 'architecte digitale WordPress près de Grasse',   // à adapter
  SIGNATURE: 'Prénom Nom\nTon activité | ton-site.fr\n00 00 00 00 00'   // à remplacer
};

var ENTETES = ['URL à analyser', 'Statut', 'Performance mobile /100', 'SEO /100', 'LCP', 'TBT', 'CLS',
  'Poids de la page', 'WordPress', 'Version WP', 'Thème', 'Constructeur', 'Extensions visibles',
  'HTTPS', 'Viewport mobile', 'Meta description', 'Nb de H1', 'Symptômes', 'Brouillon de mail', 'Analysé le'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Audit flash')
    .addItem('1. Enregistrer la clé PageSpeed', 'enregistrerCle')
    .addItem('2. Préparer l\'onglet Audits', 'preparerOnglet')
    .addItem('3. Analyser les nouvelles URL', 'analyserNouvelles')
    .addItem('Refaire les lignes en erreur', 'refaireErreurs')
    .addToUi();
  menuSourcing_();   // défini dans Sourcing.gs
}

function enregistrerCle() {
  var ui = SpreadsheetApp.getUi();
  var rep = ui.prompt('Clé API PageSpeed Insights',
    'Colle ta clé (gratuite, créée dans Google Cloud, API "PageSpeed Insights"). Sans clé, Google limite très vite les requêtes.',
    ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() !== ui.Button.OK) return;
  PropertiesService.getScriptProperties().setProperty('PSI_KEY', rep.getResponseText().trim());
  ui.alert('Clé enregistrée.');
}

function preparerOnglet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var feuille = ss.getSheetByName(CONFIG.ONGLET) || ss.insertSheet(CONFIG.ONGLET);
  feuille.getRange(1, 1, 1, ENTETES.length).setValues([ENTETES])
    .setFontWeight('bold').setBackground('#333333').setFontColor('#ffffff').setWrap(true);
  feuille.setFrozenRows(1);
  feuille.getRange(2, 18, 500, 2).setWrap(true).setVerticalAlignment('top');
  feuille.setColumnWidth(1, 240);
  feuille.setColumnWidth(18, 360);
  feuille.setColumnWidth(19, 460);
  var plage = feuille.getRange('C2:C500');
  feuille.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(50).setBackground('#f4c7c3').setRanges([plage]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(50, 89).setBackground('#fce8b2').setRanges([plage]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(90).setBackground('#b7e1cd').setRanges([plage]).build()
  ]);
}

function analyserNouvelles() { lancer(false); }
function refaireErreurs() { lancer(true); }

function lancer(refaire) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var feuille = ss.getSheetByName(CONFIG.ONGLET);
  if (!feuille) { preparerOnglet(); feuille = ss.getSheetByName(CONFIG.ONGLET); }
  var derniere = feuille.getLastRow();
  var traitees = 0;
  for (var l = 2; l <= derniere && traitees < CONFIG.MAX_PAR_LANCEMENT; l++) {
    var url = String(feuille.getRange(l, 1).getValue()).trim();
    var statut = String(feuille.getRange(l, 2).getValue());
    if (!url) continue;
    var aFaire = refaire ? statut.indexOf('Erreur') === 0 : statut === '';
    if (!aFaire) continue;
    feuille.getRange(l, 2).setValue('En cours...');
    SpreadsheetApp.flush();
    try {
      var r = auditerUrl(url);
      feuille.getRange(l, 2, 1, ENTETES.length - 1).setValues([[
        r.statut, r.perf, r.seo, r.lcp, r.tbt, r.cls, r.poids, r.wp, r.versionWp, r.theme,
        r.constructeur, r.extensions, r.https, r.viewport, r.meta, r.h1, r.symptomes, r.mail, r.date
      ]]);
    } catch (e) {
      feuille.getRange(l, 2).setValue('Erreur : ' + String(e.message || e).substring(0, 150));
    }
    traitees++;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(traitees + ' site(s) analysé(s).', 'Audit flash', 5);
}

// ---------- Analyse d'un site ----------

function auditerUrl(brut) {
  var url = normaliserUrl(brut);
  var page = recupererPage(url);
  var tech = detecter(page.html);
  var psi = pagespeed(page.urlUtilisee);
  var r = {
    statut: 'OK',
    perf: psi.perf, seo: psi.seo, lcp: psi.lcp, tbt: psi.tbt, cls: psi.cls, poids: psi.poids,
    wp: tech.wordpress ? 'Oui' : 'Non',
    versionWp: tech.versionWp || '',
    theme: tech.theme || '',
    constructeur: tech.constructeur || '',
    extensions: tech.extensions.join(', '),
    https: page.httpsOk ? 'Oui' : 'Non',
    viewport: tech.viewport ? 'Oui' : 'Non',
    meta: tech.meta ? 'Oui' : 'Non',
    h1: tech.h1,
    date: Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm')
  };
  var liste = construireSymptomes(r, tech, page);
  r.symptomes = liste.length ? liste.map(function (s) { return '- ' + s; }).join('\n') : 'Aucun problème majeur repéré.';
  r.mail = construireMail(url, liste);
  return r;
}

function normaliserUrl(brut) {
  var u = String(brut).trim();
  if (!/^https?:\/\//i.test(u)) u = 'https://' + u;
  return u;
}

function recupererPage(url) {
  var essais = [url];
  if (/^https:/i.test(url)) essais.push(url.replace(/^https:/i, 'http:'));
  var derniereErreur = null;
  for (var i = 0; i < essais.length; i++) {
    try {
      var rep = UrlFetchApp.fetch(essais[i], {
        muteHttpExceptions: true, followRedirects: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (compatible; AuditFlash/1.0)', 'Accept-Language': 'fr-FR,fr;q=0.9' }
      });
      return { html: rep.getContentText(), code: rep.getResponseCode(), httpsOk: /^https:/i.test(essais[i]), urlUtilisee: essais[i] };
    } catch (e) { derniereErreur = e; }
  }
  throw new Error('Site injoignable (' + (derniereErreur && derniereErreur.message) + ')');
}

function pagespeed(url) {
  var cle = PropertiesService.getScriptProperties().getProperty('PSI_KEY');
  var api = 'https://www.googleapis.com/pagespeedonline/v5/runPagespeed?strategy=mobile&category=PERFORMANCE&category=SEO&locale=fr&url=' +
    encodeURIComponent(url) + (cle ? '&key=' + encodeURIComponent(cle) : '');
  var rep = null;
  for (var t = 0; t < 2; t++) {
    rep = UrlFetchApp.fetch(api, { muteHttpExceptions: true });
    if (rep.getResponseCode() === 200) break;
    if (rep.getResponseCode() === 429) Utilities.sleep(6000); else break;
  }
  if (rep.getResponseCode() !== 200) throw new Error('PageSpeed a répondu ' + rep.getResponseCode() + (rep.getResponseCode() === 429 ? ' (limite atteinte, ajoute une clé)' : ''));
  var lr = JSON.parse(rep.getContentText()).lighthouseResult;
  var a = lr.audits;
  var octets = a['total-byte-weight'] && a['total-byte-weight'].numericValue;
  return {
    perf: Math.round(lr.categories.performance.score * 100),
    seo: Math.round(lr.categories.seo.score * 100),
    lcp: a['largest-contentful-paint'] ? a['largest-contentful-paint'].displayValue : '',
    tbt: a['total-blocking-time'] ? a['total-blocking-time'].displayValue : '',
    cls: a['cumulative-layout-shift'] ? a['cumulative-layout-shift'].displayValue : '',
    poids: octets ? (octets / 1048576).toFixed(1).replace('.', ',') + ' Mo' : ''
  };
}

// ---------- Détection (équivalent simplifié de Wappalyzer, sans abonnement) ----------

function detecter(html) {
  var h = html || '';
  var t = { extensions: [] };
  t.wordpress = /\/wp-content\/|\/wp-includes\/|\/wp-json\//i.test(h);

  var gen = h.match(/<meta[^>]+name=["']generator["'][^>]+content=["']WordPress ([0-9.]+)/i) ||
            h.match(/<meta[^>]+content=["']WordPress ([0-9.]+)["'][^>]+name=["']generator["']/i);
  var verCoeur = h.match(/\/wp-includes\/[^"']+\?ver=([0-9]+\.[0-9]+(?:\.[0-9]+)?)/i);
  if (gen) t.versionWp = gen[1];
  else if (verCoeur) t.versionWp = verCoeur[1] + ' (probable)';

  var th = h.match(/\/wp-content\/themes\/([a-z0-9_-]+)\//i);
  if (th) t.theme = th[1];

  var vus = {}, re = /\/wp-content\/plugins\/([a-z0-9_-]+)\/[^"' >]*/gi, m;
  while ((m = re.exec(h)) !== null) {
    var v = m[0].match(/\?ver=([0-9][0-9.]*)/);
    if (!vus[m[1]] || (v && !vus[m[1]].v)) vus[m[1]] = { v: v ? v[1] : '' };
  }
  Object.keys(vus).forEach(function (nom) { t.extensions.push(vus[nom].v ? nom + ' ' + vus[nom].v : nom); });
  t.nbExtensions = t.extensions.length;

  if (/\/plugins\/elementor\/|elementor-element|name=["']generator["'][^>]+Elementor/i.test(h)) t.constructeur = 'Elementor';
  else if (/\/themes\/divi\/|et_pb_section/i.test(h)) t.constructeur = 'Divi';
  else if (/js_composer|vc_row/i.test(h)) t.constructeur = 'WPBakery';
  else if (/bb-plugin|fl-builder/i.test(h)) t.constructeur = 'Beaver Builder';

  if (!t.wordpress) {
    if (/wixstatic\.com|static\.parastorage\.com/i.test(h)) t.autre = 'Wix';
    else if (/squarespace/i.test(h)) t.autre = 'Squarespace';
    else if (/cdn\.shopify\.com/i.test(h)) t.autre = 'Shopify';
    else if (/webflow/i.test(h)) t.autre = 'Webflow';
    else if (/jimdo/i.test(h)) t.autre = 'Jimdo';
    if (t.autre) t.constructeur = t.autre;
  }

  t.viewport = /<meta[^>]+name=["']viewport["']/i.test(h);
  t.meta = /<meta[^>]+name=["']description["'][^>]+content=["'][^"']{20,}/i.test(h) ||
           /<meta[^>]+content=["'][^"']{20,}["'][^>]+name=["']description["']/i.test(h);
  t.h1 = (h.match(/<h1[\s>]/gi) || []).length;
  return t;
}

function comparerVersions(a, b) {
  var pa = String(a).split('.').map(Number), pb = String(b).split('.').map(Number);
  for (var i = 0; i < Math.max(pa.length, pb.length); i++) {
    var x = pa[i] || 0, y = pb[i] || 0;
    if (x !== y) return x < y ? -1 : 1;
  }
  return 0;
}

function construireSymptomes(r, tech, page) {
  var s = [];
  if (!r.https || r.https === 'Non') s.push('Le site ne s\'ouvre pas correctement en HTTPS : les navigateurs affichent un avertissement de sécurité aux visiteurs.');
  if (typeof r.perf === 'number') {
    if (r.perf < 50) s.push('Site très lent sur mobile (score ' + r.perf + '/100, affichage principal en ' + r.lcp + '). Une partie des visiteurs part avant d\'avoir vu la page.');
    else if (r.perf < 90) s.push('Vitesse mobile moyenne (score ' + r.perf + '/100). Des gains importants sont possibles.');
  }
  if (r.poids) {
    var mo = parseFloat(String(r.poids).replace(',', '.'));
    if (mo >= 3) s.push('Page lourde (' + r.poids + ') : elle ralentit l\'affichage et consomme inutilement de l\'énergie.');
  }
  if (tech.autre) s.push('Site construit avec ' + tech.autre + ' : moins de maîtrise du référencement et de la performance qu\'avec un WordPress bien structuré.');
  if (tech.wordpress && tech.versionWp) {
    var vn = tech.versionWp.replace(/[^0-9.]/g, '');
    if (comparerVersions(vn, CONFIG.WP_VERSION_RECENTE) < 0) s.push('Version de WordPress ancienne (' + tech.versionWp + ') : risque de sécurité et de dysfonctionnement.');
  }
  if (tech.nbExtensions >= 15) s.push(tech.nbExtensions + ' extensions chargent du code chez le visiteur : le site est alourdi et plus exposé aux failles.');
  if (!tech.viewport) s.push('Pas d\'adaptation mobile déclarée : le site peut s\'afficher mal sur téléphone.');
  if (!tech.meta) s.push('Pas de description pour Google : l\'extrait affiché dans les résultats de recherche est subi, pas choisi.');
  if (tech.h1 === 0) s.push('Aucun titre principal (H1) : Google comprend mal le sujet de la page.');
  else if (tech.h1 > 1) s.push(tech.h1 + ' titres principaux (H1) sur la page : la structure SEO est brouillée.');
  if (typeof r.seo === 'number' && r.seo < 80) s.push('Score SEO technique de ' + r.seo + '/100 : des bases de référencement manquent.');
  return s;
}

function construireMail(url, symptomes) {
  var domaine = url.replace(/^https?:\/\//i, '').replace(/\/.*$/, '');
  var points = symptomes.slice(0, 3).map(function (x) { return '- ' + x; }).join('\n');
  if (!points) points = '- Votre site est globalement sain, mais il reste des marges de progression sur la visibilité.';
  return 'Objet : Analyse rapide de ' + domaine + '\n\n' +
    'Bonjour,\n\n' +
    'Je suis ' + CONFIG.PRESENTATION + '. J\'ai regardé ' + domaine + ' (analyse de l\'extérieur, sans rien modifier) et voici ce que j\'ai relevé :\n\n' +
    points + '\n\n' +
    'Ces points se corrigent. Je vous propose un échange de 30 minutes, gratuit, pour vous expliquer ce qui pèse le plus sur votre visibilité et dans quel ordre agir : ' + CONFIG.LIEN_RDV + '\n\n' +
    'Bien cordialement,\n' + CONFIG.SIGNATURE;
}