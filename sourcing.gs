/**
 * SOURCING
 * Conçu par Alexia Gombert (Ta Webmaster) avec Claude Code. Licence MIT.
 * A coller dans un NOUVEAU fichier de script (+ > Script) du même projet que l'Audit flash.
 * Menu "Sourcing" :
 *   1. enregistrer la clé Google Places
 *   2. préparer l'onglet "Sourcing"
 *   3. chercher de nouveaux prospects (Google Maps), 3 bis. chercher des agences sur toute une zone,
 *      3 ter. chercher des artisans (métier x commune) dans le 06 et le 83,
 *      3 quater. chercher des thérapeutes et praticiens bien-être (pratique x commune)
 *   4. enrichir les lignes (dirigeant, email, constat sur le site)
 *   5. transférer les lignes "Validé" vers l'onglet Prospects
 *   6. préparer la journée d'envois (onglet Aujourd'hui : relances dues, premiers mails, mails prêts)
 *   7. enregistrer les envois cochés "Fait" dans Prospects
 * Il ne modifie jamais les sites analysés. Il lit seulement des pages publiques.
 * Ce fichier réutilise des fonctions de l'Audit flash (recupererPage, detecter, comparerVersions, CONFIG).
 */

var SRC = {
  ONGLET: 'Sourcing',
  PROSPECTS: 'Prospects',
  AUDITS: 'Audits',
  MAX_RESULTATS: 20,
  DUREE_MAX_MS: 270000   // le script s'arrête de lui-même avant la limite de 6 minutes de Google
};

var SRC_ENTETES = ['Nom', 'Type', 'Statut', 'Adresse', 'Téléphone', 'Site web', 'Email', 'Page contact',
  'Décideur (dirigeant légal)', 'SIREN', 'Constat site', 'Pistes repérées', 'Source', 'Ajouté le', 'Note', 'ID Maps'];

var COL = { NOM: 1, TYPE: 2, STATUT: 3, ADR: 4, TEL: 5, SITE: 6, MAIL: 7, CONTACT: 8, DECIDEUR: 9,
  SIREN: 10, CONSTAT: 11, PISTES: 12, SOURCE: 13, DATE: 14, NOTE: 15, ID: 16 };

var SRC_STATUTS = ['À enrichir', 'À qualifier', 'Validé', 'Écarté', 'Transféré'];

// Liste déroulante des types de prospect (colonne D de Prospects, colonne B de Sourcing).
var TYPES_PROSPECT = ['Agence', 'Artisan', 'Commerçant', 'Praticien bien-être', 'Restaurant', 'Autre'];

// Ramène ce qui a été tapé à l'une des valeurs de la liste (sans tenir compte des accents ni des majuscules).
function typeProspect_(t) {
  var n = nomNorm_(t);
  if (!n) return 'Autre';
  for (var i = 0; i < TYPES_PROSPECT.length; i++) { if (nomNorm_(TYPES_PROSPECT[i]) === n) return TYPES_PROSPECT[i]; }
  if (/agence/.test(n)) return 'Agence';
  if (/artisan|plomb|electric|macon|menuis|peintre|carrel|couvr|paysag|chauff|serrur|renov/.test(n)) return 'Artisan';
  if (/commerc|boutique|magasin|tpe/.test(n)) return 'Commerçant';
  if (/praticien|therap|bien etre|naturo|sophro|massage|spa|yoga|coach/.test(n)) return 'Praticien bien-être';
  if (/restau|brasserie|traiteur|cafe|bar/.test(n)) return 'Restaurant';
  return 'Autre';
}

function reglesType_() {
  return SpreadsheetApp.newDataValidation().requireValueInList(TYPES_PROSPECT, true).setAllowInvalid(false).build();
}

function menuSourcing_() {
  SpreadsheetApp.getUi().createMenu('Sourcing')
    .addItem('1. Enregistrer la clé Google Places', 'src_enregistrerCle')
    .addItem('2. Préparer l\'onglet Sourcing', 'src_preparerOnglet')
    .addItem('3. Chercher de nouveaux prospects (une recherche)', 'src_chercher')
    .addItem('3 bis. Chercher des agences sur toute une zone', 'src_chercherLot')
    .addItem('3 ter. Chercher des artisans (métier x communes)', 'src_chercherArtisans')
    .addItem('3 quater. Chercher des thérapeutes (pratique x communes)', 'src_chercherTherapeutes')
    .addItem('4. Enrichir les lignes (dirigeant, email, site)', 'src_enrichir')
    .addItem('5. Transférer les lignes validées vers Prospects', 'src_transferer')
    .addSeparator()
    .addItem('6. Préparer ma journée d\'envois (onglet Aujourd\'hui)', 'env_aujourdhui')
    .addItem('7. Enregistrer les envois cochés "Fait"', 'env_marquerFaits')
    .addToUi();
}


// ---------- Préparation ----------

function src_enregistrerCle() {
  var ui = SpreadsheetApp.getUi();
  var rep = ui.prompt('Clé API Google Places',
    'Colle ta clé (créée dans Google Cloud, API "Places API (New)" activée).',
    ui.ButtonSet.OK_CANCEL);
  if (rep.getSelectedButton() !== ui.Button.OK) return;
  PropertiesService.getScriptProperties().setProperty('PLACES_KEY', rep.getResponseText().trim());
  ui.alert('Clé enregistrée.');
}

function src_preparerOnglet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = ss.getSheetByName(SRC.ONGLET) || ss.insertSheet(SRC.ONGLET);
  f.getRange(1, 1, 1, SRC_ENTETES.length).setValues([SRC_ENTETES])
    .setFontWeight('bold').setBackground('#333333').setFontColor('#ffffff').setWrap(true);
  f.setFrozenRows(1);
  var regle = SpreadsheetApp.newDataValidation().requireValueInList(SRC_STATUTS, true).setAllowInvalid(false).build();
  f.getRange(2, COL.STATUT, 1000, 1).setDataValidation(regle);
  f.getRange(2, COL.TYPE, 1000, 1).setDataValidation(reglesType_());
  f.getRange(2, COL.TEL, 1000, 1).setNumberFormat('@');
  f.getRange(2, 1, 1000, SRC_ENTETES.length).setVerticalAlignment('top');
  f.getRange(2, COL.PISTES, 1000, 1).setWrap(true);
  f.getRange(2, COL.DECIDEUR, 1000, 1).setWrap(true);
  f.setColumnWidth(COL.NOM, 220);
  f.setColumnWidth(COL.ADR, 240);
  f.setColumnWidth(COL.SITE, 200);
  f.setColumnWidth(COL.MAIL, 220);
  f.setColumnWidth(COL.DECIDEUR, 220);
  f.setColumnWidth(COL.PISTES, 300);
  f.setColumnWidth(COL.NOTE, 200);
  // Lien avec Prospects : état de la relation et date du 1er contact, mis à jour automatiquement.
  f.getRange(1, 17, 1, 2).setValues([['État de la relation (Prospects)', 'Date du 1er contact']])
    .setFontWeight('bold').setBackground('#333333').setFontColor('#ffffff').setWrap(true);
  f.getRange(2, 18, 1000, 1).setNumberFormat('dd/MM/yyyy');
  f.setColumnWidth(17, 170);
  f.setColumnWidth(18, 170);
  SpreadsheetApp.getActiveSpreadsheet().toast('Onglet Sourcing prêt.', 'Sourcing', 5);
}


// ---------- 3. Chercher (une recherche) ----------

function src_chercher() {
  var ui = SpreadsheetApp.getUi();
  var cle = cleLieux_();
  if (!cle) return;

  var r1 = ui.prompt('Type de prospect', 'Choisis dans la liste : ' + TYPES_PROSPECT.join(', ') + '.', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var type = typeProspect_(r1.getResponseText().trim());

  var r2 = ui.prompt('Recherche Google Maps', 'Exemple : agence web Cannes, plombier Grasse, spa Antibes.', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var requete = r2.getResponseText().trim();
  if (!requete) { ui.alert('La recherche est obligatoire.'); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = ss.getSheetByName(SRC.ONGLET);
  if (!f) { src_preparerOnglet(); f = ss.getSheetByName(SRC.ONGLET); }

  var res;
  try {
    res = ajouterRecherche_(requete, type, cle, f, connus_(), dateJour_());
  } catch (e) {
    ui.alert(e.message);
    return;
  }
  ss.toast(res.ajoutes + ' nouveau(x) prospect(s), ' + res.ignores + ' ignoré(s) (déjà connus ou fermés). Lance l\'étape 4.', 'Sourcing', 8);
}

// ---------- 3 bis. Chercher des agences sur toute une zone ----------

var LOT_ZONES = {
  '1': ['Nice', 'Cannes', 'Antibes', 'Sophia Antipolis', 'Grasse', 'Mougins', 'Cagnes-sur-Mer', 'Menton', 'Vence',
    'Saint-Laurent-du-Var', 'Fréjus', 'Draguignan', 'Toulon', 'Hyères'],
  '2': ['Marseille', 'Aix-en-Provence', 'Avignon', 'Montpellier', 'Nîmes', 'Toulouse', 'Perpignan', 'Béziers'],
  '3': ['Lyon', 'Paris', 'Bordeaux', 'Nantes', 'Lille', 'Strasbourg', 'Rennes', 'Grenoble', 'Annecy', 'Dijon', 'Tours', 'Rouen']
};
var LOT_VARIANTES = ['agence web', 'agence WordPress'];

function src_chercherLot() {
  var ui = SpreadsheetApp.getUi();
  var cle = cleLieux_();
  if (!cle) return;

  var r = ui.prompt('Lot d\'agences',
    'Quelle zone ?\n1 = Alpes-Maritimes et Var\n2 = Sud de la France (Marseille, Montpellier, Toulouse...)\n' +
    '3 = Grandes villes de France (Lyon, Paris, Bordeaux...)\nTape 1, 2 ou 3. Le travail se fait à distance : la marque blanche n\'a pas besoin de proximité.',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var choix = r.getResponseText().trim();
  var villes = LOT_ZONES[choix];
  if (!villes) { ui.alert('Tape 1, 2 ou 3.'); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = ss.getSheetByName(SRC.ONGLET);
  if (!f) { src_preparerOnglet(); f = ss.getSheetByName(SRC.ONGLET); }

  var props = PropertiesService.getScriptProperties();
  var cleIdx = 'LOT_' + choix;
  var total = villes.length * LOT_VARIANTES.length;
  var idx = Number(props.getProperty(cleIdx) || 0);
  if (idx >= total) idx = 0;

  var deja = connus_(), date = dateJour_();
  var debut = Date.now(), ajoutes = 0, ignores = 0, faites = 0, erreur = '';

  for (; idx < total; idx++) {
    if (Date.now() - debut > SRC.DUREE_MAX_MS - 30000) break;
    var ville = villes[Math.floor(idx / LOT_VARIANTES.length)];
    var variante = LOT_VARIANTES[idx % LOT_VARIANTES.length];
    try {
      var res = ajouterRecherche_(variante + ' ' + ville, 'Agence', cle, f, deja, date);
      ajoutes += res.ajoutes; ignores += res.ignores; faites++;
    } catch (e) { erreur = e.message; break; }
    Utilities.sleep(200);
  }
  props.setProperty(cleIdx, String(idx));

  if (erreur) { ui.alert('Arrêt après ' + faites + ' recherche(s) :\n' + erreur); }
  var reste = total - idx;
  ss.toast(faites + ' recherche(s), ' + ajoutes + ' nouvelle(s) agence(s), ' + ignores + ' ignorée(s). ' +
    (reste > 0 ? 'Relance cette étape pour continuer (' + reste + ' recherche(s) restantes), ou passe à l\'étape 4.' : 'Zone terminée : passe à l\'étape 4.'),
    'Sourcing', 10);
}

// ---------- 3 ter. Chercher des artisans (métier x communes) ----------

var ART_ZONES = {
  '1': ['Grasse', 'Mouans-Sartoux', 'Saint-Vallier-de-Thiey', 'Cabris', 'Peymeinade', 'Le Cannet', 'Mougins', 'Mandelieu-la-Napoule', 'Cannes', 'Vence'],
  '2': ['Antibes', 'Valbonne', 'Biot', 'Vallauris', 'Cagnes-sur-Mer', 'Saint-Laurent-du-Var', 'Nice', 'Menton'],
  '3': ['Fréjus', 'Saint-Raphaël', 'Draguignan', 'Les Arcs', 'Lorgues', 'Brignoles', 'Toulon', 'Hyères']
};
var ART_METIERS = ['plombier', 'électricien', 'maçon', 'menuisier', 'peintre en bâtiment', 'carreleur', 'couvreur',
  'paysagiste', 'chauffagiste', 'serrurier', 'artisan rénovation'];

var THERA_PRATIQUES = ['naturopathe', 'sophrologue', 'hypnothérapeute', 'massage bien-être', 'réflexologue',
  'soins énergétiques', 'psychopraticien', 'praticien Reiki', 'coach de vie', 'professeur de yoga'];

function src_chercherArtisans() {
  lotMetiers_('Artisan', 'artisans', ART_METIERS, 'Métier', 'ART_');
}

function src_chercherTherapeutes() {
  lotMetiers_('Praticien bien-être', 'thérapeutes', THERA_PRATIQUES, 'Pratique', 'THE_');
}

// Recherche par lot : une pratique ou un métier x les communes d'une zone. Reprend où elle s'est arrêtée.
function lotMetiers_(type, mot, liste, libelle, prefixe) {
  var ui = SpreadsheetApp.getUi();
  var cle = cleLieux_();
  if (!cle) return;

  var r = ui.prompt('Lot de ' + mot,
    'Quelle zone ?\n1 = Pays de Grasse, Cannes, Mougins, Vence\n2 = Antibes, Sophia, Cagnes, Nice, Menton\n3 = Var (Fréjus, Draguignan, Toulon...)\nTape 1, 2 ou 3.',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var zone = r.getResponseText().trim();
  var villes = ART_ZONES[zone];
  if (!villes) { ui.alert('Tape 1, 2 ou 3.'); return; }

  var r2 = ui.prompt(libelle,
    'Un seul (ex. ' + liste[0] + ') ou laisse vide pour tous : ' + liste.join(', ') + '.',
    ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var choisi = r2.getResponseText().trim();
  var items = choisi ? [choisi] : liste;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = ss.getSheetByName(SRC.ONGLET);
  if (!f) { src_preparerOnglet(); f = ss.getSheetByName(SRC.ONGLET); }

  var props = PropertiesService.getScriptProperties();
  var cleIdx = prefixe + zone + '_' + nomNorm_(choisi || 'tous').replace(/ /g, '_');
  var total = villes.length * items.length;
  var idx = Number(props.getProperty(cleIdx) || 0);
  if (idx >= total) idx = 0;

  var deja = connus_(), date = dateJour_();
  var debut = Date.now(), ajoutes = 0, ignores = 0, faites = 0, erreur = '';

  for (; idx < total; idx++) {
    if (Date.now() - debut > SRC.DUREE_MAX_MS - 30000) break;
    var ville = villes[Math.floor(idx / items.length)];
    var m = items[idx % items.length];
    try {
      var res = ajouterRecherche_(m + ' ' + ville, type, cle, f, deja, date);
      ajoutes += res.ajoutes; ignores += res.ignores; faites++;
    } catch (e) { erreur = e.message; break; }
    Utilities.sleep(200);
  }
  props.setProperty(cleIdx, String(idx));

  if (erreur) { ui.alert('Arrêt après ' + faites + ' recherche(s) :\n' + erreur); }
  var reste = total - idx;
  ss.toast(faites + ' recherche(s), ' + ajoutes + ' nouvelle(s) ligne(s), ' + ignores + ' ignorée(s). ' +
    (reste > 0 ? 'Relance cette étape pour continuer (' + reste + ' recherche(s) restantes), ou passe à l\'étape 4.' : 'Zone terminée : passe à l\'étape 4.'),
    'Sourcing', 10);
}

// Interroge Google Maps et ajoute au Sourcing les entreprises pas encore connues.
function ajouterRecherche_(requete, type, cle, f, deja, date) {
  var rep = UrlFetchApp.fetch('https://places.googleapis.com/v1/places:searchText', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      'X-Goog-Api-Key': cle,
      'X-Goog-FieldMask': 'places.id,places.displayName,places.formattedAddress,places.nationalPhoneNumber,places.websiteUri,places.businessStatus'
    },
    payload: JSON.stringify({ textQuery: requete, languageCode: 'fr', regionCode: 'FR', pageSize: SRC.MAX_RESULTATS })
  });
  if (rep.getResponseCode() !== 200) {
    throw new Error('Google Places a répondu ' + rep.getResponseCode() + ' :\n' + rep.getContentText().substring(0, 400));
  }
  var lieux = JSON.parse(rep.getContentText()).places || [];

  var lignes = [], ignores = 0;
  lieux.forEach(function (p) {
    var nom = p.displayName && p.displayName.text ? p.displayName.text : '';
    if (!nom || p.businessStatus === 'CLOSED_PERMANENTLY') { ignores++; return; }
    var site = p.websiteUri ? p.websiteUri.split('?')[0] : '';
    var dom = domaine_(site);
    if (deja.ids[p.id] || (dom && deja.domaines[dom]) || deja.noms[nomNorm_(nom)]) { ignores++; return; }
    deja.ids[p.id] = 1;
    if (dom) deja.domaines[dom] = 1;
    deja.noms[nomNorm_(nom)] = 1;
    var l = new Array(SRC_ENTETES.length);
    for (var i = 0; i < l.length; i++) l[i] = '';
    l[COL.NOM - 1] = nom;
    l[COL.TYPE - 1] = type;
    l[COL.STATUT - 1] = 'À enrichir';
    l[COL.ADR - 1] = p.formattedAddress || '';
    l[COL.TEL - 1] = p.nationalPhoneNumber || '';
    l[COL.SITE - 1] = site;
    l[COL.SOURCE - 1] = 'Google Maps : ' + requete;
    l[COL.DATE - 1] = date;
    l[COL.ID - 1] = p.id;
    lignes.push(l);
  });

  if (lignes.length) {
    var debut = Math.max(dernLigne_(f), 1) + 1;
    f.getRange(debut, COL.TEL, lignes.length, 1).setNumberFormat('@');
    f.getRange(debut, 1, lignes.length, SRC_ENTETES.length).setValues(lignes);
  }
  return { ajoutes: lignes.length, ignores: ignores };
}

function cleLieux_() {
  var cle = PropertiesService.getScriptProperties().getProperty('PLACES_KEY');
  if (!cle) SpreadsheetApp.getUi().alert('Enregistre d\'abord la clé Google Places (menu Sourcing, étape 1).');
  return cle || '';
}

function dateJour_() {
  return Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy');
}

// Ce que le carnet connaît déjà, pour ne jamais rechercher deux fois la même entreprise.
function connus_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = { ids: {}, domaines: {}, noms: {} };
  var f = ss.getSheetByName(SRC.ONGLET);
  if (f && dernLigne_(f) > 1) {
    f.getRange(2, 1, dernLigne_(f) - 1, SRC_ENTETES.length).getValues().forEach(function (l) {
      if (l[COL.ID - 1]) res.ids[l[COL.ID - 1]] = 1;
      var d = domaine_(l[COL.SITE - 1]); if (d) res.domaines[d] = 1;
      if (l[COL.NOM - 1]) res.noms[nomNorm_(l[COL.NOM - 1])] = 1;
    });
  }
  var p = ss.getSheetByName(SRC.PROSPECTS);
  if (p && p.getLastRow() > 1) {
    p.getRange(2, 1, p.getLastRow() - 1, 8).getValues().forEach(function (l) {
      if (l[0]) res.noms[nomNorm_(l[0])] = 1;
      var d = domaine_(l[7]); if (d) res.domaines[d] = 1;
    });
  }
  return res;
}

// ---------- 4. Enrichir ----------

function src_enrichir() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = ss.getSheetByName(SRC.ONGLET);
  if (!f) { SpreadsheetApp.getUi().alert('Onglet Sourcing introuvable. Lance l\'étape 2.'); return; }
  var debut = Date.now();
  var derniere = dernLigne_(f);
  var faites = 0, restantes = 0;

  for (var l = 2; l <= derniere; l++) {
    if (String(f.getRange(l, COL.STATUT).getValue()) !== 'À enrichir') continue;
    if (Date.now() - debut > SRC.DUREE_MAX_MS) { restantes++; continue; }

    var nom = String(f.getRange(l, COL.NOM).getValue());
    var adresse = String(f.getRange(l, COL.ADR).getValue());
    var site = String(f.getRange(l, COL.SITE).getValue()).trim();
    var cp = (adresse.match(/\b\d{5}\b/) || [])[0];
    var notes = [];

    // Dirigeant légal et SIREN
    var dir = null;
    try { dir = chercherDirigeant_(nom, cp); } catch (e) { notes.push('Annuaire indisponible'); }
    if (dir) {
      f.getRange(l, COL.DECIDEUR).setValue(dir.dirigeants || '');
      f.getRange(l, COL.SIREN).setValue(dir.siren ? "'" + dir.siren : '');
    } else if (notes.indexOf('Annuaire indisponible') < 0) {
      notes.push('Entreprise non retrouvée dans l\'annuaire');
    }

    // Site : email, page contact, constat
    if (!site) {
      f.getRange(l, COL.CONSTAT).setValue('Pas de site');
      f.getRange(l, COL.PISTES).setValue('Pas de site web (à vérifier sur Google)');
    } else {
      try {
        var info = lireSite_(site);
        f.getRange(l, COL.MAIL).setValue(info.emails.join(' ; '));
        f.getRange(l, COL.CONTACT).setValue(info.pageContact);
        f.getRange(l, COL.CONSTAT).setValue(info.constat);
        f.getRange(l, COL.PISTES).setValue(info.pistes.join(', '));
        if (!info.emails.length) notes.push(info.formulaire ? 'Formulaire de contact, pas d\'email affiché' : 'Aucun email trouvé');
      } catch (e) {
        f.getRange(l, COL.CONSTAT).setValue('Site injoignable');
        notes.push('Site injoignable');
      }
    }

    f.getRange(l, COL.NOTE).setValue(notes.join(' ; '));
    f.getRange(l, COL.STATUT).setValue('À qualifier');
    SpreadsheetApp.flush();
    faites++;
    Utilities.sleep(150);
  }
  ss.toast(faites + ' ligne(s) enrichie(s)' + (restantes ? ', ' + restantes + ' restante(s) : relance l\'étape 4.' : '.'), 'Sourcing', 8);
}

function chercherDirigeant_(nom, cp) {
  var url = 'https://recherche-entreprises.api.gouv.fr/search?q=' + encodeURIComponent(nom) +
    (cp ? '&code_postal=' + cp : '') + '&per_page=5';
  var rep = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (rep.getResponseCode() !== 200) throw new Error('Annuaire ' + rep.getResponseCode());
  var resultats = JSON.parse(rep.getContentText()).results || [];
  var mot = motCle_(nom);
  for (var i = 0; i < resultats.length; i++) {
    var r = resultats[i];
    var cible = nomNorm_((r.nom_complet || '') + ' ' + (r.nom_raison_sociale || ''));
    if (!mot || cible.indexOf(mot) === -1) continue;
    var noms = (r.dirigeants || []).slice(0, 2).map(function (d) {
      var base = d.denomination ? d.denomination : ((d.prenoms || '').split(' ')[0] + ' ' + (d.nom || '')).trim();
      return base + (d.qualite ? ' (' + d.qualite + ')' : '');
    });
    return { dirigeants: noms.join(' ; '), siren: r.siren || '' };
  }
  return null;
}

function lireSite_(site) {
  var url = normaliserUrl(site);
  var page = recupererPage(url);
  var dom = domaine_(page.urlUtilisee);
  var tech = detecter(page.html);
  var emails = extraireEmails_(page.html, dom);
  var pageContact = '';
  var formulaire = /<form[\s>]/i.test(page.html);

  if (!emails.length) {
    var liens = trouverLiensContact_(page.html, page.urlUtilisee, dom);
    for (var i = 0; i < liens.length && i < 2 && !emails.length; i++) {
      try {
        var p2 = recupererPage(liens[i]);
        if (!pageContact) pageContact = liens[i];
        emails = extraireEmails_(p2.html, dom);
        if (/<form[\s>]/i.test(p2.html)) formulaire = true;
        if (emails.length) pageContact = liens[i];
      } catch (e) { /* page de contact illisible : on passe */ }
    }
  }

  var constat;
  if (tech.wordpress) {
    constat = 'WordPress' + (tech.constructeur ? ' / ' + tech.constructeur : '') + (tech.versionWp ? ' ' + tech.versionWp : '');
  } else if (tech.autre) {
    constat = tech.autre;
  } else {
    constat = 'Autre ou non identifié';
  }

  var pistes = [];
  if (!page.httpsOk) pistes.push('pas de HTTPS');
  if (tech.wordpress && tech.versionWp && comparerVersions(tech.versionWp.replace(/[^0-9.]/g, ''), CONFIG.WP_VERSION_RECENTE) < 0) pistes.push('WordPress ancien');
  if (tech.nbExtensions >= 15) pistes.push(tech.nbExtensions + ' extensions visibles');
  if (!tech.viewport) pistes.push('pas adapté mobile');
  if (!tech.meta) pistes.push('pas de description Google');
  if (tech.h1 === 0) pistes.push('pas de titre H1');
  if (tech.autre) pistes.push('site ' + tech.autre);

  return { emails: emails, pageContact: pageContact, formulaire: formulaire, constat: constat, pistes: pistes };
}

function extraireEmails_(html, dom) {
  var texte = String(html || '').replace(/&#64;|&commat;|\[at\]|\(at\)/gi, '@');
  var re = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, m, vus = {}, trouves = [];
  while ((m = re.exec(texte)) !== null) {
    var e = m[0].toLowerCase();
    if (/\.(png|jpe?g|gif|svg|webp|css|js)$/.test(e)) continue;
    if (/sentry|wixpress|example\.|domain\.com|votre|@2x|u003e/.test(e)) continue;
    if (!vus[e]) { vus[e] = 1; trouves.push(e); }
  }
  trouves.sort(function (a, b) { return (b.indexOf(dom) > -1 ? 1 : 0) - (a.indexOf(dom) > -1 ? 1 : 0); });
  return trouves.slice(0, 2);
}

function trouverLiensContact_(html, base, dom) {
  var re = /href=["']([^"'#]+)["']/gi, m, vus = {}, liens = [];
  while ((m = re.exec(html)) !== null) {
    var h = m[1];
    if (!/contact|mentions|equipe|a-propos|about/i.test(h)) continue;
    if (/^(mailto|tel|javascript):/i.test(h)) continue;
    var abs = absolutiser_(h, base);
    if (!abs || domaine_(abs) !== dom || vus[abs]) continue;
    vus[abs] = 1;
    liens.push(abs);
  }
  liens.sort(function (a, b) { return (/contact/i.test(b) ? 1 : 0) - (/contact/i.test(a) ? 1 : 0); });
  return liens;
}

function absolutiser_(href, base) {
  if (/^https?:\/\//i.test(href)) return href;
  var racine = base.match(/^https?:\/\/[^\/]+/i);
  if (!racine) return '';
  if (href.indexOf('//') === 0) return base.match(/^https?:/i)[0] + href;
  if (href.charAt(0) === '/') return racine[0] + href;
  return racine[0] + '/' + href;
}

// ---------- 5. Transférer ----------

function src_transferer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = ss.getSheetByName(SRC.ONGLET);
  var p = ss.getSheetByName(SRC.PROSPECTS);
  if (!f || !p) { SpreadsheetApp.getUi().alert('Il faut un onglet Sourcing et un onglet Prospects.'); return; }
  if (dernLigne_(f) < 2) return;

  var donnees = f.getRange(2, 1, dernLigne_(f) - 1, SRC_ENTETES.length).getValues();
  var date = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy');
  var nouvelles = [], pourAudits = [], indices = [];

  donnees.forEach(function (l, i) {
    if (l[COL.STATUT - 1] !== 'Validé') return;
    var type = typeProspect_(l[COL.TYPE - 1]);
    var observations = [l[COL.CONSTAT - 1], l[COL.PISTES - 1], l[COL.NOTE - 1]].filter(function (x) { return x; }).join(' | ');
    // Colonnes de Prospects : Nom, Statut, Prochaine action, Type, Décideur, Email, Téléphone, Site web,
    // Adresse ou ville, Observations, Source de la liste, Canal, Premier contact, Relance 1, Relance 2, Montant estimé
    nouvelles.push([
      l[COL.NOM - 1], 'À contacter', /^(agence|artisan|praticien)/i.test(type) ? '1er mail' : 'Audit flash puis mail', type,
      l[COL.DECIDEUR - 1], String(l[COL.MAIL - 1]).split(' ; ')[0], l[COL.TEL - 1], l[COL.SITE - 1],
      l[COL.ADR - 1], observations, 'Script Sourcing du ' + date + ' (' + l[COL.SOURCE - 1] + ')',
      '', '', '', '', ''
    ]);
    if (!/^agence/i.test(type) && l[COL.SITE - 1]) pourAudits.push(l[COL.SITE - 1]);
    indices.push(i + 2);
  });

  if (!nouvelles.length) { ss.toast('Aucune ligne au statut "Validé".', 'Sourcing', 5); return; }

  var debut = Math.max(p.getLastRow(), 1) + 1;
  p.getRange(debut, 7, nouvelles.length, 1).setNumberFormat('@');
  p.getRange(debut, 1, nouvelles.length, 16).setValues(nouvelles);
  p.getRange(debut, PC.TYPE, nouvelles.length, 1).setDataValidation(reglesType_());

  var audits = ss.getSheetByName(SRC.AUDITS);
  var ajoutsAudits = 0;
  if (audits && pourAudits.length) {
    var existants = {};
    if (audits.getLastRow() > 1) {
      audits.getRange(2, 1, audits.getLastRow() - 1, 1).getValues().forEach(function (r) { existants[domaine_(r[0])] = 1; });
    }
    pourAudits.forEach(function (u) {
      if (existants[domaine_(u)]) return;
      audits.getRange(Math.max(audits.getLastRow(), 1) + 1, 1).setValue(u);
      ajoutsAudits++;
    });
  }

  indices.forEach(function (n) { f.getRange(n, COL.STATUT).setValue('Transféré'); });
  ss.toast(nouvelles.length + ' prospect(s) ajouté(s) à Prospects, ' + ajoutsAudits + ' URL ajoutée(s) à Audits.', 'Sourcing', 8);
}

// ---------- 6 et 7. Ma journée d'envois (agences) ----------
// L'onglet "Aujourd'hui" liste ce qu'il y a à envoyer : relances dues, premiers mails, formulaires.
// Chaque ligne a un mail prêt et un lien "Ouvrir dans Outlook". Tu relis, tu envoies, tu coches "Fait",
// puis l'étape 7 met à jour Prospects (statut, dates, prochaine action).

var ENV = {
  ONGLET: 'Aujourd\'hui',
  PREMIERS_MAX: 10,        // premiers mails par jour (avec email), agences
  ARTISANS_MAX: 10,        // premiers mails par jour, artisans (en plus des agences)
  THERAPEUTES_MAX: 10,     // premiers mails par jour, thérapeutes et praticiens bien-être
  LIEN_RDV_ARTISANS: 'https://exemple.fr/rendez-vous-audit',
  FORMULAIRES_MAX: 5,      // contacts par formulaire ou LinkedIn par jour
  DELAI_JOURS: 7,          // délai entre deux messages
  FROID_JOURS: 90,         // pas de nouveau contact avant ce délai après la relance 2
  LIEN_RDV: 'https://exemple.fr/rendez-vous-partenariat',
  LIEN_MB: 'https://exemple.fr/marque-blanche',
  // Optionnel : si tu changes de lien de rendez-vous, indique la date du changement (ex. new Date(2026, 9, 22)).
  // Les relances de messages plus anciens ajoutent alors une phrase signalant le nouveau lien. Laisser null sinon.
  ANCIEN_LIEN_AVANT: null
};

// Colonnes de l'onglet Prospects (numérotées à partir de 1)
var PC = { NOM: 1, STATUT: 2, ACTION: 3, TYPE: 4, DECIDEUR: 5, MAIL: 6, TEL: 7, SITE: 8, ADR: 9, OBS: 10,
  SOURCE: 11, CANAL: 12, C1: 13, R1: 14, R2: 15, MONTANT: 16 };

function env_aujourdhui() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var p = ss.getSheetByName(SRC.PROSPECTS);
  if (!p || p.getLastRow() < 2) { SpreadsheetApp.getUi().alert('L\'onglet Prospects est vide ou introuvable.'); return; }

  var auj = jourSeul_(new Date());
  var donnees = p.getRange(2, 1, p.getLastRow() - 1, 16).getValues();
  var relances = [], premiers = [], formulaires = [];

  donnees.forEach(function (l, i) {
    var nom = String(l[PC.NOM - 1]).trim();
    if (!nom || !/^(agence|artisan|praticien)/i.test(String(l[PC.TYPE - 1]).trim())) return;
    var estArtisan = deuxMails_(l[PC.TYPE - 1]);
    // Artisans et praticiens : seulement ceux venus du script (colonne Source). Les anciennes lignes saisies à la main restent à part.
    if (estArtisan && !/sourcing/i.test(String(l[PC.SOURCE - 1]))) return;
    var statut = String(l[PC.STATUT - 1]).trim();
    var d1 = parseDate_(l[PC.C1 - 1]), dR1 = parseDate_(l[PC.R1 - 1]), dR2 = parseDate_(l[PC.R2 - 1]);
    var item = {
      ligne: i + 2, nom: nom, mail: premierEmail_(l[PC.MAIL - 1]), decideur: l[PC.DECIDEUR - 1],
      obs: l[PC.OBS - 1], adr: l[PC.ADR - 1], d1: d1, dernier: dR1 || d1,
      artisan: estArtisan, thera: /^praticien/i.test(String(l[PC.TYPE - 1]).trim()), tel: l[PC.TEL - 1], site: l[PC.SITE - 1]
    };

    if (statut === 'À contacter') {
      item.action = item.mail ? 'Premier mail' : 'Premier contact (formulaire ou LinkedIn)';
      (item.mail ? premiers : formulaires).push(item);
    } else if (statut === 'Contacté' && d1 && joursEntre_(d1, auj) >= ENV.DELAI_JOURS) {
      item.action = 'Relance 1';
      relances.push(item);
    } else if (!estArtisan && statut === 'Relancé' && !dR2 && dR1 && joursEntre_(dR1, auj) >= ENV.DELAI_JOURS) {
      item.action = 'Relance 2';
      relances.push(item);
    }
  });

  relances.sort(function (a, b) { return a.dernier - b.dernier; });
  var parType = function (liste, max) {
    var ag = liste.filter(function (x) { return !x.artisan; }).slice(0, max);
    var ar = liste.filter(function (x) { return x.artisan && !x.thera; }).slice(0, ENV.ARTISANS_MAX);
    var th = liste.filter(function (x) { return x.thera; }).slice(0, ENV.THERAPEUTES_MAX);
    return ag.concat(ar, th);
  };
  var premiersJour = parType(premiers, ENV.PREMIERS_MAX);
  var formulairesJour = parType(formulaires, ENV.FORMULAIRES_MAX);
  var aEnvoyer = relances.concat(premiersJour, formulairesJour);

  var f = ss.getSheetByName(ENV.ONGLET) || ss.insertSheet(ENV.ONGLET);
  f.clear();
  f.getRange(1, 1, 1, 8).setValues([['Fait', 'Action', 'Prospect', 'Email', 'Mail à envoyer', 'Ouvrir dans Outlook', 'Ligne Prospects', 'Remarque']])
    .setFontWeight('bold').setBackground('#333333').setFontColor('#ffffff');
  f.setFrozenRows(1);

  if (!aEnvoyer.length) {
    ss.toast('Rien à envoyer aujourd\'hui. Lance le Sourcing (étapes 3 bis, 4, 5) pour remplir Prospects.', 'Sourcing', 8);
    return;
  }

  var valeurs = [], liens = [];
  aEnvoyer.forEach(function (it, i) {
    var r = i + 2;
    var texte;
    if (it.thera) texte = it.action === 'Relance 1' ? mailRelanceTherapeute_(it) : mailPremierTherapeute_(it);
    else if (it.artisan) texte = it.action === 'Relance 1' ? mailRelanceArtisan_(it) : mailPremierArtisan_(it);
    else texte = it.action === 'Relance 1' ? mailRelance_(1, it) : it.action === 'Relance 2' ? mailRelance_(2, it) : mailPremier_(it);
    var remarque = [];
    if (!it.mail) remarque.push(it.artisan
      ? 'Pas d\'email : appeler (' + (it.tel || 'pas de numéro') + ') ou passer par le formulaire du site. Le mail sert alors de trame.'
      : 'Pas d\'email : formulaire du site, LinkedIn ou téléphone');
    if (!it.artisan && (it.action === 'Relance 1' || it.action === 'Relance 2') && ancienLien_(it)) remarque.push('Le message contient la phrase sur le nouveau lien de rendez-vous');
    valeurs.push([false, it.action, it.nom, it.mail, texte, '', it.ligne, remarque.join(' / ')]);
    liens.push([lienMail_(it.mail, texte)]);
  });

  f.getRange(2, 1, valeurs.length, 8).setValues(valeurs);
  f.getRange(2, 6, liens.length, 1).setRichTextValues(liens);
  f.getRange(2, 1, valeurs.length, 1).insertCheckboxes();
  f.getRange(2, 1, valeurs.length, 8).setVerticalAlignment('top');
  f.getRange(2, 5, valeurs.length, 1).setWrap(true);
  f.getRange(2, 8, valeurs.length, 1).setWrap(true);
  f.getRange(2, 7, valeurs.length, 1).setFontColor('#999999');
  f.setColumnWidth(1, 50); f.setColumnWidth(2, 150); f.setColumnWidth(3, 200); f.setColumnWidth(4, 220);
  f.setColumnWidth(5, 480); f.setColumnWidth(6, 150); f.setColumnWidth(7, 70); f.setColumnWidth(8, 260);
  f.setRowHeights(2, valeurs.length, 120);

  var surplus = (premiers.length - premiersJour.length) + (formulaires.length - formulairesJour.length);
  ss.toast(relances.length + ' relance(s), ' + premiersJour.length + ' premier(s) mail(s), ' +
    formulairesJour.length + ' formulaire(s)/appel(s)' + (surplus ? ' (' + surplus + ' en attente pour demain)' : '') + '.',
    'Aujourd\'hui', 10);
  ss.setActiveSheet(f);
}

function env_marquerFaits() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = ss.getSheetByName(ENV.ONGLET);
  var p = ss.getSheetByName(SRC.PROSPECTS);
  if (!f || !p || f.getLastRow() < 2) { SpreadsheetApp.getUi().alert('Lance d\'abord l\'étape 6.'); return; }

  var auj = jourSeul_(new Date());
  var lignes = f.getRange(2, 1, f.getLastRow() - 1, 8).getValues();
  var n = 0;

  lignes.forEach(function (l, i) {
    if (l[0] !== true || /^Fait le/.test(String(l[7]))) return;
    var ligne = Number(l[6]);
    if (!ligne || String(p.getRange(ligne, PC.NOM).getValue()).trim() !== String(l[2]).trim()) return; // sécurité : la ligne a bougé
    var action = String(l[1]);

    if (/^Premier/.test(action)) {
      p.getRange(ligne, PC.STATUT).setValue('Contacté');
      p.getRange(ligne, PC.ACTION).setValue((deuxMails_(p.getRange(ligne, PC.TYPE).getValue()) ? 'Mail 2 le ' : 'Relance 1 le ') + formatCourt_(ajouterJours_(auj, ENV.DELAI_JOURS)) + ' si pas de réponse');
      p.getRange(ligne, PC.CANAL).setValue(l[3] ? 'Email' : (deuxMails_(p.getRange(ligne, PC.TYPE).getValue()) ? 'Téléphone ou formulaire' : 'Formulaire ou LinkedIn'));
      ecrireDate_(p.getRange(ligne, PC.C1), auj);
    } else if (action === 'Relance 1' && deuxMails_(p.getRange(ligne, PC.TYPE).getValue())) {
      // Artisans et praticiens : 2 messages seulement (mail 1, puis mail 2 une semaine plus tard).
      p.getRange(ligne, PC.STATUT).setValue('Froid');
      p.getRange(ligne, PC.ACTION).setValue('Aucune avant le ' + formatCourt_(ajouterJours_(auj, ENV.FROID_JOURS)));
      ecrireDate_(p.getRange(ligne, PC.R1), auj);
    } else if (action === 'Relance 1') {
      p.getRange(ligne, PC.STATUT).setValue('Relancé');
      p.getRange(ligne, PC.ACTION).setValue('Relance 2 le ' + formatCourt_(ajouterJours_(auj, ENV.DELAI_JOURS)) + ' si pas de réponse');
      ecrireDate_(p.getRange(ligne, PC.R1), auj);
    } else if (action === 'Relance 2') {
      p.getRange(ligne, PC.STATUT).setValue('Froid');
      p.getRange(ligne, PC.ACTION).setValue('Aucune avant le ' + formatCourt_(ajouterJours_(auj, ENV.FROID_JOURS)));
      ecrireDate_(p.getRange(ligne, PC.R2), auj);
    } else { return; }

    f.getRange(i + 2, 8).setValue('Fait le ' + formatCourt_(auj));
    f.getRange(i + 2, 1, 1, 8).setBackground('#e6e6e6');
    n++;
  });
  ss.toast(n + ' envoi(s) enregistré(s) dans Prospects.', 'Aujourd\'hui', 6);
}

// ---------- Les mails ----------

function signature_() {
  return CONFIG.SIGNATURE;   // définie dans Code.gs
}

function bonjour_(it) {
  var pr = prenom_(it.decideur);
  return 'Bonjour' + (pr ? ' ' + pr : '') + ',';
}

function nomCourt_(nom) {
  return String(nom).replace(/\s+-\s+.*$/, '').trim();
}

function mailPremier_(it) {
  return 'Objet : Un renfort WordPress pour ' + nomCourt_(it.nom) + '\n\n' +
    bonjour_(it) + '\n\n' +
    'Je m\'appelle ' + CONFIG.NOM + ', ' + CONFIG.PRESENTATION + '.\n\n' +
    accroche_(it.obs) + '\n\n' +
    'Je propose aux agences un renfort en marque blanche : création et refonte de sites sous Elementor, boutiques WooCommerce, ' +
    'SEO technique et maintenance. Mes sites sont éco-conçus, donc rapides, sobres et pensés pour le référencement.\n\n' +
    'Mon CV est en pièce jointe, et le détail de la collaboration est ici : ' + ENV.LIEN_MB + '\n\n' +
    proposition_(it) + '\n\n' +
    'Bien à vous,\n\n' + signature_() + '\n\n' +
    'Si ce message ne vous concerne pas, un simple mot suffit et je ne vous recontacterai pas.';
}

function mailRelance_(numero, it) {
  var nouveauLien = ancienLien_(it)
    ? 'Mon lien de réservation a changé : celui-ci remplace celui de mon premier message.\n\n' : '';
  if (numero === 1) {
    return 'Objet : Re : Un renfort WordPress pour ' + nomCourt_(it.nom) + '\n\n' +
      bonjour_(it) + '\n\n' +
      'Je me permets de revenir vers vous au sujet de mon message du ' + formatCourt_(it.d1) + '.\n\n' +
      'Pour rappel, je propose aux agences un renfort WordPress en marque blanche, avec une grille tarifaire dégressive selon le volume et sans minimum mensuel.\n\n' +
      'Une visio de 15 à 30 minutes vous conviendrait-elle ? ' + ENV.LIEN_RDV + '\n\n' + nouveauLien +
      'Bien à vous,\n\n' + signature_();
  }
  return 'Objet : Dernier message de ma part, ' + nomCourt_(it.nom) + '\n\n' +
    bonjour_(it) + '\n\n' +
    'Sans retour de votre part, je ne vous relancerai plus. Si un pic d\'activité ou un projet WordPress à déléguer se présente, ' +
    'vous pouvez me joindre directement ou réserver un créneau ici : ' + ENV.LIEN_RDV + '\n\n' + nouveauLien +
    'Bien à vous,\n\n' + signature_();
}

// ---------- Les mails artisans (vouvoiement, 2 messages : mail 1, puis mail 2 sept jours plus tard) ----------

// Une seule accroche, la plus utile, d'après ce que l'étape 4 a repéré sur le site.
function accrocheArtisan_(it) {
  var obs = String(it.obs || '');
  if (!String(it.site || '').trim() || /pas de site/i.test(obs)) {
    return 'Je n\'ai pas trouvé de site internet pour ' + nomCourt_(it.nom) + '. Quand un client cherche un artisan sur Google, ' +
      'il regarde d\'abord les avis, puis le site : quelques pages claires, vos réalisations et un bouton pour vous appeler suffisent à le rassurer.';
  }
  if (/pas de https/i.test(obs)) {
    return 'En visitant votre site, j\'ai vu que votre navigateur le signale comme « non sécurisé ». Cela peut faire fuir un visiteur avant même qu\'il vous ait contacté, et se corrige rapidement.';
  }
  if (/wordpress ancien/i.test(obs)) {
    return 'Votre site est sous WordPress, mais dans une version ancienne. C\'est une cause fréquente de lenteur et de piratage, et ça se remet en ordre sans tout refaire.';
  }
  if (/pas adapté mobile/i.test(obs)) {
    return 'En visitant votre site, j\'ai remarqué qu\'il ne s\'adapte pas bien à l\'écran d\'un téléphone. Or une grande partie de vos futurs clients vous cherche depuis leur mobile.';
  }
  if (/extensions visibles/i.test(obs)) {
    return 'Votre site WordPress semble chargé en extensions. C\'est souvent ce qui le ralentit et qui le rend plus fragile face au piratage.';
  }
  if (/pas de description google/i.test(obs)) {
    return 'Votre site n\'a pas de description pour Google : dans les résultats de recherche, il se présente donc moins bien qu\'il ne le pourrait.';
  }
  return 'J\'ai regardé votre site : il y a probablement quelques points simples à améliorer pour qu\'il vous amène plus d\'appels.';
}

function mailPremierArtisan_(it) {
  var sans = !String(it.site || '').trim() || /pas de site/i.test(String(it.obs || ''));
  var offre = sans
    ? 'Je crée des sites simples, rapides et clairs pour les artisans : vos prestations, vos réalisations, votre zone d\'intervention et un accès direct pour vous joindre. Je vous accompagne pour que vous puissiez ensuite les modifier vous-même.'
    : 'Je propose un audit flash gratuit : en quelques minutes, je vous montre ce qui ralentit votre site, ce qui limite votre visibilité sur Google, et ce qui peut être corrigé rapidement. Sans engagement.';
  return 'Objet : ' + (sans ? 'Un site pour ' : 'Votre site internet, ') + nomCourt_(it.nom) + '\n\n' +
    bonjour_(it) + '\n\n' +
    'Je m\'appelle ' + CONFIG.NOM + ', ' + CONFIG.PRESENTATION + '. J\'aide les artisans et entreprises locales à avoir un site qui travaille pour eux.\n\n' +
    accrocheArtisan_(it) + '\n\n' +
    offre + '\n\n' +
    'Seriez-vous disponible pour une visio de 15 à 30 minutes ? ' + ENV.LIEN_RDV_ARTISANS + '\n' +
    'Vous pouvez aussi me répondre directement ou m\'appeler : je m\'adapte à vos horaires de chantier.\n\n' +
    'Bien à vous,\n\n' + signature_() + '\n\n' +
    'Si ce message ne vous concerne pas, un simple mot suffit et je ne vous recontacterai pas.';
}

function mailRelanceArtisan_(it) {
  return 'Objet : Re : ' + (/pas de site/i.test(String(it.obs || '')) || !String(it.site || '').trim() ? 'Un site pour ' : 'Votre site internet, ') + nomCourt_(it.nom) + '\n\n' +
    bonjour_(it) + '\n\n' +
    'Je me permets de revenir vers vous au sujet de mon message du ' + formatCourt_(it.d1) + '. Je sais que vos journées sont chargées sur les chantiers.\n\n' +
    'Pour vous donner une idée de ma façon de travailler : je remets en ordre des sites lents ou piratés, ' +
    'et je crée des sites éco-conçus, donc rapides et simples à faire vivre.\n\n' +
    'Si la question de votre site se pose, un échange de 15 minutes suffit pour faire le point : ' + ENV.LIEN_RDV_ARTISANS + '\n\n' +
    'Sans retour de votre part, je ne vous relancerai plus. Vous gardez mes coordonnées si le besoin se présente un jour.\n\n' +
    'Bien à vous,\n\n' + signature_();
}

// Cycle à deux messages (mail 1, puis mail 2 sept jours plus tard) pour les artisans et les praticiens bien-être.
function deuxMails_(type) {
  return /^(artisan|praticien)/i.test(String(type || '').trim());
}

// ---------- Les mails thérapeutes et praticiens bien-être (vouvoiement, 2 messages) ----------

function sansSite_(it) {
  return !String(it.site || '').trim() || /pas de site/i.test(String(it.obs || ''));
}

function accrocheTherapeute_(it) {
  var obs = String(it.obs || '');
  if (sansSite_(it)) {
    return 'Je n\'ai pas trouvé de site internet pour ' + nomCourt_(it.nom) + '. Quand une personne cherche un praticien près de chez elle, ' +
      'un site clair est souvent ce qui la rassure avant de prendre contact : qui vous êtes, comment vous travaillez, et comment réserver une séance.';
  }
  if (/pas de https/i.test(obs)) {
    return 'En visitant votre site, j\'ai vu que votre navigateur le signale comme « non sécurisé ». Pour quelqu\'un qui hésite à vous contacter, ce message peut freiner, et il se corrige rapidement.';
  }
  if (/pas adapté mobile/i.test(obs)) {
    return 'En visitant votre site, j\'ai remarqué qu\'il ne s\'adapte pas bien à l\'écran d\'un téléphone. C\'est pourtant souvent depuis son mobile qu\'on cherche un praticien et qu\'on prend rendez-vous.';
  }
  if (/wordpress ancien|extensions visibles/i.test(obs)) {
    return 'Votre site est sous WordPress, mais il semble avoir besoin d\'une remise en ordre (version ancienne ou nombreuses extensions). C\'est une cause fréquente de lenteur et de fragilité, et ça se corrige sans tout refaire.';
  }
  if (/pas de description google/i.test(obs)) {
    return 'Votre site n\'a pas de description pour Google : dans les résultats de recherche, votre pratique se présente donc moins bien qu\'elle ne le pourrait.';
  }
  return 'J\'ai regardé votre site : il y a probablement quelques points simples à améliorer pour que les personnes qui vous cherchent vous trouvent plus facilement.';
}

function mailPremierTherapeute_(it) {
  return 'Objet : ' + (sansSite_(it) ? 'Un site pour ' : 'Votre visibilité en ligne, ') + nomCourt_(it.nom) + '\n\n' +
    bonjour_(it) + '\n\n' +
    'Je m\'appelle ' + CONFIG.NOM + ', ' + CONFIG.PRESENTATION + '. J\'accompagne les praticiens et thérapeutes pour que les personnes qui les cherchent les trouvent, et puissent prendre rendez-vous simplement.\n\n' +
    accrocheTherapeute_(it) + '\n\n' +
    'Selon vos besoins et vos envies, je peux vous aider sur un ou plusieurs points : un site ergonomique, pensé pour présenter votre pratique et garder l\'attention de vos visiteurs, ' +
    'une fiche Google bien renseignée pour apparaître dans votre secteur, et une prise de rendez-vous en ligne. ' +
    'Je m\'adapte à ce que vous souhaitez pour votre visibilité, sans rien vous imposer.\n\n' +
    'Seriez-vous disponible pour une visio de 15 à 30 minutes, pour que je comprenne votre pratique et la visibilité que vous souhaitez ? ' + ENV.LIEN_RDV_ARTISANS + '\n' +
    'Vous pouvez aussi me répondre directement à ce message.\n\n' +
    'Bien à vous,\n\n' + signature_() + '\n\n' +
    'Si ce message ne vous concerne pas, un simple mot suffit et je ne vous recontacterai pas.';
}

function mailRelanceTherapeute_(it) {
  return 'Objet : Re : ' + (sansSite_(it) ? 'Un site pour ' : 'Votre visibilité en ligne, ') + nomCourt_(it.nom) + '\n\n' +
    bonjour_(it) + '\n\n' +
    'Je me permets de revenir vers vous au sujet de mon message du ' + formatCourt_(it.d1) + '. Je sais que votre temps auprès des personnes que vous accompagnez passe avant tout.\n\n' +
    'Je crée des sites éco-conçus, donc rapides et sobres, que vous pourrez ensuite faire vivre vous-même : je vous accompagne pour que vous soyez autonome.\n\n' +
    'Si la visibilité de votre pratique est un sujet pour vous, un échange de 15 minutes suffit pour faire le point : ' + ENV.LIEN_RDV_ARTISANS + '\n\n' +
    'Sans retour de votre part, je ne vous relancerai plus. Vous gardez mes coordonnées si le besoin se présente un jour.\n\n' +
    'Bien à vous,\n\n' + signature_();
}

// Agence des Alpes-Maritimes ou du Var (code postal 06 ou 83) : on propose un café. Sinon, une visio.
function estLocale_(it) {
  return /\b(06|83)\d{3}\b/.test(String(it.adr || ''));
}

function proposition_(it) {
  if (estLocale_(it)) {
    return 'Nous travaillons tous derrière un écran, mais les bonnes collaborations commencent par une vraie rencontre. ' +
      'Nous sommes voisins : je serais ravie de vous retrouver autour d\'un café, chez vous ou à l\'endroit qui vous arrange. ' +
      'Si vous préférez commencer par une visio de 15 à 30 minutes : ' + ENV.LIEN_RDV;
  }
  return 'Seriez-vous ouverts à une visio de 15 à 30 minutes ? ' + ENV.LIEN_RDV;
}

function accroche_(obs) {
  if (/wordpress/i.test(String(obs || ''))) {
    return 'Votre site est construit sous WordPress, la technologie que je maîtrise : je peux absorber une partie de votre production lors des pics d\'activité, sous votre nom.';
  }
  return 'Votre agence conçoit des sites pour ses clients : je peux absorber une partie de la production WordPress lorsque votre équipe est à pleine charge, sous votre nom.';
}

function ancienLien_(it) {
  return !!(ENV.ANCIEN_LIEN_AVANT && it.d1 && it.d1 < ENV.ANCIEN_LIEN_AVANT);
}

// Prénom seulement quand le dirigeant est une personne unique au format "Prénom NOM (Qualité)".
function prenom_(decideur) {
  var d = String(decideur || '').trim();
  if (!d || d.indexOf(';') > -1) return '';
  if (/\b(SAS|SARL|EURL|SASU|SCI|SA|HOLDING|GROUPE?)\b/i.test(d)) return '';
  var m = d.match(/^([A-Za-zÀ-ÿ'’\-]{2,})\s+\S.*\(/);
  if (!m) return '';
  return m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase();
}

// ---------- Dates et outils ----------

function premierEmail_(v) {
  var m = String(v || '').match(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/);
  return m ? m[0] : '';
}

function jourSeul_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function parseDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return jourSeul_(v);
  var m = String(v || '').match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

function joursEntre_(a, b) {
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

function ajouterJours_(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}

function formatCourt_(d) {
  if (!d) return '';
  var j = d.getDate(), m = d.getMonth() + 1;
  return (j < 10 ? '0' : '') + j + '/' + (m < 10 ? '0' : '') + m;
}

function ecrireDate_(cellule, d) {
  cellule.setNumberFormat('dd/MM/yyyy');
  cellule.setValue(d);
}

// ---------- Outils ----------

function domaine_(url) {
  return String(url || '').toLowerCase().replace(/^https?:\/\/(www\.)?/, '').replace(/[\/?#].*$/, '').trim();
}

function nomNorm_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
}

// Premier mot distinctif du nom, pour vérifier qu'on a bien retrouvé la bonne entreprise.
function motCle_(nom) {
  var banal = ['agence', 'web', 'digital', 'digitale', 'communication', 'studio', 'groupe', 'sas', 'sarl', 'les', 'des', 'the'];
  var mots = nomNorm_(nom).split(' ');
  for (var i = 0; i < mots.length; i++) {
    if (mots[i].length >= 3 && banal.indexOf(mots[i]) === -1) return mots[i];
  }
  return mots[0] || '';
}

// Dernière ligne réellement remplie dans la colonne A (les formules des colonnes Q et R ne comptent pas).
function dernLigne_(f) {
  var v = f.getRange(1, 1, f.getMaxRows(), 1).getValues();
  for (var i = v.length - 1; i >= 0; i--) { if (String(v[i][0]).trim() !== '') return i + 1; }
  return 1;
}

// Lien "Ouvrir dans Outlook" : texte enrichi (fonctionne quelle que soit la langue du Sheet).
function lienMail_(email, texte) {
  var vide = SpreadsheetApp.newRichTextValue().setText('').build();
  if (!email) return vide;
  var m = String(texte).match(/^Objet : ([^\n]+)\n\n([\s\S]*)$/);
  var objet = m ? m[1] : '';
  var corps = m ? m[2] : String(texte);
  var url = 'mailto:' + email + '?subject=' + encodeURIComponent(objet) + '&body=' + encodeURIComponent(corps);
  return SpreadsheetApp.newRichTextValue().setText('Ouvrir dans Outlook').setLinkUrl(url).build();
}