# Audit flash et Sourcing

Deux scripts Google Apps Script, à installer dans un même Google Sheet, pour **repérer des prospects, analyser leur site WordPress en quelques secondes et préparer des mails de prospection personnalisés** à relire avant envoi.

Conçu par Alexia Gombert ([Ta Webmaster](https://tawebmaster.com)) avec Claude Code, pour structurer sa prospection quotidienne.

## Ce que fait chaque script

### `Code.gs` : Audit flash
Pour chaque URL collée dans l'onglet **Audits** :
- **Performance et SEO** via l'API PageSpeed Insights (mobile) : scores, LCP, TBT, CLS, poids de la page
- **Détection technique** par lecture du code public : WordPress, version, thème, constructeur de pages, extensions visibles, HTTPS, viewport mobile, meta description, nombre de H1
- **Symptômes** rédigés en langage clair, et **brouillon de mail** basé sur les trois principaux constats

### `Sourcing.gs` : sourcing et suivi
Menu **Sourcing** dans le Sheet :
1. Recherche d'entreprises via Google Maps (agences, artisans, praticiens bien-être), par métier et par commune, avec reprise là où le lot s'est arrêté
2. Enrichissement : dirigeant légal (annuaire public des entreprises), email et page contact trouvés sur le site, constat technique
3. Transfert des lignes validées vers l'onglet **Prospects**
4. Onglet **Aujourd'hui** : relances dues, premiers mails du jour, avec un lien « Ouvrir dans Outlook » ; cases à cocher pour enregistrer les envois et mettre à jour les statuts et les dates

Aucun mail n'est envoyé automatiquement : tout est relu et envoyé à la main.

Les scripts **ne modifient jamais** les sites analysés : ils lisent uniquement des pages publiques.

## Installation

1. Créer un Google Sheet, puis **Extensions → Apps Script**.
2. Créer deux fichiers de script, `Code` et `Sourcing`, et y coller le contenu des fichiers du même nom. Enregistrer, recharger le Sheet.
3. Adapter le bloc `CONFIG` en haut de `Code.gs` (nom, présentation, signature, lien de rendez-vous) et le bloc `ENV` en haut de la partie envois de `Sourcing.gs` (liens, limites par jour).
4. Menu **Audit flash** : enregistrer la clé PageSpeed, préparer l'onglet Audits.
5. Menu **Sourcing** : enregistrer la clé Google Places, préparer l'onglet Sourcing.
6. Créer à la main un onglet **Prospects** avec les colonnes : Nom, Statut, Prochaine action, Type, Décideur, Email, Téléphone, Site web, Adresse ou ville, Observations, Source de la liste, Canal, Date du 1er contact, Date relance 1, Date relance 2, Montant estimé.

Les clés d'API sont stockées dans les propriétés du script, jamais dans le code ni dans le Sheet.

## À adapter

- Les zones géographiques (communes des Alpes-Maritimes et du Var), les métiers et les pratiques recherchés
- Le texte des mails, écrits pour une activité d'architecte digitale WordPress
- La règle « proposer un café » pour les agences voisines (codes postaux 06 et 83)

## Limites

- Quelques sites par lancement, pour rester sous la limite d'exécution de 6 minutes d'Apps Script
- La détection des extensions et des versions repose sur le code visible de la page : c'est une indication, pas un audit exhaustif
- Sans clé PageSpeed, Google limite très vite les requêtes

## Usage responsable

- Ces scripts collectent et conservent des données professionnelles (noms de dirigeants, emails, téléphones). **Ne publie jamais ton Sheet** : seul le code est destiné à être partagé.
- Pour de la prospection B2B, identifie clairement l'expéditeur et propose un moyen de ne plus être contacté, comme dans les modèles fournis. Respecte le RGPD et les demandes de désinscription.
- Vérifie les conditions d'utilisation de l'API Google Places concernant la conservation des données issues de Google Maps.

## Licence

MIT.
