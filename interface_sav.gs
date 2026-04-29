/**
 * Application Web SAV — Google Apps Script
 * Déployer comme "Application Web" (exécuter en tant que : Moi).
 * Renseigner SAV_SPREADSHEET_ID avec l’ID du classeur (dans l’URL Google Sheets).
 *
 * Feuilles créées / vérifiées automatiquement :
 * - SAV 2026 (distributeurs)
 * - SAV Marketplace
 * - Demande d'Avoir
 * - Réception Atelier
 * - Envoi/Transport
 * - Logistique (une ligne par demande transport — vue équipe logistique)
 * - SAV_Modeles (liste pour le dropdown « modèle »)
 */

const SAV_SPREADSHEET_ID = "1t4RZPX5JJR81LeLfvX14sbBZq1Iq3wmVeAwrehwKnb8";
// Référentiel de production (Usines / Modèles) : classeur "calendrier de production"
// Si tes modèles sont gérés dans un autre classeur que le SAV, mets son ID ici.
const SAV_PROD_SPREADSHEET_ID = "1koOpvCtCXbQZcahM4RDLp4vY4ZJW77FUp4XAD7kj8X4";
const SAV_APP_VERSION = "sav-2026-04-29d";

const SAV_SHEET_DIST = "SAV 2026";
const SAV_SHEET_MP = "SAV Marketplace";
const SAV_SHEET_REC = "Réception Atelier";
const SAV_SHEET_EXP = "Expertise";
const SAV_SHEET_TR = "Envoi/Transport";
const SAV_SHEET_LOG = "Logistique";
// Logistique — "Autres" : arrivées / envois hors dossiers SAV (simple traçabilité de transit)
const SAV_SHEET_LOG_OTHER = "Logistique_Autres";
const SAV_SHEET_PDC = "Pièces détachées";
const SAV_SHEET_MODELS = "SAV_Modeles";
const SAV_SHEET_RDV_VISIO = "SAV_RDV_Visio";
const SAV_SHEET_PROCEDURES = "SAV_Procedures";
const SAV_SHEET_PROCEDURE_MODELS = "SAV_ProcedureModels";
const SAV_SHEET_MP_MESSAGES = "SAV_MP_Messages";
const SAV_SHEET_AVOIR_TICKETS = "Tickets Avoir";
const SAV_SHEET_ANALYTICS = "SAV_Analytics";
const SAV_SHEET_ANALYTICS_VOLUMES = "SAV_Analytics_Volumes";
const SAV_SHEET_ANALYTICS_YEAR_TICKETS = "SAV_Analytics_YearTickets";
/** Feuille « production » (même classeur que le calendrier) : colonnes Nom | Usine — une ligne = un modèle (toutes usines). */
const SAV_SHEET_MODELES_PROD = "Modeles";
const SAV_SHEET_MODELES_PROD_ALT = "Modèles";

// Dossier racine Drive unique — tout est créé automatiquement à l’intérieur
const SAV_DRIVE_ROOT_FOLDER_ID = "1XrMTHEcn2gasLwJR95zF4vi8iGvRo5KC";

// Noms des sous-dossiers créés automatiquement dans SAV_DRIVE_ROOT_FOLDER_ID
const SAV_DRIVE_SUBDIR_DIST  = "SAV Distributeur";
const SAV_DRIVE_SUBDIR_MP    = "SAV Marketplace";
const SAV_DRIVE_SUBDIR_PDC   = "PDC";
const SAV_DRIVE_SUBDIR_MISC  = "Logistique & Autres";

// Clés de cache pour les IDs de sous-dossiers (évite de rescanner à chaque appel)
const SAV_DRIVE_CACHE_DIST  = "sav:drive:dist";
const SAV_DRIVE_CACHE_MP    = "sav:drive:mp";
const SAV_DRIVE_CACHE_PDC   = "sav:drive:pdc";
const SAV_DRIVE_CACHE_MISC  = "sav:drive:misc";

// Logo Optimea (pour le site + entêtes PDF)
// Dossier fourni par toi : 1Ugcb_LwnLfr8VSu6Ph_OlNt46IBiyn7n
const SAV_LOGO_DRIVE_FOLDER_ID = "1Ugcb_LwnLfr8VSu6Ph_OlNt46IBiyn7n";
const SAV_LOGO_CACHE_KEY = "sav:logo:dataUrl";

function savGetLogoDataUrl_() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(SAV_LOGO_CACHE_KEY);
    if (cached) return String(cached);
  } catch (e) {}

  const folder = DriveApp.getFolderById(SAV_LOGO_DRIVE_FOLDER_ID);
  const files = folder.getFiles();
  let best = null;
  let bestScore = -1;

  function scoreFile_(name) {
    const n = String(name || "").toLowerCase();
    let s = 0;
    if (n.indexOf("logo") >= 0) s += 100;
    if (n.indexOf("optimea") >= 0) s += 50;
    // prefer common raster formats
    if (n.endsWith(".png")) s += 20;
    if (n.endsWith(".jpg") || n.endsWith(".jpeg")) s += 15;
    return s;
  }

  while (files.hasNext()) {
    const f = files.next();
    const mime = String(f.getMimeType ? f.getMimeType() : "");
    const name = String(f.getName ? f.getName() : "");
    const lower = name.toLowerCase();
    // ignore non-raster formats to avoid HTML rendering issues
    const ok =
      lower.endsWith(".png") || lower.endsWith(".jpg") || lower.endsWith(".jpeg") || lower.endsWith(".webp");
    if (!ok) continue;
    const s = scoreFile_(name);
    if (s > bestScore) {
      bestScore = s;
      best = f;
    }
  }

  if (!best) return "";

  const blob = best.getBlob();
  const bytes = blob.getBytes();
  const b64 = Utilities.base64Encode(bytes);
  const mimeType = String(blob.getContentType ? blob.getContentType() : "image/png");
  const dataUrl = "data:" + mimeType + ";base64," + b64;

  try {
    CacheService.getScriptCache().put(SAV_LOGO_CACHE_KEY, dataUrl, 24 * 60 * 60);
  } catch (e) {}

  return dataUrl;
}

/**
 * Exposée côté client via google.script.run pour charger le logo en asynchrone.
 * Ne jamais appeler dans doGet().
 */
function savGetLogoUrl() {
  try { return savGetLogoDataUrl_(); } catch(e) { return ""; }
}

function savPdfLogoHtml_() {
  const dataUrl = savGetLogoDataUrl_();
  if (!dataUrl) return "";
  // Centered, no clipping: fixed max sizes + display:block
  return (
    '<div style="display:flex;justify-content:center;align-items:center;margin:0 0 10px 0;">' +
    '<img src="' +
    dataUrl +
    '" style="max-width:170px;max-height:56px;width:auto;height:auto;display:block;object-fit:contain;" />' +
    "</div>"
  );
}

const HDR_DIST = [
  "Numéro de dossier",
  "Nom magasin/distributeur",
  "Adresse complète",
  "Modèle concerné",
  "Panne constatée",
  "Numéro de série",
  "Facture reçue",
  "Photo plaque reçue",
  "Date création",
  "État du dossier",
  "Dossier Drive",
  "Date clôture",
  "Mail envoyé au distributeur (accord retour atelier)",
  "Date mail accord retour atelier",
  "Référence dossier distributeur",
  "Demandes PDC",
];

const HDR_MP = [
  "Numéro de dossier",
  "Nom marketplace",
  "Nom client",
  "Email client",
  "Téléphone client",
  "Adresse client",
  "Demande d'enlèvement",
  "Modèle concerné",
  "Panne constatée",
  "Numéro facture",
  "Photo N°S reçue",
  "Date création",
  "État du dossier",
  "Dossier Drive",
  "Date clôture",
  "Demandes PDC",
];

// Module "Avoir" supprimé à la demande.
// (Anciennes constantes SAV_SHEET_AVOIR / HDR_AVOIR supprimées.)

const HDR_REC = [
  "Numéro de dossier SAV",
  "Date réception",
  "État visuel appareil",
  "Éléments accompagnants",
  "Notes techniques",
  "Technicien responsable",
];

const HDR_EXP = [
  "Numéro de dossier SAV",
  "Date expertise",
  "Technicien",
  "Travaux réalisés",
  "Pièces utilisées",
  "Tests effectués",
  "Conclusions",
  "Décision",
  "Coût pièces (€)",
  "Coût main d'œuvre (€)",
];

const HDR_TR = [
  "Numéro de dossier SAV",
  "Type envoi",
  "Point départ",
  "Date demande d'envoi",
  "Numéro tracking",
  "Statut",
  "Preuve livraison reçue",
  "Date livraison",
  "Adresse arrivée",
];

/** Une ligne = un « bon » logistique ; lien vers le dossier SAV (feuille + ligne). */
const HDR_LOG = [
  "Id",
  "N° dossier SAV",
  "Feuille dossier",
  "Ligne dossier",
  "Type envoi",
  "Point départ",
  "Résumé",
  "Modèle",
  "Statut logistique",
  "Date demande",
  "Tracking",
  "Date expédition",
  "Transporteur",
  "Coût transport (€)",
  "Notes logistique",
  "Preuve livraison",
  "Date livraison",
  "Colis prêt (SAV)",
  "Date colis prêt",
  "Adresse arrivée",
];

/** Logistique — autres arrivées / envois (hors dossiers SAV). */
const HDR_LOG_OTHER = [
  "Id",
  "Sens", // ARRIVÉE | ENVOI
  "Date",
  "Référence / Client",
  "Point départ",
  "Adresse arrivée",
  "Transporteur",
  "Tracking",
  "Notes",
];

/** Une ligne = un envoi de pièce détachée (hors dossiers SAV). */
const HDR_PDC = [
  "Id",
  "Date création",
  "Poids (kg)",
  "Transporteur",
  "Tracking",
  "Statut",
  "Date expédition",
  "Date livraison",
  "Notes",
  "Colis prêt",
  "Point départ",
  "Adresse arrivée",
  "Références",
  "Désignation",
  "Dossier SAV (optionnel)",
  "Dossier Drive",
];

const LOG_STAT_A_TRAITER = "À traiter";
const LOG_STAT_EXPEDIE = "Expédié";
const LOG_STAT_LIVRE = "Livré";
const LOG_STAT_ANNULE = "Annulé";

// Procédure 3 — RDV visio
const HDR_RDV_VISIO = ["N° dossier SAV", "Type", "Modèle", "Client", "Email", "Téléphone", "Date RDV", "Heure RDV", "État", "Notes", "Créé le", "Dernière maj"];

// SAV_Modeles: colonnes A/B/C = modèles classés par procédure 1/2/3
const HDR_MODELS = ["procedure_1", "procedure_2", "procedure_3"];
const HDR_PROCEDURES = ["Id", "Libellé", "Objectif", "Documentation", "Étapes", "Actif", "Dernière maj"];
const HDR_PROCEDURE_MODELS = ["Modèle", "ProcédureId", "Dernière maj"];
const HDR_MP_MESSAGES = ["Id", "N° dossier SAV", "Marketplace", "Client email", "Sens", "Message", "Auteur", "Date", "Lu", "Notes internes"];

const HDR_ANALYTICS = ["Généré le", "Section", "Clé", "Nombre", "%", "Base (dossiers)"];
const HDR_ANALYTICS_VOLUMES = ["Année", "Référence", "Volume produits"];
const HDR_ANALYTICS_YEAR_TICKETS = ["Année", "Créé le", "Statut", "Notes"];

const HDR_AVOIR_TICKETS = [
  "Numéro de dossier",
  "Magasin",
  "Adresse",
  "Modèle",
  "Panne",
  "Numéro de série",
  "Type",
  "Date création",
  "Statut",
  "Dossier Drive",
  "Centrale d'achat",
  "Facture achat centrale",
  "Notes",
  "Feuille dossier",
  "Ligne dossier",
];

// Emails récap quotidiens
const SAV_MAIL_RELATION_CLIENT = "[relationclient@optimea.fr](mailto:relationclient@optimea.fr)";
const SAV_MAIL_ADV = "[adv@optimea.fr](mailto:adv@optimea.fr)";
const SAV_MAIL_FROM = "cedric.gesnouin@optimea.fr";

// =========================
// EMAIL — envoi via le compte Google lié au script (MailApp)




// Cache (fluidité : limite les lectures Sheets répétées)
const SAV_CACHE_TTL_SECONDS = 120; // court TTL + invalidation sur écriture = fluide sans incohérences longues
const SAV_CACHE_KEYS = {
  MODELS: "sav:models",
  PROCS: "sav:procs",
  STATS: "sav:stats",
  KPIS: "sav:kpis",
  CASES_ALL: "sav:cases:all",
  LOG_OPEN: "sav:log:open",
  MP_OPEN: "sav:mp:open",
  CENTRALES: "sav:centrales",
};

function savCacheGet_(key) {
  try {
    const v = CacheService.getScriptCache().get(key);
    return v ? JSON.parse(v) : null;
  } catch (e) {
    return null;
  }
}

function savCachePut_(key, obj, ttlSeconds) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(obj), ttlSeconds || SAV_CACHE_TTL_SECONDS);
  } catch (e) {
    // ignore cache failures
  }
}

function savCacheInvalidateAll_() {
  try {
    CacheService.getScriptCache().removeAll(Object.keys(SAV_CACHE_KEYS).map((k) => SAV_CACHE_KEYS[k]));
  } catch (e) {
    // ignore
  }
}

function savHtmlEsc_(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
}

function savSetupDailyRecapTriggers() {
  // Installe un déclencheur quotidien (une fois) pour envoyer les récap.
  const existing = ScriptApp.getProjectTriggers().some((t) => t.getHandlerFunction() === "savSendDailyRecaps");
  if (existing) return { ok: true, message: "Déclencheur déjà installé." };
  ScriptApp.newTrigger("savSendDailyRecaps").timeBased().everyDays(1).atHour(19).create(); // ~19h (timezone script)
  return { ok: true, message: "Déclencheur installé (quotidien ~19h)." };
}

function savRemoveDailyRecapTriggers() {
  const tr = ScriptApp.getProjectTriggers();
  let removed = 0;
  tr.forEach((t) => {
    if (t.getHandlerFunction() === "savSendDailyRecaps") {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  return { ok: true, removed };
}

function savTestEmail(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const to = String(o.to || "").trim();
  if (!to) return { ok: false, message: "Email destinataire requis." };
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const subj = "[TEST] SAV OPTIMEA — email de test (" + now + ")";
  const body =
    "Bonjour,\n\n" +
    "Ceci est un email de test envoyé par l'application SAV.\n\n" +
    "- Expéditeur : " +
    SAV_MAIL_FROM +
    "\n" +
    "- Envoyé le : " +
    now +
    "\n";
  try {
    const r = savSendEmailFromSav_(to, subj, body, []);
    return r && r.ok ? { ok: true, message: "OK — email test envoyé.", from: r.from } : { ok: false, message: (r && r.message) || "Erreur envoi." };
  } catch (e) {
    const msg = e && e.message ? String(e.message) : String(e);
    const stack = e && e.stack ? String(e.stack) : "";
    return { ok: false, message: msg, stack: stack };
  }
}

function savSendDailyRecaps() {
  // Point d’entrée appelé par le déclencheur.
  ensureSavSheets_();
  savCacheInvalidateAll_(); // force la fraîcheur des données pour le récap

  const now = new Date();
  const day = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");

  // 1) Marketplace : dossiers ouverts + tri par état
  const mp = savMpListCases(true) || [];
  const mpByEtat = {};
  mp.forEach((c) => {
    const e = String(c.etat || ETAT_CREATION).trim();
    if (!mpByEtat[e]) mpByEtat[e] = [];
    mpByEtat[e].push(c);
  });

  const mpNums = mp.map((c) => String(c.numero || "").trim()).filter(Boolean);
  const mpSubject = "[SAV] Marketplace — actions à faire (" + day + ") — " + mpNums.length + " dossier(s)";
  const mpBody = savBuildMarketplaceRecapHtml_(day, mpByEtat, mpNums);
  // Envoi depuis [sav@optimea.fr](mailto:sav@optimea.fr) via Microsoft 365 (HTML conservé)
  savSendEmailFromSavHtml_(SAV_MAIL_RELATION_CLIENT, mpSubject, mpBody, []);

  // Logistique : à traiter + expédié (en route)
  const log = savListLogistique(true) || [];
  const aTraiter = log.filter((x) => String(x.statutLog || "") === LOG_STAT_A_TRAITER);
  const enRoute = log.filter((x) => String(x.statutLog || "") === LOG_STAT_EXPEDIE);
  const advSubject =
    "[SAV] Logistique — actions à faire (" +
    day +
    ") — " +
    (aTraiter.length + enRoute.length) +
    " demande(s)";
  const advBody = savBuildLogistiqueRecapHtml_(day, aTraiter, enRoute);
  savSendEmailFromSavHtml_(SAV_MAIL_ADV, advSubject, advBody, []);

  return { ok: true, day, mpCount: mpNums.length, logATraiter: aTraiter.length, logEnRoute: enRoute.length };
}

function savBuildMarketplaceRecapHtml_(day, byEtat, allNums) {
  const esc = savHtmlEsc_;
  let h = "";
  h += '<div style="font-family:Arial,sans-serif;font-size:13px;color:#0f172a">';
  h += "<h2 style=\"margin:0 0 10px 0\">Marketplace — récap actions (" + esc(day) + ")</h2>";
  h += '<div style="margin:0 0 10px 0;color:#475569">Dossiers ouverts : <b>' + esc(allNums.length) + "</b></div>";
  if (!allNums.length) {
    h += '<div style="color:#16a34a"><b>Aucune action</b> : aucun dossier Marketplace ouvert.</div>';
    h += "</div>";
    return h;
  }
  h += '<div style="margin:0 0 12px 0"><b>Numéros à traiter</b> : ' + esc(allNums.join(", ")) + "</div>";
  const etats = Object.keys(byEtat || {}).sort();
  etats.forEach((e) => {
    const rows = byEtat[e] || [];
    if (!rows.length) return;
    h += '<div style="margin:14px 0 6px 0"><b>État : ' + esc(e) + "</b> (" + esc(rows.length) + ")</div>";
    h += '<ul style="margin:6px 0 0 18px;padding:0">';
    rows.slice(0, 80).forEach((c) => {
      h += "<li><b>" + esc(c.numero) + "</b> — " + esc(c.marketplace) + " — " + esc(c.clientEmail || "") + "</li>";
    });
    if (rows.length > 80) h += "<li>… (" + esc(rows.length - 80) + " autres)</li>";
    h += "</ul>";
  });
  h += "</div>";
  return h;
}

function savBuildLogistiqueRecapHtml_(day, aTraiter, enRoute) {
  const esc = savHtmlEsc_;
  let h = "";
  h += '<div style="font-family:Arial,sans-serif;font-size:13px;color:#0f172a">';
  h += "<h2 style=\"margin:0 0 10px 0\">Logistique — récap actions (" + esc(day) + ")</h2>";
  h +=
    '<div style="margin:0 0 12px 0;color:#475569">À traiter : <b>' +
    esc(aTraiter.length) +
    "</b> · En route : <b>" +
    esc(enRoute.length) +
    "</b></div>";

  function section(title, rows) {
    let s = '<div style="margin:14px 0 6px 0"><b>' + esc(title) + "</b> (" + esc(rows.length) + ")</div>";
    if (!rows.length) return s + '<div style="color:#475569">—</div>';
    s += '<ul style="margin:6px 0 0 18px;padding:0">';
    rows.slice(0, 120).forEach((x) => {
      const hasTracking = String((x && x.tracking) || "").trim();
      const isExpedied = String((x && x.statutLog) || "").trim() === LOG_STAT_EXPEDIE;
      const miss = hasTracking ? "" : (isExpedied ? " — ⚠ tracking manquant" : "");
      s +=
        "<li><b>" +
        esc(x.numero) +
        "</b> — " +
        esc(x.typeEnvoi || "") +
        " — " +
        esc(x.pointDepart || "") +
        (x.tracking ? " — tracking " + esc(x.tracking) : "") +
        miss +
        "</li>";
    });
    if (rows.length > 120) s += "<li>… (" + esc(rows.length - 120) + " autres)</li>";
    s += "</ul>";
    return s;
  }

  h += section("Demandes à expédier (À traiter)", aTraiter);
  h += section("Demandes en route (Expédié)", enRoute);
  h += "</div>";
  return h;
}

// Procédures SAV Optimea (modèle → procédure)
const SAV_PROC_POSTAL = "PROC_1_POSTAL";
const SAV_PROC_RAPATRIEMENT = "PROC_2_RAPATRIEMENT";
const SAV_PROC_VISIO = "PROC_3_VISIO";

function savProcedureDefs_() {
  const cached = savCacheGet_(SAV_CACHE_KEYS.PROCS);
  if (cached && cached.procDefs) return cached.procDefs;
  const base = {
    [SAV_PROC_POSTAL]: {
      id: SAV_PROC_POSTAL,
      label: "Procédure 1 — SAV par courrier postal (Appareils < 50€)",
      objectif:
        "Collecter les informations du produit par voie postale pour traiter les dossiers de garantie de manière simplifiée.",
      documentation: ["Plaque signalétique envoyée par courrier (étiquette du produit)", "Photos du défaut signalé"],
      etapes: [
        "Client signale un problème (contact SAV, création du dossier).",
        "Demande de documentation (plaque signalétique + photos du défaut).",
        "Réception par courrier postal de la plaque signalétique (archivage) — pas de réception en atelier.",
        "Analyse (n° série, date d’achat, couverture garantie) puis décision.",
        "Clôture du dossier (communication + archivage).",
      ],
    },
    [SAV_PROC_RAPATRIEMENT]: {
      id: SAV_PROC_RAPATRIEMENT,
      label: "Procédure 2 — SAV normal avec rapatriement (Appareils > 50€ — Mobile/Portable)",
      objectif: "Rapatrier l’appareil complet pour expertise technique avant décision de garantie.",
      documentation: ["Facture / date d’achat", "Infos produit (n° série)", "Description du problème", "Photos (si dispo)"],
      etapes: [
        "Client signale un problème (contact SAV, création du dossier).",
        "Collecte d’informations (conditions de garantie + infos produit).",
        "Demande d’envoi de l’appareil (étiquette transport, emballage, suivi).",
        "Réception (contrôle colis, enregistrement).",
        "Expertise technique (diagnostic, tests, cause) puis décision.",
        "Solution : si couvert → réparation/remplacement ; sinon → devis ou renvoi simple.",
        "Envoi au client (préparation + expédition).",
        "Confirmation de livraison (suivi).",
        "Clôture (enregistrement + archivage).",
      ],
    },
    [SAV_PROC_VISIO]: {
      id: SAV_PROC_VISIO,
      label: "Procédure 3 — SAV avec diagnostic en visio (Appareils > 50€ — Fixe)",
      objectif:
        "Effectuer un diagnostic en visioconférence avant rapatriement pour vérifier l’installation et réduire les rejets de garantie.",
      documentation: ["Facture / date d’achat", "Infos produit (n° série)", "Description du problème", "Visio (contrôle installation)"],
      etapes: [
        "Client signale un problème (contact SAV, création du dossier).",
        "Collecte d’informations initiales (conditions de garantie + infos produit).",
        "Diagnostic visio : vérification installation/environnement/dégâts externes + pré-diagnostic.",
        "Décision post-visio : si OK → demande d’envoi ; sinon → conseils / rejet selon cas.",
        "Demande d’envoi de l’appareil (étiquette transport, emballage, suivi).",
        "Réception (contrôle colis, enregistrement).",
        "Expertise technique approfondie puis décision.",
        "Solution : si couvert → réparation/remplacement ; sinon → devis ou renvoi simple.",
        "Envoi au client + confirmation de livraison.",
        "Clôture (enregistrement + archivage).",
      ],
    },
  };
  const custom = savReadProceduresFromSheet_();
  // Custom overrides base (id match) + adds new procedures.
  Object.keys(custom).forEach((k) => (base[k] = custom[k]));
  savCachePut_(SAV_CACHE_KEYS.PROCS, Object.assign({}, cached || {}, { procDefs: base }));
  return base;
}

function savProcedureModelToProcId_() {
  const cached = savCacheGet_(SAV_CACHE_KEYS.PROCS);
  if (cached && cached.procMap) return cached.procMap;
  const m = {};
  const addAll = (procId, arr) => arr.forEach((x) => (m[String(x || "").trim()] = procId));

  // PROCÉDURE 1 : postal (<50€)
  addAll(SAV_PROC_POSTAL, [
    "OVP-B40",
    "OVP-A40",
    "OVP-AR40N",
    "OVP-AR30B",
    "OVP2-B40",
    "OVP2-BR40N",
    "OVT-A15",
    "OVT-A30",
    "OVT2-A30",
    "OVC2-A19",
    "OCE-A01-2000B",
    "OCE-A05-2000 n/b",
  ]);

  // PROCÉDURE 2 : rapatriement (>50€ mobile/portable)
  addAll(SAV_PROC_RAPATRIEMENT, [
    "OVP-C40",
    "OVP-C40LUXE",
    "OPC-A01-050",
    "OPC-A01-070",
    "OPC-A01-070HP",
    "OPC-A01-090hHP",
    "OPC-A01-120HP",
    "OPC-A01-120HPWIFI",
    "OPC-A01-120",
    "OPC-A02-140",
    "OPC-A02-160HP",
    "OPC-A02-180",
    "OPC-A01-140",
    "OPC-A01-160HP",
    "OPC-A01-180",
    "OPC-B01-050",
    "OPC-B01-070",
    "OPC-B01-090",
    "OPC-B01-120",
    "OPC-B01-140",
    "OPC-C01-071",
    "OPC-C01-091",
    "OPC-C02-121",
    "OPC-C02-121HP",
    "OPC-D01-050",
    "MMCS-12HRN8-QRD0",
    "ORA-50M",
    "ORA-56D",
    "ORA-70D",
    "ORA-440D",
    "DN12",
    "DF20",
    "DM50O",
    "DH-J04-100",
    "OCE-D01-1500",
    "OCE-D01-2000",
    "OCE-D01-2500",
    "OCE-B01-1500",
    "OCE-C01-2000",
    "OCE-C03-2000",
    "OCE-C05-2200",
    "OCE-C08-1500",
    "OCE-G01-1500",
    "OCE-E01-1500",
    "OCE-F01-1500",
    "OCI-S03",
    "OCE-H02-2200",
  ]);

  // PROCÉDURE 3 : visio (>50€ fixe)
  addAll(SAV_PROC_VISIO, [
    "OAC-300-RE1",
    "OAC-250-RE2",
    "OAC-270-SD1",
    "OAC-270-SDIN",
    "OAC-270-SDIN2",
    "OCF-IR1-90PW int",
    "OCF-IR1-90PW ext",
    "OCR-IR1-120PW int",
    "OCF-IR1-120PW ext",
    "OCF-IR1-180PW int",
    "OCF-IR1-180PW ext",
    "OCF-IR2-90PW",
    "OCF-IR2-120PW",
    "OCF-IR2-180W",
    "MULTIWIND REV",
    "OPT-ORIBT-1000",
    "OPT-ORIBT-2000",
    "OPT-ORICIS-1000",
    "OPT-ORICIS-1500",
    "OPT-ORICIS-2000",
    "OPT-ORIPR2-1000",
    "OPT-ORIPR2-1500",
    "OPT-ORIPR2-2000",
    "OPT-ORISM-2000",
    "OPT-ORISS-2000",
    "OPT-ORSS1-500/750",
    "OPT-ORSS2-1500/1750",
    "OPT-ORSS3-500/750",
    "OPT-ORSS4-1500/1750",
    "GCAT200i",
    "GCAT250i",
    "WINDPAC-EM45C",
    "WINDPAC-EM55C",
    "WINDPAC-ET90C",
  ]);

  const custom = savReadProcedureModelMapFromSheet_();
  Object.keys(custom).forEach((model) => (m[model] = custom[model]));

  // Si la feuille SAV_Modeles est remplie (procedure_1/2/3), elle prime pour le mapping.
  const fromSavModels = savReadProcedureModelMapFromSavModelsSheet_();
  Object.keys(fromSavModels).forEach((model) => (m[model] = fromSavModels[model]));

  savCachePut_(SAV_CACHE_KEYS.PROCS, Object.assign({}, cached || {}, { procMap: m }));
  return m;
}

function savGetProcedureForModel_(modele) {
  const mod = String(modele || "").trim();
  if (!mod) return null;
  const map = savProcedureModelToProcId_();
  const defs = savProcedureDefs_();
  const pid = map[mod] || "";
  return pid && defs[pid] ? defs[pid] : null;
}

function savProceduresGetAll() {
  ensureSavSheets_();
  const cached = savCacheGet_(SAV_CACHE_KEYS.PROCS);
  if (cached && cached.procDefs && cached.procMap) return { procDefs: cached.procDefs, procMap: cached.procMap };
  const procDefs = savProcedureDefs_();
  const procMap = savProcedureModelToProcId_();
  savCachePut_(SAV_CACHE_KEYS.PROCS, { procDefs, procMap });
  return { procDefs, procMap };
}

function savUpsertProcedure(payload) {
  ensureSavSheets_();
  const id = String((payload && payload.id) || "").trim();
  const label = String((payload && payload.label) || "").trim();
  if (!id) return { ok: false, message: "Id procédure requis." };
  if (!label) return { ok: false, message: "Libellé requis." };
  const objectif = String((payload && payload.objectif) || "").trim();
  const docText = String((payload && payload.documentationText) || "").trim();
  const etapesText = String((payload && payload.etapesText) || "").trim();
  const actif = payload && payload.actif !== undefined ? !!payload.actif : true;

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_PROCEDURES);
  ensureHeaderGeneric_(sh, HDR_PROCEDURES);
  const lr = sh.getLastRow();
  const now = new Date();

  let rowToUpdate = -1;
  if (lr >= 2) {
    const ids = sh.getRange(2, 1, lr - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || "").trim() === id) {
        rowToUpdate = i + 2;
        break;
      }
    }
  }

  const line = [id, label, objectif, docText, etapesText, actif ? "OUI" : "NON", now];
  if (rowToUpdate > 0) {
    sh.getRange(rowToUpdate, 1, 1, HDR_PROCEDURES.length).setValues([line]);
  } else {
    sh.appendRow(line);
  }
  savCacheInvalidateAll_();
  return { ok: true };
}

function savAssignModelsToProcedure(payload) {
  ensureSavSheets_();
  const procId = String((payload && payload.procId) || "").trim();
  if (!procId) return { ok: false, message: "Procédure requise." };
  const modelsText = String((payload && payload.modelsText) || "").trim();
  const models = modelsText
    .split(/[\n,;]+/g)
    .map((s) => String(s || "").trim())
    .filter(Boolean);
  if (!models.length) return { ok: false, message: "Aucun modèle fourni." };

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_PROCEDURE_MODELS);
  ensureHeaderGeneric_(sh, HDR_PROCEDURE_MODELS);
  const now = new Date();

  // Index existing rows by model
  const lr = sh.getLastRow();
  const idx = {};
  if (lr >= 2) {
    const vals = sh.getRange(2, 1, lr - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) idx[String(vals[i][0] || "").trim()] = i + 2;
  }

  let updated = 0;
  let inserted = 0;
  models.forEach((model) => {
    const r = idx[model];
    const line = [model, procId, now];
    if (r) {
      sh.getRange(r, 1, 1, HDR_PROCEDURE_MODELS.length).setValues([line]);
      updated++;
    } else {
      sh.appendRow(line);
      inserted++;
    }
  });
  savCacheInvalidateAll_();
  return { ok: true, updated, inserted };
}

function savUnassignModelFromProcedure(modele) {
  ensureSavSheets_();
  const model = String(modele || "").trim();
  if (!model) return { ok: false, message: "Modèle requis." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_PROCEDURE_MODELS);
  if (!sh || sh.getLastRow() < 2) return { ok: true, removed: 0 };
  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0] || "").trim() === model) {
      sh.deleteRow(i + 2);
      savCacheInvalidateAll_();
      return { ok: true, removed: 1 };
    }
  }
  savCacheInvalidateAll_();
  return { ok: true, removed: 0 };
}

function savMpListCases(openOnly) {
  ensureSavSheets_();
  if (openOnly) {
    const cached = savCacheGet_(SAV_CACHE_KEYS.MP_OPEN);
    if (cached && Array.isArray(cached)) return cached;
    const out = readCasesFrom_(SAV_SHEET_MP, "Marketplace", false);
    savCachePut_(SAV_CACHE_KEYS.MP_OPEN, out);
    return out;
  }
  return readCasesFrom_(SAV_SHEET_MP, "Marketplace", null);
}

function savMpListMessages(numero) {
  ensureSavSheets_();
  const n = String(numero || "").trim();
  if (!n) return [];
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_MP_MESSAGES);
  if (!sh || sh.getLastRow() < 2) return [];
  ensureHeaderGeneric_(sh, HDR_MP_MESSAGES);
  const lr = sh.getLastRow();
  const rangeNums = sh.getRange(2, 2, lr - 1, 1); // col B = N° dossier

  // Optimisation : on ne lit que les lignes correspondant au dossier demandé.
  // Ça évite de charger toute la feuille de messages si elle devient volumineuse.
  const finder = rangeNums.createTextFinder("^" + savEscapeRegExp_(n) + "$").useRegularExpression(true);
  const hits = finder.findAll() || [];
  if (!hits.length) return [];

  const out = [];
  const nc = HDR_MP_MESSAGES.length;
  for (let i = 0; i < hits.length; i++) {
    const row = hits[i].getRow();
    const r = sh.getRange(row, 1, 1, nc).getValues()[0];
    out.push({
      id: String(r[0] || ""),
      numero: String(r[1] || ""),
      marketplace: String(r[2] || ""),
      clientEmail: String(r[3] || ""),
      sens: String(r[4] || ""),
      message: String(r[5] || ""),
      auteur: String(r[6] || ""),
      date: formatDate_(r[7]),
      lu: String(r[8] || ""),
      notes: String(r[9] || ""),
    });
  }
  out.sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")));
  return out;
}

function savEscapeRegExp_(s) {
  return String(s || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function savMpAddMessage(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const numero = String(o.numero || "").trim();
  const marketplace = String(o.marketplace || "").trim();
  const clientEmail = String(o.clientEmail || "").trim();
  const sens = String(o.sens || "").trim(); // "OUT" (à transmettre) ou "IN" (du client)
  const message = String(o.message || "").trim();
  const auteur = String(o.auteur || "").trim();
  const sheet = String(o.sheet || "").trim();
  const rowMain = Number(o.rowMain || 0);
  if (!numero) return { ok: false, message: "Numéro dossier requis." };
  if (!message) return { ok: false, message: "Message requis." };
  if (sens !== "OUT" && sens !== "IN") return { ok: false, message: "Sens invalide." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_MP_MESSAGES);
  ensureHeaderGeneric_(sh, HDR_MP_MESSAGES);
  const msgId = Utilities.getUuid();
  const msgDate = new Date();
  sh.appendRow([
    msgId,
    numero,
    marketplace,
    clientEmail,
    sens,
    message,
    auteur || Session.getActiveUser().getEmail() || "SAV",
    msgDate,
    "NON",
    "",
  ]);

  // Archivage dans Drive (si on a le contexte dossier)
  try {
    if ((sheet === SAV_SHEET_MP || sheet === SAV_SHEET_DIST) && rowMain >= 2) {
      const { folder } = savEnsureDriveFolderAndWriteUrl_(sheet, rowMain, numero);
      const msgFolder = savGetOrCreateSubFolder_(folder, "Messages Marketplace");
      const tz = Session.getScriptTimeZone();
      const ts = Utilities.formatDate(msgDate, tz, "yyyy-MM-dd_HH-mm");
      const safeSens = sens === "IN" ? "CLIENT" : "SAV";
      const fileName = ts + "_" + safeSens + ".txt";
      const content =
        "Dossier: " +
        numero +
        "\nDate: " +
        Utilities.formatDate(msgDate, tz, "yyyy-MM-dd HH:mm") +
        "\nSens: " +
        sens +
        "\nAuteur: " +
        (auteur || Session.getActiveUser().getEmail() || "SAV") +
        "\nMarketplace: " +
        (marketplace || "") +
        "\nClient: " +
        (clientEmail || "") +
        "\n\n" +
        message +
        "\n";
      msgFolder.createFile(fileName, content, MimeType.PLAIN_TEXT);
    }
  } catch (e) {
    // ne bloque jamais l'envoi message
  }

  savCacheInvalidateAll_();
  return { ok: true, id: msgId };
}

function savMpMarkAllRead(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const numero = String(o.numero || "").trim();
  if (!numero) return { ok: false, message: "Numéro dossier requis." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_MP_MESSAGES);
  if (!sh || sh.getLastRow() < 2) return { ok: true, updated: 0 };
  ensureHeaderGeneric_(sh, HDR_MP_MESSAGES);
  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, HDR_MP_MESSAGES.length).getValues();
  let updated = 0;
  for (let i = 0; i < vals.length; i++) {
    const row = i + 2;
    const r = vals[i] || [];
    if (String(r[1] || "").trim() !== numero) continue;
    const sens = String(r[4] || "").trim();
    const lu = String(r[8] || "").trim().toUpperCase();
    if (sens === "IN" && lu !== "OUI") {
      sh.getRange(row, 9).setValue("OUI");
      updated++;
    }
  }
  if (updated) savCacheInvalidateAll_();
  return { ok: true, updated };
}

function savParseListField_(text) {
  const t = String(text || "").trim();
  if (!t) return [];
  // Accept JSON array or newline list
  if (t[0] === "[" && t[t.length - 1] === "]") {
    try {
      const arr = JSON.parse(t);
      if (Array.isArray(arr)) return arr.map((x) => String(x || "").trim()).filter(Boolean);
    } catch (e) {
      // fallthrough
    }
  }
  return t
    .split(/\n+/g)
    .map((s) => String(s || "").trim())
    .filter(Boolean);
}

function savReadProceduresFromSheet_() {
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_PROCEDURES);
  if (!sh || sh.getLastRow() < 2) return {};
  ensureHeaderGeneric_(sh, HDR_PROCEDURES);
  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, HDR_PROCEDURES.length).getValues();
  const out = {};
  vals.forEach((r) => {
    const id = String(r[0] || "").trim();
    if (!id) return;
    const actif = String(r[5] || "").trim().toUpperCase();
    if (actif === "NON" || actif === "FALSE" || actif === "0") return;
    out[id] = {
      id,
      label: String(r[1] || "").trim(),
      objectif: String(r[2] || "").trim(),
      documentation: savParseListField_(r[3]),
      etapes: savParseListField_(r[4]),
    };
  });
  return out;
}

function savReadProcedureModelMapFromSheet_() {
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_PROCEDURE_MODELS);
  if (!sh || sh.getLastRow() < 2) return {};
  ensureHeaderGeneric_(sh, HDR_PROCEDURE_MODELS);
  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, 2).getValues();
  const out = {};
  vals.forEach((r) => {
    const model = String(r[0] || "").trim();
    const pid = String(r[1] || "").trim();
    if (model && pid) out[model] = pid;
  });
  return out;
}

function savReadProcedureModelMapFromSavModelsSheet_() {
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_MODELS);
  if (!sh || sh.getLastRow() < 2) return {};
  ensureHeaderGeneric_(sh, HDR_MODELS);
  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, 3).getValues();
  const out = {};
  vals.forEach((r) => {
    const m1 = String(r[0] || "").trim();
    const m2 = String(r[1] || "").trim();
    const m3 = String(r[2] || "").trim();
    if (m1) out[m1] = SAV_PROC_POSTAL; // procedure_1
    if (m2) out[m2] = SAV_PROC_RAPATRIEMENT; // procedure_2
    if (m3) out[m3] = SAV_PROC_VISIO; // procedure_3
  });
  return out;
}

// États (colonne État des feuilles principales)
const ETAT_CREATION = "CRÉATION";
const ETAT_EN_ATTENTE = "En attente réception";
const ETAT_RECU = "Réception atelier";
const ETAT_EXPERTISE = "Expertise";
const ETAT_ATTENTE_PIECES = "EN ATTENTE - Pièces";
const ETAT_TRANSPORT_DEMANDE = "Transport - demande";
const ETAT_TRANSPORT_ROUTE = "Transport - en route";
const ETAT_TRANSPORT_LIVRE = "Transport - livré";
const ETAT_CLOTURE_AVOIR = "Clôture avec Avoir";
const ETAT_FERMÉ_HG = "FERMÉ - Hors garantie";
const ETAT_ARCHIVE = "Archivé";

function doGet(e) {
  try {
    // NE PAS appeler ensureSavSheets_() ici — trop lourd, déclenche OAuth
    let models = [];
    let procDefs = {};
    let procMap = {};
    try { models = savGetModels_(); } catch(e) { models = []; }
    try { procDefs = savProcedureDefs_(); } catch(e) { procDefs = {}; }
    try { procMap = savProcedureModelToProcId_(); } catch(e) { procMap = {}; }
    const payload = {
      version: SAV_APP_VERSION,
      models,
      procDefs,
      procMap,
      // Stats chargées à la demande côté client (bouton "Stat global").
    };
    const json = JSON.stringify(payload);
    const b64 = Utilities.base64Encode(json, Utilities.Charset.UTF_8);

    // Mode debug: afficher le HTML généré sans l'exécuter
    // Usage: .../exec?debug=1
    try {
      const dbg = e && e.parameter && String(e.parameter.debug || "").trim() === "1";
      if (dbg) {
        const html = getSavHtml_(b64);
        const esc = savHtmlEsc_;
        return HtmlService.createHtmlOutput(
          '<!doctype html><meta charset="utf-8"><title>DEBUG HTML</title>' +
            '<pre style="white-space:pre-wrap;word-break:break-word;font-family:ui-monospace,Consolas,monospace;font-size:12px;padding:12px">' +
            esc(html) +
            "</pre>"
        ).setTitle("Gestion SAV — DEBUG");
      }
    } catch (err) {
      // ignore
    }

    return HtmlService.createHtmlOutput(getSavHtml_(b64))
      .setTitle("Gestion SAV")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    const msg = String(err && err.stack ? err.stack : err && err.message ? err.message : err);
    return HtmlService.createHtmlOutput(
      '<!doctype html><meta charset="utf-8"><title>Erreur</title>' +
        '<div style="font-family:Arial,sans-serif;padding:16px;white-space:pre-wrap;color:#7f1d1d">' +
        'Erreur au chargement (' +
        SAV_APP_VERSION +
        ')\\n\\n' +
        msg.replace(/[&<>]/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[m])) +
        "</div>"
    ).setTitle("Gestion SAV — Erreur");
  }
}

function savGetStats() {
  return savGetStats_();
}

function savGetStats_() {
  const cached = savCacheGet_(SAV_CACHE_KEYS.STATS);
  if (cached) return cached;

  const DAY_MS = 24 * 60 * 60 * 1000;
  function msToDaysInt_(ms) {
    return Math.round(ms / DAY_MS);
  }
  function avgInt_(sumInt, count) {
    if (!count || count < 2) return null;
    return Math.round(sumInt / count);
  }

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const cases = savListCases() || [];
  let distEnCours = 0;
  let mpEnCours = 0;
  let clos = 0; // clôturés (avoir / hors garantie / archivés)
  // Statut "solution" sur dossiers finalisés
  let dossiersRepares = 0;
  let dossiersEchanges = 0;
  let dossiersAvoir = 0;
  let dossiersHG = 0;
  let dossiersFinalises = 0;

  const lastDecisionByNumero = {};
  try {
    const shExp = ss.getSheetByName(SAV_SHEET_EXP);
    if (shExp && shExp.getLastRow() >= 2) {
      ensureHeaderGeneric_(shExp, HDR_EXP);
      const lr = shExp.getLastRow();
      const vals = shExp.getRange(2, 1, lr - 1, HDR_EXP.length).getValues();
      // HDR_EXP: [Numero, Date expertise, ..., Decision, ...] => decision index = 7 (0-based) in savReadExpertiseByNumero_
      for (let i = 0; i < vals.length; i++) {
        const r = vals[i] || [];
        const numero = String(r[0] || "").trim();
        if (!numero) continue;
        const d = r[1];
        const decision = String(r[7] || "").trim();
        if (!decision) continue;
        const ts = d instanceof Date && !isNaN(d.getTime()) ? d.getTime() : 0;
        const prev = lastDecisionByNumero[numero];
        if (!prev || ts >= prev.ts) lastDecisionByNumero[numero] = { ts, decision };
      }
    }
  } catch (e) {
    // ignore
  }

  // Délais moyens clôture (création -> date clôture)
  let sumCloseDistDays = 0;
  let cntCloseDist = 0;
  let sumCloseMpDays = 0;
  let cntCloseMp = 0;

  try {
    const reNum = /^SAV-\d{4}-\d+$/;
    [
      { sheet: SAV_SHEET_DIST, type: "Distributeur", hdr: HDR_DIST },
      { sheet: SAV_SHEET_MP, type: "Marketplace", hdr: HDR_MP },
    ].forEach((cfg) => {
      const sh = ss.getSheetByName(cfg.sheet);
      if (!sh || sh.getLastRow() < 2) return;
      ensureHeaderGeneric_(sh, cfg.hdr);
      const idx = savHeaderIndexByName_(sh);
      const lr = sh.getLastRow();
      const width = Math.max(sh.getLastColumn(), cfg.hdr.length);
      const vals = sh.getRange(2, 1, lr - 1, width).getValues();
      const colNum = 1;
      const colDateCreation = idx["Date création"] || (cfg.sheet === SAV_SHEET_DIST ? 9 : 10);
      const colDateCloture = idx["Date clôture"] || cfg.hdr.length;
      for (let i = 0; i < vals.length; i++) {
        const r = vals[i] || [];
        const num = String(r[colNum - 1] || "").trim();
        if (!num || !reNum.test(num)) continue;
        const dCreate = r[colDateCreation - 1];
        const dClose = r[colDateCloture - 1];
        if (!(dCreate instanceof Date) || isNaN(dCreate.getTime())) continue;
        if (!(dClose instanceof Date) || isNaN(dClose.getTime())) continue;
        const ms = dClose.getTime() - dCreate.getTime();
        if (!(ms >= 0)) continue;
        const days = msToDaysInt_(ms);
        if (cfg.type === "Distributeur") {
          sumCloseDistDays += days;
          cntCloseDist++;
        } else {
          sumCloseMpDays += days;
          cntCloseMp++;
        }
      }
    });
  } catch (e) {
    // ignore
  }

  for (let i = 0; i < cases.length; i++) {
    const c = cases[i] || {};
    const etat = String(c.etat || "").trim();
    const type = String(c.type || "").trim();
    const isClos = etat === ETAT_CLOTURE_AVOIR || etat === ETAT_FERMÉ_HG;
    const isArch = etat === ETAT_ARCHIVE;
    if (isClos || isArch) {
      clos++;
      dossiersFinalises++;
      const numero = String(c.numero || "").trim();
      const dec = numero && lastDecisionByNumero[numero] ? String(lastDecisionByNumero[numero].decision || "").trim() : "";
      if (etat === ETAT_CLOTURE_AVOIR || dec === "Avoir") dossiersAvoir++;
      else if (dec === "Réparation validée") dossiersRepares++;
      else if (dec === "Échange") dossiersEchanges++;
      else if (etat === ETAT_FERMÉ_HG || dec === "Hors garantie" || dec === "Refus garantie") dossiersHG++;
      continue;
    }
    if (type === "Distributeur") distEnCours++;
    else if (type === "Marketplace") mpEnCours++;
  }

  // Logistique (transport)
  let transportATraiter = 0;
  let transportExpedie = 0;
  let transportLivre = 0;
  let transportTrackingManquant = 0;
  // Délai moyen livraison (date demande -> date livraison)
  let sumLivDays = 0;
  let cntLiv = 0;
  try {
    const rows = savListLogistique(false) || []; // on a besoin des livrés aussi
    for (let i = 0; i < rows.length; i++) {
      const s = String(rows[i] && rows[i].statutLog ? rows[i].statutLog : "").trim();
      if (s === LOG_STAT_A_TRAITER) transportATraiter++;
      else if (s === LOG_STAT_EXPEDIE) transportExpedie++;
      else if (s === LOG_STAT_LIVRE) transportLivre++;
      if (rows[i] && rows[i].missingTracking) transportTrackingManquant++;


    }
  } catch (e) {
    // ignore : stats ne doivent jamais bloquer le chargement
  }

  try {
    const sh = ss.getSheetByName(SAV_SHEET_LOG);
    if (sh && sh.getLastRow() >= 2) {
      ensureHeaderGeneric_(sh, HDR_LOG);
      const lr = sh.getLastRow();
      const vals = sh.getRange(2, 1, lr - 1, HDR_LOG.length).getValues();
      for (let i = 0; i < vals.length; i++) {
        const r = vals[i] || [];
        const stat = String(r[8] || "").trim();
        if (stat !== LOG_STAT_LIVRE) continue;
        const dDem = r[9];
        const dLiv = r[16];
        if (!(dDem instanceof Date) || isNaN(dDem.getTime())) continue;
        if (!(dLiv instanceof Date) || isNaN(dLiv.getTime())) continue;
        const ms = dLiv.getTime() - dDem.getTime();
        if (!(ms >= 0)) continue;
        sumLivDays += msToDaysInt_(ms);
        cntLiv++;
      }
    }
  } catch (e) {
    // ignore
  }

  let mpMessagesATraiter = 0;
  try {
    const openMp = {};
    try {
      const shMp = ss.getSheetByName(SAV_SHEET_MP);
      if (shMp && shMp.getLastRow() >= 2) {
        ensureHeaderGeneric_(shMp, HDR_MP);
        const idx = savHeaderIndexByName_(shMp);
        const lrMp = shMp.getLastRow();
        const widthMp = Math.max(shMp.getLastColumn(), HDR_MP.length);
        const valsMp = shMp.getRange(2, 1, lrMp - 1, widthMp).getValues();
        const colNum = 1;
        const etatCol = idx["État du dossier"] || 9;
        for (let i = 0; i < valsMp.length; i++) {
          const r = valsMp[i] || [];
          const numero = String(r[colNum - 1] || "").trim();
          if (!numero) continue;
          const etat = String(r[etatCol - 1] || "").trim();
          const isClos = etat === ETAT_CLOTURE_AVOIR || etat === ETAT_FERMÉ_HG;
          const isArch = etat === ETAT_ARCHIVE;
          if (isClos || isArch) continue;
          openMp[numero] = true;
        }
      }
    } catch (e) {}

    const unreadMap = savMpUnreadCounts_(); // { [numero]: count }
    Object.keys(unreadMap || {}).forEach((numero) => {
      if (numero && !openMp[numero]) return;
      mpMessagesATraiter += Number(unreadMap[numero] || 0);
    });
  } catch (e) {
    // ignore
  }

  let pdcTotal = 0;
  let pdcEnCours = 0;
  let pdcLivre = 0;
  let sumPdcLivDays = 0;
  let cntPdcLiv = 0;
  try {
    const sh = ss.getSheetByName(SAV_SHEET_PDC);
    if (sh && sh.getLastRow() >= 2) {
      ensureHeaderGeneric_(sh, HDR_PDC);
      const lr = sh.getLastRow();
      const vals = sh.getRange(2, 1, lr - 1, HDR_PDC.length).getValues();
      for (let i = 0; i < vals.length; i++) {
        const r = vals[i] || [];
        const id = String(r[0] || "").trim();
        if (!id) continue;
        const stat = String(r[5] || "").trim();
        // Les demandes "Annulé" sont considérées supprimées côté stats
        if (stat === LOG_STAT_ANNULE) continue;
        pdcTotal++;
        if (stat === LOG_STAT_A_TRAITER || stat === LOG_STAT_EXPEDIE) pdcEnCours++;
        else if (stat === LOG_STAT_LIVRE) {
          pdcLivre++;
          const dCreate = r[1];
          const dLiv = r[7];
          if (dCreate instanceof Date && !isNaN(dCreate.getTime()) && dLiv instanceof Date && !isNaN(dLiv.getTime())) {
            const ms = dLiv.getTime() - dCreate.getTime();
            if (ms >= 0) {
              sumPdcLivDays += msToDaysInt_(ms);
              cntPdcLiv++;
            }
          }
        }
      }
    }
  } catch (e) {
    // ignore
  }

  const out = {
    distEnCours,
    mpEnCours,
    totalOuverts: distEnCours + mpEnCours,
    dossiersClos: clos,
    dossiersFinalises,
    dossiersRepares,
    dossiersEchanges,
    dossiersAvoir,
    dossiersHG,
    delaiMoyenClotureDistJ: avgInt_(sumCloseDistDays, cntCloseDist),
    delaiMoyenClotureMpJ: avgInt_(sumCloseMpDays, cntCloseMp),
    delaiMoyenLivraisonJ: avgInt_(sumLivDays, cntLiv),
    transportATraiter,
    transportExpedie,
    transportLivre,
    transportTrackingManquant,
    mpMessagesATraiter,
    pdcTotal,
    pdcEnCours,
    pdcLivre,
    pdcDelaiMoyenLivraisonJ: avgInt_(sumPdcLivDays, cntPdcLiv),
  };
  savCachePut_(SAV_CACHE_KEYS.STATS, out);
  return out;
}

function savGetKpis() {
  return savGetKpis_();
}

function savGetKpis_() {
  const cached = savCacheGet_(SAV_CACHE_KEYS.KPIS);
  if (cached) return cached;

  const cases = savListCases() || [];
  const out = {
    total: 0,
    enCours: 0,
    finalises: 0,
    distributeurTotal: 0,
    marketplaceTotal: 0,
    // Transport / logistique
    transportATraiter: 0,
    transportExpedie: 0,
    transportLivre: 0,
    transportTrackingManquant: 0,
  };

  out.total = cases.length;

  for (let i = 0; i < cases.length; i++) {
    const c = cases[i] || {};
    const etat = String(c.etat || "").trim();
    if (savIsFinalEtat_(etat)) out.finalises++;
    else out.enCours++;

    const type = String(c.type || "").trim();
    if (type === "Distributeur") out.distributeurTotal++;
    else if (type === "Marketplace") out.marketplaceTotal++;
  }

  try {
    const rows = savListLogistique(true) || [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i] || {};
      const s = String(r.statutLog || "").trim();
      if (s === LOG_STAT_A_TRAITER) out.transportATraiter++;
      else if (s === LOG_STAT_EXPEDIE) out.transportExpedie++;
      else if (s === LOG_STAT_LIVRE) out.transportLivre++;
      if (r.missingTracking) out.transportTrackingManquant++;
    }
  } catch (e) {
    // ignore : KPIs ne doivent jamais bloquer l'app
  }

  savCachePut_(SAV_CACHE_KEYS.KPIS, out);
  return out;
}

function savComputeGlobalAnalytics_() {
  ensureSavSheets_();

  function norm_(s) {
    return String(s || "")
      .toLowerCase()
      .replace(/\\s+/g, " ")
      .trim();
  }
  function inc_(m, k, by) {
    const key = String(k || "").trim();
    if (!key) return;
    m[key] = (m[key] || 0) + (by || 1);
  }
  function top_(m, n) {
    const arr = Object.keys(m || {}).map((k) => ({ key: k, count: Number(m[k] || 0) }));
    arr.sort((a, b) => b.count - a.count || String(a.key).localeCompare(String(b.key)));
    return arr.slice(0, n || 10);
  }
  function pct_(count, total) {
    total = Number(total || 0);
    if (!total) return 0;
    return Math.round((Number(count || 0) / total) * 1000) / 10; // 1 décimale
  }

  const cases = savListCases() || [];
  const totalCases = cases.length;
  const byType = { Distributeur: 0, Marketplace: 0 };
  const byEtat = {};
  const byModele = {};
  const byPanneLabel = {};
  const byPanneFamille = {};
  const byPanneAppareil = {};
  const byDecision = {};

  // Parse panne format: "[APP / FAMILLE] LIBELLE"
  function parsePanne_(s) {
    const raw = String(s || "").trim();
    const m = raw.match(/^\[([^\]]+)\]\s*(.*)$/);
    if (!m) return { app: "", fam: "", label: raw };
    const head = String(m[1] || "").trim();
    const label = String(m[2] || "").trim();
    const parts = head.split("/").map((x) => String(x || "").trim());
    return { app: parts[0] || "", fam: parts[1] || "", label: label || raw };
  }

  for (let i = 0; i < cases.length; i++) {
    const c = cases[i] || {};
    const type = String(c.type || "").trim();
    const etat = String(c.etat || "").trim();
    const modele = String(c.modele || "").trim();
    const panne = String(c.panne || "").trim();

    if (type === "Distributeur") byType.Distributeur++;
    else if (type === "Marketplace") byType.Marketplace++;

    inc_(byEtat, etat);
    inc_(byModele, modele);

    const p = parsePanne_(panne);
    if (p.app) inc_(byPanneAppareil, p.app);
    if (p.fam) inc_(byPanneFamille, p.fam);
    if (p.label) inc_(byPanneLabel, p.label);
  }

  // Décisions d'expertise (dernière décision par dossier)
  try {
    const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
    const shExp = ss.getSheetByName(SAV_SHEET_EXP);
    if (shExp && shExp.getLastRow() >= 2) {
      ensureHeaderGeneric_(shExp, HDR_EXP);
      const lr = shExp.getLastRow();
      const vals = shExp.getRange(2, 1, lr - 1, HDR_EXP.length).getValues();
      const last = {}; // numero -> {ts, decision}
      for (let i = 0; i < vals.length; i++) {
        const r = vals[i] || [];
        const numero = String(r[0] || "").trim();
        if (!numero) continue;
        const d = r[1];
        const decision = String(r[7] || "").trim();
        if (!decision) continue;
        const ts = d instanceof Date && !isNaN(d.getTime()) ? d.getTime() : 0;
        const prev = last[numero];
        if (!prev || ts >= prev.ts) last[numero] = { ts, decision };
      }
      Object.keys(last).forEach((k) => inc_(byDecision, last[k].decision));
    }
  } catch (e) {
    // ignore
  }

  // Références les plus signalées (PDC + expertise "pièces utilisées")
  const byRef = {};
  try {
    const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
    const shPdc = ss.getSheetByName(SAV_SHEET_PDC);
    if (shPdc && shPdc.getLastRow() >= 2) {
      ensureHeaderGeneric_(shPdc, HDR_PDC);
      const lr = shPdc.getLastRow();
      const vals = shPdc.getRange(2, 1, lr - 1, HDR_PDC.length).getValues();
      for (let i = 0; i < vals.length; i++) {
        const r = vals[i] || [];
        const refs = String(r[12] || "").trim();
        if (!refs) continue;
        refs
          .split(/[,\\n;]/g)
          .map((x) => String(x || "").trim())
          .filter(Boolean)
          .forEach((ref) => inc_(byRef, ref));
      }
    }
  } catch (e) {}
  try {
    const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
    const shExp = ss.getSheetByName(SAV_SHEET_EXP);
    if (shExp && shExp.getLastRow() >= 2) {
      ensureHeaderGeneric_(shExp, HDR_EXP);
      const lr = shExp.getLastRow();
      const vals = shExp.getRange(2, 1, lr - 1, HDR_EXP.length).getValues();
      for (let i = 0; i < vals.length; i++) {
        const r = vals[i] || [];
        const pieces = String(r[4] || "").trim(); // Pièces utilisées (free text)
        if (!pieces) continue;
        // Extract tokens that look like refs (alphanum >=4)
        const hits = pieces.match(/[A-Z0-9][A-Z0-9_-]{3,}/gi) || [];
        hits.forEach((h) => inc_(byRef, String(h || "").trim().toUpperCase()));
      }
    }
  } catch (e) {}

  const out = {
    totalCases,
    byType,
    topEtats: top_(byEtat, 10).map((x) => ({ ...x, pct: pct_(x.count, totalCases) })),
    topModeles: top_(byModele, 15).map((x) => ({ ...x, pct: pct_(x.count, totalCases) })),
    topPannes: top_(byPanneLabel, 20).map((x) => ({ ...x, pct: pct_(x.count, totalCases) })),
    topFamilles: top_(byPanneFamille, 15).map((x) => ({ ...x, pct: pct_(x.count, totalCases) })),
    topAppareils: top_(byPanneAppareil, 10).map((x) => ({ ...x, pct: pct_(x.count, totalCases) })),
    topDecisions: top_(byDecision, 10).map((x) => ({ ...x, pct: pct_(x.count, Math.max(1, Object.keys(byDecision).reduce((s,k)=>s+byDecision[k],0))) })),
    topRefs: top_(byRef, 20),
  };

  return out;
}

function savBuildGlobalAnalyticsSnapshot() {
  ensureSavSheets_();
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  let sh = ss.getSheetByName(SAV_SHEET_ANALYTICS);
  if (!sh) sh = ss.insertSheet(SAV_SHEET_ANALYTICS);
  ensureHeaderGeneric_(sh, HDR_ANALYTICS);

  const a = savComputeGlobalAnalytics_();
  const gen = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  const base = Number(a && a.totalCases ? a.totalCases : 0);

  const rows = [];
  function pushSection_(section, arr, includePct) {
    (arr || []).forEach((x) => {
      rows.push([gen, section, String(x.key || ""), Number(x.count || 0), includePct ? Number(x.pct || 0) : "", base]);
    });
  }
  pushSection_("Pannes (top)", a.topPannes, true);
  pushSection_("Familles pannes (top)", a.topFamilles, true);
  pushSection_("Types appareils (top)", a.topAppareils, true);
  pushSection_("Modèles (top)", a.topModeles, true);
  pushSection_("Décisions expertise (top)", a.topDecisions, true);
  pushSection_("États dossier (top)", a.topEtats, true);
  pushSection_("Références (top)", a.topRefs, false);

  // réécrit le contenu (hors entête)
  try {
    // Nettoyage zones + graphiques (sans toucher l'entête)
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(HDR_ANALYTICS.length, 20)).clear({ contentsOnly: true });
    (sh.getCharts() || []).forEach((c) => sh.removeChart(c));
  } catch (e) {}

  if (rows.length) sh.getRange(2, 1, rows.length, HDR_ANALYTICS.length).setValues(rows);

  // Mise en forme simple
  try {
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, HDR_ANALYTICS.length);
    sh.getRange(1, 1, 1, HDR_ANALYTICS.length).setFontWeight("bold");
  } catch (e) {}

  // Blocs synthèse + graphiques
  try {
    function writeBlock_(title, arr, startRow) {
      const startCol = 8; // colonne H
      sh.getRange(startRow, startCol, 1, 2).setValues([[title, "Nombre"]]).setFontWeight("bold");
      const data = (arr || []).map((x) => [String(x.key || ""), Number(x.count || 0)]);
      if (data.length) sh.getRange(startRow + 1, startCol, data.length, 2).setValues(data);
      return { row: startRow, col: startCol, height: Math.max(2, data.length + 1) };
    }

    const blocks = [];
    let r0 = 2;
    blocks.push({ name: "Pannes (top)", kind: "bar", ...writeBlock_("Top pannes", a.topPannes || [], r0) });
    r0 += blocks[blocks.length - 1].height + 2;
    blocks.push({ name: "Familles pannes (top)", kind: "pie", ...writeBlock_("Familles pannes", a.topFamilles || [], r0) });
    r0 += blocks[blocks.length - 1].height + 2;
    blocks.push({ name: "Modèles (top)", kind: "bar", ...writeBlock_("Modèles", a.topModeles || [], r0) });
    r0 += blocks[blocks.length - 1].height + 2;
    blocks.push({ name: "Références (top)", kind: "bar", ...writeBlock_("Références", a.topRefs || [], r0) });

    // Graphiques à droite des blocs (colonne K)
    blocks.forEach((b, i) => {
      const dataRange = sh.getRange(b.row, b.col, b.height, 2); // inclut titre
      const chartPosCol = 11; // K
      const chartPosRow = b.row;
      let builder = sh.newChart().addRange(dataRange).setOption("legend", { position: "right" }).setPosition(chartPosRow, chartPosCol, 0, 0);
      if (b.kind === "pie") builder = builder.asPieChart().setOption("pieHole", 0.35);
      else builder = builder.asBarChart().setOption("bars", "horizontal");
      sh.insertChart(builder.build());
    });
  } catch (e) {
    // ne doit jamais casser l'analyse
  }

  // cache court pour l'UI
  savCachePut_("sav:analytics:v1", a, 300);
  return a;
}

function savGetGlobalAnalytics(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const force = !!o.force;
  if (force) return savBuildGlobalAnalyticsSnapshot();
  const cached = savCacheGet_("sav:analytics:v1");
  if (cached) return cached;
  // si pas de cache, renvoie vide : l'analyse doit être lancée explicitement
  return { totalCases: 0, topPannes: [], topFamilles: [], topAppareils: [], topModeles: [], topDecisions: [], topRefs: [], topEtats: [], byType: { Distributeur: 0, Marketplace: 0 } };
}

function savEnsureAnnualCloseTicket_() {
  ensureSavSheets_();
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_ANALYTICS_YEAR_TICKETS);
  ensureHeaderGeneric_(sh, HDR_ANALYTICS_YEAR_TICKETS);
  const yearNow = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy"));
  const targetYear = yearNow - 1;
  if (targetYear < 2000) return { ok: false, message: "Année cible invalide." };
  const lr = sh.getLastRow();
  if (lr >= 2) {
    const vals = sh.getRange(2, 1, lr - 1, HDR_ANALYTICS_YEAR_TICKETS.length).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      const r = vals[i] || [];
      const y = Number(r[0] || 0);
      if (y === targetYear) {
        return { ok: true, year: targetYear, createdAt: formatDate_(r[1]), statut: String(r[2] || ""), notes: String(r[3] || "") };
      }
    }
  }
  const createdAt = new Date();
  sh.appendRow([targetYear, createdAt, "À faire", "Renseigner volumes par référence puis lancer l'analyse annuelle."]);
  savCacheInvalidateAll_();
  return { ok: true, year: targetYear, createdAt: formatDate_(createdAt), statut: "À faire", notes: "Renseigner volumes par référence puis lancer l'analyse annuelle." };
}

function savGetAnnualCloseTicket() {
  return savEnsureAnnualCloseTicket_();
}

function savListVolumesForYear(year) {
  ensureSavSheets_();
  const y = Number(year || 0);
  if (!y) return [];
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_ANALYTICS_VOLUMES);
  ensureHeaderGeneric_(sh, HDR_ANALYTICS_VOLUMES);
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  const vals = sh.getRange(2, 1, lr - 1, HDR_ANALYTICS_VOLUMES.length).getValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i] || [];
    if (Number(r[0] || 0) !== y) continue;
    const ref = String(r[1] || "").trim();
    if (!ref) continue;
    out.push({ row: i + 2, year: y, ref, volume: Number(r[2] || 0) });
  }
  out.sort((a, b) => String(a.ref).localeCompare(String(b.ref)));
  return out;
}

function savUpsertVolumeForYear(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const y = Number(o.year || 0);
  const ref = String(o.ref || "").trim();
  const volume = Number(o.volume || 0);
  if (!y || !ref) return { ok: false, message: "Année + référence requises." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_ANALYTICS_VOLUMES);
  ensureHeaderGeneric_(sh, HDR_ANALYTICS_VOLUMES);
  const lr = sh.getLastRow();
  if (lr >= 2) {
    const vals = sh.getRange(2, 1, lr - 1, HDR_ANALYTICS_VOLUMES.length).getValues();
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i] || [];
      if (Number(r[0] || 0) === y && String(r[1] || "").trim() === ref) {
        sh.getRange(i + 2, 3).setValue(volume);
        savCacheInvalidateAll_();
        return { ok: true, updated: true };
      }
    }
  }
  sh.appendRow([y, ref, volume]);
  savCacheInvalidateAll_();
  return { ok: true, inserted: true };
}

function savBuildAnnualAnalyticsSnapshot(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const y = Number(o.year || 0);
  if (!y) return { ok: false, message: "Année requise." };

  // Compte refs signalées sur l'année (PDC + Expertise pièces)
  const byRef = {};
  function inc_(m, k) {
    const key = String(k || "").trim();
    if (!key) return;
    m[key] = (m[key] || 0) + 1;
  }

  try {
    const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
    const shPdc = ss.getSheetByName(SAV_SHEET_PDC);
    if (shPdc && shPdc.getLastRow() >= 2) {
      ensureHeaderGeneric_(shPdc, HDR_PDC);
      const lr = shPdc.getLastRow();
      const vals = shPdc.getRange(2, 1, lr - 1, HDR_PDC.length).getValues();
      for (let i = 0; i < vals.length; i++) {
        const r = vals[i] || [];
        const dCreate = r[1];
        if (!(dCreate instanceof Date) || isNaN(dCreate.getTime())) continue;
        if (dCreate.getFullYear() !== y) continue;
        const refs = String(r[12] || "").trim();
        if (!refs) continue;
        refs
          .split(/[,\\n;]/g)
          .map((x) => String(x || "").trim())
          .filter(Boolean)
          .forEach((ref) => inc_(byRef, ref));
      }
    }
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
    const shExp = ss.getSheetByName(SAV_SHEET_EXP);
    if (shExp && shExp.getLastRow() >= 2) {
      ensureHeaderGeneric_(shExp, HDR_EXP);
      const lr = shExp.getLastRow();
      const vals = shExp.getRange(2, 1, lr - 1, HDR_EXP.length).getValues();
      for (let i = 0; i < vals.length; i++) {
        const r = vals[i] || [];
        const d = r[1];
        if (!(d instanceof Date) || isNaN(d.getTime())) continue;
        if (d.getFullYear() !== y) continue;
        const pieces = String(r[4] || "").trim();
        if (!pieces) continue;
        const hits = pieces.match(/[A-Z0-9][A-Z0-9_-]{3,}/gi) || [];
        hits.forEach((h) => inc_(byRef, String(h || "").trim().toUpperCase()));
      }
    }
  } catch (e) {}

  // Volumes saisis
  const vols = savListVolumesForYear(y) || [];
  const volMap = {};
  vols.forEach((v) => (volMap[String(v.ref || "").trim()] = Number(v.volume || 0)));

  // Ecrit une section annuelle dans SAV_Analytics (sans casser le reste)
  try {
    const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
    let sh = ss.getSheetByName(SAV_SHEET_ANALYTICS);
    if (!sh) sh = ss.insertSheet(SAV_SHEET_ANALYTICS);
    ensureHeaderGeneric_(sh, HDR_ANALYTICS);

    const gen = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    const keys = Object.keys(byRef).sort((a, b) => (byRef[b] || 0) - (byRef[a] || 0) || a.localeCompare(b));
    const rows = [];
    for (let i = 0; i < Math.min(50, keys.length); i++) {
      const ref = keys[i];
      const count = Number(byRef[ref] || 0);
      const vol = Number(volMap[ref] || 0);
      const rate = vol > 0 ? Math.round((count / vol) * 10000) / 100 : ""; // % avec 2 décimales
      rows.push([gen, "ANNUEL " + String(y), ref, count, rate, vol]);
    }
    // append sous les données existantes
    const start = sh.getLastRow() + 2;
    if (rows.length) sh.getRange(start, 1, rows.length, HDR_ANALYTICS.length).setValues(rows);
  } catch (e) {}

  savCacheInvalidateAll_();
  return { ok: true, year: y, refs: Object.keys(byRef).length, volumes: vols.length };
}

function savImportVolumesFromCsv(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const y = Number(o.year || 0);
  const file = o && o.file ? o.file : null; // {name,mimeType,base64}
  if (!y) return { ok: false, message: "Année requise." };
  if (!file || !file.base64) return { ok: false, message: "Fichier requis." };

  const name = String(file.name || "").trim();
  const mimeType = String(file.mimeType || "").trim().toLowerCase();
  // On supporte CSV / texte uniquement (Excel .xlsx non parsé ici)
  const looksLikeCsv = mimeType.indexOf("csv") !== -1 || mimeType.indexOf("text") !== -1 || /\.csv$/i.test(name);
  if (!looksLikeCsv) {
    return { ok: false, message: "Format non supporté pour import direct. Exporte en CSV (colonnes: Année, Référence, Volume produits) puis réessaie." };
  }

  let text = "";
  try {
    const bytes = Utilities.base64Decode(String(file.base64 || ""));
    text = Utilities.newBlob(bytes).getDataAsString("UTF-8");
  } catch (e) {
    return { ok: false, message: "Impossible de lire le CSV." };
  }

  let rows = [];
  try {
    rows = Utilities.parseCsv(text);
  } catch (e) {
    return { ok: false, message: "CSV invalide (séparateur/encodage). Essaie un CSV UTF-8 séparé par virgules." };
  }
  if (!rows || rows.length < 2) return { ok: false, message: "CSV vide." };

  // Détecte l'entête
  const header = rows[0].map((x) => String(x || "").trim().toLowerCase());
  function col_(names) {
    for (let i = 0; i < header.length; i++) {
      const h = header[i];
      for (let j = 0; j < names.length; j++) if (h === names[j]) return i;
    }
    return -1;
  }
  const colYear = col_(["année", "annee", "year"]);
  const colRef = col_(["référence", "reference", "ref"]);
  const colVol = col_(["volume produits", "volume", "qty", "quantité", "quantite"]);
  if (colYear === -1 || colRef === -1 || colVol === -1) {
    return { ok: false, message: "Entêtes CSV attendues: Année, Référence, Volume produits." };
  }

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_ANALYTICS_VOLUMES);
  ensureHeaderGeneric_(sh, HDR_ANALYTICS_VOLUMES);

  let inserted = 0;
  let updated = 0;
  let ignored = 0;

  // Index existant (année+ref -> row)
  const lr = sh.getLastRow();
  const idx = {};
  if (lr >= 2) {
    const vals = sh.getRange(2, 1, lr - 1, HDR_ANALYTICS_VOLUMES.length).getValues();
    for (let i = 0; i < vals.length; i++) {
      const ry = Number(vals[i][0] || 0);
      const rr = String(vals[i][1] || "").trim();
      if (!ry || !rr) continue;
      idx[String(ry) + "||" + rr] = i + 2;
    }
  }

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const ry = Number(String(r[colYear] || "").trim());
    const ref = String(r[colRef] || "").trim();
    const vol = Number(String(r[colVol] || "").replace(",", ".").trim());
    if (!ry || !ref || !(vol >= 0)) {
      ignored++;
      continue;
    }
    if (ry !== y) {
      // ignore les autres années (import annuel ciblé)
      ignored++;
      continue;
    }
    const key = String(ry) + "||" + ref;
    const row = idx[key];
    if (row) {
      sh.getRange(row, 3).setValue(vol);
      updated++;
    } else {
      sh.appendRow([ry, ref, vol]);
      inserted++;
    }
  }

  savCacheInvalidateAll_();
  return { ok: true, inserted, updated, ignored };
}

function savCreateAvoirTicket(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const sheetName = String(o.sheet || "").trim();
  const row = Number(o.row || 0);
  const numero = String(o.numero || "").trim();
  const notes = String(o.notes || "").trim();
  const centraleAchat = String(o.centraleAchat || "").trim();
  const factureAchatCentrale = String(o.factureAchatCentrale || "").trim();
  if (!sheetName || !row || row < 2 || !numero) return { ok: false, message: "Paramètres invalides (dossier)." };
  if (!centraleAchat) return { ok: false, message: "Centrale d'achat requise pour créer une demande d’avoir." };

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const shMain = ss.getSheetByName(sheetName);
  if (!shMain) return { ok: false, message: "Feuille dossier introuvable." };

  const idx = savHeaderIndexByName_(shMain);
  const width = shMain.getLastColumn();
  const r = shMain.getRange(row, 1, 1, width).getValues()[0] || [];

  let magasin = "";
  let adresse = "";
  let modele = "";
  let panne = "";
  let serie = "";
  let type = "";

  if (sheetName === SAV_SHEET_DIST) {
    magasin = String(r[(idx["Nom magasin/distributeur"] || 2) - 1] || "").trim();
    adresse = String(r[(idx["Adresse complète"] || 3) - 1] || "").trim();
    modele = String(r[(idx["Modèle concerné"] || 4) - 1] || "").trim();
    panne = String(r[(idx["Panne constatée"] || 5) - 1] || "").trim();
    serie = String(r[(idx["Numéro de série"] || 6) - 1] || "").trim();
    type = "Distributeur";
  } else if (sheetName === SAV_SHEET_MP) {
    magasin = "Marketplace";
    adresse = "";
    modele = "";
    panne = String(r[(idx["Panne constatée"] || 6) - 1] || "").trim();
    serie = "";
    type = "Marketplace";
  } else {
    type = sheetName;
  }

  const driveCol = idx["Dossier Drive"] || ((sheetName === SAV_SHEET_DIST ? HDR_DIST.length : HDR_MP.length) - 1);
  const driveUrl = String(r[driveCol - 1] || "").trim();
  const now = new Date();

  const shT = ss.getSheetByName(SAV_SHEET_AVOIR_TICKETS);
  ensureHeaderGeneric_(shT, HDR_AVOIR_TICKETS);

  shT.appendRow([
    numero,
    magasin,
    adresse,
    modele,
    panne,
    serie,
    type,
    now,
    "À faire",
    driveUrl,
    centraleAchat,
    factureAchatCentrale,
    notes,
    sheetName,
    row,
  ]);

  // Met le dossier en "Clôture avec Avoir" (comme avant), mais le travail est suivi dans Tickets Avoir.
  try {
    savUpdateEtat_({ sheet: sheetName, row: row, etat: ETAT_CLOTURE_AVOIR });
  } catch (e) {
    // ne bloque pas la création du ticket
  }

  savCacheInvalidateAll_();
  return { ok: true };
}

function savListAvoirTicketsPendingByDistributor() {
  ensureSavSheets_();
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_AVOIR_TICKETS);
  if (!sh || sh.getLastRow() <= 1) return [];
  ensureHeaderGeneric_(sh, HDR_AVOIR_TICKETS);

  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, HDR_AVOIR_TICKETS.length).getValues();
  const groups = {};
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    const statut = String(r[8] || "").trim();
    if (statut && statut !== "À faire") continue;
    const magasin = String(r[1] || "").trim() || "—";
    const centrale = String(r[10] || "").trim() || "—";
    if (!groups[centrale]) groups[centrale] = [];
    groups[centrale].push({
      ticketRow: i + 2,
      numero: String(r[0] || "").trim(),
      magasin: magasin,
      adresse: String(r[2] || ""),
      modele: String(r[3] || ""),
      panne: String(r[4] || ""),
      serie: String(r[5] || ""),
      type: String(r[6] || ""),
      dateCreation: formatDate_(r[7]),
      statut: statut || "À faire",
      driveUrl: String(r[9] || ""),
      centraleAchat: String(r[10] || ""),
      factureAchatCentrale: String(r[11] || ""),
      notes: String(r[12] || ""),
      sheet: String(r[13] || ""),
      row: Number(r[14] || 0),
    });
  }

  return Object.keys(groups)
    .sort((a, b) => a.localeCompare(b))
    .map((centraleAchat) => {
      const items = groups[centraleAchat] || [];
      items.sort((a, b) => String(b.dateCreation || "").localeCompare(String(a.dateCreation || "")));
      return { centraleAchat, items };
    });
}

function savListCentralesAchat() {
  ensureSavSheets_();
  const cached = savCacheGet_(SAV_CACHE_KEYS.CENTRALES);
  if (cached && Array.isArray(cached)) return cached;
  try {
    const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
    const sh = ss.getSheetByName(SAV_SHEET_AVOIR_TICKETS);
    if (!sh || sh.getLastRow() < 2) return [];
    ensureHeaderGeneric_(sh, HDR_AVOIR_TICKETS);
    const lr = sh.getLastRow();
    const vals = sh.getRange(2, 11, lr - 1, 1).getValues(); // col 11 = Centrale d'achat
    const seen = {};
    for (let i = 0; i < vals.length; i++) {
      const v = String(vals[i][0] || "").trim();
      if (!v) continue;
      seen[v] = true;
    }
    const out = Object.keys(seen).sort((a, b) => a.localeCompare(b));
    savCachePut_(SAV_CACHE_KEYS.CENTRALES, out);
    return out;
  } catch (e) {
    return [];
  }
}

function savSetAvoirTicketStatut(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const ticketRow = Number(o.ticketRow || 0);
  const statut = String(o.statut || "").trim();
  if (!ticketRow || ticketRow < 2 || !statut) return { ok: false, message: "Paramètres invalides." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_AVOIR_TICKETS);
  if (!sh || sh.getLastRow() < ticketRow) return { ok: false, message: "Ticket introuvable." };
  ensureHeaderGeneric_(sh, HDR_AVOIR_TICKETS);
  // colonne 9 = Statut
  sh.getRange(ticketRow, 9).setValue(statut);

  // Si l'avoir est fait : on archive automatiquement le dossier, sans nécessiter une action manuelle.
  // Colonnes (HDR_AVOIR_TICKETS) :
  // 13 = Feuille dossier, 14 = Ligne dossier (1-based)
  try {
    if (statut === "Terminé") {
      const r = sh.getRange(ticketRow, 1, 1, HDR_AVOIR_TICKETS.length).getValues()[0] || [];
      const sheetName = String(r[13] || "").trim();
      const row = Number(r[14] || 0);
      if (sheetName && row && row >= 2) {
        // On archive le dossier (sauvegarde aussi Date clôture si vide via savUpdateEtat)
        savUpdateEtat_({ sheet: sheetName, row: row, etat: ETAT_ARCHIVE });
      }
    }
  } catch (e) {
    // ne bloque jamais la validation du ticket
  }

  savCacheInvalidateAll_();
  return { ok: true };
}

function savDeleteAvoirTicket(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const ticketRow = Number(o.ticketRow || 0);
  if (!ticketRow || ticketRow < 2) return { ok: false, message: "Ligne ticket invalide." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_AVOIR_TICKETS);
  if (!sh) return { ok: false, message: "Feuille Tickets Avoir introuvable." };
  if (sh.getLastRow() < ticketRow) return { ok: false, message: "Ticket introuvable." };
  sh.deleteRow(ticketRow);
  savCacheInvalidateAll_();
  return { ok: true };
}

function savGetModels() {
  ensureSavSheets_();
  return savGetModels_();
}

function savGetProcedureForModel(modele) {
  ensureSavSheets_();
  return savGetProcedureForModel_(modele);
}

function savGetModels_() {
  const prodId = String(SAV_PROD_SPREADSHEET_ID || SAV_SPREADSHEET_ID || "").trim();
  const cacheKey = SAV_CACHE_KEYS.MODELS + ":" + prodId;
  const cached = savCacheGet_(cacheKey);
  if (cached && Array.isArray(cached)) return cached;
  const ssSav = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  let ssProd = ssSav;
  if (prodId && prodId !== String(SAV_SPREADSHEET_ID || "").trim()) {
    try {
      ssProd = SpreadsheetApp.openById(prodId);
    } catch (e) {
      // fallback: le classeur "production" peut être inaccessible selon les droits du compte exécutant la web app
      ssProd = ssSav;
    }
  }
  const seen = {};
  const add = (name) => {
    const n = String(name || "").trim();
    if (n) seen[n] = true;
  };

  /**
   * Lit tous les noms de modèles (une ou plusieurs colonnes) depuis le bloc de données réellement utilisé.
   * Ignore la ligne d’en-tête si la 1ère ligne contient des intitulés (Nom/procedure_1/2/3).
   */
  function addModelsFromSheet_(sh, colCount) {
    if (!sh) return;
    const lr = sh.getLastRow();
    if (!lr) return;
    const cc = Math.max(1, Number(colCount || 1));
    // Plus fiable que getDataRange() si la feuille contient des zones “vides” / mise en forme.
    const data = sh.getRange(1, 1, lr, cc).getValues();
    if (!data || !data.length) return;
    let start = 0;
    const h0 = String((data[0] && data[0][0]) || "").trim().toLowerCase();
    const h1 = String((data[0] && data[0][1]) || "").trim().toLowerCase();
    const h2 = String((data[0] && data[0][2]) || "").trim().toLowerCase();
    const looksLikeHeader =
      h0 === "nom" ||
      h0 === "procedure_1" ||
      h0 === "procédure_1" ||
      h1 === "procedure_2" ||
      h1 === "procédure_2" ||
      h2 === "procedure_3" ||
      h2 === "procédure_3";
    if (looksLikeHeader) start = 1;
    for (let i = start; i < data.length; i++) {
      const row = data[i] || [];
      for (let c = 0; c < cc; c++) add(row[c]);
    }
  }

  // 1) Référentiel production : feuille Modeles/Modèles (classeur prod)
  addModelsFromSheet_(ssProd.getSheetByName(SAV_SHEET_MODELES_PROD), 1);
  addModelsFromSheet_(ssProd.getSheetByName(SAV_SHEET_MODELES_PROD_ALT), 1);

  // 2) Toute feuille dont le nom commence par « Modeles_ » (ex. export par usine), colonne A = nom modèle
  const all = ssProd.getSheets();
  for (let s = 0; s < all.length; s++) {
    const nm = all[s].getName();
    if (/^Modeles_/i.test(nm)) addModelsFromSheet_(all[s], 1);
  }

  // 3) Feuille dédiée SAV (modèles classés par procédure 1/2/3)
  const shSav = ssSav.getSheetByName(SAV_SHEET_MODELS);
  if (shSav) {
    ensureHeaderGeneric_(shSav, HDR_MODELS);
    addModelsFromSheet_(shSav, 3);
  }

  const out = Object.keys(seen).sort();
  savCachePut_(cacheKey, out);
  return out;
}

function savDebugModelsSources() {
  ensureSavSheets_();
  const ssSav = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const prodId = String(SAV_PROD_SPREADSHEET_ID || SAV_SPREADSHEET_ID || "").trim();
  const ssProd = prodId ? SpreadsheetApp.openById(prodId) : ssSav;
  const report = [];

  // Prod
  function countFromProd_(name) {
    const sh = ssProd.getSheetByName(name);
    if (!sh) return { sheet: "[PROD] " + name, exists: false, lastRow: 0, nonEmpty: 0 };
    const lr = sh.getLastRow();
    if (!lr) return { sheet: "[PROD] " + name, exists: true, lastRow: 0, nonEmpty: 0 };
    const colA = sh.getRange(1, 1, lr, 1).getValues();
    let start = 0;
    const h0 = String(colA[0][0] || "").trim().toLowerCase();
    if (h0 === "nom") start = 1;
    let nonEmpty = 0;
    for (let i = start; i < colA.length; i++) if (String(colA[i][0] || "").trim()) nonEmpty++;
    return { sheet: "[PROD] " + name, exists: true, lastRow: lr, nonEmpty };
  }
  function countFromSav_(name) {
    const sh = ssSav.getSheetByName(name);
    if (!sh) return { sheet: "[SAV] " + name, exists: false, lastRow: 0, nonEmpty: 0 };
    const lr = sh.getLastRow();
    if (!lr) return { sheet: "[SAV] " + name, exists: true, lastRow: 0, nonEmpty: 0 };
    const colA = sh.getRange(1, 1, lr, 1).getValues();
    let start = 0;
    const h0 = String(colA[0][0] || "").trim().toLowerCase();
    if (h0 === "nom") start = 1;
    let nonEmpty = 0;
    for (let i = start; i < colA.length; i++) if (String(colA[i][0] || "").trim()) nonEmpty++;
    return { sheet: "[SAV] " + name, exists: true, lastRow: lr, nonEmpty };
  }

  report.push(countFromProd_(SAV_SHEET_MODELES_PROD));
  report.push(countFromProd_(SAV_SHEET_MODELES_PROD_ALT));
  report.push(countFromSav_(SAV_SHEET_MODELS));

  const all = ssProd.getSheets();
  let modelesSheets = 0;
  let modelesRows = 0;
  for (let i = 0; i < all.length; i++) {
    const nm = all[i].getName();
    if (/^Modeles_/i.test(nm)) {
      modelesSheets++;
      const r = countFromProd_(nm);
      modelesRows += r.nonEmpty;
    }
  }

  return {
    totalDistinct: savGetModels_().length,
    sources: report,
    modelesSheets,
    modelesRows,
  };
}

function savClearCache() {
  // Fonction “standalone” : ne dépend pas de savCacheInvalidateAll_ (pratique si copie partielle du fichier).
  try {
    const prodId = String(SAV_PROD_SPREADSHEET_ID || SAV_SPREADSHEET_ID || "").trim();
    const keys = []
      .concat(Object.keys(SAV_CACHE_KEYS || {}).map((k) => SAV_CACHE_KEYS[k]))
      .concat([String((SAV_CACHE_KEYS && SAV_CACHE_KEYS.MODELS) || "sav:models") + ":" + prodId]);
    CacheService.getScriptCache().removeAll(keys);
    return { ok: true, removedKeys: keys.length };
  } catch (e) {
    // fallback : at least touch cache entry to ensure no throw
    try {
      CacheService.getScriptCache().put("sav:clearcache:ping", String(new Date().getTime()), 1);
    } catch (e2) {}
    return { ok: false, message: String(e && e.message ? e.message : e) };
  }
}

// (Analyse/KPIs supprimés à la demande : les fonctions savGetStats* / savGetKpis* ont été retirées.)

/** Compte les lignes avec numéro de dossier : ouverts vs archivés. */
function countOpenArch_(sheetName) {
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() <= 1) return { open: 0, arch: 0 };
  const lr = sh.getLastRow();
  const idx = savHeaderIndexByName_(sh);
  const etatCol = idx["État du dossier"] || (sheetName === SAV_SHEET_DIST ? 10 : 9);
  const nums = sh.getRange(2, 1, lr - 1, 1).getValues();
  const etats = sh.getRange(2, etatCol, lr - 1, 1).getValues();
  let open = 0;
  let arch = 0;
  for (let i = 0; i < nums.length; i++) {
    const num = String(nums[i][0] || "").trim();
    if (!num) continue;
    const e = String(etats[i][0] || "").trim();
    if (e === ETAT_ARCHIVE) arch++;
    else open++;
  }
  return { open, arch };
}

/** Tous les dossiers (ouverts + archivés) — le client filtre l’affichage. */
function savListCases() {
  ensureSavSheets_();
  const cached = savCacheGet_(SAV_CACHE_KEYS.CASES_ALL);
  if (cached && Array.isArray(cached)) return cached;
  const mpUnread = savMpUnreadCounts_();
  const out = [];
  out.push.apply(out, readCasesFrom_(SAV_SHEET_DIST, "Distributeur", null, null));
  out.push.apply(out, readCasesFrom_(SAV_SHEET_MP, "Marketplace", null, mpUnread));
  out.sort((a, b) => String(b.dateCreation || "").localeCompare(String(a.dateCreation || "")));
  savCachePut_(SAV_CACHE_KEYS.CASES_ALL, out);
  return out;
}

function savMpUnreadCounts_() {
  // { [numero]: count } pour messages IN non lus
  try {
    const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
    const sh = ss.getSheetByName(SAV_SHEET_MP_MESSAGES);
    if (!sh || sh.getLastRow() < 2) return {};
    ensureHeaderGeneric_(sh, HDR_MP_MESSAGES);
    const lr = sh.getLastRow();
    const vals = sh.getRange(2, 1, lr - 1, HDR_MP_MESSAGES.length).getValues();
    const m = {};
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i] || [];
      const numero = String(r[1] || "").trim(); // N° dossier
      if (!numero) continue;
      const sens = String(r[4] || "").trim(); // Sens
      const lu = String(r[8] || "").trim().toUpperCase(); // Lu
      if (sens === "IN" && lu !== "OUI") m[numero] = (m[numero] || 0) + 1;
    }
    return m;
  } catch (e) {
    return {};
  }
}

function readCasesFrom_(sheetName, typeLabel, includeArchived, mpUnreadMap) {
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const lr = sh.getLastRow();
  if (lr <= 1) return [];
  const width = Math.max(sh.getLastColumn(), sheetName === SAV_SHEET_DIST ? HDR_DIST.length : HDR_MP.length);
  const vals = sh.getRange(2, 1, lr - 1, width).getValues();
  const idx = savHeaderIndexByName_(sh); // 1-based
  const list = [];
  for (let i = 0; i < vals.length; i++) {
    const row = i + 2;
    const r = vals[i];
    const etatCol = idx["État du dossier"] || (sheetName === SAV_SHEET_DIST ? 10 : 9);
    const dateCol = idx["Date création"] || (sheetName === SAV_SHEET_DIST ? 9 : 8);
    const etat = String(r[etatCol - 1] || "").trim();
    if (includeArchived === false && etat === ETAT_ARCHIVE) continue;
    if (includeArchived === true && etat !== ETAT_ARCHIVE) continue;
    const base = {
      row,
      sheet: sheetName,
      type: typeLabel,
      numero: String(r[0] || ""),
      etat: etat || ETAT_CREATION,
      dateCreation: formatDate_(r[dateCol - 1]),
    };
    if (sheetName === SAV_SHEET_DIST) {
      list.push(
        Object.assign({}, base, {
          magasin: String(r[(idx["Nom magasin/distributeur"] || 2) - 1] || ""),
          adresse: String(r[(idx["Adresse complète"] || 3) - 1] || ""),
          modele: String(r[(idx["Modèle concerné"] || 4) - 1] || ""),
          panne: String(r[(idx["Panne constatée"] || 5) - 1] || ""),
          serie: String(r[(idx["Numéro de série"] || 6) - 1] || ""),
          refDistributeur: String(r[(idx["Référence dossier distributeur"] || 0) - 1] || ""),
          facture: String(r[(idx["Facture reçue"] || 7) - 1] || ""),
          photoPlaque: String(r[(idx["Photo plaque reçue"] || 8) - 1] || ""),
          mailAccordRetour: String(r[(idx["Mail envoyé au distributeur (accord retour atelier)"] || 0) - 1] || ""),
          dateMailAccordRetour: formatDateIfDate_(r[(idx["Date mail accord retour atelier"] || 0) - 1]),
        })
      );
    } else {
      list.push(
        Object.assign({}, base, {
          marketplace: String(r[(idx["Nom marketplace"] || 2) - 1] || ""),
          clientNom: String(r[(idx["Nom client"] || 3) - 1] || ""),
          clientEmail: String(r[(idx["Email client"] || 4) - 1] || ""),
          clientTel: String(r[(idx["Téléphone client"] || 5) - 1] || ""),
          clientAdresse: String(r[(idx["Adresse client"] || 6) - 1] || ""),
          demandeEnlevement: String(r[(idx["Demande d'enlèvement"] || 7) - 1] || ""),
          modele: String(r[(idx["Modèle concerné"] || 8) - 1] || ""),
          panne: String(r[(idx["Panne constatée"] || 9) - 1] || ""),
          facture: String(r[(idx["Numéro facture"] || 10) - 1] || ""),
          photoNs: String(r[(idx["Photo N°S reçue"] || 11) - 1] || ""),
          mpUnread: Number((mpUnreadMap && mpUnreadMap[String(r[0] || "").trim()]) || 0),
        })
      );
    }
  }
  return list;
}

function formatDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  }
  return String(v || "");
}

function savMissingRequired_(pairs) {
  const missing = [];
  (pairs || []).forEach((p) => {
    const label = String((p && p.label) || "").trim();
    const v = p ? p.value : "";
    const s = String(v == null ? "" : v).trim();
    if (!label) return;
    if (!s) missing.push(label);
  });
  return missing;
}

function savCreateDistributor(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const distRef = String(o.refDistributeur || "").trim();
  const missing = savMissingRequired_([
    { label: "Magasin", value: o.magasin },
    { label: "Adresse complète", value: o.adresse },
    { label: "Modèle", value: o.modele },
    { label: "Panne constatée", value: o.panne },
    { label: "Numéro de série", value: o.serie },
  ]);
  if (missing.length) {
    throw new Error("Impossible de valider la demande : champs obligatoires manquants — " + missing.join(", "));
  }
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_DIST);
  const num = nextSavNumberGlobal_(ss);
  const now = new Date();
  const facture = o.facture ? "OUI" : "NON";
  const photo = o.photoPlaque ? "OUI" : "NON";
  const row = [
    num,
    String(o.magasin || "").trim(),
    String(o.adresse || "").trim(),
    String(o.modele || "").trim(),
    String(o.panne || "").trim(),
    String(o.serie || "").trim(),
    facture,
    photo,
    now,
    ETAT_CREATION,
    "",
    "",
    "",
    "",
    distRef,
    "",
  ];
  sh.appendRow(row);
  savCacheInvalidateAll_();
  const rowDist = sh.getLastRow();
  try { savEnsureDriveFolderAndWriteUrl_(SAV_SHEET_DIST, rowDist, num); } catch(e) {}
  return { ok: true, numero: num, row: rowDist, sheet: SAV_SHEET_DIST };
}

function savCreateMarketplace(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const missing = savMissingRequired_([
    { label: "Marketplace", value: o.marketplace },
    { label: "Nom client", value: o.clientNom },
    { label: "Email client", value: o.clientEmail },
    { label: "Téléphone client", value: o.clientTel },
    { label: "Adresse client", value: o.clientAdresse },
    { label: "Modèle", value: o.modele },
    { label: "Panne constatée", value: o.panne },
  ]);
  if (missing.length) {
    throw new Error("Impossible de valider la demande : champs obligatoires manquants — " + missing.join(", "));
  }
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_MP);
  const now = new Date();
  const photoNs = o.photoNs ? "OUI" : "NON";
  const enlevement = o.demandeEnlevement ? "OUI" : "NON";

  // Déduplication: évite double-clic / double appel qui crée 2 dossiers
  const dupe = savIsDuplicateMarketplaceRequest_(sh, o, now);
  if (dupe) {
    // IMPORTANT: en cas de déduplication dossier, on ne recrée pas de demande transport.
    // La demande transport doit venir uniquement de la création initiale Marketplace.
    return { ok: true, deduped: true, numero: dupe.numero, row: dupe.rowMain, sheet: SAV_SHEET_MP };
  }

  const num = nextSavNumberGlobal_(ss);
  const row = [
    num,
    String(o.marketplace || "").trim(),
    String(o.clientNom || "").trim(),
    String(o.clientEmail || "").trim(),
    String(o.clientTel || "").trim(),
    String(o.clientAdresse || "").trim(),
    enlevement,
    String(o.modele || "").trim(),
    String(o.panne || "").trim(),
    String(o.numFacture || "").trim(),
    photoNs,
    now,
    ETAT_CREATION,
    "",
    "",
    "",
  ];
  sh.appendRow(row);
  savCacheInvalidateAll_();
  const rowMain = sh.getLastRow();

  // Si demande d'enlèvement: crée automatiquement une demande transport depuis l'adresse client
  if (enlevement === "OUI") {
    const DEST = "34 rue du moulin des bruyeres, Courbevoie 92400";
    try {
      savAddTransport({
        numero: num,
        typeEnvoi: "Enlèvement",
        pointDepart: String(o.clientAdresse || "").trim(),
        adresseArrivee: DEST,
        tracking: "",
        statut: "Demande",
        preuve: false,
        sheet: SAV_SHEET_MP,
        rowMain: rowMain,
        updateEtat: ETAT_TRANSPORT_DEMANDE,
      });
    } catch (e) {
      // Ne bloque pas la création du dossier si la demande transport échoue
    }
  }

  try { savEnsureDriveFolderAndWriteUrl_(SAV_SHEET_MP, rowMain, num); } catch(e) {}
  return { ok: true, numero: num, row: rowMain, sheet: SAV_SHEET_MP };
}

function savIsDuplicateMarketplaceRequest_(sh, o, now) {
  try {
    const email = String(o && o.clientEmail ? o.clientEmail : "").trim().toLowerCase();
    if (!email) return null;
    const tel = String(o && o.clientTel ? o.clientTel : "").trim();
    const nom = String(o && o.clientNom ? o.clientNom : "").trim();
    const mp = String(o && o.marketplace ? o.marketplace : "").trim();
    const adr = String(o && o.clientAdresse ? o.clientAdresse : "").trim();
    const modele = String(o && o.modele ? o.modele : "").trim();
    const panne = String(o && o.panne ? o.panne : "").trim();
    const numFacture = String(o && o.numFacture ? o.numFacture : "").trim();
    const photoNs = o && o.photoNs ? "OUI" : "NON";
    const enlevement = o && o.demandeEnlevement ? "OUI" : "NON";

    const lr = sh.getLastRow();
    if (lr < 2) return null;
    const from = Math.max(2, lr - 30); // check last 30 rows only
    const vals = sh.getRange(from, 1, lr - from + 1, HDR_MP.length).getValues();

    for (let i = vals.length - 1; i >= 0; i--) {
      const r = vals[i] || [];
      const rNum = String(r[0] || "").trim();
      if (!rNum) continue;
      const rMp = String(r[1] || "").trim();
      const rNom = String(r[2] || "").trim();
      const rEmail = String(r[3] || "").trim().toLowerCase();
      const rTel = String(r[4] || "").trim();
      const rAdr = String(r[5] || "").trim();
      const rEnl = String(r[6] || "").trim();
      const rModele = String(r[7] || "").trim();
      const rPanne = String(r[8] || "").trim();
      const rFact = String(r[9] || "").trim();
      const rPhoto = String(r[10] || "").trim();
      const rDate = r[11];

      if (rEmail !== email) continue;
      if (rTel !== tel) continue;
      if (rNom !== nom) continue;
      if (rMp !== mp) continue;
      if (rAdr !== adr) continue;
      if (rEnl !== enlevement) continue;
      if (rModele !== modele) continue;
      if (rPanne !== panne) continue;
      if (rFact !== numFacture) continue;
      if (rPhoto !== photoNs) continue;
      if (!(rDate instanceof Date) || isNaN(rDate.getTime())) continue;

      if (Math.abs(now.getTime() - rDate.getTime()) <= 2 * 60 * 1000) {
        return { numero: rNum, rowMain: from + i };
      }
      return null;
    }
    return null;
  } catch (e) {
    return null;
  }
}

function savUploadAttachments(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const sheetName = String(o.sheet || "").trim();
  const row = Number(o.row || 0);
  const numero = String(o.numero || "").trim();
  const files = Array.isArray(o.files) ? o.files : [];
  if (!sheetName || !row || row < 2 || !numero) return { ok: false, message: "Paramètres invalides (dossier)." };
  if (!files.length) return { ok: false, message: "Aucun fichier." };

  const parent = savGetDriveParentFolderForSheet_(sheetName);
  const folder = savGetOrCreateSubFolder_(parent, numero);

  let created = 0;
  for (let i = 0; i < files.length; i++) {
    const f = files[i] || {};
    const name = String(f.name || ("piece-" + (i + 1))).trim();
    const mimeType = String(f.mimeType || "application/octet-stream").trim();
    const b64 = String(f.base64 || "").trim();
    if (!b64) continue;
    const bytes = Utilities.base64Decode(b64);
    const blob = Utilities.newBlob(bytes, mimeType, name);
    folder.createFile(blob);
    created++;
  }

  const url = folder.getUrl();

  // écrire le lien Drive dans la feuille du dossier (colonne "Dossier Drive")
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  const idx = savHeaderIndexByName_(sh);
  const driveCol = idx["Dossier Drive"] || ((sheetName === SAV_SHEET_DIST ? HDR_DIST.length : HDR_MP.length) - 1);
  sh.getRange(row, driveCol).setValue(url);
  savCacheInvalidateAll_();
  return { ok: true, folderUrl: url, created };
}

function savGetOrCreateSubFolder_(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parentFolder.createFolder(name);
}

/**
 * Retourne (et crée si besoin) le sous-dossier Drive correspondant à un type.
 * Tout passe par le dossier racine SAV_DRIVE_ROOT_FOLDER_ID.
 * Les IDs des sous-dossiers sont mis en cache 6h pour éviter des appels Drive répétés.
 */
function savGetDriveSubFolder_(subName, cacheKey) {
  // 1. Tenter le cache (ID stocké → accès direct sans rescan)
  try {
    const cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return DriveApp.getFolderById(cached);
  } catch(e) {}

  // 2. Ouvrir le dossier racine et créer/récupérer le sous-dossier
  const root = DriveApp.getFolderById(SAV_DRIVE_ROOT_FOLDER_ID);
  const folder = savGetOrCreateSubFolder_(root, subName);

  // 3. Mettre en cache l'ID 6h
  try { CacheService.getScriptCache().put(cacheKey, folder.getId(), 6 * 3600); } catch(e) {}
  return folder;
}

function savGetDriveParentFolderForSheet_(sheetName) {
  const s = String(sheetName || "").trim();
  if (s === SAV_SHEET_MP) return savGetDriveSubFolder_(SAV_DRIVE_SUBDIR_MP, SAV_DRIVE_CACHE_MP);
  return savGetDriveSubFolder_(SAV_DRIVE_SUBDIR_DIST, SAV_DRIVE_CACHE_DIST);
}

function savGetDriveParentFolderForMisc_() {
  return savGetDriveSubFolder_(SAV_DRIVE_SUBDIR_MISC, SAV_DRIVE_CACHE_MISC);
}

function savGetDriveParentFolderForPdc_() {
  return savGetDriveSubFolder_(SAV_DRIVE_SUBDIR_PDC, SAV_DRIVE_CACHE_PDC);
}

function pdcNextNumber_(ss) {
  const y = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy");
  const prefix = "PDC-" + y + "-";
  const sh = ss.getSheetByName(SAV_SHEET_PDC);
  if (!sh || sh.getLastRow() < 2) return prefix + "001";
  ensureHeaderGeneric_(sh, HDR_PDC);
  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, 1).getValues();
  let max = 0;
  const re = new RegExp("^PDC-" + y + "-(\\d+)$");
  for (let i = 0; i < vals.length; i++) {
    const id = String(vals[i][0] || "").trim();
    const m = id.match(re);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }
  return prefix + String(max + 1).padStart(3, "0");
}

function pdcEnsureDriveFolderAndWriteUrl_(sh, row, id) {
  const parent = savGetDriveParentFolderForPdc_();
  const folder = savGetOrCreateSubFolder_(parent, id);
  const url = folder.getUrl();
  // colonne "Dossier Drive" (dernière)
  sh.getRange(row, HDR_PDC.length).setValue(url);
  return { folder, folderUrl: url };
}

function pdcCreate(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_PDC);
  ensureHeaderGeneric_(sh, HDR_PDC);
  const now = new Date();
  const poids = o.poids === "" || o.poids == null ? "" : Number(o.poids);
  const tr = String(o.transporteur || "DISTRIBUTEUR").trim();
  const tracking = String(o.tracking || "").trim();
  const notes = String(o.notes || "").trim();
  const pointDepart = String(o.pointDepart || "").trim();
  const adresseArrivee = String(o.adresseArrivee || "").trim();
  const refs = String(o.refs || "").trim();
  const designation = String(o.designation || "").trim();
  const colisPret = o.colisPret ? "OUI" : "NON";
  const dossierSav = String(o.dossierSav || "").trim();

  // Déduplication forte (quand la demande est renseignée) :
  // si une demande identique existe déjà (non livrée), ne jamais ajouter une 2e ligne
  // Déduplication: évite double-clic / double appel qui crée 2 lignes
  const dupe = pdcFindDuplicateRequest_(sh, {
    now,
    poids,
    transporteur: tr,
    tracking,
    notes,
    colisPret,
    pointDepart,
    adresseArrivee,
    refs,
    designation,
  });
  if (dupe) return { ok: true, deduped: true, id: dupe.id, row: dupe.row };

  const id = pdcNextNumber_(ss);
  sh.appendRow([id, now, poids, tr, tracking, LOG_STAT_A_TRAITER, "", "", notes, colisPret, pointDepart, adresseArrivee, refs, designation, dossierSav, ""]);
  const row = sh.getLastRow();

  // Si PDC rattachée à un dossier SAV : stocke l'ID PDC dans le dossier
  try {
    if (dossierSav) savAppendPdcToCase_(ss, dossierSav, id);
  } catch (e) {}

  // crée le dossier Drive et écrit l'URL
  try {
    pdcEnsureDriveFolderAndWriteUrl_(sh, row, id);
  } catch (e) {}
  savCacheInvalidateAll_();
  return { ok: true, id, row };
}

function pdcFindDuplicateRequest_(sh, o) {
  try {
    function norm_(s) {
      return String(s || "")
        .replace(/\s+/g, " ")
        .trim()
        .toLowerCase();
    }
    const now = o && o.now instanceof Date ? o.now : new Date();
    const lr = sh.getLastRow();
    if (lr < 2) return null;
    const from = Math.max(2, lr - 20);
    const vals = sh.getRange(from, 1, lr - from + 1, HDR_PDC.length).getValues();

    const poids = o && o.poids !== undefined ? String(o.poids) : "";
    const transporteur = norm_(o && o.transporteur !== undefined ? o.transporteur : "");
    const tracking = norm_(o && o.tracking !== undefined ? o.tracking : "");
    const notes = norm_(o && o.notes !== undefined ? o.notes : "");
    const colisPret = norm_(o && o.colisPret !== undefined ? o.colisPret : "");
    const pointDepart = norm_(o && o.pointDepart !== undefined ? o.pointDepart : "");
    const adresseArrivee = norm_(o && o.adresseArrivee !== undefined ? o.adresseArrivee : "");
    const refs = norm_(o && o.refs !== undefined ? o.refs : "");
    const designation = norm_(o && o.designation !== undefined ? o.designation : "");

    for (let i = vals.length - 1; i >= 0; i--) {
      const r = vals[i] || [];
      const id = String(r[0] || "").trim();
      if (!id) continue;
      const rDate = r[1];
      if (!(rDate instanceof Date) || isNaN(rDate.getTime())) continue;

      const statut = norm_(r[5]);
      if (statut === norm_(LOG_STAT_LIVRE) || statut === norm_(LOG_STAT_ANNULE)) continue;

      const rPoids = String(r[2] === "" || r[2] == null ? "" : r[2]);
      const rTr = norm_(r[3]);
      const rTracking = norm_(r[4]);
      const rNotes = norm_(r[8]);
      const rReady = norm_(r[9]);
      const rFrom = norm_(r[10]);
      const rTo = norm_(r[11]);
      const rRefs = norm_(r[12]);
      const rDes = norm_(r[13]);

      if (rPoids !== poids) continue;
      if (rTr !== transporteur) continue;
      if (rTracking !== tracking) continue;
      if (rNotes !== notes) continue;
      if (rReady !== colisPret) continue;
      if (rFrom !== pointDepart) continue;
      if (rTo !== adresseArrivee) continue;
      if (rRefs !== refs) continue;
      if (rDes !== designation) continue;

      if (Math.abs(now.getTime() - rDate.getTime()) <= 2 * 60 * 1000) {
        return { id, row: from + i };
      }
      return null;
    }
    return null;
  } catch (e) {
    return null;
  }
}

function pdcList(openOnly) {
  ensureSavSheets_();
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_PDC);
  if (!sh || sh.getLastRow() < 2) return [];
  ensureHeaderGeneric_(sh, HDR_PDC);
  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, HDR_PDC.length).getValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const row = i + 2;
    const r = vals[i] || [];
    const stat = String(r[5] || "").trim();
    if (openOnly && (stat === LOG_STAT_LIVRE || stat === LOG_STAT_ANNULE)) continue;
    out.push({
      pdcRow: row,
      id: String(r[0] || "").trim(),
      dateCreation: formatDate_(r[1]),
      poids: r[2] === "" ? "" : Number(r[2] || 0),
      transporteur: String(r[3] || "").trim(),
      tracking: String(r[4] || "").trim(),
      statut: stat,
      dateExpedition: formatDateIfDate_(r[6]),
      dateLivraison: formatDateIfDate_(r[7]),
      notes: String(r[8] || ""),
      colisPret: String(r[9] || "").trim(),
      pointDepart: String(r[10] || ""),
      adresseArrivee: String(r[11] || ""),
      refs: String(r[12] || ""),
      designation: String(r[13] || ""),
      dossierSav: String(r[14] || "").trim(),
      driveUrl: String(r[15] || ""),
    });
  }
  out.sort((a, b) => String(b.dateCreation || "").localeCompare(String(a.dateCreation || "")));
  return out;
}

function pdcListPage(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const tab = String(o.tab || "open").trim().toLowerCase(); // open|a|e|l|x|all
  const limit = Math.min(200, Math.max(10, Number(o.limit || 50)));
  const offset = Math.max(0, Number(o.offset || 0));
  const year = o.year === "" || o.year == null ? null : Number(o.year);

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_PDC);
  if (!sh || sh.getLastRow() < 2) return { rows: [], hasMore: false };
  ensureHeaderGeneric_(sh, HDR_PDC);

  function wantStat_(s) {
    const stat = String(s || "").trim();
    if (tab === "all") return true;
    if (tab === "open") return stat === LOG_STAT_A_TRAITER || stat === LOG_STAT_EXPEDIE;
    if (tab === "a") return stat === LOG_STAT_A_TRAITER;
    if (tab === "e") return stat === LOG_STAT_EXPEDIE;
    if (tab === "l") return stat === LOG_STAT_LIVRE;
    if (tab === "x") return stat === LOG_STAT_ANNULE;
    return stat === LOG_STAT_A_TRAITER || stat === LOG_STAT_EXPEDIE;
  }

  function wantYear_(d) {
    if (!year) return true;
    if (!(d instanceof Date) || isNaN(d.getTime())) return false;
    return d.getFullYear() === year;
  }

  function toObj_(rowIndex1, r) {
    const stat = String(r[5] || "").trim();
    return {
      pdcRow: rowIndex1,
      id: String(r[0] || "").trim(),
      dateCreation: formatDate_(r[1]),
      poids: r[2] === "" ? "" : Number(r[2] || 0),
      transporteur: String(r[3] || "").trim(),
      tracking: String(r[4] || "").trim(),
      statut: stat,
      dateExpedition: formatDateIfDate_(r[6]),
      dateLivraison: formatDateIfDate_(r[7]),
      notes: String(r[8] || ""),
      colisPret: String(r[9] || "").trim(),
      pointDepart: String(r[10] || ""),
      adresseArrivee: String(r[11] || ""),
      refs: String(r[12] || ""),
      designation: String(r[13] || ""),
      dossierSav: String(r[14] || "").trim(),
      driveUrl: String(r[15] || ""),
    };
  }

  // Scan depuis le bas pour éviter de charger des années entières si inutile
  const lr = sh.getLastRow();
  const chunk = 500;
  let start = lr;
  let skipped = 0;
  const out = [];
  let morePossible = false;

  while (start >= 2 && out.length < limit) {
    const from = Math.max(2, start - chunk + 1);
    const num = start - from + 1;
    const vals = sh.getRange(from, 1, num, HDR_PDC.length).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      const row1 = from + i;
      const r = vals[i] || [];
      const id = String(r[0] || "").trim();
      if (!id) continue;
      const dCreate = r[1];
      if (!wantYear_(dCreate)) continue;
      const stat = String(r[5] || "").trim();
      if (!wantStat_(stat)) continue;

      if (skipped < offset) {
        skipped++;
        continue;
      }
      out.push(toObj_(row1, r));
      if (out.length >= limit) break;
    }
    start = from - 1;
  }

  // Détermine s'il reste potentiellement d'autres lignes
  if (start >= 2) morePossible = true;
  return { rows: out, hasMore: morePossible };
}

function pdcSave(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const pdcRow = Number(o.pdcRow || 0);
  const action = String(o.action || "").toLowerCase();
  if (!pdcRow || pdcRow < 2) return { ok: false, message: "Ligne PDC invalide." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_PDC);
  if (!sh || sh.getLastRow() < pdcRow) return { ok: false, message: "Ligne PDC introuvable." };
  ensureHeaderGeneric_(sh, HDR_PDC);
  const r = sh.getRange(pdcRow, 1, 1, HDR_PDC.length).getValues()[0] || [];
  const id = String(r[0] || "").trim();
  const currentStatut = String(r[5] || "").trim();
  if (currentStatut === LOG_STAT_ANNULE) return { ok: false, message: "Demande PDC annulée (modification impossible)." };
  const poids = o.poids !== undefined ? (o.poids === "" ? "" : Number(o.poids)) : r[2];
  const transporteur = o.transporteur !== undefined ? String(o.transporteur || "").trim() : String(r[3] || "").trim();
  const tracking = o.tracking !== undefined ? String(o.tracking || "").trim() : String(r[4] || "").trim();
  const notes = o.notes !== undefined ? String(o.notes || "").trim() : String(r[8] || "").trim();
  const colisPret = o.colisPret !== undefined ? (o.colisPret ? "OUI" : "NON") : String(r[9] || "").trim();
  const pointDepart = o.pointDepart !== undefined ? String(o.pointDepart || "").trim() : String(r[10] || "").trim();
  const adresseArrivee = o.adresseArrivee !== undefined ? String(o.adresseArrivee || "").trim() : String(r[11] || "").trim();
  const refs = o.refs !== undefined ? String(o.refs || "").trim() : String(r[12] || "").trim();
  const designation = o.designation !== undefined ? String(o.designation || "").trim() : String(r[13] || "").trim();
  const dossierSav = o.dossierSav !== undefined ? String(o.dossierSav || "").trim() : String(r[14] || "").trim();

  sh.getRange(pdcRow, 3).setValue(poids);
  sh.getRange(pdcRow, 4).setValue(transporteur);
  sh.getRange(pdcRow, 5).setValue(tracking);
  sh.getRange(pdcRow, 9).setValue(notes);
  sh.getRange(pdcRow, 10).setValue(colisPret);
  sh.getRange(pdcRow, 11).setValue(pointDepart);
  sh.getRange(pdcRow, 12).setValue(adresseArrivee);
  sh.getRange(pdcRow, 13).setValue(refs);
  sh.getRange(pdcRow, 14).setValue(designation);
  sh.getRange(pdcRow, 15).setValue(dossierSav);

  if (action === "expedier") {
    sh.getRange(pdcRow, 6).setValue(LOG_STAT_EXPEDIE);
    sh.getRange(pdcRow, 7).setValue(new Date());
  } else if (action === "livrer") {
    sh.getRange(pdcRow, 6).setValue(LOG_STAT_LIVRE);
    const dl = new Date(String(o.dateLivraison || ""));
    if (isNaN(dl.getTime())) return { ok: false, message: "Date livraison invalide." };
    sh.getRange(pdcRow, 8).setValue(dl);
  } else if (action === "save") {
    // rien
  } else {
    return { ok: false, message: "Action invalide (save|expedier|livrer)." };
  }

  // s'assure dossier Drive
  try {
    if (id) pdcEnsureDriveFolderAndWriteUrl_(sh, pdcRow, id);
  } catch (e) {}

  savCacheInvalidateAll_();
  return { ok: true };
}

function pdcCancel(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const pdcRow = Number(o.pdcRow || 0);
  const reason = String(o.reason || "").trim();
  if (!pdcRow || pdcRow < 2) return { ok: false, message: "Ligne PDC invalide." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_PDC);
  if (!sh || sh.getLastRow() < pdcRow) return { ok: false, message: "Ligne PDC introuvable." };
  ensureHeaderGeneric_(sh, HDR_PDC);
  const r = sh.getRange(pdcRow, 1, 1, HDR_PDC.length).getValues()[0] || [];
  const currentStatut = String(r[5] || "").trim();
  if (currentStatut === LOG_STAT_LIVRE) return { ok: false, message: "Demande déjà livrée (annulation impossible)." };
  if (currentStatut === LOG_STAT_ANNULE) return { ok: true, cancelled: true };
  const prevNotes = String(r[8] || "").trim();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  const add = "[ANNULATION " + stamp + "]" + (reason ? " " + reason : "");
  const newNotes = prevNotes ? (prevNotes + "\n" + add) : add;
  sh.getRange(pdcRow, 6).setValue(LOG_STAT_ANNULE);
  sh.getRange(pdcRow, 9).setValue(newNotes);
  savCacheInvalidateAll_();
  return { ok: true, cancelled: true };
}

function pdcUploadDocs(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const pdcRow = Number(o.pdcRow || 0);
  const files = Array.isArray(o.files) ? o.files : [];
  if (!pdcRow || pdcRow < 2) return { ok: false, message: "Ligne PDC invalide." };
  if (!files.length) return { ok: false, message: "Aucun fichier." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_PDC);
  if (!sh || sh.getLastRow() < pdcRow) return { ok: false, message: "Ligne PDC introuvable." };
  ensureHeaderGeneric_(sh, HDR_PDC);
  const r = sh.getRange(pdcRow, 1, 1, HDR_PDC.length).getValues()[0] || [];
  const id = String(r[0] || "").trim();
  if (!id) return { ok: false, message: "Id PDC manquant." };

  const parent = savGetDriveParentFolderForPdc_();
  const folder = savGetOrCreateSubFolder_(parent, id);
  const docs = savGetOrCreateSubFolder_(folder, "Documents");

  let created = 0;
  for (let i = 0; i < files.length; i++) {
    const f = files[i] || {};
    const name = String(f.name || ("doc-" + (i + 1))).trim();
    const mimeType = String(f.mimeType || "application/octet-stream").trim();
    const b64 = String(f.base64 || "").trim();
    if (!b64) continue;
    const bytes = Utilities.base64Decode(b64);
    docs.createFile(Utilities.newBlob(bytes, mimeType, name));
    created++;
  }
  try {
    sh.getRange(pdcRow, HDR_PDC.length).setValue(folder.getUrl());
  } catch (e) {}
  savCacheInvalidateAll_();
  return { ok: true, created, folderUrl: folder.getUrl() };
}

function savHeaderIndexByName_(sh) {
  const row = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];
  const m = {};
  for (let i = 0; i < row.length; i++) {
    const k = String(row[i] || "").trim();
    if (k) m[k] = i + 1; // 1-based col
  }
  return m;
}

function savEnsureDriveFolderAndWriteUrl_(sheetName, row, numero) {
  const parent = savGetDriveParentFolderForSheet_(sheetName);
  const folder = savGetOrCreateSubFolder_(parent, numero);
  const url = folder.getUrl();

  // écrire le lien Drive dans la feuille du dossier (colonne "Dossier Drive")
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  const idx = savHeaderIndexByName_(sh);
  const driveCol = idx["Dossier Drive"] || ((sheetName === SAV_SHEET_DIST ? HDR_DIST.length : HDR_MP.length) - 1);
  sh.getRange(row, driveCol).setValue(url);
  return { folder, folderUrl: url };
}

/**
 * Backfill Drive folders for existing cases in Sheets.
 * Creates the folder if missing and writes "Dossier Drive" URL.
 *
 * payload:
 * - sheet: optional ("SAV 2026" or "SAV Marketplace"), default: both
 * - startRow: optional (>=2), default 2
 * - limit: optional, default 200
 * - dryRun: optional boolean (if true, doesn't write)
 */
function savBackfillDriveFolders(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const sheetFilter = String(o.sheet || "").trim();
  const startRow = Math.max(2, Number(o.startRow || 2));
  const limit = Math.max(10, Math.min(800, Number(o.limit || 200)));
  const dryRun = !!o.dryRun;
  const recheckExisting = !!o.recheckExisting;
  const forceRecreate = !!o.forceRecreate;

  const lock = LockService.getScriptLock();
  lock.waitLock(30 * 1000);
  try {
    const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
    const targets = [];
    if (!sheetFilter || sheetFilter === SAV_SHEET_DIST) targets.push(SAV_SHEET_DIST);
    if (!sheetFilter || sheetFilter === SAV_SHEET_MP) targets.push(SAV_SHEET_MP);

    const result = { ok: true, dryRun, sheets: [], totalScanned: 0, totalCreated: 0, totalAlready: 0, totalSkipped: 0 };

    targets.forEach((sheetName) => {
      const sh = ss.getSheetByName(sheetName);
      if (!sh) return;
      ensureHeaderGeneric_(sh, sheetName === SAV_SHEET_DIST ? HDR_DIST : HDR_MP);
      const idx = savHeaderIndexByName_(sh);
      const numCol = idx["Numéro de dossier"] || 1;
      const driveCol = idx["Dossier Drive"] || ((sheetName === SAV_SHEET_DIST ? HDR_DIST.length : HDR_MP.length) - 1);
      const lr = sh.getLastRow();
      if (lr < startRow) {
        result.sheets.push({ sheet: sheetName, scanned: 0, created: 0, already: 0, skipped: 0, nextStartRow: null });
        return;
      }
      const toRead = Math.min(limit, lr - startRow + 1);
      const vals = sh.getRange(startRow, 1, toRead, Math.max(numCol, driveCol)).getValues();
      let created = 0;
      let already = 0;
      let skipped = 0;

      // Parent differs by sheet (Distributeur vs Marketplace)
      const parent = savGetDriveParentFolderForSheet_(sheetName);

      function extractFolderId_(driveUrl) {
        const u = String(driveUrl || "").trim();
        if (!u) return "";
        // Typical formats:
        // https://drive.google.com/drive/folders/<id>
        // https://drive.google.com/drive/folders/<id>?...
        const m = u.match(/\/folders\/([^\/\?\#]+)/i);
        if (m && m[1]) return String(m[1]).trim();
        // fallback: last path segment
        const parts = u.split("/");
        return String(parts[parts.length - 1] || "").split("?")[0].trim();
      }

      for (let i = 0; i < vals.length; i++) {
        const row = startRow + i;
        const r = vals[i] || [];
        const numero = String(r[numCol - 1] || "").trim();
        const driveUrl = String(r[driveCol - 1] || "").trim();
        if (!numero) {
          skipped++;
          continue;
        }
        if (driveUrl) {
          if (!recheckExisting && !forceRecreate) {
            already++;
            continue;
          }
          // Recheck: if the folder doesn't exist anymore / not accessible, recreate.
          if (!forceRecreate) {
            try {
              const fid = extractFolderId_(driveUrl);
              if (fid) {
                const f = DriveApp.getFolderById(fid);
                // If the name matches, consider OK; else we recreate to be safe.
                const nameOk = (() => {
                  try {
                    return String(f.getName && f.getName() || "").trim() === numero;
                  } catch (e) {
                    return true;
                  }
                })();
                if (nameOk) {
                  already++;
                  continue;
                }
              }
            } catch (e) {
              // will recreate below
            }
          }
          // if we reach here: recreate + update url (either forceRecreate=true or recheck failed)
        }
        // Create or reuse folder by name = numero
        try {
          const folder = savGetOrCreateSubFolder_(parent, numero);
          const url = folder.getUrl();
          if (!dryRun) sh.getRange(row, driveCol).setValue(url);
          created++;
        } catch (e) {
          // If Drive fails (permissions/quota), stop and return a helpful error
          const msg = e && e.message ? String(e.message) : String(e);
          throw new Error("Drive error sur " + sheetName + " ligne " + row + " (dossier " + numero + ") : " + msg);
        }
        // avoid Drive quotas in bulk operations
        if ((i + 1) % 25 === 0) Utilities.sleep(200);
      }

      const scanned = vals.length;
      const nextStartRow = (startRow + scanned) <= lr ? (startRow + scanned) : null;
      result.sheets.push({ sheet: sheetName, scanned, created, already, skipped, nextStartRow });
      result.totalScanned += scanned;
      result.totalCreated += created;
      result.totalAlready += already;
      result.totalSkipped += skipped;
    });

    savCacheInvalidateAll_();
    return result;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function savDriveAccessCheck() {
  ensureSavSheets_();

  function checkFolder(label, getFolderFn) {
    try {
      const f = getFolderFn();
      const name = f.getName();
      const id = f.getId();
      // Test écriture : crée un sous-dossier temporaire et le met à la corbeille
      const testName = "_SAV_TEST_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
      let sub;
      try {
        sub = f.createFolder(testName);
        sub.setTrashed(true);
      } catch (eWrite) {
        return { label, ok: false, id, name, phase: "write", message: String(eWrite.message || eWrite) };
      }
      return { label, ok: true, name, id };
    } catch (e) {
      return { label, ok: false, phase: "read", message: String(e.message || e) };
    }
  }

  return {
    ok: true,
    rootId: SAV_DRIVE_ROOT_FOLDER_ID,
    checks: [
      checkFolder("Dossier racine SAV",      () => DriveApp.getFolderById(SAV_DRIVE_ROOT_FOLDER_ID)),
      checkFolder("SAV Distributeur",        () => savGetDriveParentFolderForSheet_(SAV_SHEET_DIST)),
      checkFolder("SAV Marketplace",         () => savGetDriveParentFolderForSheet_(SAV_SHEET_MP)),
      checkFolder("PDC",                     () => savGetDriveParentFolderForPdc_()),
      checkFolder("Logistique & Autres",     () => savGetDriveParentFolderForMisc_()),
    ],
  };
}

function savGeneratePlaqueSheetPdf(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const sheetName = String(o.sheet || "").trim();
  const row = Number(o.row || 0);
  const numero = String(o.numero || "").trim();
  if (!sheetName || !row || row < 2 || !numero) return { ok: false, message: "Paramètres invalides (dossier)." };

  // Pour l'instant : procédure 1 = dossiers Distributeur
  if (sheetName !== SAV_SHEET_DIST) return { ok: false, message: "Feuille plaque : disponible uniquement pour les dossiers Distributeur." };

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return { ok: false, message: "Feuille dossier introuvable." };
  ensureHeaderGeneric_(sh, HDR_DIST);

  const r = sh.getRange(row, 1, 1, HDR_DIST.length).getValues()[0];
  const magasin = String(r[1] || "").trim();
  const adresse = String(r[2] || "").trim();
  const modele = String(r[3] || "").trim();
  const panne = String(r[4] || "").trim();
  const serie = String(r[5] || "").trim();

  const { folder, folderUrl } = savEnsureDriveFolderAndWriteUrl_(sheetName, row, numero);

  const esc = savHtmlEsc_;
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  const html =
    '<html><head><meta charset="utf-8"/>' +
    '<style>' +
    'body{font-family:Arial,sans-serif;font-size:12px;color:#0f172a;margin:24px}' +
    'h1{margin:0 0 8px 0;font-size:16px;text-align:center}' +
    '.muted{color:#475569}' +
    '.sec{margin-top:12px;border-top:1px solid #cbd5e1;padding-top:10px}' +
    '.row{margin:4px 0}' +
    '.lbl{display:inline-block;width:160px;font-weight:bold}' +
    '.box{margin-top:10px;border:2px solid #0f172a;height:320px;border-radius:10px}' +
    '.box .hint{padding:10px;color:#475569}' +
    '.sign{margin-top:14px;border-top:1px solid #cbd5e1;padding-top:10px}' +
    '</style></head><body>' +
    '<h1>FEUILLE RETOUR — PLAQUE SIGNALÉTIQUE</h1>' +
    '<div class="muted" style="text-align:center">Dossier SAV : <b>' + esc(numero) + '</b> · Date : ' + esc(today) + "</div>" +
    '<div class="sec"><div style="font-weight:900;margin-bottom:6px">ADRESSE DE RETOUR</div>' +
    '<div class="row"><span class="lbl">Envoyer à :</span> 34 rue du moulin des bruyere, Courbevoie 92400</div>' +
    "</div>" +
    '<div class="sec"><div style="font-weight:900;margin-bottom:6px">INFOS DOSSIER</div>' +
    '<div class="row"><span class="lbl">Magasin / Distributeur :</span> ' + esc(magasin) + "</div>" +
    '<div class="row"><span class="lbl">Adresse :</span> ' + esc(adresse) + "</div>" +
    "</div>" +
    '<div class="sec"><div style="font-weight:900;margin-bottom:6px">INFOS APPAREIL</div>' +
    '<div class="row"><span class="lbl">Modèle :</span> ' + esc(modele) + "</div>" +
    '<div class="row"><span class="lbl">N° série :</span> ' + esc(serie) + "</div>" +
    "</div>" +
    '<div class="sec"><div style="font-weight:900;margin-bottom:6px">PANNE</div>' +
    '<div class="row"><span class="lbl">Description :</span> ' + esc(panne) + "</div>" +
    "</div>" +
    '<div class="sec"><div style="font-weight:900;margin-bottom:6px">COLLER LA PLAQUE SIGNALÉTIQUE ICI</div>' +
    '<div class="box"><div class="hint">Espace réservé pour coller la plaque (photo / photocopie).</div></div>' +
    "</div>" +
    '<div class="sign"><div style="font-weight:900;margin-bottom:6px">SIGNATURE</div>' +
    '<div class="row"><span class="lbl">Signature :</span> ____________________________</div>' +
    '<div class="row"><span class="lbl">Date :</span> ____________________________</div>' +
    "</div>" +
    "</body></html>";

  const blob = Utilities.newBlob(html, "text/html", "feuille.html").getAs("application/pdf");
  const fileName = "Feuille_Retour_Plaque_" + numero + ".pdf";
  const f = folder.createFile(blob.setName(fileName));
  return { ok: true, folderUrl, fileUrl: f.getUrl(), fileName };
}

function savReadCaseMainInfo_(sheetName, row) {
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return null;
  ensureHeaderGeneric_(sh, sheetName === SAV_SHEET_DIST ? HDR_DIST : HDR_MP);
  const idx = savHeaderIndexByName_(sh);
  const width = sheetName === SAV_SHEET_DIST ? HDR_DIST.length : HDR_MP.length;
  const r = sh.getRange(row, 1, 1, width).getValues()[0] || [];
  const out = { sheetName, row };
  if (sheetName === SAV_SHEET_DIST) {
    out.type = "Distributeur";
    out.numero = String(r[0] || "").trim();
    out.magasin = String(r[(idx["Nom magasin/distributeur"] || 2) - 1] || "").trim();
    out.adresse = String(r[(idx["Adresse complète"] || 3) - 1] || "").trim();
    out.modele = String(r[(idx["Modèle concerné"] || 4) - 1] || "").trim();
    out.panne = String(r[(idx["Panne constatée"] || 5) - 1] || "").trim();
    out.serie = String(r[(idx["Numéro de série"] || 6) - 1] || "").trim();
    // pas d'email client dans les dossiers distributeur (par design actuel)
    out.clientNom = "";
    out.clientEmail = "";
    out.clientTel = "";
  } else {
    out.type = "Marketplace";
    out.numero = String(r[0] || "").trim();
    out.marketplace = String(r[(idx["Nom marketplace"] || 2) - 1] || "").trim();
    out.clientNom = String(r[(idx["Nom client"] || 3) - 1] || "").trim();
    out.clientEmail = String(r[(idx["Email client"] || 4) - 1] || "").trim();
    out.clientTel = String(r[(idx["Téléphone client"] || 5) - 1] || "").trim();
    out.adresse = String(r[(idx["Adresse client"] || 6) - 1] || "").trim();
    out.modele = String(r[(idx["Modèle concerné"] || 8) - 1] || "").trim();
    out.panne = String(r[(idx["Panne constatée"] || 9) - 1] || "").trim();
    out.serie = ""; // pas stocké systématiquement en marketplace
  }
  return out;
}

function savGetProcIdForModel_(modele) {
  const m = String(modele || "").trim();
  if (!m) return "";
  const map = savProcedureModelToProcId_();
  return String(map[m] || "").trim();
}

function savReturnAddressHtml_() {
  // Adresse imposée par le prompt
  return (
    "OPTIMEA - SAV<br/>" +
    "34 Rue du Moulin des Bruyères<br/>" +
    "92400 COURBEVOIE<br/>" +
    "FRANCE"
  );
}

function savReturnAddressText_() {
  return "OPTIMEA - SAV\n34 Rue du Moulin des Bruyères\n92400 COURBEVOIE\nFRANCE";
}


function savSendEmailFromSav_(to, subject, bodyText, attachments) {
  try {
    const dest = String(to || "").trim();
    if (!dest) return { ok: false, message: "Destinataire email requis." };

    const atts = attachments && attachments.length ? attachments : [];
    const opts = { name: "OPTIMEA - SAV" };
    if (atts && atts.length) opts.attachments = atts;

    MailApp.sendEmail({ to: dest, subject: String(subject || ""), body: String(bodyText || ""), ...opts });
    let sender = "MailApp";
    try { sender = Session.getActiveUser().getEmail(); } catch (e) {}
    return { ok: true, from: sender };
  } catch (e) {
    const msg = e && e.message ? String(e.message) : String(e);
    return { ok: false, message: msg };
  }
}

function savSendEmailFromSavHtml_(to, subject, htmlBody, attachments) {
  try {
    const dest = String(to || "").trim();
    if (!dest) return { ok: false, message: "Destinataire email requis." };

    const atts = attachments && attachments.length ? attachments : [];
    const opts = { name: "OPTIMEA - SAV" };
    if (atts && atts.length) opts.attachments = atts;

    MailApp.sendEmail({
      to: dest,
      subject: String(subject || ""),
      htmlBody: String(htmlBody || ""),
      ...opts,
    });
    let sender = "MailApp";
    try { sender = Session.getActiveUser().getEmail(); } catch (e) {}
    return { ok: true, from: sender };
  } catch (e) {
    const msg = e && e.message ? String(e.message) : String(e);
    return { ok: false, message: msg };
  }
}

function savCreatePdfFile_(folder, fileName, html) {
  const blob = Utilities.newBlob(html, "text/html", "doc.html").getAs("application/pdf");
  return folder.createFile(blob.setName(fileName));
}

function savGenerateProc1PostalRequestPdf(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const sheetName = String(o.sheet || "").trim();
  const row = Number(o.row || 0);
  const numero = String(o.numero || "").trim();
  if (!sheetName || !row || row < 2 || !numero) return { ok: false, message: "Paramètres invalides (dossier)." };

  const info = savReadCaseMainInfo_(sheetName, row);
  if (!info) return { ok: false, message: "Dossier introuvable." };
  const pid = savGetProcIdForModel_(info.modele);
  if (pid !== SAV_PROC_POSTAL) return { ok: false, message: "Ce dossier n'est pas en Procédure 1 (postal)." };

  const { folder, folderUrl } = savEnsureDriveFolderAndWriteUrl_(sheetName, row, numero);
  const docs = savGetOrCreateSubFolder_(folder, "Étiquettes & courriers");

  const esc = savHtmlEsc_;
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  const html =
    '<html><head><meta charset="utf-8"/>' +
    '<style>' +
    'body{font-family:Arial,sans-serif;color:#0f172a;margin:26px}' +
    'h1{margin:0 0 8px 0;font-size:18px;text-align:center}' +
    '.muted{color:#475569}' +
    '.card{margin-top:14px;border:1px solid #cbd5e1;border-radius:12px;padding:12px}' +
    '.row{margin:6px 0;font-size:13px}' +
    '.lbl{display:inline-block;width:150px;font-weight:700}' +
    '.big{font-size:20px;font-weight:900;letter-spacing:1px}' +
    '.addr{line-height:1.4}' +
    '</style></head><body>' +
    savPdfLogoHtml_() +
    '<h1>DEMANDE DE RETOUR - PLAQUE SIGNALÉTIQUE</h1>' +
    '<div class="muted" style="text-align:center">Dossier SAV : <span class="big">#' +
    esc(numero) +
    '</span> · Date : ' +
    esc(today) +
    "</div>" +
    '<div class="card">' +
    '<div class="row"><span class="lbl">Appareil :</span> ' +
    esc(info.modele) +
    "</div>" +
    '<div class="row"><span class="lbl">Client :</span> ' +
    esc(info.clientNom || info.magasin || "—") +
    "</div>" +
    '<div class="row"><span class="lbl">Téléphone :</span> ' +
    esc(info.clientTel || "") +
    "</div>" +
    '<div class="row"><span class="lbl">Email :</span> ' +
    esc(info.clientEmail || "") +
    "</div>" +
    '<div class="row"><span class="lbl">Adresse :</span> ' +
    esc(info.adresse || "") +
    "</div>" +
    "</div>" +
    '<div class="card">' +
    '<div style="font-weight:900;margin-bottom:8px">Adresse de retour</div>' +
    '<div class="addr">' +
    savReturnAddressHtml_() +
    "</div>" +
    "</div>" +
    '<div class="card">' +
    '<div style="font-weight:900;margin-bottom:8px">Instruction</div>' +
    '<div class="row">Veuillez nous renvoyer la <b>plaque signalétique</b> de votre appareil par <b>courrier postal</b> à l’adresse ci-dessus.</div>' +
    "</div>" +
    "</body></html>";

  const fileName = "PROC1_Demande_Retour_Plaque_" + numero + ".pdf";
  const f = savCreatePdfFile_(docs, fileName, html);
  return { ok: true, folderUrl, fileUrl: f.getUrl(), fileName };
}

function savGenerateReturnLabelPdf(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const sheetName = String(o.sheet || "").trim();
  const row = Number(o.row || 0);
  const numero = String(o.numero || "").trim();
  if (!sheetName || !row || row < 2 || !numero) return { ok: false, message: "Paramètres invalides (dossier)." };

  const info = savReadCaseMainInfo_(sheetName, row);
  if (!info) return { ok: false, message: "Dossier introuvable." };
  const pid = savGetProcIdForModel_(info.modele);
  if (pid !== SAV_PROC_RAPATRIEMENT && pid !== SAV_PROC_VISIO) {
    return { ok: false, message: "Étiquette retour uniquement pour Procédure 2 & 3." };
  }

  const { folder, folderUrl } = savEnsureDriveFolderAndWriteUrl_(sheetName, row, numero);
  const docs = savGetOrCreateSubFolder_(folder, "Étiquettes & courriers");

  const esc = savHtmlEsc_;
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const title = "ÉTIQUETTE DE RETOUR - SAV OPTIMEA";

  const html =
    '<html><head><meta charset="utf-8"/>' +
    '<style>' +
    'body{font-family:Arial,sans-serif;color:#0f172a;margin:24px}' +
    'h1{margin:0 0 6px 0;font-size:18px;text-align:center}' +
    '.muted{color:#475569}' +
    '.box{margin-top:12px;border:2px solid #0f172a;border-radius:14px;padding:14px}' +
    '.big{font-size:22px;font-weight:900;letter-spacing:1px}' +
    '.row{margin:6px 0;font-size:13px}' +
    '.lbl{display:inline-block;width:120px;font-weight:800}' +
    '.addr{margin-top:10px;line-height:1.4}' +
    '</style></head><body>' +
    savPdfLogoHtml_() +
    '<h1>' +
    esc(title) +
    "</h1>" +
    '<div style="text-align:center;margin-top:2px;font-weight:900">À COLLER À UN ENDROIT VISIBLE</div>' +
    '<div class="muted" style="text-align:center">À coller sur l’appareil / colis · Généré le ' +
    esc(today) +
    "</div>" +
    '<div class="box">' +
    '<div class="row"><span class="lbl">Dossier :</span> <span class="big">#' +
    esc(numero) +
    "</span></div>" +
    '<div class="row"><span class="lbl">Modèle :</span> ' +
    esc(info.modele) +
    "</div>" +
    (info.serie ? '<div class="row"><span class="lbl">N° série :</span> ' + esc(info.serie) + "</div>" : "") +
    '<div class="row"><span class="lbl">Client :</span> ' +
    esc(info.clientNom || info.magasin || "—") +
    "</div>" +
    (info.clientTel ? '<div class="row"><span class="lbl">Tél :</span> ' + esc(info.clientTel) + "</div>" : "") +
    (info.clientEmail ? '<div class="row"><span class="lbl">Email :</span> ' + esc(info.clientEmail) + "</div>" : "") +
    '<div class="row"><span class="lbl">Retour à :</span></div>' +
    '<div class="addr">' +
    savReturnAddressHtml_() +
    "</div>" +
    "</div>" +
    "</body></html>";

  const procTag = pid === SAV_PROC_VISIO ? "PROC3" : "PROC2";
  const fileName = procTag + "_Etiquette_Retour_" + numero + ".pdf";
  // Avoid duplicates: if the expected file already exists, reuse it.
  try {
    const it = docs.getFilesByName(fileName);
    if (it && it.hasNext()) {
      const f = it.next();
      return { ok: true, folderUrl, fileUrl: f.getUrl(), fileName, fileId: f.getId(), already: true };
    }
  } catch (e) {
    // ignore and regenerate
  }
  const f = savCreatePdfFile_(docs, fileName, html);
  return { ok: true, folderUrl, fileUrl: f.getUrl(), fileName, fileId: f.getId() };
}

function savGenerateLogistiqueShippingLabelPdf(payload) {
  // Wrapper public (appelable via google.script.run depuis le client)
  ensureSavSheets_();
  return savGenerateLogistiqueShippingLabelPdf_(payload || {});
}

function savGenerateLogistiqueShippingLabelPdf_(o) {
  const sheetName = String(o && o.sheet ? o.sheet : "").trim();
  const row = Number(o && o.rowMain ? o.rowMain : 0);
  const numero = String(o && o.numero ? o.numero : "").trim();
  const typeEnvoi = String(o && o.typeEnvoi ? o.typeEnvoi : "").trim();
  const toAddr = String(o && o.adresseArrivee ? o.adresseArrivee : "").trim();
  if (!sheetName || !row || row < 2 || !numero) return null;
  if (!toAddr) return null;

  // Pull client/dossier info from the SAV sheet when available
  const infoMain = savReadCaseMainInfo_(sheetName, row) || {};

  const { folder, folderUrl } = savEnsureDriveFolderAndWriteUrl_(sheetName, row, numero);
  const docs = savGetOrCreateSubFolder_(folder, "Étiquettes & courriers");

  const esc = savHtmlEsc_;
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  const html =
    '<html><head><meta charset="utf-8"/>' +
    '<style>' +
    'body{font-family:Arial,sans-serif;color:#0f172a;margin:24px}' +
    'h1{margin:0 0 6px 0;font-size:18px;text-align:center}' +
    '.muted{color:#475569}' +
    '.box{margin-top:12px;border:2px solid #0f172a;border-radius:14px;padding:14px}' +
    '.big{font-size:22px;font-weight:900;letter-spacing:1px}' +
    '.row{margin:8px 0;font-size:13px}' +
    '.lbl{display:block;font-weight:900;margin-bottom:4px}' +
    '.addr{white-space:pre-wrap;line-height:1.35}' +
    '</style></head><body>' +
    savPdfLogoHtml_() +
    '<h1>ÉTIQUETTE D’EXPÉDITION - LOGISTIQUE</h1>' +
    '<div style="text-align:center;margin-top:2px;font-weight:900">À COLLER À UN ENDROIT VISIBLE</div>' +
    '<div class="muted" style="text-align:center">Dossier #'+esc(numero)+' · '+esc(typeEnvoi||'')+' · '+esc(today)+'</div>' +
    '<div class="box">' +
    '<div class="row"><span class="big">#'+esc(numero)+'</span></div>' +
    '<div class="row"><div class="lbl">EXPÉDITEUR (OPTIMEA)</div><div class="addr">'+esc(savReturnAddressText_())+'</div></div>' +
    '<div class="row"><div class="lbl">SOLUTION (choisie)</div><div class="addr">'+esc(String(typeEnvoi||'').trim() || '—')+'</div></div>' +
    '<div class="row"><div class="lbl">DESTINATAIRE (CLIENT)</div><div class="addr">'+
      esc(String(infoMain.clientNom||infoMain.magasin||'').trim())
      + (String(infoMain.clientNom||infoMain.magasin||'').trim() ? '\n' : '')
      + esc(toAddr) +
    '</div></div>' +
    '<div class="row"><div class="lbl">Adresse (client / dossier)</div><div class="addr">'+esc(toAddr)+'</div></div>' +
    '</div>' +
    '</body></html>';

  const fileName = "Logistique_Etiquette_Expedition_" + numero + ".pdf";
  // Avoid duplicates: if the expected file already exists, reuse it.
  try {
    const it = docs.getFilesByName(fileName);
    if (it && it.hasNext()) {
      const f = it.next();
      return { ok: true, folderUrl, fileUrl: f.getUrl(), fileName, fileId: f.getId(), already: true };
    }
  } catch (e) {
    // ignore and regenerate
  }
  const f = savCreatePdfFile_(docs, fileName, html);
  return { ok: true, folderUrl, fileUrl: f.getUrl(), fileName, fileId: f.getId() };
}

function savBackfillMissingLogistiqueShippingLabels(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const dryRun = !!o.dryRun;
  const limit = Math.max(1, Number(o.limit || 50));

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_LOG);
  if (!sh) return { ok: false, message: "Feuille Logistique introuvable." };
  ensureHeaderGeneric_(sh, HDR_LOG);

  const lr = sh.getLastRow();
  if (lr < 2) return { ok: true, totalRows: 0, processed: 0, created: 0, already: 0, skipped: 0, dryRun };

  const nc = HDR_LOG.length;
  const vals = sh.getRange(2, 1, lr - 1, nc).getValues();

  let created = 0;
  let already = 0;
  let skipped = 0;
  let processed = 0;

  for (let i = 0; i < vals.length; i++) {
    if (processed >= limit) break;
    const r = vals[i] || [];

    const stat = String(r[8] || "").trim();
    // On régénère même si le dossier est "plus loin" (ex: Livré),
    // mais on ignore seulement les lignes annulées/vides.
    if (!stat) continue;
    if (stat === LOG_STAT_ANNULE) continue;

    const numero = String(r[1] || "").trim();
    const sheetName = String(r[2] || "").trim();
    const rowMain = Number(r[3] || 0);
    const typeEnvoi = String(r[4] || "").trim();
    const adresseArrivee = String(r[19] || "").trim();

    if (!numero || !sheetName || !rowMain || rowMain < 2) {
      skipped++;
      continue;
    }
    if (!adresseArrivee) {
      skipped++;
      continue;
    }

    processed++;
    if (dryRun) {
      created++;
      continue;
    }

    const res = savGenerateLogistiqueShippingLabelPdf_({
      sheet: sheetName,
      rowMain: rowMain,
      numero: numero,
      typeEnvoi: typeEnvoi,
      adresseArrivee: adresseArrivee,
    });

    if (res && res.ok) {
      if (res.already) already++;
      else created++;
    } else {
      skipped++;
    }

    if ((processed % 10) === 0) Utilities.sleep(150);
  }

  savCacheInvalidateAll_();
  return { ok: true, dryRun, totalRows: vals.length, processed, created, already, skipped };
}

function savGenerateCaseRecapPdf(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const sheetName = String(o.sheet || "").trim();
  const row = Number(o.row || 0);
  const numero = String(o.numero || "").trim();
  if (!sheetName || !row || row < 2 || !numero) return { ok: false, message: "Paramètres invalides (dossier)." };

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const shMain = ss.getSheetByName(sheetName);
  if (!shMain) return { ok: false, message: "Feuille dossier introuvable." };
  ensureHeaderGeneric_(shMain, sheetName === SAV_SHEET_DIST ? HDR_DIST : HDR_MP);

  const mainWidth = sheetName === SAV_SHEET_DIST ? HDR_DIST.length : HDR_MP.length;
  const mainRow = shMain.getRange(row, 1, 1, mainWidth).getValues()[0] || [];
  const idxMain = savHeaderIndexByName_(shMain);
  const etatCol = idxMain["État du dossier"] || (sheetName === SAV_SHEET_DIST ? 10 : 10);
  const dateCol = idxMain["Date création"] || (sheetName === SAV_SHEET_DIST ? 9 : 10);

  const etat = String(mainRow[etatCol - 1] || "").trim();
  const dateCreation = mainRow[dateCol - 1];
  const dateCreationStr = formatDate_(dateCreation);

  const info = { numero, sheetName, row, etat, dateCreationStr };

  if (sheetName === SAV_SHEET_DIST) {
    info.type = "Distributeur";
    info.magasin = String(mainRow[(idxMain["Nom magasin/distributeur"] || 2) - 1] || "").trim();
    info.adresse = String(mainRow[(idxMain["Adresse complète"] || 3) - 1] || "").trim();
    info.modele = String(mainRow[(idxMain["Modèle concerné"] || 4) - 1] || "").trim();
    info.panne = String(mainRow[(idxMain["Panne constatée"] || 5) - 1] || "").trim();
    info.serie = String(mainRow[(idxMain["Numéro de série"] || 6) - 1] || "").trim();
    info.facture = String(mainRow[(idxMain["Facture reçue"] || 7) - 1] || "").trim();
    info.photo = String(mainRow[(idxMain["Photo plaque reçue"] || 8) - 1] || "").trim();
  } else {
    info.type = "Marketplace";
    info.marketplace = String(mainRow[(idxMain["Nom marketplace"] || 2) - 1] || "").trim();
    info.clientNom = String(mainRow[(idxMain["Nom client"] || 3) - 1] || "").trim();
    info.clientEmail = String(mainRow[(idxMain["Email client"] || 4) - 1] || "").trim();
    info.clientTel = String(mainRow[(idxMain["Téléphone client"] || 5) - 1] || "").trim();
    info.clientAdresse = String(mainRow[(idxMain["Adresse client"] || 6) - 1] || "").trim();
    info.demandeEnlevement = String(mainRow[(idxMain["Demande d'enlèvement"] || 7) - 1] || "").trim();
    info.modele = String(mainRow[(idxMain["Modèle concerné"] || 8) - 1] || "").trim();
    info.panne = String(mainRow[(idxMain["Panne constatée"] || 9) - 1] || "").trim();
    info.numFacture = String(mainRow[(idxMain["Numéro facture"] || 10) - 1] || "").trim();
    info.photoNs = String(mainRow[(idxMain["Photo N°S reçue"] || 11) - 1] || "").trim();
  }

  const { folder, folderUrl } = savEnsureDriveFolderAndWriteUrl_(sheetName, row, numero);

  // Read related steps
  const rec = savReadReceptionByNumero_(ss, numero);
  const exp = savReadExpertiseByNumero_(ss, numero);
  const trs = savReadTransportsByNumero_(ss, numero);
  const logs = savReadLogistiqueByNumero_(ss, numero);
  const mpMsgs = sheetName === SAV_SHEET_MP ? savReadMarketplaceMessagesByNumero_(ss, numero) : [];

  const esc = savHtmlEsc_;

  function rowLine(label, value) {
    return '<div class="row"><span class="lbl">' + esc(label) + "</span> " + esc(value) + "</div>";
  }

  function section(title, bodyHtml) {
    return '<div class="sec"><div class="sectitle">' + esc(title) + "</div>" + (bodyHtml || '<div class="muted">—</div>') + "</div>";
  }

  function listBlock(arr) {
    if (!arr || !arr.length) return '<div class="muted">—</div>';
    let h = '<ul style="margin:6px 0 0 18px;padding:0">';
    arr.forEach((x) => (h += "<li>" + esc(x) + "</li>"));
    h += "</ul>";
    return h;
  }

  function table(headers, rows) {
    if (!rows || !rows.length) return '<div class="muted">—</div>';
    let h = '<table class="t"><thead><tr>';
    headers.forEach((x) => (h += "<th>" + esc(x) + "</th>"));
    h += "</tr></thead><tbody>";
    rows.forEach((r) => {
      h += "<tr>";
      r.forEach((c) => (h += "<td>" + esc(c) + "</td>"));
      h += "</tr>";
    });
    h += "</tbody></table>";
    return h;
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  let html = "";
  html += '<html><head><meta charset="utf-8"/>';
  html += "<style>";
  html += 'body{font-family:Arial,sans-serif;font-size:12px;color:#0f172a;margin:24px}';
  html += 'h1{margin:0 0 6px 0;font-size:16px}';
  html += ".muted{color:#475569}";
  html += ".sec{margin-top:12px;border-top:1px solid #cbd5e1;padding-top:10px}";
  html += ".sectitle{font-weight:900;margin-bottom:6px}";
  html += ".row{margin:4px 0}";
  html += ".lbl{display:inline-block;width:170px;font-weight:bold}";
  html += ".t{width:100%;border-collapse:collapse;margin-top:6px}";
  html += ".t th,.t td{border:1px solid #cbd5e1;padding:6px;vertical-align:top}";
  html += ".t th{background:#f1f5f9;text-align:left}";
  html += "</style></head><body>";

  html += "<h1>RAPPORT TECHNIQUE — DOSSIER SAV " + esc(numero) + "</h1>";
  html += '<div class="muted">Généré le ' + esc(today) + " · Dossier Drive : " + esc(folderUrl) + "</div>";

  // Dossier (nouvelle demande)
  let dossierHtml = "";
  dossierHtml += rowLine("Type", info.type || "");
  dossierHtml += rowLine("État", info.etat || "");
  dossierHtml += rowLine("Date création", info.dateCreationStr || "");
  if (info.type === "Distributeur") {
    dossierHtml += rowLine("Magasin / Distributeur", info.magasin || "");
    dossierHtml += rowLine("Adresse", info.adresse || "");
    dossierHtml += rowLine("Modèle", info.modele || "");
    dossierHtml += rowLine("Numéro de série", info.serie || "");
    dossierHtml += rowLine("Panne constatée", info.panne || "");
    dossierHtml += rowLine("Facture reçue", info.facture || "");
    dossierHtml += rowLine("Photo plaque reçue", info.photo || "");
  } else {
    dossierHtml += rowLine("Marketplace", info.marketplace || "");
    dossierHtml += rowLine("Nom client", info.clientNom || "");
    dossierHtml += rowLine("Email client", info.clientEmail || "");
    dossierHtml += rowLine("Téléphone client", info.clientTel || "");
    dossierHtml += rowLine("Adresse client", info.clientAdresse || "");
    dossierHtml += rowLine("Demande d'enlèvement", info.demandeEnlevement || "");
    dossierHtml += rowLine("Modèle", info.modele || "");
    dossierHtml += rowLine("Panne constatée", info.panne || "");
    dossierHtml += rowLine("Numéro facture", info.numFacture || "");
    dossierHtml += rowLine("Photo N°S reçue", info.photoNs || "");
  }
  html += section("Nouvelle demande", dossierHtml);

  // Réception atelier
  const recRows = (rec || []).map((x) => [
    String(x.dateReception || ""),
    String(x.etatVisuel || ""),
    String(x.elements || ""),
    String(x.notes || ""),
    String(x.technicien || ""),
  ]);
  html += section("Réception atelier", table(["Date", "État visuel", "Éléments", "Notes", "Technicien"], recRows));

  // Expertise
  const expRows = (exp || []).map((x) => [
    String(x.dateExpertise || ""),
    String(x.technicien || ""),
    String(x.travaux || ""),
    String(x.pieces || ""),
    String(x.tests || ""),
    String(x.conclusions || ""),
    String(x.decision || ""),
    String(x.coutPieces || ""),
    String(x.coutMainOeuvre || ""),
  ]);
  html += section(
    "Expertise",
    table(
      ["Date", "Technicien", "Travaux", "Pièces", "Tests", "Conclusions", "Décision", "Coût pièces", "Coût main d'œuvre"],
      expRows
    )
  );

  // Envoi / transport
  const trRows = (trs || []).map((x) => [
    String(x.dateDemande || ""),
    String(x.typeEnvoi || ""),
    String(x.pointDepart || ""),
    String(x.adresseArrivee || ""),
    String(x.tracking || ""),
    String(x.statut || ""),
    String(x.preuve || ""),
    String(x.dateLivraison || ""),
  ]);
  html += section(
    "Envoi / transport",
    table(["Date demande", "Type", "Adresse départ", "Adresse arrivée", "Tracking", "Statut", "Preuve", "Date livraison"], trRows)
  );

  // Logistique
  const logRows = (logs || []).map((x) => [
    String(x.dateDemande || ""),
    String(x.typeEnvoi || ""),
    String(x.pointDepart || ""),
    String(x.statutLog || ""),
    String(x.tracking || ""),
    String(x.transporteur || ""),
    String(x.coutTransport || ""),
    String(x.dateExpedition || ""),
    String(x.preuveLivraison || ""),
    String(x.dateLivraison || ""),
  ]);
  html += section(
    "Logistique",
    table(
      ["Date demande", "Type", "Départ", "Statut", "Tracking", "Transporteur", "Coût", "Date expédition", "Preuve", "Date livraison"],
      logRows
    )
  );

  if (mpMsgs && mpMsgs.length) {
    const msgRows = mpMsgs.map((m) => [
      String(m.date || ""),
      String(m.sens || ""),
      String(m.auteur || ""),
      String(m.message || ""),
    ]);
    html += section("Messages Marketplace", table(["Date", "Sens", "Auteur", "Message"], msgRows));
  }

  html += "</body></html>";

  const blob = Utilities.newBlob(html, "text/html", "rapport.html").getAs("application/pdf");
  const fileName = "Rapport_Technique_" + numero + ".pdf";
  const f = folder.createFile(blob.setName(fileName));
  return { ok: true, folderUrl, fileUrl: f.getUrl(), fileName };
}

function savGetCaseTimelineHtml(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const sheetName = String(o.sheet || "").trim();
  const row = Number(o.row || 0);
  const numero = String(o.numero || "").trim();
  if (!sheetName || !row || row < 2 || !numero) return { ok: false, message: "Paramètres invalides (dossier)." };

  try {
    const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
    const shMain = ss.getSheetByName(sheetName);
    if (!shMain) return { ok: false, message: "Feuille dossier introuvable." };
    ensureHeaderGeneric_(shMain, sheetName === SAV_SHEET_DIST ? HDR_DIST : HDR_MP);
    const idxMain = savHeaderIndexByName_(shMain);
    const mainWidth = sheetName === SAV_SHEET_DIST ? HDR_DIST.length : HDR_MP.length;
    const mainRow = shMain.getRange(row, 1, 1, mainWidth).getValues()[0] || [];

    const dateCol = idxMain["Date création"] || (sheetName === SAV_SHEET_DIST ? 9 : 12);
    const etatCol = idxMain["État du dossier"] || (sheetName === SAV_SHEET_DIST ? 10 : 13);
    const closeCol = idxMain["Date clôture"] || (sheetName === SAV_SHEET_DIST ? HDR_DIST.length : HDR_MP.length);

    const created = mainRow[dateCol - 1];
    const etat = String(mainRow[etatCol - 1] || "").trim();
    const closed = mainRow[closeCol - 1];

    const esc = savHtmlEsc_;

    function row_(d, step, details, urlLabel, url) {
      const dateStr = d instanceof Date && !isNaN(d.getTime()) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : String(d || "");
      const link = url ? ('<a href="' + esc(url) + '" target="_blank">' + esc(urlLabel || "Ouvrir") + "</a>") : "";
      return (
        "<tr>" +
        "<td style='white-space:nowrap'>" + esc(dateStr || "—") + "</td>" +
        "<td style='font-weight:900'>" + esc(step || "") + "</td>" +
        "<td style='white-space:pre-wrap'>" + esc(details || "") + "</td>" +
        "<td>" + link + "</td>" +
        "</tr>"
      );
    }

    const { folderUrl } = savEnsureDriveFolderAndWriteUrl_(sheetName, row, numero);
    // Liens Drive utiles
    let proofUrl = "";
    let msgUrl = "";
    try {
      const parent = savGetDriveParentFolderForSheet_(sheetName);
        const dossier = savGetOrCreateSubFolder_(parent, numero);
        proofUrl = savGetOrCreateSubFolder_(dossier, "Preuves livraison").getUrl();
        msgUrl = savGetOrCreateSubFolder_(dossier, "Messages Marketplace").getUrl();
    } catch (e) {}

    const rec = savReadReceptionByNumero_(ss, numero) || [];
    const exp = savReadExpertiseByNumero_(ss, numero) || [];
    const trs = savReadTransportsByNumero_(ss, numero) || [];
    const logs = savReadLogistiqueByNumero_(ss, numero) || [];

    // Tickets avoir : lecture rapide par numéro (col 1)
    let avoirRows = [];
    try {
      const shA = ss.getSheetByName(SAV_SHEET_AVOIR_TICKETS);
      if (shA && shA.getLastRow() >= 2) {
        ensureHeaderGeneric_(shA, HDR_AVOIR_TICKETS);
        const lrA = shA.getLastRow();
        const valsA = shA.getRange(2, 1, lrA - 1, HDR_AVOIR_TICKETS.length).getValues();
        for (let i = 0; i < valsA.length; i++) {
          const r = valsA[i] || [];
          if (String(r[0] || "").trim() !== numero) continue;
          avoirRows.push({
            date: r[7],
            statut: String(r[8] || "").trim(),
            centrale: String(r[10] || "").trim(),
            facture: String(r[11] || "").trim(),
          });
        }
      }
    } catch (e) {}

    let html = "";
    html += "<div style='display:flex;gap:10px;flex-wrap:wrap;align-items:center;justify-content:space-between;margin-bottom:10px'>";
    html += "<div><div style='font-weight:950'>Historique dossier</div><div class='muted'>Dossier " + esc(numero) + " · État : " + esc(etat) + "</div></div>";
    html += "<div class='row' style='gap:10px;flex-wrap:wrap'>";
    html += folderUrl ? ("<a class='btn' href='" + esc(folderUrl) + "' target='_blank'>Ouvrir dossier Drive</a>") : "";
    html += proofUrl ? ("<a class='btn' href='" + esc(proofUrl) + "' target='_blank'>Preuves livraison</a>") : "";
    if (sheetName === SAV_SHEET_MP && msgUrl) html += "<a class='btn' href='" + esc(msgUrl) + "' target='_blank'>Messages Marketplace</a>";
    html += "</div></div>";

    html += "<div style='overflow:auto;border:1px solid var(--border);border-radius:12px;background:var(--card)'>";
    html += "<table style='width:100%;border-collapse:collapse;font-size:12px'>";
    html += "<thead><tr style='background:rgba(15,23,42,.04)'><th style='text-align:left;padding:10px'>Date</th><th style='text-align:left;padding:10px'>Étape</th><th style='text-align:left;padding:10px'>Détails</th><th style='text-align:left;padding:10px'>Lien</th></tr></thead><tbody>";

    html += row_(created, "Création dossier", "", "Drive", folderUrl);
    if (closed instanceof Date && !isNaN(closed.getTime())) html += row_(closed, "Clôture dossier", "", "", "");

    rec.forEach((x) => {
      html += row_(x.dateReception, "Réception atelier", [x.etatVisuel, x.elements, x.notes].filter(Boolean).join("\n"), "", "");
    });
    exp.forEach((x) => {
      html += row_(x.dateExpertise, "Expertise", ["Technicien: " + (x.technicien || ""), "Décision: " + (x.decision || ""), x.conclusions || ""].filter(Boolean).join("\n"), "", "");
    });
    trs.forEach((x) => {
      html += row_(x.dateDemande, "Envoi / transport (" + (x.typeEnvoi || "") + ")", ["Statut: " + (x.statut || ""), "Tracking: " + (x.tracking || ""), "Arrivée: " + (x.adresseArrivee || "")].filter(Boolean).join("\n"), "", "");
    });
    logs.forEach((x) => {
      html += row_(x.dateDemande, "Logistique (" + (x.typeEnvoi || "") + ")", ["Statut: " + (x.statutLog || ""), "Transporteur: " + (x.transporteur || ""), "Tracking: " + (x.tracking || "")].filter(Boolean).join("\n"), "Preuves", proofUrl);
      if (x.dateExpedition) html += row_(x.dateExpedition, "Expédition", "", "", "");
      if (x.dateLivraison) html += row_(x.dateLivraison, "Livraison", "", "Preuves", proofUrl);
    });
    avoirRows.forEach((x) => {
      html += row_(x.date, "Ticket avoir", ["Statut: " + (x.statut || ""), "Centrale: " + (x.centrale || ""), x.facture ? ("Facture: " + x.facture) : ""].filter(Boolean).join("\n"), "", "");
    });

    html += "</tbody></table></div>";
    return { ok: true, html };
  } catch (e) {
    return { ok: false, message: "Erreur génération historique." };
  }
}

function savReadReceptionByNumero_(ss, numero) {
  try {
    const sh = ss.getSheetByName(SAV_SHEET_REC);
    if (!sh || sh.getLastRow() < 2) return [];
    ensureHeaderGeneric_(sh, HDR_REC);
    const lr = sh.getLastRow();
    const vals = sh.getRange(2, 1, lr - 1, HDR_REC.length).getValues();
    const out = [];
    const n = String(numero || "").trim();
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      if (String(r[0] || "").trim() !== n) continue;
      out.push({
        dateReception: formatDateIfDate_(r[1]),
        etatVisuel: String(r[2] || ""),
        elements: String(r[3] || ""),
        notes: String(r[4] || ""),
        technicien: String(r[5] || ""),
      });
    }
    return out;
  } catch (e) {
    return [];
  }
}

function savReadExpertiseByNumero_(ss, numero) {
  try {
    const sh = ss.getSheetByName(SAV_SHEET_EXP);
    if (!sh || sh.getLastRow() < 2) return [];
    ensureHeaderGeneric_(sh, HDR_EXP);
    const lr = sh.getLastRow();
    const vals = sh.getRange(2, 1, lr - 1, HDR_EXP.length).getValues();
    const out = [];
    const n = String(numero || "").trim();
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i] || [];
      if (String(r[0] || "").trim() !== n) continue;
      out.push({
        dateExpertise: formatDateIfDate_(r[1]),
        technicien: String(r[2] || ""),
        travaux: String(r[3] || ""),
        pieces: String(r[4] || ""),
        tests: String(r[5] || ""),
        conclusions: String(r[6] || ""),
        decision: String(r[7] || ""),
        coutPieces: r[8] === "" || r[8] == null ? "" : String(r[8]),
        coutMainOeuvre: r[9] === "" || r[9] == null ? "" : String(r[9]),
      });
    }
    out.sort((a, b) => String(a.dateExpertise || "").localeCompare(String(b.dateExpertise || "")));
    return out;
  } catch (e) {
    return [];
  }
}

function savReadTransportsByNumero_(ss, numero) {
  try {
    const sh = ss.getSheetByName(SAV_SHEET_TR);
    if (!sh || sh.getLastRow() < 2) return [];
    ensureHeaderGeneric_(sh, HDR_TR);
    const lr = sh.getLastRow();
    const vals = sh.getRange(2, 1, lr - 1, HDR_TR.length).getValues();
    const out = [];
    const n = String(numero || "").trim();
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      if (String(r[0] || "").trim() !== n) continue;
      out.push({
        typeEnvoi: String(r[1] || ""),
        pointDepart: String(r[2] || ""),
        dateDemande: formatDateIfDate_(r[3]),
        tracking: String(r[4] || ""),
        statut: String(r[5] || ""),
        preuve: String(r[6] || ""),
        dateLivraison: formatDateIfDate_(r[7]),
        adresseArrivee: String(r[8] || ""),
      });
    }
    out.sort((a, b) => String(a.dateDemande || "").localeCompare(String(b.dateDemande || "")));
    return out;
  } catch (e) {
    return [];
  }
}

function savReadLogistiqueByNumero_(ss, numero) {
  try {
    const sh = ss.getSheetByName(SAV_SHEET_LOG);
    if (!sh || sh.getLastRow() < 2) return [];
    ensureHeaderGeneric_(sh, HDR_LOG);
    const lr = sh.getLastRow();
    const vals = sh.getRange(2, 1, lr - 1, HDR_LOG.length).getValues();
    const out = [];
    const n = String(numero || "").trim();
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      if (String(r[1] || "").trim() !== n) continue;
      out.push({
        typeEnvoi: String(r[4] || ""),
        pointDepart: String(r[5] || ""),
        statutLog: String(r[8] || ""),
        dateDemande: formatDate_(r[9]),
        tracking: String(r[10] || ""),
        dateExpedition: formatDateIfDate_(r[11]),
        transporteur: String(r[12] || ""),
        coutTransport: Number(r[13] || 0),
        preuveLivraison: String(r[15] || ""),
        dateLivraison: formatDateIfDate_(r[16]),
      });
    }
    out.sort((a, b) => String(a.dateDemande || "").localeCompare(String(b.dateDemande || "")));
    return out;
  } catch (e) {
    return [];
  }
}

function savReadMarketplaceMessagesByNumero_(ss, numero) {
  try {
    const sh = ss.getSheetByName(SAV_SHEET_MP_MESSAGES);
    if (!sh || sh.getLastRow() < 2) return [];
    ensureHeaderGeneric_(sh, HDR_MP_MESSAGES);
    const lr = sh.getLastRow();
    const vals = sh.getRange(2, 1, lr - 1, HDR_MP_MESSAGES.length).getValues();
    const out = [];
    const n = String(numero || "").trim();
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      if (String(r[1] || "").trim() !== n) continue;
      out.push({
        sens: String(r[4] || ""),
        message: String(r[5] || ""),
        auteur: String(r[6] || ""),
        date: formatDate_(r[7]),
      });
    }
    out.sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")));
    return out;
  } catch (e) {
    return [];
  }
}

function savUpdateEtat(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const sheet = String(o.sheet || "");
  const row = Number(o.row);
  const etat = String(o.etat || "");
  if (!sheet || !row || row < 2 || !etat) return { ok: false, message: "Paramètres invalides." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheet);

  // Bloque la progression si docs requis manquants (sauf règles spécifiques logistique gérées ailleurs)
  savAssertDocsBeforeNextStep_({ sheet, row, etat, sh });

  const idx = savHeaderIndexByName_(sh);
  const col = idx["État du dossier"] || (sheet === SAV_SHEET_DIST ? 10 : 9);
  sh.getRange(row, col).setValue(etat);
  // Si clôture : enregistrer date de clôture (si vide)
  if (savIsFinalEtat_(etat)) {
    const closeCol = idx["Date clôture"] || (sheet === SAV_SHEET_DIST ? HDR_DIST.length : HDR_MP.length); // dernière colonne = Date clôture
    const existing = sh.getRange(row, closeCol).getValue();
    if (!(existing instanceof Date) || isNaN(existing.getTime())) {
      sh.getRange(row, closeCol).setValue(new Date());
    }
  }
  savCacheInvalidateAll_();
  return { ok: true };
}

function savSetMailAccordRetour(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const sheet = String(o.sheet || "").trim();
  const row = Number(o.row || 0);
  const sent = !!o.sent;
  const emailTo = String(o.emailTo || "").trim();
  if (!sheet || !row || row < 2) return { ok: false, message: "Paramètres invalides." };
  if (sheet !== SAV_SHEET_DIST) return { ok: false, message: "Fonction disponible uniquement pour les dossiers Distributeur." };

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheet);
  if (!sh) return { ok: false, message: "Feuille introuvable." };
  ensureHeaderGeneric_(sh, HDR_DIST);
  const idx = savHeaderIndexByName_(sh);
  const colSent = idx["Mail envoyé au distributeur (accord retour atelier)"] || 0;
  const colDate = idx["Date mail accord retour atelier"] || 0;
  if (!colSent || !colDate) return { ok: false, message: "Colonnes mail accord retour introuvables." };

  sh.getRange(row, colSent).setValue(sent ? "OUI" : "NON");
  if (sent) {
    const existing = sh.getRange(row, colDate).getValue();
    if (!(existing instanceof Date) || isNaN(existing.getTime())) sh.getRange(row, colDate).setValue(new Date());
    // passage automatique à l'étape "En attente réception"
    savUpdateEtat_({ sheet: sheet, row: row, etat: ETAT_EN_ATTENTE });

    // Génère l'étiquette retour (PROC 2 & 3) au moment de "l'envoi"
    try {
      const width = HDR_DIST.length;
      const mainRow = sh.getRange(row, 1, 1, width).getValues()[0] || [];
      const idx2 = savHeaderIndexByName_(sh);
      const numero = String(mainRow[0] || "").trim();
      const modele = String(mainRow[(idx2["Modèle concerné"] || 4) - 1] || "").trim();
      const pid = savGetProcIdForModel_(modele);
      if (pid === SAV_PROC_RAPATRIEMENT || pid === SAV_PROC_VISIO) {
        const lab = savGenerateReturnLabelPdf({ sheet: sheet, row: row, numero: numero });
        // Envoi email optionnel (si adresse fournie)
        if (emailTo && lab && lab.ok && lab.fileId) {
          const file = DriveApp.getFileById(lab.fileId);
          const subj = "SAV OPTIMEA — Étiquette de retour à coller — Dossier #" + numero;
          const body =
            "Bonjour,\n\n" +
            "Veuillez trouver en pièce jointe l’étiquette de retour pour le dossier SAV #" +
            numero +
            ".\n" +
            "Merci de l’imprimer et de la COLLER À UN ENDROIT VISIBLE sur le colis.\n\n" +
            "Adresse de retour :\n" +
            "OPTIMEA - SAV\n34 Rue du Moulin des Bruyères\n92400 COURBEVOIE\nFRANCE\n\n" +
            "Cordialement,\nSAV OPTIMEA";
          const sendRes = savSendEmailFromSav_(emailTo, subj, body, [file.getAs(MimeType.PDF)]);
          if (!sendRes || !sendRes.ok) throw new Error((sendRes && sendRes.message) || "Erreur envoi email.");
        }
      }
    } catch (e) {
      // ne bloque pas le passage d'étape
    }
  }
  savCacheInvalidateAll_();
  return { ok: true };
}

function savAssertDocsBeforeNextStep_(ctx) {
  const sheet = String((ctx && ctx.sheet) || "").trim();
  const row = Number((ctx && ctx.row) || 0);
  const etat = String((ctx && ctx.etat) || "").trim();
  const sh = (ctx && ctx.sh) || null;
  if (!sheet || !row || row < 2 || !etat || !sh) return;

  // On bloque uniquement quand on quitte les étapes "amont"
  const etatsRequiringDocs = [ETAT_RECU, ETAT_EXPERTISE, ETAT_TRANSPORT_DEMANDE, ETAT_TRANSPORT_ROUTE, ETAT_TRANSPORT_LIVRE, ETAT_CLOTURE_AVOIR];
  if (etatsRequiringDocs.indexOf(etat) === -1) return;

  if (sheet === SAV_SHEET_DIST) {
    // Colonnes: Facture reçue (7), Photo plaque (8)
    const r = sh.getRange(row, 1, 1, HDR_DIST.length).getValues()[0];
    const facture = String(r[6] || "").trim(); // OUI/NON
    const photoPlaque = String(r[7] || "").trim(); // OUI/NON
    const missing = [];
    if (facture !== "OUI") missing.push("Facture reçue");
    if (photoPlaque !== "OUI") missing.push("Photo plaque reçue");
    if (missing.length) throw new Error("Impossible de passer à l'étape suivante : document(s) manquant(s) — " + missing.join(", "));
  } else if (sheet === SAV_SHEET_MP) {
    const idx = savHeaderIndexByName_(sh);
    const r = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    const numFacture = String(r[(idx["Numéro facture"] || 6) - 1] || "").trim();
    const photoNs = String(r[(idx["Photo N°S reçue"] || 7) - 1] || "").trim(); // OUI/NON
    const missing = [];
    if (!numFacture) missing.push("N° facture");
    if (photoNs !== "OUI") missing.push("Photo N°S reçue");
    if (missing.length) throw new Error("Impossible de passer à l'étape suivante : document(s) manquant(s) — " + missing.join(", "));
  }
}

function savIsFinalEtat_(etat) {
  const e = String(etat || "").trim();
  return e === ETAT_CLOTURE_AVOIR || e === ETAT_FERMÉ_HG || e === ETAT_ARCHIVE;
}

function savAddReception(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_REC);
  const row = [
    String(o.numero || ""),
    o.dateReception instanceof Date ? o.dateReception : new Date(String(o.dateReception || "")),
    String(o.etatVisuel || ""),
    String(o.elements || ""),
    String(o.notes || ""),
    String(o.technicien || "").trim(),
  ];
  sh.appendRow(row);
  if (o.sheet && o.rowMain) {
    savUpdateEtat_({ sheet: o.sheet, row: o.rowMain, etat: ETAT_RECU });
  }
  return { ok: true };
}

function savSaveExpertise(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_EXP);
  const dateExpertise = o.dateExpertise instanceof Date ? o.dateExpertise : o.dateExpertise ? new Date(String(o.dateExpertise)) : new Date();
  sh.appendRow([
    String(o.numero || ""),
    dateExpertise,
    String(o.technicien || "").trim(),
    String(o.travaux || ""),
    String(o.pieces || ""),
    String(o.tests || ""),
    String(o.conclusions || ""),
    String(o.decision || ""),
    o.coutPieces === "" || o.coutPieces == null ? "" : Number(o.coutPieces),
    o.coutMainOeuvre === "" || o.coutMainOeuvre == null ? "" : Number(o.coutMainOeuvre),
  ]);
  savCacheInvalidateAll_();
  return { ok: true };
}

function savGetExpertiseLastForNumero(numero) {
  ensureSavSheets_();
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const rows = savReadExpertiseByNumero_(ss, numero) || [];
  return rows.length ? rows[rows.length - 1] : null;
}

function savUpdateEtat_(o) {
  savUpdateEtat(o);
}

// Module "Avoir" supprimé à la demande (création avoir).

function savAddTransport(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_TR);
  const now = new Date();

  let lastLabel = null;
  function applySideEffects_() {
    if (o.sheet && o.rowMain && o.updateEtat) {
      savUpdateEtat_({ sheet: o.sheet, row: o.rowMain, etat: o.updateEtat });
    }
    /** File logistique : une ligne « bon » par demande transport SAV. */
    if (o.sheet && o.rowMain && o.updateEtat === ETAT_TRANSPORT_DEMANDE) {
      logistiqueQueue_(o);

      // Générer une étiquette d’expédition (si adresse arrivée fournie)
      // et l’envoyer à la logistique si souhaité (par défaut ADV).
      try {
        const lab = savGenerateLogistiqueShippingLabelPdf_(o);
        if (lab && lab.ok && lab.fileId) {
          lastLabel = lab;
          const file = DriveApp.getFileById(lab.fileId);
          const subj = "Logistique — étiquette expédition à coller — Dossier #" + String(o.numero || "");
          const body =
            "Bonjour,\n\n" +
            "Merci de traiter la demande logistique.\n" +
            "Étiquette d’expédition jointe : À COLLER À UN ENDROIT VISIBLE.\n\n" +
            "Dossier : #" +
            String(o.numero || "") +
            "\n" +
            "Type : " +
            String(o.typeEnvoi || "") +
            "\n" +
            "Destinataire :\n" +
            String(o.adresseArrivee || "") +
            "\n\n" +
            "Expéditeur :\n" +
            savReturnAddressText_() +
            "\n\n" +
            "Lien dossier Drive : " +
            String(lab.folderUrl || "") +
            "\n";

          // Envoi à ADV/logistique (modifiable via payload.emailTo)
          const emailTo = String(o.emailTo || SAV_MAIL_ADV || "").trim();
          if (emailTo) {
            try {
              savSendEmailFromSav_(emailTo, subj, body, [file.getAs(MimeType.PDF)]);
            } catch (eEmail) {}
          }
        }
      } catch (e) {}
    }
  }

  // Déduplication forte : si une demande identique existe déjà, ne jamais ajouter une 2e ligne
  if (savHasSameOpenTransportRequest_(sh, o)) {
    applySideEffects_();
    return { ok: true, deduped: true, dedupedReason: "existing", label: lastLabel };
  }
  // Déduplication: évite double-clic / double appel qui crée 2 lignes
  if (savIsDuplicateTransportRequest_(sh, o, now)) {
    // On applique quand même les effets de bord (idempotent)
    applySideEffects_();
    return { ok: true, deduped: true, label: lastLabel };
  }
  sh.appendRow([
    String(o.numero || ""),
    String(o.typeEnvoi || ""),
    String(o.pointDepart || ""),
    now,
    String(o.tracking || ""),
    String(o.statut || "Demande"),
    o.preuve ? "OUI" : "NON",
    o.dateLivraison ? (o.dateLivraison instanceof Date ? o.dateLivraison : new Date(String(o.dateLivraison))) : "",
    String(o.adresseArrivee || ""),
  ]);
  applySideEffects_();
  return { ok: true, label: lastLabel };
}

function savHasSameOpenTransportRequest_(sh, o) {
  try {
    function norm_(s) {
      return String(s || "")
        .replace(/\s+/g, " ")
        .trim()
        .toLowerCase();
    }
    const n = String(o && o.numero ? o.numero : "").trim();
    if (!n) return false;
    const typeEnvoi = norm_(o && o.typeEnvoi ? o.typeEnvoi : "");
    const pointDepart = norm_(o && o.pointDepart ? o.pointDepart : "");
    const adresseArrivee = norm_(o && o.adresseArrivee ? o.adresseArrivee : "");
    const statut = norm_(o && o.statut ? o.statut : "Demande");
    const tracking = norm_(o && o.tracking ? o.tracking : "");

    const lr = sh.getLastRow();
    if (lr < 2) return false;
    const from = Math.max(2, lr - 200); // plus large que 20
    const vals = sh.getRange(from, 1, lr - from + 1, HDR_TR.length).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      const r = vals[i] || [];
      const rn = String(r[0] || "").trim();
      if (rn !== n) continue;
      const rType = norm_(r[1]);
      const rPt = norm_(r[2]);
      const rTrk = norm_(r[4]);
      const rStat = norm_(r[5]);
      const rArr = norm_(r[8]);
      if (rType !== typeEnvoi) continue;
      if (rPt !== pointDepart) continue;
      if (rArr !== adresseArrivee) continue;
      if (rStat !== statut) continue;
      if (rTrk !== tracking) continue;
      return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}

function savIsDuplicateTransportRequest_(sh, o, now) {
  try {
    function norm_(s){
      return String(s||'')
        .replace(/\s+/g,' ')
        .trim()
        .toLowerCase();
    }
    const n = String(o && o.numero ? o.numero : "").trim();
    if (!n) return false;
    const typeEnvoi = norm_(o && o.typeEnvoi ? o.typeEnvoi : "");
    const pointDepart = norm_(o && o.pointDepart ? o.pointDepart : "");
    const adresseArrivee = norm_(o && o.adresseArrivee ? o.adresseArrivee : "");
    const statut = norm_(o && o.statut ? o.statut : "Demande");
    const tracking = norm_(o && o.tracking ? o.tracking : "");

    const lr = sh.getLastRow();
    if (lr < 2) return false;
    const from = Math.max(2, lr - 20); // check last 20 rows only
    const vals = sh.getRange(from, 1, lr - from + 1, HDR_TR.length).getValues();

    for (let i = vals.length - 1; i >= 0; i--) {
      const r = vals[i] || [];
      const rn = String(r[0] || "").trim();
      if (rn !== n) continue;
      const rType = norm_(r[1]);
      const rPt = norm_(r[2]);
      const rDate = r[3];
      const rTrk = norm_(r[4]);
      const rStat = norm_(r[5]);
      const rArr = norm_(r[8]);
      // same payload "shape"
      if (rType !== typeEnvoi) continue;
      if (rPt !== pointDepart) continue;
      if (rArr !== adresseArrivee) continue;
      if (rStat !== statut) continue;
      if (rTrk !== tracking) continue;
      if (!(rDate instanceof Date) || isNaN(rDate.getTime())) return false;
      // within 2 minutes => likely double submit
      if (Math.abs(now.getTime() - rDate.getTime()) <= 2 * 60 * 1000) return true;
      return false;
    }
    return false;
  } catch (e) {
    return false;
  }
}

function readDossierRow_(sheetName, row) {
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh || row < 2) return null;
  const width = sheetName === SAV_SHEET_DIST ? HDR_DIST.length : HDR_MP.length;
  return sh.getRange(row, 1, 1, width).getValues()[0];
}

function buildResumeForDossier_(sheetName, row) {
  const r = readDossierRow_(sheetName, row);
  if (!r) return "";
  if (sheetName === SAV_SHEET_DIST) {
    return [String(r[1] || "").trim(), String(r[3] || "").trim()].filter(Boolean).join(" · ");
  }
  return [String(r[2] || "").trim(), String(r[1] || "").trim()].filter(Boolean).join(" · ");
}

function readModeleForDossier_(sheetName, row) {
  const r = readDossierRow_(sheetName, row);
  if (!r) return "";
  if (sheetName === SAV_SHEET_DIST) return String(r[3] || "").trim();
  if (sheetName === SAV_SHEET_MP) return String(r[7] || "").trim();
  return "";
}

function logistiqueQueue_(o) {
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_LOG);
  ensureHeaderGeneric_(sh, HDR_LOG);
  // Déduplication logistique: une seule ligne "ouverte" par numéro + typeEnvoi + pointDepart + adresseArrivee
  try {
    const n = String(o && o.numero ? o.numero : "").trim();
    if (n && sh.getLastRow() >= 2) {
      const lr = sh.getLastRow();
      const from = Math.max(2, lr - 50);
      const vals = sh.getRange(from, 1, lr - from + 1, HDR_LOG.length).getValues();
      for (let i = vals.length - 1; i >= 0; i--) {
        const r = vals[i] || [];
        const rn = String(r[1] || "").trim(); // N° dossier
        if (rn !== n) continue;
        const rType = String(r[4] || "").trim();
        const rPt = String(r[5] || "").trim();
        const rStat = String(r[8] || "").trim();
        const rArr = String(r[19] || "").trim();
        if (rStat === LOG_STAT_LIVRE) continue;
        if (
          rType === String(o.typeEnvoi || "").trim() &&
          rPt === String(o.pointDepart || "").trim() &&
          rArr === String(o.adresseArrivee || "").trim()
        ) {
          return; // already queued
        }
      }
    }
  } catch (e) {}
  const resume = buildResumeForDossier_(o.sheet, o.rowMain);
  const modele = readModeleForDossier_(o.sheet, o.rowMain);
  sh.appendRow([
    Utilities.getUuid(),
    String(o.numero || ""),
    String(o.sheet || ""),
    Number(o.rowMain || 0),
    String(o.typeEnvoi || ""),
    String(o.pointDepart || ""),
    resume,
    modele,
    LOG_STAT_A_TRAITER,
    new Date(),
    String(o.tracking || ""),
    "",
    String(o.transporteur || "DISTRIBUTEUR"),
    Number(o.coutTransport || 0),
    "",
    "NON",
    "",
    "NON",
    "",
    String(o.adresseArrivee || ""),
  ]);
}

function savListLogistique(openOnly) {
  ensureSavSheets_();
  if (openOnly) {
    const cached = savCacheGet_(SAV_CACHE_KEYS.LOG_OPEN);
    if (cached && Array.isArray(cached)) return cached;
  }
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_LOG);
  if (!sh || sh.getLastRow() <= 1) return [];
  const lr = sh.getLastRow();
  const nc = HDR_LOG.length;
  const vals = sh.getRange(2, 1, lr - 1, nc).getValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const row = i + 2;
    const r = vals[i];
    const idRaw = String(r[0] || "").trim();
    const numeroRaw = String(r[1] || "").trim();
    const stat = String(r[8] || "").trim();
    if (openOnly && stat === LOG_STAT_LIVRE) continue;
    // Ignore lignes vides / parasites (souvent créées par des mises en forme / anciennes plages)
    if (!idRaw && !numeroRaw && !stat) continue;
    out.push({
      logRow: row,
      id: idRaw,
      numero: numeroRaw,
      feuille: String(r[2] || ""),
      ligneDossier: Number(r[3] || 0),
      typeEnvoi: String(r[4] || ""),
      pointDepart: String(r[5] || ""),
      resume: String(r[6] || ""),
      modele: String(r[7] || ""),
      statutLog: stat,
      dateDemande: formatDate_(r[9]),
      tracking: String(r[10] || ""),
      dateExpedition: formatDateIfDate_(r[11]),
      transporteur: String(r[12] || ""),
      coutTransport: Number(r[13] || 0),
      notesLog: String(r[14] || ""),
      preuveLivraison: String(r[15] || ""),
      dateLivraison: formatDateIfDate_(r[16]),
      colisPret: String(r[17] || ""),
      dateColisPret: formatDateIfDate_(r[18]),
      adresseArrivee: String(r[19] || ""),
      missingTracking: stat === LOG_STAT_EXPEDIE && !String(r[10] || "").trim(),
    });
  }
  out.sort((a, b) => String(b.dateDemande || "").localeCompare(String(a.dateDemande || "")));
  if (openOnly) savCachePut_(SAV_CACHE_KEYS.LOG_OPEN, out);
  return out;
}

function savDeleteLogistique(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const logRow = Number(o.logRow || 0);
  if (!logRow || logRow < 2) return { ok: false, message: "Ligne logistique invalide." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_LOG);
  if (!sh) return { ok: false, message: "Feuille Logistique introuvable." };
  if (sh.getLastRow() < logRow) return { ok: false, message: "Ligne logistique introuvable." };
  sh.deleteRow(logRow);
  savCacheInvalidateAll_();
  return { ok: true };
}

/**
 * Logistique — Autres (hors dossiers SAV)
 * Une ligne = une arrivée ou un envoi, pour traçabilité simple.
 */
function savOtherLogList(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const sens = String(o.sens || "").trim().toUpperCase(); // "ARRIVÉE" | "ARRIVEE" | "ENVOI" | ""(all)

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_LOG_OTHER);
  if (!sh || sh.getLastRow() <= 1) return [];
  ensureHeaderGeneric_(sh, HDR_LOG_OTHER);

  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, HDR_LOG_OTHER.length).getValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const row = i + 2;
    const r = vals[i] || [];
    const id = String(r[0] || "").trim();
    const s = String(r[1] || "").trim();
    if (!id && !s && !String(r[2] || "").trim()) continue;
    const sNorm = String(s || "").trim().toUpperCase();
    if (sens) {
      const want = sens === "ARRIVEE" ? "ARRIVÉE" : sens;
      const got = sNorm === "ARRIVEE" ? "ARRIVÉE" : sNorm;
      if (want !== got) continue;
    }
    out.push({
      otherRow: row,
      id,
      sens: s || "",
      date: formatDateIfDate_(r[2]) || String(r[2] || ""),
      ref: String(r[3] || ""),
      pointDepart: String(r[4] || ""),
      adresseArrivee: String(r[5] || ""),
      transporteur: String(r[6] || ""),
      tracking: String(r[7] || ""),
      notes: String(r[8] || ""),
    });
  }
  out.sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")));
  return out;
}

function savOtherLogCreate(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const sens = String(o.sens || "").trim();
  if (!sens) return { ok: false, message: "Sens requis (ARRIVÉE ou ENVOI)." };

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_LOG_OTHER);
  if (!sh) return { ok: false, message: "Feuille Logistique_Autres introuvable." };
  ensureHeaderGeneric_(sh, HDR_LOG_OTHER);

  const id = nextId_("OTH", ss, SAV_SHEET_LOG_OTHER, 1) || ("OTH-" + new Date().getTime());
  const date = o.date ? o.date : new Date();
  sh.appendRow([
    id,
    sens,
    date,
    String(o.ref || ""),
    String(o.pointDepart || ""),
    String(o.adresseArrivee || ""),
    String(o.transporteur || ""),
    String(o.tracking || ""),
    String(o.notes || ""),
  ]);
  savCacheInvalidateAll_();
  return { ok: true, id };
}

function savOtherLogSave(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const otherRow = Number(o.otherRow || 0);
  if (!otherRow || otherRow < 2) return { ok: false, message: "Ligne invalide." };

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_LOG_OTHER);
  if (!sh) return { ok: false, message: "Feuille Logistique_Autres introuvable." };
  ensureHeaderGeneric_(sh, HDR_LOG_OTHER);
  if (sh.getLastRow() < otherRow) return { ok: false, message: "Ligne introuvable." };

  // On conserve Id + Sens, on met à jour le reste
  const cur = sh.getRange(otherRow, 1, 1, HDR_LOG_OTHER.length).getValues()[0] || [];
  const id = String(cur[0] || "").trim();
  const sens = String(cur[1] || "").trim();
  sh.getRange(otherRow, 1, 1, HDR_LOG_OTHER.length).setValues([[
    id,
    sens,
    o.date ? o.date : cur[2],
    String(o.ref || ""),
    String(o.pointDepart || ""),
    String(o.adresseArrivee || ""),
    String(o.transporteur || ""),
    String(o.tracking || ""),
    String(o.notes || ""),
  ]]);

  savCacheInvalidateAll_();
  return { ok: true };
}

function savOtherLogDelete(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const otherRow = Number(o.otherRow || 0);
  if (!otherRow || otherRow < 2) return { ok: false, message: "Ligne invalide." };

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_LOG_OTHER);
  if (!sh) return { ok: false, message: "Feuille Logistique_Autres introuvable." };
  if (sh.getLastRow() < otherRow) return { ok: false, message: "Ligne introuvable." };
  sh.deleteRow(otherRow);
  savCacheInvalidateAll_();
  return { ok: true };
}

function savUploadLogistiqueProof(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const logRow = Number(o.logRow || 0);
  const files = Array.isArray(o.files) ? o.files : [];
  if (!logRow || logRow < 2) return { ok: false, message: "Ligne logistique invalide." };
  if (!files.length) return { ok: false, message: "Aucun fichier." };

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_LOG);
  if (!sh) return { ok: false, message: "Feuille Logistique introuvable." };
  ensureHeaderGeneric_(sh, HDR_LOG);
  if (sh.getLastRow() < logRow) return { ok: false, message: "Ligne logistique introuvable." };

  const r = sh.getRange(logRow, 1, 1, HDR_LOG.length).getValues()[0] || [];
  const numero = String(r[1] || "").trim();
  if (!numero) return { ok: false, message: "Numéro dossier manquant sur la ligne logistique." };
  const sheetName = String(r[2] || "").trim(); // Feuille dossier (Distributeur/Marketplace)

  const parent = savGetDriveParentFolderForSheet_(sheetName || SAV_SHEET_DIST);
  const dossier = savGetOrCreateSubFolder_(parent, numero);
  const proofFolder = savGetOrCreateSubFolder_(dossier, "Preuves livraison");

  let created = 0;
  for (let i = 0; i < files.length; i++) {
    const f = files[i] || {};
    const name = String(f.name || ("preuve-" + (i + 1))).trim();
    const mimeType = String(f.mimeType || "application/octet-stream").trim();
    const b64 = String(f.base64 || "").trim();
    if (!b64) continue;
    const bytes = Utilities.base64Decode(b64);
    const blob = Utilities.newBlob(bytes, mimeType, name);
    proofFolder.createFile(blob);
    created++;
  }

  // Marque "Preuve livraison reçue" à OUI (colonne 16) si on a uploadé quelque chose
  if (created > 0) {
    sh.getRange(logRow, 16).setValue("OUI");
  }

  savCacheInvalidateAll_();
  return { ok: true, created, folderUrl: proofFolder.getUrl() };
}

function savGetTabBadges() {
  ensureSavSheets_();
  let list = 0;
  let log = 0;
  let pdc = 0;
  let avoir = 0;
  let mp = 0;
  try {
    const rows = savListLogistique(true) || [];
    // actions: expédier (À traiter) + livrer (Expédié)
    for (let i = 0; i < rows.length; i++) {
      const s = String(rows[i] && rows[i].statutLog ? rows[i].statutLog : "").trim();
      if (s === LOG_STAT_A_TRAITER || s === LOG_STAT_EXPEDIE) log++;
    }
  } catch (e) {}
  try {
    const rows = pdcList(true) || [];
    // actions: À traiter + Expédié
    for (let i = 0; i < rows.length; i++) {
      const s = String(rows[i] && rows[i].statut ? rows[i].statut : "").trim();
      if (s === LOG_STAT_A_TRAITER || s === LOG_STAT_EXPEDIE) pdc++;
    }
  } catch (e) {}
  try {
    const groups = savListAvoirTicketsPendingByDistributor() || [];
    for (let i = 0; i < groups.length; i++) {
      const items = groups[i] && Array.isArray(groups[i].items) ? groups[i].items : [];
      avoir += items.length;
    }
  } catch (e) {}
  try {
    // Messages IN non lus = actions à traiter côté Marketplace
    const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
    // Ne compter que les dossiers Marketplace encore "ouverts" (sinon badge reste après clôture)
    const openMp = {};
    try {
      const shMp = ss.getSheetByName(SAV_SHEET_MP);
      if (shMp && shMp.getLastRow() >= 2) {
        ensureHeaderGeneric_(shMp, HDR_MP);
        const idx = savHeaderIndexByName_(shMp);
        const lrMp = shMp.getLastRow();
        const widthMp = Math.max(shMp.getLastColumn(), HDR_MP.length);
        const valsMp = shMp.getRange(2, 1, lrMp - 1, widthMp).getValues();
        const colNum = 1;
        const etatCol = idx["État du dossier"] || 9;
        for (let i = 0; i < valsMp.length; i++) {
          const r = valsMp[i] || [];
          const numero = String(r[colNum - 1] || "").trim();
          if (!numero) continue;
          const etat = String(r[etatCol - 1] || "").trim();
          const isClos = etat === ETAT_CLOTURE_AVOIR || etat === ETAT_FERMÉ_HG;
          const isArch = etat === ETAT_ARCHIVE;
          if (isClos || isArch) continue;
          openMp[numero] = true;
        }
      }
    } catch (e) {}
    const unreadMap = savMpUnreadCounts_(); // { [numero]: count }
    const seen = {};
    Object.keys(unreadMap || {}).forEach((numero) => {
      if (numero && !openMp[numero]) return;
      const n = Number(unreadMap[numero] || 0);
      if (n > 0) {
        mp += n; // badge Marketplace = nb de messages IN non lus
        seen[numero] = true; // badge "Mes dossiers" = nb de dossiers avec message à lire
      }
    });
    list = Object.keys(seen).length;
  } catch (e) {}
  return { list, log, pdc, avoir, mp };
}

// Module "Avoir" supprimé à la demande (listing + archivage).

function formatDateIfDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(v || "").trim();
}

function updateLastTransportForNumero_(numero, patch) {
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_TR);
  const lr = sh.getLastRow();
  if (lr < 2) return;
  const nums = sh.getRange(2, 1, lr, 1).getValues();
  let lastRow = -1;
  const n = String(numero || "").trim();
  for (let i = nums.length - 1; i >= 0; i--) {
    if (String(nums[i][0] || "").trim() === n) {
      lastRow = i + 2;
      break;
    }
  }
  if (lastRow < 2) return;
  if (patch.tracking !== undefined) sh.getRange(lastRow, 5).setValue(patch.tracking);
  if (patch.statut !== undefined) sh.getRange(lastRow, 6).setValue(patch.statut);
  if (patch.preuve !== undefined) sh.getRange(lastRow, 7).setValue(patch.preuve ? "OUI" : "NON");
  if (patch.dateLivraison !== undefined && patch.dateLivraison !== "")
    sh.getRange(lastRow, 8).setValue(patch.dateLivraison instanceof Date ? patch.dateLivraison : new Date(String(patch.dateLivraison)));
}

function savSaveLogistique(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const logRow = Number(o.logRow);
  if (!logRow || logRow < 2) return { ok: false, message: "Ligne logistique invalide." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_LOG);
  const r = sh.getRange(logRow, 1, 1, HDR_LOG.length).getValues()[0];
  const feuille = String(r[2] || "");
  const ligne = Number(r[3] || 0);
  const numero = String(r[1] || "");
  const action = String(o.action || "").toLowerCase();

  const tracking = String(o.tracking !== undefined ? o.tracking : r[10] || "").trim();
  const dateExpStr = String(o.dateExpedition || "").trim();
  const transporteur = String(o.transporteur !== undefined ? o.transporteur : r[12] || "").trim();
  const coutTransport = Number(o.coutTransport !== undefined ? o.coutTransport : r[13] || 0);
  const notes = String(o.notes !== undefined ? o.notes : r[14] || "").trim();
  const preuve = !!o.preuveLivraison;
  const dateLivStr = String(o.dateLivraison || "").trim();

  sh.getRange(logRow, 11).setValue(tracking);
  sh.getRange(logRow, 13).setValue(transporteur);
  sh.getRange(logRow, 14).setValue(coutTransport);
  sh.getRange(logRow, 15).setValue(notes);

  // Transporteurs externes à facturer : coût obligatoire
  const tr = String(transporteur || "").trim().toUpperCase();
  const isBillable = tr && tr !== "DISTRIBUTEUR" && tr !== "GRATUIT";
  if (isBillable && (!coutTransport || isNaN(coutTransport) || coutTransport <= 0)) {
    return { ok: false, message: "Coût transport requis pour transporteur externe (" + transporteur + ")." };
  }

  if (dateExpStr) {
    const d = new Date(dateExpStr);
    if (!isNaN(d.getTime())) sh.getRange(logRow, 12).setValue(d);
  }

  if (action === "save") {
    // Simple sauvegarde des champs (sans changer de statut)
  } else if (action === "expedier") {
    if (!dateExpStr) sh.getRange(logRow, 12).setValue(new Date());
    sh.getRange(logRow, 9).setValue(LOG_STAT_EXPEDIE);
    sh.getRange(logRow, 16).setValue(preuve ? "OUI" : "NON");
    // Sync étape dossier (uniquement si la ligne est rattachée à un dossier SAV)
    if ((feuille === SAV_SHEET_DIST || feuille === SAV_SHEET_MP) && ligne >= 2) {
      savUpdateEtat_({ sheet: feuille, row: ligne, etat: ETAT_TRANSPORT_ROUTE });
    }
    // Marketplace: génère un ticket/message automatique "colis en route"
    try {
      if (feuille === SAV_SHEET_MP && ligne >= 2) {
        savMpAutoTicketColisEnRoute_({
          numero,
          sheet: feuille,
          rowMain: ligne,
          tracking,
          transporteur,
          dateExpedition: dateExpStr,
        });
      }
    } catch (e) {
      // ne bloque jamais la validation logistique
    }
    updateLastTransportForNumero_(numero, { tracking, statut: "En route", preuve });
  } else if (action === "livrer") {
    if (!dateLivStr) return { ok: false, message: "Date livraison requise." };
    const dl = new Date(dateLivStr);
    if (isNaN(dl.getTime())) return { ok: false, message: "Date livraison invalide." };
    sh.getRange(logRow, 9).setValue(LOG_STAT_LIVRE);
    sh.getRange(logRow, 16).setValue(preuve ? "OUI" : "NON");
    sh.getRange(logRow, 17).setValue(dl);
    if ((feuille === SAV_SHEET_DIST || feuille === SAV_SHEET_MP) && ligne >= 2) {
      // Clôture automatique : livré => archive + PDF récap archivé dans Drive
      savUpdateEtat_({ sheet: feuille, row: ligne, etat: ETAT_TRANSPORT_LIVRE });
      try {
        savUpdateEtat_({ sheet: feuille, row: ligne, etat: ETAT_ARCHIVE });
      } catch (e) {}
      try {
        savGenerateCaseRecapPdf({ sheet: feuille, row: ligne, numero: numero });
      } catch (e) {}
    }
    updateLastTransportForNumero_(numero, { tracking, statut: "Livré", preuve, dateLivraison: dl });
  } else {
    return { ok: false, message: "Action invalide (save|expedier|livrer)." };
  }

  savCacheInvalidateAll_();
  return { ok: true };
}

function savMpAutoTicketColisEnRoute_(ctx) {
  // Crée un message OUT automatique demandant d'informer le client (Marketplace).
  // Déduplication: ne crée pas 2 fois le même ticket.
  const numero = String(ctx && ctx.numero ? ctx.numero : "").trim();
  const sheet = String(ctx && ctx.sheet ? ctx.sheet : "").trim();
  const rowMain = Number(ctx && ctx.rowMain ? ctx.rowMain : 0);
  if (!numero || sheet !== SAV_SHEET_MP || rowMain < 2) return;

  const tracking = String(ctx && ctx.tracking ? ctx.tracking : "").trim();
  const transporteur = String(ctx && ctx.transporteur ? ctx.transporteur : "").trim();
  const dateExpedition = String(ctx && ctx.dateExpedition ? ctx.dateExpedition : "").trim();

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const shCase = ss.getSheetByName(SAV_SHEET_MP);
  if (!shCase) return;
  ensureHeaderGeneric_(shCase, HDR_MP);
  const idx = savHeaderIndexByName_(shCase);
  const width = Math.max(shCase.getLastColumn(), HDR_MP.length);
  const row = shCase.getRange(rowMain, 1, 1, width).getValues()[0] || [];
  const marketplace = String(row[(idx["Nom marketplace"] || 2) - 1] || "").trim();
  const clientEmail = String(row[(idx["Email client"] || 4) - 1] || "").trim();

  const shMsg = ss.getSheetByName(SAV_SHEET_MP_MESSAGES);
  if (!shMsg) return;
  ensureHeaderGeneric_(shMsg, HDR_MP_MESSAGES);

  // Dédup: cherche si un ticket AUTO:EN_ROUTE existe déjà (dernières 200 lignes)
  try {
    const lr = shMsg.getLastRow();
    if (lr >= 2) {
      const from = Math.max(2, lr - 200);
      const vals = shMsg.getRange(from, 1, lr - from + 1, HDR_MP_MESSAGES.length).getValues();
      for (let i = vals.length - 1; i >= 0; i--) {
        const r = vals[i] || [];
        if (String(r[1] || "").trim() !== numero) continue; // col B = N° dossier
        const sens = String(r[4] || "").trim(); // Sens
        const notes = String(r[9] || "").trim(); // Notes internes
        if (sens === "OUT" && notes === "AUTO:EN_ROUTE") return; // déjà créé
      }
    }
  } catch (e) {}

  const parts = [];
  parts.push("Colis en route — informer le client.");
  if (transporteur) parts.push("Transporteur : " + transporteur);
  if (tracking) parts.push("Tracking : " + tracking);
  if (dateExpedition) parts.push("Date expédition : " + dateExpedition);
  const msg = parts.join("\n");

  shMsg.appendRow([
    Utilities.getUuid(),
    numero,
    marketplace,
    clientEmail,
    "OUT",
    msg,
    "SAV",
    new Date(),
    "NON",
    "AUTO:EN_ROUTE",
  ]);

  // Archive dans Drive (dans le dossier SAV Marketplace)
  try {
    const { folder } = savEnsureDriveFolderAndWriteUrl_(sheet, rowMain, numero);
    const msgFolder = savGetOrCreateSubFolder_(folder, "Messages Marketplace");
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
    msgFolder.createFile(ts + "_AUTO_COLIS_EN_ROUTE.txt", msg + "\n", MimeType.PLAIN_TEXT);
  } catch (e) {}

  savCacheInvalidateAll_();
}

function savSetColisPret(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const numero = String(o.numero || "").trim();
  const pret = !!o.pret;
  if (!numero) return { ok: false, message: "Numéro dossier requis." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return { ok: false, message: "Aucune demande logistique pour ce dossier." };
  ensureHeaderGeneric_(sh, HDR_LOG);
  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, HDR_LOG.length).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const row = i + 2;
    const r = vals[i] || [];
    if (String(r[1] || "").trim() !== numero) continue;
    const stat = String(r[8] || "").trim();
    if (stat === LOG_STAT_LIVRE) continue;
    sh.getRange(row, 18).setValue(pret ? "OUI" : "NON"); // Colis prêt (SAV)
    sh.getRange(row, 19).setValue(pret ? new Date() : ""); // Date colis prêt
    savCacheInvalidateAll_();
    return { ok: true };
  }
  return { ok: false, message: "Aucune demande logistique ouverte pour ce dossier." };
}

function savGetLogistiqueLastForNumero(numero) {
  ensureSavSheets_();
  const n = String(numero || "").trim();
  if (!n) return null;
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return null;
  ensureHeaderGeneric_(sh, HDR_LOG);
  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, HDR_LOG.length).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i] || [];
    if (String(r[1] || "").trim() !== n) continue;
    return {
      statutLog: String(r[8] || "").trim(),
      tracking: String(r[10] || "").trim(),
      dateExpedition: formatDateIfDate_(r[11]),
      transporteur: String(r[12] || "").trim(),
      coutTransport: Number(r[13] || 0),
      preuveLivraison: String(r[15] || "").trim(),
      dateLivraison: formatDateIfDate_(r[16]),
      colisPret: String(r[17] || "").trim(),
      dateColisPret: formatDateIfDate_(r[18]),
    };
  }
  return null;
}

function savAddModele(nom) {
  ensureSavSheets_();
  const n = String(nom || "").trim();
  if (!n) return { ok: false, message: "Nom requis." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_MODELS);
  sh.appendRow([n]);
  return { ok: true };
}

// =========================
// PROCÉDURE 3 — RDV VISIO
// =========================

function savRdvVisioGetByNumero(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const numero = String(o.numero || "").trim();
  if (!numero) return null;
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_RDV_VISIO);
  if (!sh || sh.getLastRow() < 2) return null;
  ensureHeaderGeneric_(sh, HDR_RDV_VISIO);
  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, HDR_RDV_VISIO.length).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i] || [];
    if (String(r[0] || "").trim() !== numero) continue;
    return {
      numero: String(r[0] || "").trim(),
      type: String(r[1] || "").trim(),
      modele: String(r[2] || "").trim(),
      client: String(r[3] || "").trim(),
      email: String(r[4] || "").trim(),
      tel: String(r[5] || "").trim(),
      date: formatDateIfDate_(r[6]) || String(r[6] || "").trim(),
      heure: String(r[7] || "").trim(),
      etat: String(r[8] || "").trim(),
      notes: String(r[9] || "").trim(),
    };
  }
  return null;
}

function savRdvVisioUpsert(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const numero = String(o.numero || "").trim();
  if (!numero) return { ok: false, message: "Numéro dossier requis." };
  const date = String(o.date || "").trim(); // YYYY-MM-DD (input[type=date])
  const heure = String(o.heure || "").trim(); // HH:MM
  if (!date || !heure) return { ok: false, message: "Date et heure requises." };

  // Garde-fous : pas de week-end, plage 09:00–17:00
  try {
    const d = new Date(date + "T00:00:00");
    const day = d.getDay(); // 0 dim, 6 sam
    if (day === 0 || day === 6) return { ok: false, message: "RDV impossible le week-end." };
  } catch (e) {}
  if (!/^\d{2}:\d{2}$/.test(heure)) return { ok: false, message: "Heure invalide (HH:MM)." };
  const hh = Number(heure.split(":")[0] || 0);
  const mm = Number(heure.split(":")[1] || 0);
  if (hh < 9 || hh > 17 || (hh === 17 && mm > 0)) return { ok: false, message: "Heure hors plage (09:00–17:00)." };

  const type = String(o.type || "").trim();
  const modele = String(o.modele || "").trim();
  const client = String(o.client || "").trim();
  const email = String(o.email || "").trim();
  const tel = String(o.tel || "").trim();
  const etat = String(o.etat || "").trim() || "À VENIR";
  const notes = String(o.notes || "").trim();

  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_RDV_VISIO);
  if (!sh) return { ok: false, message: "Feuille RDV introuvable." };
  ensureHeaderGeneric_(sh, HDR_RDV_VISIO);

  const now = new Date();
  const lr = sh.getLastRow();
  let rowToUpdate = -1;
  if (lr >= 2) {
    const nums = sh.getRange(2, 1, lr - 1, 1).getValues();
    for (let i = nums.length - 1; i >= 0; i--) {
      if (String(nums[i][0] || "").trim() === numero) {
        rowToUpdate = i + 2;
        break;
      }
    }
  }

  const line = [numero, type, modele, client, email, tel, date, heure, etat, notes, rowToUpdate > 0 ? sh.getRange(rowToUpdate, 11).getValue() || now : now, now];
  if (rowToUpdate > 0) sh.getRange(rowToUpdate, 1, 1, HDR_RDV_VISIO.length).setValues([line]);
  else sh.appendRow(line);

  savCacheInvalidateAll_();
  return { ok: true };
}

function savRdvVisioList(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const onlyOpen = o.onlyOpen !== undefined ? !!o.onlyOpen : false;
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(SAV_SHEET_RDV_VISIO);
  if (!sh || sh.getLastRow() < 2) return [];
  ensureHeaderGeneric_(sh, HDR_RDV_VISIO);
  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, HDR_RDV_VISIO.length).getValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i] || [];
    const et = String(r[8] || "").trim() || "À VENIR";
    if (onlyOpen && et === "ANNULÉ") continue;
    out.push({
      numero: String(r[0] || "").trim(),
      type: String(r[1] || "").trim(),
      modele: String(r[2] || "").trim(),
      client: String(r[3] || "").trim(),
      email: String(r[4] || "").trim(),
      tel: String(r[5] || "").trim(),
      date: formatDateIfDate_(r[6]) || String(r[6] || "").trim(),
      heure: String(r[7] || "").trim(),
      etat: et,
      notes: String(r[9] || "").trim(),
    });
  }
  // tri date+heure asc
  out.sort((a, b) => String(a.date + " " + a.heure).localeCompare(String(b.date + " " + b.heure)));
  return out;
}

function savArchive(payload) {
  return savUpdateEtat(payload);
}

function savSetDistributorRef(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const sheet = String(o.sheet || "").trim();
  const row = Number(o.row || 0);
  const ref = String(o.refDistributeur || "").trim();
  if (!sheet || !row || row < 2) return { ok: false, message: "Paramètres invalides." };
  if (sheet !== SAV_SHEET_DIST) return { ok: false, message: "Référence disponible uniquement pour les dossiers Distributeur." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheet);
  if (!sh) return { ok: false, message: "Feuille introuvable." };
  ensureHeaderGeneric_(sh, HDR_DIST);
  const idx = savHeaderIndexByName_(sh);
  const col = idx["Référence dossier distributeur"] || 0;
  if (!col) return { ok: false, message: "Colonne référence introuvable." };
  sh.getRange(row, col).setValue(ref);
  savCacheInvalidateAll_();
  return { ok: true };
}

function savFindCaseByNumero(payload) {
  ensureSavSheets_();
  const o = payload || {};
  const numero = String(o.numero || "").trim();
  if (!numero) return { ok: false, message: "Numéro requis." };
  const ss = SpreadsheetApp.openById(SAV_SPREADSHEET_ID);
  try {
    const sheets = [
      { name: SAV_SHEET_DIST, hdr: HDR_DIST, type: "Distributeur" },
      { name: SAV_SHEET_MP, hdr: HDR_MP, type: "Marketplace" },
    ];
    for (let i = 0; i < sheets.length; i++) {
      const cfg = sheets[i];
      const sh = ss.getSheetByName(cfg.name);
      if (!sh || sh.getLastRow() < 2) continue;
      ensureHeaderGeneric_(sh, cfg.hdr);
      const rangeNums = sh.getRange(2, 1, sh.getLastRow() - 1, 1); // col A = numéro
      const finder = rangeNums.createTextFinder("^" + savEscapeRegExp_(numero) + "$").useRegularExpression(true);
      const hit = finder.findNext();
      if (hit) {
        const row = hit.getRow();
        return { ok: true, sheet: cfg.name, row, type: cfg.type, numero };
      }
    }
    return { ok: false, message: "Dossier introuvable : " + numero };
  } catch (e) {
    return { ok: false, message: "Erreur recherche dossier." };
  }
}

function savAppendPdcToCase_(ss, numero, pdcId) {
  try {
    const n = String(numero || "").trim();
    const id = String(pdcId || "").trim();
    if (!n || !id) return false;
    const cfgs = [
      { name: SAV_SHEET_DIST, hdr: HDR_DIST },
      { name: SAV_SHEET_MP, hdr: HDR_MP },
    ];
    for (let i = 0; i < cfgs.length; i++) {
      const cfg = cfgs[i];
      const sh = ss.getSheetByName(cfg.name);
      if (!sh || sh.getLastRow() < 2) continue;
      ensureHeaderGeneric_(sh, cfg.hdr);
      const idx = savHeaderIndexByName_(sh);
      const colPdc = idx["Demandes PDC"] || 0;
      if (!colPdc) continue;

      const rangeNums = sh.getRange(2, 1, sh.getLastRow() - 1, 1);
      const finder = rangeNums.createTextFinder("^" + savEscapeRegExp_(n) + "$").useRegularExpression(true);
      const hit = finder.findNext();
      if (!hit) continue;
      const row = hit.getRow();
      const cell = sh.getRange(row, colPdc);
      const existing = String(cell.getValue() || "").trim();
      const items = existing ? existing.split(/\s*\n\s*/g).filter(Boolean) : [];
      if (items.indexOf(id) === -1) items.push(id);
      cell.setValue(items.join("\n"));
      return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}

function nextSavNumber_(ss, sheetName) {
  // Numérotation globale: on numérote sur toutes les feuilles SAV
  const y = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy");
  const prefix = "SAV-" + y;
  let max = 0;
  const re = /^SAV-\d{4}-(\d+)$/;

  function scanSheet_(name) {
    try {
      const sh = ss.getSheetByName(name);
      if (!sh) return;
      const lr = sh.getLastRow();
      if (lr <= 1) return;
      const vals = sh.getRange(2, 1, lr - 1, 1).getValues();
      for (let i = 0; i < vals.length; i++) {
        const m = String(vals[i][0] || "").trim().match(re);
        if (m) max = Math.max(max, parseInt(m[1], 10));
      }
    } catch (e) {}
  }

  // Feuille demandée (compat) + les deux principales
  scanSheet_(sheetName);
  scanSheet_(SAV_SHEET_DIST);
  scanSheet_(SAV_SHEET_MP);

  // (optionnel) si d'autres feuilles numérotées sont ajoutées plus tard, on les ajoute ici.

  return prefix + "-" + String(max + 1).padStart(3, "0");
}

function nextSavNumberGlobal_(ss) {
  return nextSavNumber_(ss, SAV_SHEET_DIST);
}

function nextId_(prefix, ss, sheetName, colIdx) {
  const y = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy");
  const fullPrefix = prefix + "-" + y + "-";
  const sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return fullPrefix + "001";
  const lr = sh.getLastRow();
  const vals = sh.getRange(2, colIdx, lr - 1, 1).getValues();
  let max = 0;
  const re = new RegExp("^" + prefix + "-" + y + "-(\\d+)$");
  for (let i = 0; i < vals.length; i++) {
    const m = String(vals[i][0] || "").trim().match(re);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }
  return fullPrefix + String(max + 1).padStart(3, "0");
}

function ensureSavSheets_() {
  const id = String(SAV_SPREADSHEET_ID || "").trim();
  if (!id || id === "REMPLACE_PAR_ID_DU_CLASSEUR") {
    throw new Error("Configure SAV_SPREADSHEET_ID dans interface_sav.gs");
  }
  const ss = SpreadsheetApp.openById(id);
  const names = [
    [SAV_SHEET_DIST, HDR_DIST],
    [SAV_SHEET_MP, HDR_MP],
    [SAV_SHEET_REC, HDR_REC],
    [SAV_SHEET_EXP, HDR_EXP],
    [SAV_SHEET_TR, HDR_TR],
    [SAV_SHEET_LOG, HDR_LOG],
    [SAV_SHEET_LOG_OTHER, HDR_LOG_OTHER],
    [SAV_SHEET_PDC, HDR_PDC],
    [SAV_SHEET_MODELS, HDR_MODELS],
    [SAV_SHEET_RDV_VISIO, HDR_RDV_VISIO],
    [SAV_SHEET_PROCEDURES, HDR_PROCEDURES],
    [SAV_SHEET_PROCEDURE_MODELS, HDR_PROCEDURE_MODELS],
    [SAV_SHEET_MP_MESSAGES, HDR_MP_MESSAGES],
    [SAV_SHEET_AVOIR_TICKETS, HDR_AVOIR_TICKETS],
    [SAV_SHEET_ANALYTICS, HDR_ANALYTICS],
    [SAV_SHEET_ANALYTICS_VOLUMES, HDR_ANALYTICS_VOLUMES],
    [SAV_SHEET_ANALYTICS_YEAR_TICKETS, HDR_ANALYTICS_YEAR_TICKETS],
  ];
  for (const [name, hdr] of names) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    ensureHeaderGeneric_(sh, hdr);
  }
  seedModelsIfEmpty_(ss);
}

function seedModelsIfEmpty_(ss) {
  const sh = ss.getSheetByName(SAV_SHEET_MODELS);
  if (sh.getLastRow() > 1) return;
  const p1 = ["OVP-B40", "OVP-A40", "OVP-AR40N", "OVP-AR30B", "OVP2-B40", "OVP2-BR40N", "OVT-A15", "OVT-A30", "OVT2-A30", "OVC2-A19", "OCE-A01-2000B", "OCE-A05-2000"];
  const p2 = [
    "OVP-C40",
    "OVP-C40LUXE",
    "OPC-A01-050",
    "OPC-A01-070",
    "OPC-A01-070HP",
    "OPC-A01-090hHP",
    "OPC-A01-120HP",
    "OPC-A01-120HPWIFI",
    "OPC-A01-120",
    "OPC-A02-140",
    "OPC-A02-160HP",
    "OPC-A02-180",
    "OPC-A01-140",
    "OPC-A01-160HP",
    "OPC-A01-180",
    "OPC-B01-050",
    "OPC-B01-070",
    "OPC-B01-090",
    "OPC-B01-120",
    "OPC-B01-140",
    "OPC-C01-071",
    "OPC-C01-091",
    "OPC-C02-121",
    "OPC-C02-121HP",
    "OPC-D01-050",
    "MMCS-12HRN8-QRD0",
    "ORA-50M",
    "ORA-56D",
    "ORA-70D",
    "ORA-440D",
    "DN12",
    "DF20",
    "DM50O",
    "DH-J04-100",
    "OCE-D01-1500",
    "OCE-D01-2000",
    "OCE-D01-2500",
    "OCE-B01-1500",
    "OCE-C01-2000",
    "OCE-C03-2000",
    "OCE-C05-2200",
    "OCE-C08-1500",
    "OCE-G01-1500",
    "OCE-E01-1500",
    "OCE-F01-1500",
    "OCI-S03",
    "OCE-H02-2200",
  ];
  const p3 = [
    "OAC-300-RE1",
    "OAC-250-RE2",
    "OAC-270-SD1",
    "OAC-270-SDIN",
    "OAC-270-SDIN2",
    "OCF-IR1-90PW int",
    "OCF-IR1-90PW ext",
    "OCR-IR1-120PW int",
    "OCF-IR1-120PW ext",
    "OCF-IR1-180PW int",
    "OCF-IR1-180PW ext",
    "OCF-IR2-90PW",
    "OCF-IR2-120PW",
    "OCF-IR2-180W",
    "MULTIWIND REV",
    "OPT-ORIBT-1000",
    "OPT-ORIBT-2000",
    "OPT-ORICIS-1000",
    "OPT-ORICIS-1500",
    "OPT-ORICIS-2000",
    "OPT-ORIPR2-1000",
    "OPT-ORIPR2-1500",
    "OPT-ORIPR2-2000",
    "OPT-ORISM-2000",
    "OPT-ORISS-2000",
    "OPT-ORSS1-500/750",
    "OPT-ORSS2-1500/1750",
    "OPT-ORSS3-500/750",
    "OPT-ORSS4-1500/1750",
    "GCAT200i",
    "GCAT250i",
    "WINDPAC-EM45C",
    "WINDPAC-EM55C",
    "WINDPAC-ET90C",
  ];
  const max = Math.max(p1.length, p2.length, p3.length);
  const seed = [];
  for (let i = 0; i < max; i++) seed.push([p1[i] || "", p2[i] || "", p3[i] || ""]);
  sh.getRange(2, 1, seed.length, 3).setValues(seed);
}

function ensureHeaderGeneric_(sh, headers) {
  const width = headers.length;
  const firstRow = sh.getRange(1, 1, 1, width).getValues()[0];
  const isEmpty = firstRow.every((v) => v === "" || v === null);
  const matches = firstRow.every((v, i) => String(v || "").trim() === headers[i]);
  if (isEmpty || !matches) {
    sh.getRange(1, 1, 1, width).setValues([headers]);
    sh.getRange(1, 1, 1, width).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
}

function getSavHtml_(payloadB64) {
  const siteLogoHtml = ""; // Logo chargé de façon asynchrone côté client
  var html = `<!DOCTYPE html>
<html lang="fr">
<head>
  <meta http-equiv="cache-control" content="no-cache, no-store, must-revalidate" />
  <meta http-equiv="pragma" content="no-cache" />
  <meta http-equiv="expires" content="0" />
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Gestion SAV</title>
  <style>
    :root{
      --bg:#f6f8fc;
      --panel:#ffffff;
      --card:#f8fafc;
      --border:rgba(15,23,42,.12);
      --text:#0f172a;
      --muted:#475569;
      --accent:#2563eb; /* bleu doux */
      --ok:#16a34a;
      --warn:#d97706;
      --bad:#dc2626;
      --shadow:0 10px 30px rgba(15,23,42,.06);
    }
    *{box-sizing:border-box}
    body{margin:0;font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;background:var(--bg);color:var(--text);min-height:100vh}
    /* Utilise mieux l'écran (moins de marges latérales) */
    .app{max-width:none;margin:0;padding:14px 16px}
    .top{display:flex;flex-wrap:wrap;gap:12px;align-items:center;justify-content:space-between;margin-bottom:16px}
    h1{margin:0;font-size:1.25rem}
    .tabs{display:flex;gap:8px;flex-wrap:wrap}
    .tab{position:relative;border:1px solid var(--border);background:var(--panel);color:var(--text);padding:10px 14px;border-radius:12px;cursor:pointer;font-weight:800}
    .tab.on{border-color:rgba(37,99,235,.35);background:rgba(37,99,235,.10);color:var(--accent)}
    .tabBadge{position:absolute;top:-6px;right:-6px;min-width:18px;height:18px;padding:0 6px;border-radius:999px;background:var(--bad);color:#fff;font-size:11px;font-weight:950;display:none;align-items:center;justify-content:center;box-shadow:0 6px 14px rgba(220,38,38,.25)}
    .grid{display:grid;grid-template-columns:minmax(680px,1fr) 320px;gap:14px;align-items:stretch}
    @media(max-width:960px){.grid{grid-template-columns:1fr}}
    @media(max-width:1100px){
      .app > .row{flex-direction:column}
    }
    #mainGrid.logistique-mode{grid-template-columns:1fr}
    #mainGrid.logistique-mode .panel:nth-child(2){display:none}
    .panel{background:var(--panel);border:1px solid var(--border);border-radius:16px;overflow:hidden;box-shadow:var(--shadow)}
    /* Donne plus de hauteur au panneau "Nouvelle demande" */
    #mainPanel{min-height:calc(100vh - 190px)}
    #mainBody{min-height:calc(100vh - 260px)}
    .ph{padding:12px 14px;border-bottom:1px solid var(--border);font-weight:800;font-size:13px}
    .pb{padding:14px}
    label{display:block;font-size:11px;font-weight:800;color:var(--muted);margin:10px 0 6px;text-transform:uppercase;letter-spacing:.04em}
    input,select,textarea{width:100%;padding:10px 12px;border-radius:12px;border:1px solid var(--border);background:var(--card);color:var(--text);font-size:13px}
    input:focus,select:focus,textarea:focus{outline:none;border-color:rgba(37,99,235,.45);box-shadow:0 0 0 3px rgba(37,99,235,.12)}
    textarea{min-height:88px;resize:vertical}
    .btn{border:1px solid var(--border);background:#ffffff;color:var(--text);padding:10px 14px;border-radius:12px;cursor:pointer;font-weight:800}
    .btn:hover{background:#f1f5f9}
    .btn.primary{border-color:rgba(37,99,235,.35);background:rgba(37,99,235,.10);color:var(--accent)}
    .row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
    .muted{color:var(--muted);font-size:12px}
    .badge{display:inline-block;padding:4px 10px;border-radius:999px;font-size:11px;font-weight:900;background:rgba(148,163,184,.12);border:1px solid var(--border)}
    .badge.ok{background:rgba(22,163,74,.12);border-color:rgba(22,163,74,.28);color:#166534}
    .badge.warn{background:rgba(217,119,6,.12);border-color:rgba(217,119,6,.28);color:#92400e}
    .list{display:flex;flex-direction:column;gap:8px;max-height:420px;overflow:auto}
    .item{padding:10px 12px;border-radius:12px;border:1px solid var(--border);background:var(--card);cursor:pointer}
    .item:hover{border-color:rgba(37,99,235,.25);background:#ffffff}
    .tl{display:flex;flex-direction:column;gap:8px;font-size:12px}
    .tl div{display:flex;gap:8px;align-items:center}
    .dot{width:8px;height:8px;border-radius:999px;background:var(--muted)}
    .dot.on{background:var(--accent)}
    .hide{display:none!important}
    .diag{display:none;margin-bottom:12px;padding:12px;border-radius:12px;border:1px solid rgba(220,38,38,.28);background:rgba(220,38,38,.08);font-size:12px;white-space:pre-wrap;color:#7f1d1d}
  </style>
</head>
<body>
  <div class="app">
    <div class="top">
      <div>
        <div id="siteLogo" style="display:flex;justify-content:center;margin:0 0 10px 0"></div>
        <script>
          // Charge le logo de façon asynchrone (ne bloque pas le chargement)
          try {
            window.setTimeout(function() {
              google.script.run
                .withSuccessHandler(function(url) {
                  if (!url) return;
                  var el = document.getElementById("siteLogo");
                  if (el) el.innerHTML = '<img src="' + url + '" style="max-width:170px;max-height:56px;object-fit:contain" />';
                })
                .withFailureHandler(function() {})
                .savGetLogoUrl();
            }, 1500);
          } catch(e) {}
        </script>
        <h1>Gestion SAV</h1>
        <div class="muted" id="ver">—</div>
      </div>
      <div class="tabs">
        <button type="button" class="tab on" data-tab="new">Nouvelle</button>
        <button type="button" class="tab" data-tab="list">Mes dossiers<span class="tabBadge" id="badge_list"></span></button>
        <button type="button" class="tab" data-tab="arch">Archivés</button>
        <button type="button" class="tab" data-tab="rdv">📅 Calendrier RDV</button>
        <button type="button" class="tab" data-tab="log">Logistique<span class="tabBadge" id="badge_log"></span></button>
        <button type="button" class="tab" data-tab="pdc">Envois pièces détachées<span class="tabBadge" id="badge_pdc"></span></button>
        <button type="button" class="tab" data-tab="avoir">Avoir (tickets)<span class="tabBadge" id="badge_avoir"></span></button>
        <button type="button" class="tab" data-tab="mp">Marketplace<span class="tabBadge" id="badge_mp"></span></button>
        <button type="button" class="tab" data-tab="sys">Système</button>
      </div>
    </div>
    <div class="diag" id="diag"></div>
    <div class="row" style="align-items:flex-start;gap:14px">
      <div class="panel" id="statsSide" style="width:320px;position:sticky;top:12px">
        <div class="ph">Statistiques</div>
        <div class="pb" id="statsBody"></div>
      </div>
      <div class="grid" id="mainGrid" style="flex:1">
      <div class="panel" id="mainPanel">
        <div class="ph" id="mainTitle">Nouvelle demande</div>
        <div class="pb" id="mainBody"></div>
      </div>
      <div class="panel">
        <div class="ph">Dossier sélectionné</div>
        <div class="pb">
          <div class="muted" id="sideNum">Aucun</div>
          <div style="margin-top:10px"><span class="badge" id="sideEtat">—</span></div>
          <div id="sideProc" class="hide" style="margin-top:12px;padding:10px 12px;border:1px solid rgba(56,189,248,.35);background:rgba(56,189,248,.08);border-radius:12px"></div>
          <div class="tl" id="timeline" style="margin-top:14px"></div>
          <div id="sideActions" style="margin-top:14px;display:flex;flex-direction:column;gap:8px"></div>
        </div>
      </div>
      </div>
    </div>
  </div>
  <script>
    function showErr(e){
      var d=document.getElementById('diag');
      d.style.display='block';
      d.textContent=(e&&e.message)?e.message:String(e);
    }
    function showPdfLink(res){
      // Affiche un bandeau cliquable avec le lien de téléchargement du PDF
      var d=document.getElementById('diag');
      if(!res||!res.fileUrl){ showErr('PDF généré mais lien introuvable.'); return; }
      d.style.display='block';
      d.style.background='#f0fdf4';
      d.style.color='#166534';
      d.style.borderColor='#bbf7d0';
      d.innerHTML='<b>PDF prêt :</b> '+(res.fileName||'document.pdf')+
        ' &nbsp;<a href="'+res.fileUrl+'" target="_blank" rel="noopener" '+
        'style="color:#15803d;font-weight:700;text-decoration:underline">Télécharger / Ouvrir ↗</a>'+
        ' &nbsp;<span style="cursor:pointer;opacity:.5" onclick="this.parentNode.style.display='none'">✕</span>';
    }
    function b64ToObj(b64){
      var bin=atob(b64);
      // Décodage UTF-8 robuste (TextDecoder pas dispo partout, ex: IE/compat mode)
      try{
        if(typeof TextDecoder!=='undefined'){
          var bytes=new Uint8Array(bin.length);
          for(var i=0;i<bin.length;i++) bytes[i]=bin.charCodeAt(i);
          var txt=new TextDecoder('utf-8').decode(bytes);
          return JSON.parse(txt);
        }
      }catch(e){}
      // Fallback: decodeURIComponent(escape(...)) (suffisant ici, payload JSON)
      var txt2;
      try{
        txt2=decodeURIComponent(escape(bin));
      }catch(e2){
        txt2=bin; // dernier recours (ASCII)
      }
      return JSON.parse(txt2);
    }
    var boot=b64ToObj(${JSON.stringify(payloadB64)});
    document.getElementById('ver').textContent='Version '+boot.version;

    var models=boot.models||[];
    // Important: si on a déjà des modèles dans le payload initial, on évite un appel serveur
    // au chargement (peut déclencher un panneau OAuth et donner l'impression "rien ne charge").
    var modelsLoadedOnce=!!(models && models.length);
    var procDefs=boot.procDefs||{};
    var procMap=boot.procMap||{};
    var stats=null; // chargées à la demande
    var current=null;
    var centralesAchat=[]; // suggestions "mémo"

    function buildStatsHtml_(s){
      if(!s) return '';
      function line(label, value, cls){
        return (
          '<div style="display:flex;justify-content:space-between;gap:10px;align-items:center;padding:8px 10px;border:1px solid var(--border);border-radius:12px;background:var(--card)">' +
          '<div class="muted" style="font-weight:900;letter-spacing:.04em;text-transform:uppercase;font-size:11px">' +
          esc(label) +
          "</div>" +
          '<span class="badge' +
          (cls ? " " + cls : "") +
          '">' +
          esc(value) +
          "</span>" +
          "</div>"
        );
      }
      var html='';
      html += '<div style="display:flex;flex-direction:column;gap:8px">';
      html += '<div class="muted" style="font-weight:950;margin:0 0 2px 2px">Dossiers</div>';
      html += line('Total des dossiers ouverts', s.totalOuverts||0, '');
      html += line('Dossiers en cours — distributeur', s.distEnCours||0, '');
      html += line('Dossiers en cours — marketplace', s.mpEnCours||0, '');
      html += line('Dossiers clos', s.dossiersClos||0, 'ok');
      if((s.dossiersRepares||0)||(s.dossiersEchanges||0)||(s.dossiersAvoir||0)||(s.dossiersHG||0)){
        html += '<div style="height:6px"></div>';
        html += '<div class="muted" style="font-weight:950;margin:0 0 2px 2px">Dossiers finalisés (résultat)</div>';
        html += line('Réparés', s.dossiersRepares||0, 'ok');
        html += line('Échangés', s.dossiersEchanges||0, 'ok');
        html += line('Avoir', s.dossiersAvoir||0, 'ok');
        if(s.dossiersHG||0) html += line('Hors garantie / Refus', s.dossiersHG||0, 'warn');
      }
      html += '<div style="height:6px"></div>';
      html += '<div class="muted" style="font-weight:950;margin:0 0 2px 2px">Transport et logistique</div>';
      html += line('Demandes à traiter', s.transportATraiter||0, 'warn');
      html += line('Demandes en route', s.transportExpedie||0, '');
      html += line('Demandes livrées', s.transportLivre||0, 'ok');
      if(s.delaiMoyenLivraisonJ!=null){
        html += line('Délai moyen de livraison (jours)', s.delaiMoyenLivraisonJ, '');
      }
      if(s.transportTrackingManquant){
        html += line('Expédié sans tracking', s.transportTrackingManquant||0, 'warn');
      }
      html += '<div style="height:6px"></div>';
      html += '<div class="muted" style="font-weight:950;margin:0 0 2px 2px">Pièces détachées</div>';
      html += line('Total (PDC)', s.pdcTotal||0, '');
      html += line('En cours (PDC)', s.pdcEnCours||0, (s.pdcEnCours? 'warn':'ok'));
      html += line('Livrées (PDC)', s.pdcLivre||0, 'ok');
      if(s.pdcDelaiMoyenLivraisonJ!=null){
        html += line('Délai moyen de livraison PDC (jours)', s.pdcDelaiMoyenLivraisonJ, '');
      }
      html += '<div style="height:6px"></div>';
      html += '<div class="muted" style="font-weight:950;margin:0 0 2px 2px">Marketplace</div>';
      html += line('Messages à traiter', s.mpMessagesATraiter||0, (s.mpMessagesATraiter? 'warn':'ok'));
      html += '<div style="height:6px"></div>';
      html += '<div class="muted" style="font-weight:950;margin:0 0 2px 2px">Délais moyens de clôture</div>';
      if(s.delaiMoyenClotureDistJ!=null) html += line('Distributeur (jours)', s.delaiMoyenClotureDistJ, '');
      if(s.delaiMoyenClotureMpJ!=null) html += line('Marketplace (jours)', s.delaiMoyenClotureMpJ, '');
      html += '</div>';
      return html;
    }

    // Suggestions "centrales d'achat" (mémo)
    // IMPORTANT: on ne charge plus ça au démarrage pour éviter de bloquer l'UI
    // si l'environnement OAuth/iframe Google est capricieux.
    function loadCentralesAchat_(){
      try{
        google.script.run.withSuccessHandler(function(list){
          centralesAchat = Array.isArray(list)?list:[];
        }).withFailureHandler(function(){}).savListCentralesAchat();
      }catch(e){}
    }

    function renderStatsInto_(containerEl, s){
      if(!containerEl) return;
      containerEl.innerHTML = buildStatsHtml_(s) || '<div class="muted">Aucune statistique.</div>';
    }

    function renderStatsShell_(){
      var wrap=document.getElementById('statsBody');
      if(!wrap) return;
      // Ne pas dupliquer
      if(document.getElementById('btnLoadStatsLeft')) return;
      wrap.innerHTML =
        '<div class="row" style="gap:8px;align-items:center;justify-content:space-between;margin-bottom:8px">' +
          '<div class="muted" style="font-weight:950">Statistiques</div>' +
          '<button class="btn primary" type="button" id="btnLoadStatsLeft" style="white-space:nowrap">Rafraîchir</button>' +
        '</div>' +
        '<div class="muted" id="statsLeftStatus" style="margin:0 0 8px 0">Chargement des statistiques...</div>' +
        '<div id="statsLeftOut"></div>';

      var btn=document.getElementById('btnLoadStatsLeft');
      var status=document.getElementById('statsLeftStatus');
      var out=document.getElementById('statsLeftOut');
      function load(){
        if(!btn) return;
        btn.disabled=true;
        btn.textContent='Chargement...';
        if(status) status.textContent='Chargement des statistiques...';
        google.script.run.withSuccessHandler(function(s){
          stats=s||null;
          if(status) status.textContent='';
          renderStatsInto_(out, stats);
          btn.disabled=false;
          btn.textContent='Rafraîchir';
        }).withFailureHandler(function(e){
          if(status) status.textContent=(e&&e.message)?e.message:String(e);
          btn.disabled=false;
          btn.textContent='Rafraîchir';
          showErr(e);
        }).savGetStats();
      }
      if(btn) btn.onclick=load;
      // Auto-chargement, mais différé: laisse l'UI se rendre d'abord
      // (sinon le panneau OAuth Google peut donner l'impression "rien ne charge").
      try{ window.setTimeout(load, 600); }catch(e){ try{ load(); }catch(e2){} }
    }

    var ETAT_CREATION="CRÉATION";
    var ETAT_ATT="En attente réception";
    var ETAT_RECU="Réception atelier";
    var ETAT_EXP="Expertise";
    var ETAT_TR_D="Transport - demande";
    var ETAT_TR_R="Transport - en route";
    var ETAT_TR_L="Transport - livré";
    var ETAT_AVOIR="Clôture avec Avoir";
    var ETAT_HG="FERMÉ - Hors garantie";
    var ETAT_ARCH="Archivé";

    // Procédures (côté navigateur)
    var SAV_PROC_POSTAL="PROC_1_POSTAL";
    var SAV_PROC_RAPATRIEMENT="PROC_2_RAPATRIEMENT";
    var SAV_PROC_VISIO="PROC_3_VISIO";

    function stepFromEtat(e){
      if(e===ETAT_CREATION) return 1;
      if(e===ETAT_ATT) return 3;
      if(e===ETAT_RECU) return 4;
      if(e===ETAT_EXP||e===ETAT_TR_D) return 5;
      if(e===ETAT_TR_R) return 6;
      if(e===ETAT_TR_L||e===ETAT_AVOIR||e===ETAT_HG||e===ETAT_ARCH) return 7;
      return 2;
    }

    function renderTimeline(etat){
      var n=stepFromEtat(etat);
      var labels=['Saisie','Début','Attente réception','Réception','Expertise','Transport','Clôture'];
      var el=document.getElementById('timeline');
      el.innerHTML='';
      for(var i=0;i<labels.length;i++){
        var on=(i+1)<=n;
        el.innerHTML+='<div><span class="dot'+(on?' on':'')+'"></span><span>'+labels[i]+'</span></div>';
      }
    }

    function setCurrent(c){
      current=c;
      document.getElementById('sideNum').textContent=c?c.numero:'Aucun';
      var b=document.getElementById('sideEtat');
      if(!c){
        b.textContent='—';b.className='badge';
        document.getElementById('sideActions').innerHTML='';
        var sp=document.getElementById('sideProc');
        if(sp){ sp.classList.add('hide'); sp.innerHTML=''; }
        renderTimeline('');
        return;
      }
      b.textContent=c.etat||'—';
      b.className='badge'+(c.etat===ETAT_ATT?' warn':'')+(c.etat===ETAT_AVOIR||c.etat===ETAT_ARCH?' ok':'');
      renderTimeline(c.etat||ETAT_CREATION);
      renderSideActions(c);

      // Pré-charge (à la demande) les centrales d'achat quand on ouvre un dossier en expertise,
      // pour éviter un appel serveur au démarrage tout en gardant l'aide au prompt "Avoir".
      try{
        if(c && String(c.etat||'')===ETAT_EXP && !window.__centralesLoaded){
          window.__centralesLoaded=true;
          loadCentralesAchat_();
        }
      }catch(e){}

      // Procédure visible en permanence dans la sidebar
      var sp=document.getElementById('sideProc');
      if(sp){
        var mod=(c&&c.modele)?String(c.modele).trim():'';
        var pid=procMap[String(mod||'').trim()]||'';
        var p=pid?procDefs[pid]:null;
        if(mod){
          var h='';
          h+='<div style="font-weight:900;margin-bottom:6px">Logique SAV</div>';
          if(p){
            h+='<div style="font-size:12px"><b>'+esc(p.label)+'</b></div>';
            if(p.objectif) h+='<div class="muted" style="margin-top:6px">'+esc(p.objectif)+'</div>';
            if(p.documentation&&p.documentation.length){
              h+='<div style="margin-top:8px;font-size:12px"><b>Docs à demander</b><ul style="margin:6px 0 0 18px;padding:0">';
              p.documentation.forEach(function(x){ h+='<li>'+esc(x)+'</li>'; });
              h+='</ul></div>';
            }
            if(p.etapes&&p.etapes.length){
              h+='<div style="margin-top:8px;font-size:12px"><b>Déroulé</b><ol style="margin:6px 0 0 18px;padding:0">';
              p.etapes.forEach(function(x){ h+='<li>'+esc(x)+'</li>'; });
              h+='</ol></div>';
            }
          }else{
            h+='<div class="muted" style="font-size:12px">Aucune procédure définie pour ce modèle.</div>';
          }
          sp.innerHTML=h;
          sp.classList.remove('hide');
        }else{
          sp.classList.add('hide');
          sp.innerHTML='';
        }
      }
    }

    function findByNumero_(arr, num){
      var n=String(num||'');
      if(!arr||!arr.length) return null;
      for(var i=0;i<arr.length;i++){
        var x=arr[i];
        if(x && String(x.numero)===n) return x;
      }
      return null;
    }

    function renderSideActions(c){
      var box=document.getElementById('sideActions');
      box.innerHTML='';
      var e=c.etat||ETAT_CREATION;
      var pid='';
      if(c && c.type==='Distributeur'){
        var mod=(c&&c.modele)?String(c.modele).trim():'';
        pid=procMap[String(mod||'').trim()]||'';
      }

      function checkbox(label, checked, onChange){
        var wrap=document.createElement('div');
        wrap.style.padding='10px 12px';
        wrap.style.border='1px solid var(--border)';
        wrap.style.borderRadius='12px';
        wrap.style.background='rgba(255,255,255,.03)';
        var id='cb_'+Math.random().toString(36).slice(2);
        wrap.innerHTML =
          '<label style="margin:0;display:flex;align-items:flex-start;gap:8px;font-size:13px;font-weight:800;color:var(--text);text-transform:none;letter-spacing:0" for="'+id+'">'+
            '<input type="checkbox" id="'+id+'" style="width:auto;margin:2px 0 0 0" '+(checked?'checked':'')+' />'+
            esc(label)+
          '</label>';
        box.appendChild(wrap);
        var cb=document.getElementById(id);
        if(cb){
          cb.onchange=function(){
            if(cb.disabled) return;
            cb.disabled=true;
            onChange(cb.checked, function(){ cb.disabled=false; });
          };
        }
      }

      function btn(label, fn){
        var b=document.createElement('button');
        b.className='btn primary';
        b.type='button';
        b.textContent=label;
        b.onclick=function(){
          if(b.disabled) return;
          b.disabled=true;
          var oldLabel=label;
          b.textContent='...';
          try{ fn(b, oldLabel); }
          catch(err){ b.disabled=false; b.textContent=oldLabel; throw err; }
        };
        box.appendChild(b);
      }
      function reenable_(b, oldLabel){
        try{ if(b){ b.disabled=false; b.textContent=oldLabel||'Valider'; } }catch(e){}
      }

      if(e===ETAT_CREATION){
        // Procédures 2 & 3 : checkbox "mail envoyé" -> passe automatiquement en attente réception
        if(pid===SAV_PROC_RAPATRIEMENT || pid===SAV_PROC_VISIO){
          var already = String(c.mailAccordRetour||'').trim().toUpperCase()==='OUI';
          checkbox(
            "Mail envoyer au distributeur pour accord de retour en atelier pour expertise",
            already,
            function(val, done){
              if(!val){ done(); return; } // on ne gère pas la décoche (pas de retour arrière)
              var emailTo = prompt('Email destinataire (optionnel) pour envoyer l\\'étiquette retour en PDF :','')||'';
              google.script.run.withSuccessHandler(function(){
                reloadList(function(){
                  var u=findByNumero_(allCases, c.numero);
                  if(u){ setCurrent(u); viewDetail(u); }
                  else { c.etat=ETAT_ATT; setCurrent(c); viewDetail(c); }
                });
                done();
              }).withFailureHandler(function(err){
                done();
                showErr(err);
              }).savSetMailAccordRetour({sheet:c.sheet,row:c.row,sent:true,emailTo:emailTo});
            }
          );
        }else{
        if(pid==='PROC_1_POSTAL'){
          btn('Demande retour plaque (PDF) →', function(b, oldLabel){
            google.script.run.withSuccessHandler(function(res){
              if(res&&res.ok){ showPdfLink(res); }
              else{ showErr((res&&res.message)?res.message:'Erreur génération PDF'); }
              reenable_(b, oldLabel);
            }).withFailureHandler(function(e){
              reenable_(b, oldLabel);
              showErr(e);
            }).savGenerateProc1PostalRequestPdf({sheet:c.sheet,row:c.row,numero:c.numero});
          });
          btn('Feuille retour plaque (PDF) →', function(b, oldLabel){
            google.script.run.withSuccessHandler(function(res){
              if(res&&res.ok){ showPdfLink(res); }
              else{ showErr((res&&res.message)?res.message:'Erreur génération PDF'); }
              reenable_(b, oldLabel);
            }).withFailureHandler(function(e){
              reenable_(b, oldLabel);
              showErr(e);
            }).savGeneratePlaqueSheetPdf({sheet:c.sheet,row:c.row,numero:c.numero});
          });
        }
        btn('En attente de réception →', function(b, oldLabel){
          google.script.run.withSuccessHandler(function(){
            reloadList(function(){
              var u=findByNumero_(allCases, c.numero);
              if(u){ setCurrent(u); viewDetail(u); }
              else { c.etat=ETAT_ATT; setCurrent(c); viewDetail(c); }
            });
            reenable_(b, oldLabel);
          }).withFailureHandler(function(e){
            reenable_(b, oldLabel);
            showErr(e);
          }).savUpdateEtat({sheet:c.sheet,row:c.row,etat:ETAT_ATT});
        });
        }
      }
      if(e===ETAT_ATT){
        // Procédure 1 (postal) : pas de réception atelier
        if(pid!=='PROC_1_POSTAL'){
          btn('Bon de retour (PDF) →', function(b, oldLabel){
            google.script.run
              .withSuccessHandler(function(res){
                if(res&&res.ok){ showPdfLink(res); }
                else{ showErr((res&&res.message)?res.message:'Erreur génération PDF'); }
                reenable_(b, oldLabel);
              })
              .withFailureHandler(function(e){
                reenable_(b, oldLabel);
                showErr(e);
              })
              .savGenerateReturnLabelPdf({sheet:c.sheet,row:c.row,numero:c.numero});
          });
          btn('Appareil reçu en atelier →', function(b, oldLabel){ reenable_(b, oldLabel); openReceptionForm(c); });
        }
      }
      if(e===ETAT_RECU){
        btn('Passer en expertise →', function(b, oldLabel){
          google.script.run.withSuccessHandler(function(){
            reloadList(function(){
              var u=findByNumero_(allCases, c.numero);
              if(u){ setCurrent(u); viewDetail(u); }
              else { c.etat=ETAT_EXP; setCurrent(c); viewDetail(c); }
            });
            reenable_(b, oldLabel);
          }).withFailureHandler(function(e){
            reenable_(b, oldLabel);
            showErr(e);
          }).savUpdateEtat({sheet:c.sheet,row:c.row,etat:ETAT_EXP});
        });
      }
      if(e===ETAT_EXP){
        btn('Renseigner expertise', function(b, oldLabel){ reenable_(b, oldLabel); openExpertiseForm(c); });
        btn('Créer ticket avoir', function(b, oldLabel){
          var hint = centralesAchat && centralesAchat.length
            ? ('\\n\\nCentrales déjà utilisées :\\n- ' + centralesAchat.slice(0, 12).join('\\n- ') + (centralesAchat.length>12?'\\n…':''))
            : '';
          var def = (window.__lastCentraleAchat||'');
          var centrale=prompt('Centrale d\\'achat (obligatoire) ?'+hint, def)||'';
          if(!String(centrale||'').trim()){
            reenable_(b, oldLabel);
            showErr('Centrale d\\'achat requise.');
            return;
          }
          try{ window.__lastCentraleAchat=String(centrale||'').trim(); }catch(e){}
          var facture=prompt('Facture d\\'achat centrale (optionnel) ?')||'';
          var notes=prompt('Notes pour l\\'avoir (optionnel) ?')||'';
          google.script.run.withSuccessHandler(function(res){
            if(res && res.ok){
              reloadList(function(){
                showTab('avoir');
              });
            }else{
              showErr((res&&res.message)?res.message:'Erreur création ticket');
            }
            reenable_(b, oldLabel);
          }).withFailureHandler(function(e){
            reenable_(b, oldLabel);
            showErr(e);
          }).savCreateAvoirTicket({sheet:c.sheet,row:c.row,numero:c.numero,centraleAchat:centrale,factureAchatCentrale:facture,notes:notes});
        });
        btn('Bon de transport / échange', function(b, oldLabel){ reenable_(b, oldLabel); openTransportForm(c,'Échange'); });
        btn('Réparation (transport)', function(b, oldLabel){ reenable_(b, oldLabel); openTransportForm(c,'Réparation'); });
        btn('En attente pièces', function(b, oldLabel){
          google.script.run.withSuccessHandler(function(){
            reloadList(function(){
              var u=findByNumero_(allCases, c.numero);
              if(u){ setCurrent(u); viewDetail(u); }
              else { c.etat='EN ATTENTE - Pièces'; setCurrent(c); viewDetail(c); }
            });
            reenable_(b, oldLabel);
          }).withFailureHandler(function(e){
            reenable_(b, oldLabel);
            showErr(e);
          }).savUpdateEtat({sheet:c.sheet,row:c.row,etat:'EN ATTENTE - Pièces'});
        });
        btn('Clôturer hors garantie', function(b, oldLabel){
          if(!confirm('Clôturer ce dossier (hors garantie) ?'))return;
          // Génère le bon d'envoi (PDF) vers l'adresse de destination, puis clôture.
          var destAddr = (c && c.type==='Distributeur') ? (c.adresse||'') : (c.clientAdresse||c.adresse||'');
          if(!String(destAddr||'').trim()){
            reenable_(b, oldLabel);
            showErr('Adresse de destination requise (hors garantie).');
            return;
          }
          google.script.run.withSuccessHandler(function(res){
            try{
              if(res && res.ok && res.fileUrl){ showPdfLink(res); }
            }catch(e){}
            reloadList(function(){
              var u=findByNumero_(allCases, c.numero);
              if(u){ setCurrent(u); viewDetail(u); }
              else { c.etat=ETAT_HG; setCurrent(c); viewDetail(c); }
            });
            reenable_(b, oldLabel);
          }).withFailureHandler(function(e){
            reenable_(b, oldLabel);
            showErr(e);
          }).savGenerateLogistiqueShippingLabelPdf({
            sheet:c.sheet,
            rowMain:c.row,
            numero:c.numero,
            typeEnvoi:'Hors garantie',
            adresseArrivee:destAddr,
          });
        });
      }
      // Simplification : la progression Transport (en route / livré) est pilotée dans l’onglet Logistique.
      // On évite ainsi de valider 2 fois la même étape (détail dossier + logistique).
      if((e===ETAT_TR_D||e===ETAT_TR_R)&&e!==ETAT_TR_L){
        btn('Gérer transport (Logistique) →', function(b, oldLabel){
          showTab('log');
          reenable_(b, oldLabel);
        });
      }
      if((e===ETAT_TR_L||e===ETAT_AVOIR||e===ETAT_HG)&&e!==ETAT_ARCH){
        btn('Archiver', function(b, oldLabel){
          google.script.run.withSuccessHandler(function(){ reloadList(); setCurrent(null); showTab('list');})
          .withFailureHandler(function(e){ reenable_(b, oldLabel); showErr(e); })
          .savUpdateEtat({sheet:c.sheet,row:c.row,etat:ETAT_ARCH});
        });
        // Rapport technique : utile côté Distributeur uniquement (pas Marketplace)
        if(c.type==='Distributeur'){
          btn('Rapport technique (PDF) →', function(b, oldLabel){
            google.script.run.withSuccessHandler(function(res){
              if(res&&res.ok){ showPdfLink(res); }
              else{ showErr((res&&res.message)?res.message:'Erreur génération PDF'); }
              reenable_(b, oldLabel);
            }).withFailureHandler(function(e){
              reenable_(b, oldLabel);
              showErr(e);
            }).savGenerateCaseRecapPdf({sheet:c.sheet,row:c.row,numero:c.numero});
          });
        }
      }
    }

    function openReceptionForm(c){
      document.getElementById('mainTitle').textContent='Réception atelier';
      document.getElementById('mainBody').innerHTML=
        '<label>Date réception</label><input type="date" id="rDate" />'+
        '<label>État visuel</label><textarea id="rVis"></textarea>'+
        '<label>Éléments accompagnants</label><textarea id="rElt" placeholder="Chargeur, câbles..."></textarea>'+
        '<label>Notes</label><textarea id="rNotes"></textarea>'+
        '<label>Technicien</label><input type="text" id="rTech" />'+
        '<div style="margin-top:12px"><button class="btn primary" type="button" id="rSave">Enregistrer</button></div>';
      document.getElementById('rSave').onclick=function(){
        google.script.run.withSuccessHandler(function(){
          reloadList(function(){
            var u=findByNumero_(allCases, c.numero);
            if(u){ setCurrent(u); viewDetail(u); }
            else { c.etat=ETAT_RECU; setCurrent(c); viewDetail(c); }
          });
        }).withFailureHandler(showErr).savAddReception({
          numero:c.numero,dateReception:document.getElementById('rDate').value,
          etatVisuel:document.getElementById('rVis').value,
          elements:document.getElementById('rElt').value,
          notes:document.getElementById('rNotes').value,
          technicien:document.getElementById('rTech').value,
          sheet:c.sheet,rowMain:c.row
        });
      };
    }

    function openExpertiseForm(c){
      document.getElementById('mainTitle').textContent='Expertise';
      document.getElementById('mainBody').innerHTML='<div class="muted">Chargement...</div>';
      google.script.run.withSuccessHandler(function(last){
        last=last||{};
        document.getElementById('mainBody').innerHTML=
          '<label>Date expertise</label><input type="date" id="eDate" value="'+esc(last.dateExpertise||'')+'" />'+
          '<label>Technicien</label><input type="text" id="eTech" value="'+esc(last.technicien||'')+'" />'+
          '<label>Travaux réalisés</label><textarea id="eTrav">'+esc(last.travaux||'')+'</textarea>'+
          '<label>Pièces utilisées</label><textarea id="ePieces" placeholder="Références, quantités, détails...">'+esc(last.pieces||'')+'</textarea>'+
          '<label>Tests effectués</label><textarea id="eTests" placeholder="Mesures, tests de sécurité, essais...">'+esc(last.tests||'')+'</textarea>'+
          '<label>Conclusions</label><textarea id="eConc" placeholder="Cause, diagnostic final...">'+esc(last.conclusions||'')+'</textarea>'+
          '<label>Décision</label><select id="eDec">'+
            '<option value="">—</option>'+
            '<option>Réparation validée</option>'+
            '<option>Échange</option>'+
            '<option>Refus garantie</option>'+
            '<option>Hors garantie</option>'+
            '<option>Avoir</option>'+
          '</select>'+
          '<div class="row" style="gap:10px;align-items:flex-end">'+
            '<div style="flex:1;min-width:200px"><label>Coût pièces (€)</label><input type="number" step="0.01" id="eCoutPieces" value="'+esc(last.coutPieces||'')+'" /></div>'+
            '<div style="flex:1;min-width:200px"><label>Coût main d\\'œuvre (€)</label><input type="number" step="0.01" id="eCoutMO" value="'+esc(last.coutMainOeuvre||'')+'" /></div>'+
          '</div>'+
          '<div style="margin-top:12px"><button class="btn primary" type="button" id="eSave">Enregistrer l\\'expertise</button></div>';
        var sel=document.getElementById('eDec');
        if(sel && last.decision){ sel.value=String(last.decision); }
        var btn=document.getElementById('eSave');
        btn.onclick=function(){
          if(btn.disabled) return;
          btn.disabled=true;
          var old=btn.textContent;
          btn.textContent='...';
          google.script.run.withSuccessHandler(function(){
            reloadList(function(){
              var u=findByNumero_(allCases, c.numero);
              if(u){ setCurrent(u); viewDetail(u); }
              else { setCurrent(c); viewDetail(c); }
            });
          }).withFailureHandler(function(e){
            btn.disabled=false;
            btn.textContent=old;
            showErr(e);
          }).savSaveExpertise({
            numero:c.numero,
            dateExpertise:document.getElementById('eDate').value,
            technicien:document.getElementById('eTech').value,
            travaux:document.getElementById('eTrav').value,
            pieces:document.getElementById('ePieces').value,
            tests:document.getElementById('eTests').value,
            conclusions:document.getElementById('eConc').value,
            decision:document.getElementById('eDec').value,
            coutPieces:document.getElementById('eCoutPieces').value,
            coutMainOeuvre:document.getElementById('eCoutMO').value
          });
        };
      }).withFailureHandler(showErr).savGetExpertiseLastForNumero(c.numero);
    }

    function openTransportForm(c, type){
      document.getElementById('mainTitle').textContent='Envoi / transport';
      var defaultArr = (c && c.type==='Distributeur') ? (c.adresse||'') : (c.clientAdresse||'');
      var defaultFrom = 'Sealogis';
      document.getElementById('mainBody').innerHTML=
        '<div class="row" style="gap:10px;align-items:stretch;flex-wrap:wrap">'+
          '<div style="flex:1;min-width:280px;padding:10px 12px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.03)">'+
            '<div style="font-weight:900;margin-bottom:8px">Adresse de départ</div>'+
            '<label>Point départ *</label>'+
            '<select id="tP"><option>Sealogis</option><option>Optimea</option></select>'+
            '<label>Adresse de départ (détails)</label><textarea id="tFrom" placeholder="Adresse complète de départ..."></textarea>'+
          '</div>'+
          '<div style="flex:1;min-width:280px;padding:10px 12px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.03)">'+
            '<div style="font-weight:900;margin-bottom:8px">Adresse d\\'arrivée</div>'+
            '<label>Adresse d\\'arrivée *</label><textarea id="tTo" placeholder="Adresse complète d\\'arrivée..."></textarea>'+
          '</div>'+
        '</div>'+
        '<label>Tracking (optionnel)</label><input type="text" id="tTr" />'+
        '<div style="margin-top:12px"><button class="btn primary" id="tSave">Créer demande</button></div>';

      // préremplissage
      try{
        document.getElementById('tFrom').value = defaultFrom;
        document.getElementById('tTo').value = defaultArr;
      }catch(e){}
      var sel=document.getElementById('tP');
      if(sel){
        sel.onchange=function(){
          var v=String(sel.value||'');
          var from=document.getElementById('tFrom');
          if(from) from.value=v;
        };
      }

      document.getElementById('tSave').onclick=function(){
        var btn=document.getElementById('tSave');
        if(btn){ btn.disabled=true; btn.textContent='Création...'; }
        google.script.run
          .withSuccessHandler(function(res){
            if(btn){ btn.disabled=false; btn.textContent='Créer demande'; }
            // Si une étiquette a été générée, on affiche le lien PDF cliquable.
            if(res && res.label && res.label.fileUrl){
              var main=document.getElementById('mainBody');
              if(main){
                main.innerHTML =
                  '<div class="evt" style="margin-bottom:14px">'+
                    '<div style="font-weight:950">Étiquette générée</div>'+
                    '<div class="muted">PDF : '+esc(res.label.fileName||'')+'</div>'+
                    '<div class="row" style="margin-top:10px">'+
                      '<a class="btn primary" href="'+esc(res.label.fileUrl)+'" target="_blank" rel="noopener">Télécharger / Imprimer</a>'+
                    '</div>'+
                  '</div>'+
                  '<div style="margin-top:10px">'+
                    '<button class="btn" type="button" id="backToDossier">Retour au dossier</button>'+
                  '</div>';
                var back=document.getElementById('backToDossier');
                if(back){
                  back.onclick=function(){
                    reloadList(function(){
                      var u=findByNumero_(allCases, c.numero);
                      if(u){ setCurrent(u); viewDetail(u); }
                      else { c.etat=ETAT_TR_D; setCurrent(c); viewDetail(c); }
                    }, false);
                  };
                }
                return;
              }
            }
            reloadList(function(){
              var u=findByNumero_(allCases, c.numero);
              if(u){ setCurrent(u); viewDetail(u); }
              else { c.etat=ETAT_TR_D; setCurrent(c); viewDetail(c); }
            }, false);
          })
          .withFailureHandler(function(e){
            if(btn){ btn.disabled=false; btn.textContent='Créer demande'; }
            showErr(e);
          })
          .savAddTransport({
            numero:c.numero,
            typeEnvoi:type,
            pointDepart:document.getElementById('tFrom').value || document.getElementById('tP').value,
            adresseArrivee:document.getElementById('tTo').value,
            tracking:document.getElementById('tTr').value,statut:'Demande',preuve:false,
            sheet:c.sheet,rowMain:c.row,updateEtat:ETAT_TR_D
          });
      };

    function viewDetail(c){
      document.getElementById('mainTitle').textContent='Dossier '+c.numero;
      var html='<div class="muted">Type : '+c.type+'</div>';
      if(c.type==='Distributeur'){
        html+='<p><b>Magasin</b> '+esc(c.magasin)+'<br><b>Réf dossier distributeur</b> <span id="refDistTxt">'+esc(c.refDistributeur||'')+'</span><br><b>Adresse</b> '+esc(c.adresse)+'<br><b>Modèle</b> '+esc(c.modele)+'<br><b>Panne</b> '+esc(c.panne)+'<br><b>Série</b> '+esc(c.serie)+'</p>';
        html+='<div style="margin-top:-6px;margin-bottom:10px;padding:10px 12px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.03)">';
        html+='<div class="muted" style="font-weight:950;margin:0 0 8px 2px">Référence dossier distributeur</div>';
        html+='<div class="row" style="gap:10px;flex-wrap:wrap;align-items:center">';
        html+='<input id="refDistInp" placeholder="Référence interne / centrale / commande..." style="flex:1;min-width:220px" value="'+esc(c.refDistributeur||'')+'" />';
        html+='<button class="btn" type="button" id="refDistSave">Enregistrer</button>';
        html+='</div>';
        html+='<div class="muted" id="refDistRes" style="margin-top:6px"></div>';
        html+='</div>';
        var pid=procMap[String(c.modele||'').trim()]||'';
        var p=pid?procDefs[pid]:null;
        html+='<div style="margin-top:10px;padding:10px 12px;border:1px solid rgba(56,189,248,.35);background:rgba(56,189,248,.08);border-radius:12px">';
        html+='<div style="font-weight:900;margin-bottom:6px">Logique SAV</div>';
        if(p){
          html+='<div style="font-size:12px"><b>'+esc(p.label)+'</b></div>';
          html+='<div class="muted" style="margin-top:6px">'+esc(p.objectif||'')+'</div>';
        }else{
          html+='<div class="muted" style="font-size:12px">Aucune procédure définie pour ce modèle.</div>';
        }
        html+='</div>';
      }else{
        html+='<p><b>Marketplace</b> '+esc(c.marketplace)+'<br><b>Client</b> '+esc(c.clientNom)+'<br><b>Email</b> '+esc(c.clientEmail)+'</p>';
      }
      // Procédure + RDV visio (PROC 3) : affiché pour tous les dossiers qui ont un modèle
      try{
        var pid2=procMap[String(c.modele||'').trim()]||'';
        if(pid2===SAV_PROC_VISIO){
          html+='<div style="margin-top:12px;padding:10px 12px;border:1px solid rgba(99,102,241,.35);background:rgba(99,102,241,.08);border-radius:12px">';
          html+='<div style="font-weight:900;margin-bottom:6px">Procédure 3 — RDV visio</div>';
          html+='<div class="muted" style="font-size:12px">Planifie le diagnostic visio puis retrouve tous les RDV dans l’onglet “📅 Calendrier RDV”.</div>';
          html+='<div class="row" style="gap:10px;flex-wrap:wrap;align-items:end;margin-top:10px">';
          html+='<div style="min-width:160px"><label>Date RDV *</label><input type="date" id="rdvDate" /></div>';
          html+='<div style="min-width:160px"><label>Heure RDV *</label><select id="rdvHeure"></select></div>';
          html+='<div style="min-width:180px"><label>État</label><select id="rdvEtat"><option>À VENIR</option><option>EFFECTUÉ</option><option>ANNULÉ</option></select></div>';
          html+='</div>';
          html+='<label style="margin-top:10px">Notes</label><textarea id="rdvNotes" placeholder="Notes du diagnostic / contexte..."></textarea>';
          html+='<div class="row" style="gap:10px;flex-wrap:wrap;align-items:center;margin-top:10px">';
          html+='<button class="btn primary" type="button" id="rdvSave">Confirmer le rendez-vous</button>';
          html+='<button class="btn" type="button" id="rdvOpenTab">Ouvrir calendrier RDV</button>';
          html+='</div>';
          html+='<div class="muted" id="rdvRes" style="margin-top:8px"></div>';
          html+='</div>';
        }
      }catch(e){}
      html+='<p><b>État</b> '+esc(c.etat)+' · <b>Créé</b> '+esc(c.dateCreation)+'</p>';
      html+='<div id="logInfo" class="muted" style="margin-top:10px"></div>';
      html+='<div id="colisPretWrap" style="margin-top:10px"></div>';
      html+='<div style="margin-top:12px" class="row" id="caseTools" style="gap:10px;flex-wrap:wrap"></div>';
      html+='<div id="caseTimelineWrap" class="hide" style="margin-top:12px"></div>';
      if(c.type!=='Distributeur'){
        html+='<div id="mpTicketWrap" style="margin-top:12px"></div>';
      }
      document.getElementById('mainBody').innerHTML=html;
      setCurrent(c);

      // Bind RDV visio si PROC 3
      try{
        var pid3=procMap[String(c.modele||'').trim()]||'';
        if(pid3===SAV_PROC_VISIO){
          var sel=document.getElementById('rdvHeure');
          if(sel){
            var times=[];
            for(var h=9; h<=17; h++){
              for(var m=0; m<60; m+=30){
                if(h===17 && m>0) continue;
                var hh=String(h).padStart(2,'0');
                var mm=String(m).padStart(2,'0');
                times.push(hh+':'+mm);
              }
            }
            sel.innerHTML = '<option value="">—</option>' + times.map(function(t){ return '<option value="'+esc(t)+'">'+esc(t)+'</option>'; }).join('');
          }
          var bOpen=document.getElementById('rdvOpenTab');
          if(bOpen) bOpen.onclick=function(){ showTab('rdv'); };

          // load existing RDV if any
          google.script.run.withSuccessHandler(function(r){
            if(!r) return;
            try{
              if(document.getElementById('rdvDate') && r.date){
                // accept either yyyy-mm-dd or dd/mm/yyyy (we keep best effort)
                if(String(r.date||'').indexOf('-')>0) document.getElementById('rdvDate').value=String(r.date);
              }
              if(document.getElementById('rdvHeure') && r.heure) document.getElementById('rdvHeure').value=String(r.heure);
              if(document.getElementById('rdvEtat') && r.etat) document.getElementById('rdvEtat').value=String(r.etat);
              if(document.getElementById('rdvNotes')) document.getElementById('rdvNotes').value=String(r.notes||'');
            }catch(e){}
          }).withFailureHandler(function(){}).savRdvVisioGetByNumero({numero:c.numero});

          var bSave=document.getElementById('rdvSave');
          if(bSave){
            bSave.onclick=function(){
              var res=document.getElementById('rdvRes');
              var date=document.getElementById('rdvDate')?document.getElementById('rdvDate').value:'';
              var heure=document.getElementById('rdvHeure')?document.getElementById('rdvHeure').value:'';
              var etat=document.getElementById('rdvEtat')?document.getElementById('rdvEtat').value:'';
              var notes=document.getElementById('rdvNotes')?document.getElementById('rdvNotes').value:'';
              if(!date || !heure){ showErr('Date et heure RDV requises.'); return; }
              bSave.disabled=true;
              if(res) res.textContent='Enregistrement...';
              google.script.run.withSuccessHandler(function(r){
                bSave.disabled=false;
                if(r&&r.ok){
                  if(res) res.textContent='OK — RDV enregistré.';
                }else{
                  if(res) res.textContent='';
                  showErr((r&&r.message)?r.message:'Erreur RDV');
                }
              }).withFailureHandler(function(e){
                bSave.disabled=false;
                if(res) res.textContent='';
                showErr(e);
              }).savRdvVisioUpsert({
                numero:c.numero,
                type:c.type,
                modele:c.modele,
                client:(c.clientNom||c.magasin||''),
                email:(c.clientEmail||''),
                tel:(c.clientTel||''),
                date:date,
                heure:heure,
                etat:etat,
                notes:notes
              });
            };
          }
        }
      }catch(e){}

      // Outils dossier : Historique + impression + PDF
      // Accessible uniquement pour les dossiers archivés (sinon la barre d'étapes suffit).
      if(String(c.etat||'')==='Archivé'){
        try{
          var tools=document.getElementById('caseTools');
          if(tools){
            tools.style.gap='10px';
            tools.style.flexWrap='wrap';
            tools.innerHTML =
              '<button class="btn" type="button" id="btnTimeline">Historique / étapes</button>'+
              '<button class="btn" type="button" id="btnTimelinePrint">Imprimer</button>'+
              '<button class="btn" type="button" id="btnTimelinePdf">Exporter PDF</button>';
          }
          var btnT=document.getElementById('btnTimeline');
          var btnP=document.getElementById('btnTimelinePrint');
          var btnPdf=document.getElementById('btnTimelinePdf');
          var wrap=document.getElementById('caseTimelineWrap');
          function ensureTimeline_(cb){
            if(!wrap) return;
            if(wrap.getAttribute('data-loaded')==='1'){ if(cb) cb(); return; }
            wrap.innerHTML='<div class="muted">Chargement de l\\'historique...</div>';
            google.script.run.withSuccessHandler(function(r){
              if(r&&r.ok){
                wrap.setAttribute('data-loaded','1');
                wrap.innerHTML = r.html || '<div class="muted">Aucun historique.</div>';
                if(cb) cb();
              }else{
                wrap.innerHTML='';
                showErr((r&&r.message)?r.message:'Erreur');
              }
            }).withFailureHandler(showErr).savGetCaseTimelineHtml({sheet:c.sheet,row:c.row,numero:c.numero});
          }
          if(btnT){
            btnT.onclick=function(){
              if(!wrap) return;
              var show = wrap.classList.contains('hide');
              if(show){
                wrap.classList.remove('hide');
                ensureTimeline_();
              }else{
                wrap.classList.add('hide');
              }
            };
          }
          if(btnP){
            btnP.onclick=function(){
              ensureTimeline_(function(){
                if(!wrap) return;
                var w=window.open('','_blank');
                if(!w) return;
                w.document.write('<!doctype html><meta charset="utf-8"><title>Historique '+esc(c.numero)+'</title>');
                w.document.write('<style>body{font-family:Arial,sans-serif;margin:24px} .btn{display:none} a{color:#2563eb} table{page-break-inside:auto} tr{page-break-inside:avoid}</style>');
                w.document.write(wrap.innerHTML);
                w.document.close();
                w.focus();
                w.print();
              });
            };
          }
          if(btnPdf){
            btnPdf.onclick=function(){
              google.script.run.withSuccessHandler(function(res){
                if(res&&res.ok){
                  showPdfLink(res);
                }else{
                  showErr((res&&res.message)?res.message:'Erreur génération PDF');
                }
              }).withFailureHandler(showErr).savGenerateCaseRecapPdf({sheet:c.sheet,row:c.row,numero:c.numero});
            };
          }
        }catch(e){}
      }else{
        // cache les conteneurs si dossier non archivé
        try{
          var tools=document.getElementById('caseTools');
          if(tools) tools.innerHTML='';
          var wrap=document.getElementById('caseTimelineWrap');
          if(wrap){ wrap.innerHTML=''; wrap.classList.add('hide'); }
        }catch(e){}
      }

      // Référence dossier distributeur (persistante + utilisée pour recherche)
      if(c.type==='Distributeur'){
        var bRef=document.getElementById('refDistSave');
        if(bRef){
          bRef.onclick=function(){
            var inp=document.getElementById('refDistInp');
            var res=document.getElementById('refDistRes');
            if(!inp) return;
            bRef.disabled=true;
            var v=inp.value||'';
            if(res) res.textContent='Enregistrement...';
            google.script.run.withSuccessHandler(function(r){
              bRef.disabled=false;
              if(r&&r.ok){
                if(res) res.textContent='OK';
                try{ c.refDistributeur=String(v||''); }catch(e){}
                var t=document.getElementById('refDistTxt');
                if(t) t.textContent=String(v||'');
                try{ reloadList(function(){}); }catch(e){}
              }else{
                if(res) res.textContent='';
                showErr((r&&r.message)?r.message:'Erreur');
              }
            }).withFailureHandler(function(e){
              bRef.disabled=false;
              if(res) res.textContent='';
              showErr(e);
            }).savSetDistributorRef({sheet:c.sheet,row:c.row,refDistributeur:v});
          };
        }
      }

      // Synchronisation logistique : affiche date départ / livraison + checkbox "colis prêt"
      google.script.run.withSuccessHandler(function(li){
        if(!li){
          var el=document.getElementById('logInfo');
          if(el) el.textContent='';
          var w=document.getElementById('colisPretWrap');
          if(w) w.innerHTML='';
          return;
        }
        var info=[];
        if(li.statutLog) info.push('Logistique : '+li.statutLog);
        if(li.dateExpedition) info.push('Départ : '+li.dateExpedition);
        if(li.dateLivraison) info.push('Livraison : '+li.dateLivraison);
        if(li.tracking) info.push('Tracking : '+li.tracking);
        var el=document.getElementById('logInfo');
        if(el) el.textContent=info.join(' · ');

        var w=document.getElementById('colisPretWrap');
        if(w){
          var checked = String(li.colisPret||'').trim().toUpperCase()==='OUI';
          w.innerHTML =
            '<label style="margin:0;display:flex;align-items:center;gap:8px;font-size:13px;font-weight:800;color:var(--text);text-transform:none;letter-spacing:0">' +
              '<input type="checkbox" id="colisPret" style="width:auto;margin:0" '+(checked?'checked':'')+' />' +
              'Colis prêt (prévenir la logistique)' +
            '</label>' +
            (li.dateColisPret ? ('<div class="muted" style="margin-top:6px">Déclaré prêt le '+esc(li.dateColisPret)+'</div>') : '');
          var cb=document.getElementById('colisPret');
          if(cb){
            cb.onchange=function(){
              cb.disabled=true;
              google.script.run.withSuccessHandler(function(){
                cb.disabled=false;
                // refresh log info
                google.script.run.withSuccessHandler(function(li2){
                  if(li2){
                    var info=[];
                    if(li2.statutLog) info.push('Logistique : '+li2.statutLog);
                    if(li2.dateExpedition) info.push('Départ : '+li2.dateExpedition);
                    if(li2.dateLivraison) info.push('Livraison : '+li2.dateLivraison);
                    if(li2.tracking) info.push('Tracking : '+li2.tracking);
                    var el=document.getElementById('logInfo');
                    if(el) el.textContent=info.join(' · ');
                  }
                }).withFailureHandler(showErr).savGetLogistiqueLastForNumero(c.numero);
              }).withFailureHandler(function(e){
                cb.disabled=false;
                showErr(e);
              }).savSetColisPret({numero:c.numero,pret:cb.checked});
            };
          }
        }
      }).withFailureHandler(showErr).savGetLogistiqueLastForNumero(c.numero);

      // Messagerie Marketplace (dossier Marketplace)
      if(c.type!=='Distributeur'){
        google.script.run.withSuccessHandler(function(msgs){
          msgs=msgs||[];
          var box=document.getElementById('mpTicketWrap');
          if(!box) return;
          var unread=0;
          msgs.forEach(function(m){
            if(String(m.sens||'')==='IN' && String(m.lu||'').toUpperCase()!=='OUI') unread++;
          });
          var h='';
          h+='<div style="padding:10px 12px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.03)">';
          h+='<div class="row" style="justify-content:space-between;align-items:center">';
          h+='<div style="font-weight:900">Messagerie Marketplace</div>';
          h+= unread ? ('<span class="badge warn">'+esc(unread)+' non lu(s)</span>') : '<span class="badge ok">Tout lu</span>';
          h+='</div>';
          h+='<div class="muted" style="margin-top:6px">Fil lié à ce dossier. Les messages sont archivés dans Drive.</div>';

          h+='<div style="margin-top:10px">';
          if(!msgs.length){
            h+='<div class="muted">Aucun message.</div>';
          }else{
            // affiche les 8 derniers pour rester léger
            var start=Math.max(0, msgs.length-8);
            for(var i=start;i<msgs.length;i++){
              var m=msgs[i]||{};
              var out=(String(m.sens||'')==='OUT');
              h+='<div style="margin-top:8px;padding:10px 12px;border-radius:12px;border:1px solid '+(out?'rgba(34,197,94,.35)':'rgba(245,158,11,.35)')+';background:'+(out?'rgba(34,197,94,.08)':'rgba(245,158,11,.08)')+'">';
              h+='<div style="font-size:12px;font-weight:900">'+(out?'SAV → Client':'Client → SAV')+'</div>';
              h+='<div class="muted" style="margin-top:2px">'+esc(m.date||'')+' · '+esc(m.auteur||'')+'</div>';
              h+='<div style="margin-top:8px;white-space:pre-wrap;font-size:13px">'+esc(m.message||'')+'</div>';
              h+='</div>';
            }
          }
          h+='</div>';

          h+='<div style="margin-top:12px">';
          h+='<label>Message (SAV → client)</label><textarea id="dMpMsg" placeholder="Écris ton message…"></textarea>';
          h+='<div class="row" style="margin-top:10px;gap:10px;flex-wrap:wrap">';
          h+='<button class="btn primary" type="button" id="dMpSend">Envoyer</button>';
          h+='<button class="btn" type="button" id="dMpOpen">Ouvrir dans Marketplace →</button>';
          h+='</div>';
          h+='<div class="muted" id="dMpRes" style="margin-top:8px"></div>';
          h+='</div>';

          h+='</div>';
          box.innerHTML=h;

          // marque les messages IN comme lus (ticket traité) dès ouverture du dossier
          if(unread){
            google.script.run.withSuccessHandler(function(){ try{ refreshTabBadges_(); }catch(e){} }).withFailureHandler(function(){}).savMpMarkAllRead({numero:c.numero});
          }

          var send=document.getElementById('dMpSend');
          if(send){
            send.onclick=function(){
              var txt=document.getElementById('dMpMsg').value||'';
              if(!String(txt).trim()){ showErr('Message requis.'); return; }
              send.disabled=true;
              var old=send.textContent;
              send.textContent='...';
              var payload={numero:c.numero,marketplace:c.marketplace,clientEmail:c.clientEmail,sens:'OUT',message:txt,auteur:'',sheet:c.sheet,rowMain:c.row};
              google.script.run.withSuccessHandler(function(res){
                send.disabled=false;
                send.textContent=old;
                if(res&&res.ok){
                  document.getElementById('dMpRes').textContent='OK — message enregistré.';
                  viewDetail(c); // refresh
                  try{ refreshTabBadges_(); }catch(e){}
                }else{
                  showErr((res&&res.message)?res.message:'Erreur');
                }
              }).withFailureHandler(function(e){
                send.disabled=false;
                send.textContent=old;
                showErr(e);
              }).savMpAddMessage(payload);
            };
          }
          var open=document.getElementById('dMpOpen');
          if(open){
            open.onclick=function(){ try{ showTab('mp'); }catch(e){} };
          }
        }).withFailureHandler(showErr).savMpListMessages(c.numero);
      }
    }

    function esc(s){ return String(s||'').replace(/[&<>"']/g,function(m){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]);}); }

    var LOG_A='À traiter';
    var LOG_E='Expédié';
    var LOG_L='Livré';
    var LOG_X='Annulé';

    function renderLogistique(){
      document.getElementById('mainTitle').textContent='Logistique';
      var mode = window.__logMode || 'sav';
      document.getElementById('mainBody').innerHTML=
        '<div class="row" style="gap:10px;flex-wrap:wrap;margin-bottom:10px">'+
          '<button class="btn '+(mode==='sav'?'primary':'')+'" type="button" id="logModeSav">Transport SAV</button>'+
          '<button class="btn '+(mode==='pdc'?'primary':'')+'" type="button" id="logModePdc">Pièces détachées</button>'+
          '<button class="btn '+(mode==='other'?'primary':'')+'" type="button" id="logModeOther">Autres</button>'+
        '</div>'+
        '<div class="muted">Chargement...</div>';

      var bSav=document.getElementById('logModeSav');
      var bPdc=document.getElementById('logModePdc');
      if(bSav) bSav.onclick=function(){ window.__logMode='sav'; renderLogistique(); };
      if(bPdc) bPdc.onclick=function(){ window.__logMode='pdc'; renderLogistique(); };
      var bO=document.getElementById('logModeOther');
      if(bO) bO.onclick=function(){ window.__logMode='other'; renderLogistique(); };

      if(mode==='other'){
        renderLogistiqueOther_();
        return;
      }

      var fetch = (mode==='pdc') ? 'pdc' : 'sav';
      var run = (fetch==='pdc')
        ? google.script.run.withFailureHandler(showErr).withSuccessHandler(function(res){
            res=res||{};
            var rows=res.rows||[];
            var hasMore=!!res.hasMore;
            var html=document.getElementById('mainBody').innerHTML;
            // garde les 3 boutons (évite split fragile sur HTML)
            var top = '<div class="row" style="gap:10px;flex-wrap:wrap;margin-bottom:10px">'+
                        '<button class="btn '+(mode==='sav'?'primary':'')+'" type="button" id="logModeSav">Transport SAV</button>'+
                        '<button class="btn '+(mode==='pdc'?'primary':'')+'" type="button" id="logModePdc">Pièces détachées</button>'+
                        '<button class="btn '+(mode==='other'?'primary':'')+'" type="button" id="logModeOther">Autres</button>'+
                      '</div>';
            var h=top;
            h += '<div class="evt" style="margin-bottom:14px">';
            h += '<div style="font-weight:900">Pièces détachées — envois</div>';
            h += '<div class="muted">Création et suivi des envois de pièces (hors dossiers SAV).</div>';
            h += '<div style="margin-top:8px;padding:10px 12px;border:1px solid rgba(37,99,235,.25);border-radius:12px;background:rgba(37,99,235,.04)">';
            h +=   '<div style="font-weight:900;margin-bottom:6px">Procédure logistique (PDC)</div>';
            h +=   '<div class="muted">1) Vérifier que “Colis prêt” = OUI et que les adresses / références sont complètes.<br>2) Renseigner le transporteur et le poids.<br>3) Valider expédition.<br>4) Une fois le transport effectué, renseigner le tracking (si nécessaire) puis valider livraison.<br>5) Archiver les documents (devis / BL / facture) dans la zone Documents.</div>';
            h += '</div>';
            h += '<div class="row" style="margin-top:10px;gap:10px;flex-wrap:wrap">';
            h +=   '<button class="btn primary" type="button" id="pdcNew">Créer un envoi PDC</button>';
            h += '</div>';
            h += '</div>';

            var tab = window.__pdcTab || 'open';
            var limit = Number(window.__pdcLimit||50);
            var offset = Number(window.__pdcOffset||0);
            var y = new Date().getFullYear();
            var year = (window.__pdcYear===undefined)?y:window.__pdcYear;
            var yearLabel = (year===null || year===0)?'Toutes':String(year);

            h += '<div class="row" style="gap:10px;flex-wrap:wrap;margin:4px 0 12px 0">';
            h +=   '<button class="btn '+(tab==='open'?'primary':'')+'" type="button" data-pdc-tab="open">En cours</button>';
            h +=   '<button class="btn '+(tab==='a'?'primary':'')+'" type="button" data-pdc-tab="a">À traiter</button>';
            h +=   '<button class="btn '+(tab==='e'?'primary':'')+'" type="button" data-pdc-tab="e">Expédié</button>';
            h +=   '<button class="btn '+(tab==='l'?'primary':'')+'" type="button" data-pdc-tab="l">Livré</button>';
            h +=   '<button class="btn '+(tab==='x'?'primary':'')+'" type="button" data-pdc-tab="x">Annulé</button>';
            h +=   '<button class="btn '+(tab==='all'?'primary':'')+'" type="button" data-pdc-tab="all">Tout</button>';
            h +=   '<span class="badge">Année '+esc(yearLabel)+'</span>';
            h += '</div>';

            if(!rows.length){
              h += '<div class="muted">Aucun envoi pour ce filtre.</div>';
              document.getElementById('mainBody').innerHTML=h;
            }else{
              rows.forEach(function(x){
                h += '<div class="evt" style="margin-bottom:14px">';
                var s=String(x.statut||'');
                var klass = (s===LOG_A?'warn':(s===LOG_E?'':'ok'));
                if(s===LOG_X) klass='muted';
                h += '<div class="row"><div class="name">'+esc(x.id)+'</div><span class="badge '+klass+'">'+esc(s)+'</span></div>';
                h += '<div class="muted">Créé le '+esc(x.dateCreation||'')+'</div>';
            h += '<label>Point de départ</label><textarea id="pdcFrom_'+x.pdcRow+'">'+esc(x.pointDepart||'')+'</textarea>';
            h += '<label>Adresse arrivée</label><textarea id="pdcTo_'+x.pdcRow+'">'+esc(x.adresseArrivee||'')+'</textarea>';
            h += '<label>Références</label><input type="text" id="pdcRefs_'+x.pdcRow+'" value="'+esc(x.refs||'')+'" />';
            h += '<label>Désignation</label><input type="text" id="pdcDes_'+x.pdcRow+'" value="'+esc(x.designation||'')+'" />';
            var ready = String(x.colisPret||'').toUpperCase()==='OUI';
            h += '<label class="row" style="margin-top:10px"><input type="checkbox" id="pdcReady_'+x.pdcRow+'" '+(ready?'checked':'')+' /> Colis prêt à l’envoi</label>';
            h += '<label>Poids (kg)</label><input type="number" step="0.01" id="pdcW_'+x.pdcRow+'" value="'+esc((x.poids===''||x.poids==null)?'':String(x.poids))+'" />';
            h += '<label>Transporteur</label><input type="text" id="pdcT_'+x.pdcRow+'" value="'+esc(x.transporteur||'DISTRIBUTEUR')+'" />';
                if(String(x.statut||'')===LOG_E){
                  h += '<label>Tracking</label><input type="text" id="pdcK_'+x.pdcRow+'" value="'+esc(x.tracking||'')+'" />';
                }
                h += '<label>Notes</label><textarea id="pdcN_'+x.pdcRow+'">'+esc(x.notes||'')+'</textarea>';
                h += '<div class="row" style="margin-top:10px;gap:10px;flex-wrap:wrap">';
                h +=   '<button class="btn" type="button" data-pdc-act="save" data-row="'+x.pdcRow+'">Enregistrer</button>';
                if(String(x.statut||'')===LOG_A){
                  h += '<button class="btn primary" type="button" data-pdc-act="exp" data-row="'+x.pdcRow+'">Valider expédition</button>';
                }else if(String(x.statut||'')===LOG_E){
                  h += '<button class="btn primary" type="button" data-pdc-act="liv" data-row="'+x.pdcRow+'">Valider livraison</button>';
                }
                if(String(x.statut||'')!==LOG_L && String(x.statut||'')!==LOG_X){
                  h += '<button class="btn" type="button" data-pdc-act="cancel" data-row="'+x.pdcRow+'">Annuler / supprimer</button>';
                }
                h += '</div>';
                h += '<div style="margin-top:10px;padding:10px 12px;border:1px dashed var(--border);border-radius:12px;background:rgba(255,255,255,.03)">';
                h +=   '<div style="font-weight:900;margin-bottom:6px">Documents (devis / BL / facture)</div>';
                h +=   '<input type="file" id="pdcF_'+x.pdcRow+'" data-pdc-files="1" multiple />';
                h +=   '<div class="muted" id="pdcR_'+x.pdcRow+'" style="margin-top:6px"></div>';
                h += '</div>';
                h += '</div>';
              });
              if(hasMore){
                h += '<div class="row" style="gap:10px;flex-wrap:wrap;margin:6px 0 14px 0">';
                h +=   '<button class="btn" type="button" id="pdcMore">Charger plus</button>';
                h += '</div>';
              }
              document.getElementById('mainBody').innerHTML=h;
            }

            var bn=document.getElementById('pdcNew');
            if(bn){
              bn.onclick=function(){
                google.script.run.withSuccessHandler(function(){ renderLogistique(); }).withFailureHandler(showErr).pdcCreate({transporteur:'DISTRIBUTEUR',poids:'',tracking:'',notes:''});
              };
            }
            // rebind nav buttons
            var bSav2=document.getElementById('logModeSav');
            var bPdc2=document.getElementById('logModePdc');
            var bO2=document.getElementById('logModeOther');
            if(bSav2) bSav2.onclick=function(){ window.__logMode='sav'; renderLogistique(); };
            if(bPdc2) bPdc2.onclick=function(){ window.__logMode='pdc'; renderLogistique(); };
            if(bO2) bO2.onclick=function(){ window.__logMode='other'; renderLogistique(); };
            document.querySelectorAll('button[data-pdc-tab]').forEach(function(b){
              b.onclick=function(){
                window.__pdcTab=String(b.getAttribute('data-pdc-tab')||'open');
                window.__pdcOffset=0;
                renderLogistique();
              };
            });
            var more=document.getElementById('pdcMore');
            if(more){
              more.onclick=function(){
                window.__pdcOffset = Number(window.__pdcOffset||0) + Number(window.__pdcLimit||50);
                renderLogistique();
              };
            }
            document.querySelectorAll('button[data-pdc-act]').forEach(function(b){
              b.onclick=function(){
                var row=Number(b.getAttribute('data-row')||0);
                var act=String(b.getAttribute('data-pdc-act')||'');
                if(act==='cancel'){
                  if(!confirm('Annuler cette demande PDC ?')) return;
                  var reason=prompt('Motif (optionnel) ?')||'';
                  google.script.run.withSuccessHandler(function(){ renderLogistique(); }).withFailureHandler(showErr).pdcCancel({pdcRow:row,reason:reason});
                  return;
                }
                var payload={
                  pdcRow:row,
                  action:(act==='exp'?'expedier':(act==='liv'?'livrer':'save')),
                  poids:document.getElementById('pdcW_'+row).value,
                  transporteur:document.getElementById('pdcT_'+row).value,
                  notes:document.getElementById('pdcN_'+row).value,
                  colisPret:document.getElementById('pdcReady_'+row).checked,
                  pointDepart:document.getElementById('pdcFrom_'+row).value,
                  adresseArrivee:document.getElementById('pdcTo_'+row).value,
                  refs:document.getElementById('pdcRefs_'+row).value,
                  designation:document.getElementById('pdcDes_'+row).value
                };
                if(document.getElementById('pdcK_'+row)) payload.tracking=document.getElementById('pdcK_'+row).value;
                if(act==='liv'){
                  var dl=prompt('Date livraison (AAAA-MM-JJ) ?');
                  if(!dl) return;
                  payload.dateLivraison=dl;
                }
                google.script.run.withSuccessHandler(function(){ renderLogistique(); }).withFailureHandler(showErr).pdcSave(payload);
              };
            });
            document.querySelectorAll('input[data-pdc-files="1"]').forEach(function(inp){
              inp.onchange=function(){
                var m=String(inp.id||'').match(/^pdcF_(\d+)$/);
                if(!m) return;
                var row=Number(m[1]||0);
                var fl=(inp.files&&inp.files.length)?Array.from(inp.files):[];
                if(!fl.length) return;
                var res=document.getElementById('pdcR_'+row);
                if(res) res.textContent='Envoi...';
                uploadFiles_(fl, function(payloadFiles){
                  google.script.run.withSuccessHandler(function(u){
                    if(u&&u.ok){
                      if(res) res.textContent='OK — '+(u.created||0)+' fichier(s) archivé(s). '+(u.folderUrl?('Drive : '+u.folderUrl):'');
                      try{ inp.value=''; }catch(e){}
                    }else{
                      if(res) res.textContent='';
                      showErr((u&&u.message)?u.message:'Erreur upload');
                    }
                  }).withFailureHandler(showErr).pdcUploadDocs({pdcRow:row,files:payloadFiles});
                });
              };
            });
          }).pdcListPage({tab:(window.__pdcTab||'open'),limit:(window.__pdcLimit||50),offset:(window.__pdcOffset||0),year:(window.__pdcYear===undefined?new Date().getFullYear():window.__pdcYear)})
        : google.script.run.withSuccessHandler(function(rows){
        if(!rows||!rows.length){
          document.getElementById('mainBody').innerHTML='<div class="muted">Aucune demande en cours (toutes les lignes sont livrées, ou aucune demande transport).</div>';
          return;
        }
        // Regroupe par dossier (numéro SAV) pour séparer visuellement les demandes
        function groupByNumero_(list){
          var g={};
          (list||[]).forEach(function(x){
            if(!x) return;
            var n=String(x.numero||'').trim()||'—';
            if(!g[n]) g[n]=[];
            g[n].push(x);
          });
          var keys=Object.keys(g);
          keys.sort(function(a,b){ return String(b).localeCompare(String(a)); });
          return {keys:keys, groups:g};
        }

        var missingTrk=0;
        rows.forEach(function(x){ if(x && x.missingTracking) missingTrk++; });
        var html='';
        if(missingTrk){
          html += '<div class="evt" style="margin-bottom:14px;border-color:rgba(245,158,11,.45);background:rgba(245,158,11,.06)">';
          html += '<div style="font-weight:900">Alerte</div>';
          html += '<div class="muted">Tracking manquant sur <b>'+esc(missingTrk)+'</b> dossier(s) déjà <b>expédié(s)</b>. (Le tracking peut être ajouté après enlèvement.)</div>';
          html += '</div>';
        }

        var gb = groupByNumero_(rows);
        gb.keys.forEach(function(num){
          var items = gb.groups[num] || [];
          items.sort(function(a,b){
            var da=String((a&&a.dateDemande)||'');
            var db=String((b&&b.dateDemande)||'');
            if(db!==da) return db.localeCompare(da);
            return Number((b&&b.logRow)||0) - Number((a&&a.logRow)||0);
          });

          html += '<div class="evt" style="margin-bottom:14px;border-color:rgba(148,163,184,.45);background:rgba(2,6,23,.12)">';
          html +=   '<div class="row"><div class="name">'+esc(num)+'</div><span class="badge">'+esc(items.length)+' demande(s)</span></div>';
          html += '</div>';

          items.forEach(function(x){
          html += '<div class="evt" style="margin-bottom:14px">';
          html += '<div class="row"><div class="name">'+esc(x.typeEnvoi||'Transport')+'</div>';
          html += x.missingTracking ? '<span class="badge warn">Tracking manquant</span>' : '<span class="badge warn">'+esc(x.statutLog)+'</span>';
          if(String(x.colisPret||'').toUpperCase()==='OUI'){
            html += '<span class="badge ok">Colis prêt</span>';
          }else{
            html += '<span class="badge">Colis non prêt</span>';
          }
          html += '</div>';
          html += '<div class="muted">Créé le '+esc(x.dateDemande)+'</div>';
          html += '<div style="margin-top:6px;font-size:12px"><b>Destinataire / résumé</b> '+esc(x.resume)+'</div>';
          html += '<div style="font-size:12px"><b>Modèle</b> '+esc(x.modele)+'</div>';

          // Cadres séparés Arrivée / Enlèvement (évite quiproquos)
          var arr = String(x.adresseArrivee||'').trim();
          var enl = String(x.pointDepart||'').trim();
          // fallback pour les anciennes lignes
          if(!arr && String(x.typeEnvoi||'').toLowerCase().indexOf('enlèvement')>-1){
            arr = '34 rue du moulin des bruyeres, Courbevoie 92400';
          }
          html += '<div class="row" style="gap:10px;align-items:stretch;flex-wrap:wrap;margin-top:10px">';
          html +=   '<div style="flex:1;min-width:260px;padding:10px 12px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.03)">';
          html +=     '<div style="font-weight:900;margin-bottom:6px">Arrivée</div>';
          html +=     '<div style="white-space:pre-wrap;font-size:12px">'+esc(arr||'—')+'</div>';
          html +=   '</div>';
          html +=   '<div style="flex:1;min-width:260px;padding:10px 12px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.03)">';
          html +=     '<div style="font-weight:900;margin-bottom:6px">Enlèvement</div>';
          html +=     '<div style="white-space:pre-wrap;font-size:12px">'+esc(enl||'—')+'</div>';
          html +=   '</div>';
          html += '</div>';

          // Workflow:
          // - À traiter => saisie expédition (transporteur, coût, date expédition, notes) + validation expédition
          // - Expédié => tracking (sauvegardable), preuve + date livraison + validation livraison
          if(x.statutLog===LOG_A){
            var trVal = String(x.transporteur||'').trim();
            html += '<label>Transporteur (DISTRIBUTEUR = gratuit)</label>';
            html += '<select id="trp_'+x.logRow+'">';
            function opt(v,label){
              v=String(v||'');
              label=(label==null)?v:String(label);
              var sel = (String(trVal).toUpperCase()===String(v).toUpperCase()) ? ' selected' : '';
              return '<option value="'+esc(v)+'"'+sel+'>'+esc(label)+'</option>';
            }
            html += opt('DISTRIBUTEUR','DISTRIBUTEUR');
            html += opt('DPD','DPD');
            html += opt('La Poste (Colissimo)','La Poste (Colissimo)');
            html += opt('Chronopost','Chronopost');
            html += opt('UPS','UPS');
            html += opt('FedEx','FedEx');
            html += opt('DHL','DHL');
            html += opt('GLS','GLS');
            html += opt('TNT','TNT');
            html += opt('Mondial Relay','Mondial Relay');
            html += opt('Geodis','Geodis');
            html += opt('Kuehne+Nagel','Kuehne+Nagel');
            html += opt('XPO Logistics','XPO Logistics');
            html += opt('Colis Privé','Colis Privé');
            html += opt('So Colissimo','So Colissimo');
            html += opt('Hermes/Evri','Hermes/Evri');
            // Compat: si une ancienne valeur existe et n'est pas dans la liste, on la conserve
            if(trVal && ['DISTRIBUTEUR','DPD','LA POSTE (COLISSIMO)','CHRONOPOST','UPS','FEDEX','DHL','GLS','TNT','MONDIAL RELAY','GEODIS','KUEHNE+NAGEL','XPO LOGISTICS','COLIS PRIVÉ','SO COLISSIMO','HERMES/EVRI'].indexOf(String(trVal).toUpperCase())===-1){
              html += opt(trVal, trVal);
            }
            html += '</select>';
            html += '<label>Coût transport (€)</label><input type="number" step="0.01" id="ct_'+x.logRow+'" value="'+esc((x.coutTransport!=null)?String(x.coutTransport):'')+'" />';
            html += '<label>Date expédition</label><input type="date" id="dexp_'+x.logRow+'" value="'+esc(x.dateExpedition||'')+'" />';
            html += '<label>Notes logistique</label><textarea id="note_'+x.logRow+'">'+esc(x.notesLog)+'</textarea>';
          }else{
            // Affiche un récap des infos expédition (lecture seule) pour éviter les confusions
            var recap=[];
            if(x.transporteur) recap.push('Transporteur : '+esc(x.transporteur));
            if(x.coutTransport!=null && String(x.coutTransport)!=='') recap.push('Coût : '+esc(String(x.coutTransport))+'€');
            if(x.dateExpedition) recap.push('Date expédition : '+esc(x.dateExpedition));
            if(recap.length){
              html += '<div class="muted" style="margin-top:10px;font-size:12px"><b>Expédition</b> '+recap.join(' · ')+'</div>';
            }
          }
          if(x.statutLog===LOG_E){
            html += '<label>Tracking</label><input type="text" id="trk_'+x.logRow+'" value="'+esc(x.tracking||'')+'" />';
            html += '<div class="row" style="margin-top:8px"><button class="btn" type="button" data-act="save" data-row="'+x.logRow+'">Enregistrer tracking</button></div>';
            html += '<label class="row"><input type="checkbox" id="prf_'+x.logRow+'" '+(String(x.preuveLivraison||'')==='OUI'?'checked':'')+'/> Preuve livraison reçue</label>';
            html += '<label>Date livraison</label><input type="date" id="dlv_'+x.logRow+'" value="'+esc(x.dateLivraison||'')+'" />';
            html += '<div style="margin-top:10px;padding:10px 12px;border:1px dashed var(--border);border-radius:12px;background:rgba(255,255,255,.03)">';
            html +=   '<div style="font-weight:900;margin-bottom:6px">Déposer la preuve (photo/PDF)</div>';
            html +=   '<div class="muted" style="margin-bottom:8px">Glisse-dépose ici, ou sélectionne un fichier. Les fichiers seront enregistrés dans le dossier Drive du SAV.</div>';
            html +=   '<input type="file" id="pf_'+x.logRow+'" data-proof="1" multiple />';
            html += '</div>';
          }
          if(x.statutLog===LOG_A){
            html += '<div class="row" style="margin-top:10px;gap:10px;flex-wrap:wrap">';
            html +=   '<button class="btn primary" type="button" data-act="exp" data-row="'+x.logRow+'">Valider expédition</button>';
            html +=   '<button class="btn" type="button" data-act="del" data-row="'+x.logRow+'">Supprimer</button>';
            html += '</div>';
          } else if(x.statutLog===LOG_E){
            html += '<div class="row" style="margin-top:10px;gap:10px;flex-wrap:wrap">';
            html +=   '<button class="btn primary" type="button" data-act="liv" data-row="'+x.logRow+'">Valider livraison</button>';
            html +=   '<button class="btn" type="button" data-act="del" data-row="'+x.logRow+'">Supprimer</button>';
            html += '</div>';
          } else {
            html += '<div class="row" style="margin-top:10px"><button class="btn" type="button" data-act="del" data-row="'+x.logRow+'">Supprimer</button></div>';
          }
          html += '</div>';
          });
        });
        document.getElementById('mainBody').innerHTML=html;
        document.querySelectorAll('button[data-act="exp"]').forEach(function(b){
          b.onclick=function(){
            if(b.disabled) return;
            b.disabled=true;
            var old=b.textContent;
            b.textContent='...';
            var lr=Number(b.getAttribute('data-row'));
            google.script.run.withSuccessHandler(function(){ renderLogistique(); }).withFailureHandler(function(e){
              b.disabled=false;
              b.textContent=old;
              showErr(e);
            }).savSaveLogistique({
              logRow:lr,
              action:'expedier',
              transporteur:document.getElementById('trp_'+lr).value,
              coutTransport:document.getElementById('ct_'+lr).value,
              dateExpedition:document.getElementById('dexp_'+lr).value,
              notes:document.getElementById('note_'+lr).value
            });
          };
        });
        document.querySelectorAll('button[data-act="liv"]').forEach(function(b){
          b.onclick=function(){
            if(b.disabled) return;
            b.disabled=true;
            var old=b.textContent;
            b.textContent='...';
            var lr=Number(b.getAttribute('data-row'));
            google.script.run.withSuccessHandler(function(){ renderLogistique(); }).withFailureHandler(function(e){
              b.disabled=false;
              b.textContent=old;
              showErr(e);
            }).savSaveLogistique({
              logRow:lr,
              action:'livrer',
              preuveLivraison:document.getElementById('prf_'+lr).checked,
              dateLivraison:document.getElementById('dlv_'+lr).value
            });
          };
        });
        document.querySelectorAll('button[data-act="del"]').forEach(function(b){
          b.onclick=function(){
            var lr=Number(b.getAttribute('data-row'));
            if(!lr) return;
            if(!confirm('Supprimer cette demande logistique ?')) return;
            b.disabled=true;
            var old=b.textContent;
            b.textContent='...';
            google.script.run.withSuccessHandler(function(){ renderLogistique(); }).withFailureHandler(function(e){
              b.disabled=false;
              b.textContent=old;
              showErr(e);
            }).savDeleteLogistique({logRow:lr});
          };
        });

        document.querySelectorAll('button[data-act="save"]').forEach(function(b){
          b.onclick=function(){
            if(b.disabled) return;
            b.disabled=true;
            var old=b.textContent;
            b.textContent='...';
            var lr=Number(b.getAttribute('data-row'));
            google.script.run.withSuccessHandler(function(){ renderLogistique(); }).withFailureHandler(function(e){
              b.disabled=false;
              b.textContent=old;
              showErr(e);
            }).savSaveLogistique({
              logRow:lr,
              action:'save',
              tracking:document.getElementById('trk_'+lr).value
            });
          };
        });

        // Upload preuve livraison (statut expédié)
        document.querySelectorAll('input[data-proof="1"]').forEach(function(inp){
          inp.onchange=function(){
            var m = String(inp.id||'').match(/^pf_(\d+)$/);
            if(!m) return;
            var lr = Number(m[1]||0);
            if(!lr) return;
            var files = inp.files;
            if(!files || !files.length) return;
            inp.disabled=true;
            uploadFiles_(files, function(payloadFiles){
              google.script.run.withSuccessHandler(function(){
                inp.disabled=false;
                renderLogistique();
              }).withFailureHandler(function(e){
                inp.disabled=false;
                showErr(e);
              }).savUploadLogistiqueProof({logRow:lr, files:payloadFiles});
            });
          };
        });
      }).withFailureHandler(showErr).savListLogistique(true);

      // lance la requête
      if(fetch==='pdc') return;
    }

    function renderLogistiqueOther_(){
      document.getElementById('mainTitle').textContent='Logistique — Autres';
      var type = String(window.__logOtherType||'').trim().toUpperCase();
      if(type!=='ENVOI' && type!=='ARRIVÉE' && type!=='ARRIVEE'){
        var pick = prompt('Autres : ARRIVÉE ou ENVOI ?', 'ENVOI') || '';
        pick = String(pick||'').trim().toUpperCase();
        if(pick==='ARRIVEE') pick='ARRIVÉE';
        if(pick!=='ARRIVÉE' && pick!=='ENVOI'){ showErr('Choix invalide (ARRIVÉE ou ENVOI).'); return; }
        window.__logOtherType = pick;
        type = pick;
      }
      if(type==='ARRIVEE') type='ARRIVÉE';
      var top =
        '<div class="row" style="gap:10px;flex-wrap:wrap;margin-bottom:10px">'+
          '<button class="btn" type="button" id="logModeSav">Transport SAV</button>'+
          '<button class="btn" type="button" id="logModePdc">Pièces détachées</button>'+
          '<button class="btn primary" type="button" id="logModeOther">Autres</button>'+
        '</div>';
      var sub =
        '<div class="row" style="gap:10px;flex-wrap:wrap;margin:4px 0 12px 0">'+
          '<button class="btn '+(type==='ENVOI'?'primary':'')+'" type="button" id="otherTypeEnvoi">Envois</button>'+
          '<button class="btn '+(type==='ARRIVÉE'?'primary':'')+'" type="button" id="otherTypeArr">Arrivées</button>'+
          '<button class="btn primary" type="button" id="otherNew">+ Ajouter</button>'+
        '</div>';
      document.getElementById('mainBody').innerHTML = top +
        '<div class="evt" style="margin-bottom:14px">' +
          '<div style="font-weight:900">Autres — '+esc(type)+'</div>' +
          '<div class="muted">Traçabilité simple des colis passés “par chez nous” (hors dossiers SAV, sans workflow complet).</div>' +
        '</div>' +
        sub +
        '<div class="muted">Chargement...</div>';

      // nav
      var bSav=document.getElementById('logModeSav');
      var bPdc=document.getElementById('logModePdc');
      var bO=document.getElementById('logModeOther');
      if(bSav) bSav.onclick=function(){ window.__logMode='sav'; renderLogistique(); };
      if(bPdc) bPdc.onclick=function(){ window.__logMode='pdc'; renderLogistique(); };
      if(bO) bO.onclick=function(){ window.__logMode='other'; renderLogistique(); };
      var bE=document.getElementById('otherTypeEnvoi');
      var bA=document.getElementById('otherTypeArr');
      if(bE) bE.onclick=function(){ window.__logOtherType='ENVOI'; renderLogistiqueOther_(); };
      if(bA) bA.onclick=function(){ window.__logOtherType='ARRIVÉE'; renderLogistiqueOther_(); };

      var bNew=document.getElementById('otherNew');
      if(bNew){
        bNew.onclick=function(){
          var today = new Date();
          var y=today.getFullYear(), m=String(today.getMonth()+1).padStart(2,'0'), d=String(today.getDate()).padStart(2,'0');
          var date = prompt('Date (AAAA-MM-JJ) ?', (y+'-'+m+'-'+d)) || '';
          date = String(date||'').trim();
          var ref = prompt('Référence / client (optionnel) ?', '') || '';
          var pd = prompt(type==='ENVOI'?'Point de départ (optionnel) ?':'Provenance / point départ (optionnel) ?', '') || '';
          var aa = prompt(type==='ENVOI'?'Adresse arrivée (optionnel) ?':'Adresse arrivée (optionnel) ?', '') || '';
          var trp = prompt('Transporteur (optionnel) ?', '') || '';
          var trk = prompt('Tracking (optionnel) ?', '') || '';
          var notes = prompt('Notes (optionnel) ?', '') || '';
          google.script.run.withSuccessHandler(function(r){
            if(r&&r.ok){ renderLogistiqueOther_(); }
            else showErr((r&&r.message)?r.message:'Erreur création');
          }).withFailureHandler(showErr).savOtherLogCreate({
            sens:type,
            date:date,
            ref:ref,
            pointDepart:pd,
            adresseArrivee:aa,
            transporteur:trp,
            tracking:trk,
            notes:notes
          });
        };
      }

      google.script.run.withSuccessHandler(function(rows){
        rows = rows || [];
        var h = top +
          '<div class="evt" style="margin-bottom:14px">' +
            '<div style="font-weight:900">Autres — '+esc(type)+'</div>' +
            '<div class="muted">Traçabilité simple des colis passés “par chez nous”.</div>' +
          '</div>' +
          sub;
        if(!rows.length){
          h += '<div class="muted">Aucune ligne.</div>';
          document.getElementById('mainBody').innerHTML=h;
        }else{
          rows.forEach(function(x){
            h += '<div class="evt" style="margin-bottom:14px">';
            h += '<div class="row"><div class="name">'+esc(x.id)+'</div><span class="badge">'+esc(x.sens||type)+'</span></div>';
            h += '<div class="muted">Date : '+esc(x.date||'')+'</div>';
            h += '<label>Référence / client</label><input type="text" id="othRef_'+x.otherRow+'" value="'+esc(x.ref||'')+'" />';
            h += '<label>Point départ</label><textarea id="othPd_'+x.otherRow+'">'+esc(x.pointDepart||'')+'</textarea>';
            h += '<label>Adresse arrivée</label><textarea id="othAa_'+x.otherRow+'">'+esc(x.adresseArrivee||'')+'</textarea>';
            h += '<label>Transporteur</label><input type="text" id="othTrp_'+x.otherRow+'" value="'+esc(x.transporteur||'')+'" />';
            h += '<label>Tracking</label><input type="text" id="othTrk_'+x.otherRow+'" value="'+esc(x.tracking||'')+'" />';
            h += '<label>Notes</label><textarea id="othNotes_'+x.otherRow+'">'+esc(x.notes||'')+'</textarea>';
            h += '<div class="row" style="margin-top:10px;gap:10px;flex-wrap:wrap">';
            h +=   '<button class="btn primary" type="button" data-oth-act="save" data-row="'+x.otherRow+'">Enregistrer</button>';
            h +=   '<button class="btn" type="button" data-oth-act="del" data-row="'+x.otherRow+'">Supprimer</button>';
            h += '</div>';
            h += '</div>';
          });
          document.getElementById('mainBody').innerHTML=h;
        }

        // rebind nav + actions after render
        var bSav2=document.getElementById('logModeSav');
        var bPdc2=document.getElementById('logModePdc');
        var bO2=document.getElementById('logModeOther');
        if(bSav2) bSav2.onclick=function(){ window.__logMode='sav'; renderLogistique(); };
        if(bPdc2) bPdc2.onclick=function(){ window.__logMode='pdc'; renderLogistique(); };
        if(bO2) bO2.onclick=function(){ window.__logMode='other'; renderLogistique(); };
        var bE2=document.getElementById('otherTypeEnvoi');
        var bA2=document.getElementById('otherTypeArr');
        if(bE2) bE2.onclick=function(){ window.__logOtherType='ENVOI'; renderLogistiqueOther_(); };
        if(bA2) bA2.onclick=function(){ window.__logOtherType='ARRIVÉE'; renderLogistiqueOther_(); };
        var bNew2=document.getElementById('otherNew');
        if(bNew2) bNew2.onclick = bNew ? bNew.onclick : function(){};

        document.querySelectorAll('button[data-oth-act]').forEach(function(b){
          b.onclick=function(){
            var row=Number(b.getAttribute('data-row')||0);
            var act=String(b.getAttribute('data-oth-act')||'');
            if(!row) return;
            if(act==='del'){
              if(!confirm('Supprimer cette ligne ?')) return;
              google.script.run.withSuccessHandler(function(r){
                if(r&&r.ok) renderLogistiqueOther_();
                else showErr((r&&r.message)?r.message:'Erreur suppression');
              }).withFailureHandler(showErr).savOtherLogDelete({otherRow:row});
              return;
            }
            google.script.run.withSuccessHandler(function(r){
              if(r&&r.ok) renderLogistiqueOther_();
              else showErr((r&&r.message)?r.message:'Erreur');
            }).withFailureHandler(showErr).savOtherLogSave({
              otherRow:row,
              ref:document.getElementById('othRef_'+row).value,
              pointDepart:document.getElementById('othPd_'+row).value,
              adresseArrivee:document.getElementById('othAa_'+row).value,
              transporteur:document.getElementById('othTrp_'+row).value,
              tracking:document.getElementById('othTrk_'+row).value,
              notes:document.getElementById('othNotes_'+row).value
            });
          };
        });
      }).withFailureHandler(showErr).savOtherLogList({sens:type});
    }

    function renderMarketplace(){
      document.getElementById('mainTitle').textContent='Marketplace — suivi & messages client';
      document.getElementById('mainBody').innerHTML='<div class="muted">Chargement...</div>';
      google.script.run.withSuccessHandler(function(rows){
        rows=rows||[];
        if(!rows.length){
          document.getElementById('mainBody').innerHTML='<div class="muted">Aucun dossier Marketplace.</div>';
          return;
        }
        // tri par date desc
        rows.sort(function(a,b){ return String(b.dateCreation||'').localeCompare(String(a.dateCreation||'')); });

        var listHtml='<div class="list" style="max-height:260px">';
        rows.forEach(function(c){
          listHtml+='<div class="item" data-num="'+esc(c.numero)+'"><b>'+esc(c.numero)+'</b> · '+esc(c.marketplace)+'<div class="muted">'+esc(c.clientEmail||'')+' · '+esc(c.etat||'')+'</div></div>';
        });
        listHtml+='</div>';

        var pane=
          '<div class="row" style="align-items:flex-start;gap:14px;flex-wrap:wrap">'+
            '<div style="flex:1;min-width:320px">'+
              '<div class="muted" style="margin-bottom:8px">Dossiers Marketplace (ouverts)</div>'+
              listHtml+
            '</div>'+
            '<div style="flex:2;min-width:340px">'+
              '<div id="mpDetail" class="muted">Sélectionne un dossier pour voir le fil.</div>'+
            '</div>'+
          '</div>';

        document.getElementById('mainBody').innerHTML=pane;
        document.querySelectorAll('.item[data-num]').forEach(function(el){
          el.onclick=function(){
            var num=el.getAttribute('data-num');
            var c=findByNumero_(rows, String(num));
            if(c) openMpThread(c);
          };
        });
      }).withFailureHandler(showErr).savMpListCases(true);
    }

    function renderAvoirTickets(){
      document.getElementById('mainTitle').textContent='Avoir — tickets à traiter';
      document.getElementById('mainBody').innerHTML='<div class="muted">Chargement...</div>';
      google.script.run.withSuccessHandler(function(groups){
        groups=groups||[];
        if(!groups.length){
          document.getElementById('mainBody').innerHTML='<div class="muted">Aucun ticket avoir en attente.</div>';
          return;
        }
        var html='';
        groups.forEach(function(g){
          html+='<div class="evt" style="margin-bottom:14px">';
          html+='<div class="row"><div class="name">'+esc(g.centraleAchat||'—')+'</div><span class="badge">'+esc((g.items||[]).length)+'</span></div>';
          (g.items||[]).forEach(function(x){
            var safe=String(x.numero||'').replace(/[^a-zA-Z0-9_-]/g,'_');
            html+='<div style="margin-top:10px;padding:10px 12px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.03)">';
            html+='<div class="row"><div style="font-weight:900">'+esc(x.numero)+'</div><span class="badge warn">'+esc(x.statut||'À faire')+'</span></div>';
            html+='<div class="muted">Créé le '+esc(x.dateCreation||'')+(x.driveUrl?(' · <a href="'+esc(x.driveUrl)+'" target="_blank">Dossier Drive</a>'):'')+'</div>';
            if(x.magasin && String(x.magasin||'').trim() && String(x.magasin||'').trim()!=='—'){
              html+='<div style="margin-top:6px;font-size:12px"><b>Magasin</b> '+esc(x.magasin)+'</div>';
            }
            if(x.modele||x.panne||x.serie){
              html+='<div style="margin-top:6px;font-size:12px"><b>Modèle</b> '+esc(x.modele||'—')+' · <b>Panne</b> '+esc(x.panne||'—')+' · <b>Série</b> '+esc(x.serie||'—')+'</div>';
            }
            if(x.adresse){
              html+='<div style="margin-top:6px;font-size:12px"><b>Adresse</b> '+esc(x.adresse)+'</div>';
            }
            if(x.centraleAchat){
              html+='<div style="margin-top:6px;font-size:12px"><b>Centrale d\\'achat</b> '+esc(x.centraleAchat)+'</div>';
            }
            if(x.factureAchatCentrale){
              html+='<div style="margin-top:4px;font-size:12px"><b>Facture achat centrale</b> '+esc(x.factureAchatCentrale)+'</div>';
            }
            if(x.notes){
              html+='<div style="margin-top:6px;font-size:12px"><b>Notes</b> '+esc(x.notes)+'</div>';
            }

            html+='<div style="margin-top:10px;padding:10px 12px;border:1px dashed rgba(37,99,235,.35);border-radius:12px;background:rgba(37,99,235,.03)">';
            html+='<div style="font-weight:900;margin-bottom:6px">Coller l\\'avoir (PDF/scan) dans le dossier</div>';
            html+='<div class="muted">Glisse le PDF/scans de l\\'avoir, il sera enregistré dans le dossier Drive du dossier SAV.</div>';
            html+='<div class="row" style="gap:10px;align-items:center;flex-wrap:wrap;margin-top:10px">';
            html+='<div id="avtDrop_'+safe+'" style="flex:1;min-width:220px;padding:10px 12px;border:1px dashed rgba(37,99,235,.35);border-radius:12px;background:rgba(37,99,235,.04);color:var(--muted);font-size:12px;cursor:pointer">Glisser l\\'avoir ici (ou cliquer)</div>';
            html+='<input type="file" id="avtFiles_'+safe+'" multiple class="hide" />';
            html+='<button class="btn primary" type="button" data-avt-up="1" data-safe="'+safe+'" data-num="'+esc(x.numero)+'" data-sheet="'+esc(x.sheet)+'" data-row="'+esc(x.row)+'" data-ticket-row="'+esc(x.ticketRow||0)+'">Archiver</button>';
            html+='<button class="btn" type="button" data-avt-done="1" data-ticket-row="'+esc(x.ticketRow||0)+'">Valider avoir fait</button>';
            html+='<button class="btn" type="button" data-avt-del="1" data-ticket-row="'+esc(x.ticketRow||0)+'">Supprimer demande</button>';
            html+='</div>';
            html+='<div class="muted" id="avtRes_'+safe+'" style="margin-top:8px"></div>';
            html+='</div>';

            html+='</div>';
          });
          html+='</div>';
        });
        document.getElementById('mainBody').innerHTML=html;

        (groups||[]).forEach(function(g){
          (g.items||[]).forEach(function(x){
            var safe=String(x.numero||'').replace(/[^a-zA-Z0-9_-]/g,'_');
            bindDropZone_('avtDrop_'+safe, 'avtFiles_'+safe);
          });
        });

        document.querySelectorAll('button[data-avt-up="1"]').forEach(function(b){
          b.onclick=function(){
            var safe=b.getAttribute('data-safe');
            var numero=b.getAttribute('data-num');
            var sheet=b.getAttribute('data-sheet');
            var row=Number(b.getAttribute('data-row')||0);
            var ticketRow=Number(b.getAttribute('data-ticket-row')||0);
            var input=document.getElementById('avtFiles_'+safe);
            var fl=(input&&input.files)?Array.from(input.files):[];
            if(!fl.length){ showErr('Aucun fichier sélectionné.'); return; }
            var resEl=document.getElementById('avtRes_'+safe);
            if(resEl) resEl.textContent='Envoi...';
            uploadFiles_(fl, function(payloadFiles){
              google.script.run.withSuccessHandler(function(u){
                if(u&&u.ok){
                  if(resEl) resEl.textContent='OK — '+(u.created||0)+' fichier(s) archivé(s). '+(u.folderUrl?('Drive : '+u.folderUrl):'');
                  try{ input.value=''; }catch(e){}
                  bindDropZone_('avtDrop_'+safe, 'avtFiles_'+safe);
                  google.script.run.withSuccessHandler(function(){ renderAvoirTickets(); }).withFailureHandler(showErr).savSetAvoirTicketStatut({ticketRow:ticketRow,statut:'Terminé'});
                }else{
                  showErr((u&&u.message)?u.message:'Erreur upload');
                  if(resEl) resEl.textContent='';
                }
              }).withFailureHandler(showErr).savUploadAttachments({sheet:sheet,row:row,numero:numero,files:payloadFiles});
            });
          };
        });

        document.querySelectorAll('button[data-avt-done="1"]').forEach(function(b){
          b.onclick=function(){
            var ticketRow=Number(b.getAttribute('data-ticket-row')||0);
            if(!ticketRow) return;
            if(!confirm('Valider que l\\'avoir est fait (ticket terminé) ?')) return;
            google.script.run.withSuccessHandler(function(){
              renderAvoirTickets();
            }).withFailureHandler(showErr).savSetAvoirTicketStatut({ticketRow:ticketRow,statut:'Terminé'});
          };
        });

        document.querySelectorAll('button[data-avt-del="1"]').forEach(function(b){
          b.onclick=function(){
            var ticketRow=Number(b.getAttribute('data-ticket-row')||0);
            if(!ticketRow) return;
            if(!confirm('Supprimer cette demande d\\'avoir ?')) return;
            google.script.run.withSuccessHandler(function(res){
              if(res && res.ok){
                renderAvoirTickets();
              }else{
                showErr((res&&res.message)?res.message:'Erreur suppression');
              }
            }).withFailureHandler(showErr).savDeleteAvoirTicket({ticketRow:ticketRow});
          };
        });
      }).withFailureHandler(showErr).savListAvoirTicketsPendingByDistributor();
    }

    function openMpThread(c){
      var box=document.getElementById('mpDetail');
      if(!box) return;
      box.innerHTML='<div class="muted">Chargement du fil...</div>';
      google.script.run.withSuccessHandler(function(msgs){
        msgs=msgs||[];
        var h='';
        h+='<div style="font-weight:900;margin-bottom:6px">Dossier '+esc(c.numero)+'</div>';
        h+='<div class="muted">Marketplace : '+esc(c.marketplace)+' · Client : '+esc(c.clientEmail||'')+' · État : '+esc(c.etat||'')+'</div>';
        h+='<div style="margin-top:10px;padding:10px 12px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.03)">';
        h+='<div style="font-weight:900;margin-bottom:8px">Fil de messages</div>';
        if(!msgs.length){
          h+='<div class="muted">Aucun message pour ce dossier.</div>';
        }else{
          msgs.forEach(function(m){
            var out=(String(m.sens||'')==='OUT');
            h+='<div style="margin-bottom:10px;padding:10px 12px;border-radius:12px;border:1px solid '+(out?'rgba(34,197,94,.35)':'rgba(245,158,11,.35)')+';background:'+(out?'rgba(34,197,94,.08)':'rgba(245,158,11,.08)')+'">';
            h+='<div style="font-size:12px;font-weight:900">'+(out?'À transmettre au client':'Message du client')+'</div>';
            h+='<div class="muted" style="margin-top:2px">'+esc(m.date||'')+' · '+esc(m.auteur||'')+'</div>';
            h+='<div style="margin-top:8px;white-space:pre-wrap;font-size:13px">'+esc(m.message||'')+'</div>';
            h+='</div>';
          });
        }
        h+='</div>';

        h+='<div style="margin-top:12px;padding:10px 12px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.03)">';
        h+='<div style="font-weight:900;margin-bottom:8px">Ajouter un message</div>';
        h+='<label>Sens</label><select id="mpSens"><option value="OUT">À transmettre au client</option><option value="IN">Message reçu du client</option></select>';
        h+='<label>Message</label><textarea id="mpMsg" placeholder="Saisis le message…"></textarea>';
        h+='<label>Auteur (optionnel)</label><input id="mpAuth" placeholder="ex: Prénom / Équipe" />';
        h+='<div class="row" style="margin-top:10px"><button class="btn primary" type="button" id="mpSend">Enregistrer</button></div>';
        h+='<div class="muted" id="mpRes" style="margin-top:8px"></div>';
        h+='</div>';

        h+='<div style="margin-top:12px;padding:10px 12px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.03)">';
        h+='<div style="font-weight:900;margin-bottom:8px">Archivage documents</div>';
        h+='<div class="muted">Glisse-dépose les documents à archiver pour ce dossier (ils seront enregistrés dans Drive).</div>';
        h+='<div class="row" style="gap:10px;align-items:center;flex-wrap:wrap;margin-top:10px">';
        h+='<div id="mpDrop" style="flex:1;min-width:220px;padding:10px 12px;border:1px dashed rgba(37,99,235,.35);border-radius:12px;background:rgba(37,99,235,.04);color:var(--muted);font-size:12px;cursor:pointer">Glisser vos fichiers ici (ou cliquer)</div>';
        h+='<input type="file" id="mpFiles" multiple class="hide" />';
        h+='<button class="btn primary" type="button" id="mpUp">Archiver</button>';
        h+='</div>';
        h+='<div class="muted" id="mpUpRes" style="margin-top:8px"></div>';
        h+='</div>';

        box.innerHTML=h;
        bindDropZone_('mpDrop','mpFiles');
        document.getElementById('mpSend').onclick=function(){
          var payload={
            numero:c.numero,
            marketplace:c.marketplace,
            clientEmail:c.clientEmail,
            sens:document.getElementById('mpSens').value,
            message:document.getElementById('mpMsg').value,
            auteur:document.getElementById('mpAuth').value,
            sheet:c.sheet,
            rowMain:c.row
          };
          google.script.run.withSuccessHandler(function(res){
            if(res&&res.ok){
              document.getElementById('mpRes').textContent='OK — message enregistré.';
              openMpThread(c);
              try{ refreshTabBadges_(); }catch(e){}
            }else{
              showErr((res&&res.message)?res.message:'Erreur');
            }
          }).withFailureHandler(showErr).savMpAddMessage(payload);
        };

        document.getElementById('mpUp').onclick=function(){
          var input=document.getElementById('mpFiles');
          var fl=(input&&input.files)?Array.from(input.files):[];
          if(!fl.length){ showErr('Aucun fichier sélectionné.'); return; }
          document.getElementById('mpUpRes').textContent='Envoi...';
          uploadFiles_(fl, function(payloadFiles){
            google.script.run.withSuccessHandler(function(u){
              if(u&&u.ok){
                document.getElementById('mpUpRes').textContent='OK — '+(u.created||0)+' fichier(s) archivé(s). '+(u.folderUrl?('Drive : '+u.folderUrl):'');
                // reset input + zone
                try{ input.value=''; }catch(e){}
                bindDropZone_('mpDrop','mpFiles');
              }else{
                showErr((u&&u.message)?u.message:'Erreur upload');
                document.getElementById('mpUpRes').textContent='';
              }
            }).withFailureHandler(showErr).savUploadAttachments({sheet:c.sheet,row:c.row,numero:c.numero,files:payloadFiles});
          });
        };
      }).withFailureHandler(showErr).savMpListMessages(c.numero);
    }

    function renderNewForm(){
      document.getElementById('mainTitle').textContent='Nouvelle demande';
      // Robustesse + complétude : on recharge une fois la liste des modèles depuis le serveur
      // (utile si le payload initial est incomplet ou si la liste est très longue).
      if(!modelsLoadedOnce && !(models && models.length)){
        document.getElementById('mainBody').innerHTML='<div class="muted">Chargement de la liste des modèles...</div>';
        google.script.run.withSuccessHandler(function(list){
          models=(list||[]).slice().sort();
          modelsLoadedOnce=true;
          renderNewForm();
        }).withFailureHandler(showErr).savGetModels();
        return;
      }
      var opts=models.map(function(m){return '<option value="'+esc(m)+'">'+esc(m)+'</option>';}).join('');
      var distList=[
        'Leroy Merlin',
        'Castorama',
        'Brico Dépôt',
        'Bricomarché',
        'OBI',
        'Mr Bricolage',
        'Intermarché',
        'Carrefour',
        'Auchan',
        'E.Leclerc',
        'Darty',
        'Amazon',
        'Cdiscount',
        'La Redoute'
      ];
      var distOpts = '<option value="">—</option>' + distList.map(function(n){ return '<option value="'+esc(n)+'">'+esc(n)+'</option>'; }).join('') + '<option value="__autre__">Autre...</option>';
      var mpList=[
        'Boulanger',
        'Darty',
        'Fnac',
        'Cdiscount',
        'Manomano',
        'Leroy Merlin',
        'Castorama',
        'Bricomarché',
        'Mr Bricolage',
        'Rueducommerce',
        'Conforama',
        'But',
        'Ikea',
        'Amazon',
        'eBay'
      ];
      var mpOpts = '<option value="">—</option>' + mpList.map(function(n){ return '<option value="'+esc(n)+'">'+esc(n)+'</option>'; }).join('') + '<option value="__autre__">Autre...</option>';
      var PANNE_DATA={
        "CLIMATISEURS":{
          "Pannes électriques":[
            "Pas de mise sous tension","Absence d'alimentation électrique","Fusible grillé","Disjoncteur déclenché","Câble d'alimentation endommagé","Prise électrique défectueuse","Court-circuit à la mise sous tension","Défaut d'isolement électrique","Courant de fuite anormal","Disjoncteur différentiel qui saute"
          ],
          "Cartes électroniques":[
            "Carte de puissance défaillante","Carte de commande HS","Condensateur électrolytique explosé","Condensateur gonflé","Soudure froide sur PCB","Circuit imprimé endommagé ou fissurée","Composant brûlé ou noirci","Transistor grillé","MOSFET défaillant","Diode grillée","Relais de puissance collé","Relais qui ne ferme pas","Module de puissance HS","Convertisseur DC/DC défaillant","Alimentation 12V absente"
          ],
          "Compresseur":[
            "Compresseur ne démarre pas","Compresseur bloqué","Compresseur grillé","Compresseur qui grince","Compresseur qui fait du bruit anormal","Bobine moteur en court-circuit","Arbre compresseur bloqué","Compresseur thermiquement bloqué"
          ],
          "Circuit frigorifique":[
            "Fuite de gaz réfrigérant","Perte de fluide frigorigène","Fuite mineure (micro-fuite)","Fuite majeure (tuyau percé)","Fuite sur soudure","Fuite au raccordement","Perte progressive de gaz","Absence de gaz réfrigérant","Charge insuffisante de gaz","Tuyau frigorifique gelé","Circuit obstrué","Tuyau écrasé ou plié","Présence d'eau dans circuit","Humidité excessive du circuit","Vanne d'expansion électronique bloquée"
          ],
          "Refroidissement/Performance":[
            "Climatiseur ne refroidit pas","Refroidissement insuffisant","Perte progressive de froid","Température intérieure ne baisse plus","Froid seulement au démarrage","Évaporateur gelé","Accumulation de givre","Condenseur externe encrassé","Ventilateur condenseur ne tourne pas","Perte d'efficacité du condenseur"
          ],
          "Ventilateur intérieur":[
            "Ventilateur ne tourne pas","Ventilateur qui tourne lentement","Ventilateur qui s'arrête et redémarre","Ventilateur qui fait du bruit","Moteur ventilateur défaillant","Pale cassée","Pale fissurée","Pale déformée","Déséquilibre des pales","Roulement ventilateur usé"
          ],
          "Drainage":[
            "Eau qui s'accumule dans l'appareil","Fuite d'eau au sol","Tuyau de drainage obstrué","Bac de collecte plein","Pente incorrecte du tuyau","Tuyau de drainage gelé","Joint de tuyau défaillant","Accumulation de condensation"
          ],
          "Affichage/Interface":[
            "Écran LED ne s'allume pas","Affichage clignotant","Affichage flou","Segments LED manquants","Télécommande ne répond pas","Boutons ne réagissent pas","Affichage de caractères bizarres","Écran qui s'éteint après quelques secondes"
          ],
          "Thermostat":[
            "Thermostat défaillant","Mauvaise régulation de température","Décalage de température","Thermostat qui bascule en continu","Capteur de température cassé","Sonde de température détachée","Appareil qui fait des cycles courts","Surcharge thermique activée"
          ],
          "Bruits/Vibrations":[
            "Sifflement anormal","Cliquetis métallique","Claquement sec","Grincement","Couinement aigu","Bourdonnement excessif","Bruit de frottement","Bruit de vibration"
          ]
        },
        "RADIATEURS":{
          "Pannes électriques":[
            "Pas de mise sous tension","Fusible grillé","Disjoncteur déclenché","Câble endommagé","Prise défectueuse","Court-circuit","Défaut d'isolement","Fuite de courant"
          ],
          "Cartes électroniques":[
            "Carte de puissance défaillante","Carte de commande HS","Condensateur explosé","Condensateur gonflé","Soudure froide","Circuit imprimé fissurée","Composant brûlé","Transistor grillé","Diode défaillante","Relais collé"
          ],
          "Résistance chauffante":[
            "Résistance cassée","Résistance grillée","Résistance en circuit ouvert","Absence complète de chaleur","Chaleur très faible","Chaleur intermittente","Mauvais contact résistance","Soudure de résistance cassée"
          ],
          "Ventilateur (soufflant)":[
            "Ventilateur ne tourne pas","Ventilateur tourne lentement","Ventilateur s'arrête","Ventilateur bruyant","Moteur défaillant","Pale cassée","Pale déformée","Déséquilibre des pales"
          ],
          "Thermostat":[
            "Thermostat défaillant","Mauvaise régulation thermique","Température ne baisse pas","Thermostat qui bascule","Capteur température cassé","Sonde thermique détachée","Protection thermique bloquée","Radiateur qui s'éteint rapidement"
          ],
          "Sécurité/Protection":[
            "Protection thermique activée","Thermostat de sécurité déclenché","Surchauffe détectée","Capteur surchauffe bloqué","Disjoncteur thermique déclenché","Arrêt sécurité","Protection qui ne réarme pas"
          ],
          "Bruits":[
            "Cliquetis","Claquement au démarrage","Bruit de dilatation","Crépitement électrique","Grincement moteur","Vibrations anormales","Bruit de pulsation"
          ],
          "Mécanique (mobile)":[
            "Roulettes cassées","Roulettes bloquées","Socle instable","Radiateur penche","Poignée cassée","Châssis déformé","Vis desserrées"
          ]
        },
        "VENTILATEURS":{
          "Pannes électriques":[
            "Pas d'alimentation","Fusible grillé","Disjoncteur déclenché","Câble endommagé","Prise défectueuse","Court-circuit","Défaut d'isolement"
          ],
          "Cartes électroniques":[
            "Carte de puissance HS","Carte de commande défaillante","Condensateur explosé","Condensateur gonflé","Soudure froide","Composant brûlé","Transistor grillé","Diode défaillante","Relais collé"
          ],
          "Moteur":[
            "Moteur ne démarre pas","Moteur très lentement","Moteur par à-coups","Moteur qui s'arrête et redémarre","Moteur bloqué","Moteur grillé","Bobine moteur en circuit ouvert","Arbre moteur tordu","Roulement usé","Roulement grippé"
          ],
          "Pales":[
            "Pale cassée","Pale fissurée","Pale tordue","Pale pliée","Pale déformée","Pale qui vibre","Pale qui frotte","Pale endommagée"
          ],
          "Vitesse/Régulation":[
            "Régulateur défaillant","Commutateur ne répond pas","Variateur ne fonctionne pas","Une seule vitesse fonctionne","Pas d'augmentation de vitesse","Vitesse maximale impossible","Vitesse minimale impossible","Rhéostat défaillant"
          ],
          "Oscillation/Pivot":[
            "Oscillation bloquée","Oscillation lente","Ventilateur ne pivote pas","Moteur d'oscillation défaillant","Engrenage usé","Engrenage cassé","Chaîne lâche","Chaîne qui saute"
          ],
          "Mécanique (mobile)":[
            "Roulettes cassées","Roulettes bloquées","Socle instable","Ventilateur penche","Poignée cassée","Tirant défaillant","Châssis déformé","Vis desserrées"
          ],
          "Support/Fixation (fixe)":[
            "Support mural cassé","Fixation murale HS","Vis manquantes","Chevilles arrachées","Support qui descend","Boulons desserrés","Plaque de support cassée","Bras extensible cassé"
          ],
          "Grillage":[
            "Grille cassée","Grille déformée","Grille qui vibre","Grille desserrée","Rouille de grille","Grille mal fixée","Grille endommagée","Trous dans grille"
          ],
          "Bruits":[
            "Bruit excessif","Vibrations anormales","Grincement","Crissement","Couinement aigu","Bourdonnement","Cliquetis","Claquement"
          ]
        }
      };
      document.getElementById('mainBody').innerHTML=
        '<div class="row" style="margin-bottom:10px">'+
        '<label style="margin:0"><input type="radio" name="br" value="dist" checked /> Distributeur</label>'+
        '<label style="margin:0"><input type="radio" name="br" value="mp" /> Marketplace</label>'+
        '<label style="margin:0"><input type="radio" name="br" value="pdc" /> Pièces détachées</label></div>'+
        '<div id="formDist">'+
        '<label>Magasin *</label><select id="dMag">'+distOpts+'</select>'+
        '<div id="dMagOtherWrap" class="hide"><label>Autre magasin *</label><input id="dMagOther" placeholder="Saisir le nom du distributeur..." /></div>'+
        '<label>Référence dossier distributeur</label><input id="dRef" placeholder="Référence interne / centrale / commande..." />'+
        '<label>Adresse *</label><textarea id="dAdr"></textarea>'+
        '<label>Modèle *</label>'+
        '<div class="row" style="gap:8px;align-items:stretch">'+
          '<select id="dMod" style="flex:1"><option value="">—</option>'+opts+'</select>'+
          '<button class="btn" type="button" id="btnReloadModels" style="white-space:nowrap">Rafraîchir</button>'+
        '</div>'+
        '<div id="procBox" class="hide" style="margin-top:10px;padding:10px 12px;border:1px solid rgba(56,189,248,.35);background:rgba(56,189,248,.08);border-radius:12px"></div>'+
        '<label>Type appareil *</label><select id="dApp"><option value="">—</option><option value="CLIMATISEURS">🔵 Climatiseurs</option><option value="RADIATEURS">🟡 Radiateurs</option><option value="VENTILATEURS">🟢 Ventilateurs</option></select>'+
        '<label>Famille panne *</label><select id="dPanType"><option value="">—</option></select>'+
        '<label>Panne *</label><select id="dPanSel"><option value="">—</option></select>'+
        '<div id="dPanOtherWrap" class="hide"><label>Autre panne *</label><textarea id="dPanOther" placeholder="Décrire la panne..."></textarea></div>'+
        '<label>N° série *</label><input id="dSer" />'+
        '<div class="row" style="gap:16px;align-items:center;flex-wrap:wrap">'+
          '<label style="margin:0;display:flex;align-items:center;gap:8px;font-size:13px;font-weight:800;color:var(--text);text-transform:none;letter-spacing:0" for="dFac">'+
            '<input type="checkbox" id="dFac" style="width:auto;margin:0" />'+
            'Facture reçue'+
          '</label>'+
        '</div>'+
        '<div class="row" style="gap:16px;align-items:center;flex-wrap:wrap;margin-top:8px">'+
          '<label style="margin:0;display:flex;align-items:center;gap:8px;font-size:13px;font-weight:800;color:var(--text);text-transform:none;letter-spacing:0" for="dPho">'+
            '<input type="checkbox" id="dPho" style="width:auto;margin:0" />'+
            'Photo plaque'+
          '</label>'+
          '<div id="dDrop" style="flex:1;min-width:220px;padding:10px 12px;border:1px dashed rgba(37,99,235,.35);border-radius:12px;background:rgba(37,99,235,.04);color:var(--muted);font-size:12px;cursor:pointer">Glisser vos fichiers ici (ou cliquer)</div>'+
          '<input type="file" id="dFiles" multiple class="hide" />'+
        '</div>'+
        '</div>'+
        '<div id="formMp" class="hide">'+
        '<label>Marketplace *</label><select id="mMp">'+mpOpts+'</select>'+
        '<label>Nom client *</label><input id="mNom" />'+
        '<label>Email *</label><input id="mMail" type="email" />'+
        '<label>Téléphone *</label><input id="mTel" />'+
        '<label>Adresse client *</label><textarea id="mAdr" placeholder="Adresse complète (rue, ville, code postal, pays)"></textarea>'+
        '<div class="row" style="gap:16px;align-items:center;flex-wrap:wrap;margin-top:8px">'+
          '<label style="margin:0;display:flex;align-items:center;gap:8px;font-size:13px;font-weight:800;color:var(--text);text-transform:none;letter-spacing:0" for="mEnl">'+
            '<input type="checkbox" id="mEnl" style="width:auto;margin:0" />'+
            'Demande d\\'enlèvement'+
          '</label>'+
        '</div>'+
        '<label>Modèle *</label>'+
        '<div class="row" style="gap:8px;align-items:stretch">'+
          '<select id="mMod" style="flex:1"><option value="">—</option>'+opts+'</select>'+
          '<button class="btn" type="button" id="btnReloadModelsMp" style="white-space:nowrap">Rafraîchir</button>'+
        '</div>'+
        '<label>Type appareil *</label><select id="mApp"><option value="">—</option><option value="CLIMATISEURS">🔵 Climatiseurs</option><option value="RADIATEURS">🟡 Radiateurs</option><option value="VENTILATEURS">🟢 Ventilateurs</option></select>'+
        '<label>Famille panne *</label><select id="mPanType"><option value="">—</option></select>'+
        '<label>Panne *</label><select id="mPanSel"><option value="">—</option></select>'+
        '<div id="mPanOtherWrap" class="hide"><label>Autre panne *</label><textarea id="mPanOther" placeholder="Décrire la panne..."></textarea></div>'+
        '<label>N° facture *</label><input id="mFac" />'+
        '<div class="row" style="gap:16px;align-items:center;flex-wrap:wrap;margin-top:8px">'+
          '<label style="margin:0;display:flex;align-items:center;gap:8px;font-size:13px;font-weight:800;color:var(--text);text-transform:none;letter-spacing:0" for="mPho">'+
            '<input type="checkbox" id="mPho" style="width:auto;margin:0" />'+
            'Photo N°S'+
          '</label>'+
          '<div id="mDrop" style="flex:1;min-width:220px;padding:10px 12px;border:1px dashed rgba(37,99,235,.35);border-radius:12px;background:rgba(37,99,235,.04);color:var(--muted);font-size:12px;cursor:pointer">Glisser vos fichiers ici (ou cliquer)</div>'+
          '<input type="file" id="mFiles" multiple class="hide" />'+
        '</div>'+
        '</div>'+
        '<div id="formPdc" class="hide">'+
        '<div class="muted">Créer une demande d\\'envoi de pièces détachées (hors dossiers SAV). Le suivi se fait uniquement dans Logistique.</div>'+
        '<div style="margin-top:10px;padding:10px 12px;border:1px solid rgba(37,99,235,.25);border-radius:12px;background:rgba(37,99,235,.04)">'+
          '<div style="font-weight:900;margin-bottom:6px">Procédure technicien (demande PDC)</div>'+
          '<div class="muted">'+
            '1) Renseigner le point de départ + l’adresse d’arrivée.<br>'+
            '2) (Optionnel) Cocher “Rattacher à un dossier SAV” puis saisir le numéro SAV : le logiciel vérifie le dossier avant enregistrement.<br>'+
            '3) Ajouter les pièces : 1 ligne = 1 référence + 1 désignation (bouton “Ajouter une pièce”).<br>'+
            '4) Cocher “Colis prêt” si la pièce est disponible.<br>'+
            '5) Renseigner poids + transporteur + notes, puis enregistrer : tu obtiens un numéro PDC. Le suivi se fait dans Logistique → Pièces détachées (expédition/livraison + documents).'+
          '</div>'+
        '</div>'+
        '<label>Point de départ</label><textarea id="pdcFrom" placeholder="Adresse / service expéditeur..."></textarea>'+
        '<label>Adresse d\\'arrivée</label><textarea id="pdcTo" placeholder="Adresse complète destinataire..."></textarea>'+
        '<label style="margin-top:10px" class="row"><input type="checkbox" id="pdcLinkCase" /> Rattacher à un dossier SAV</label>'+
        '<div id="pdcLinkWrap" class="hide">'+
          '<label>Numéro de dossier SAV *</label><input id="pdcCaseNum" placeholder="ex: SAV-2026-001" />'+
          '<div class="muted" id="pdcCaseCheck" style="margin-top:6px"></div>'+
        '</div>'+
        '<div class="muted" style="font-weight:950;margin:10px 0 6px 2px">Pièces</div>'+
        '<div id="pdcPieces"></div>'+
        '<div class="row" style="gap:10px;flex-wrap:wrap;margin-top:8px">'+
          '<button class="btn" type="button" id="pdcAddPiece">+ Ajouter une pièce</button>'+
        '</div>'+
        '<label style="margin-top:10px" class="row"><input type="checkbox" id="pdcReady" /> Colis prêt à l’envoi</label>'+
        '<label>Poids (kg)</label><input id="pdcPoids" type="number" step="0.01" placeholder="ex: 1.25" />'+
        '<label>Transporteur</label><select id="pdcTrp">'+
          '<option value="DISTRIBUTEUR">DISTRIBUTEUR</option>'+
          '<option value="DPD">DPD</option>'+
          '<option value="La Poste (Colissimo)">La Poste (Colissimo)</option>'+
          '<option value="Chronopost">Chronopost</option>'+
          '<option value="UPS">UPS</option>'+
          '<option value="FedEx">FedEx</option>'+
          '<option value="DHL">DHL</option>'+
          '<option value="GLS">GLS</option>'+
          '<option value="TNT">TNT</option>'+
          '<option value="Mondial Relay">Mondial Relay</option>'+
          '<option value="Geodis">Geodis</option>'+
          '<option value="Kuehne+Nagel">Kuehne+Nagel</option>'+
          '<option value="XPO Logistics">XPO Logistics</option>'+
          '<option value="Colis Privé">Colis Privé</option>'+
          '<option value="So Colissimo">So Colissimo</option>'+
          '<option value="Hermes/Evri">Hermes/Evri</option>'+
        '</select>'+
        '<label>Notes</label><textarea id="pdcNotes" placeholder="Référence pièce, destinataire, adresse, contenu..."></textarea>'+
        '</div>'+
        '<div style="margin-top:12px" class="row">'+
        '<button class="btn primary" type="button" id="btnCreate">Enregistrer le dossier</button>'+
        '<input type="text" id="newModel" placeholder="Ajouter un modèle à la liste..." style="max-width:240px" />'+
        '<button class="btn" type="button" id="btnAddModel">+ Modèle</button></div>';

      function syncBranch(){
        var br=document.querySelector('input[name="br"]:checked').value;
        document.getElementById('formDist').classList.toggle('hide',br!=='dist');
        document.getElementById('formMp').classList.toggle('hide',br!=='mp');
        document.getElementById('formPdc').classList.toggle('hide',br!=='pdc');
        // Pièces détachées: pas de gestion modèles
        var nm=document.getElementById('newModel');
        var bm=document.getElementById('btnAddModel');
        if(nm) nm.disabled = (br==='pdc');
        if(bm) bm.disabled = (br==='pdc');
      }
      document.querySelectorAll('input[name="br"]').forEach(function(r){ r.onchange=syncBranch; });
      syncBranch();

      // Pièces PDC : lignes ref + désignation
      function addPdcPieceRow_(ref, des){
        var box=document.getElementById('pdcPieces');
        if(!box) return;
        var idx = box.querySelectorAll('div[data-pdc-piece="1"]').length + 1;
        var row=document.createElement('div');
        row.setAttribute('data-pdc-piece','1');
        row.className='row';
        row.style.gap='10px';
        row.style.flexWrap='wrap';
        row.style.alignItems='center';
        row.style.marginTop='8px';
        row.innerHTML =
          '<input class="pdcRef" placeholder="Référence (ex: 12345)" style="flex:1;min-width:180px" value="'+esc(ref||'')+'" />' +
          '<input class="pdcDes" placeholder="Désignation" style="flex:2;min-width:240px" value="'+esc(des||'')+'" />' +
          '<button class="btn" type="button" data-pdc-rm="1" title="Supprimer">×</button>';
        box.appendChild(row);
        row.querySelectorAll('button[data-pdc-rm="1"]').forEach(function(b){
          b.onclick=function(){ try{ row.remove(); }catch(e){ row.parentNode && row.parentNode.removeChild(row); } };
        });
      }
      function collectPdcPieces_(){
        var box=document.getElementById('pdcPieces');
        if(!box) return {refs:'', designation:''};
        var refs=[];
        var dess=[];
        box.querySelectorAll('div[data-pdc-piece="1"]').forEach(function(r){
          var refEl=r.querySelector('.pdcRef');
          var desEl=r.querySelector('.pdcDes');
          var ref=refEl?String(refEl.value||'').trim():'';
          var des=desEl?String(desEl.value||'').trim():'';
          if(!ref && !des) return;
          refs.push(ref);
          dess.push(des);
        });
        return {refs:refs.join('\\n'), designation:dess.join('\\n')};
      }
      // init (1 ligne) + bouton ajouter
      try{ addPdcPieceRow_('',''); }catch(e){}
      var addBtn=document.getElementById('pdcAddPiece');
      if(addBtn){
        addBtn.onclick=function(){ addPdcPieceRow_('',''); };
      }

      // Rattachement dossier SAV (validation en 2 étapes simple)
      var linkCb=document.getElementById('pdcLinkCase');
      var linkWrap=document.getElementById('pdcLinkWrap');
      var caseInp=document.getElementById('pdcCaseNum');
      var caseCheck=document.getElementById('pdcCaseCheck');
      if(linkCb && linkWrap && caseInp){
        function syncLink_(){
          var on = !!linkCb.checked;
          linkWrap.classList.toggle('hide', !on);
          caseInp.disabled = !on;
          if(!on){
            caseInp.value='';
            if(caseCheck) caseCheck.textContent='';
          }
        }
        linkCb.onchange=syncLink_;
        syncLink_();
      }

      function syncMagOther(){
        var sel=document.getElementById('dMag');
        var wrap=document.getElementById('dMagOtherWrap');
        if(!sel || !wrap) return;
        var v=String(sel.value||'');
        wrap.classList.toggle('hide', v!=='__autre__');
        if(v!=='__autre__'){
          var inp=document.getElementById('dMagOther');
          if(inp) inp.value='';
        }
      }
      var dMagSel=document.getElementById('dMag');
      if(dMagSel) dMagSel.onchange=syncMagOther;
      syncMagOther();

      function initPannePicker(prefix){
        var app=document.getElementById(prefix+'App');
        var typeSel=document.getElementById(prefix+'PanType');
        var panSel=document.getElementById(prefix+'PanSel');
        var otherWrap=document.getElementById(prefix+'PanOtherWrap');
        var other=document.getElementById(prefix+'PanOther');
        if(!app||!typeSel||!panSel) return;

        function setOptions(el, arr){
          el.innerHTML = '<option value="">—</option>' + (arr||[]).map(function(x){ return '<option value="'+esc(x)+'">'+esc(x)+'</option>'; }).join('') + '<option value="__autre__">Autre...</option>';
        }
        function onApp(){
          var a=String(app.value||'');
          var types=a && PANNE_DATA[a] ? Object.keys(PANNE_DATA[a]).sort() : [];
          setOptions(typeSel, types);
          setOptions(panSel, []);
          if(otherWrap) otherWrap.classList.add('hide');
          if(other) other.value='';
        }
        function onType(){
          var a=String(app.value||'');
          var t=String(typeSel.value||'');
          var pans=(a && t && PANNE_DATA[a] && PANNE_DATA[a][t]) ? PANNE_DATA[a][t] : [];
          setOptions(panSel, pans);
          if(otherWrap) otherWrap.classList.add('hide');
          if(other) other.value='';
        }
        function onPan(){
          var v=String(panSel.value||'');
          if(!otherWrap) return;
          otherWrap.classList.toggle('hide', v!=='__autre__');
          if(v!=='__autre__' && other) other.value='';
        }
        app.onchange=onApp;
        typeSel.onchange=onType;
        panSel.onchange=onPan;
        onApp();
      }
      initPannePicker('d');
      initPannePicker('m');

      bindDropZone_('dDrop','dFiles');
      bindDropZone_('mDrop','mFiles');

      document.getElementById('btnReloadModels').onclick=function(){
        document.getElementById('btnReloadModels').textContent='...';
        google.script.run.withSuccessHandler(function(list){
          models=(list||[]).slice().sort();
          modelsLoadedOnce=true;
          renderNewForm();
        }).withFailureHandler(function(e){
          document.getElementById('btnReloadModels').textContent='Rafraîchir';
          showErr(e);
        }).savGetModels();
      };

      var btnReloadModelsMp = document.getElementById('btnReloadModelsMp');
      if(btnReloadModelsMp){
        btnReloadModelsMp.onclick=function(){
          btnReloadModelsMp.textContent='...';
          google.script.run.withSuccessHandler(function(list){
            models=(list||[]).slice().sort();
            modelsLoadedOnce=true;
            renderNewForm();
          }).withFailureHandler(function(e){
            btnReloadModelsMp.textContent='Rafraîchir';
            showErr(e);
          }).savGetModels();
        };
      }

      function renderProcedureForSelectedModel(){
        var mod=(document.getElementById('dMod')&&document.getElementById('dMod').value)||'';
        var pid=procMap[String(mod||'').trim()]||'';
        var p=pid?procDefs[pid]:null;
        var box=document.getElementById('procBox');
        if(!box)return;
        if(!mod){
          box.classList.add('hide');
          box.innerHTML='';
          return;
        }
        var h='';
        h+='<div style="font-weight:900;margin-bottom:6px">Logique SAV</div>';
        if(p){
          h+='<div style="font-size:12px"><b>'+esc(p.label)+'</b></div>';
          if(p.objectif) h+='<div class="muted" style="margin-top:6px">'+esc(p.objectif)+'</div>';
          if(p.documentation&&p.documentation.length){
            h+='<div style="margin-top:8px;font-size:12px"><b>Docs à demander</b><ul style="margin:6px 0 0 18px;padding:0">';
            p.documentation.forEach(function(x){ h+='<li>'+esc(x)+'</li>'; });
            h+='</ul></div>';
          }
          if(p.etapes&&p.etapes.length){
            h+='<div style="margin-top:8px;font-size:12px"><b>Étapes</b><ol style="margin:6px 0 0 18px;padding:0">';
            p.etapes.forEach(function(x){ h+='<li>'+esc(x)+'</li>'; });
            h+='</ol></div>';
          }
        }else{
          h+='<div class="muted" style="font-size:12px">Aucune procédure définie pour ce modèle.</div>';
        }
        box.innerHTML=h;
        box.classList.remove('hide');
      }
      document.getElementById('dMod').onchange=renderProcedureForSelectedModel;
      renderProcedureForSelectedModel();

      document.getElementById('btnAddModel').onclick=function(){
        var n=document.getElementById('newModel').value.trim();
        if(!n)return;
        google.script.run.withSuccessHandler(function(res){
          if(res&&res.ok){ models.push(n); models.sort(); document.getElementById('dMod').innerHTML='<option value="">—</option>'+models.map(function(m){return '<option value="'+esc(m)+'">'+esc(m)+'</option>';}).join(''); document.getElementById('newModel').value=''; }
        }).withFailureHandler(showErr).savAddModele(n);
      };

      document.getElementById('btnCreate').onclick=function(){
        var btn=document.getElementById('btnCreate');
        if(btn && btn.disabled) return;
        var oldTxt = btn ? btn.textContent : '';
        if(btn){ btn.disabled=true; btn.textContent='Enregistrement...'; }
        function done_(){
          try{ if(btn){ btn.disabled=false; btn.textContent=oldTxt||'Enregistrer le dossier'; } }catch(e){}
        }
        var br=document.querySelector('input[name="br"]:checked').value;
        if(br==='dist'){
          var miss=[];
          var magSel=String(document.getElementById('dMag').value||'').trim();
          var magOther=String((document.getElementById('dMagOther')&&document.getElementById('dMagOther').value)||'').trim();
          var mag=(magSel==='__autre__')?magOther:magSel;
          if(!mag) miss.push('Magasin');
          if(!String(document.getElementById('dAdr').value||'').trim()) miss.push('Adresse complète');
          if(!String(document.getElementById('dMod').value||'').trim()) miss.push('Modèle');
          var dApp=String(document.getElementById('dApp').value||'').trim();
          var dType=String(document.getElementById('dPanType').value||'').trim();
          var dPanSel=String(document.getElementById('dPanSel').value||'').trim();
          var dPanOther=String((document.getElementById('dPanOther')&&document.getElementById('dPanOther').value)||'').trim();
          var dPan=(dPanSel==='__autre__')?dPanOther:dPanSel;
          if(!dApp) miss.push('Type appareil');
          if(!dType) miss.push('Famille panne');
          if(!dPan) miss.push('Panne constatée');
          if(!String(document.getElementById('dSer').value||'').trim()) miss.push('Numéro de série');
          if(miss.length){ done_(); showErr('Impossible de valider la demande : champs obligatoires manquants — '+miss.join(', ')); return; }
          google.script.run.withSuccessHandler(function(res){
            // Upload éventuel des pièces jointes
            var input=document.getElementById('dFiles');
            var fl=(input&&input.files)?Array.from(input.files):[];
            if(!fl.length){
              alert('Dossier créé : '+res.numero);
              reloadList(function(){ showTab('list'); });
              done_();
              return;
            }
            uploadFiles_(fl, function(payloadFiles){
              google.script.run.withSuccessHandler(function(u){
                if(u&&u.ok){
                  alert('Dossier créé : '+res.numero+'\\nFichiers enregistrés : '+(u.created||0)+'\\nDossier Drive : '+(u.folderUrl||''));
                }else{
                  alert('Dossier créé : '+res.numero+'\\nUpload fichiers : échec');
                }
                reloadList(function(){ showTab('list'); });
                done_();
              }).withFailureHandler(function(e){ done_(); showErr(e); }).savUploadAttachments({sheet:res.sheet,row:res.row,numero:res.numero,files:payloadFiles});
            });
          }).withFailureHandler(function(e){ done_(); showErr(e); }).savCreateDistributor({
            magasin:mag,
            refDistributeur:document.getElementById('dRef').value,
            adresse:document.getElementById('dAdr').value,
            modele:document.getElementById('dMod').value,
            panne:'['+dApp+' / '+dType+'] '+dPan,
            serie:document.getElementById('dSer').value,
            facture:document.getElementById('dFac').checked,
            photoPlaque:document.getElementById('dPho').checked
          });
        }else if(br==='pdc'){
          var pieces = collectPdcPieces_();
          var payload={
            poids:document.getElementById('pdcPoids').value,
            transporteur:document.getElementById('pdcTrp').value,
            tracking:'',
            notes:document.getElementById('pdcNotes').value,
            pointDepart:document.getElementById('pdcFrom').value,
            adresseArrivee:document.getElementById('pdcTo').value,
            refs:pieces.refs,
            designation:pieces.designation,
            colisPret:document.getElementById('pdcReady').checked,
            dossierSav:''
          };
          var linkOn = !!(document.getElementById('pdcLinkCase') && document.getElementById('pdcLinkCase').checked);
          var caseCheck=document.getElementById('pdcCaseCheck');
          if(linkOn){
            var num = String((document.getElementById('pdcCaseNum')&&document.getElementById('pdcCaseNum').value)||'').trim();
            if(!num){ done_(); showErr('Numéro de dossier SAV requis (rattachement).'); return; }
            if(caseCheck) caseCheck.textContent='Vérification du dossier...';
            google.script.run.withSuccessHandler(function(r){
              if(!(r&&r.ok)){
                done_();
                if(caseCheck) caseCheck.textContent='';
                showErr((r&&r.message)?r.message:'Dossier introuvable');
                return;
              }
              if(caseCheck) caseCheck.textContent='OK — dossier trouvé : '+String(r.type||'');
              payload.dossierSav = num;
              google.script.run.withSuccessHandler(function(res){
                done_();
                if(res && res.ok){
                  alert('Demande PDC créée : '+res.id+'\\nLe suivi est dans Logistique → Pièces détachées.');
                  try{ window.__logMode='pdc'; }catch(e){}
                  showTab('log');
                }else{
                  showErr((res&&res.message)?res.message:'Erreur création PDC');
                }
              }).withFailureHandler(function(e){ done_(); showErr(e); }).pdcCreate(payload);
            }).withFailureHandler(function(e){ done_(); showErr(e); }).savFindCaseByNumero({numero:num});
            return;
          }
          google.script.run.withSuccessHandler(function(res){
            done_();
            if(res && res.ok){
              alert('Demande PDC créée : '+res.id+'\\nLe suivi est dans Logistique → Pièces détachées.');
              try{ window.__logMode='pdc'; }catch(e){}
              showTab('log');
            }else{
              showErr((res&&res.message)?res.message:'Erreur création PDC');
            }
          }).withFailureHandler(function(e){
            done_();
            showErr(e);
          }).pdcCreate(payload);
        }else{
          var miss=[];
          if(!String(document.getElementById('mMp').value||'').trim()) miss.push('Marketplace');
          if(!String(document.getElementById('mNom').value||'').trim()) miss.push('Nom client');
          if(!String(document.getElementById('mMail').value||'').trim()) miss.push('Email client');
          if(!String(document.getElementById('mTel').value||'').trim()) miss.push('Téléphone client');
          if(!String(document.getElementById('mAdr').value||'').trim()) miss.push('Adresse client');
          if(!String(document.getElementById('mMod').value||'').trim()) miss.push('Modèle');
          var mApp=String(document.getElementById('mApp').value||'').trim();
          var mType=String(document.getElementById('mPanType').value||'').trim();
          var mPanSel=String(document.getElementById('mPanSel').value||'').trim();
          var mPanOther=String((document.getElementById('mPanOther')&&document.getElementById('mPanOther').value)||'').trim();
          var mPan=(mPanSel==='__autre__')?mPanOther:mPanSel;
          if(!mApp) miss.push('Type appareil');
          if(!mType) miss.push('Famille panne');
          if(!mPan) miss.push('Panne constatée');
          if(miss.length){ done_(); showErr('Impossible de valider la demande : champs obligatoires manquants — '+miss.join(', ')); return; }
          google.script.run.withSuccessHandler(function(res){
            var input=document.getElementById('mFiles');
            var fl=(input&&input.files)?Array.from(input.files):[];
            if(!fl.length){
              alert('Dossier créé : '+res.numero);
              reloadList(function(){ showTab('list'); });
              done_();
              return;
            }
            uploadFiles_(fl, function(payloadFiles){
              google.script.run.withSuccessHandler(function(u){
                if(u&&u.ok){
                  alert('Dossier créé : '+res.numero+'\\nFichiers enregistrés : '+(u.created||0)+'\\nDossier Drive : '+(u.folderUrl||''));
                }else{
                  alert('Dossier créé : '+res.numero+'\\nUpload fichiers : échec');
                }
                reloadList(function(){ showTab('list'); });
                done_();
              }).withFailureHandler(function(e){ done_(); showErr(e); }).savUploadAttachments({sheet:res.sheet,row:res.row,numero:res.numero,files:payloadFiles});
            });
          }).withFailureHandler(function(e){ done_(); showErr(e); }).savCreateMarketplace({
            marketplace:document.getElementById('mMp').value,
            clientNom:document.getElementById('mNom').value,
            clientEmail:document.getElementById('mMail').value,
            clientTel:document.getElementById('mTel').value,
            clientAdresse:document.getElementById('mAdr').value,
            demandeEnlevement:document.getElementById('mEnl').checked,
            modele:document.getElementById('mMod').value,
            panne:'['+mApp+' / '+mType+'] '+mPan,
            numFacture:document.getElementById('mFac').value,
            photoNs:document.getElementById('mPho').checked
          });
        }
      };
    }

    function uploadFiles_(files, cb){
      // Convertit File -> {name,mimeType,base64} puis callback(payloadFiles)
      var out=[];
      var i=0;
      function next(){
        if(i>=files.length){ cb(out); return; }
        var f=files[i++];
        var r=new FileReader();
        r.onload=function(){
          try{
            var dataUrl=String(r.result||'');
            var parts=dataUrl.split(',');
            var b64=(parts.length>1)?parts[1]:'';
            out.push({name:f.name,mimeType:f.type||'application/octet-stream',base64:b64});
          }catch(e){}
          next();
        };
        r.onerror=function(){ next(); };
        r.readAsDataURL(f);
      }
      next();
    }

    function bindDropZone_(zoneId, inputId){
      var z=document.getElementById(zoneId);
      var inp=document.getElementById(inputId);
      if(!z || !inp) return;
      function renderCount(){
        var n=(inp.files&&inp.files.length)?inp.files.length:0;
        z.textContent = n ? (n+' fichier(s) sélectionné(s) — glisser pour remplacer / cliquer') : 'Glisser vos fichiers ici (ou cliquer)';
      }
      z.onclick=function(){ inp.click(); };
      inp.onchange=function(){ renderCount(); };
      z.addEventListener('dragover', function(e){ e.preventDefault(); z.style.background='rgba(37,99,235,.08)'; });
      z.addEventListener('dragleave', function(){ z.style.background='rgba(37,99,235,.04)'; });
      z.addEventListener('drop', function(e){
        e.preventDefault();
        z.style.background='rgba(37,99,235,.04)';
        if(e.dataTransfer && e.dataTransfer.files){
          inp.files = e.dataTransfer.files;
          renderCount();
        }
      });
      renderCount();
    }

    function refreshProcedures_(cb){
      google.script.run.withSuccessHandler(function(res){
        procDefs=(res&&res.procDefs)||{};
        procMap=(res&&res.procMap)||{};
        if(cb) cb();
      }).withFailureHandler(showErr).savProceduresGetAll();
    }

    function renderProcedures(){
      document.getElementById('mainTitle').textContent='Paramétrage — Procédures SAV';
      var pkeys=Object.keys(procDefs||{}).sort();
      var opts=pkeys.map(function(k){
        var p=procDefs[k]||{};
        return '<option value="'+esc(k)+'">'+esc((p.label||k))+'</option>';
      }).join('');

      document.getElementById('mainBody').innerHTML=
        '<div class="muted">Créer/mettre à jour une procédure et rattacher des modèles. Données stockées dans les onglets <b>SAV_Procedures</b> et <b>SAV_ProcedureModels</b> du classeur.</div>'+
        '<div style="margin-top:14px;padding:12px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.03)">'+
          '<div style="font-weight:900;margin-bottom:8px">1) Créer / modifier une procédure</div>'+
          '<label>Id procédure</label><input id="pId" placeholder="PROC_4_NOUVELLE" />'+
          '<label>Libellé</label><input id="pLabel" placeholder="Procédure X — ..." />'+
          '<label>Objectif</label><textarea id="pObj" placeholder="Objectif de la procédure..."></textarea>'+
          '<label>Documentation (1 élément par ligne)</label><textarea id="pDoc" placeholder="Facture\\nPlaque signalétique\\nPhotos..."></textarea>'+
          '<label>Étapes (1 étape par ligne)</label><textarea id="pSteps" placeholder="1) ...\\n2) ..."></textarea>'+
          '<label class="row"><input type="checkbox" id="pAct" checked /> Active</label>'+
          '<div class="row" style="margin-top:10px">'+
            '<button class="btn primary" type="button" id="btnSaveProc">Enregistrer</button>'+
            '<button class="btn" type="button" id="btnLoadProc">Charger</button>'+
            '<select id="pPick" style="max-width:520px"><option value="">— choisir une procédure —</option>'+opts+'</select>'+
          '</div>'+
        '</div>'+
        '<div style="margin-top:14px;padding:12px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.03)">'+
          '<div style="font-weight:900;margin-bottom:8px">2) Rattacher des modèles à une procédure</div>'+
          '<label>Procédure</label><select id="mProc"><option value="">—</option>'+opts+'</select>'+
          '<label>Modèles (un par ligne, ou séparés par virgule)</label><textarea id="mModels" placeholder="OPC-A01-120\\nOVP-C40\\n..."></textarea>'+
          '<div class="row" style="margin-top:10px">'+
            '<button class="btn primary" type="button" id="btnAssign">Rattacher</button>'+
          '</div>'+
          '<div class="muted" id="assignRes" style="margin-top:8px"></div>'+
        '</div>'+
        '<div style="margin-top:14px;padding:12px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.03)">'+
          '<div style="font-weight:900;margin-bottom:8px">3) Retirer un modèle de son rattachement</div>'+
          '<label>Modèle</label><input id="uModel" placeholder="ex: OPC-A01-120" />'+
          '<div class="row" style="margin-top:10px">'+
            '<button class="btn" type="button" id="btnUnassign">Retirer</button>'+
          '</div>'+
          '<div class="muted" id="unassignRes" style="margin-top:8px"></div>'+
        '</div>';

      function fillFromProcId(pid){
        var p=procDefs[pid]||null;
        if(!p) return;
        document.getElementById('pId').value=p.id||pid;
        document.getElementById('pLabel').value=p.label||'';
        document.getElementById('pObj').value=p.objectif||'';
        document.getElementById('pDoc').value=(p.documentation||[]).join('\\n');
        document.getElementById('pSteps').value=(p.etapes||[]).join('\\n');
        document.getElementById('pAct').checked=true;
      }

      document.getElementById('btnLoadProc').onclick=function(){
        var pid=document.getElementById('pPick').value;
        if(!pid) return;
        fillFromProcId(pid);
      };

      document.getElementById('btnSaveProc').onclick=function(){
        var payload={
          id:document.getElementById('pId').value,
          label:document.getElementById('pLabel').value,
          objectif:document.getElementById('pObj').value,
          documentationText:document.getElementById('pDoc').value,
          etapesText:document.getElementById('pSteps').value,
          actif:document.getElementById('pAct').checked
        };
        google.script.run.withSuccessHandler(function(res){
          if(res&&res.ok){
            refreshProcedures_(function(){ renderProcedures(); });
          }else{
            showErr((res&&res.message)?res.message:'Erreur enregistrement procédure');
          }
        }).withFailureHandler(showErr).savUpsertProcedure(payload);
      };

      document.getElementById('btnAssign').onclick=function(){
        var pid=document.getElementById('mProc').value;
        var mt=document.getElementById('mModels').value;
        google.script.run.withSuccessHandler(function(res){
          if(res&&res.ok){
            document.getElementById('assignRes').textContent='OK — '+res.inserted+' ajoutés, '+res.updated+' mis à jour.';
            refreshProcedures_(function(){});
          }else{
            showErr((res&&res.message)?res.message:'Erreur rattachement');
          }
        }).withFailureHandler(showErr).savAssignModelsToProcedure({procId:pid,modelsText:mt});
      };

      document.getElementById('btnUnassign').onclick=function(){
        var m=document.getElementById('uModel').value;
        google.script.run.withSuccessHandler(function(res){
          if(res&&res.ok){
            document.getElementById('unassignRes').textContent = res.removed ? 'OK — rattachement retiré.' : 'Aucun rattachement trouvé.';
            refreshProcedures_(function(){});
          }else{
            showErr((res&&res.message)?res.message:'Erreur');
          }
        }).withFailureHandler(showErr).savUnassignModelFromProcedure(m);
      };
    }

    var allCases=[];
    var listQuery='';
    var listSearchFocusedOnce=false;
    var globalQuery='';
    var lastListFetchMs=0;
    var listFetchInFlight=false;
    var listFetchWaiters=[];
    function reloadList(cb, force){
      var now=Date.now();
      // si données récentes: ne pas refetch (rend l'UI beaucoup plus fluide)
      if(!force && allCases && allCases.length && (now-lastListFetchMs)<25000){
        if(cb) cb();
        return;
      }
      if(cb) listFetchWaiters.push(cb);
      if(listFetchInFlight) return;
      listFetchInFlight=true;
      google.script.run.withSuccessHandler(function(rows){
        allCases=rows||[];
        lastListFetchMs=Date.now();
        listFetchInFlight=false;
        // Robustesse : la fonction KPIs peut ne pas exister selon versions.
        try{ if(typeof refreshKpis_ === 'function') refreshKpis_(); }catch(e){}
        // exécute les callbacks en attente
        var w=listFetchWaiters.slice(0);
        listFetchWaiters=[];
        w.forEach(function(fn){ try{ if(fn) fn(); }catch(e){} });
        var tab=document.querySelector('.tab.on');
        if(tab&&tab.getAttribute('data-tab')==='list') renderList(false);
        if(tab&&tab.getAttribute('data-tab')==='arch') renderList(true);
      }).withFailureHandler(function(e){
        listFetchInFlight=false;
        listFetchWaiters=[];
        // Affiche l'erreur ET garde une UI utilisable (barre de recherche visible).
        try{
          var tab=document.querySelector('.tab.on');
          if(tab&&tab.getAttribute('data-tab')==='list') renderList(false);
          if(tab&&tab.getAttribute('data-tab')==='arch') renderList(true);
        }catch(err){}
        showErr(e);
      }).savListCases();
    }

    function renderList(arch){
      document.getElementById('mainTitle').textContent=arch?'Dossiers archivés':'Mes dossiers';
      var base=allCases.filter(function(c){ var a=(c.etat==='Archivé'); return arch?a:!a; });

      var q=String(listQuery||'').trim().toLowerCase();
      var tokens=q ? q.split(/\s+/g).filter(Boolean) : [];
      function textFor(c){
        var parts=[];
        parts.push(String(c.numero||''));
        // "Distributeur" : magasin ; "Marketplace" : marketplace
        parts.push(String(c.magasin||''));
        parts.push(String(c.marketplace||''));
        parts.push(String(c.modele||''));
        parts.push(String(c.refDistributeur||''));
        return parts.join(' ').toLowerCase();
      }

      var rows=base;
      if(tokens.length){
        rows=base.filter(function(c){
          var t=textFor(c);
          for(var i=0;i<tokens.length;i++){
            if(t.indexOf(tokens[i])===-1) return false;
          }
          return true;
        });
      }

      var h='';
      h+='<div class="row" style="gap:10px;align-items:stretch;margin-bottom:10px">';
      h+='<input id="listSearch" placeholder="Rechercher par numéro de dossier, distributeur ou modèle..." style="flex:1" value="'+esc(listQuery||'')+'" />';
      h+='<button class="btn" type="button" id="listSearchClear" style="white-space:nowrap">Effacer</button>';
      h+='</div>';
      h+='<div class="muted" style="margin:0 0 10px 2px">'+esc(rows.length)+' dossier(s) affiché(s)</div>';

      if(!base.length){
        h+='<div class="muted">Chargement des dossiers…</div>';
        document.getElementById('mainBody').innerHTML=h;
      }else if(!rows.length){
        h+='<div class="muted">Aucun dossier ne correspond à la recherche.</div>';
        document.getElementById('mainBody').innerHTML=h;
      } else {
        h+='<div class="list">';
        rows.forEach(function(c){
          var unread = Number(c && c.mpUnread ? c.mpUnread : 0);
          var pill = unread ? (' <span class="badge warn" style="margin-left:6px">MP '+esc(String(unread))+'</span>') : '';
          h+='<div class="item" data-idx="'+c.numero+'"><b>'+esc(c.numero)+'</b>'+pill+' · '+esc(c.type)+' · '+esc(c.etat)+'<div class="muted">'+esc(c.dateCreation)+'</div></div>';
        });
        h+='</div>';
        document.getElementById('mainBody').innerHTML=h;
      }

      var sEl=document.getElementById('listSearch');
      if(sEl){
        sEl.oninput=function(){
          listQuery=sEl.value||'';
          renderList(arch);
        };
        // garde le focus uniquement au premier rendu (évite comportements étranges selon navigateur)
        if(!listSearchFocusedOnce){
          listSearchFocusedOnce=true;
          try{
            sEl.focus();
            var v=sEl.value||'';
            sEl.setSelectionRange(v.length,v.length);
          }catch(e){}
        }
      }
      var cEl=document.getElementById('listSearchClear');
      if(cEl){
        cEl.onclick=function(){
          listQuery='';
          renderList(arch);
        };
      }

      document.querySelectorAll('.item').forEach(function(el){
        el.onclick=function(){
          var num=el.getAttribute('data-idx');
          // chercher dans allCases (compat: pas de Array.prototype.find)
          var c=null;
          for(var i=0;i<allCases.length;i++){
            if(String(allCases[i] && allCases[i].numero || '')===String(num||'')){ c=allCases[i]; break; }
          }
          if(c) viewDetail(c);
        };
      });
    }

    function showTab(name){
      var mg=document.getElementById('mainGrid');
      if(mg) mg.classList.remove('logistique-mode');
      document.querySelectorAll('.tab').forEach(function(t){ t.classList.toggle('on',t.getAttribute('data-tab')===name); });
      document.getElementById('diag').style.display='none';
      refreshTabBadges_();

      if(name==='new'){ setCurrent(null); renderNewForm(); return; }
      if(name==='list'){
        setCurrent(null);
        renderList(false); // UI immédiate (barre de recherche visible)
        // reloadList rend déjà la liste à la fin du fetch (ou immédiatement si cache récent)
        // => on ne redemande pas renderList ici (évite double rendu => navigation plus fluide)
        reloadList(null, false);
        return;
      }
      if(name==='arch'){
        setCurrent(null);
        renderList(true); // UI immédiate (barre de recherche visible)
        // même logique que 'list' (évite double rendu)
        reloadList(null, false);
        return;
      }
      if(name==='rdv'){
        setCurrent(null);
        renderRdvCalendar_();
        return;
      }
      if(name==='log'){
        if(mg) mg.classList.add('logistique-mode');
        setCurrent(null);
        renderLogistique();
        return;
      }
      if(name==='pdc'){
        if(mg) mg.classList.add('logistique-mode');
        setCurrent(null);
        renderPdc_();
        return;
      }
      if(name==='avoir'){
        if(mg) mg.classList.add('logistique-mode');
        setCurrent(null);
        renderAvoirTickets();
        return;
      }
      if(name==='mp'){
        if(mg) mg.classList.add('logistique-mode');
        setCurrent(null);
        renderMarketplace();
        return;
      }
      if(name==='sys'){
        setCurrent(null);
        renderSystem_();
        return;
      }
    }

    function renderRdvCalendar_(){
      document.getElementById('mainTitle').textContent='📅 Calendrier des rendez-vous (Procédure 3 — Visio)';
      document.getElementById('mainBody').innerHTML='<div class="muted">Chargement...</div>';
      google.script.run.withSuccessHandler(function(rows){
        rows = rows || [];
        var h='';
        h += '<div class="evt" style="margin-bottom:14px">';
        h += '<div style="font-weight:900">Rendez-vous visio</div>';
        h += '<div class="muted">Liste des RDV visio (Procédure 3). Les RDV se créent depuis un dossier (Proc 3) via “Planifier RDV”.</div>';
        h += '</div>';
        if(!rows.length){
          h += '<div class="muted">Aucun rendez-vous.</div>';
          document.getElementById('mainBody').innerHTML=h;
          return;
        }
        h += '<div style="overflow:auto;border:1px solid var(--border);border-radius:12px">';
        h += '<table style="width:100%;border-collapse:collapse;font-size:13px">';
        h += '<thead><tr style="background:rgba(255,255,255,.04)">';
        h += '<th style="text-align:left;padding:10px;border-bottom:1px solid var(--border)">Date</th>';
        h += '<th style="text-align:left;padding:10px;border-bottom:1px solid var(--border)">Heure</th>';
        h += '<th style="text-align:left;padding:10px;border-bottom:1px solid var(--border)">Dossier</th>';
        h += '<th style="text-align:left;padding:10px;border-bottom:1px solid var(--border)">Client</th>';
        h += '<th style="text-align:left;padding:10px;border-bottom:1px solid var(--border)">Modèle</th>';
        h += '<th style="text-align:left;padding:10px;border-bottom:1px solid var(--border)">État</th>';
        h += '<th style="text-align:left;padding:10px;border-bottom:1px solid var(--border)">Notes</th>';
        h += '</tr></thead><tbody>';
        rows.forEach(function(x){
          var badge = 'badge';
          if(String(x.etat||'')==='À VENIR') badge += ' info';
          else if(String(x.etat||'')==='EFFECTUÉ') badge += ' ok';
          else if(String(x.etat||'')==='ANNULÉ') badge += ' warn';
          h += '<tr>';
          h += '<td style="padding:10px;border-bottom:1px solid rgba(255,255,255,.06)">'+esc(x.date||'')+'</td>';
          h += '<td style="padding:10px;border-bottom:1px solid rgba(255,255,255,.06)">'+esc(x.heure||'')+'</td>';
          h += '<td style="padding:10px;border-bottom:1px solid rgba(255,255,255,.06)"><button class="btn" type="button" data-open-num="'+esc(x.numero||'')+'">#'+esc(x.numero||'')+'</button></td>';
          h += '<td style="padding:10px;border-bottom:1px solid rgba(255,255,255,.06)">'+esc(x.client||'—')+'</td>';
          h += '<td style="padding:10px;border-bottom:1px solid rgba(255,255,255,.06)">'+esc(x.modele||'')+'</td>';
          h += '<td style="padding:10px;border-bottom:1px solid rgba(255,255,255,.06)"><span class="'+badge+'">'+esc(x.etat||'')+'</span></td>';
          h += '<td style="padding:10px;border-bottom:1px solid rgba(255,255,255,.06);white-space:pre-wrap">'+esc(x.notes||'')+'</td>';
          h += '</tr>';
        });
        h += '</tbody></table></div>';
        document.getElementById('mainBody').innerHTML=h;

        document.querySelectorAll('button[data-open-num]').forEach(function(b){
          b.onclick=function(){
            var num=b.getAttribute('data-open-num');
            if(!num) return;
            reloadList(function(){
              var c=findByNumero_(allCases, num);
              if(c){ setCurrent(c); viewDetail(c); showTab('list'); }
              else showErr('Dossier introuvable dans la liste. Recharge “Mes dossiers”.');
            }, false);
          };
        });
      }).withFailureHandler(showErr).savRdvVisioList({onlyOpen:false});
    }

    function renderPdc_(){
      document.getElementById('mainTitle').textContent='Envois pièces détachées';
      document.getElementById('mainBody').innerHTML='<div class="muted">Chargement...</div>';
      google.script.run.withSuccessHandler(function(rows){
        rows=rows||[];
        var h='';
        h += '<div class="evt" style="margin-bottom:14px">';
        h += '<div style="font-weight:900">Pièces détachées — envois</div>';
        h += '<div class="muted">Suivi des envois PDC non livrés (et expédiés en cours). Création possible via Nouvelle → Pièces détachées.</div>';
        h += '</div>';
        if(!rows.length){
          h += '<div class="muted">Aucun envoi en cours.</div>';
          document.getElementById('mainBody').innerHTML=h;
          return;
        }
        rows.forEach(function(x){
          h += '<div class="evt" style="margin-bottom:14px">';
          h += '<div class="row"><div class="name">'+esc(x.id)+'</div><span class="badge warn">'+esc(x.statut||'')+'</span></div>';
          h += '<div class="muted">Créé le '+esc(x.dateCreation||'')+'</div>';
          h += '<div style="margin-top:6px;font-size:12px"><b>Désignation</b> '+esc(x.designation||'—')+'</div>';
          h += '<div style="font-size:12px"><b>Références</b> '+esc(x.refs||'—')+'</div>';
          h += '<div style="margin-top:8px;font-size:12px"><b>Départ</b> '+esc(x.pointDepart||'—')+'</div>';
          h += '<div style="font-size:12px"><b>Arrivée</b> '+esc(x.adresseArrivee||'—')+'</div>';
          h += '</div>';
        });
        document.getElementById('mainBody').innerHTML=h;
      }).withFailureHandler(showErr).pdcList(true);
    }

    function setTabBadge_(tabId, count){
      var el=document.getElementById('badge_'+tabId);
      if(!el) return;
      count = Number(count||0);
      if(count>0){
        el.textContent=String(count);
        el.style.display='inline-flex';
      }else{
        el.textContent='';
        el.style.display='none';
      }
    }

    function refreshTabBadges_(){
      // Throttle: évite un aller-retour serveur à chaque clic d’onglet
      var now=Date.now();
      if(window.__lastBadgeFetchMs && (now-window.__lastBadgeFetchMs)<15000) return;
      if(window.__badgeFetchInFlight) return;
      window.__badgeFetchInFlight=true;
      google.script.run.withSuccessHandler(function(r){
        window.__lastBadgeFetchMs=Date.now();
        window.__badgeFetchInFlight=false;
        r=r||{};
        setTabBadge_('list', r.list||0);
        setTabBadge_('log', r.log||0);
        setTabBadge_('pdc', r.pdc||0);
        setTabBadge_('avoir', r.avoir||0);
        setTabBadge_('mp', r.mp||0);
      }).withFailureHandler(function(){
        window.__badgeFetchInFlight=false;
      }).savGetTabBadges();
    }

    document.querySelectorAll('.tab').forEach(function(t){
      t.onclick=function(){ showTab(t.getAttribute('data-tab')); };
    });
    // Auto: badges après rendu initial (léger)
    try{ window.setTimeout(refreshTabBadges_, 900); }catch(e){ try{ refreshTabBadges_(); }catch(e2){} }

    function renderGlobalSearch_(){
      // Ne pas dupliquer la barre si déjà présente
      if(document.getElementById('globalSearch')) return;
      var wrap=document.getElementById('statsBody');
      if(!wrap) return;
      // injecte au-dessus des stats existantes (sans dépendre du moment où renderStats_ est appelé)
      var box = document.createElement('div');
      box.innerHTML =
        '<div style="margin-bottom:10px;padding:10px 10px;border:1px solid var(--border);border-radius:12px;background:var(--card)">' +
          '<div class="muted" style="font-weight:950;margin:0 0 8px 2px">Recherche dossiers</div>' +
          '<div class="row" style="gap:8px;align-items:stretch">' +
            '<input id="globalSearch" placeholder="Numéro de dossier, distributeur ou modèle..." style="flex:1" />' +
            '<button class="btn primary" type="button" id="globalSearchGo" style="white-space:nowrap">Rechercher</button>' +
          '</div>' +
          '<div class="row" style="gap:8px;align-items:center;margin-top:8px">' +
            '<button class="btn" type="button" id="globalSearchClear" style="white-space:nowrap">Effacer</button>' +
            '<div class="muted" style="font-size:12px">Astuce : plusieurs mots possibles.</div>' +
          '</div>' +
        '</div>';
      // prepend
      wrap.insertBefore(box.firstChild, wrap.firstChild);

      var input=document.getElementById('globalSearch');
      var go=document.getElementById('globalSearchGo');
      var clear=document.getElementById('globalSearchClear');
      function run(){
        globalQuery = input ? (input.value||'') : '';
        listQuery = globalQuery;
        showTab('list');
      }
      if(input){
        input.value = globalQuery || listQuery || '';
        input.onkeydown=function(ev){
          ev = ev || window.event;
          var k = ev.key || ev.keyCode;
          if(k==='Enter' || k===13){ run(); }
        };
      }
      if(go) go.onclick=run;
      if(clear) clear.onclick=function(){
        globalQuery='';
        listQuery='';
        if(input) input.value='';
        // si on est déjà sur list/arch, rafraîchir l'affichage
        try{
          var tab=document.querySelector('.tab.on');
          var t=tab?tab.getAttribute('data-tab'):'';
          if(t==='list') renderList(false);
          if(t==='arch') renderList(true);
        }catch(e){}
      };
    }

    function renderSystem_(){
      document.getElementById('mainTitle').textContent='Système';
      var mailFromClient = 'cedric.gesnouin@optimea.fr';
      document.getElementById('mainBody').innerHTML =
        '<div class="row" style="gap:10px;align-items:stretch;flex-wrap:wrap">' +
          '<button class="btn" type="button" id="sysProcedures">Procédures</button>' +
          '<button class="btn primary" type="button" id="sysStats">Stat global</button>' +
          '<button class="btn" type="button" id="sysTestMail">Tester email ('+esc(mailFromClient)+')</button>' +
          '<button class="btn" type="button" id="sysBackfillDrive">Créer dossiers Drive manquants</button>' +
          '<button class="btn" type="button" id="sysBackfillLogLabels">Créer étiquettes logistique manquantes</button>' +
        '</div>' +
        '<div id="sysBox" style="margin-top:12px"></div>';

      var box=document.getElementById('sysBox');
      var bProc=document.getElementById('sysProcedures');
      var bStats=document.getElementById('sysStats');
      var b3=document.getElementById('sysTestMail');
      var bBF=document.getElementById('sysBackfillDrive');
      var bLogLabels=document.getElementById('sysBackfillLogLabels');

      if(bProc) bProc.onclick=function(){
        refreshProcedures_(function(){ renderProcedures(); });
      };
      if(bStats) bStats.onclick=function(){
        renderSystemStats_();
      };

      if(b3) b3.onclick=function(){
        var to = prompt('Envoyer le test à quelle adresse email ? (ex: toi)', '') || '';
        to = String(to||'').trim();
        if(!to){ showErr('Email requis.'); return; }
        try{ b3.disabled=true; b3.textContent='Envoi...'; }catch(e){}
        if(box) box.innerHTML = '<div class="muted">Envoi en cours...</div>';
        google.script.run.withSuccessHandler(function(r){
          try{ b3.disabled=false; b3.textContent='Tester email ('+mailFromClient+')'; }catch(e){}
          if(r && r.ok){
            if(box) box.innerHTML = '<div class="evt"><div style="font-weight:900">OK — email envoyé</div><div class="muted">Expéditeur : '+esc(r.from||mailFromClient)+'</div></div>';
          }else{
            if(box) box.innerHTML = '<div class="evt" style="border-color:rgba(239,68,68,.35);background:rgba(239,68,68,.08)"><div style="font-weight:900">Erreur</div><div class="muted">'+esc((r&&r.message)?r.message:'Erreur envoi')+'</div></div>';
          }
        }).withFailureHandler(function(e){
          try{ b3.disabled=false; b3.textContent='Tester email ('+mailFromClient+')'; }catch(e){}
          if(box) box.innerHTML = '<div class="evt" style="border-color:rgba(239,68,68,.35);background:rgba(239,68,68,.08)"><div style="font-weight:900">Erreur</div><div class="muted">'+esc(e&&e.message?e.message:String(e))+'</div></div>';
        }).savTestEmail({to:to});
      };

      if(box){
        box.innerHTML = '<div class="muted">Choisis une action ci-dessus (procédures / stats / Drive / email).</div>';
      }

      if(bBF) bBF.onclick=function(){
        if(box) box.innerHTML =
          '<div class="evt" style="margin-bottom:14px">' +
          '<div style="font-weight:900">Créer dossiers Drive manquants</div>' +
          '<div class="muted">Backfill : crée les dossiers Drive pour les dossiers déjà présents dans les feuilles.</div>' +
          '</div>' +
          '<div class="muted" id="bfRes">Traitement...</div>';

        function renderRes(r){
          var res=document.getElementById('bfRes');
          if(!res) return;
          r=r||{};
          if(!r.ok){ res.textContent='Erreur.'; return; }
          var lines=[];
          lines.push('Scanné: '+String(r.totalScanned||0)+' · Créés: '+String(r.totalCreated||0)+' · Déjà OK: '+String(r.totalAlready||0)+' · Ignorés: '+String(r.totalSkipped||0)+(r.dryRun?' (simulation)':''));
          (r.sheets||[]).forEach(function(s){
            lines.push('- '+String(s.sheet||'')+' : scanné '+String(s.scanned||0)+', créés '+String(s.created||0)+', déjà '+String(s.already||0)+', ignorés '+String(s.skipped||0));
          });
          res.textContent = lines.join('\\n');
        }

        google.script.run.withSuccessHandler(function(r){
          renderRes(r);
        }).withFailureHandler(function(e){
          if(box){
            var rr=document.getElementById('bfRes');
            if(rr) rr.textContent='';
          }
          showErr(e);
        }).savBackfillDriveFolders({limit:800,dryRun:false,recheckExisting:true, forceRecreate:true});
      };

      if(bLogLabels) bLogLabels.onclick=function(){
        if(box) box.innerHTML =
          '<div class="evt" style="margin-bottom:14px">' +
          '<div style="font-weight:900">Étiquettes logistique</div>' +
          '<div class="muted">Backfill : crée les PDF d’<b>étiquette d’expédition</b> dans “Étiquettes & courriers” pour les demandes logistique concernées.</div>' +
          '</div>' +
          '<div class="muted" id="bfRes">Traitement...</div>';

        function renderRes(r){
          var res=document.getElementById('bfRes');
          if(!res) return;
          r=r||{};
          if(!r.ok){ res.textContent='Erreur.'; return; }
          res.textContent =
            'Traité: '+String(r.processed||0)+' · Créés: '+String(r.created||0)+' · Déjà OK: '+String(r.already||0)+' · Ignorés: '+String(r.skipped||0);
        }

        google.script.run.withSuccessHandler(function(r){
          renderRes(r);
        }).withFailureHandler(function(e){
          showErr(e);
        }).savBackfillMissingLogistiqueShippingLabels({limit:200,dryRun:false});
      };
    }

    function renderSystemStats_(){
      document.getElementById('mainTitle').textContent='Système — Statistiques globales';
      document.getElementById('mainBody').innerHTML =
        '<div class="row" style="gap:10px;align-items:center;flex-wrap:wrap">' +
          '<button class="btn" type="button" id="backSys">Retour</button>' +
          '<button class="btn primary" type="button" id="loadSysStats">Charger les statistiques globales</button>' +
          '<button class="btn" type="button" id="loadSysAnalytics">Analyses pannes / refs</button>' +
          '<div class="muted" id="sysStatsStatus"></div>' +
        '</div>' +
        '<div id="sysStatsOut" style="margin-top:12px"></div>'+
        '<div id="sysYearClose" style="margin-top:12px"></div>';

      var back=document.getElementById('backSys');
      if(back) back.onclick=function(){ renderSystem_(); };

      var btn=document.getElementById('loadSysStats');
      var btnA=document.getElementById('loadSysAnalytics');
      var status=document.getElementById('sysStatsStatus');
      var out=document.getElementById('sysStatsOut');
      var yc=document.getElementById('sysYearClose');
      if(btn){
        btn.onclick=function(){
          btn.disabled=true;
          btn.textContent='Chargement...';
          if(status) status.textContent='Chargement des statistiques...';
          var statsDone=false;
          var kpisDone=false;
          var gotStats=null;
          var gotKpis=null;

          function maybeRender(){
            if(!(statsDone && kpisDone)) return;
            if(status) status.textContent='';
            btn.disabled=false;
            btn.textContent='Rafraîchir les statistiques globales';
            if(!out) return;

            function card(title, value, color){
              return (
                '<div style="flex:1;min-width:180px;padding:10px 12px;border:1px solid var(--border);border-radius:14px;background:var(--panel);box-shadow:var(--shadow)">' +
                  '<div class="muted" style="font-size:11px;font-weight:900;letter-spacing:.04em;text-transform:uppercase">' + esc(title) + '</div>' +
                  '<div style="margin-top:4px;font-size:20px;font-weight:950;color:' + esc(color || 'var(--text)') + '">' + esc(value) + '</div>' +
                '</div>'
              );
            }

            var h='';
            if(gotKpis){
              h += '<div style="display:flex;flex-wrap:wrap;gap:10px;margin-bottom:12px">';
              h += card('Dossiers total', gotKpis.total||0, 'var(--text)');
              h += card('En cours', gotKpis.enCours||0, 'var(--accent)');
              h += card('Finalisés', gotKpis.finalises||0, 'var(--ok)');
              h += card('Distributeur', gotKpis.distributeurTotal||0, 'var(--text)');
              h += card('Marketplace', gotKpis.marketplaceTotal||0, 'var(--text)');
              h += card('Transport à traiter', gotKpis.transportATraiter||0, 'var(--warn)');
              h += card('Transport expédié', gotKpis.transportExpedie||0, 'var(--accent)');
              h += card('Transport livré', gotKpis.transportLivre||0, 'var(--ok)');
              if(gotKpis.transportTrackingManquant){
                h += card('Expédié sans tracking', gotKpis.transportTrackingManquant||0, 'var(--warn)');
              }
              h += '</div>';
            }
            out.innerHTML = h + '<div id="sysStatsLines"></div>';
            renderStatsInto_(document.getElementById('sysStatsLines'), gotStats);
          }

          google.script.run.withSuccessHandler(function(s){
            gotStats=s||null;
            stats=gotStats;
            statsDone=true;
            maybeRender();
          }).withFailureHandler(function(e){
            statsDone=true;
            if(status) status.textContent=(e&&e.message)?e.message:String(e);
            btn.disabled=false;
            btn.textContent='Charger les statistiques globales';
            showErr(e);
          }).savGetStats();

          google.script.run.withSuccessHandler(function(k){
            gotKpis=k||null;
            kpisDone=true;
            maybeRender();
          }).withFailureHandler(function(e){
            kpisDone=true;
            // On n'empêche pas l'affichage des stats si les KPI échouent
            maybeRender();
          }).savGetKpis();
        };
      }

      if(btnA){
        btnA.onclick=function(){
          btnA.disabled=true;
          btnA.textContent='Chargement...';
          if(status) status.textContent='Chargement des analyses...';
          google.script.run.withSuccessHandler(function(a){
            btnA.disabled=false;
            btnA.textContent='Analyses pannes / refs';
            if(status) status.textContent='';
            a=a||{};
            if(!out) return;

            function barRow(title, count, pct){
              var p = Number(pct||0);
              var w = Math.max(2, Math.min(100, p));
              return (
                '<div style="margin-top:8px;padding:10px 12px;border:1px solid var(--border);border-radius:12px;background:var(--panel)">'+
                  '<div class="row" style="justify-content:space-between;gap:10px;align-items:center">'+
                    '<div style="font-weight:900">'+esc(title)+'</div>'+
                    '<div class="muted">'+esc(count)+' · '+esc(p.toFixed(1))+'%</div>'+
                  '</div>'+
                  '<div style="height:8px;border-radius:999px;background:rgba(15,23,42,.06);margin-top:8px;overflow:hidden">'+
                    '<div style="height:8px;width:'+esc(String(w))+'%;background:var(--accent)"></div>'+
                  '</div>'+
                '</div>'
              );
            }

            function block(title, items, keyLabel){
              var h='';
              h+='<div class="evt" style="margin-bottom:14px">';
              h+='<div style="font-weight:950;margin-bottom:6px">'+esc(title)+'</div>';
              if(!items||!items.length){ h+='<div class="muted">—</div>'; h+='</div>'; return h; }
              items.forEach(function(x){
                h += barRow(String(x.key||keyLabel||'—'), Number(x.count||0), Number(x.pct||0));
              });
              h+='</div>';
              return h;
            }

            function listBlock(title, items){
              var h='';
              h+='<div class="evt" style="margin-bottom:14px">';
              h+='<div style="font-weight:950;margin-bottom:6px">'+esc(title)+'</div>';
              if(!items||!items.length){ h+='<div class="muted">—</div>'; h+='</div>'; return h; }
              h+='<div class="list">';
              items.forEach(function(x){
                h+='<div class="item"><b>'+esc(String(x.key||''))+'</b><div class="muted">'+esc(String(x.count||0))+' signalement(s)</div></div>';
              });
              h+='</div></div>';
              return h;
            }

            var total = Number(a.totalCases||0);
            var header =
              '<div class="evt" style="margin-bottom:14px">'+
                '<div style="font-weight:950">Analyses (taux / top)</div>'+
                '<div class="muted">Base : '+esc(total)+' dossier(s) (ouverts + archivés). Les % sont des parts de cette base.</div>'+
              '</div>';

            var html = header;

            // Répartition Distributeur / Marketplace
            if(a.byType){
              var bt=a.byType;
              var btTotal=Number((bt.Distributeur||0))+(bt.Marketplace||0);
              if(btTotal>0){
                html += '<div class="evt" style="margin-bottom:14px">';
                html += '<div style="font-weight:950;margin-bottom:6px">Répartition Distributeur / Marketplace</div>';
                html += '<div style="display:flex;gap:10px;flex-wrap:wrap">';
                html += '<div style="flex:1;min-width:140px;padding:10px 12px;border:1px solid var(--border);border-radius:14px;background:var(--panel);text-align:center"><div class="muted" style="font-size:11px;font-weight:900;text-transform:uppercase">Distributeur</div><div style="font-size:20px;font-weight:950">'+esc(bt.Distributeur||0)+'</div><div class="muted">'+esc(Math.round((bt.Distributeur||0)/btTotal*100))+'%</div></div>';
                html += '<div style="flex:1;min-width:140px;padding:10px 12px;border:1px solid var(--border);border-radius:14px;background:var(--panel);text-align:center"><div class="muted" style="font-size:11px;font-weight:900;text-transform:uppercase">Marketplace</div><div style="font-size:20px;font-weight:950">'+esc(bt.Marketplace||0)+'</div><div class="muted">'+esc(Math.round((bt.Marketplace||0)/btTotal*100))+'%</div></div>';
                html += '</div></div>';
              }
            }

            html += block('Pannes les plus signalées', a.topPannes||[]);
            html += block('Familles de pannes', a.topFamilles||[]);
            html += block('Types appareils', a.topAppareils||[]);
            html += block('Modèles les plus concernés', a.topModeles||[]);
            html += block('Décisions expertise (dernière)', a.topDecisions||[]);
            html += block('Répartition par état dossier', a.topEtats||[]);
            html += listBlock('Références les plus signalées (PDC + Expertise)', a.topRefs||[]);

            out.innerHTML = html;
          }).withFailureHandler(function(e){
            btnA.disabled=false;
            btnA.textContent='Analyses pannes / refs';
            if(status) status.textContent=(e&&e.message)?e.message:String(e);
            showErr(e);
          }).savBuildGlobalAnalyticsSnapshot();
        };
      }

      // Ticket de clôture annuelle (créé automatiquement pour N-1 au passage en N)
      try{
        google.script.run.withSuccessHandler(function(t){
          if(!yc) return;
          if(!(t&&t.ok)){ yc.innerHTML=''; return; }
          var y = Number(t.year||0);
          yc.innerHTML =
            '<div class="evt" style="margin-bottom:14px">'+
              '<div style="font-weight:950">Clôture annuelle SAV</div>'+
              '<div class="muted">Ticket automatique pour l\\'année '+esc(String(y))+' (créé au changement d\\'année). Remplis les volumes par référence puis génère l\\'analytics annuelle.</div>'+
              '<div style="margin-top:10px;padding:10px 12px;border:1px dashed rgba(37,99,235,.35);border-radius:12px;background:rgba(37,99,235,.04)">'+
                '<div style="font-weight:900;margin-bottom:6px">Importer volumes (CSV)</div>'+
                '<div class="muted">Fichier CSV avec entêtes : <b>Année</b>, <b>Référence</b>, <b>Volume produits</b>. (Excel .xlsx : exporter en CSV)</div>'+
                '<div class="row" style="gap:10px;align-items:center;flex-wrap:wrap;margin-top:10px">'+
                  '<div id="ycDrop" style="flex:1;min-width:220px;padding:10px 12px;border:1px dashed rgba(37,99,235,.35);border-radius:12px;background:rgba(255,255,255,.04);color:var(--muted);font-size:12px;cursor:pointer">Glisser le CSV ici (ou cliquer)</div>'+
                  '<input type="file" id="ycFiles" class="hide" />'+
                  '<button class="btn" type="button" id="ycImport">Importer</button>'+
                '</div>'+
                '<div class="muted" id="ycImportRes" style="margin-top:8px"></div>'+
              '</div>'+
              '<div class="row" style="gap:10px;flex-wrap:wrap;margin-top:10px">'+
                '<button class="btn" type="button" id="ycOpenVol">Ouvrir feuille volumes</button>'+
                '<button class="btn primary" type="button" id="ycGen">Générer analytics annuelle '+esc(String(y))+'</button>'+
              '</div>'+
              '<div class="muted" id="ycRes" style="margin-top:8px"></div>'+
            '</div>';

          // Drop zone import CSV
          try{ bindDropZone_('ycDrop','ycFiles'); }catch(e){}

          var b1=document.getElementById('ycOpenVol');
          if(b1){
            b1.onclick=function(){
              // Texte d'info uniquement (évite les anciennes interpolations non résolues côté navigateur)
              alert('Ouvre la feuille Google Sheets : SAV_Analytics_Volumes\\nColonnes : Année | Référence | Volume produits');
            };
          }

          var bImp=document.getElementById('ycImport');
          if(bImp){
            bImp.onclick=function(){
              var input=document.getElementById('ycFiles');
              var res=document.getElementById('ycImportRes');
              var fl=(input&&input.files&&input.files.length)?Array.from(input.files):[];
              if(!fl.length){ showErr('Sélectionne un CSV.'); return; }
              if(res) res.textContent='Import...';
              bImp.disabled=true;
              uploadFiles_([fl[0]], function(payloadFiles){
                var f = payloadFiles && payloadFiles.length ? payloadFiles[0] : null;
                google.script.run.withSuccessHandler(function(r){
                  bImp.disabled=false;
                  if(r&&r.ok){
                    if(res) res.textContent='OK — '+esc(String(r.inserted||0))+' ajouté(s), '+esc(String(r.updated||0))+' mis à jour, '+esc(String(r.ignored||0))+' ignoré(s).';
                    try{ input.value=''; }catch(e){}
                  }else{
                    if(res) res.textContent='';
                    showErr((r&&r.message)?r.message:'Erreur import');
                  }
                }).withFailureHandler(function(e){
                  bImp.disabled=false;
                  if(res) res.textContent='';
                  showErr(e);
                }).savImportVolumesFromCsv({year:y, file:f});
              });
            };
          }

          var b2=document.getElementById('ycGen');
          if(b2){
            b2.onclick=function(){
              b2.disabled=true;
              b2.textContent='Génération...';
              var res=document.getElementById('ycRes');
              if(res) res.textContent='Calcul en cours...';
              google.script.run.withSuccessHandler(function(r){
                b2.disabled=false;
                b2.textContent='Générer analytics annuelle '+esc(String(y));
                if(r&&r.ok){
                  if(res) res.textContent='OK — '+esc(String(r.refs||0))+' ref(s) trouvées, '+esc(String(r.volumes||0))+' volume(s) saisis. Résultat ajouté dans SAV_Analytics (section ANNUEL '+esc(String(y))+').';
                }else{
                  if(res) res.textContent='';
                  showErr((r&&r.message)?r.message:'Erreur');
                }
              }).withFailureHandler(function(e){
                b2.disabled=false;
                b2.textContent='Générer analytics annuelle '+esc(String(y));
                if(res) res.textContent='';
                showErr(e);
              }).savBuildAnnualAnalyticsSnapshot({year:y});
            };
          }
        }).withFailureHandler(function(){}).savGetAnnualCloseTicket();
      }catch(e){}
    }

    // En cas d'erreur JS (ex: concat HTML cassée), on veut afficher le message
    // au lieu de laisser une page "vide".
    try{
      window.onerror = function(message, source, lineno, colno, error){
        try{ showErr(error||message); }catch(e){}
      };
      window.addEventListener('error', function(ev){
        try{ showErr(ev.error||ev.message||ev); }catch(e){}
      });
    }catch(e){}

    try{ renderNewForm(); }catch(e){ showErr(e); }
    // Panneau gauche : stats présentes comme avant (à la demande via bouton).
    try{ renderStatsShell_(); }catch(e){}
  </script>
</body>
</html>`;
  // Sécurité : si une interpolation côté client se retrouvait accidentellement
  // dans le HTML final (cache / ancienne version), on la neutralise avant d'envoyer.
  // On neutralise uniquement des placeholders connus (anti erreur '$'),
  // sans toucher aux éventuels template-literals valides côté navigateur.
  html = String(html)
    .replace(/\$\{\s*siteLogoHtml\s*\}/g, "")
    .replace(/\$\{\s*JSON\s*\.stringify\s*\(\s*payloadB64\s*\)\s*\}/g, "")
    .replace(/\$\{\s*SAV_SHEET_ANALYTICS_VOLUMES\s*\}/g, "");
  return html;
}