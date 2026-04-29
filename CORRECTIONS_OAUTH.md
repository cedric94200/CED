# Corrections blocage Google OAuth — Google Apps Script SAV

## Problème
Au chargement de la Web App, certains appels serveur déclenchent
l'écran OAuth de Google à l'intérieur d'un iframe, ce qui bloque
l'affichage de la page.

## Cause 1 : Logo Drive chargé dans doGet()

### Fichier concerné : interface_sav.gs
### Fonction : doGet() → getSavHtml_() → savPdfLogoHtml_() → savGetLogoDataUrl_()

**AVANT (problématique) :**
```javascript
function doGet(e) {
  try {
    ensureSavSheets_();
    const payload = { ... };
    // getSavHtml_ appelle savPdfLogoHtml_() qui accède à Drive → OAuth
    return HtmlService.createHtmlOutput(getSavHtml_(b64))
  }
}
```

**APRÈS (corrigé) :**
```javascript
function doGet(e) {
  try {
    // NE PAS appeler ensureSavSheets_ ici (trop lourd au chargement)
    // Récupérer uniquement le strict minimum
    let models = [];
    let procDefs = {};
    let procMap = {};
    try { models = savGetModels_(); } catch(e) {}
    try { procDefs = savProcedureDefs_(); } catch(e) {}
    try { procMap = savProcedureModelToProcId_(); } catch(e) {}

    const payload = {
      version: SAV_APP_VERSION,
      models,
      procDefs,
      procMap,
      logoUrl: "",  // Logo chargé côté client via savGetLogoUrlForClient()
    };
    const json = JSON.stringify(payload);
    const b64 = Utilities.base64Encode(json, Utilities.Charset.UTF_8);
    return HtmlService.createHtmlOutput(getSavHtml_(b64))
      .setTitle("Gestion SAV")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    // ... gestion erreur existante
  }
}
```

## Cause 2 : savGetLogoDataUrl_() — appel Drive bloquant

**SOLUTION :** Exposer une fonction `savGetLogoUrl()` appelée en JS 
*après* le rendu de la page (via `google.script.run`), et injecter 
le logo dans le DOM une fois reçu.

### Ajouter cette fonction côté serveur :
```javascript
function savGetLogoUrl() {
  // Appelé de manière asynchrone côté client, jamais dans doGet
  try {
    return savGetLogoDataUrl_();
  } catch(e) {
    return "";
  }
}
```

### Côté HTML/JS (dans getSavHtml_) — remplacer ${siteLogoHtml} par :
```html
<div id="siteLogo" style="display:flex;justify-content:center;margin:0 0 10px 0">
  <!-- Logo injecté après chargement -->
</div>
<script>
  // Charge le logo de façon asynchrone (APRÈS le rendu de la page)
  try {
    window.setTimeout(function() {
      google.script.run
        .withSuccessHandler(function(url) {
          if (!url) return;
          var el = document.getElementById('siteLogo');
          if (el) el.innerHTML = '<img src="' + url + '" style="max-width:170px;max-height:56px;object-fit:contain" />';
        })
        .withFailureHandler(function() {})
        .savGetLogoUrl();
    }, 1500); // délai après rendu initial
  } catch(e) {}
</script>
```

## Cause 3 : ensureSavSheets_() dans doGet

Cette fonction vérifie/crée 17 feuilles à chaque chargement.
Elle peut déclencher des scopes supplémentaires.

**SOLUTION :** La déplacer dans un trigger `onOpen` ou l'appeler
via un bouton "Initialiser" dans l'onglet Système.

### Remplacer dans doGet :
```javascript
// SUPPRIMER cette ligne dans doGet :
ensureSavSheets_();

// À la place, créer un trigger onOpen dans le projet :
function onOpen() {
  // Exécuté une fois à l'ouverture du classeur (pas de la web app)
  // Optionnel : SpreadsheetApp.getUi().createMenu('SAV').addItem('Initialiser feuilles', 'ensureSavSheets_').addToUi();
}
```

### Dans chaque fonction API (savCreateDistributor, etc.),
### ensureSavSheets_() peut rester car ces fonctions sont appelées
### APRÈS le consentement OAuth initial.

## Cause 4 : Chargement stats automatique au démarrage

Dans le JS client, `renderStatsShell_()` lance un `setTimeout(load, 600)`
qui déclenche `savGetStats()` dès le chargement.

`savGetStats()` lit plusieurs feuilles Sheets → lourd.

**Le code a déjà un setTimeout(600ms) — bien. Mais il faut s'assurer
que le 1er appel google.script.run n'est PAS dans doGet ou dans un
appel synchrone du chargement.**

## Cause 5 : Déploiement — paramètres essentiels

Dans Apps Script > Déployer > Gérer les déploiements :
- **Exécuter en tant que** : Moi (propriétaire du script)  ← OBLIGATOIRE
- **Accès** : Toute personne (sans connexion Google)  
  OU "Toute personne avec un compte Google" selon besoin

Si "Exécuter en tant que : Moi" est bien configuré, les utilisateurs
ne verront JAMAIS l'écran OAuth. C'est souvent la cause n°1.

## Résumé des modifications à faire dans Cursor

1. **doGet()** : Supprimer `ensureSavSheets_()`, wrapper les 3 appels
   dans des try/catch individuels
2. **getSavHtml_()** : Remplacer `${siteLogoHtml}` par un div vide + 
   script async
3. **Ajouter** la fonction `savGetLogoUrl()` (wrapper async)
4. **Vérifier le déploiement** : "Exécuter en tant que : Moi"

