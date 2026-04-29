# SCHÉMAS DES PROCÉDURES SAV — OPTIMEA
# Modélisation OUI / NON — à compléter / corriger

---

## PROCÉDURE 1 — SAV POSTAL (Appareils < 50 €)

```
[Client signale un problème]
          │
          ▼
  ┌───────────────────────────────┐
  │  Dossier créé dans l'app SAV  │   → état : CRÉATION
  └───────────────────────────────┘
          │
          ▼
  Produit couvert par la garantie ?
       │           │
      OUI          NON ──────────────────────────────→ [Informer le client : hors garantie]
       │                                                         │
       ▼                                                         ▼
  Demande envoi plaque signalétique + photos par courrier    [FIN / Archivé]
          │
          ▼
  Documents reçus ?
       │       │
      OUI      NON ──→ [Relance client]  ──→ (retour à l'attente)
       │
       ▼
  ┌─────────────────────────────────────────┐
  │  En attente réception                   │   → état : En attente réception
  └─────────────────────────────────────────┘
          │
          ▼
  Analyse : n° série / date d'achat / garantie valide ?
       │           │
      OUI          NON ──────────────────────────────→ [Refus garantie — courrier au client]
       │                                                         │
       ▼                                                         ▼
  Défaut couvert (vice de fabrication) ?                   [FIN / Archivé]
       │           │
      OUI          NON ──────────────────────────────→ [Hors garantie — informer client]
       │                                                         │
       ▼                                                         ▼
  Quelle solution ?                                        [FIN / Archivé]
  ┌────────────────────┬──────────────────┐
  │                    │                  │
AVOIR             REMPLACEMENT         AUTRE
  │                    │                  │
  ▼                    ▼                  ▼
Ticket avoir       Envoi produit    (à définir)
créé               neuf au client
  │                    │
  └────────┬───────────┘
           ▼
  ┌───────────────────┐
  │  Clôture + archive │   → état : Archivé
  └───────────────────┘
```

**⚠️ Points à vérifier / compléter :**
- [ ] Délai max avant relance client si pas de courrier reçu ?
- [ ] Qui décide du refus garantie ? Validation hiérarchique ?
- [ ] La solution "avoir" sur PROC_1 est-elle toujours un avoir financier ou parfois un remplacement ?
- [ ] Le client est-il informé par email, courrier, ou téléphone ?
- [ ] Y a-t-il un délai légal à respecter (garantie légale 2 ans) ?

---

## PROCÉDURE 2 — SAV RAPATRIEMENT (Appareils > 50 € — Mobile / Portable)

```
[Client signale un problème]
          │
          ▼
  ┌───────────────────────────────┐
  │  Dossier créé dans l'app SAV  │   → état : CRÉATION
  └───────────────────────────────┘
          │
          ▼
  Informations collectées :
  Facture / n° série / description / photos
          │
          ▼
  Conditions de garantie OK ?
       │           │
      OUI          NON ──────────────────────────────→ [Informer client : hors garantie]
       │                                                         │
       ▼                                                         ▼
  Envoi étiquette transport au client                      [FIN / Archivé]
  (PDF généré dans l'app)
          │
          ▼
  ┌─────────────────────────────────────────┐
  │  En attente réception                   │   → état : En attente réception
  └─────────────────────────────────────────┘
          │
          ▼
  Colis reçu en atelier ?
       │       │
      OUI      NON ──→ [Suivi transport — relance transporteur / client]
       │
       ▼
  Bon de retour (PDF) + contrôle colis à l'arrivée
          │
          ▼
  ┌─────────────────────────────────────────┐
  │  Réception atelier                      │   → état : Réception atelier
  │  (date, état visuel, éléments, tech.)   │
  └─────────────────────────────────────────┘
          │
          ▼
  ┌─────────────────────────────────────────┐
  │  Expertise technique                    │   → état : Expertise
  │  (diagnostic, tests, pièces, cause)     │
  └─────────────────────────────────────────┘
          │
          ▼
  Pièces disponibles ?
       │       │
      OUI      NON ──→ ┌──────────────────────────────┐
       │                │  EN ATTENTE — Pièces         │  → état : EN ATTENTE - Pièces
       │                └──────────────────────────────┘
       │                         │
       │                  Pièces reçues ?
       │                    │       │
       │                   OUI      NON ──→ [Relance fournisseur]
       │                    │
       └────────────────────┘
          │
          ▼
  Décision expertise :
  ┌──────────────┬──────────────┬───────────────┬───────────────┐
  │              │              │               │               │
RÉPARATION    ÉCHANGE         AVOIR         HORS GARANTIE   REFUS GARANTIE
  │              │              │               │               │
  ▼              ▼              ▼               ▼               ▼
Réparation    Préparer       Ticket avoir   Bon envoi HG    Courrier refus
effectuée     appareil       créé           (PDF)           au client
  │           de remplac.       │               │               │
  │              │              ▼               ▼               ▼
  │              │         Clôture avoir   FERMÉ HG         Archivé
  │              │
  └──────┬───────┘
         ▼
  ┌──────────────────────────────────────┐
  │  Transport — demande                 │   → état : Transport - demande
  │  (Bon de transport PDF généré)       │
  └──────────────────────────────────────┘
         │
         ▼
  ┌──────────────────────────────────────┐
  │  Transport — en route                │   → état : Transport - en route
  └──────────────────────────────────────┘
         │
         ▼
  Livraison confirmée ?
       │       │
      OUI      NON ──→ [Suivi transporteur]
       │
       ▼
  ┌──────────────────────────────────────┐
  │  Transport — livré                   │   → état : Transport - livré
  └──────────────────────────────────────┘
         │
         ▼
  Rapport technique (PDF) généré si Distributeur
         │
         ▼
  ┌───────────────────┐
  │  Archivé           │   → état : Archivé
  └───────────────────┘
```

**⚠️ Points à vérifier / compléter :**
- [ ] Qui envoie l'étiquette transport au client ? Email automatique ou manuel ?
- [ ] Délai max pour l'expertise (engagement de délai client) ?
- [ ] Si "Hors garantie" : le client est prévenu avant l'envoi du bon HG ? Il peut refuser le retour ?
- [ ] "Refus garantie" vs "Hors garantie" : quelle différence concrète dans votre process ?
- [ ] Si pièces en attente > X jours : escalade automatique ?
- [ ] Le rapport technique PDF est-il envoyé automatiquement ou manuellement ?
- [ ] L'avoir est-il transmis à la comptabilité ? Comment ?

---

## PROCÉDURE 3 — SAV VISIO (Appareils > 50 € — Fixe)

```
[Client signale un problème]
          │
          ▼
  ┌───────────────────────────────┐
  │  Dossier créé dans l'app SAV  │   → état : CRÉATION
  └───────────────────────────────┘
          │
          ▼
  Informations collectées :
  Facture / n° série / description du problème
          │
          ▼
  Conditions de garantie OK ?
       │           │
      OUI          NON ──────────────────────────────→ [Informer client : hors garantie]
       │                                                         │
       ▼                                                         ▼
  ┌─────────────────────────────────────────────────────┐   [FIN / Archivé]
  │  ÉTAPE VISIO : rendez-vous planifié                 │
  │  Vérification installation / environnement /        │
  │  dégâts externes / pré-diagnostic                   │
  └─────────────────────────────────────────────────────┘
          │
          ▼
  Problème visible / confirmé en visio ?
       │           │
      OUI          NON
       │            │
       │            ▼
       │   Cause identifiable sans rapatriement ?
       │        │           │
       │       OUI          NON ──→ [Conseils / réglages / reset]
       │        │                             │
       │        ▼                      Problème résolu ?
       │   Mauvaise installation ?          │       │
       │     │         │                  OUI      NON ──→ (retour au rapatriement)
       │    OUI        NON                 │
       │     │          │                 ▼
       │     ▼          │           [Dossier clos / Archivé]
       │  Rejet garantie│
       │  (mauvaise     │
       │  utilisation)  │
       │     │          │
       └─────┴──────────┘
                  │
                  ▼
  Rapatriement décidé ?
       │           │
      OUI          NON ──→ [Fin — conseils / archivage]
       │
       ▼
  Envoi étiquette transport au client
  (PDF généré dans l'app)
          │
          ▼
  ┌─────────────────────────────────────────┐
  │  En attente réception                   │   → état : En attente réception
  └─────────────────────────────────────────┘
          │
          ▼
  Colis reçu en atelier ?
       │       │
      OUI      NON ──→ [Suivi transport — relance]
       │
       ▼
  Bon de retour (PDF) + contrôle colis
          │
          ▼
  ┌─────────────────────────────────────────┐
  │  Réception atelier                      │   → état : Réception atelier
  └─────────────────────────────────────────┘
          │
          ▼
  ┌─────────────────────────────────────────┐
  │  Expertise technique approfondie        │   → état : Expertise
  │  (diagnostic confirmé post-visio)       │
  └─────────────────────────────────────────┘
          │
          ▼
  Pièces disponibles ?
       │       │
      OUI      NON ──→ ┌──────────────────────────────┐
       │                │  EN ATTENTE — Pièces         │  → état : EN ATTENTE - Pièces
       │                └──────────────────────────────┘
       │                         │
       │                  Pièces reçues ?
       │                    │       │
       │                   OUI      NON ──→ [Relance fournisseur]
       │                    │
       └────────────────────┘
          │
          ▼
  Décision expertise :
  ┌──────────────┬──────────────┬───────────────┬───────────────┐
  │              │              │               │               │
RÉPARATION    ÉCHANGE         AVOIR         HORS GARANTIE   REFUS GARANTIE
  │              │              │               │               │
  ▼              ▼              ▼               ▼               ▼
Réparation    Préparer       Ticket avoir   Bon envoi HG    Courrier refus
effectuée     appareil       créé           (PDF)           au client
  │           de remplac.       │               │               │
  │              │              ▼               ▼               ▼
  │              │         Clôture avoir   FERMÉ HG         Archivé
  │              │
  └──────┬───────┘
         ▼
  ┌──────────────────────────────────────┐
  │  Transport — demande                 │   → état : Transport - demande
  └──────────────────────────────────────┘
         │
         ▼
  ┌──────────────────────────────────────┐
  │  Transport — en route                │   → état : Transport - en route
  └──────────────────────────────────────┘
         │
         ▼
  Livraison confirmée ?
       │       │
      OUI      NON ──→ [Suivi transporteur]
       │
       ▼
  ┌──────────────────────────────────────┐
  │  Transport — livré                   │   → état : Transport - livré
  └──────────────────────────────────────┘
         │
         ▼
  Rapport technique (PDF) si Distributeur
         │
         ▼
  ┌───────────────────┐
  │  Archivé           │   → état : Archivé
  └───────────────────┘
```

**⚠️ Points à vérifier / compléter :**
- [ ] Comment le RDV visio est-il planifié ? Outil dédié (Teams, Zoom, Meet) ? L'app doit-elle gérer la planification ?
- [ ] Qui conduit la visio ? Toujours le même technicien ?
- [ ] Le résultat de la visio est-il tracé dans l'app ? (notes, photos, décision)
- [ ] Si rejet après visio (mauvaise installation) : un courrier de refus est-il généré ?
- [ ] Après la visio, si le problème est résolu : le dossier est clos comment ? Pas d'étape "Transport" ?
- [ ] La différence PROC_2 vs PROC_3 post-rapatriement : est-ce exactement la même chose, ou y a-t-il des étapes spécifiques PROC_3 ?

---

## RÉCAPITULATIF DES ÉTATS (toutes procédures)

```
CRÉATION
   │
   ├─ (PROC_1 seulement) ──────────────────────────────→ En attente réception
   │                                                             │ (pas de réception atelier)
   │                                                             ↓ (analyse puis décision directe)
   │
   ├─ (PROC_2 / PROC_3) ──→ En attente réception ──→ Réception atelier
   │                                                       │
   │                                              Expertise
   │                                                 │         │
   │                                         EN ATTENTE    Décision
   │                                         - Pièces          │
   │                                              │      ┌─────┴──────────────────┐
   │                                              └──────┤                        │
   │                                                     ▼                        ▼
   │                                            Transport - demande       Clôture avec Avoir
   │                                                     │                  FERMÉ - HG
   │                                            Transport - en route
   │                                                     │
   │                                            Transport - livré
   │
   └──────────────────────────────────────────────────── Archivé
```

---

## QUESTIONS GLOBALES À TRANCHER

| # | Question | Réponse |
|---|----------|---------|
| 1 | Y a-t-il un délai de réponse max au client après création du dossier ? | |
| 2 | Qui crée le dossier : le client lui-même ou un agent SAV ? | |
| 3 | Le client Marketplace et le Distributeur ont-ils des workflows différents ? | |
| 4 | La garantie est-elle toujours 2 ans ou variable selon le produit ? | |
| 5 | En cas de "Hors garantie", un devis de réparation est-il proposé avant retour ? | |
| 6 | L'avoir : qui le valide ? Montant fixe ou calculé ? | |
| 7 | Y a-t-il un suivi des coûts SAV par dossier (pièces + MO + transport) ? | |
| 8 | Le transporteur est-il toujours le même ou choisi au cas par cas ? | |
| 9 | Y a-t-il une procédure d'urgence / escalade si client mécontent ? | |
| 10 | PROC_3 post-rapatriement = identique PROC_2 ou différente ? | |
