# Todo Dashboard — Design Spec

## Contexte

Dashboard web personnel affichant les tâches Microsoft To Do dans un semainier visuel. Remplace un système de post-it physiques sur planche en bois par un affichage permanent, agréable et toujours à jour.

**Problème** : Microsoft To Do fonctionne mais manque de visibilité passive — il faut l'ouvrir activement. Le dashboard est conçu pour rester affiché en onglet épinglé.

**Solution** : Un fichier `index.html` unique, hébergé sur GitHub Pages, qui se connecte à l'API Microsoft Graph via MSAL.js pour lire et compléter les tâches.

---

## Stack technique

- **Un seul fichier** : `index.html` (HTML + CSS + JS intégrés)
- **MSAL.js** via CDN (`@azure/msal-browser`) — auth PKCE, pas de client secret
- **Microsoft Graph API** — lecture/écriture des tâches
- **Aucun backend**, aucune base de données, aucun framework
- **Hébergement** : GitHub Pages (`https://FlyinthisSky.github.io/todo-dashboard/`)

### Constantes configurables (haut du fichier)

```js
const CLIENT_ID = "2ca92e37-7ae8-4155-95ab-538a74cda14d";
const REDIRECT_URI = "https://FlyinthisSky.github.io/todo-dashboard/";
const REFRESH_INTERVAL = 15; // minutes
```

---

## Authentification

- **Flux** : Authorization Code avec PKCE (standard SPA)
- **Authority** : `https://login.microsoftonline.com/consumers` (comptes personnels uniquement)
- **Scopes** : `["Tasks.Read", "Tasks.ReadWrite"]`
- **Cache** : `sessionStorage` (géré par MSAL)
- **Refresh** : MSAL acquiert silencieusement un nouveau token si expiré ; si échec → re-login

---

## Endpoints Microsoft Graph

| Endpoint | Usage |
|----------|-------|
| `GET /me/todo/lists` | Récupérer toutes les listes |
| `GET /me/todo/lists/{listId}/tasks` | Tâches d'une liste (filtrer `status eq 'notStarted'`) |
| `PATCH /me/todo/lists/{listId}/tasks/{taskId}` | Marquer une tâche comme complétée |

### Champs utilisés par tâche

`title`, `status`, `dueDateTime`, `importance`, `body`, `categories`

---

## Layout

### Desktop

```
┌──────────────────────────────────────────────────────┐
│  🗓️ Ma Semaine          [🔄 Refresh] [👤 Déconnexion] │
├────────┬─────┬─────┬─────┬─────┬─────┬─────┬────────┤
│ 📥     │ LUN │ MAR │ MER │ JEU │ VEN │ SAM │  DIM   │
│ INBOX  │  7  │  8  │  9  │ 10  │ 11  │ 12  │  13    │
│        │     │     │     │     │     │     │        │
│ [card] │[card│[card│     │[card│[card│     │        │
│ [card] │ card│     │     │     │     │     │        │
│ [card] │    ]│     │     │     │     │     │        │
├────────┴─────┴─────┴─────┴─────┴─────┴─────┴────────┤
│  Auto-refresh dans 12min │ Dernière MàJ: 14:32      │
└──────────────────────────────────────────────────────┘
```

- **Sidebar inbox** à gauche (~200px fixe) : tâches sans `dueDateTime`
- **7 colonnes** (Lundi → Dimanche) : CSS Grid, colonnes égales
- **Header** : titre "Ma Semaine", bouton refresh, bouton déconnexion
- **Footer** : timer de prochain auto-refresh, timestamp de dernière mise à jour

### Mobile (< 768px)

- L'inbox passe en barre horizontale scrollable en haut
- Les jours s'affichent en sections verticales empilées (pleine largeur)
- Les cartes s'empilent verticalement dans chaque section

---

## Design visuel

### Fond

- Couleur de base : dégradé `#1e1812` → `#2a2118` (bois sombre)
- Texture SVG subtile en overlay (grain de bois simulé)

### Cartes de tâches

- **Coins arrondis** : 8px
- **Ombres portées** : `0 4px 12px rgba(0,0,0,0.3)`
- **Gradient de fond** par liste (voir palette ci-dessous)
- **Contenu** : emoji + nom de liste (label), titre de la tâche, badge importance

### Palette de couleurs par liste

| Liste | Gradient | Texte |
|-------|----------|-------|
| Ménage | `#FDFD96` → `#f0e68c` | sombre (`#333`) |
| Boulot | `#FF6B6B` → `#ee5a5a` | blanc |
| Courses | `#55efc4` → `#00b894` | sombre (`#1a1a2e`) |
| Projet perso | `#48dbfb` → `#0abde3` | blanc |
| Sortie | `#a29bfe` → `#6c5ce7` | blanc |

Les nouvelles listes détectées via l'API reçoivent une couleur automatique depuis une palette étendue (orange `#ffa502`, rose `#fd79a8`, etc.).

### Indicateurs d'importance

- **Haute** : badge "⚡ Haute priorité" (texte)
- **Normale** : badge "● Normale" (discret)
- **Basse** : badge "○ Basse" (très discret)

### Tâches en retard

- Bordure rouge pulsante (animation CSS `pulse`)
- Badge "En retard" affiché sur la carte
- Les tâches en retard apparaissent dans la colonne du jour d'échéance original

### Typographie

- Police principale : **Quicksand** (Google Fonts) — arrondie, distinctive, lisible
- Fallback : `system-ui, sans-serif`

### Animations

- **Apparition des cartes** : fade-in + léger slide-up au chargement
- **Complétion** : la carte est barrée visuellement (texte barré + opacité réduite), reste visible ~1.5s, puis fade-out et suppression du DOM
- **Refresh** : les nouvelles cartes apparaissent en fade-in

---

## Flux utilisateur

### 1. Première visite (non authentifié)

Écran centré avec :
- Logo / titre du dashboard
- Bouton "Se connecter avec Microsoft"
- Style cohérent avec le thème (fond sombre, bouton coloré)

### 2. Authentification

- MSAL ouvre une popup (ou redirect selon config)
- Token acquis → stocké en sessionStorage
- Redirect vers le dashboard

### 3. Chargement des données

1. `GET /me/todo/lists` → récupérer toutes les listes
2. Pour chaque liste : `GET /me/todo/lists/{listId}/tasks?$filter=status eq 'notStarted'`
3. Associer chaque liste à une couleur (mapping par nom ou par ordre)
4. Trier les tâches : celles avec `dueDateTime` dans la colonne du jour, celles sans dans l'inbox

### 4. Affichage

- Calculer la semaine en cours (lundi → dimanche)
- Distribuer les tâches dans les colonnes
- Les tâches dont `dueDateTime` est passé et `status` = `notStarted` → marquées "en retard"

### 5. Interaction : compléter une tâche

1. Clic sur la carte
2. `PATCH /me/todo/lists/{listId}/tasks/{taskId}` avec `{ "status": "completed" }`
3. Animation : texte barré + opacité réduite pendant 1.5s
4. Fade-out et suppression du DOM

### 6. Auto-refresh

- Toutes les 15 minutes (configurable via `REFRESH_INTERVAL`)
- Re-fetch silencieux de toutes les tâches
- Mise à jour du DOM sans flash (diff logique : ajouter les nouvelles, retirer les complétées)
- Footer affiche le countdown et le timestamp de dernière MàJ

---

## Gestion des erreurs

| Situation | Comportement |
|-----------|-------------|
| Token expiré | MSAL `acquireTokenSilent` → si échec, re-login automatique |
| API Graph indisponible | Message discret en footer : "Impossible de rafraîchir, nouvelle tentative dans X min" |
| Aucune tâche dans une colonne | Message subtil : "Rien de prévu" (texte grisé) |
| Aucune tâche du tout | Message d'encouragement centré |
| Erreur PATCH (complétion) | Toast discret : "Impossible de marquer comme complétée, réessayez" |

---

## Responsive

| Breakpoint | Comportement |
|------------|-------------|
| ≥ 1024px | Layout complet : sidebar inbox + 7 colonnes |
| 768px–1023px | Sidebar inbox réduite (~150px), colonnes compressées |
| < 768px | Inbox en barre horizontale en haut, jours en sections verticales empilées |

---

## Fichiers livrés

Un seul fichier : `index.html` à la racine du repo `todo-dashboard`, prêt pour GitHub Pages.
