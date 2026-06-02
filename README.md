# Copy-SharePointSite

Script PowerShell **robuste** pour dupliquer un site SharePoint Online complet vers un **nouveau site cible** : structure, contenu et permissions.

Code organisé en **streams** (`#region`), chaque bloc autonome et commenté, avec tous les standards : journalisation, authentification, interface de suivi, gestion d'erreurs et mode simulation.

---

## ✨ Fonctionnalités

- **Provisioning** d'un nouveau site cible (Communication Site ou Team Site), idempotent
- **Extraction** du modèle PnP du site source (listes, champs, content types, pages, navigation, sécurité, fichiers)
- **Application** du modèle sur la cible
- **Copie du contenu** des bibliothèques de documents (fichiers + dossiers, copie côté serveur)
- **Journalisation** : fichier `.log` horodaté + transcript + console colorée
- **Authentification** interactive PnP (compatible MFA), support App Registration (`-ClientId`)
- **Interface de suivi** : `Write-Progress` multi-niveaux + rapport final
- **Robustesse** : retry avec back-off exponentiel (anti-throttling 429), `-DryRun`, `-WhatIf`, validation des URLs

---

## 📋 Pré-requis

- PowerShell **7.2+** (requis par PnP.PowerShell ; Windows PowerShell 5.1 n'est pas supporté)
- Module **PnP.PowerShell**
  ```powershell
  Install-Module PnP.PowerShell -Scope CurrentUser
  ```
- Droits **administrateur tenant SharePoint** (création du site cible)

---

## 🚀 Utilisation

### Test à blanc (aucune écriture côté cible)

```powershell
.\Copy-SharePointSite.ps1 `
  -SourceUrl https://contoso.sharepoint.com/sites/Modele `
  -TargetUrl https://contoso.sharepoint.com/sites/Modele-Copie `
  -Owner admin@contoso.com `
  -DryRun
```

### Duplication réelle, avec contenu

```powershell
.\Copy-SharePointSite.ps1 `
  -SourceUrl https://contoso.sharepoint.com/sites/Modele `
  -TargetUrl https://contoso.sharepoint.com/sites/Modele-Copie `
  -Owner admin@contoso.com `
  -IncludeContent
```

> 💡 Toujours valider en `-DryRun` sur un site jetable avant un run réel.

### Traitement par lot (CSV)

Enchaîne plusieurs opérations (Site **et/ou** Teams) décrites dans un fichier CSV. Chaque ligne est exécutée comme un run complet ; un échec n'interrompt pas le lot, et un résumé global est affiché puis exporté en CSV (`_SPCopy_Logs\BatchResult_*.csv`).

```powershell
# Valider tout le lot en simulation, puis le rejouer en réel
.\Copy-SharePointSite.ps1 -ConfigCsv .\samples\operations.csv -DryRun
.\Copy-SharePointSite.ps1 -ConfigCsv .\samples\operations.csv
```

Colonnes reconnues (voir [`samples/operations.csv`](samples/operations.csv)) :

| Colonne | Pour | Description |
|---------|------|-------------|
| `Mode` | toutes | `Site` (défaut) ou `Team` |
| `SourceUrl`, `TargetUrl` | Site | URLs source / nouvelle cible |
| `TargetTitle`, `TargetType`, `Owner` | Site | options du site cible |
| `IncludeContent` | Site | copie du contenu (`true`/`false`) |
| `SourceTeamId`, `NewTeamName`, `TenantUrl`, `Visibility` | Team | clonage d'équipe |
| `IncludePermissions` | toutes | copie des ACL |
| `DryRun` | toutes | simulation pour cette ligne |

> Les colonnes booléennes acceptent `true/1/yes/oui/o/x`. Les switches globaux `-DryRun` et `-IncludePermissions` s'ajoutent (OR) à chaque ligne — pratique pour forcer une simulation de tout le lot.

---

## ⚙️ Paramètres

| Paramètre | Requis | Description |
|-----------|:------:|-------------|
| `-SourceUrl` | ✅ | URL du site source |
| `-TargetUrl` | ✅ | URL du **nouveau** site cible à créer |
| `-TargetTitle` | | Titre d'affichage (déduit de l'URL si absent) |
| `-TargetType` | | `CommunicationSite` (défaut) ou `TeamSite` |
| `-Owner` | | UPN du propriétaire du nouveau site |
| `-IncludeContent` | | Inclut pages et fichiers de branding dans le modèle |
| `-ClientId` | | ClientId d'une App Registration Entra ID (requis pour l'auth certificat) |
| `-AppName` | | Nom convivial de l'app (étiquette logs/rapport) |
| `-Thumbprint` | | Empreinte du certificat → auth **app-only** (serveur, sans MFA) |
| `-TenantId` | | ID de tenant (GUID) ou domaine `*.onmicrosoft.com` pour l'auth certificat |
| `-ForceContentCopy` | | (Team) Copie **synchrone forcée** du contenu SharePoint après le clone |
| `-MuteNotifications` | | Désactive l'email de bienvenue du groupe (via `Set-UnifiedGroup`) |
| `-LogPath` | | Dossier des logs / modèle exporté (défaut : `.\_SPCopy_Logs`) |
| `-DryRun` | | Simulation : n'écrit **rien** côté cible |

> **Authentification** : par défaut **interactive** (navigateur + MFA). Si `-Thumbprint` est fourni avec `-ClientId`, le script bascule en **app-only par certificat** (le certificat doit être présent dans le magasin de la machine). Le domaine tenant est déduit de l'URL ; si le préfixe SharePoint diffère du tenant (ex. `groupeeiffage.sharepoint.com` mais tenant `eiffage.onmicrosoft.com`), précisez `-TenantId` (GUID ou domaine). L'app doit avoir les permissions d'application adéquates (`Sites.FullControl.All`, et `Group.ReadWrite.All` pour Teams).

---

## 🧱 Architecture (streams)

| Stream | Rôle |
|--------|------|
| 1 — Journalisation | `Write-Log`, `Invoke-WithRetry` (back-off exponentiel) |
| 2 — Pré-requis | Vérifie / charge `PnP.PowerShell` |
| 3 — Authentification | Connexion interactive PnP, déduction de l'URL admin |
| 4 — Extraction source | `Get-PnPSiteTemplate` → fichier `.pnp` |
| 5 — Provisioning cible | `New-PnPSite` (idempotent, respecte `-DryRun`) |
| 6 — Application modèle | `Invoke-PnPSiteTemplate` |
| 7 — Copie contenu | `Copy-PnPFile`, progression imbriquée + stats |
| 8 — Vérification & rapport | Comparatif source/cible + résumé final |
| 9 — Orchestration | Enchaînement, transcript, gestion d'erreur globale |

---

## ⚠️ Limites connues

- `Get-PnPSiteTemplate` ne capture pas tout : workflows, certains web parts custom, **versions d'historique** des fichiers.
- La copie forcée de fichiers utilise un **download/upload** cross-site (lecture source, écriture cible).
- **Source archivée / lecture seule** : la **lecture** (clone, extraction, copie de contenu) fonctionne. Une **équipe archivée** est en lecture seule → l'ajout d'owner/membre échoue tant qu'elle n'est pas désarchivée (le script avertit via `isArchived`). Un **site archivé** (SharePoint Advanced Management) doit être réactivé avant lecture. La cible étant nouvellement créée, l'écriture n'est jamais bloquée par un verrou.
- **Notifications** : seul l'email de bienvenue du groupe est désactivable (`-MuteNotifications` via `Set-UnifiedGroup`). Les notifications Teams in-app d'ajout ne sont pas supprimables par API.

---

## 📄 Licence

Usage interne — Pyl.Tech.
