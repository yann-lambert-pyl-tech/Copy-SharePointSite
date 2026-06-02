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
| `-ClientId` | | ClientId d'une App Registration Entra ID (auth interactive) |
| `-LogPath` | | Dossier des logs / modèle exporté (défaut : `.\_SPCopy_Logs`) |
| `-DryRun` | | Simulation : n'écrit **rien** côté cible |

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
- La copie de fichiers utilise `Copy-PnPFile` (côté serveur, sans téléchargement local).

---

## 📄 Licence

Usage interne — Pyl.Tech.
