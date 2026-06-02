#Requires -Version 7.2
<#
.SYNOPSIS
    Duplique un site SharePoint Online complet, OU clone une équipe Microsoft Teams.

.DESCRIPTION
    Script robuste, structuré en "streams" (#region), couvrant les standards :
      - Authentification (PnP.PowerShell interactif, MFA-friendly)
      - Journalisation (fichier log horodaté + transcript + console colorée)
      - Interface de suivi (Write-Progress multi-niveaux + résumé final)
      - Gestion d'erreurs + retry avec back-off exponentiel (transitoires only)
      - Mode -WhatIf / -DryRun (aucune écriture côté cible)
      - ACL optionnelles (-IncludePermissions)

    Deux modes (parameter sets) :

    MODE SITE (défaut) — duplique un site SharePoint vers un NOUVEAU site :
      1. Connexion site SOURCE
      2. Extraction du modèle PnP (listes, champs, content types, pages,
         navigation, paramètres ; + SiteSecurity si -IncludePermissions)
      3. Provisioning du NOUVEAU site CIBLE (New-PnPSite)
      4. Application du modèle (Invoke-PnPSiteTemplate)
      5. Copie du contenu des bibliothèques (download/upload cross-site)
      6. Vérification + rapport final

    MODE TEAM (-SourceTeamId) — clone une équipe Microsoft Teams via Graph :
      Canaux, onglets, apps et paramètres clonés ; membres/owners (ACL) seulement
      si -IncludePermissions. Le site SharePoint de l'équipe est recréé d'office.

    MODE LOT (-ConfigCsv) — enchaîne plusieurs opérations décrites dans un CSV :
      Une ligne = une opération (Site ou Team via la colonne Mode). Chaque opération
      est exécutée comme un run complet (bannière, étapes, rapport) ; un résumé
      global est affiché et exporté en CSV à la fin. Un échec n'interrompt pas le lot.

.PARAMETER ConfigCsv
    (Mode LOT) Chemin d'un CSV décrivant les opérations. Colonnes reconnues :
    Mode, SourceUrl, TargetUrl, TargetTitle, TargetType, Owner, IncludeContent,
    SourceTeamId, NewTeamName, TenantUrl, Visibility, IncludePermissions, DryRun.
    Les switches globaux -DryRun et -IncludePermissions s'ajoutent (OR) à chaque ligne.

.PARAMETER IncludePermissions
    Copie les ACL. Mode SITE : inclut le handler SiteSecurity (groupes/rôles).
    Mode TEAM : ajoute 'members' aux éléments clonés (membres + owners).
    Par défaut désactivé (le nouveau site/équipe repart sur des permissions propres).

.PARAMETER SourceTeamId
    (Mode TEAM) GUID de l'équipe Microsoft Teams source à cloner.

.PARAMETER NewTeamName
    (Mode TEAM) Nom d'affichage de la nouvelle équipe clonée.

.PARAMETER TenantUrl
    (Mode TEAM) URL racine SharePoint du tenant pour la connexion Graph.
    Ex: https://contoso.sharepoint.com

.PARAMETER Visibility
    (Mode TEAM) Visibilité de la nouvelle équipe : Private (défaut) ou Public.

.PARAMETER SourceUrl
    URL complète du site SharePoint source. Ex: https://contoso.sharepoint.com/sites/Source

.PARAMETER TargetUrl
    URL complète du NOUVEAU site cible à créer. Ex: https://contoso.sharepoint.com/sites/Cible

.PARAMETER TargetTitle
    Titre d'affichage du nouveau site. Défaut : dérivé de l'URL cible.

.PARAMETER TargetType
    Type du nouveau site : TeamSite (groupe M365) ou CommunicationSite. Défaut : CommunicationSite.

.PARAMETER Owner
    UPN du propriétaire du nouveau site (requis pour TeamSite sans groupe / CommunicationSite).

.PARAMETER IncludeContent
    Inclure les éléments de liste / fichiers dans l'extraction du modèle (structure + données).

.PARAMETER ClientId
    ClientId (GUID) de l'App Registration Entra ID. Requis pour l'auth par certificat ;
    optionnel en interactif (sinon ClientId par défaut de PnP.PowerShell).

.PARAMETER AppName
    (Optionnel) Nom convivial de l'App Registration, utilisé comme étiquette dans les
    logs / le rapport. N'intervient pas dans l'authentification elle-même.

.PARAMETER Thumbprint
    (Optionnel) Empreinte du certificat (présent dans le magasin de certificats) pour
    une authentification APP-ONLY par certificat — idéale sur serveur, sans MFA.
    Requiert -ClientId ; le domaine tenant est déduit de l'URL (sauf si -TenantId fourni).

.PARAMETER TenantId
    (Optionnel) ID de tenant (GUID) ou domaine xxx.onmicrosoft.com, utilisé comme
    -Tenant lors de l'auth par certificat. À renseigner quand le préfixe SharePoint
    diffère du domaine du tenant (déduction impossible depuis l'URL).

.PARAMETER ForceContentCopy
    (Mode TEAM) Après le clone, copie SYNCHRONE et forcée du contenu des bibliothèques
    SharePoint (site source -> site de la nouvelle équipe), avec progression. Garantit
    la copie des fichiers sans dépendre du traitement asynchrone de Graph.

.PARAMETER MuteNotifications
    Désactive l'email de bienvenue du groupe lors de l'ajout de membres/owner (via
    Set-UnifiedGroup si Exchange Online est connecté). Les notifications Teams in-app
    ne sont pas supprimables via API.

.PARAMETER LogPath
    Dossier de sortie des logs et du modèle exporté. Défaut : .\_SPCopy_Logs

.PARAMETER DryRun
    Simulation : exécute extraction + diagnostics mais N'ÉCRIT RIEN côté cible.

.EXAMPLE
    # Site, avec contenu ET permissions (ACL)
    .\Copy-SharePointSite.ps1 -SourceUrl https://contoso.sharepoint.com/sites/Modele `
        -TargetUrl https://contoso.sharepoint.com/sites/Modele-Copie `
        -Owner admin@contoso.com -IncludeContent -IncludePermissions

.EXAMPLE
    # Site, test à blanc, sans rien créer côté cible
    .\Copy-SharePointSite.ps1 -SourceUrl ... -TargetUrl ... -Owner admin@contoso.com -DryRun

.EXAMPLE
    # Teams, clone complet avec les membres (ACL)
    .\Copy-SharePointSite.ps1 -SourceTeamId 0a1b2c3d-4e5f-6789-abcd-ef0123456789 `
        -NewTeamName "Projet Alpha (copie)" -TenantUrl https://contoso.sharepoint.com `
        -IncludePermissions

.EXAMPLE
    # Lot : enchaîne toutes les opérations du CSV, en simulation pour valider d'abord
    .\Copy-SharePointSite.ps1 -ConfigCsv .\operations.csv -DryRun

.EXAMPLE
    # Serveur : authentification APP-ONLY par certificat (sans MFA), tenant explicite
    .\Copy-SharePointSite.ps1 -SourceTeamId 0a1b2c3d-4e5f-6789-abcd-ef0123456789 `
        -NewTeamName "Projet Alpha (copie)" -TenantUrl https://contoso.sharepoint.com `
        -ClientId 66e9174a-d89b-4eb1-93b2-edc831f1aa85 `
        -AppName "SP-Rollback-App" -Thumbprint A1B2C3D4E5F6...90 `
        -TenantId 11112222-3333-4444-5555-666677778888 -IncludePermissions

.EXAMPLE
    # Teams : clone + copie forcée du contenu + owner, sans polluer de notifications
    .\Copy-SharePointSite.ps1 -SourceTeamId 0a1b2c3d-4e5f-6789-abcd-ef0123456789 `
        -NewTeamName "Projet Alpha (copie)" -TenantUrl https://contoso.sharepoint.com `
        -Owner chef.projet@contoso.com -ForceContentCopy -MuteNotifications -IncludePermissions

.NOTES
    Auteur  : Pyl.Tech
    Pré-requis : Install-Module PnP.PowerShell -Scope CurrentUser
    Auth    : PnP interactif (navigateur + MFA). Admin tenant requis pour créer le site.
#>

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Site')]
param(
    # --- Jeu de paramètres SITE (duplication d'un site SharePoint) ---
    [Parameter(Mandatory = $true, ParameterSetName = 'Site')]
    [ValidatePattern('^https://.+\.sharepoint\.com/(sites|teams)/.+')]
    [string]$SourceUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'Site')]
    [ValidatePattern('^https://.+\.sharepoint\.com/(sites|teams)/.+')]
    [string]$TargetUrl,

    [Parameter(Mandatory = $false, ParameterSetName = 'Site')]
    [string]$TargetTitle,

    [Parameter(Mandatory = $false, ParameterSetName = 'Site')]
    [ValidateSet('CommunicationSite', 'TeamSite')]
    [string]$TargetType = 'CommunicationSite',

    # -Owner / -IncludeContent : communs (Owner sert au site OU à l'équipe ;
    # IncludeContent n'a d'effet qu'en mode Site, ignoré en mode Team).
    [Parameter(Mandatory = $false)]
    [string]$Owner,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeContent,

    # --- Jeu de paramètres TEAM (clonage d'une équipe Microsoft Teams via Graph) ---
    [Parameter(Mandatory = $true, ParameterSetName = 'Team')]
    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string]$SourceTeamId,

    [Parameter(Mandatory = $true, ParameterSetName = 'Team')]
    [string]$NewTeamName,

    [Parameter(Mandatory = $true, ParameterSetName = 'Team')]
    [ValidatePattern('^https://.+\.sharepoint\.com')]
    [string]$TenantUrl,

    [Parameter(Mandatory = $false, ParameterSetName = 'Team')]
    [ValidateSet('Private', 'Public')]
    [string]$Visibility = 'Private',

    # --- Jeu de paramètres CSV (traitement par lot) ---
    # Chaque ligne du CSV = une opération (Site ou Team). Colonnes reconnues :
    #   Mode, SourceUrl, TargetUrl, TargetTitle, TargetType, Owner, IncludeContent,
    #   SourceTeamId, NewTeamName, TenantUrl, Visibility, IncludePermissions, DryRun
    # Les colonnes booléennes acceptent : true/1/yes/oui/o/x (sinon false).
    [Parameter(Mandatory = $true, ParameterSetName = 'Csv')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$ConfigCsv,

    # --- Paramètres communs à tous les modes ---
    # ACL : copie des permissions (groupes/rôles du site, ou membres/owners de l'équipe).
    [Parameter(Mandatory = $false)]
    [switch]$IncludePermissions,

    # Mode TEAM : copie SYNCHRONE et forcée du contenu SharePoint après le clone
    # (ne dépend pas de l'asynchrone Graph pour les fichiers).
    [Parameter(Mandatory = $false)]
    [switch]$ForceContentCopy,

    # Désactive l'email de bienvenue du groupe lors de l'ajout de membres/owner
    # (via Set-UnifiedGroup si Exchange Online est connecté). Limite le bruit.
    [Parameter(Mandatory = $false)]
    [switch]$MuteNotifications,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    # --- Authentification par CERTIFICAT (app-only, idéal serveur sans MFA) ---
    # Si -Thumbprint est fourni : connexion app-only via ClientId + Thumbprint + Tenant
    # (le domaine tenant est déduit de l'URL). -AppName sert d'étiquette (logs/rapport).
    [Parameter(Mandatory = $false)]
    [string]$AppName,

    [Parameter(Mandatory = $false)]
    [string]$Thumbprint,

    # -TenantId : ID de tenant (GUID) ou domaine xxx.onmicrosoft.com pour l'auth par
    # certificat. À fournir quand le domaine ne se déduit pas de l'URL (préfixe
    # SharePoint différent du tenant). Sinon, déduit automatiquement de l'URL.
    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Join-Path -Path $PSScriptRoot -ChildPath '_SPCopy_Logs'),

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

# Arrête au premier appel .NET/cmdlet non géré (les try/catch restent maîtres).
$ErrorActionPreference = 'Stop'
$script:StartTime = Get-Date

#region ░░ STREAM 1 : Journalisation ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# Log unifié : fichier horodaté + console colorée. Toutes les couches l'utilisent.

if (-not (Test-Path -Path $LogPath)) {
    New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
}
$script:LogFile      = Join-Path $LogPath ("SPCopy_{0:yyyyMMdd_HHmmss}.log" -f $script:StartTime)
$script:TemplateFile = Join-Path $LogPath ("SiteTemplate_{0:yyyyMMdd_HHmmss}.pnp" -f $script:StartTime)

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP', 'DEBUG')][string]$Level = 'INFO'
    )
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line  = "[{0}] [{1,-5}] {2}" -f $stamp, $Level, $Message
    Add-Content -Path $script:LogFile -Value $line -Encoding UTF8

    switch ($Level) {
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'STEP'  { Write-Host $line -ForegroundColor Cyan }
        'DEBUG' { Write-Verbose $line }
        default { Write-Host $line -ForegroundColor Gray }
    }
}

function Test-TransientError {
    # Vrai uniquement pour les erreurs rejouables (throttling / indispo / timeout réseau).
    param([Parameter(Mandatory = $true)]$ErrorRecord)
    $msg    = $ErrorRecord.Exception.Message
    $status = $null
    try { $status = [int]$ErrorRecord.Exception.Response.StatusCode } catch {}
    if ($status -in 429, 500, 502, 503, 504) { return $true }
    return ($msg -match '(?i)throttl|too many requests|temporarily|timed?\s*out|service unavailable|connection (was )?(reset|closed)|operation has timed out')
}

function Invoke-WithRetry {
    # Exécute un scriptblock avec retry + back-off exponentiel, UNIQUEMENT sur erreurs transitoires.
    # Une erreur définitive (404, auth, paramètre invalide) est relancée immédiatement.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Action,
        [string]$Operation = 'opération',
        [int]$MaxAttempts = 4,
        [int]$BaseDelaySec = 3
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $Action
        }
        catch {
            $isTransient = Test-TransientError -ErrorRecord $_
            if (-not $isTransient) {
                Write-Log "Échec non transitoire de '$Operation' : $($_.Exception.Message)" 'ERROR'
                throw
            }
            if ($attempt -eq $MaxAttempts) {
                Write-Log "Échec définitif de '$Operation' après $MaxAttempts tentatives : $($_.Exception.Message)" 'ERROR'
                throw
            }
            $delay = $BaseDelaySec * [math]::Pow(2, $attempt - 1)
            Write-Log "Erreur transitoire sur '$Operation' (tentative $attempt/$MaxAttempts) : $($_.Exception.Message). Nouvel essai dans ${delay}s." 'WARN'
            Start-Sleep -Seconds $delay
        }
    }
}
#endregion

#region ░░ STREAM 1bis : Interface de suivi (UI) ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# Suivi d'exécution pour l'administrateur :
#   - Bannière de démarrage (paramètres de la session)
#   - Liste d'étapes pilotée (Pending/Running/Done/Failed/Skipped) + durée par étape
#   - Barre de progression globale auto-calculée (Write-Progress, Id 1)
#   - Checklist finale récapitulative
# Le tracker est la SEULE source de la barre Id 1 ; les fonctions métier ne
# pilotent que les barres imbriquées (Id 2).

$script:Steps            = [System.Collections.Generic.List[object]]::new()
$script:CurrentStepIndex = 0
$script:ActivityName     = 'Exécution'

function Write-Banner {
    # Affiche un en-tête lisible avec le contexte de la session.
    param([Parameter(Mandatory = $true)][string]$Title, [hashtable]$Fields)
    $bar = '═' * 64
    Write-Host ""
    Write-Host $bar -ForegroundColor Cyan
    Write-Host ("  {0}" -f $Title) -ForegroundColor Cyan
    Write-Host $bar -ForegroundColor Cyan
    if ($Fields) {
        foreach ($k in $Fields.Keys) {
            Write-Host ("  {0,-16}: {1}" -f $k, $Fields[$k]) -ForegroundColor Gray
        }
        Write-Host $bar -ForegroundColor Cyan
    }
    Write-Host ""
}

function Initialize-StepTracker {
    # Déclare la liste ordonnée des étapes à suivre pour la session courante.
    param(
        [Parameter(Mandatory = $true)][string[]]$StepLabels,
        [string]$Activity = 'Exécution'
    )
    $script:Steps.Clear()
    $script:ActivityName = $Activity
    $i = 0
    foreach ($label in $StepLabels) {
        $i++
        $script:Steps.Add([pscustomobject]@{
            Index  = $i
            Label  = $label
            Status = 'Pending'
            Start  = $null
            End    = $null
        })
    }
    $script:CurrentStepIndex = 0
    Update-ProgressUI
}

function Update-ProgressUI {
    # Recalcule et rafraîchit la barre de progression globale (Id 1).
    $total = $script:Steps.Count
    if ($total -eq 0) { return }
    $finished = ($script:Steps | Where-Object { $_.Status -in 'Done', 'Skipped', 'Failed' }).Count
    $pct      = [int](($finished / $total) * 100)
    $curLabel = if ($script:CurrentStepIndex -ge 1) { $script:Steps[$script:CurrentStepIndex - 1].Label } else { 'Initialisation' }
    $elapsed  = (Get-Date) - $script:StartTime
    Write-Progress -Id 1 -Activity $script:ActivityName `
        -Status ("Étape {0}/{1} : {2}  —  écoulé {3:mm\:ss}" -f $script:CurrentStepIndex, $total, $curLabel, $elapsed) `
        -PercentComplete $pct
}

function Invoke-Step {
    # Exécute une étape : marque Running, journalise, chronomètre, marque Done/Failed,
    # met à jour la barre, et renvoie le résultat du scriptblock à l'appelant.
    param(
        [Parameter(Mandatory = $true)][int]$Index,
        [Parameter(Mandatory = $true)][scriptblock]$Action,
        [switch]$ContinueOnError
    )
    $step = $script:Steps[$Index - 1]
    $script:CurrentStepIndex = $Index
    $step.Status = 'Running'
    $step.Start  = Get-Date
    Update-ProgressUI
    Write-Log ("Étape {0}/{1} ▶ {2}" -f $Index, $script:Steps.Count, $step.Label) 'STEP'

    try {
        $result = & $Action
        $step.Status = 'Done'
        $step.End    = Get-Date
        Write-Log ("Étape {0}/{1} ✓ {2} ({3:n1}s)" -f $Index, $script:Steps.Count, $step.Label, ($step.End - $step.Start).TotalSeconds) 'OK'
        Update-ProgressUI
        return $result
    }
    catch {
        $step.Status = 'Failed'
        $step.End    = Get-Date
        Write-Log ("Étape {0}/{1} ✗ {2} : {3}" -f $Index, $script:Steps.Count, $step.Label, $_.Exception.Message) 'ERROR'
        Update-ProgressUI
        if ($ContinueOnError) { return $null }
        throw
    }
}

function Set-StepSkipped {
    # Marque une étape comme volontairement sautée (ex. mode DryRun).
    param([Parameter(Mandatory = $true)][int]$Index, [string]$Reason)
    $step = $script:Steps[$Index - 1]
    $step.Status = 'Skipped'
    Write-Log ("Étape {0}/{1} ⊘ {2}{3}" -f $Index, $script:Steps.Count, $step.Label, $(if ($Reason) { " ($Reason)" } else { '' })) 'WARN'
    Update-ProgressUI
}

function Show-StepChecklist {
    # Imprime la checklist finale : icône d'état + libellé + durée par étape.
    if ($script:Steps.Count -eq 0) { return }
    Write-Progress -Id 1 -Activity $script:ActivityName -Completed
    $bar = '─' * 64
    Write-Host ""
    Write-Host "  SUIVI D'EXÉCUTION (étapes)" -ForegroundColor Cyan
    Write-Host $bar -ForegroundColor DarkGray
    foreach ($s in $script:Steps) {
        $icon, $color = switch ($s.Status) {
            'Done'    { '✓', 'Green' }
            'Failed'  { '✗', 'Red' }
            'Skipped' { '⊘', 'Yellow' }
            'Running' { '▶', 'Cyan' }
            default   { '·', 'DarkGray' }
        }
        $dur = if ($s.Start -and $s.End) { "{0,6:n1}s" -f ($s.End - $s.Start).TotalSeconds } else { '       ' }
        Write-Host ("   {0}  {1,-2} {2,-44} {3}" -f $icon, $s.Index, $s.Label, $dur) -ForegroundColor $color
    }
    Write-Host $bar -ForegroundColor DarkGray
}
#endregion

#region ░░ STREAM 2 : Pré-requis ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# Vérifie / importe le module PnP.PowerShell. Le reste du script en dépend.

function Initialize-Prerequisites {
    Write-Log "Vérification des pré-requis (module PnP.PowerShell)..." 'STEP'
    $module = Get-Module -ListAvailable -Name 'PnP.PowerShell' |
        Sort-Object Version -Descending | Select-Object -First 1

    if (-not $module) {
        Write-Log "Module PnP.PowerShell introuvable." 'ERROR'
        Write-Log "Installez-le : Install-Module PnP.PowerShell -Scope CurrentUser" 'INFO'
        throw "Pré-requis manquant : PnP.PowerShell"
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
    Write-Log "PnP.PowerShell v$($module.Version) chargé." 'OK'
}
#endregion

#region ░░ STREAM 3 : Authentification ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# Connexion PnP interactive (navigateur + MFA). Retourne une connexion réutilisable.

function Get-TenantDomain {
    # Déduit le domaine tenant (xxx.onmicrosoft.com) depuis une URL SharePoint.
    param([Parameter(Mandatory = $true)][string]$AnyUrl)
    if ($AnyUrl -match '^https://([^.]+)\.sharepoint\.com') {
        $name = $Matches[1] -replace '-admin$', ''
        return "$name.onmicrosoft.com"
    }
    throw "Impossible de déduire le domaine tenant depuis : $AnyUrl"
}

function New-PnPConnectionParams {
    # Construit les paramètres de Connect-PnPOnline selon le mode d'auth :
    #   - Certificat (app-only) si -Thumbprint fourni  -> idéal serveur, sans MFA
    #   - Interactif (navigateur + MFA) sinon
    param([Parameter(Mandatory = $true)][string]$Url)
    $p = @{ Url = $Url; ReturnConnection = $true; ErrorAction = 'Stop' }
    if ($Thumbprint) {
        if (-not $ClientId) { throw "Auth par certificat : -ClientId est requis avec -Thumbprint." }
        # -Tenant : ID/domaine explicite si fourni, sinon déduit de l'URL.
        $tenant = if ($TenantId) { $TenantId } else { Get-TenantDomain -AnyUrl $Url }
        $p['ClientId']   = $ClientId
        $p['Thumbprint'] = $Thumbprint
        $p['Tenant']     = $tenant
        Write-Log ("Auth par CERTIFICAT (app-only{0}) : ClientId=$ClientId, Thumbprint=$Thumbprint, Tenant=$tenant" -f $(if ($AppName) { " '$AppName'" } else { '' })) 'INFO'
    }
    else {
        $p['Interactive'] = $true
        if ($ClientId) { $p['ClientId'] = $ClientId }
        Write-Log "Auth INTERACTIVE (navigateur + MFA)." 'INFO'
    }
    return $p
}

function Connect-Site {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [string]$Label = 'site'
    )
    Write-Log "Connexion au $Label : $Url" 'STEP'
    $params = New-PnPConnectionParams -Url $Url
    $conn = Invoke-WithRetry -Operation "connexion $Label" -Action { Connect-PnPOnline @params }
    $web  = Get-PnPWeb -Connection $conn
    Write-Log "Connecté au $Label : '$($web.Title)'." 'OK'
    return $conn
}

function Get-AdminUrl {
    # Déduit l'URL d'admin tenant à partir d'une URL de site (xxx-admin.sharepoint.com).
    param([Parameter(Mandatory = $true)][string]$AnyUrl)
    if ($AnyUrl -match '^https://([^.]+)\.sharepoint\.com') {
        return "https://$($Matches[1])-admin.sharepoint.com"
    }
    throw "Impossible de déduire l'URL d'admin depuis : $AnyUrl"
}
#endregion

#region ░░ STREAM 4 : Extraction du modèle (SOURCE) ░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# Exporte la structure (+contenu optionnel) du site source dans un fichier .pnp.

function Export-SourceTemplate {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Connection)

    Write-Log "Extraction du modèle PnP du site source..." 'STEP'

    # ACL : on n'inclut le handler SiteSecurity (groupes, rôles, attributions) que si demandé.
    $handlerList = [System.Collections.Generic.List[string]]@(
        'Lists', 'Fields', 'ContentTypes', 'Pages', 'PageContents', 'Navigation',
        'RegionalSettings', 'SupportedUILanguages', 'Files', 'WebSettings', 'Publishing'
    )
    if ($IncludePermissions) {
        $handlerList.Add('SiteSecurity')
        Write-Log "ACL activées : permissions du site incluses (groupes/rôles)." 'INFO'
        # Vérification : on remonte en INFO les comptes (adresses) présents sur le site source.
        try {
            $users = @(Get-PnPUser -Connection $Connection | Where-Object { $_.Email -and $_.PrincipalType -eq 'User' })
            Write-Log ("ACL source — comptes avec adresse ({0}) :" -f $users.Count) 'INFO'
            foreach ($u in $users) { Write-Log ("   - {0} <{1}>" -f $u.Title, $u.Email) 'INFO' }
        }
        catch {
            Write-Log "Impossible de lister les comptes du site source : $($_.Exception.Message)" 'WARN'
        }
    }
    else {
        Write-Log "ACL désactivées : permissions NON copiées (héritage par défaut du nouveau site)." 'INFO'
    }
    $handlers = $handlerList -join ','
    $params = @{
        Out           = $script:TemplateFile
        Handlers      = $handlers
        Force         = $true
        Connection    = $Connection
        ErrorAction   = 'Stop'
    }
    if ($IncludeContent) {
        $params['IncludeAllPages'] = $true
        $params['PersistBrandingFiles'] = $true
        Write-Log "Mode contenu activé : pages + fichiers de branding inclus." 'INFO'
    }

    Invoke-WithRetry -Operation 'export du modèle' -Action { Get-PnPSiteTemplate @params } | Out-Null

    if (-not (Test-Path $script:TemplateFile)) {
        throw "Le modèle n'a pas été généré : $script:TemplateFile"
    }
    $size = [math]::Round((Get-Item $script:TemplateFile).Length / 1KB, 1)
    Write-Log "Modèle exporté : $script:TemplateFile (${size} Ko)." 'OK'
}
#endregion

#region ░░ STREAM 5 : Provisioning du site CIBLE ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# Crée le nouveau site. Requiert une connexion ADMIN tenant. Respecte -DryRun.

function New-TargetSite {
    Write-Log "Provisioning du site cible ($TargetType) : $TargetUrl" 'STEP'

    if (-not $TargetTitle) {
        $TargetTitle = ($TargetUrl.TrimEnd('/').Split('/')[-1]) -replace '%20', ' '
        Write-Log "Titre cible déduit : '$TargetTitle'." 'INFO'
    }

    if ($DryRun) {
        Write-Log "[DRYRUN] Site cible NON créé (simulation)." 'WARN'
        return
    }
    if (-not $PSCmdlet.ShouldProcess($TargetUrl, "Créer le site $TargetType")) { return }

    $adminUrl   = Get-AdminUrl -AnyUrl $TargetUrl
    $adminConn  = Connect-Site -Url $adminUrl -Label 'admin tenant'

    # Idempotence : ne pas recréer si déjà présent.
    $exists = $null
    try { $exists = Get-PnPTenantSite -Url $TargetUrl -Connection $adminConn -ErrorAction SilentlyContinue } catch {}
    if ($exists) {
        Write-Log "Le site cible existe déjà : $TargetUrl — provisioning sauté." 'WARN'
        return
    }

    # NB : seul CommunicationSite accepte -Url ; TeamSite dérive son URL de -Alias.
    $siteParams = @{ Title = $TargetTitle; Connection = $adminConn; ErrorAction = 'Stop' }
    switch ($TargetType) {
        'CommunicationSite' {
            $siteParams['Type'] = 'CommunicationSite'
            $siteParams['Url']  = $TargetUrl
            if ($Owner) { $siteParams['Owner'] = $Owner }
        }
        'TeamSite' {
            $siteParams['Type']  = 'TeamSite'
            $siteParams['Alias'] = ($TargetUrl.TrimEnd('/').Split('/')[-1])
            if ($Owner) { $siteParams['Owner'] = $Owner }
        }
    }

    Invoke-WithRetry -Operation 'création du site' -Action { New-PnPSite @siteParams } | Out-Null
    Write-Log "Site cible créé. Attente de la propagation (statut Active)..." 'OK'

    # Polling du statut plutôt qu'un sleep fixe : on attend que le site soit réellement provisionné.
    $deadline = (Get-Date).AddMinutes(5)
    do {
        Start-Sleep -Seconds 5
        $site = $null
        try { $site = Get-PnPTenantSite -Url $TargetUrl -Connection $adminConn -ErrorAction SilentlyContinue } catch {}
        $status = if ($site) { "$($site.Status)" } else { 'Provisioning' }
    } while ($status -ne 'Active' -and (Get-Date) -lt $deadline)

    if ($status -eq 'Active') { Write-Log "Site cible actif." 'OK' }
    else { Write-Log "Délai d'attente dépassé (statut: $status). L'application du modèle peut échouer." 'WARN' }
}
#endregion

#region ░░ STREAM 6 : Application du modèle (CIBLE) ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# Applique le .pnp extrait sur le site cible (structure + éléments).

function Invoke-TargetTemplate {
    Write-Log "Application du modèle sur le site cible..." 'STEP'

    if ($DryRun) {
        Write-Log "[DRYRUN] Modèle NON appliqué (simulation)." 'WARN'
        return
    }
    if (-not $PSCmdlet.ShouldProcess($TargetUrl, 'Appliquer le modèle PnP')) { return }

    $targetConn = Connect-Site -Url $TargetUrl -Label 'site cible'
    $params = @{
        Path                          = $script:TemplateFile
        Connection                    = $targetConn
        ClearNavigation               = $true
        ErrorAction                   = 'Stop'
    }
    Invoke-WithRetry -Operation 'application du modèle' -Action { Invoke-PnPSiteTemplate @params } | Out-Null
    Write-Log "Modèle appliqué sur la cible." 'OK'
    return $targetConn
}
#endregion

#region ░░ STREAM 7 : Copie du contenu des bibliothèques ░░░░░░░░░░░░░░░░░░░░░░░
# Copie fichiers+dossiers de chaque bibliothèque de documents source -> cible.
# Sous-barre de progression (Id 2) imbriquée dans la barre globale (Id 1).

function Copy-LibrariesContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$SourceConn,
        [Parameter(Mandatory = $true)]$TargetConn
    )
    Write-Log "Copie du contenu des bibliothèques de documents..." 'STEP'

    # Bibliothèques de documents non masquées (BaseTemplate 101).
    $libs = Get-PnPList -Connection $SourceConn |
        Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden }

    $stats = [ordered]@{ Libraries = 0; Files = 0; Folders = 0; Errors = 0; SkippedBytes = 0 }

    foreach ($lib in $libs) {
        $stats.Libraries++
        Write-Log "Bibliothèque : '$($lib.Title)' ($($lib.ItemCount) éléments)." 'INFO'

        $items = Invoke-WithRetry -Operation "lecture '$($lib.Title)'" -Action {
            Get-PnPListItem -List $lib -PageSize 500 -Connection $SourceConn -Fields 'FileLeafRef', 'FileRef', 'FSObjType'
        }

        # Préfixe de chemin (web) décodé : FileRef est non-échappé, AbsolutePath est échappé.
        $srcWebPath = [uri]::UnescapeDataString(([uri]$SourceUrl).AbsolutePath)

        # Dossiers d'abord, puis par profondeur croissante : la structure existe avant les fichiers.
        $ordered = $items | Sort-Object `
            @{ Expression = { if ($_['FSObjType'] -eq 1) { 0 } else { 1 } } }, `
            @{ Expression = { "$($_['FileRef'])".Split('/').Count } }

        $n = 0; $total = [math]::Max($ordered.Count, 1)
        foreach ($item in $ordered) {
            $n++
            $serverRel = $item['FileRef']
            $isFolder  = ($item['FSObjType'] -eq 1)
            $leaf       = $item['FileLeafRef']
            Write-Progress -Id 2 -ParentId 1 -Activity "Bibliothèque '$($lib.Title)'" `
                -Status "$n/$total : $leaf" -PercentComplete (($n / $total) * 100)

            # Chemin relatif au web (sans le préfixe du site source), puis dossier parent côté cible.
            $relToWeb     = $serverRel.Substring($srcWebPath.Length).TrimStart('/')
            $parentRelDir = if ($relToWeb.Contains('/')) { $relToWeb.Substring(0, $relToWeb.LastIndexOf('/')) } else { '' }

            if ($DryRun) {
                Write-Log "[DRYRUN] Copierait : $serverRel -> <cible>/$relToWeb" 'DEBUG'
                if ($isFolder) { $stats.Folders++ } else { $stats.Files++ }
                continue
            }

            try {
                if ($isFolder) {
                    # Crée l'arborescence (site-relative) côté cible.
                    Resolve-PnPFolder -SiteRelativePath $relToWeb -Connection $TargetConn -ErrorAction Stop | Out-Null
                    $stats.Folders++
                }
                else {
                    # 1) Garantir le dossier parent côté cible (supprime toute dépendance d'ordre).
                    if ($parentRelDir) {
                        Resolve-PnPFolder -SiteRelativePath $parentRelDir -Connection $TargetConn -ErrorAction Stop | Out-Null
                    }
                    # 2) Copie cross-site fiable : download (source) -> upload (cible).
                    #    Évite la limite same-site-collection de Copy-PnPFile.
                    Invoke-WithRetry -Operation "copie '$leaf'" -Action {
                        $stream = Get-PnPFile -Url $serverRel -AsMemoryStream -Connection $SourceConn -ErrorAction Stop
                        Add-PnPFile -FileName $leaf -Folder $parentRelDir -Stream $stream `
                            -Connection $TargetConn -ErrorAction Stop | Out-Null
                    } | Out-Null
                    $stats.Files++
                }
            }
            catch {
                $stats.Errors++
                Write-Log "Erreur sur '$serverRel' : $($_.Exception.Message)" 'ERROR'
            }
        }
        Write-Progress -Id 2 -ParentId 1 -Activity "Bibliothèque '$($lib.Title)'" -Completed
    }

    Write-Log ("Copie terminée : {0} biblio, {1} fichiers, {2} dossiers, {3} erreur(s)." -f `
        $stats.Libraries, $stats.Files, $stats.Folders, $stats.Errors) 'OK'
    return $stats
}
#endregion

#region ░░ STREAM 8 : Vérification & rapport ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# Comparatif rapide source/cible + résumé final lisible.

function Write-FinalReport {
    [CmdletBinding()]
    param($SourceConn, $TargetConn, $CopyStats)

    Write-Log "Vérification post-copie..." 'STEP'

    $duration = (Get-Date) - $script:StartTime
    $line = '═' * 64

    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host "  RAPPORT DE DUPLICATION SHAREPOINT" -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
    Write-Host ("  Source        : {0}" -f $SourceUrl)
    Write-Host ("  Cible         : {0}" -f $TargetUrl)
    Write-Host ("  Mode          : {0}" -f $(if ($DryRun) { 'DRYRUN (simulation)' } else { 'Réel' }))
    Write-Host ("  Durée         : {0:hh\:mm\:ss}" -f $duration)

    if ($CopyStats) {
        Write-Host ("  Bibliothèques : {0}" -f $CopyStats.Libraries)
        Write-Host ("  Fichiers      : {0}" -f $CopyStats.Files)
        Write-Host ("  Dossiers      : {0}" -f $CopyStats.Folders)
        $errColor = if ($CopyStats.Errors -gt 0) { 'Red' } else { 'Green' }
        Write-Host ("  Erreurs       : {0}" -f $CopyStats.Errors) -ForegroundColor $errColor
    }

    if (-not $DryRun -and $TargetConn) {
        try {
            $srcLists = (Get-PnPList -Connection $SourceConn | Where-Object { -not $_.Hidden }).Count
            $tgtLists = (Get-PnPList -Connection $TargetConn | Where-Object { -not $_.Hidden }).Count
            Write-Host ("  Listes src/cible : {0} / {1}" -f $srcLists, $tgtLists)
        } catch {}
    }

    Write-Host ("  Log           : {0}" -f $script:LogFile)
    Write-Host ("  Modèle        : {0}" -f $script:TemplateFile)
    Write-Host $line -ForegroundColor Cyan
    Write-Host ""
}
#endregion

#region ░░ STREAM 8bis : Clonage d'une équipe Teams (Graph) ░░░░░░░░░░░░░░░░░░░░░
# Clone une équipe Microsoft Teams complète via l'API Graph (clone team).
# Le clone inclut canaux, onglets, apps et paramètres ; les membres/owners (ACL)
# ne sont clonés que si -IncludePermissions est fourni. Le SharePoint sous-jacent
# (site de l'équipe) est recréé automatiquement par le clonage.

function Connect-Tenant {
    # Connexion interactive pour les opérations Graph/Teams (token Graph via PnP).
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Url)
    Write-Log "Connexion (Graph/Teams) : $Url" 'STEP'
    $params = New-PnPConnectionParams -Url $Url
    $conn = Invoke-WithRetry -Operation 'connexion tenant' -Action { Connect-PnPOnline @params }
    Write-Log "Connecté au tenant pour Graph/Teams." 'OK'
    return $conn
}

function Write-DetectedTeamAcl {
    # Remonte en INFO les owners et members de l'équipe source (vérification ACL).
    param(
        [Parameter(Mandatory = $true)][string]$TeamId,
        [Parameter(Mandatory = $true)]$Connection
    )
    Write-Log "Lecture des ACL de l'équipe source (vérification)..." 'INFO'
    foreach ($rel in 'owners', 'members') {
        try {
            $resp = Invoke-WithRetry -Operation "lecture $rel source" -Action {
                Invoke-PnPGraphMethod -Url "v1.0/groups/$TeamId/$rel`?`$select=displayName,userPrincipalName,mail&`$top=999" -Method Get -Connection $Connection
            }
            $people = @($resp.value)
            Write-Log ("ACL source — {0} ({1}) :" -f $rel, $people.Count) 'INFO'
            foreach ($p in $people) {
                $addr = if ($p.userPrincipalName) { $p.userPrincipalName } elseif ($p.mail) { $p.mail } else { '(sans adresse)' }
                Write-Log ("   - {0} <{1}>" -f $p.displayName, $addr) 'INFO'
            }
        }
        catch {
            Write-Log "Impossible de lire '$rel' de l'équipe source : $($_.Exception.Message)" 'WARN'
        }
    }
}

function Copy-Team {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory = $true)]$Connection)

    Write-Log "Clonage de l'équipe Teams source ($SourceTeamId)..." 'STEP'

    # Vérifie que l'équipe source existe + état d'archivage.
    $srcTeam = Invoke-WithRetry -Operation 'lecture équipe source' -Action {
        Invoke-PnPGraphMethod -Url "v1.0/teams/$($SourceTeamId)?`$select=displayName,isArchived" -Method Get -Connection $Connection
    }
    Write-Log "Équipe source : '$($srcTeam.displayName)'." 'OK'
    if ($srcTeam.isArchived) {
        # Source archivée = lecture seule : le clone et la lecture du contenu fonctionnent,
        # mais c'est à signaler (rien ne sera modifié côté source de toute façon).
        Write-Log "Équipe source ARCHIVÉE (lecture seule) : clonage/lecture OK, aucune écriture côté source." 'WARN'
    }

    # mailNickname : alias dérivé du nouveau nom (alphanumérique uniquement).
    $alias = ($NewTeamName -replace '[^a-zA-Z0-9]', '')
    if (-not $alias) { $alias = "team$($script:StartTime.ToString('yyyyMMddHHmmss'))" }

    # -Owner en mode Team : le clone Graph rend l'appelant propriétaire ;
    # l'owner explicite est attribué après coup (étape dédiée via Set-GroupOwner).
    if ($Owner) {
        Write-Log "Propriétaire '$Owner' : sera attribué à la nouvelle équipe après le clone." 'INFO'
    }

    # ACL : les membres/owners ne sont clonés que si demandé.
    $parts = @('apps', 'tabs', 'settings', 'channels')
    if ($IncludePermissions) {
        $parts += 'members'
        Write-Log "ACL activées : membres et owners de l'équipe inclus dans le clone." 'INFO'
        # Vérification : on remonte en INFO les identités détectées côté source.
        Write-DetectedTeamAcl -TeamId $SourceTeamId -Connection $Connection
    }
    else {
        Write-Log "ACL désactivées : seul l'appelant sera owner de la nouvelle équipe." 'INFO'
    }

    $body = @{
        displayName  = $NewTeamName
        mailNickname = $alias
        partsToClone = ($parts -join ',')
        visibility   = $Visibility
        description  = "Clone de '$($srcTeam.displayName)' — $($script:StartTime.ToString('yyyy-MM-dd HH:mm'))"
    }

    if ($DryRun) {
        Write-Log "[DRYRUN] Clonerait l'équipe -> '$NewTeamName' (alias '$alias', parts: $($body.partsToClone))." 'WARN'
        return [ordered]@{ Mode = 'Team'; Source = $srcTeam.displayName; NewTeam = $NewTeamName; Alias = $alias; Parts = $body.partsToClone; Status = 'DRYRUN' }
    }
    if (-not $PSCmdlet.ShouldProcess($NewTeamName, "Cloner l'équipe Teams $SourceTeamId")) { return }

    # Le clone est asynchrone : Graph renvoie une opération (teamsAsyncOperation).
    Invoke-WithRetry -Operation 'clone équipe' -Action {
        Invoke-PnPGraphMethod -Url "v1.0/teams/$SourceTeamId/clone" -Method Post -Content $body -Connection $Connection
    } | Out-Null

    Write-Log "Demande de clonage envoyée (traitement asynchrone côté Microsoft 365)." 'OK'

    return [ordered]@{ Mode = 'Team'; Source = $srcTeam.displayName; NewTeam = $NewTeamName; Alias = $alias; Parts = $body.partsToClone; Status = 'Submitted' }
}

function Wait-ClonedGroup {
    # Attend l'apparition du groupe cloné (clone asynchrone) via recherche par mailNickname.
    # Renvoie l'objet groupe (id, displayName). Lève si timeout.
    param(
        [Parameter(Mandatory = $true)][string]$Alias,
        [Parameter(Mandatory = $true)]$Connection,
        [int]$TimeoutMinutes = 10
    )
    Write-Log "Attente de la disponibilité de l'équipe clonée (clone asynchrone)..." 'INFO'
    $filterValue = [uri]::EscapeDataString("mailNickname eq '$Alias'")
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $group = $null
    do {
        Start-Sleep -Seconds 15
        $resp = Invoke-WithRetry -Operation 'recherche du groupe cloné' -Action {
            Invoke-PnPGraphMethod -Url "v1.0/groups?`$filter=$filterValue&`$select=id,displayName" -Method Get -Connection $Connection
        }
        if ($resp.value -and $resp.value.Count -ge 1) { $group = $resp.value[0] }
        else { Write-Log "Équipe pas encore prête, nouvelle vérification dans 15s..." 'INFO' }
    } while (-not $group -and (Get-Date) -lt $deadline)

    if (-not $group) { throw "Équipe clonée introuvable (mailNickname '$Alias') après $TimeoutMinutes min." }
    Write-Log "Groupe cloné détecté : '$($group.displayName)' ($($group.id))." 'OK'
    return $group
}

function Disable-GroupWelcomeMail {
    # Mute l'email de bienvenue du groupe (seul levier fiable pour limiter le bruit).
    # Utilise Exchange Online (Set-UnifiedGroup) si disponible ; sinon avertit honnêtement.
    param([Parameter(Mandatory = $true)][string]$GroupId)
    if (Get-Command -Name 'Set-UnifiedGroup' -ErrorAction SilentlyContinue) {
        try {
            Set-UnifiedGroup -Identity $GroupId -UnifiedGroupWelcomeMessageEnabled:$false -ErrorAction Stop
            Write-Log "Notifications : email de bienvenue du groupe désactivé (Set-UnifiedGroup)." 'OK'
        }
        catch { Write-Log "Impossible de désactiver l'email de bienvenue : $($_.Exception.Message)" 'WARN' }
    }
    else {
        Write-Log "Mute partiel : Set-UnifiedGroup indisponible (Exchange Online non connecté). L'email de bienvenue ne peut pas être désactivé ; les notifications Teams in-app ne sont pas supprimables via API." 'WARN'
    }
}

function Set-GroupOwner {
    # Ajoute un utilisateur en owner ET membre d'un groupe (un owner Teams doit être membre).
    # Idempotent ('already exist' ignoré). -Mute désactive l'email de bienvenue au préalable.
    param(
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)][string]$OwnerUpn,
        [Parameter(Mandatory = $true)]$Connection,
        [switch]$Mute
    )
    Write-Log "Attribution du propriétaire '$OwnerUpn'..." 'STEP'
    if ($Mute) { Disable-GroupWelcomeMail -GroupId $GroupId }

    $user = Invoke-WithRetry -Operation 'résolution du compte owner' -Action {
        Invoke-PnPGraphMethod -Url "v1.0/users/$($OwnerUpn)?`$select=id,userPrincipalName" -Method Get -Connection $Connection
    }
    $ref = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.id)" }

    foreach ($rel in 'members', 'owners') {
        try {
            Invoke-WithRetry -Operation "ajout dans $rel" -Action {
                Invoke-PnPGraphMethod -Url "v1.0/groups/$GroupId/$rel/`$ref" -Method Post -Content $ref -Connection $Connection
            } | Out-Null
            Write-Log "Ajouté dans '$rel' : $OwnerUpn." 'OK'
        }
        catch {
            if ("$($_.Exception.Message)" -match 'already exist') { Write-Log "$OwnerUpn déjà présent dans '$rel'." 'INFO' }
            else { throw }
        }
    }
    Write-Log "Propriétaire '$OwnerUpn' attribué." 'OK'
}

function Get-GroupSiteUrl {
    # Renvoie l'URL du site SharePoint associé à un groupe/équipe.
    param(
        [Parameter(Mandatory = $true)][string]$GroupId,
        [Parameter(Mandatory = $true)]$Connection
    )
    $site = Invoke-WithRetry -Operation 'lecture site du groupe' -Action {
        Invoke-PnPGraphMethod -Url "v1.0/groups/$GroupId/sites/root?`$select=webUrl" -Method Get -Connection $Connection
    }
    if (-not $site.webUrl) { throw "URL SharePoint introuvable pour le groupe $GroupId." }
    return $site.webUrl
}

function Copy-TeamContent {
    # Copie FORCÉE et synchrone du contenu : bibliothèques de documents du site de l'équipe
    # source -> site de la nouvelle équipe (réutilise Copy-LibrariesContent). Renvoie les stats.
    param(
        [Parameter(Mandatory = $true)][string]$SourceTeamId,
        [Parameter(Mandatory = $true)][string]$NewGroupId,
        [Parameter(Mandatory = $true)]$Connection
    )
    Write-Log "Copie forcée du contenu SharePoint des équipes..." 'STEP'
    $srcUrl = Get-GroupSiteUrl -GroupId $SourceTeamId -Connection $Connection
    $dstUrl = Get-GroupSiteUrl -GroupId $NewGroupId  -Connection $Connection
    Write-Log "Site source : $srcUrl" 'INFO'
    Write-Log "Site cible  : $dstUrl" 'INFO'

    $srcConn = Connect-Site -Url $srcUrl -Label 'site équipe source'
    $dstConn = Connect-Site -Url $dstUrl -Label 'site équipe cible'
    return Copy-LibrariesContent -SourceConn $srcConn -TargetConn $dstConn
}
#endregion

#region ░░ STREAM 8ter : Traitement par lot (CSV) ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# Lit un CSV d'opérations, exécute chacune comme un run complet (via
# Invoke-CopyOperation), et produit un résumé global exporté en CSV.
# Une opération en échec n'interrompt pas le lot.

function ConvertTo-Bool {
    # Interprète une valeur de cellule CSV comme un booléen tolérant.
    param($Value)
    return ("$Value").Trim() -match '^(1|true|vrai|yes|oui|y|o|x)$'
}

function Set-OperationContext {
    # Aligne les variables de script (lues par les fonctions métier) sur la ligne courante.
    param(
        [Parameter(Mandatory = $true)]$Op,
        [int]$Index = 1
    )
    # Les paramètres validés (ValidatePattern / ValidateSet) conservent leur contrainte
    # sur la VARIABLE : toute ré-assignation la re-déclenche. On la retire ici car le
    # contexte par opération réutilise ces variables avec des valeurs vides selon le mode
    # (ex. SourceUrl vide en mode Team, SourceTeamId vide en mode Site).
    foreach ($vn in 'SourceUrl', 'TargetUrl', 'TargetType', 'SourceTeamId', 'TenantUrl', 'Visibility') {
        $vv = Get-Variable -Name $vn -Scope script -ErrorAction SilentlyContinue
        if ($vv) { $vv.Attributes.Clear() }
    }

    $script:OpMode      = if ($Op.Mode) { "$($Op.Mode)".Trim() } else { 'Site' }
    $script:SourceUrl   = $Op.SourceUrl
    $script:TargetUrl   = $Op.TargetUrl
    $script:TargetTitle = $Op.TargetTitle
    $script:TargetType  = if ($Op.TargetType) { "$($Op.TargetType)".Trim() } else { 'CommunicationSite' }
    $script:Owner       = $Op.Owner
    $script:SourceTeamId = $Op.SourceTeamId
    $script:NewTeamName  = $Op.NewTeamName
    $script:TenantUrl    = $Op.TenantUrl
    $script:Visibility   = if ($Op.Visibility) { "$($Op.Visibility)".Trim() } else { 'Private' }
    # Booléens : valeur de ligne OU switch global (sécurité : -DryRun force tout le lot).
    $script:IncludeContent     = ConvertTo-Bool $Op.IncludeContent
    $script:IncludePermissions = (ConvertTo-Bool $Op.IncludePermissions) -or $script:GlobalAcl
    $script:DryRun             = (ConvertTo-Bool $Op.DryRun) -or $script:GlobalDryRun
    # Modèle PnP unique par opération (évite l'écrasement entre lignes).
    $script:TemplateFile = Join-Path $LogPath ("SiteTemplate_{0:yyyyMMdd_HHmmss}_op{1:D2}.pnp" -f $script:StartTime, $Index)
}

function Import-OperationCsv {
    # Lit et valide (minimalement) le CSV d'opérations.
    param([Parameter(Mandatory = $true)][string]$Path)
    $rows = Import-Csv -Path $Path
    if (-not $rows) { throw "CSV vide ou illisible : $Path" }

    $ops = [System.Collections.Generic.List[object]]::new()
    $n = 0
    foreach ($row in $rows) {
        $n++
        $mode = if ($row.PSObject.Properties.Name -contains 'Mode' -and $row.Mode) { "$($row.Mode)".Trim() } else { 'Site' }
        if ($mode -eq 'Team') {
            if (-not $row.SourceTeamId -or -not $row.NewTeamName -or -not $row.TenantUrl) {
                throw "Ligne $n (Team) : colonnes SourceTeamId, NewTeamName et TenantUrl requises."
            }
        }
        else {
            if (-not $row.SourceUrl -or -not $row.TargetUrl) {
                throw "Ligne $n (Site) : colonnes SourceUrl et TargetUrl requises."
            }
        }
        $ops.Add($row)
    }
    return $ops
}

function Invoke-CopyOperation {
    # Exécute UNE opération (Site ou Team) de bout en bout et renvoie un objet résultat.
    # Ne lève jamais : un échec est capturé et reporté dans le résultat (Status=FAILED).
    param(
        [Parameter(Mandatory = $true)]$Op,
        [int]$Index = 1,
        [int]$Count = 1
    )
    Set-OperationContext -Op $Op -Index $Index
    $aclLabel  = if ($IncludePermissions) { 'Copiées' } else { 'Non copiées' }
    $authLabel = if ($Thumbprint) { "Certificat$(if ($AppName) { " ($AppName)" })" } else { 'Interactif' }
    $result = [ordered]@{
        Index = $Index; Mode = $script:OpMode; Target = ''; Status = 'OK'
        Files = 0; Errors = 0; Message = ''
    }
    if ($Count -gt 1) { Write-Log ("══════ Opération {0}/{1} ({2}) ══════" -f $Index, $Count, $script:OpMode) 'STEP' }

    try {
        if ($script:OpMode -eq 'Team') {
            # ----- Opération TEAM -----
            $result.Target = $NewTeamName
            Write-Banner -Title "CLONAGE MICROSOFT TEAMS" -Fields ([ordered]@{
                'Équipe source' = $SourceTeamId
                'Nouvelle'      = $NewTeamName
                'Tenant'        = $TenantUrl
                'Visibilité'    = $Visibility
                'ACL (membres)' = $aclLabel
                'Auth'          = $authLabel
                'Mode'          = $(if ($DryRun) { 'DRYRUN (simulation)' } else { 'Réel' })
                'Log'           = $script:LogFile
            })
            Initialize-StepTracker -Activity 'Clonage Teams' -StepLabels @(
                'Pré-requis (module PnP)',
                'Connexion tenant (Graph)',
                'Clonage de l''équipe',
                'Attribution du propriétaire',
                'Copie forcée du contenu'
            )
            Invoke-Step 1 { Initialize-Prerequisites }
            $tenantConn = Invoke-Step 2 { Connect-Tenant -Url $TenantUrl }
            $teamResult = Invoke-Step 3 { Copy-Team -Connection $tenantConn }

            # Si owner et/ou copie forcée : on attend que l'équipe clonée soit disponible.
            $newGroup = $null
            $needGroup = ($Owner -or $ForceContentCopy) -and -not $DryRun -and $teamResult -and $teamResult.Status -eq 'Submitted'
            if ($needGroup) {
                $newGroup = Wait-ClonedGroup -Alias $teamResult.Alias -Connection $tenantConn
            }

            # Étape 4 : attribuer le propriétaire (owner + membre).
            if ($Owner -and $newGroup) {
                Invoke-Step 4 { Set-GroupOwner -GroupId $newGroup.id -OwnerUpn $Owner -Connection $tenantConn -Mute:$MuteNotifications }
            }
            else {
                Set-StepSkipped 4 $(if (-not $Owner) { 'aucun -Owner spécifié' } elseif ($DryRun) { 'DryRun' } else { 'clone non soumis' })
            }

            # Étape 5 : copie forcée et synchrone du contenu SharePoint.
            $teamCopyStats = $null
            if ($ForceContentCopy -and $newGroup) {
                $teamCopyStats = Invoke-Step 5 { Copy-TeamContent -SourceTeamId $SourceTeamId -NewGroupId $newGroup.id -Connection $tenantConn }
                if ($teamCopyStats) { $result.Files = $teamCopyStats.Files; $result.Errors = $teamCopyStats.Errors }
            }
            else {
                Set-StepSkipped 5 $(if (-not $ForceContentCopy) { '-ForceContentCopy non demandé' } elseif ($DryRun) { 'DryRun' } else { 'clone non soumis' })
            }

            Show-StepChecklist
            Write-Banner -Title "RAPPORT DE CLONAGE TEAMS" -Fields ([ordered]@{
                'Équipe source' = $(if ($teamResult) { $teamResult.Source } else { '-' })
                'Nouvelle'      = $NewTeamName
                'Éléments'      = $(if ($teamResult) { $teamResult.Parts } else { '-' })
                'Statut'        = $(if ($teamResult) { $teamResult.Status } else { '-' })
                'ACL (membres)' = $aclLabel
                'Contenu forcé' = $(if ($teamCopyStats) { "$($teamCopyStats.Files) fichier(s), $($teamCopyStats.Errors) erreur(s)" } else { 'non' })
            })
        }
        else {
            # ----- Opération SITE -----
            $result.Target = $TargetUrl
            Write-Banner -Title "DUPLICATION DE SITE SHAREPOINT" -Fields ([ordered]@{
                'Source'  = $SourceUrl
                'Cible'   = $TargetUrl
                'Type'    = $TargetType
                'Contenu' = $(if ($IncludeContent) { 'Inclus' } else { 'Structure seule' })
                'ACL'     = $aclLabel
                'Auth'    = $authLabel
                'Mode'    = $(if ($DryRun) { 'DRYRUN (simulation)' } else { 'Réel' })
                'Log'     = $script:LogFile
            })
            Initialize-StepTracker -Activity 'Duplication SharePoint' -StepLabels @(
                'Pré-requis (module PnP)',
                'Connexion au site source',
                'Extraction du modèle',
                'Provisioning du site cible',
                'Application du modèle',
                'Copie du contenu',
                'Vérification & rapport'
            )
            Invoke-Step 1 { Initialize-Prerequisites }
            $sourceConn = Invoke-Step 2 { Connect-Site -Url $SourceUrl -Label 'site source' }
            Invoke-Step 3 { Export-SourceTemplate -Connection $sourceConn }
            Invoke-Step 4 { New-TargetSite }
            $targetConn = Invoke-Step 5 { Invoke-TargetTemplate }

            $copyStats = $null
            if ($targetConn -and -not $DryRun) {
                $copyStats = Invoke-Step 6 { Copy-LibrariesContent -SourceConn $sourceConn -TargetConn $targetConn }
            }
            elseif ($DryRun) {
                $copyStats = Invoke-Step 6 { Copy-LibrariesContent -SourceConn $sourceConn -TargetConn $sourceConn }
            }
            else {
                Set-StepSkipped 6 'aucune connexion cible'
            }

            Invoke-Step 7 { Write-FinalReport -SourceConn $sourceConn -TargetConn $targetConn -CopyStats $copyStats }
            Show-StepChecklist

            if ($copyStats) { $result.Files = $copyStats.Files; $result.Errors = $copyStats.Errors }
        }
        $result.Status = if ($DryRun) { 'DRYRUN' } else { 'OK' }
    }
    catch {
        $result.Status  = 'FAILED'
        $result.Message = $_.Exception.Message
        Write-Log "ÉCHEC de l'opération $Index : $($_.Exception.Message)" 'ERROR'
        Write-Log "Trace : $($_.ScriptStackTrace)" 'DEBUG'
        Show-StepChecklist
    }
    finally {
        # Repart d'un contexte propre pour l'opération suivante.
        try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
    }
    return [pscustomobject]$result
}

function Show-BatchSummary {
    # Récap final du lot (console) + export CSV des résultats.
    param([Parameter(Mandatory = $true)]$Results)
    $bar = '═' * 64
    Write-Host ""
    Write-Host $bar -ForegroundColor Cyan
    Write-Host "  RÉSUMÉ DU TRAITEMENT PAR LOT" -ForegroundColor Cyan
    Write-Host $bar -ForegroundColor Cyan
    foreach ($r in $Results) {
        $c = switch ($r.Status) { 'OK' { 'Green' } 'DRYRUN' { 'Yellow' } 'FAILED' { 'Red' } default { 'Gray' } }
        Write-Host ("  #{0,-2} [{1,-6}] {2,-5} {3}" -f $r.Index, $r.Status, $r.Mode, $r.Target) -ForegroundColor $c
        if ($r.Status -eq 'FAILED' -and $r.Message) {
            Write-Host ("         └─ {0}" -f $r.Message) -ForegroundColor DarkGray
        }
    }
    $ok = ($Results | Where-Object { $_.Status -eq 'OK' }).Count
    $dr = ($Results | Where-Object { $_.Status -eq 'DRYRUN' }).Count
    $ko = ($Results | Where-Object { $_.Status -eq 'FAILED' }).Count
    Write-Host $bar -ForegroundColor Cyan
    Write-Host ("  Total {0}  |  OK {1}  |  DryRun {2}  |  Échecs {3}" -f $Results.Count, $ok, $dr, $ko)
    Write-Host $bar -ForegroundColor Cyan
    Write-Host ""

    $outCsv = Join-Path $LogPath ("BatchResult_{0:yyyyMMdd_HHmmss}.csv" -f $script:StartTime)
    $Results | Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
    Write-Log "Résultats du lot exportés : $outCsv" 'OK'
}
#endregion

#region ░░ STREAM 9 : Orchestration (Main) ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# Dispatch : run unique (Site/Team) OU traitement par lot (CSV). Transcript + erreurs.

try {
    Start-Transcript -Path (Join-Path $LogPath ("Transcript_{0:yyyyMMdd_HHmmss}.log" -f $script:StartTime)) -Force | Out-Null

    # Switches globaux : s'appliquent (OR) à chaque opération.
    $script:GlobalDryRun = $DryRun.IsPresent
    $script:GlobalAcl    = $IncludePermissions.IsPresent

    if ($PSCmdlet.ParameterSetName -eq 'Csv') {
        # ---------- MODE LOT ----------
        Write-Log "Mode LOT — lecture du CSV : $ConfigCsv" 'STEP'
        $ops = Import-OperationCsv -Path $ConfigCsv
        Write-Log ("{0} opération(s) à traiter." -f $ops.Count) 'INFO'

        $results = [System.Collections.Generic.List[object]]::new()
        $i = 0
        foreach ($op in $ops) {
            $i++
            $results.Add( (Invoke-CopyOperation -Op $op -Index $i -Count $ops.Count) )
        }

        Show-BatchSummary -Results $results
        $failed = ($results | Where-Object { $_.Status -eq 'FAILED' }).Count
        Write-Log ("=== Lot terminé ({0} échec(s)) ===" -f $failed) $(if ($failed) { 'WARN' } else { 'OK' })
        exit ([int]($failed -gt 0))
    }
    else {
        # ---------- RUN UNIQUE ----------
        $op = @{
            Mode               = $PSCmdlet.ParameterSetName   # 'Site' ou 'Team'
            SourceUrl          = $SourceUrl
            TargetUrl          = $TargetUrl
            TargetTitle        = $TargetTitle
            TargetType         = $TargetType
            Owner              = $Owner
            IncludeContent     = $IncludeContent.IsPresent
            SourceTeamId       = $SourceTeamId
            NewTeamName        = $NewTeamName
            TenantUrl          = $TenantUrl
            Visibility         = $Visibility
            IncludePermissions = $IncludePermissions.IsPresent
            DryRun             = $DryRun.IsPresent
        }
        $r = Invoke-CopyOperation -Op $op
        Write-Log "=== Terminé (statut: $($r.Status)) ===" $(if ($r.Status -eq 'FAILED') { 'ERROR' } else { 'OK' })
        exit ([int]($r.Status -eq 'FAILED'))
    }
}
catch {
    Write-Log "ÉCHEC GLOBAL : $($_.Exception.Message)" 'ERROR'
    Write-Log "Trace : $($_.ScriptStackTrace)" 'DEBUG'
    exit 1
}
finally {
    try { Stop-Transcript | Out-Null } catch {}
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
}
#endregion
