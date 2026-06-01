#Requires -Version 5.1
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
    (Optionnel) ClientId d'une App Registration Entra ID pour l'auth interactive PnP.
    Si absent, le ClientId par défaut de PnP.PowerShell est utilisé.

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

    [Parameter(Mandatory = $false, ParameterSetName = 'Site')]
    [string]$Owner,

    [Parameter(Mandatory = $false, ParameterSetName = 'Site')]
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

    # --- Paramètres communs aux deux modes ---
    # ACL : copie des permissions (groupes/rôles du site, ou membres/owners de l'équipe).
    [Parameter(Mandatory = $false)]
    [switch]$IncludePermissions,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

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

function Connect-Site {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [string]$Label = 'site'
    )
    Write-Log "Connexion interactive au $Label : $Url" 'STEP'
    $params = @{ Url = $Url; Interactive = $true; ReturnConnection = $true; ErrorAction = 'Stop' }
    if ($ClientId) { $params['ClientId'] = $ClientId }

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
    Write-Progress -Id 1 -Activity 'Duplication SharePoint' -Status 'Extraction du modèle source' -PercentComplete 20

    # ACL : on n'inclut le handler SiteSecurity (groupes, rôles, attributions) que si demandé.
    $handlerList = [System.Collections.Generic.List[string]]@(
        'Lists', 'Fields', 'ContentTypes', 'Pages', 'PageContents', 'Navigation',
        'RegionalSettings', 'SupportedUILanguages', 'Files', 'WebSettings', 'Publishing'
    )
    if ($IncludePermissions) {
        $handlerList.Add('SiteSecurity')
        Write-Log "ACL activées : permissions du site incluses (groupes/rôles)." 'INFO'
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
    Write-Progress -Id 1 -Activity 'Duplication SharePoint' -Status 'Création du site cible' -PercentComplete 45

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
    Write-Progress -Id 1 -Activity 'Duplication SharePoint' -Status 'Application du modèle' -PercentComplete 65

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
    Write-Progress -Id 1 -Activity 'Duplication SharePoint' -Status 'Copie des fichiers' -PercentComplete 80

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
    Write-Progress -Id 1 -Activity 'Duplication SharePoint' -Status 'Vérification' -PercentComplete 95

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

    Write-Progress -Id 1 -Activity 'Duplication SharePoint' -Completed
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
    Write-Log "Connexion interactive (Graph/Teams) : $Url" 'STEP'
    $params = @{ Url = $Url; Interactive = $true; ReturnConnection = $true; ErrorAction = 'Stop' }
    if ($ClientId) { $params['ClientId'] = $ClientId }
    $conn = Invoke-WithRetry -Operation 'connexion tenant' -Action { Connect-PnPOnline @params }
    Write-Log "Connecté au tenant pour Graph/Teams." 'OK'
    return $conn
}

function Copy-Team {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory = $true)]$Connection)

    Write-Log "Clonage de l'équipe Teams source ($SourceTeamId)..." 'STEP'
    Write-Progress -Id 1 -Activity 'Clonage Teams' -Status 'Préparation du clone' -PercentComplete 30

    # Vérifie que l'équipe source existe.
    $srcTeam = Invoke-WithRetry -Operation 'lecture équipe source' -Action {
        Invoke-PnPGraphMethod -Url "v1.0/teams/$SourceTeamId" -Method Get -Connection $Connection
    }
    Write-Log "Équipe source : '$($srcTeam.displayName)'." 'OK'

    # mailNickname : alias dérivé du nouveau nom (alphanumérique uniquement).
    $alias = ($NewTeamName -replace '[^a-zA-Z0-9]', '')
    if (-not $alias) { $alias = "team$($script:StartTime.ToString('yyyyMMddHHmmss'))" }

    # ACL : les membres/owners ne sont clonés que si demandé.
    $parts = @('apps', 'tabs', 'settings', 'channels')
    if ($IncludePermissions) {
        $parts += 'members'
        Write-Log "ACL activées : membres et owners de l'équipe inclus dans le clone." 'INFO'
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
        return [ordered]@{ Mode = 'Team'; Source = $srcTeam.displayName; NewTeam = $NewTeamName; Parts = $body.partsToClone; Status = 'DRYRUN' }
    }
    if (-not $PSCmdlet.ShouldProcess($NewTeamName, "Cloner l'équipe Teams $SourceTeamId")) { return }

    Write-Progress -Id 1 -Activity 'Clonage Teams' -Status 'Envoi de la demande de clone' -PercentComplete 60
    # Le clone est asynchrone : Graph renvoie une opération (teamsAsyncOperation).
    Invoke-WithRetry -Operation 'clone équipe' -Action {
        Invoke-PnPGraphMethod -Url "v1.0/teams/$SourceTeamId/clone" -Method Post -Content $body -Connection $Connection
    } | Out-Null

    Write-Log "Demande de clonage envoyée (traitement asynchrone côté Microsoft 365)." 'OK'
    Write-Log "La nouvelle équipe '$NewTeamName' apparaîtra dans Teams sous quelques minutes." 'INFO'

    return [ordered]@{ Mode = 'Team'; Source = $srcTeam.displayName; NewTeam = $NewTeamName; Parts = $body.partsToClone; Status = 'Submitted' }
}
#endregion

#region ░░ STREAM 9 : Orchestration (Main) ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░
# Enchaîne les streams, gère le transcript et les erreurs globales.

try {
    Start-Transcript -Path (Join-Path $LogPath ("Transcript_{0:yyyyMMdd_HHmmss}.log" -f $script:StartTime)) -Force | Out-Null

    Write-Log "=== Démarrage — mode $($PSCmdlet.ParameterSetName) | DryRun: $DryRun | ACL: $IncludePermissions ===" 'STEP'
    Initialize-Prerequisites

    if ($PSCmdlet.ParameterSetName -eq 'Team') {
        # --- Mode TEAM : clonage d'une équipe Microsoft Teams via Graph ---
        $tenantConn = Connect-Tenant -Url $TenantUrl
        $teamResult = Copy-Team -Connection $tenantConn

        Write-Host ""
        $bar = '═' * 64
        Write-Host $bar -ForegroundColor Cyan
        Write-Host "  RAPPORT DE CLONAGE TEAMS" -ForegroundColor Cyan
        Write-Host $bar -ForegroundColor Cyan
        if ($teamResult) {
            Write-Host ("  Équipe source : {0}" -f $teamResult.Source)
            Write-Host ("  Nouvelle      : {0}" -f $teamResult.NewTeam)
            Write-Host ("  Éléments      : {0}" -f $teamResult.Parts)
            Write-Host ("  Statut        : {0}" -f $teamResult.Status)
        }
        Write-Host ("  ACL (membres) : {0}" -f $(if ($IncludePermissions) { 'Copiées' } else { 'Non copiées' }))
        Write-Host ("  Log           : {0}" -f $script:LogFile)
        Write-Host $bar -ForegroundColor Cyan
        Write-Host ""
        Write-Progress -Id 1 -Activity 'Clonage Teams' -Completed
    }
    else {
        # --- Mode SITE : duplication d'un site SharePoint ---
        Write-Log ("Source: {0} | Cible: {1}" -f $SourceUrl, $TargetUrl) 'INFO'

        $sourceConn = Connect-Site -Url $SourceUrl -Label 'site source'
        Export-SourceTemplate -Connection $sourceConn

        New-TargetSite
        $targetConn = Invoke-TargetTemplate

        $copyStats = $null
        if ($targetConn -and -not $DryRun) {
            $copyStats = Copy-LibrariesContent -SourceConn $sourceConn -TargetConn $targetConn
        }
        elseif ($DryRun) {
            $copyStats = Copy-LibrariesContent -SourceConn $sourceConn -TargetConn $sourceConn
        }

        Write-FinalReport -SourceConn $sourceConn -TargetConn $targetConn -CopyStats $copyStats
    }

    Write-Log "=== Terminé avec succès ===" 'OK'
    exit 0
}
catch {
    Write-Log "ÉCHEC GLOBAL : $($_.Exception.Message)" 'ERROR'
    Write-Log "Trace : $($_.ScriptStackTrace)" 'DEBUG'
    Write-Progress -Id 1 -Activity 'Duplication SharePoint' -Completed
    exit 1
}
finally {
    try { Stop-Transcript | Out-Null } catch {}
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
}
#endregion
