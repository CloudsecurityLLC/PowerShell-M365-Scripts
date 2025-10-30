<#
.SYNOPSIS
    Script d'audit complet SharePoint Online pour préparation migration
    
.DESCRIPTION
    Ce script analyse en profondeur l'environnement SharePoint Online:
    - Inventaire de tous les sites (sites collection, sous-sites)
    - Analyse des volumes de données par site
    - Analyse des permissions et groupes
    - Identification des contenus obsolètes
    - Rapport détaillé pour planification migration
    
.PARAMETER TenantUrl
    URL du tenant SharePoint (ex: https://contoso.sharepoint.com)
    
.PARAMETER ReportPath
    Chemin du dossier pour les rapports générés
    
.PARAMETER AnalyzePermissions
    Switch pour inclure l'analyse détaillée des permissions
    
.EXAMPLE
    .\Audit-SharePointEnvironment.ps1 -TenantUrl "https://contoso.sharepoint.com" -ReportPath "C:\Reports" -AnalyzePermissions
    
.NOTES
    Auteur: Teguy EKANZA - Consultant Expert M365
    Version: 1.0
    Date: 2025-10-23
    Prérequis: Module PnP.PowerShell (Install-Module PnP.PowerShell -Scope CurrentUser)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$TenantUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$ReportPath,
    
    [Parameter(Mandatory=$false)]
    [switch]$AnalyzePermissions
)

$ErrorActionPreference = "Continue"
$Global:LogFile = Join-Path $ReportPath "AuditLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$Global:SitesReport = @()
$Global:PermissionsReport = @()
$Global:StorageReport = @()

#region Fonctions

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor Cyan }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
    }
    
    Add-Content -Path $Global:LogFile -Value $logMessage
}

function Get-SiteCollectionDetails {
    param([string]$SiteUrl)
    
    try {
        Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
        $site = Get-PnPSite -Includes Usage, Owner
        $web = Get-PnPWeb -Includes Created, LastItemModifiedDate
        
        $lists = Get-PnPList -Includes ItemCount
        $totalItems = ($lists | Measure-Object -Property ItemCount -Sum).Sum
        
        $daysSinceLastModified = ((Get-Date) - $web.LastItemModifiedDate).Days
        
        $siteDetails = [PSCustomObject]@{
            SiteUrl = $SiteUrl
            Title = $web.Title
            Owner = if($site.Owner) { $site.Owner.Email } else { "Non défini" }
            Created = $web.Created
            LastModified = $web.LastItemModifiedDate
            DaysSinceLastModified = $daysSinceLastModified
            StorageUsedMB = [math]::Round($site.Usage.Storage / 1MB, 2)
            TotalLists = $lists.Count
            TotalItems = $totalItems
            Status = if($daysSinceLastModified -gt 180) { "Inactif" } elseif($daysSinceLastModified -gt 90) { "Peu actif" } else { "Actif" }
        }
        
        $Global:SitesReport += $siteDetails
        Write-Log "✓ Analysé: $($web.Title)" -Level SUCCESS
        
        return $siteDetails
    }
    catch {
        Write-Log "Erreur analyse $SiteUrl : $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Get-SitePermissions {
    param([string]$SiteUrl)
    
    try {
        Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
        $web = Get-PnPWeb
        $roleAssignments = Get-PnPProperty -ClientObject $web -Property RoleAssignments
        
        foreach ($roleAssignment in $roleAssignments) {
            Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings, Member
            
            $member = $roleAssignment.Member
            $roles = $roleAssignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name
            
            $permissionDetails = [PSCustomObject]@{
                SiteUrl = $SiteUrl
                PrincipalType = $member.PrincipalType
                PrincipalName = $member.Title
                LoginName = $member.LoginName
                Roles = $roles -join ", "
                IsGroup = $member.PrincipalType -eq "SharePointGroup"
            }
            
            $Global:PermissionsReport += $permissionDetails
        }
        
        Write-Log "Permissions analysées: $SiteUrl" -Level SUCCESS
    }
    catch {
        Write-Log "Erreur permissions $SiteUrl : $($_.Exception.Message)" -Level ERROR
    }
}

#endregion

#region Programme principal

try {
    if (-not (Test-Path $ReportPath)) {
        New-Item -ItemType Directory -Path $ReportPath -Force | Out-Null
        Write-Log "Dossier rapports créé: $ReportPath" -Level INFO
    }
    
    Write-Log "========================================" -Level INFO
    Write-Log "DÉBUT AUDIT SHAREPOINT ONLINE" -Level INFO
    Write-Log "========================================" -Level INFO
    Write-Log "Tenant: $TenantUrl" -Level INFO
    
    $adminUrl = $TenantUrl -replace "\.sharepoint\.com", "-admin.sharepoint.com"
    Write-Log "Connexion: $adminUrl" -Level INFO
    Connect-PnPOnline -Url $adminUrl -Interactive
    
    Write-Log "Récupération liste des sites..." -Level INFO
    $allSites = Get-PnPTenantSite -IncludeOneDriveSites:$false | Where-Object { $_.Template -notlike "*SRCHCEN*" }
    Write-Log "Sites trouvés: $($allSites.Count)" -Level SUCCESS
    
    $counter = 0
    foreach ($site in $allSites) {
        $counter++
        $percentComplete = [math]::Round(($counter / $allSites.Count) * 100, 2)
        Write-Progress -Activity "Analyse sites SharePoint" -Status "Site $counter sur $($allSites.Count)" -PercentComplete $percentComplete
        
        Get-SiteCollectionDetails -SiteUrl $site.Url
        
        if ($AnalyzePermissions) {
            Get-SitePermissions -SiteUrl $site.Url
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    Write-Progress -Activity "Analyse sites SharePoint" -Completed
    
    Write-Log "========================================" -Level INFO
    Write-Log "GÉNÉRATION RAPPORTS" -Level INFO
    Write-Log "========================================" -Level INFO
    
    $sitesReportFile = Join-Path $ReportPath "Sites_Inventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $Global:SitesReport | Export-Csv -Path $sitesReportFile -NoTypeInformation -Encoding UTF8
    Write-Log "✓ Rapport sites: $sitesReportFile" -Level SUCCESS
    
    if ($AnalyzePermissions) {
        $permissionsReportFile = Join-Path $ReportPath "Permissions_Analysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $Global:PermissionsReport | Export-Csv -Path $permissionsReportFile -NoTypeInformation -Encoding UTF8
        Write-Log "✓ Rapport permissions: $permissionsReportFile" -Level SUCCESS
    }
    
    $totalStorage = ($Global:SitesReport | Measure-Object -Property StorageUsedMB -Sum).Sum
    $inactiveSites = ($Global:SitesReport | Where-Object { $_.Status -eq "Inactif" }).Count
    $activeSites = ($Global:SitesReport | Where-Object { $_.Status -eq "Actif" }).Count
    
    $summary = @"
========================================
RAPPORT DE SYNTHÈSE
========================================
Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Tenant: $TenantUrl

STATISTIQUES:
- Total sites: $($Global:SitesReport.Count)
- Sites actifs: $activeSites
- Sites inactifs: $inactiveSites
- Stockage total: $([math]::Round($totalStorage / 1024, 2)) GB

RECOMMANDATIONS:
1. Archiver les $inactiveSites sites inactifs
2. Réviser les permissions sur sites sensibles
3. Planifier migration sites actifs en priorité
4. Mettre en place politique de rétention

FICHIERS GÉNÉRÉS:
- Log: $Global:LogFile
- Sites: $sitesReportFile
$(if($AnalyzePermissions){"- Permissions: $permissionsReportFile"})

========================================
"@
    
    $summaryFile = Join-Path $ReportPath "Summary_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $summary | Out-File -FilePath $summaryFile -Encoding UTF8
    
    Write-Host "`n$summary" -ForegroundColor Green
    Write-Log "✓ Rapport synthèse: $summaryFile" -Level SUCCESS
    
    Write-Log "========================================" -Level SUCCESS
    Write-Log "AUDIT TERMINÉ AVEC SUCCÈS" -Level SUCCESS
    Write-Log "========================================" -Level SUCCESS
}
catch {
    Write-Log "ERREUR CRITIQUE: $($_.Exception.Message)" -Level ERROR
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

#endregion
