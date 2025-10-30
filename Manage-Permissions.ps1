<#
.SYNOPSIS
    Gestion avancée permissions SharePoint Online
    
.DESCRIPTION
    Script pour auditer et gérer les permissions:
    - Audit complet permissions
    - Détection permissions excessives
    - Identification partages externes
    - Nettoyage permissions obsolètes
    - Rapports conformité
    
.PARAMETER SiteUrl
    URL du site SharePoint
    
.PARAMETER Action
    Action: Audit, Report, RemoveExternal, Cleanup
    
.EXAMPLE
    .\Manage-Permissions.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/finance" -Action Report -ReportPath "C:\Reports"
    
.NOTES
    Auteur: Teguy EKANZA
    Version: 1.0
    Date: 2025-10-23
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [ValidateSet("Audit","Report","RemoveExternal","Cleanup")]
    [string]$Action,
    
    [Parameter(Mandatory=$false)]
    [string]$ReportPath = "C:\PermissionsReports"
)

$ErrorActionPreference = "Continue"
$Global:PermissionsData = @()
$Global:ExternalSharing = @()
$Global:ExcessivePermissions = @()

function Write-PermissionLog {
    param([string]$Message, [string]$Level = "INFO")
    
    $colors = @{
        "INFO" = "Cyan"
        "WARNING" = "Yellow"
        "ERROR" = "Red"
        "SUCCESS" = "Green"
    }
    
    Write-Host "[$Level] $Message" -ForegroundColor $colors[$Level]
}

function Get-SitePermissionReport {
    param([string]$Url)
    
    try {
        Connect-PnPOnline -Url $Url -Interactive
        
        Write-PermissionLog "Analyse permissions: $Url" -Level INFO
        
        $web = Get-PnPWeb
        $roleAssignments = Get-PnPProperty -ClientObject $web -Property RoleAssignments
        
        foreach ($assignment in $roleAssignments) {
            Get-PnPProperty -ClientObject $assignment -Property RoleDefinitionBindings, Member
            
            $member = $assignment.Member
            $roles = $assignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name
            
            $hasFullControl = $roles -contains "Full Control"
            $isExternal = $member.LoginName -like "*#ext#*"
            
            $permissionEntry = [PSCustomObject]@{
                SiteUrl = $Url
                PrincipalType = $member.PrincipalType
                PrincipalName = $member.Title
                LoginName = $member.LoginName
                Email = if($member.Email) { $member.Email } else { "N/A" }
                Roles = $roles -join ", "
                HasFullControl = $hasFullControl
                IsExternal = $isExternal
                IsGroup = $member.PrincipalType -eq "SharePointGroup"
            }
            
            $Global:PermissionsData += $permissionEntry
            
            if ($hasFullControl) {
                $Global:ExcessivePermissions += $permissionEntry
            }
            
            if ($isExternal) {
                $Global:ExternalSharing += $permissionEntry
            }
        }
        
        $lists = Get-PnPList | Where-Object { $_.Hidden -eq $false }
        
        foreach ($list in $lists) {
            if ($list.HasUniqueRoleAssignments) {
                Write-PermissionLog "  Permissions uniques: $($list.Title)" -Level WARNING
                
                $listAssignments = Get-PnPProperty -ClientObject $list -Property RoleAssignments
                
                foreach ($assignment in $listAssignments) {
                    Get-PnPProperty -ClientObject $assignment -Property RoleDefinitionBindings, Member
                    
                    $member = $assignment.Member
                    $roles = $assignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name
                    
                    $permissionEntry = [PSCustomObject]@{
                        SiteUrl = $Url
                        Location = "List: $($list.Title)"
                        PrincipalType = $member.PrincipalType
                        PrincipalName = $member.Title
                        LoginName = $member.LoginName
                        Email = if($member.Email) { $member.Email } else { "N/A" }
                        Roles = $roles -join ", "
                        HasFullControl = $roles -contains "Full Control"
                        IsExternal = $member.LoginName -like "*#ext#*"
                        IsGroup = $member.PrincipalType -eq "SharePointGroup"
                    }
                    
                    $Global:PermissionsData += $permissionEntry
                }
            }
        }
        
        Write-PermissionLog "✓ Analyse terminée" -Level SUCCESS
        Write-PermissionLog "  Total: $($Global:PermissionsData.Count)" -Level INFO
        Write-PermissionLog "  Full Control: $($Global:ExcessivePermissions.Count)" -Level WARNING
        Write-PermissionLog "  Externes: $($Global:ExternalSharing.Count)" -Level WARNING
    }
    catch {
        Write-PermissionLog "Erreur: $($_.Exception.Message)" -Level ERROR
    }
}

function Remove-ExternalSharingLinks {
    param([string]$Url)
    
    try {
        Connect-PnPOnline -Url $Url -Interactive
        
        Write-PermissionLog "Suppression partages externes..." -Level INFO
        
        $users = Get-PnPUser | Where-Object { $_.LoginName -like "*#ext#*" }
        
        Write-PermissionLog "Externes trouvés: $($users.Count)" -Level WARNING
        
        foreach ($user in $users) {
            try {
                Remove-PnPUser -Identity $user.LoginName -Force
                Write-PermissionLog "  ✓ Supprimé: $($user.Title)" -Level SUCCESS
            }
            catch {
                Write-PermissionLog "  ✗ Erreur: $($user.Title)" -Level ERROR
            }
        }
        
        Write-PermissionLog "✓ Nettoyage terminé" -Level SUCCESS
    }
    catch {
        Write-PermissionLog "Erreur: $($_.Exception.Message)" -Level ERROR
    }
}

function Export-PermissionReports {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    $reportFile = Join-Path $Path "Permissions_Complete_$timestamp.csv"
    $Global:PermissionsData | Export-Csv -Path $reportFile -NoTypeInformation -Encoding UTF8
    Write-PermissionLog "✓ Rapport: $reportFile" -Level SUCCESS
    
    if ($Global:ExcessivePermissions.Count -gt 0) {
        $fcFile = Join-Path $Path "Permissions_FullControl_$timestamp.csv"
        $Global:ExcessivePermissions | Export-Csv -Path $fcFile -NoTypeInformation -Encoding UTF8
        Write-PermissionLog "✓ Full Control: $fcFile" -Level WARNING
    }
    
    if ($Global:ExternalSharing.Count -gt 0) {
        $extFile = Join-Path $Path "Permissions_External_$timestamp.csv"
        $Global:ExternalSharing | Export-Csv -Path $extFile -NoTypeInformation -Encoding UTF8
        Write-PermissionLog "✓ Externes: $extFile" -Level WARNING
    }
    
    $summary = @"
========================================
RAPPORT PERMISSIONS SHAREPOINT
========================================
Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Site: $SiteUrl

STATISTIQUES:
- Total attributions: $($Global:PermissionsData.Count)
- Groupes: $(($Global:PermissionsData | Where-Object {$_.IsGroup}).Count)
- Utilisateurs: $(($Global:PermissionsData | Where-Object {-not $_.IsGroup}).Count)

ALERTES:
- Full Control: $($Global:ExcessivePermissions.Count)
  $(if($Global:ExcessivePermissions.Count -gt 3){"⚠️ Limiter à 2-3 personnes max"})
  
- Partages externes: $($Global:ExternalSharing.Count)
  $(if($Global:ExternalSharing.Count -gt 0){"⚠️ Vérifier conformité"})

RECOMMANDATIONS:
1. Limiter Full Control à 2-3 admin
2. Utiliser groupes vs utilisateurs individuels
3. Réviser partages externes régulièrement
4. Minimiser permissions uniques
5. Documenter exceptions

FICHIERS:
- Complet: $reportFile
$(if($Global:ExcessivePermissions.Count -gt 0){"- Full Control: $fcFile"})
$(if($Global:ExternalSharing.Count -gt 0){"- Externes: $extFile"})

========================================
"@
    
    $summaryFile = Join-Path $Path "Summary_$timestamp.txt"
    $summary | Out-File -FilePath $summaryFile -Encoding UTF8
    
    Write-Host "`n$summary" -ForegroundColor Cyan
}

try {
    Write-PermissionLog "========================================" -Level INFO
    Write-PermissionLog "GESTION PERMISSIONS SHAREPOINT" -Level INFO
    Write-PermissionLog "========================================" -Level INFO
    Write-PermissionLog "Site: $SiteUrl" -Level INFO
    Write-PermissionLog "Action: $Action" -Level INFO
    
    switch ($Action) {
        "Audit" {
            Get-SitePermissionReport -Url $SiteUrl
        }
        
        "Report" {
            Get-SitePermissionReport -Url $SiteUrl
            Export-PermissionReports -Path $ReportPath
        }
        
        "RemoveExternal" {
            Get-SitePermissionReport -Url $SiteUrl
            
            if ($Global:ExternalSharing.Count -gt 0) {
                Write-PermissionLog "`n⚠️  $($Global:ExternalSharing.Count) externes détectés" -Level WARNING
                $confirm = Read-Host "Confirmer suppression? (O/N)"
                
                if ($confirm -eq "O") {
                    Remove-ExternalSharingLinks -Url $SiteUrl
                }
            }
            else {
                Write-PermissionLog "Aucun externe détecté" -Level SUCCESS
            }
            
            Export-PermissionReports -Path $ReportPath
        }
        
        "Cleanup" {
            Get-SitePermissionReport -Url $SiteUrl
            Write-PermissionLog "`nAnalyse nettoyage..." -Level INFO
            Export-PermissionReports -Path $ReportPath
        }
    }
    
    Write-PermissionLog "`n✓ Opération terminée" -Level SUCCESS
}
catch {
    Write-PermissionLog "ERREUR: $($_.Exception.Message)" -Level ERROR
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
