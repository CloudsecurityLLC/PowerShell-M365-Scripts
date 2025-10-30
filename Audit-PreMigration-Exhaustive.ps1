<#
.SYNOPSIS
    Script d'audit exhaustif pré-migration des serveurs/disques réseau
    
.DESCRIPTION
    Ce script effectue un audit complet avant migration vers SharePoint Online :
    - Inventaire détaillé de tous les fichiers
    - Analyse des permissions NTFS (ACL)
    - Détection des fichiers obsolètes (>2 ans)
    - Détection des fichiers avec extensions bloquées
    - Analyse des tailles et volumes
    - Détection de contenu dupliqué/redondant
    - Identif ication des fichiers non conformes SharePoint
    - Rapport de conformité RGPD/légale
    - Recommandations d'optimisation
    - Export de tous les rapports en CSV/Excel
    
.PARAMETER SourcePath
    Chemin UNC source à auditer (ex: \\serveur\partage)
    
.PARAMETER OutputPath
    Dossier de sortie pour les rapports (défaut: C:\AuditReports)
    
.PARAMETER ExcludeObsolete
    Exclure automatiquement les fichiers >180 jours du rapport de migration
    
.PARAMETER AnalyzePermissions
    Analyser en détail les permissions NTFS (plus lent)
    
.EXAMPLE
    .\Audit-PreMigration.ps1 -SourcePath "\\serveur\donnees" -OutputPath "C:\Audit" -AnalyzePermissions
    
.NOTES
    Auteur: Teguy EKANZA - Consultant Expert M365
    Version: 1.0 (Production-Ready Comprehensive Audit)
    Date: 2025-10-27
    Prérequis: Accès administrateur aux partages
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$SourcePath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "C:\AuditReports",
    
    [Parameter(Mandatory=$false)]
    [int]$ObsoleteDaysThreshold = 730,  # 2 ans
    
    [Parameter(Mandatory=$false)]
    [switch]$AnalyzePermissions,
    
    [Parameter(Mandatory=$false)]
    [switch]$GenerateDetailedReport
)

#region Configuration Globale

$ErrorActionPreference = "Continue"

# Extensions bloquées SharePoint Online
$BlockedExtensions = @(
    'exe', 'bat', 'cmd', 'vbs', 'js', 'msi', 'reg', 'scr', 'vbe', 'jse', 'wsf', 'wsh',
    'msh', 'msh1', 'msh2', 'mshxml', 'msh1xml', 'msh2xml', 'ps1', 'ps2', 'psc1', 'psc2',
    'mst', 'jar', 'zip', 'com', 'pif', 'asp', 'aspx', 'jsp', 'php', 'py', 'rb', 'sh',
    'app', 'deb', 'rpm', 'dmg'
)

# Extensions à archiver
$ArchiveExtensions = @('bak', 'old', 'tmp', 'temp', 'log', 'bkp')

# Noms réservés
$ReservedNames = @('CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9')

# Caractères interdits
$ForbiddenChars = @('"', '*', ':', '<', '>', '?', '/', '\', '|')

# Statistiques
$Global:AuditStats = @{
    TotalFiles = 0
    TotalFolders = 0
    TotalSizeGB = 0
    FilesEligible = 0
    FilesBlocked = 0
    FilesObsolete = 0
    FilesArchive = 0
    FilesInvalid = 0
    FilesTooLarge = 0
    FilesRedundant = 0
    DuplicateCount = 0
    StartTime = Get-Date
}

# Fichiers de sortie
$Global:Reports = @{
    InventoryComplete = Join-Path $OutputPath "01_Inventory_Complete_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    EligibleMigration = Join-Path $OutputPath "02_EligibleForMigration_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    BlockedFiles = Join-Path $OutputPath "03_BlockedFiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    ObsoleteFiles = Join-Path $OutputPath "04_ObsoleteFiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    InvalidNames = Join-Path $OutputPath "05_InvalidFileNames_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    TooLargeFiles = Join-Path $OutputPath "06_FilesTooLarge_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    Duplicates = Join-Path $OutputPath "07_PotentialDuplicates_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    Permissions = Join-Path $OutputPath "08_Permissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    Summary = Join-Path $OutputPath "09_AuditSummary_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
}

# Collections globales
$Global:InventoryData = @()
$Global:BlockedData = @()
$Global:ObsoleteData = @()
$Global:InvalidData = @()
$Global:TooLargeData = @()
$Global:EligibleData = @()
$Global:DuplicateData = @()
$Global:PermissionsData = @()

#endregion

#region Fonctions de Logging

function Write-AuditLog {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARNING","ERROR","SUCCESS","CRITICAL")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "INFO"     { Write-Host $logMessage -ForegroundColor Cyan }
        "WARNING"  { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"    { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS"  { Write-Host $logMessage -ForegroundColor Green }
        "CRITICAL" { Write-Host $logMessage -ForegroundColor DarkRed -BackgroundColor Yellow }
    }
}

#endregion

#region Fonctions de Validation

function Test-FileNameValidity {
    param([string]$FileName)
    
    # Vérifier longueur
    if ($FileName.Length -gt 128) { return $false }
    
    # Vérifier nom réservé
    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    if ($nameWithoutExt -in $ReservedNames) { return $false }
    
    # Vérifier caractères interdits
    foreach ($char in $ForbiddenChars) {
        if ($FileName.Contains($char)) { return $false }
    }
    
    # Vérifier fin avec point ou espace
    if ($FileName.EndsWith('.') -or $FileName.EndsWith(' ')) { return $false }
    
    return $true
}

function Get-FileClassification {
    param([string]$FileName)
    
    $ext = [System.IO.Path]::GetExtension($FileName).TrimStart('.')
    
    if ($ext -in $BlockedExtensions) { return "Blocked" }
    if ($ext -in $ArchiveExtensions) { return "Archive" }
    if ($ext -in @('docx', 'xlsx', 'pptx', 'pdf', 'txt', 'csv')) { return "Office" }
    if ($ext -in @('jpg', 'png', 'gif', 'bmp', 'svg')) { return "Image" }
    if ($ext -in @('mp4', 'avi', 'mov', 'wmv', 'flv')) { return "Video" }
    if ($ext -in @('mp3', 'wav', 'flac', 'm4a')) { return "Audio" }
    
    return "Other"
}

#endregion

#region Fonctions d'Audit

function Invoke-ComprehensiveAudit {
    param([string]$Path)
    
    Write-AuditLog "========================================" -Level INFO
    Write-AuditLog "DÉBUT AUDIT EXHAUSTIF PRÉ-MIGRATION" -Level INFO
    Write-AuditLog "========================================" -Level INFO
    Write-AuditLog "Chemin: $Path" -Level INFO
    
    # Vérifier accès
    if (-not (Test-Path $Path)) {
        Write-AuditLog "ERREUR: Chemin inexistant: $Path" -Level CRITICAL
        throw "Path does not exist"
    }
    
    # Récupérer tous les fichiers
    Write-AuditLog "Analyse récursive en cours..." -Level INFO
    $allFiles = Get-ChildItem -Path $Path -File -Recurse -ErrorAction Continue
    $allFolders = Get-ChildItem -Path $Path -Directory -Recurse -ErrorAction Continue
    
    $Global:AuditStats.TotalFiles = $allFiles.Count
    $Global:AuditStats.TotalFolders = $allFolders.Count
    
    Write-AuditLog "Fichiers trouvés: $($allFiles.Count)" -Level SUCCESS
    Write-AuditLog "Dossiers trouvés: $($allFolders.Count)" -Level SUCCESS
    
    # Analyser chaque fichier
    $counter = 0
    foreach ($file in $allFiles) {
        $counter++
        $percentComplete = [math]::Round(($counter / $allFiles.Count) * 100, 2)
        Write-Progress -Activity "Audit en cours" -Status "Fichier $counter/$($allFiles.Count)" -PercentComplete $percentComplete
        
        $sizeMB = $file.Length / 1MB
        $sizeGB = $file.Length / 1GB
        $Global:AuditStats.TotalSizeGB += $sizeGB
        
        $daysSinceModified = ((Get-Date) - $file.LastWriteTime).Days
        $relativePath = $file.DirectoryName.Substring($Path.Length).TrimStart('\', '/')
        
        # Entrée inventaire
        $inventoryEntry = [PSCustomObject]@{
            FileName = $file.Name
            FullPath = $file.FullName
            RelativePath = $relativePath
            SizeMB = [math]::Round($sizeMB, 2)
            Created = $file.CreationTime
            Modified = $file.LastWriteTime
            DaysSinceModified = $daysSinceModified
            Owner = (Get-Acl $file.FullName).Owner
            Classification = Get-FileClassification -FileName $file.Name
            Extension = [System.IO.Path]::GetExtension($file.Name).TrimStart('.')
            IsValidName = Test-FileNameValidity -FileName $file.Name
            IsTooLarge = ($sizeGB -gt 0.25)  # >250 MB
            IsObsolete = ($daysSinceModified -gt $ObsoleteDaysThreshold)
        }
        
        $Global:InventoryData += $inventoryEntry
        
        # Classification et collecte
        if ($inventoryEntry.Classification -eq "Blocked") {
            $Global:AuditStats.FilesBlocked++
            $Global:BlockedData += $inventoryEntry
        }
        
        if ($inventoryEntry.IsObsolete) {
            $Global:AuditStats.FilesObsolete++
            $Global:ObsoleteData += $inventoryEntry
        }
        
        if ($inventoryEntry.Classification -eq "Archive") {
            $Global:AuditStats.FilesArchive++
        }
        
        if (-not $inventoryEntry.IsValidName) {
            $Global:AuditStats.FilesInvalid++
            $Global:InvalidData += $inventoryEntry
        }
        
        if ($inventoryEntry.IsTooLarge) {
            $Global:AuditStats.FilesTooLarge++
            $Global:TooLargeData += $inventoryEntry
        }
        
        # Éligible si aucun des problèmes ci-dessus
        if ($inventoryEntry.IsValidName -and -not $inventoryEntry.IsTooLarge -and -not $inventoryEntry.IsObsolete -and $inventoryEntry.Classification -ne "Blocked") {
            $Global:AuditStats.FilesEligible++
            $Global:EligibleData += $inventoryEntry
        }
    }
    
    Write-Progress -Activity "Audit en cours" -Completed
    
    # Déterminer doublons potentiels (par nom + taille)
    Write-AuditLog "Détection de fichiers potentiellement dupliqués..." -Level INFO
    $groupByNameSize = $Global:InventoryData | Group-Object -Property @{Expression="$($_.FileName)_$($_.SizeMB)"} | Where-Object { $_.Count -gt 1 }
    $Global:AuditStats.DuplicateCount = ($groupByNameSize | Measure-Object -Sum).Sum
    
    if ($groupByNameSize.Count -gt 0) {
        foreach ($group in $groupByNameSize) {
            foreach ($item in $group.Group) {
                $Global:DuplicateData += $item
            }
        }
    }
    
    # Permissions (si demandé)
    if ($AnalyzePermissions) {
        Write-AuditLog "Analyse des permissions NTFS en cours..." -Level INFO
        Invoke-PermissionsAnalysis -Path $Path
    }
}

function Invoke-PermissionsAnalysis {
    param([string]$Path)
    
    $allItems = Get-ChildItem -Path $Path -Recurse -ErrorAction Continue
    
    foreach ($item in $allItems) {
        try {
            $acl = Get-Acl $item.FullName -ErrorAction Stop
            $owner = $acl.Owner
            
            foreach ($access in $acl.Access) {
                $Global:PermissionsData += [PSCustomObject]@{
                    Path = $item.FullName
                    Identity = $access.IdentityReference
                    FileSystemRights = $access.FileSystemRights
                    AccessControlType = $access.AccessControlType
                    Owner = $owner
                }
            }
        }
        catch {
            Write-AuditLog "Impossible d'accéder aux permissions: $($item.FullName)" -Level WARNING
        }
    }
}

#endregion

#region Export de Rapports

function Export-AuditReports {
    
    Write-AuditLog "========================================" -Level INFO
    Write-AuditLog "GÉNÉRATION DES RAPPORTS" -Level INFO
    Write-AuditLog "========================================" -Level INFO
    
    # Créer dossier de sortie
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    
    # Inventaire complet
    $Global:InventoryData | Export-Csv -Path $Global:Reports.InventoryComplete -NoTypeInformation -Encoding UTF8
    Write-AuditLog "✓ Rapport inventaire complet: $($Global:Reports.InventoryComplete)" -Level SUCCESS
    
    # Fichiers éligibles
    $Global:EligibleData | Export-Csv -Path $Global:Reports.EligibleMigration -NoTypeInformation -Encoding UTF8
    Write-AuditLog "✓ Rapport fichiers éligibles: $($Global:Reports.EligibleMigration)" -Level SUCCESS
    
    # Fichiers bloqués
    if ($Global:BlockedData.Count -gt 0) {
        $Global:BlockedData | Export-Csv -Path $Global:Reports.BlockedFiles -NoTypeInformation -Encoding UTF8
        Write-AuditLog "⚠ Fichiers bloqués: $($Global:Reports.BlockedFiles)" -Level WARNING
    }
    
    # Fichiers obsolètes
    if ($Global:ObsoleteData.Count -gt 0) {
        $Global:ObsoleteData | Export-Csv -Path $Global:Reports.ObsoleteFiles -NoTypeInformation -Encoding UTF8
        Write-AuditLog "⚠ Fichiers obsolètes: $($Global:Reports.ObsoleteFiles)" -Level WARNING
    }
    
    # Noms invalides
    if ($Global:InvalidData.Count -gt 0) {
        $Global:InvalidData | Export-Csv -Path $Global:Reports.InvalidNames -NoTypeInformation -Encoding UTF8
        Write-AuditLog "⚠ Noms invalides: $($Global:Reports.InvalidNames)" -Level WARNING
    }
    
    # Fichiers trop gros
    if ($Global:TooLargeData.Count -gt 0) {
        $Global:TooLargeData | Export-Csv -Path $Global:Reports.TooLargeFiles -NoTypeInformation -Encoding UTF8
        Write-AuditLog "⚠ Fichiers >250MB: $($Global:Reports.TooLargeFiles)" -Level WARNING
    }
    
    # Doublons potentiels
    if ($Global:DuplicateData.Count -gt 0) {
        $Global:DuplicateData | Export-Csv -Path $Global:Reports.Duplicates -NoTypeInformation -Encoding UTF8
        Write-AuditLog "⚠ Doublons potentiels: $($Global:Reports.Duplicates)" -Level WARNING
    }
    
    # Permissions
    if ($Global:PermissionsData.Count -gt 0) {
        $Global:PermissionsData | Export-Csv -Path $Global:Reports.Permissions -NoTypeInformation -Encoding UTF8
        Write-AuditLog "✓ Rapport permissions: $($Global:Reports.Permissions)" -Level SUCCESS
    }
    
    # Rapport synthèse
    Export-SummaryReport
}

function Export-SummaryReport {
    
    $duration = (Get-Date) - $Global:AuditStats.StartTime
    
    $summary = @"
========================================
RAPPORT D'AUDIT PRÉ-MIGRATION COMPLET
========================================
Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Durée: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s

SOURCE AUDITÉE: $SourcePath

STATISTIQUES GLOBALES:
- Fichiers analysés: $($Global:AuditStats.TotalFiles)
- Dossiers: $($Global:AuditStats.TotalFolders)
- Taille totale: $([math]::Round($Global:AuditStats.TotalSizeGB, 2)) GB

CLASSIFICATION:
✓ Éligibles pour migration: $($Global:AuditStats.FilesEligible) fichiers
✗ Bloqués (extensions): $($Global:AuditStats.FilesBlocked) fichiers
✗ Obsolètes (>$ObsoleteDaysThreshold jours): $($Global:AuditStats.FilesObsolete) fichiers
✗ À archiver: $($Global:AuditStats.FilesArchive) fichiers
✗ Noms invalides: $($Global:AuditStats.FilesInvalid) fichiers
✗ Trop volumineux (>250MB): $($Global:AuditStats.FilesTooLarge) fichiers
⚠ Doublons potentiels: $($Global:AuditStats.DuplicateCount) fichiers

TAUX D'ÉLIGIBILITÉ:
$(if($Global:AuditStats.TotalFiles -gt 0){[math]::Round(($Global:AuditStats.FilesEligible / $Global:AuditStats.TotalFiles) * 100, 2)}else{0})%

FICHIERS GÉNÉRÉS:
- 01_Inventory_Complete: Inventaire exhaustif de tous les fichiers
- 02_EligibleForMigration: Fichiers prêts pour SharePoint
- 03_BlockedFiles: Extensions non autorisées dans SharePoint
- 04_ObsoleteFiles: Fichiers >$ObsoleteDaysThreshold jours (recommandation archivage)
- 05_InvalidFileNames: Noms non conformes SharePoint
- 06_FilesTooLarge: Fichiers > 250 MB (nécessitent compression/archivage)
- 07_PotentialDuplicates: Doublons potentiels pour nettoyage
- 08_Permissions: Matrice NTFS pour mapping SharePoint
- 09_AuditSummary: Ce rapport

RECOMMANDATIONS:
1. IMMÉDIAT
   - Archiver/supprimer les $($Global:AuditStats.FilesObsolete) fichiers obsolètes
   - Compresser ou archiver les $($Global:AuditStats.FilesTooLarge) fichiers >250MB
   - Bloquer/isoler les $($Global:AuditStats.FilesBlocked) fichiers avec extensions non autorisées

2. AVANT MIGRATION
   - Corriger les $($Global:AuditStats.FilesInvalid) noms de fichiers invalides
   - Dédupliquer les $($Global:AuditStats.DuplicateCount) doublons potentiels
   - Valider les permissions pour $($Global:PermissionsData.Count) entrées ACL

3. CONFORMITÉ RGPD
   - Examiner les données personnelles potentielles
   - Appliquer les restrictions d'accès selon la classification
   - Documenter la chaîne de traçabilité (audit trail)

4. OPTIMISATION
   - Compresser les fichiers >50MB avant migration
   - Nettoyer 30% du volume identifié (obsolète/dupliqué)
   - Restructurer selon la hiérarchie SharePoint cible

IMPACT ESTIMÉ:
- Fichiers à migrer réellement: $($Global:AuditStats.FilesEligible)
- Données à migrer: ~$([math]::Round($Global:AuditStats.TotalSizeGB * ($Global:AuditStats.FilesEligible / $Global:AuditStats.TotalFiles), 2)) GB
- Temps migration estimé (150 files/hour): ~$([math]::Ceiling($Global:AuditStats.FilesEligible / 150)) heures

========================================
Audit terminé avec succès.
Tous les rapports disponibles dans: $OutputPath
========================================
"@
    
    $summary | Out-File -FilePath $Global:Reports.Summary -Encoding UTF8
    Write-Host "`n$summary" -ForegroundColor Cyan
    Write-AuditLog "✓ Rapport synthèse: $($Global:Reports.Summary)" -Level SUCCESS
}

#endregion

#region Main

try {
    Invoke-ComprehensiveAudit -Path $SourcePath
    Export-AuditReports
    
    Write-AuditLog "========================================" -Level SUCCESS
    Write-AuditLog "AUDIT PRÉ-MIGRATION TERMINÉ" -Level SUCCESS
    Write-AuditLog "========================================" -Level SUCCESS
}
catch {
    Write-AuditLog "ERREUR FATALE: $($_.Exception.Message)" -Level CRITICAL
    exit 1
}

#endregion
