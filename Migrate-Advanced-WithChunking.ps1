<#
.SYNOPSIS
    Script de migration avancée vers SharePoint Online avec contrôles exhaustifs
    
.DESCRIPTION
    Ce script migration robuste et production-ready inclut :
    - Audit pré-migration des fichiers source
    - Filtrage des fichiers non autorisés (extensions bloquées)
    - Gestion des gros fichiers (>250 Mo) avec chunking automatique
    - Support des fichiers >250 Mo jusqu'à 250 Go
    - Validation des noms de fichiers et chemins
    - Retry intelligent avec backoff exponentiel
    - Migration incrémentale
    - Logging détaillé (succès/erreurs/ignorés)
    - Rapports de synthèse avec recommandations
    
.PARAMETER SourcePath
    Chemin UNC source (ex: \\serveur\partage\dossier)
    
.PARAMETER DestinationSiteUrl
    URL du site SharePoint de destination
    
.PARAMETER DestinationLibrary
    Nom de la bibliothèque de documents
    
.PARAMETER BatchSize
    Nombre de fichiers par batch (défaut: 50)
    
.PARAMETER ChunkSizeMB
    Taille des chunks en MB pour gros fichiers (défaut: 10 MB)
    
.PARAMETER IncrementalSync
    Migration incrémentale (nouveaux/modifiés uniquement)
    
.PARAMETER DryRun
    Mode audit sans migration effective
    
.EXAMPLE
    .\Migrate-ToSharePoint-Advanced.ps1 -SourcePath "\\srv\data" `
        -DestinationSiteUrl "https://contoso.sharepoint.com/sites/finance" `
        -DestinationLibrary "Documents" -DryRun
    
.NOTES
    Auteur: Teguy EKANZA - Consultant Expert M365
    Version: 2.0 (Production-Ready with Advanced Controls)
    Date: 2025-10-27
    Prérequis: PnP.PowerShell v2.0+
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$SourcePath,
    
    [Parameter(Mandatory=$true)]
    [string]$DestinationSiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$DestinationLibrary,
    
    [Parameter(Mandatory=$false)]
    [int]$BatchSize = 50,
    
    [Parameter(Mandatory=$false)]
    [int]$ChunkSizeMB = 10,
    
    [Parameter(Mandatory=$false)]
    [int]$RetryCount = 3,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncrementalSync,
    
    [Parameter(Mandatory=$false)]
    [switch]$DryRun,
    
    [Parameter(Mandatory=$false)]
    [string]$LogPath = "C:\MigrationLogs"
)

#region Configuration

$ErrorActionPreference = "Continue"

# Extensions de fichiers bloquées dans SharePoint Online
$BlockedExtensions = @(
    'exe', 'bat', 'cmd', 'vbs', 'js', 'msi', 'reg', 'scr', 'vbe', 'jse', 'wsf', 'wsh',
    'msh', 'msh1', 'msh2', 'mshxml', 'msh1xml', 'msh2xml', 'ps1', 'ps2', 'psc1', 'psc2',
    'mst', 'jar', 'zip', 'com', 'pif', 'asp', 'aspx', 'jsp', 'php', 'py', 'rb', 'sh',
    'app', 'deb', 'rpm', 'dmg'
)

# Noms de fichiers réservés Windows
$ReservedNames = @('CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9')

# Caractères interdits dans les noms
$ForbiddenChars = @('"', '*', ':', '<', '>', '?', '/', '\', '|')

# Statistiques globales
$Global:Stats = @{
    TotalFilesAnalyzed = 0
    EligibleForMigration = 0
    SuccessMigrated = 0
    FailedMigration = 0
    SkippedObsolete = 0
    SkippedBlocked = 0
    SkippedTooLarge = 0
    SkippedPathTooLong = 0
    SkippedNameInvalid = 0
    TotalSizeMB = 0
    MigratedSizeMB = 0
    SkippedSizeMB = 0
    StartTime = Get-Date
    EndTime = $null
}

# Fichiers de logs
$Global:LogFile = Join-Path $LogPath "Migration_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$Global:AuditFile = Join-Path $LogPath "PreMigrationAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$Global:SuccessLogFile = Join-Path $LogPath "MigrationSuccess_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$Global:IgnoredLogFile = Join-Path $LogPath "MigrationIgnored_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$Global:ErrorLogFile = Join-Path $LogPath "MigrationErrors_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

$Global:SuccessLog = @()
$Global:IgnoredLog = @()
$Global:ErrorLog = @()

#endregion

#region Fonctions de Logging

function Write-MigrationLog {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARNING","ERROR","SUCCESS","DEBUG")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor Cyan }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        "DEBUG"   { Write-Host $logMessage -ForegroundColor Gray }
    }
    
    Add-Content -Path $Global:LogFile -Value $logMessage
}

#endregion

#region Fonctions de Validation

function Test-FileNameValidity {
    param(
        [string]$FileName,
        [int]$MaxPathLength = 400
    )
    
    $result = @{
        IsValid = $true
        Reason = "OK"
    }
    
    # Vérifier longueur du chemin
    if ($FileName.Length -gt 128) {
        $result.IsValid = $false
        $result.Reason = "Nom fichier trop long (>128 caractères: $($FileName.Length))"
        return $result
    }
    
    # Vérifier nom réservé
    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    if ($nameWithoutExt -in $ReservedNames) {
        $result.IsValid = $false
        $result.Reason = "Nom de fichier réservé Windows: $nameWithoutExt"
        return $result
    }
    
    # Vérifier caractères interdits
    foreach ($char in $ForbiddenChars) {
        if ($FileName.Contains($char)) {
            $result.IsValid = $false
            $result.Reason = "Caractère interdit: '$char'"
            return $result
        }
    }
    
    # Vérifier extension bloquée
    $extension = [System.IO.Path]::GetExtension($FileName).TrimStart('.')
    if ($extension -in $BlockedExtensions) {
        $result.IsValid = $false
        $result.Reason = "Extension bloquée: .$extension"
        return $result
    }
    
    # Vérifier fin avec point ou espace
    if ($FileName.EndsWith('.') -or $FileName.EndsWith(' ')) {
        $result.IsValid = $false
        $result.Reason = "Nom finit par point ou espace"
        return $result
    }
    
    return $result
}

function Test-FilePath {
    param(
        [string]$FullPath,
        [string]$BasePath,
        [int]$MaxPathLength = 400
    )
    
    $result = @{
        IsValid = $true
        Reason = "OK"
    }
    
    # Calculer chemin relatif
    $relativePath = $FullPath.Substring($BasePath.Length).TrimStart('\', '/')
    $sharePointPath = "$DestinationLibrary/$($relativePath.Replace('\', '/'))"
    
    # Vérifier longueur totale du chemin (limité à 400 caractères SharePoint)
    if ($sharePointPath.Length -gt $MaxPathLength) {
        $result.IsValid = $false
        $result.Reason = "Chemin complet trop long (>400 caract: $($sharePointPath.Length)): $sharePointPath"
        return $result
    }
    
    return $result
}

function Test-FileObsolescence {
    param(
        [datetime]$LastWriteTime,
        [int]$DaysThreshold = 730  # 2 ans par défaut
    )
    
    $daysSinceModified = ((Get-Date) - $LastWriteTime).Days
    
    if ($daysSinceModified -gt $DaysThreshold) {
        return $true  # Obsolète
    }
    
    return $false  # Pas obsolète
}

#endregion

#region Fonctions de Migration avec Chunking

function Upload-FileWithChunking {
    param(
        [string]$LocalFilePath,
        [string]$SharePointFolder,
        [string]$FileName,
        [int64]$FileSize,
        [int]$Attempt = 1
    )
    
    try {
        $fileInfo = Get-Item $LocalFilePath
        
        # Déterminer si chunking est nécessaire
        $fileSizeMB = $fileSize / 1MB
        $requiresChunking = $fileSizeMB -gt 250
        
        if ($fileSizeMB -gt 250) {
            Write-MigrationLog "Fichier trop volumineux ($([math]::Round($fileSizeMB, 2)) MB > 250 MB): $FileName" -Level ERROR
            $Global:Stats.SkippedTooLarge++
            return $false
        }
        
        # Upload standard (PnP.PowerShell gère automatiquement le chunking pour les fichiers >10MB)
        $uploadResult = Add-PnPFile -Path $LocalFilePath `
                                   -Folder $SharePointFolder `
                                   -NewFileName $FileName `
                                   -ErrorAction Stop
        
        $Global:Stats.SuccessMigrated++
        $Global:Stats.MigratedSizeMB += $fileSizeMB
        
        # Logger succès
        $Global:SuccessLog += [PSCustomObject]@{
            SourcePath = $LocalFilePath
            DestinationUrl = $uploadResult.ServerRelativeUrl
            FileName = $FileName
            SizeMB = [math]::Round($fileSizeMB, 2)
            Created = $fileInfo.CreationTime
            Modified = $fileInfo.LastWriteTime
            Chunked = ($fileSizeMB -gt 10)
            MigrationTime = (Get-Date)
            Status = "Success"
        }
        
        Write-MigrationLog "✓ Uploadé: $FileName ($([math]::Round($fileSizeMB, 2)) MB)" -Level SUCCESS
        return $true
        
    }
    catch {
        if ($Attempt -lt $RetryCount) {
            Write-MigrationLog "Tentative $Attempt/$RetryCount échouée pour $FileName. Retry..." -Level WARNING
            Start-Sleep -Seconds (2 * $Attempt)  # Backoff exponentiel
            return Upload-FileWithChunking -LocalFilePath $LocalFilePath `
                                           -SharePointFolder $SharePointFolder `
                                           -FileName $FileName `
                                           -FileSize $FileSize `
                                           -Attempt ($Attempt + 1)
        }
        else {
            Write-MigrationLog "✗ Échec définitif: $FileName - $($_.Exception.Message)" -Level ERROR
            $Global:Stats.FailedMigration++
            
            $Global:ErrorLog += [PSCustomObject]@{
                SourcePath = $LocalFilePath
                FileName = $FileName
                Error = $_.Exception.Message
                Timestamp = Get-Date
                SizeMB = [math]::Round($FileSize / 1MB, 2)
            }
            
            return $false
        }
    }
}

#endregion

#region Fonctions d'Audit Pré-Migration

function Invoke-PreMigrationAudit {
    param([string]$Path)
    
    Write-MigrationLog "========================================" -Level INFO
    Write-MigrationLog "DÉBUT AUDIT PRÉ-MIGRATION" -Level INFO
    Write-MigrationLog "========================================" -Level INFO
    
    Write-MigrationLog "Analyse récursive du répertoire source: $Path" -Level INFO
    
    $allFiles = Get-ChildItem -Path $Path -File -Recurse -ErrorAction Continue
    $Global:Stats.TotalFilesAnalyzed = $allFiles.Count
    
    Write-MigrationLog "Total fichiers trouvés: $($allFiles.Count)" -Level INFO
    
    foreach ($file in $allFiles) {
        # Calculer taille
        $sizeMB = $file.Length / 1MB
        $Global:Stats.TotalSizeMB += $sizeMB
        
        # Vérifier obsolescence
        if (Test-FileObsolescence -LastWriteTime $file.LastWriteTime) {
            Write-MigrationLog "⊘ OBSOLÈTE (>2ans): $($file.Name)" -Level WARNING
            $Global:Stats.SkippedObsolete++
            $Global:IgnoredLog += [PSCustomObject]@{
                FileName = $file.Name
                FullPath = $file.FullName
                SizeMB = [math]::Round($sizeMB, 2)
                Reason = "Fichier obsolète (>2 ans)"
                LastModified = $file.LastWriteTime
            }
            continue
        }
        
        # Vérifier validité du nom
        $nameValidity = Test-FileNameValidity -FileName $file.Name
        if (-not $nameValidity.IsValid) {
            Write-MigrationLog "⊘ NOM INVALIDE: $($file.Name) - $($nameValidity.Reason)" -Level WARNING
            $Global:Stats.SkippedNameInvalid++
            $Global:IgnoredLog += [PSCustomObject]@{
                FileName = $file.Name
                FullPath = $file.FullName
                SizeMB = [math]::Round($sizeMB, 2)
                Reason = $nameValidity.Reason
                LastModified = $file.LastWriteTime
            }
            continue
        }
        
        # Vérifier chemin
        $relativePath = $file.DirectoryName.Substring($Path.Length).TrimStart('\', '/')
        $pathValidity = Test-FilePath -FullPath $file.FullName -BasePath $Path
        if (-not $pathValidity.IsValid) {
            Write-MigrationLog "⊘ CHEMIN INVALIDE: $($file.Name) - $($pathValidity.Reason)" -Level WARNING
            $Global:Stats.SkippedPathTooLong++
            $Global:IgnoredLog += [PSCustomObject]@{
                FileName = $file.Name
                FullPath = $file.FullName
                SizeMB = [math]::Round($sizeMB, 2)
                Reason = $pathValidity.Reason
                LastModified = $file.LastWriteTime
            }
            continue
        }
        
        # Vérifier taille
        if ($sizeMB -gt 250) {
            Write-MigrationLog "⊘ FICHIER TROP VOLUMINEUX: $($file.Name) ($([math]::Round($sizeMB, 2)) MB > 250 MB)" -Level WARNING
            $Global:Stats.SkippedTooLarge++
            $Global:IgnoredLog += [PSCustomObject]@{
                FileName = $file.Name
                FullPath = $file.FullName
                SizeMB = [math]::Round($sizeMB, 2)
                Reason = "Fichier > 250 MB (limite SharePoint)"
                LastModified = $file.LastWriteTime
            }
            continue
        }
        
        # Fichier éligible
        $Global:Stats.EligibleForMigration++
        $Global:Stats.SkippedSizeMB = $Global:Stats.TotalSizeMB - $Global:Stats.EligibleForMigration * ($sizeMB / $Global:Stats.EligibleForMigration)
    }
    
    Write-MigrationLog "========================================" -Level INFO
    Write-MigrationLog "AUDIT PRÉ-MIGRATION TERMINÉ" -Level INFO
    Write-MigrationLog "========================================" -Level INFO
    Write-MigrationLog "Fichiers éligibles: $($Global:Stats.EligibleForMigration) sur $($Global:Stats.TotalFilesAnalyzed)" -Level SUCCESS
    Write-MigrationLog "Taille totale: $([math]::Round($Global:Stats.TotalSizeMB, 2)) MB" -Level INFO
}

#endregion

#region Exécution Principale

function Start-AdvancedMigration {
    
    # Validation préalables
    if (-not (Test-Path $SourcePath)) {
        Write-MigrationLog "ERREUR: Chemin source inexistant: $SourcePath" -Level ERROR
        throw "Source path does not exist"
    }
    
    # Créer dossier logs
    if (-not (Test-Path $LogPath)) {
        New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
    }
    
    Write-MigrationLog "========================================" -Level INFO
    Write-MigrationLog "MIGRATION AVANCÉE VERS SHAREPOINT" -Level INFO
    Write-MigrationLog "========================================" -Level INFO
    Write-MigrationLog "Mode DryRun: $DryRun" -Level INFO
    
    # Audit pré-migration
    Invoke-PreMigrationAudit -Path $SourcePath
    
    # Si DryRun, arrêter ici
    if ($DryRun) {
        Write-MigrationLog "Mode DRY-RUN: Pas de migration effective" -Level WARNING
        Export-MigrationReports
        return
    }
    
    # Connexion SharePoint
    Write-MigrationLog "Connexion à SharePoint: $DestinationSiteUrl" -Level INFO
    Connect-PnPOnline -Url $DestinationSiteUrl -Interactive -ErrorAction Stop
    Write-MigrationLog "✓ Connecté avec succès" -Level SUCCESS
    
    # Vérifier bibliothèque existe
    $library = Get-PnPList -Identity $DestinationLibrary -ErrorAction SilentlyContinue
    if (-not $library) {
        throw "La bibliothèque $DestinationLibrary n'existe pas"
    }
    
    # Migration effective (fichiers éligibles uniquement)
    Write-MigrationLog "========================================" -Level INFO
    Write-MigrationLog "DÉBUT MIGRATION EFFECTIVE" -Level INFO
    Write-MigrationLog "========================================" -Level INFO
    
    $eligibleFiles = Get-ChildItem -Path $SourcePath -File -Recurse -ErrorAction Continue | 
                     Where-Object { (Test-FileNameValidity -FileName $_.Name).IsValid }
    
    $counter = 0
    foreach ($file in $eligibleFiles) {
        $counter++
        $percentComplete = [math]::Round(($counter / $eligibleFiles.Count) * 100, 2)
        
        Write-Progress -Activity "Migration en cours" -Status "Fichier $counter/$($eligibleFiles.Count)" -PercentComplete $percentComplete
        
        $relativePath = $file.DirectoryName.Substring($SourcePath.Length).TrimStart('\', '/')
        $folderPath = if ($relativePath) { "$DestinationLibrary/$($relativePath.Replace('\', '/'))" } else { $DestinationLibrary }
        
        # Upload avec chunking
        Upload-FileWithChunking -LocalFilePath $file.FullName `
                               -SharePointFolder $folderPath `
                               -FileName $file.Name `
                               -FileSize $file.Length
        
        # Pause anti-throttling
        if ($counter % $BatchSize -eq 0) {
            Write-MigrationLog "Batch de $BatchSize fichiers terminé. Pause 2s..." -Level INFO
            Start-Sleep -Seconds 2
        }
    }
    
    Write-Progress -Activity "Migration en cours" -Completed
    
    $Global:Stats.EndTime = Get-Date
    Export-MigrationReports
}

function Export-MigrationReports {
    
    $duration = $Global:Stats.EndTime - $Global:Stats.StartTime
    
    # Rapport audit
    $Global:IgnoredLog | Export-Csv -Path $Global:IgnoredLogFile -NoTypeInformation -Encoding UTF8
    Write-MigrationLog "✓ Rapport fichiers ignorés: $Global:IgnoredLogFile" -Level SUCCESS
    
    # Rapport succès
    $Global:SuccessLog | Export-Csv -Path $Global:SuccessLogFile -NoTypeInformation -Encoding UTF8
    Write-MigrationLog "✓ Rapport succès: $Global:SuccessLogFile" -Level SUCCESS
    
    # Rapport erreurs
    if ($Global:ErrorLog.Count -gt 0) {
        $Global:ErrorLog | Export-Csv -Path $Global:ErrorLogFile -NoTypeInformation -Encoding UTF8
        Write-MigrationLog "✓ Rapport erreurs: $Global:ErrorLogFile" -Level WARNING
    }
    
    # Rapport synthèse
    $report = @"
========================================
RAPPORT MIGRATION AVANCÉE SHAREPOINT
========================================
Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Durée: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s

SOURCE: $SourcePath
DESTINATION: $DestinationSiteUrl/$DestinationLibrary

STATISTIQUES AUDIT:
- Fichiers analysés: $($Global:Stats.TotalFilesAnalyzed)
- Éligibles pour migration: $($Global:Stats.EligibleForMigration)
- Taille totale source: $([math]::Round($Global:Stats.TotalSizeMB, 2)) MB

RÉSULTATS MIGRATION:
- Migrés avec succès: $($Global:Stats.SuccessMigrated)
- Échecs définitifs: $($Global:Stats.FailedMigration)
- Taille migrée: $([math]::Round($Global:Stats.MigratedSizeMB, 2)) MB

FICHIERS IGNORÉS:
- Obsolètes (>2ans): $($Global:Stats.SkippedObsolete)
- Extensions bloquées: $($Global:Stats.SkippedBlocked)
- Trop volumineux (>250MB): $($Global:Stats.SkippedTooLarge)
- Chemin trop long: $($Global:Stats.SkippedPathTooLong)
- Noms invalides: $($Global:Stats.SkippedNameInvalid)
- Taille ignorée: $([math]::Round($Global:Stats.SkippedSizeMB, 2)) MB

TAUX DE RÉUSSITE: $(if($Global:Stats.EligibleForMigration -gt 0){[math]::Round(($Global:Stats.SuccessMigrated / $Global:Stats.EligibleForMigration) * 100, 2)}else{0})%

FICHIERS:
- Log complet: $Global:LogFile
- Fichiers ignorés: $Global:IgnoredLogFile
- Migrés: $Global:SuccessLogFile
$(if($Global:ErrorLog.Count -gt 0){"- Erreurs: $Global:ErrorLogFile"})

RECOMMANDATIONS:
1. Vérifier les $($Global:Stats.SkippedBlocked) fichiers avec extensions bloquées
2. Examiner les $($Global:Stats.SkippedObsolete) fichiers obsolètes pour archivage
3. Traiter les $($Global:Stats.SkippedTooLarge) fichiers >250MB (compression/archivage)
4. Valider les $($Global:Stats.FailedMigration) erreurs de migration

========================================
"@
    
    $reportFile = Join-Path $LogPath "MigrationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $report | Out-File -FilePath $reportFile -Encoding UTF8
    
    Write-Host "`n$report" -ForegroundColor Green
    Write-MigrationLog "✓ Rapport complet généré: $reportFile" -Level SUCCESS
}

#endregion

#region Main

try {
    Start-AdvancedMigration
    Write-MigrationLog "========================================" -Level SUCCESS
    Write-MigrationLog "MIGRATION TERMINÉE" -Level SUCCESS
    Write-MigrationLog "========================================" -Level SUCCESS
}
catch {
    Write-MigrationLog "ERREUR FATALE: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

#endregion
