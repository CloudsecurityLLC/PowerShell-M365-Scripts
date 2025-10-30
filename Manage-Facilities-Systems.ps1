<#
.SYNOPSIS
    Script de gestion et monitoring Affichage Dynamique & Téléphonie VoIP
    
.DESCRIPTION
    Ce script supervise les systèmes d'affichage dynamique et la téléphonie VoIP/Teams Phone,
    analyse la qualité de service, génère des rapports et identifie les problèmes.
    Adapté pour Product Manager Facilities & Communication chez Veolia.
    
.PARAMETER System
    Système à monitorer: DigitalSignage, VoIP, Both
    
.PARAMETER Action
    Action à effectuer: Monitor, QualityCheck, Analytics, Alert, FullAudit
    
.PARAMETER ExportPath
    Chemin d'export des rapports (par défaut : C:\FacilitiesReports)
    
.PARAMETER DaysAnalysis
    Nombre de jours d'historique à analyser (par défaut : 7)
    
.EXAMPLE
    .\Manage-Facilities-Systems.ps1 -System DigitalSignage -Action Monitor
    
.EXAMPLE
    .\Manage-Facilities-Systems.ps1 -System VoIP -Action QualityCheck
    
.EXAMPLE
    .\Manage-Facilities-Systems.ps1 -System Both -Action FullAudit -ExportPath "D:\Reports"
    
.NOTES
    Auteur: Consultant Facilities & Telecom
    Version: 2.0
    Date: 2025-10-26
    
    Prerequisites:
    - PowerShell 5.0 ou supérieur
    - Pour VoIP: Module Microsoft.Graph (Teams Phone)
    - Pour Digital Signage: Accès API fournisseur
    
    Note: Script utilise données simulées pour démo
    Pour production, connecter aux APIs réelles
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("DigitalSignage", "VoIP", "Both")]
    [string]$System,
    
    [Parameter(Mandatory=$true)]
    [ValidateSet("Monitor", "QualityCheck", "Analytics", "Alert", "FullAudit")]
    [string]$Action,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = "C:\FacilitiesReports",
    
    [Parameter(Mandatory=$false)]
    [int]$DaysAnalysis = 7
)

#========================================
# CONFIGURATION GLOBALE
#========================================

$ErrorActionPreference = "Continue"
$VerbosePreference = "Continue"

$Global:Config = @{
    # Digital Signage
    SignageAPIURL = "https://signage.veolia.local/api"
    SignageRefreshInterval = 300
    SignageSLAUptime = 99.0
    
    # VoIP / Teams Phone
    VoIPAPIURL = "https://teams.microsoft.com/api"
    CallQualityThresholdMOS = 3.5
    CallQualityTargetMOS = 4.0
    MaxAcceptableLatency = 150
    MaxAcceptableJitter = 30
    MaxAcceptablePacketLoss = 1.0
}

$Colors = @{
    Success = "Green"
    Warning = "Yellow"
    Error = "Red"
    Info = "Cyan"
    Highlight = "Magenta"
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Info", "Success", "Warning", "Error")]
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage -ForegroundColor $Colors[$Level]
}

#========================================
# SECTION 1: AFFICHAGE DYNAMIQUE
#========================================

function Generate-DigitalSignageData {
    try {
        Write-Log "Génération données affichage dynamique..." -Level "Info"
        
        $screens = @()
        $locations = @("Paris-HQ Hall", "Aubervilliers Reception", "Lyon Cafétéria", "Bordeaux Entrée", "Marseille Bureau")
        $contentTypes = @("Actualités", "Météo", "Menu Cafét", "Événements", "KPIs Business", "Sécurité")
        
        for ($i = 1; $i -le 30; $i++) {
            $uptime = Get-Random -Minimum 95 -Maximum 100
            $status = if ($uptime -gt 98) { "Online" } 
                     elseif ($uptime -gt 95) { "Degraded" }
                     else { "Offline" }
            
            $screen = [PSCustomObject]@{
                ScreenID = "SCREEN-$('{0:D3}' -f $i)"
                ScreenName = "Écran $('{0:D2}' -f $i)"
                Location = $locations | Get-Random
                Size = @("32''", "42''", "55''", "65''") | Get-Random
                Resolution = "1920x1080"
                Status = $status
                UptimePercent = [math]::Round($uptime, 2)
                LastUpdate = (Get-Date).AddMinutes(-(Get-Random -Minimum 1 -Maximum 60))
                CurrentContent = $contentTypes | Get-Random
                ContentRotations = Get-Random -Minimum 50 -Maximum 500
                Brightness = Get-Random -Minimum 50 -Maximum 100
                Temperature = Get-Random -Minimum 20 -Maximum 45
                NetworkLatency = Get-Random -Minimum 5 -Maximum 150
            }
            $screens += $screen
        }
        
        Write-Log "✓ $($screens.Count) écrans générés" -Level "Success"
        return $screens
    }
    catch {
        Write-Log "✗ Erreur génération données signage: $_" -Level "Error"
        return $null
    }
}

function Generate-SignageContent {
    param([int]$Days = 7)
    
    try {
        Write-Log "Génération historique contenus affichés..." -Level "Info"
        
        $contents = @()
        $contentTypes = @("Actualités", "Météo", "Menu Cafét", "Événements", "KPIs Business", "Sécurité")
        
        for ($d = 0; $d -lt $Days; $d++) {
            $date = (Get-Date).AddDays(-$d)
            
            foreach ($type in $contentTypes) {
                $content = [PSCustomObject]@{
                    ContentID = "CONTENT-$(Get-Date $date -Format 'yyyyMMdd')-$(Get-Random -Minimum 100 -Maximum 999)"
                    Type = $type
                    Date = $date
                    DisplayCount = Get-Random -Minimum 100 -Maximum 2000
                    AvgDisplayTime = Get-Random -Minimum 5 -Maximum 60
                    InteractionRate = [math]::Round((Get-Random -Minimum 0 -Maximum 30) / 100, 3)
                    UpdateFrequency = switch ($type) {
                        "Actualités" { "Horaire" }
                        "Météo" { "30min" }
                        "Menu Cafét" { "Quotidien" }
                        "Événements" { "Hebdo" }
                        "KPIs Business" { "Temps réel" }
                        "Sécurité" { "Mensuel" }
                    }
                }
                $contents += $content
            }
        }
        
        Write-Log "✓ $($contents.Count) contenus générés" -Level "Success"
        return $contents
    }
    catch {
        Write-Log "✗ Erreur génération contenus: $_" -Level "Error"
        return $null
    }
}

function Monitor-DigitalSignage {
    param([array]$Screens)
    
    try {
        Write-Log "═══════════════════════════════════════════════════════" -Level "Info"
        Write-Log "MONITORING AFFICHAGE DYNAMIQUE" -Level "Highlight"
        Write-Log "═══════════════════════════════════════════════════════" -Level "Info"
        Write-Host ""
        
        $stats = @{
            TotalScreens = $Screens.Count
            OnlineScreens = ($Screens | Where-Object { $_.Status -eq "Online" }).Count
            DegradedScreens = ($Screens | Where-Object { $_.Status -eq "Degraded" }).Count
            OfflineScreens = ($Screens | Where-Object { $_.Status -eq "Offline" }).Count
            AvgUptime = [math]::Round(($Screens | Measure-Object -Property UptimePercent -Average).Average, 2)
            AvgNetworkLatency = [math]::Round(($Screens | Measure-Object -Property NetworkLatency -Average).Average, 2)
        }
        
        Write-Log "ÉTAT DU PARC:" -Level "Info"
        Write-Host "  • Total écrans: $($stats.TotalScreens)"
        Write-Host "  • En ligne: $($stats.OnlineScreens)" -ForegroundColor Green
        Write-Host "  • Dégradés: $($stats.DegradedScreens)" -ForegroundColor Yellow
        Write-Host "  • Hors ligne: $($stats.OfflineScreens)" -ForegroundColor Red
        Write-Host "  • Uptime moyen: $($stats.AvgUptime)%"
        Write-Host "  • Latence réseau moyenne: $($stats.AvgNetworkLatency)ms"
        Write-Host ""
        
        # Alertes
        $alerts = $Screens | Where-Object { 
            $_.Status -in @("Offline", "Degraded") -or 
            $_.Temperature -gt 40 -or 
            $_.NetworkLatency -gt 100 
        }
        
        if ($alerts.Count -gt 0) {
            Write-Log "⚠ ALERTES DÉTECTÉES:" -Level "Warning"
            $alerts | Select-Object -First 5 | ForEach-Object {
                Write-Host "  • $($_.ScreenName) [$($_.Location)]" -ForegroundColor Yellow
                Write-Host "    Status: $($_.Status) | Temp: $($_.Temperature)°C | Latence: $($_.NetworkLatency)ms"
            }
            Write-Host ""
        }
        
        return @{
            Statistics = $stats
            Alerts = $alerts
        }
    }
    catch {
        Write-Log "✗ Erreur monitoring signage: $_" -Level "Error"
        return $null
    }
}

function Analyze-SignageUsage {
    param([array]$Contents)
    
    try {
        Write-Log "Analyse utilisation contenus affichage..." -Level "Info"
        Write-Host ""
        
        $analysis = $Contents | Group-Object Type | ForEach-Object {
            [PSCustomObject]@{
                ContentType = $_.Name
                TotalDisplays = ($_.Group | Measure-Object -Property DisplayCount -Sum).Sum
                AvgDisplayTime = [math]::Round(($_.Group | Measure-Object -Property AvgDisplayTime -Average).Average, 2)
                AvgInteractionRate = [math]::Round(($_.Group | Measure-Object -Property InteractionRate -Average).Average * 100, 2)
                UpdateFrequency = $_.Group[0].UpdateFrequency
            }
        } | Sort-Object TotalDisplays -Descending
        
        Write-Log "TOP CONTENUS PAR AFFICHAGE:" -Level "Info"
        $analysis | Select-Object -First 5 | ForEach-Object {
            Write-Host "  • $($_.ContentType): $($_.TotalDisplays) affichages | $($_.AvgDisplayTime)s moyen"
        }
        Write-Host ""
        
        return $analysis
    }
    catch {
        Write-Log "✗ Erreur analyse signage: $_" -Level "Error"
        return $null
    }
}

#========================================
# SECTION 2: TÉLÉPHONIE VoIP
#========================================

function Generate-VoIPData {
    param([int]$Days = 7)
    
    try {
        Write-Log "Génération données téléphonie VoIP..." -Level "Info"
        
        $calls = @()
        $locations = @("Paris-HQ", "Aubervilliers", "Lyon", "Bordeaux", "Marseille")
        
        for ($d = 0; $d -lt $Days; $d++) {
            $date = (Get-Date).AddDays(-$d)
            $callsPerDay = Get-Random -Minimum 200 -Maximum 500
            
            for ($c = 0; $c -lt $callsPerDay; $c++) {
                # MOS (Mean Opinion Score): 1-5, 4+ = bon
                $mos = [math]::Round((Get-Random -Minimum 250 -Maximum 500) / 100, 2)
                
                $call = [PSCustomObject]@{
                    CallID = "CALL-$(Get-Date $date -Format 'yyyyMMdd')-$(Get-Random -Minimum 10000 -Maximum 99999)"
                    Date = $date
                    Caller = "user$(Get-Random -Minimum 1 -Maximum 500)@veolia.com"
                    Callee = "user$(Get-Random -Minimum 1 -Maximum 500)@veolia.com"
                    Location = $locations | Get-Random
                    Duration = Get-Random -Minimum 30 -Maximum 3600
                    CallType = @("Internal", "External", "Conference") | Get-Random
                    Direction = @("Inbound", "Outbound") | Get-Random
                    MOS = $mos
                    Latency = Get-Random -Minimum 10 -Maximum 300
                    Jitter = Get-Random -Minimum 1 -Maximum 50
                    PacketLoss = [math]::Round((Get-Random -Minimum 0 -Maximum 50) / 10, 2)
                    Quality = if ($mos -ge 4.0) { "Excellent" }
                             elseif ($mos -ge 3.5) { "Good" }
                             elseif ($mos -ge 3.0) { "Fair" }
                             else { "Poor" }
                    CallDropped = (Get-Random -Minimum 0 -Maximum 100) -lt 2
                }
                $calls += $call
            }
        }
        
        Write-Log "✓ $($calls.Count) appels générés" -Level "Success"
        return $calls
    }
    catch {
        Write-Log "✗ Erreur génération données VoIP: $_" -Level "Error"
        return $null
    }
}

function Generate-VoIPUsers {
    try {
        Write-Log "Génération données utilisateurs VoIP..." -Level "Info"
        
        $users = @()
        $locations = @("Paris-HQ", "Aubervilliers", "Lyon", "Bordeaux", "Marseille")
        
        for ($i = 1; $i -le 100; $i++) {
            $user = [PSCustomObject]@{
                UserID = "user$i@veolia.com"
                DisplayName = "Utilisateur $i"
                Location = $locations | Get-Random
                PhoneNumber = "+33$(Get-Random -Minimum 100000000 -Maximum 999999999)"
                TeamsPhoneEnabled = $true
                VoicemailEnabled = (Get-Random -Minimum 0 -Maximum 100) -gt 20
                CallsThisWeek = Get-Random -Minimum 5 -Maximum 100
                AvgCallDuration = Get-Random -Minimum 120 -Maximum 1800
                MissedCallsRate = [math]::Round((Get-Random -Minimum 0 -Maximum 30) / 100, 3)
            }
            $users += $user
        }
        
        Write-Log "✓ $($users.Count) utilisateurs générés" -Level "Success"
        return $users
    }
    catch {
        Write-Log "✗ Erreur génération utilisateurs: $_" -Level "Error"
        return $null
    }
}

function Monitor-VoIPQuality {
    param([array]$Calls)
    
    try {
        Write-Log "═══════════════════════════════════════════════════════" -Level "Info"
        Write-Log "MONITORING TÉLÉPHONIE VoIP / TEAMS PHONE" -Level "Highlight"
        Write-Log "═══════════════════════════════════════════════════════" -Level "Info"
        Write-Host ""
        
        $stats = @{
            TotalCalls = $Calls.Count
            AvgMOS = [math]::Round(($Calls | Measure-Object -Property MOS -Average).Average, 2)
            ExcellentQuality = ($Calls | Where-Object { $_.Quality -eq "Excellent" }).Count
            GoodQuality = ($Calls | Where-Object { $_.Quality -eq "Good" }).Count
            FairQuality = ($Calls | Where-Object { $_.Quality -eq "Fair" }).Count
            PoorQuality = ($Calls | Where-Object { $_.Quality -eq "Poor" }).Count
            CallDropRate = [math]::Round((($Calls | Where-Object { $_.CallDropped }).Count / $Calls.Count) * 100, 2)
            AvgLatency = [math]::Round(($Calls | Measure-Object -Property Latency -Average).Average, 2)
            AvgJitter = [math]::Round(($Calls | Measure-Object -Property Jitter -Average).Average, 2)
            AvgPacketLoss = [math]::Round(($Calls | Measure-Object -Property PacketLoss -Average).Average, 2)
        }
        
        Write-Log "QUALITÉ D'APPEL:" -Level "Info"
        Write-Host "  • Total appels analysés: $($stats.TotalCalls)"
        Write-Host "  • MOS moyen: $($stats.AvgMOS)/5.0" -ForegroundColor $(if ($stats.AvgMOS -ge 4.0) { "Green" } elseif ($stats.AvgMOS -ge 3.5) { "Yellow" } else { "Red" })
        Write-Host "  • Qualité Excellente: $($stats.ExcellentQuality) ($([math]::Round(($stats.ExcellentQuality / $stats.TotalCalls) * 100, 2))%)"
        Write-Host "  • Qualité Bonne: $($stats.GoodQuality) ($([math]::Round(($stats.GoodQuality / $stats.TotalCalls) * 100, 2))%)"
        Write-Host "  • Qualité Passable: $($stats.FairQuality) ($([math]::Round(($stats.FairQuality / $stats.TotalCalls) * 100, 2))%)"
        Write-Host "  • Qualité Médiocre: $($stats.PoorQuality) ($([math]::Round(($stats.PoorQuality / $stats.TotalCalls) * 100, 2))%)" -ForegroundColor $(if ($stats.PoorQuality -gt 0) { "Red" } else { "Green" })
        Write-Host ""
        
        Write-Log "MÉTRIQUES RÉSEAU:" -Level "Info"
        Write-Host "  • Taux appels coupés: $($stats.CallDropRate)%" -ForegroundColor $(if ($stats.CallDropRate -gt 1) { "Red" } else { "Green" })
        Write-Host "  • Latence moyenne: $($stats.AvgLatency)ms" -ForegroundColor $(if ($stats.AvgLatency -gt $Global:Config.MaxAcceptableLatency) { "Red" } else { "Green" })
        Write-Host "  • Jitter moyen: $($stats.AvgJitter)ms" -ForegroundColor $(if ($stats.AvgJitter -gt $Global:Config.MaxAcceptableJitter) { "Yellow" } else { "Green" })
        Write-Host "  • Perte paquets moyenne: $($stats.AvgPacketLoss)%" -ForegroundColor $(if ($stats.AvgPacketLoss -gt $Global:Config.MaxAcceptablePacketLoss) { "Red" } else { "Green" })
        Write-Host ""
        
        # Alertes qualité
        $poorCalls = $Calls | Where-Object { 
            $_.MOS -lt $Global:Config.CallQualityThresholdMOS -or 
            $_.Latency -gt $Global:Config.MaxAcceptableLatency -or
            $_.PacketLoss -gt $Global:Config.MaxAcceptablePacketLoss
        }
        
        if ($poorCalls.Count -gt 0) {
            Write-Log "⚠ $($poorCalls.Count) APPELS AVEC PROBLÈMES DE QUALITÉ:" -Level "Warning"
            $poorCalls | Select-Object -First 5 | ForEach-Object {
                Write-Host "  • $($_.CallID) [$($_.Location)]" -ForegroundColor Yellow
                Write-Host "    MOS: $($_.MOS) | Latence: $($_.Latency)ms | PacketLoss: $($_.PacketLoss)%"
            }
            Write-Host ""
        }
        
        return @{
            Statistics = $stats
            PoorQualityCalls = $poorCalls
        }
    }
    catch {
        Write-Log "✗ Erreur monitoring VoIP: $_" -Level "Error"
        return $null
    }
}

function Analyze-VoIPUsage {
    param(
        [array]$Calls,
        [array]$Users
    )
    
    try {
        Write-Log "Analyse utilisation VoIP..." -Level "Info"
        Write-Host ""
        
        # Analyse par localisation
        $locationAnalysis = $Calls | Group-Object Location | ForEach-Object {
            [PSCustomObject]@{
                Location = $_.Name
                CallCount = $_.Count
                AvgMOS = [math]::Round(($_.Group | Measure-Object -Property MOS -Average).Average, 2)
                AvgDuration = [math]::Round(($_.Group | Measure-Object -Property Duration -Average).Average / 60, 2)
                QualityIssues = ($_.Group | Where-Object { $_.Quality -in @("Fair", "Poor") }).Count
            }
        } | Sort-Object CallCount -Descending
        
        Write-Log "TOP 3 SITES PAR VOLUME D'APPELS:" -Level "Info"
        $locationAnalysis | Select-Object -First 3 | ForEach-Object {
            Write-Host "  • $($_.Location): $($_.CallCount) appels | MOS: $($_.AvgMOS)"
        }
        Write-Host ""
        
        # Analyse par type d'appel
        $typeAnalysis = $Calls | Group-Object CallType | ForEach-Object {
            [PSCustomObject]@{
                CallType = $_.Name
                Count = $_.Count
                Percentage = [math]::Round(($_.Count / $Calls.Count) * 100, 2)
                AvgDuration = [math]::Round(($_.Group | Measure-Object -Property Duration -Average).Average / 60, 2)
            }
        }
        
        # Top utilisateurs
        $topUsers = $Users | Sort-Object CallsThisWeek -Descending | Select-Object -First 10
        
        return @{
            LocationAnalysis = $locationAnalysis
            TypeAnalysis = $typeAnalysis
            TopUsers = $topUsers
        }
    }
    catch {
        Write-Log "✗ Erreur analyse VoIP: $_" -Level "Error"
        return $null
    }
}

#========================================
# FONCTION: EXPORT RAPPORTS
#========================================

function Export-FacilitiesReports {
    param(
        [string]$Path,
        [string]$SystemType,
        [object]$SignageData,
        [object]$VoIPData
    )
    
    try {
        Write-Log "Export des rapports..." -Level "Info"
        Write-Host ""
        
        if (!(Test-Path $Path)) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        
        # Exports Digital Signage
        if ($SystemType -in @("DigitalSignage", "Both") -and $SignageData) {
            $signagePath = Join-Path $Path "DigitalSignage_Screens_$timestamp.csv"
            $SignageData.Screens | Export-Csv -Path $signagePath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            Write-Log "  ✓ Écrans: $signagePath" -Level "Success"
            
            if ($SignageData.Contents) {
                $contentPath = Join-Path $Path "DigitalSignage_Contents_$timestamp.csv"
                $SignageData.ContentAnalysis | Export-Csv -Path $contentPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                Write-Log "  ✓ Contenus: $contentPath" -Level "Success"
            }
        }
        
        # Exports VoIP
        if ($SystemType -in @("VoIP", "Both") -and $VoIPData) {
            $voipPath = Join-Path $Path "VoIP_Calls_$timestamp.csv"
            $VoIPData.Calls | Export-Csv -Path $voipPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            Write-Log "  ✓ Appels: $voipPath" -Level "Success"
            
            if ($VoIPData.Users) {
                $usersPath = Join-Path $Path "VoIP_Users_$timestamp.csv"
                $VoIPData.Users | Export-Csv -Path $usersPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                Write-Log "  ✓ Utilisateurs: $usersPath" -Level "Success"
            }
        }
        
        # Rapport synthétique
        $summaryPath = Join-Path $Path "Facilities_Summary_$timestamp.txt"
        $summary = @"
═══════════════════════════════════════════════════════════════════════════════
RAPPORT FACILITIES - VEOLIA
Système(s) analysé(s): $SystemType
═══════════════════════════════════════════════════════════════════════════════
Généré le: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")

$(if ($SystemType -in @("DigitalSignage", "Both") -and $SignageData) {@"
═══════════════════════════════════════════════════════════════════════════════
AFFICHAGE DYNAMIQUE
═══════════════════════════════════════════════════════════════════════════════
Total écrans: $($SignageData.MonitorData.Statistics.TotalScreens)
En ligne: $($SignageData.MonitorData.Statistics.OnlineScreens)
Uptime moyen: $($SignageData.MonitorData.Statistics.AvgUptime)%
Alertes: $($SignageData.MonitorData.Alerts.Count)
"@})

$(if ($SystemType -in @("VoIP", "Both") -and $VoIPData) {@"
═══════════════════════════════════════════════════════════════════════════════
TÉLÉPHONIE VoIP / TEAMS PHONE
═══════════════════════════════════════════════════════════════════════════════
Total appels: $($VoIPData.QualityData.Statistics.TotalCalls)
MOS moyen: $($VoIPData.QualityData.Statistics.AvgMOS)/5.0
Qualité excellente: $([math]::Round(($VoIPData.QualityData.Statistics.ExcellentQuality / $VoIPData.QualityData.Statistics.TotalCalls) * 100, 2))%
Taux appels coupés: $($VoIPData.QualityData.Statistics.CallDropRate)%
Problèmes qualité: $($VoIPData.QualityData.PoorQualityCalls.Count)
"@})

═══════════════════════════════════════════════════════════════════════════════
"@
        
        $summary | Out-File -FilePath $summaryPath -Encoding UTF8
        Write-Log "  ✓ Rapport synthétique: $summaryPath" -Level "Success"
        
        Write-Log "✓ Tous les rapports exportés!" -Level "Success"
        return $true
    }
    catch {
        Write-Log "✗ Erreur export: $_" -Level "Error"
        return $false
    }
}

#========================================
# PROGRAMME PRINCIPAL
#========================================

function Main {
    Clear-Host
    
    Write-Host @"
    
╔═══════════════════════════════════════════════════════════════════════════════╗
║                                                                               ║
║          GESTION FACILITIES - AFFICHAGE DYNAMIQUE & TÉLÉPHONIE VoIP           ║
║                    Monitoring, Quality Check & Analytics                      ║
║                                                                               ║
╚═══════════════════════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Cyan
    
    Write-Log "Configuration" -Level "Info"
    Write-Log "  • Système: $System" -Level "Info"
    Write-Log "  • Action: $Action" -Level "Info"
    Write-Log "  • Export Path: $ExportPath" -Level "Info"
    Write-Log "  • Période analyse: $DaysAnalysis jours" -Level "Info"
    Write-Host ""
    
    $signageData = $null
    $voipData = $null
    
    # Traitement Digital Signage
    if ($System -in @("DigitalSignage", "Both")) {
        Write-Host ""
        $screens = Generate-DigitalSignageData
        $contents = Generate-SignageContent -Days $DaysAnalysis
        
        Write-Host ""
        
        switch ($Action) {
            { $_ -in @("Monitor", "FullAudit") } {
                $monitorData = Monitor-DigitalSignage -Screens $screens
                $signageData = @{
                    Screens = $screens
                    Contents = $contents
                    MonitorData = $monitorData
                }
            }
            { $_ -in @("Analytics", "FullAudit") } {
                $contentAnalysis = Analyze-SignageUsage -Contents $contents
                if (-not $signageData) { $signageData = @{} }
                $signageData.ContentAnalysis = $contentAnalysis
                $signageData.Screens = $screens
                $signageData.Contents = $contents
            }
        }
    }
    
    # Traitement VoIP
    if ($System -in @("VoIP", "Both")) {
        Write-Host ""
        $calls = Generate-VoIPData -Days $DaysAnalysis
        $users = Generate-VoIPUsers
        
        Write-Host ""
        
        switch ($Action) {
            { $_ -in @("QualityCheck", "FullAudit") } {
                $qualityData = Monitor-VoIPQuality -Calls $calls
                $voipData = @{
                    Calls = $calls
                    Users = $users
                    QualityData = $qualityData
                }
            }
            { $_ -in @("Analytics", "FullAudit") } {
                $usageAnalysis = Analyze-VoIPUsage -Calls $calls -Users $users
                if (-not $voipData) { $voipData = @{} }
                $voipData.UsageAnalysis = $usageAnalysis
                $voipData.Calls = $calls
                $voipData.Users = $users
            }
        }
    }
    
    # Export
    Write-Host ""
    Export-FacilitiesReports -Path $ExportPath -SystemType $System -SignageData $signageData -VoIPData $voipData
    
    Write-Host ""
    Write-Log "✓ Action '$Action' terminée avec succès!" -Level "Success"
    Write-Host ""
    Write-Log "Consultez les rapports dans: $ExportPath" -Level "Info"
    Write-Host ""
}

# Point d'entrée
try {
    Main
}
catch {
    Write-Log "✗ Erreur non gérée: $_" -Level "Error"
    exit 1
}