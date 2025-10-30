<#
.SYNOPSIS
    Script de gestion et monitoring des réservations de salles de réunion
    
.DESCRIPTION
    Ce script supervise l'utilisation des salles de réunion, identifie les no-shows,
    optimise l'occupation et génère des rapports d'utilisation détaillés.
    Adapté pour Product Manager Facilities chez Veolia.
    
.PARAMETER Action
    Action à effectuer: Monitor, Analytics, Optimize, NoShow, FullAudit
    Monitor: Surveillance temps-réel des réservations
    Analytics: Analyse d'utilisation et tendances
    Optimize: Recommandations d'optimisation
    NoShow: Détection et gestion des no-shows
    FullAudit: Audit complet (toutes actions)
    
.PARAMETER ExportPath
    Chemin d'export des rapports (par défaut : C:\RoomBookingReports)
    
.PARAMETER DaysAnalysis
    Nombre de jours d'historique à analyser (par défaut : 30)
    
.PARAMETER NoShowThreshold
    Seuil en minutes pour considérer un no-show (par défaut : 15)
    
.EXAMPLE
    .\Manage-RoomBooking.ps1 -Action Monitor
    
.EXAMPLE
    .\Manage-RoomBooking.ps1 -Action FullAudit -ExportPath "D:\Reports"
    
.EXAMPLE
    .\Manage-RoomBooking.ps1 -Action Analytics -DaysAnalysis 90
    
.NOTES
    Auteur: Consultant Facilities Management
    Version: 2.0
    Date: 2025-10-26
    
    Prerequisites:
    - PowerShell 5.0 ou supérieur
    - Module Microsoft.Graph (pour Outlook Calendar / Exchange)
    - Permissions: Calendars.Read.All
    
    Note: Script utilise données simulées pour démo
    Pour production, connecter à Exchange Online/Outlook API
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Monitor", "Analytics", "Optimize", "NoShow", "FullAudit")]
    [string]$Action,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = "C:\RoomBookingReports",
    
    [Parameter(Mandatory=$false)]
    [int]$DaysAnalysis = 30,
    
    [Parameter(Mandatory=$false)]
    [int]$NoShowThreshold = 15
)

#========================================
# CONFIGURATION GLOBALE
#========================================

$ErrorActionPreference = "Continue"
$VerbosePreference = "Continue"

$Global:Config = @{
    BookingSystemURL = "https://roombooking.veolia.local"
    WorkingHoursStart = 8
    WorkingHoursEnd = 19
    NoShowPenaltyEnabled = $true
    OptimalOccupancyRate = 0.70
    MinBookingMinutes = 30
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
# FONCTION 1: GÉNÉRATION DONNÉES SALLES
#========================================

function Generate-RoomData {
    try {
        Write-Log "Génération des données de salles de réunion..." -Level "Info"
        
        $rooms = @()
        $locations = @("Paris-HQ", "Aubervilliers-LeV", "Lyon", "Bordeaux", "Marseille")
        $roomTypes = @("Petit (4 pers)", "Moyen (8 pers)", "Grand (12 pers)", "XL (20 pers)", "Auditorium (50 pers)")
        
        for ($i = 1; $i -le 50; $i++) {
            $capacity = switch ($roomTypes | Get-Random) {
                "Petit (4 pers)" { 4 }
                "Moyen (8 pers)" { 8 }
                "Grand (12 pers)" { 12 }
                "XL (20 pers)" { 20 }
                "Auditorium (50 pers)" { 50 }
            }
            
            $room = [PSCustomObject]@{
                RoomID = "ROOM-$('{0:D3}' -f $i)"
                RoomName = "Salle $('{0:D2}' -f (Get-Random -Minimum 1 -Maximum 99))"
                Location = $locations | Get-Random
                Floor = Get-Random -Minimum 1 -Maximum 10
                Capacity = $capacity
                Type = $roomTypes | Get-Random
                Equipment = @("Écran", "Visio", "Tableau blanc") | Get-Random -Count (Get-Random -Minimum 1 -Maximum 3)
                Status = if ((Get-Random -Minimum 0 -Maximum 100) -gt 10) { "Available" } else { "Maintenance" }
            }
            $rooms += $room
        }
        
        Write-Log "✓ $($rooms.Count) salles générées" -Level "Success"
        return $rooms
    }
    catch {
        Write-Log "✗ Erreur génération salles: $_" -Level "Error"
        return $null
    }
}

#========================================
# FONCTION 2: GÉNÉRATION RÉSERVATIONS
#========================================

function Generate-BookingData {
    param(
        [array]$Rooms,
        [int]$Days = 30
    )
    
    try {
        Write-Log "Génération des données de réservations ($Days jours)..." -Level "Info"
        
        $bookings = @()
        $startDate = (Get-Date).AddDays(-$Days)
        
        # Générer réservations historiques
        for ($d = 0; $d -lt $Days; $d++) {
            $currentDate = $startDate.AddDays($d)
            
            # Skip weekends
            if ($currentDate.DayOfWeek -in @("Saturday", "Sunday")) {
                continue
            }
            
            # Générer 3-8 réservations par jour
            $bookingsPerDay = Get-Random -Minimum 3 -Maximum 8
            
            for ($b = 0; $b -lt $bookingsPerDay; $b++) {
                $room = $Rooms | Get-Random
                $startHour = Get-Random -Minimum $Global:Config.WorkingHoursStart -Maximum ($Global:Config.WorkingHoursEnd - 2)
                $duration = @(30, 60, 90, 120) | Get-Random
                
                $bookingStart = $currentDate.Date.AddHours($startHour)
                $bookingEnd = $bookingStart.AddMinutes($duration)
                
                $organizer = "user$(Get-Random -Minimum 1 -Maximum 100)@veolia.com"
                $attendees = Get-Random -Minimum 2 -Maximum $room.Capacity
                
                # Simuler no-show (10% des cas)
                $isNoShow = (Get-Random -Minimum 0 -Maximum 100) -lt 10
                
                # Simuler check-in (85% si pas no-show)
                $checkedIn = if (-not $isNoShow) {
                    (Get-Random -Minimum 0 -Maximum 100) -lt 85
                } else {
                    $false
                }
                
                $booking = [PSCustomObject]@{
                    BookingID = "BK-$(Get-Date -Format 'yyyyMMdd')-$(Get-Random -Minimum 1000 -Maximum 9999)"
                    RoomID = $room.RoomID
                    RoomName = $room.RoomName
                    Location = $room.Location
                    Capacity = $room.Capacity
                    Organizer = $organizer
                    AttendeesCount = $attendees
                    StartTime = $bookingStart
                    EndTime = $bookingEnd
                    DurationMinutes = $duration
                    Status = if ($bookingStart -lt (Get-Date)) {
                        if ($isNoShow) { "No-Show" }
                        elseif ($checkedIn) { "Completed" }
                        else { "Completed (No Check-in)" }
                    } else {
                        "Upcoming"
                    }
                    CheckedIn = $checkedIn
                    IsNoShow = $isNoShow
                    OccupancyRate = [math]::Round(($attendees / $room.Capacity) * 100, 2)
                }
                $bookings += $booking
            }
        }
        
        Write-Log "✓ $($bookings.Count) réservations générées" -Level "Success"
        return $bookings
    }
    catch {
        Write-Log "✗ Erreur génération réservations: $_" -Level "Error"
        return $null
    }
}

#========================================
# FONCTION 3: MONITORING TEMPS RÉEL
#========================================

function Monitor-RoomBookings {
    param(
        [array]$Rooms,
        [array]$Bookings
    )
    
    try {
        Write-Log "Monitoring des réservations en cours..." -Level "Info"
        Write-Host ""
        
        $now = Get-Date
        $todayBookings = $Bookings | Where-Object {
            $_.StartTime.Date -eq $now.Date
        }
        
        $currentBookings = $todayBookings | Where-Object {
            $_.StartTime -le $now -and $_.EndTime -ge $now
        }
        
        $upcomingBookings = $todayBookings | Where-Object {
            $_.StartTime -gt $now -and $_.StartTime -lt $now.AddHours(2)
        }
        
        $stats = @{
            TotalRooms = $Rooms.Count
            AvailableRooms = ($Rooms | Where-Object { $_.Status -eq "Available" }).Count - $currentBookings.Count
            CurrentlyBooked = $currentBookings.Count
            TodayTotal = $todayBookings.Count
            UpcomingNext2h = $upcomingBookings.Count
            TodayOccupancyRate = if ($todayBookings.Count -gt 0) {
                [math]::Round((($todayBookings | Measure-Object -Property DurationMinutes -Sum).Sum / 
                    (($Global:Config.WorkingHoursEnd - $Global:Config.WorkingHoursStart) * 60 * $Rooms.Count)) * 100, 2)
            } else { 0 }
        }
        
        Write-Log "ÉTAT ACTUEL:" -Level "Info"
        Write-Host "  • Salles totales: $($stats.TotalRooms)"
        Write-Host "  • Salles disponibles maintenant: $($stats.AvailableRooms)"
        Write-Host "  • Réservations en cours: $($stats.CurrentlyBooked)"
        Write-Host "  • Réservations aujourd'hui: $($stats.TodayTotal)"
        Write-Host "  • Prochaines 2h: $($stats.UpcomingNext2h)"
        Write-Host "  • Taux occupation aujourd'hui: $($stats.TodayOccupancyRate)%"
        Write-Host ""
        
        if ($currentBookings.Count -gt 0) {
            Write-Log "RÉSERVATIONS EN COURS:" -Level "Info"
            $currentBookings | Select-Object -First 5 | ForEach-Object {
                Write-Host "  • $($_.RoomName) [$($_.Location)]"
                Write-Host "    $($_.StartTime.ToString('HH:mm')) - $($_.EndTime.ToString('HH:mm')) | $($_.AttendeesCount) pers"
            }
            Write-Host ""
        }
        
        return @{
            Statistics = $stats
            CurrentBookings = $currentBookings
            UpcomingBookings = $upcomingBookings
        }
    }
    catch {
        Write-Log "✗ Erreur monitoring: $_" -Level "Error"
        return $null
    }
}

#========================================
# FONCTION 4: ANALYTICS D'UTILISATION
#========================================

function Analyze-RoomUsage {
    param(
        [array]$Rooms,
        [array]$Bookings
    )
    
    try {
        Write-Log "Analyse d'utilisation des salles..." -Level "Info"
        Write-Host ""
        
        # Analyse par salle
        $roomAnalysis = $Rooms | ForEach-Object {
            $room = $_
            $roomBookings = $Bookings | Where-Object { $_.RoomID -eq $room.RoomID }
            
            [PSCustomObject]@{
                RoomID = $room.RoomID
                RoomName = $room.RoomName
                Location = $room.Location
                Capacity = $room.Capacity
                TotalBookings = $roomBookings.Count
                TotalHours = [math]::Round(($roomBookings | Measure-Object -Property DurationMinutes -Sum).Sum / 60, 2)
                AvgOccupancyRate = if ($roomBookings.Count -gt 0) {
                    [math]::Round(($roomBookings | Measure-Object -Property OccupancyRate -Average).Average, 2)
                } else { 0 }
                NoShowCount = ($roomBookings | Where-Object { $_.IsNoShow }).Count
                NoShowRate = if ($roomBookings.Count -gt 0) {
                    [math]::Round((($roomBookings | Where-Object { $_.IsNoShow }).Count / $roomBookings.Count) * 100, 2)
                } else { 0 }
                Utilization = "Normal"
            }
        } | Sort-Object TotalBookings -Descending
        
        # Top 5 salles les plus utilisées
        Write-Log "TOP 5 SALLES LES PLUS UTILISÉES:" -Level "Info"
        $roomAnalysis | Select-Object -First 5 | ForEach-Object {
            Write-Host "  • $($_.RoomName) [$($_.Location)]: $($_.TotalBookings) réservations | $($_.TotalHours)h"
        }
        Write-Host ""
        
        # Analyse par localisation
        $locationAnalysis = $Bookings | Group-Object Location | ForEach-Object {
            [PSCustomObject]@{
                Location = $_.Name
                TotalBookings = $_.Count
                AvgDuration = [math]::Round(($_.Group | Measure-Object -Property DurationMinutes -Average).Average, 2)
                NoShowRate = [math]::Round((($_.Group | Where-Object { $_.IsNoShow }).Count / $_.Count) * 100, 2)
            }
        } | Sort-Object TotalBookings -Descending
        
        # Analyse par heure de la journée
        $hourlyAnalysis = $Bookings | Group-Object { $_.StartTime.Hour } | ForEach-Object {
            [PSCustomObject]@{
                Hour = "$($_.Name):00"
                BookingCount = $_.Count
            }
        } | Sort-Object { [int]$_.Hour.Split(':')[0] }
        
        Write-Log "✓ Analyses générées" -Level "Success"
        
        return @{
            RoomAnalysis = $roomAnalysis
            LocationAnalysis = $locationAnalysis
            HourlyAnalysis = $hourlyAnalysis
        }
    }
    catch {
        Write-Log "✗ Erreur analyse: $_" -Level "Error"
        return $null
    }
}

#========================================
# FONCTION 5: OPTIMISATION
#========================================

function Optimize-RoomConfiguration {
    param(
        [array]$Rooms,
        [object]$Analytics
    )
    
    try {
        Write-Log "Génération recommandations d'optimisation..." -Level "Info"
        Write-Host ""
        
        $recommendations = @()
        
        # Salles sous-utilisées
        $underutilized = $Analytics.RoomAnalysis | Where-Object { 
            $_.TotalBookings -lt 10 
        } | Select-Object -First 5
        
        foreach ($room in $underutilized) {
            $recommendations += [PSCustomObject]@{
                RoomID = $room.RoomID
                RoomName = $room.RoomName
                Issue = "Sous-utilisation"
                CurrentBookings = $room.TotalBookings
                Recommendation = "Revoir nécessité de cette salle ou améliorer visibilité"
                Priority = "Moyenne"
                PotentialSaving = "Réaffectation espace possible"
            }
        }
        
        # Salles sur-utilisées
        $overutilized = $Analytics.RoomAnalysis | Where-Object { 
            $_.TotalBookings -gt 50 
        } | Select-Object -First 5
        
        foreach ($room in $overutilized) {
            $recommendations += [PSCustomObject]@{
                RoomID = $room.RoomID
                RoomName = $room.RoomName
                Issue = "Sur-utilisation"
                CurrentBookings = $room.TotalBookings
                Recommendation = "Ajouter salle similaire ou limiter durées réservations"
                Priority = "Haute"
                PotentialSaving = "Amélioration satisfaction utilisateurs"
            }
        }
        
        # Salles avec high no-show rate
        $highNoShow = $Analytics.RoomAnalysis | Where-Object { 
            $_.NoShowRate -gt 15 
        } | Select-Object -First 5
        
        foreach ($room in $highNoShow) {
            $recommendations += [PSCustomObject]@{
                RoomID = $room.RoomID
                RoomName = $room.RoomName
                Issue = "Taux no-show élevé ($($room.NoShowRate)%)"
                CurrentBookings = $room.TotalBookings
                Recommendation = "Activer système de pénalité ou check-in obligatoire"
                Priority = "Haute"
                PotentialSaving = "Réduction $($room.NoShowCount) no-shows"
            }
        }
        
        Write-Log "✓ $($recommendations.Count) recommandations générées" -Level "Success"
        return $recommendations
    }
    catch {
        Write-Log "✗ Erreur optimisation: $_" -Level "Error"
        return $null
    }
}

#========================================
# FONCTION 6: DÉTECTION NO-SHOWS
#========================================

function Detect-NoShows {
    param(
        [array]$Bookings
    )
    
    try {
        Write-Log "Détection des no-shows..." -Level "Info"
        Write-Host ""
        
        $noShows = $Bookings | Where-Object { $_.IsNoShow -eq $true }
        
        if ($noShows.Count -gt 0) {
            Write-Log "⚠ NO-SHOWS DÉTECTÉS:" -Level "Warning"
            
            # Analyse par organisateur
            $noShowByOrganizer = $noShows | Group-Object Organizer | ForEach-Object {
                [PSCustomObject]@{
                    Organizer = $_.Name
                    NoShowCount = $_.Count
                    TotalWastedMinutes = ($_.Group | Measure-Object -Property DurationMinutes -Sum).Sum
                    LastNoShow = ($_.Group | Sort-Object StartTime -Descending | Select-Object -First 1).StartTime
                    Action = if ($_.Count -gt 3) { "Avertissement" } else { "Surveillance" }
                }
            } | Sort-Object NoShowCount -Descending
            
            Write-Host ""
            Write-Host "  TOP 5 UTILISATEURS NO-SHOWS:" -ForegroundColor Yellow
            $noShowByOrganizer | Select-Object -First 5 | ForEach-Object {
                Write-Host "    • $($_.Organizer): $($_.NoShowCount) no-shows | $($_.TotalWastedMinutes) min perdues"
            }
            Write-Host ""
            
            $totalWasted = ($noShows | Measure-Object -Property DurationMinutes -Sum).Sum
            Write-Log "Total temps perdu no-shows: $totalWasted minutes ($([math]::Round($totalWasted / 60, 2)) heures)" -Level "Warning"
            
            return @{
                NoShows = $noShows
                ByOrganizer = $noShowByOrganizer
                TotalWastedMinutes = $totalWasted
            }
        }
        else {
            Write-Log "✓ Aucun no-show détecté" -Level "Success"
            return $null
        }
    }
    catch {
        Write-Log "✗ Erreur détection no-shows: $_" -Level "Error"
        return $null
    }
}

#========================================
# FONCTION 7: EXPORT RAPPORTS
#========================================

function Export-AllReports {
    param(
        [string]$Path,
        [array]$Rooms,
        [array]$Bookings,
        [hashtable]$MonitorData,
        [hashtable]$Analytics,
        [array]$Recommendations,
        [object]$NoShowData
    )
    
    try {
        Write-Log "Export des rapports..." -Level "Info"
        Write-Host ""
        
        if (!(Test-Path $Path)) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        
        # Export salles
        $roomPath = Join-Path $Path "RoomBooking_Rooms_$timestamp.csv"
        $Rooms | Export-Csv -Path $roomPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        Write-Log "  ✓ Salles: $roomPath" -Level "Success"
        
        # Export réservations
        $bookingPath = Join-Path $Path "RoomBooking_Bookings_$timestamp.csv"
        $Bookings | Export-Csv -Path $bookingPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        Write-Log "  ✓ Réservations: $bookingPath" -Level "Success"
        
        # Export analyses
        if ($Analytics) {
            $analyticsPath = Join-Path $Path "RoomBooking_Analytics_$timestamp.csv"
            $Analytics.RoomAnalysis | Export-Csv -Path $analyticsPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            Write-Log "  ✓ Analyses: $analyticsPath" -Level "Success"
        }
        
        # Export recommandations
        if ($Recommendations) {
            $recoPath = Join-Path $Path "RoomBooking_Recommendations_$timestamp.csv"
            $Recommendations | Export-Csv -Path $recoPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            Write-Log "  ✓ Recommandations: $recoPath" -Level "Success"
        }
        
        # Export no-shows
        if ($NoShowData) {
            $noShowPath = Join-Path $Path "RoomBooking_NoShows_$timestamp.csv"
            $NoShowData.ByOrganizer | Export-Csv -Path $noShowPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            Write-Log "  ✓ No-Shows: $noShowPath" -Level "Success"
        }
        
        # Rapport synthétique
        $summaryPath = Join-Path $Path "RoomBooking_Summary_$timestamp.txt"
        $summary = @"
═══════════════════════════════════════════════════════════════════════════════
RAPPORT GESTION SALLES DE RÉUNION - VEOLIA
═══════════════════════════════════════════════════════════════════════════════
Généré le: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
Période analysée: $DaysAnalysis jours

═══════════════════════════════════════════════════════════════════════════════
PARC DE SALLES
═══════════════════════════════════════════════════════════════════════════════
Total salles: $($Rooms.Count)
Salles disponibles: $(($Rooms | Where-Object { $_.Status -eq "Available" }).Count)
Salles en maintenance: $(($Rooms | Where-Object { $_.Status -eq "Maintenance" }).Count)

═══════════════════════════════════════════════════════════════════════════════
UTILISATION
═══════════════════════════════════════════════════════════════════════════════
Total réservations: $($Bookings.Count)
Taux occupation moyen: $([math]::Round(($Bookings | Measure-Object -Property OccupancyRate -Average).Average, 2))%
Durée moyenne réservation: $([math]::Round(($Bookings | Measure-Object -Property DurationMinutes -Average).Average, 2)) min

NO-SHOWS:
Total no-shows: $(($Bookings | Where-Object { $_.IsNoShow }).Count)
Taux no-show: $([math]::Round((($Bookings | Where-Object { $_.IsNoShow }).Count / $Bookings.Count) * 100, 2))%
$(if ($NoShowData) { "Temps perdu: $([math]::Round($NoShowData.TotalWastedMinutes / 60, 2)) heures" })

═══════════════════════════════════════════════════════════════════════════════
TOP 3 SALLES LES PLUS UTILISÉES
═══════════════════════════════════════════════════════════════════════════════
$(if ($Analytics) {
    $Analytics.RoomAnalysis | Select-Object -First 3 | ForEach-Object {
        "• $($_.RoomName) [$($_.Location)]: $($_.TotalBookings) réservations`n"
    }
})

═══════════════════════════════════════════════════════════════════════════════
RECOMMANDATIONS PRIORITAIRES
═══════════════════════════════════════════════════════════════════════════════
$(if ($Recommendations) {
    $Recommendations | Where-Object { $_.Priority -eq "Haute" } | Select-Object -First 5 | ForEach-Object {
        "• $($_.RoomName): $($_.Recommendation)`n"
    }
})

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
║              GESTION SALLES DE RÉUNION - VEOLIA                               ║
║          Monitoring, Analytics & Optimisation                                 ║
║                                                                               ║
╚═══════════════════════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Cyan
    
    Write-Log "Configuration" -Level "Info"
    Write-Log "  • Action: $Action" -Level "Info"
    Write-Log "  • Export Path: $ExportPath" -Level "Info"
    Write-Log "  • Période analyse: $DaysAnalysis jours" -Level "Info"
    Write-Host ""
    
    # Génération données
    $rooms = Generate-RoomData
    $bookings = Generate-BookingData -Rooms $rooms -Days $DaysAnalysis
    
    Write-Host ""
    
    switch ($Action) {
        "Monitor" {
            Write-Log "Exécution: MONITOR" -Level "Info"
            Write-Host ""
            $monitorData = Monitor-RoomBookings -Rooms $rooms -Bookings $bookings
            Export-AllReports -Path $ExportPath -Rooms $rooms -Bookings $bookings -MonitorData $monitorData
        }
        
        "Analytics" {
            Write-Log "Exécution: ANALYTICS" -Level "Info"
            Write-Host ""
            $analytics = Analyze-RoomUsage -Rooms $rooms -Bookings $bookings
            Export-AllReports -Path $ExportPath -Rooms $rooms -Bookings $bookings -Analytics $analytics
        }
        
        "Optimize" {
            Write-Log "Exécution: OPTIMIZE" -Level "Info"
            Write-Host ""
            $analytics = Analyze-RoomUsage -Rooms $rooms -Bookings $bookings
            $recommendations = Optimize-RoomConfiguration -Rooms $rooms -Analytics $analytics
            Export-AllReports -Path $ExportPath -Rooms $rooms -Bookings $bookings -Analytics $analytics -Recommendations $recommendations
        }
        
        "NoShow" {
            Write-Log "Exécution: NO-SHOW DETECTION" -Level "Info"
            Write-Host ""
            $noShowData = Detect-NoShows -Bookings $bookings
            Export-AllReports -Path $ExportPath -Rooms $rooms -Bookings $bookings -NoShowData $noShowData
        }
        
        "FullAudit" {
            Write-Log "Exécution: FULL AUDIT" -Level "Info"
            Write-Host ""
            $monitorData = Monitor-RoomBookings -Rooms $rooms -Bookings $bookings
            $analytics = Analyze-RoomUsage -Rooms $rooms -Bookings $bookings
            $recommendations = Optimize-RoomConfiguration -Rooms $rooms -Analytics $analytics
            $noShowData = Detect-NoShows -Bookings $bookings
            Export-AllReports -Path $ExportPath -Rooms $rooms -Bookings $bookings -MonitorData $monitorData -Analytics $analytics -Recommendations $recommendations -NoShowData $noShowData
        }
    }
    
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