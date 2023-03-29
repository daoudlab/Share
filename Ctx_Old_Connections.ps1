# =======================================================
# NAME: unregistered&Maintenance_V01.ps1
#
# AUTHOR: Nicolas DAOUDAL - EDR France
# DATE: 10/05/2022
#
# KEYWORDS: Citrix, Registration , Maintenance
# VERSION 1.0
# 
# COMMENTS: This script will liste Unregistered and In maintenance VDA in PROD and REC.
# It can be runed from any account, it will use Xml Credentials of srv-xddc accounts 
# Output log file will be set in \\bernstein\Utilisateurs\DSI\Production\Citrix\Scripts\Logs
#Requires -Version 2.0
# =======================================================
# Variables :
# ========================================================
#region Variable

<#Description du fonctionnement:
Fichiers nécessaires :
"C:\scripts\computers.txt" : contient une liste de noms d'ordinateurs à interroger.
"C:\scripts\Mailcreds.xml" : contient les informations d'identification pour l'envoi de courriels.
"fonctions.psm1" : un module de fonctions à importer.

Variables modifiables :
$interactiveLogging : booléen pour activer ou désactiver l'affichage des logs dans la console.
$workingDirectory : chemin du dossier de travail.
$From : adresse e-mail de l'expéditeur des courriels.
$To : adresse e-mail du destinataire des courriels.
$SmtpServer : adresse SMTP du serveur de messagerie.
$subject : sujet du courriel.
$WorkingFolders : liste des dossiers de travail à créer.
$
#>
# Activer l'affichage des logs dans la console. À désactiver pour les tâches planifiées.
$interactiveLogging = $true

# Chemin du dossier de travail
$workingDirectory = "\\bernstein\Utilisateurs\DSI\Production\Citrix\Scripts"

 # Retention Settings
 $Retention = (Get-Date).AddDays(-120)
 $Time4Maint = $Retention.AddDays(-90)
 $Time4clean = $Time4Maint.AddDays(-180)

# Nom du fichier PowerShell en cours d'exécution #(peut etre modifié)
$fileName = $MyInvocation.MyCommand.Name -replace ('.ps1','')

# Paramètres pour l'envoi de courriels
$subject = $fileName + "Report"
$smtpServer = $EdrMail.Smtp
#$Port = $EdrMail.Port # empty for port 25
$from = $EdrMail.From
$to = "n.daoudal@edr.com"
$Priority = "normal" # low, normal, high.

# DO NOT MODIFY BELOW
# ========================================================

# Liste des dossiers de travail à créer
$WorkingFolders = "Reports", "Logs", "Modules"

# Création des dossiers s'ils n'existent pas déjà
foreach ($folder in $WorkingFolders) {
    $path = Join-Path -Path $workingDirectory -ChildPath $folder
    if (-not (Test-Path $path)) {
        New-Item -ItemType Directory -Path $path | Out-Null
        Write-Warning "Le dossier $path à été crée"
    }
}

# Définition des noms de fichiers avec la date et l'heure actuelles
$date = Get-Date -Format "yyyyMMdd-HHmmss"
$logfile = Join-Path -Path $workingDirectory -ChildPath "Logs\$($fileName)_$date.log"
$reportfile = Join-Path -Path $workingDirectory -ChildPath "Reports\$($fileName)_$date.HtmlOutput"
$FunctionModule = Join-Path -Path $workingDirectory -ChildPath "Modules\fonctions.psm1"

# Importation du module de fonctions :
Import-Module $FunctionModule -DisableNameChecking -ErrorAction Stop

 $DailyExport = Get-LatestFile -Directory "$workingDirectory\Reports" -Filename "Citrix_Daily_Machines_Reporting"
 $AdUsersInfos = "$workingDirectory\Lists\AdUsersInfos.csv"

#endregion

# Code :
# ========================================================
Write-Log -Message "Le script commence à $(Get-Date)" -Level RESULT
$stopwatch = [Diagnostics.Stopwatch]::StartNew()

$UsersInfos = Import-Csv -Path $AdUsersInfos -Delimiter ";"
 
$BrokerInfos = Html2Table -HtmlFile $DailyExport

# Check if BrokerInfos file is not empty
if ($BrokerInfos) {
    # Create empty array to hold results
    $Results = @()

    foreach ($Desktop in $BrokerInfos) {
        if ($Desktop.LastConnectionTime -as [datetime] -lt $RetentionDate) {
            $DesktopObject = New-Object -TypeName PSObject -Property @{
                MachineName = $Desktop.MachineName
                Environment = $Desktop.Environment
                LastConnectionTime = $Desktop.LastConnectionTime
                AssociatedUsers = @()
            }
    
            foreach ($User in $Desktop.Users.Split(";")) {
                $LoginName = $User.Split(";")[0]
                $UserName = $User.Split(";")[1]
    
                $UserObject = New-Object -TypeName PSObject -Property @{
                    LoginName = $LoginName
                    Name = $UserName
                    AccountStatus = $true
                    Note = ""
                    Action = ""
                }
    
                $Account = $UsersInfos | Where-Object { $_.LoginName -eq $LoginName }
    
                if ($Account.Expired -eq $true -or $Account.Disabled -eq $true) {
    
                    $UserObject.Note = "Account is expired or disabled."
                    $UserObject.AccountStatus = $false
                }
    
                if ($Desktop.LastConnectionTime -as [datetime]  -lt $CleanupDate) {
                    $UserObject.Action = "Cleanup"
                }
                elseif ($Desktop.LastConnectionTime -as [datetime] -lt $MaintenanceDate) {
                    $UserObject.Action = "Maintenance"
                }
                else {
                    $UserObject.Action = "Mail2User"
                }
    
                $DesktopObject.AssociatedUsers += $UserObject
            }
    
            $Results += $DesktopObject
        }
    }

 
}
else {
    Write-Warning "BrokerInfos file is empty."
}

# Exportation des résultats vers un fichier HtmlOutput
$HtmlOutput = $Results | ConvertTo-Html -As Table -Head $EdRcss -Title $fileName
$HtmlOutput | Out-File -FilePath $reportfile

# Configuration des variables pour l'envoi d'email
$body = $HtmlOutput | Out-String #| Where-Object { $_.LastConnectionTime -as [datetime] -lt $Time4Maint } |Sort-Object { $_."LastConnectionTime" -as [datetime] } -Descending | Out-String

# Fin du script
$stopwatch.Stop()
$duration = $stopwatch.Elapsed
Write-Log -Message "Le script a terminé en $duration." -Level RESULT

# Envoi de l'email
Send-Email -smtpServer $smtpServer -from $from -to $to -subject $subject -body $body -Attachments $reportfile -Priority $Priority
