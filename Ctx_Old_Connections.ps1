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


# Strings
# =======================================================


# Retention Settings
$Retention = (Get-Date).AddDays(-120)
$Time4Maint = $Retention.AddDays(-90)
$Time4clean = $Time4Maint.AddDays(-180)

#Do not modify below
# =======================================================
$Date = Get-Date
$sw = [Diagnostics.Stopwatch]::StartNew()
$Results = $Null
$Results = @()
$Errors = $Null
$Unnasigned = $null
$continue = $false






#Script Start
# =======================================================

Write-Log -Message "Script Starting at $date" -Level RESULT
$BrokerInfos = Import-Csv -Path $DailyExport -Delimiter ";"
$UsersInfos = Import-Csv -Path $AdUsersInfos -Delimiter ";"


if (-not ([string]::IsNullOrEmpty($BrokerInfos))) {


           



    #No recent connections

    $Old = $BrokerInfos | Where-Object ({ $_.LastConnectionTime -as [datetime] -lt $Retention }) | Select-Object MachineName, Env, LastConnectionTime, AssociatedUserNames, AssociatedUserFullNames, DesktopGroupName, Catalog | Sort-Object { $_."LastConnectionTime" -as [datetime] } -Descending

    Write-Log -Message "collecting No recent connections Machine ($($Old.Count))" -Level INFO





    foreach ($User in $Old) {



        $Job = $null
        $Job = New-Object -TypeName psobject



        #check if multiples accounts are linked to same VDI
        if ($User.AssociatedUserNames -like "*;*") {
            $splitLogin = $User.AssociatedUserNames -split ";"
            $UserLogin = $splitLogin[0].replace("LCF\", "")
            $SplitName = $User.AssociatedUserFullNames -split ";"
            $UserFullName = $SplitName[0]

            if ($null -ne $splitLogin[1]) { $Job | Add-Member -MemberType NoteProperty -Name "AssociatedUserNames1" -Value $splitLogin[1] -Force }
            if ($null -ne $splitLogin[2]) { $Job | Add-Member -MemberType NoteProperty -Name "AssociatedUserNames2" -Value $splitLogin[2] -Force }
            if ($null -ne $splitLogin[3]) { $Job | Add-Member -MemberType NoteProperty -Name "AssociatedUserNames3" -Value $splitLogin[3] -Force }
            $Job | Add-Member -MemberType NoteProperty -Name "Notes" -Value "multiple Users: $($SplitName.count)" -Force
        }
        else {

            $UserLogin = $User.AssociatedUserNames.replace("LCF\", "") 
            $UserFullName = $User.AssociatedUserFullNames
        }

        $UserInfo = $UsersInfos | Where-Object { $_.SamAccountName -eq $UserLogin } | Select-Object *

        # check if account is expired or Disabled

        if (-not ([string]::IsNullOrEmpty($UserInfo.AccountExpirationDate)) -and ($UserInfo.AccountExpirationDate -as [datetime] -lt $Date -or $UserInfo.Enabled -eq "False") ) {
            $action = "Clean Now!"
            $Job | Add-Member -MemberType NoteProperty -Name "Notes" -Value "Account expired or disabled" -Force
            $Job | Add-Member -MemberType NoteProperty -Name "AccountStatus" -Value $false -Force
        }else {
   

            $Job | Add-Member -MemberType NoteProperty -Name "AccountStatus" -Value $UserInfo.Enabled -Force

            #Set actions according to retention time
            # 5 Month = Notify user by mail
            # 6 Month = Set VDI in maintenance
            # 12 Month = Cleanup and release VDI

            if ($User.LastConnectionTime -as [datetime] -lt $Time4clean) {
                $action = "Cleanup"
            }elseif ($User.LastConnectionTime -as [datetime] -ge $Time4clean -and $User.LastConnectionTime -as [datetime] -lt $Time4Maint) {
                $action = "Maintenance"
            }else { $action = "Mail2User" }
    
        }

        $Job | Add-Member -MemberType NoteProperty -Name "MachineName" -Value $User.MachineName -Force
        $Job | Add-Member -MemberType NoteProperty -Name "Environnement" -Value $User.Environnement -Force
        $Job | Add-Member -MemberType NoteProperty -Name "DesktopGroupName" -Value $User.DesktopGroupName -Force
        $Job | Add-Member -MemberType NoteProperty -Name "CatalogName" -Value $User.CatalogName -Force
        $Job | Add-Member -MemberType NoteProperty -Name "LastCitrixConnectionTime" -Value $User.LastConnectionTime -Force
        $Job | Add-Member -MemberType NoteProperty -Name "AssociatedUserNames" -Value $UserLogin -Force
        $Job | Add-Member -MemberType NoteProperty -Name "AssociatedUserFullNames" -Value $UserFullName -Force
        $Job | Add-Member -MemberType NoteProperty -Name "LastADLogonDate" -Value $UserInfo.LastLogonDate -Force
    
        $Job | Add-Member -MemberType NoteProperty -Name "AccountExpirationDate" -Value $UserInfo.AccountExpirationDate -Force
        $Job | Add-Member -MemberType NoteProperty -Name "lockedout" -Value $UserInfo.lockedout -Force
        #check if mail is empty
        if (-not ([string]::IsNullOrEmpty($UserInfo.EmailAddress)) -or ([string]::IsNullOrEmpty($UserFullName))) {
            $Job | Add-Member -MemberType NoteProperty -Name "EmailAddress" -Value $UserInfo.EmailAddress -Force }
        else {
        
            $Name, $FirstName = ($UserFullName.ToLower()).split(",") -replace ("[éèêë]", "e") -replace ("[ûüùú]", "u") -replace ("[àâäáã]", "a") -replace ("[öòôóõ]", "o") -replace ("ç", "c") -replace ("[îìï]", "i") -replace ("[ñ]", "n") -replace ('[^a-zA-Z0-9]', '')
        
            $Email = "$($FirstName[0]).$($Name)@EDR.COM" 
            $Job | Add-Member -MemberType NoteProperty -Name "EmailAddress" -Value $Email -Force
        }


        $Job | Add-Member -MemberType NoteProperty -Name "Action" -Value $action -Force       
            
        $Results += $Job

   
    }

    # Exportation des résultats vers un fichier HTML
    $html = $Results | ConvertTo-Html -As Table -Head $EdRcss -Title $fileName
    $html | Out-File -FilePath $reportfile

    # Configuration des variables pour l'envoi d'email
    $body = $html | Out-String

    # Fin du script
    $stopwatch.Stop()
    $duration = $stopwatch.Elapsed
    Write-Log -Message "Le script a terminé en $duration." -Level RESULT

    # Envoi de l'email
    Send-Email -smtpServer $smtpServer -from $from -to $to -subject $subject -body $body -Attachments $reportfile -Priority $Priority

}