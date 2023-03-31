##################################################
#  Infos
##################################################
# Read the key from the file and decrypt the password file
$GmailKey = Get-Content \\daoudnas\Softwares\Softs\GitHub\github\Powershell\Modules\aes.key
$GmailPassword = Get-Content \\daoudnas\Softwares\Softs\GitHub\github\Powershell\Modules\gmailpassword.txt | ConvertTo-SecureString -Key $GmailKey

# Create a PSCredential object
$GmailCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $gmail.username, $GmailPassword

#gmail
$global:gmail = @{
    username = "daoudnas1664@gmail.com"
    smtp     = "smtp.gmail.com"
    cred     = $GmailCredential
    port     = "587"
}
$VcenterKey = Get-Content \\daoudnas\Softwares\Softs\GitHub\github\Powershell\Modules\aesvc.key
$VcenterPassword = Get-Content \\daoudnas\Softwares\Softs\GitHub\github\Powershell\Modules\vcpassword.txt | ConvertTo-SecureString -Key $VcenterKey

# Create a PSCredential object
$VcenterCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $Vcenter.username, $VcenterPassword
#vcenter
$global:Vcenter = @{
    username = "administrator@vsphere.local"
    Server = "vcenter.vsphere.local"
    Cred   = $VcenterCredential
}

##################################################
#  FUNCTIONS
##################################################
#region Funtions
Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True)]
        [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "RESULT")]
        [String]
        $Level,

        [Parameter(Mandatory = $True)]
        [string]
        $Message
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"
    
    if ($InteractiveLogging -and $PSBoundParameters.ContainsKey('Level')) {
        Write-Host $Line -ForegroundColor @{
            'INFO'   = 'Gray'
            'WARN'   = 'Yellow'
            'ERROR'  = 'Red'
            'FATAL'  = 'Magenta'
            'RESULT' = 'Green'
        }[$Level]
    }
    
    if ($LogFile) {
        $Line | Out-File -FilePath $LogFile -Append
    }
    else {
        Write-Host $Line -ForegroundColor @{
            'INFO'   = 'Gray'
            'WARN'   = 'Yellow'
            'ERROR'  = 'Red'
            'FATAL'  = 'Magenta'
            'RESULT' = 'Green'
        }[$Level]
    }
}

function Disable-Indexing {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Drive
    )

    $volume = Get-WmiObject -Class Win32_Volume -Filter "DriveLetter='$Drive'"
    if ($volume.IndexingEnabled) {
        Write-Log -Message "Indexing of drive $Drive disabled." -Level RESULT
        $volume | Set-WmiInstance -Arguments @{IndexingEnabled = $false } | Out-Null
        $Global:Info += "<br>Indexing of drive $Drive disabled."
    }
    else {
        Write-Log -Message "Drive $Drive indexing was already disabled." -Level RESULT
        $Global:Info += "<br>Indexing of drive $Drive already disabled."
    }
}

function Optimize-Drive {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$DriveLetter
    )

    $defragAnalysis = defrag $DriveLetter /A /V
    $trimAnalysis = Get-StorageReliabilityCounter -DriveLetter $DriveLetter | Select-Object -ExpandProperty DeviceScavengedRatio

    Write-Log -Message "Drive $DriveLetter analysis:" -Level INFO
    Write-Log -Message "----------------------------------" -Level INFO
    Write-Log -Message "Defrag analysis:" -Level INFO
    Write-Log -Message $defragAnalysis -Level INFO
    Write-Log -Message "Trim analysis:" -Level INFO
    Write-Log -Message $trimAnalysis -Level INFO

    if ($defragAnalysis -like "*defragmentation is not needed*") {
        Write-Log -Message "No defragmentation required." -Level INFO
    }
    else {
        Write-Log -Message "Defragmentation required." -Level WARN
        defrag $DriveLetter /V
    }

    if ($trimAnalysis -ge 1) {
        Write-Log -Message "Trim required." -Level WARN
        Optimize-Volume -DriveLetter $DriveLetter -ReTrim -Verbose
    }
    else {
        Write-Log -Message "No trim required." -Level INFO
    }
}

  
function Send-Email {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SmtpServer,
        
        [Parameter(Mandatory = $true)]
        [string]$From,
        
        [Parameter(Mandatory = $true)]
        [string]$To,
        
        [Parameter(Mandatory = $true)]
        [string]$Subject,
        
        [Parameter(Mandatory = $true)]
        [string]$Body,
        
        [string[]]$Attachments,
        
        [ValidateSet('Low', 'Normal', 'High')]
        [string]$Priority = 'Normal',
        
        [switch]$UseSsl,
        
        [int]$Port = 25,
        
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        # Create a new SMTP client object
        $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $Port)
        
        if ($UseSsl) {
        $smtpClient.EnableSsl = $true
        }

        if ($Credential) {
            $smtpClient.Credentials = $Credential
        }
        
        # Create a new MailMessage object
        $mailMessage = New-Object System.Net.Mail.MailMessage($From, $To)
        $mailMessage.Subject = $Subject
        $mailMessage.Body = $Body
        $mailMessage.IsBodyHtml = $true
        $mailMessage.Priority = [System.Net.Mail.MailPriority]::$Priority
        
        if ($Attachments) {
            foreach ($attachment in $Attachments) {
                $mailMessage.Attachments.Add($attachment)
            }
        }
        
        # Send the email
        $smtpClient.Send($mailMessage)
    }
    catch {
        Write-Error $_.Exception.Message
    }
}


function HTML2Table {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HtmlFile
    )

    # Lire le contenu du fichier HTML
    $html = New-Object -ComObject "HTMLFile"
    $source = Get-Content -Path $HtmlFile -Raw
    try {
        $html.IHTMLDocument2_write($source) 2> $null
    }
    catch {
        $encoded = [Text.Encoding]::Unicode.GetBytes($source)
        $html.write($encoded)
    }

    # Accéder à la première table du fichier HTML
    $table = $html.getElementsByTagName("table")[0]

    # Créer un tableau pour stocker les données de la table
    $data = @()

    # Parcourir les lignes de la table en ignorant la première ligne (l'en-tête)
    for ($rowIndex = 1; $rowIndex -lt $table.rows.length; $rowIndex++) {
        # Récupérer la ligne courante
        $row = $table.rows[$rowIndex]

        # Créer un objet pour stocker les données de la ligne
        $rowData = New-Object PSObject

        # Parcourir les cellules de la ligne
        for ($cellIndex = 0; $cellIndex -lt $row.cells.length; $cellIndex++) {
            # Récupérer le nom de la colonne à partir de l'en-tête de la table
            $columnName = $table.rows[0].cells[$cellIndex].innerText

            # Récupérer la valeur de la cellule
            $cellValue = $row.cells[$cellIndex].innerText

            # Ajouter la valeur de la cellule à l'objet rowData
            $rowData | Add-Member -MemberType NoteProperty -Name $columnName -Value $cellValue
        }

        # Ajouter l'objet rowData au tableau data
        $data += $rowData
    }

    # Retourner le tableau data
    return $data
}


function Get-LatestFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    # Récupérer tous les fichiers contenant le nom de fichier spécifié
    $files = Get-ChildItem -Path $Directory -Filter "*$FileName*"

    # Trier les fichiers par date de dernière modification en ordre décroissant
    $sortedFiles = $files | Sort-Object LastWriteTime -Descending

    # Retourner le premier fichier de la liste (le plus récent)
    return $sortedFiles[0]
}


#endregion

##################################################
#  STYLE
##################################################
$Font = '"Calibri", Tahoma, Geneva, Verdana, sans-serif'
$Font2 = '"Calibri Light", Tahoma, Geneva, Verdana, sans-serif'
$ColumnFont = '"Bahnschrift", Tahoma, Geneva, Verdana, sans-serif'
$global:EdRcss = @"
<style>
    /* Styles pour le titre */
    h1 {
        font-weight: bold;
        color: #F7CD13;
        background-color: #001B47;
        padding: 20px;
        font-family: $Font;
        font-size: 36px;
        text-align: center;
        text-shadow: 2px 2px 5px #757575;
        border-radius: 10px;
        margin-bottom: 30px;
    }

    /* Ajout du logo en haut de la page */
    .logo {
        display: flex;
        justify-content: center;
        margin-bottom: 30px;
    }

    .logo img {
        height: 100px;
        width: auto;
    }

    /* Styles pour la table */
    table {
        border-collapse: collapse;
        width: 100%;
        table-layout: auto;
    }

    /* Alternance de couleurs entre chaque ligne de résultat */
        table tr:nth-child(even) {
            background-color: #f9f9f9;
            color: #0D47A1;
        }

        table tr:nth-child(odd) {
            color: #F7CD13;
        }

        table td, table th {
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }

    /* Styles pour les en-têtes de colonne */
    table th {
        background-color: #F7CD13;
        color: #CF0123;
        font-weight: bold;
        font-family: $ColumnFont;
        font-size: 20px;
        text-align: center;
        padding: 15px;
        box-shadow: 2px 2px 5px #757575;
        border-radius: 10px;
    }

    /* Styles pour les cellules de la table */
    table td {
        font-family: $Font2;
        font-size: 18px;
        text-align: left;
        padding: 15px;
        border-bottom: 1px solid #dddddd;
    }
    
    body {
  background-color: #001B47;
}
</style>
<div class="logo">
    <img src="https://www.edmond-de-rothschild.com/style%20library/edrcom_common/img/logo.png">
</div>
"@

$global:EdRcssNoLogo = @"
<style>
    /* Styles pour le titre */
    h1 {
        font-weight: bold;
        color: #F7CD13;
        background-color: #001B47;
        padding: 20px;
        font-family: $Font;
        font-size: 36px;
        text-align: center;
        text-shadow: 2px 2px 5px #757575;
        border-radius: 10px;
        margin-bottom: 30px;
    }


  

    /* Styles pour la table */
    table {
        border-collapse: collapse;
        width: 100%;
        table-layout: auto;
    }

    /* Alternance de couleurs entre chaque ligne de résultat */
        table tr:nth-child(even) {
            background-color: #f9f9f9;
            color: #0D47A1;
        }

        table tr:nth-child(odd) {
            color: #F7CD13;
        }

        table td, table th {
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }

    /* Styles pour les en-têtes de colonne */
    table th {
        background-color: #F7CD13;
        color: #CF0123;
        font-weight: bold;
        font-family: $ColumnFont;
        font-size: 20px;
        text-align: center;
        padding: 15px;
        box-shadow: 2px 2px 5px #757575;
        border-radius: 10px;
    }

    /* Styles pour les cellules de la table */
    table td {
        font-family: $Font2;
        font-size: 18px;
        text-align: left;
        padding: 15px;
        border-bottom: 1px solid #dddddd;
    }
    
    body {
  background-color: #001B47;
}
</style>
"@

$global:DaoudCSS = @"


<style>
    /* Styles pour le titre */
    h1 {
        font-weight: bold;
        color: #F77F00;
        background-color: #ffffff;
        padding: 20px;
        font-family: $Font;
        font-size: 36px;
        text-align: center;
        text-shadow: 2px 2px 5px #757575;
        border-radius: 10px;
        margin-bottom: 30px;
    }

    /* Styles pour la table */
    table {
        border-collapse: collapse;
        width: 100%;
        table-layout: auto;
    }

    /* Alternance de couleurs entre chaque ligne de résultat */
    table tr:nth-child(even) {
        background-color: #f2f2f2;
    }

    table tr:nth-child(odd) {
        background-color: #ffffff;
    }

    /* Styles pour les en-têtes de colonne */
    table th {
        background-color: #F77F00;
        color: #ffffff;
        font-weight: bold;
        font-family: $ColumnFont;
        font-size: 20px;
        text-align: center;
        padding: 15px;
        box-shadow: 2px 2px 5px #757575;
        border-radius: 10px;
        cursor: pointer;
    }

    /* Styles pour les cellules de la table */
    table td {
        font-family: $Font2;
        font-size: 18px;
        text-align: left;
        padding: 15px;
        border-bottom: 1px solid #dddddd;
    }

    /* Styles pour la colonne triée */
    table th.sorted-asc,
    table th.sorted-desc {
        background-color: #CF0123;
        color: #ffffff;
    }

    /* Styles pour le filtre de colonne */
    .filterable {
        position: relative;
    }

    .filterable input[type=text] {
        width: 100%;
        padding: 5px 10px;
        border: none;
        border-radius: 3px;
        font-size: 14px;
        margin-bottom: 15px;
    }

    .filterable .filter-icon {
        position: absolute;
        top: 50%;
        right: 10px;
        transform: translateY(-50%);
        color: #888;
        font-size: 16px;
        cursor: pointer;
    }

    /* Styles pour le message de filtre vide */
    .filter-empty {
        font-style: italic;
        color: #888;
        text-align: center;
        margin-top: 20px;
    }

    body {
        background-color: #ffffff;
    }
</style>

"@