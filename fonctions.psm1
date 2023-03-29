##################################################
#  Infos
##################################################
# Read the key from the file and decrypt the password file
$GmailKey = Get-Content \\daoudnas\Softwares\Softs\scripts\Powershell\Modules\aes.key
$GmailPassword = Get-Content \\daoudnas\Softwares\Softs\scripts\Powershell\Modules\gmailpassword.txt | ConvertTo-SecureString -Key $GmailKey

# Create a PSCredential object
$GmailUsername = "$GmailCredential.UserName"
$GmailCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $GmailUsername, $GmailPassword

#gmail
$global:gmail = @{
    username = "daoudnas1664@gmail.com"
    smtp     = "smtp.gmail.com"
    cred     = $GmailCredential
    port     = "587"
}
$VcenterKey = Get-Content \\daoudnas\Softwares\Softs\scripts\Powershell\Modules\aesvc.key
$VcenterPassword = Get-Content \\daoudnas\Softwares\Softs\scripts\Powershell\Modules\vcpassword.txt | ConvertTo-SecureString -Key $VcenterKey

# Create a PSCredential object
$VcenterUsername = "$VcenterCredential.UserName"
$VcenterCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $VcenterUsername, $VcenterPassword
#vcenter
$global:Vcenter = @{
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
        $mailParams = @{
            SmtpServer  = $SmtpServer
            From        = $From
            To          = $To
            Subject     = $Subject
            Body        = $Body
            BodyAsHtml  = $true
            Priority    = $Priority
            Encoding    = [System.Text.Encoding]::UTF8
            ErrorAction = 'Stop'
        }
        
        if ($Attachments) {
            $mailParams.Add('Attachments', $Attachments)
        }
        
        if ($UseSsl) {
            $mailParams.Add('UseSsl', $true)
        }
        
        if ($Credential) {
            $mailParams.Add('Credential', $Credential)
        }
        
        if ($Port) {
            $mailParams.Add('Port', $Port)
        }
        
        Send-MailMessage @mailParams
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

$global:EdRcss = @"
<style>
    /* Styles pour le titre */
    h1 {
        font-weight: bold;
        color: #F7CD13;
        background-color: #001B47;
        padding: 20px;
        font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
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
        font-family: "Bahnschrift", Tahoma, Geneva, Verdana, sans-serif;
        font-size: 20px;
        text-align: center;
        padding: 15px;
        box-shadow: 2px 2px 5px #757575;
        border-radius: 10px;
    }

    /* Styles pour les cellules de la table */
    table td {
        font-family: "Bahnschrift", Tahoma, Geneva, Verdana, sans-serif;
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



