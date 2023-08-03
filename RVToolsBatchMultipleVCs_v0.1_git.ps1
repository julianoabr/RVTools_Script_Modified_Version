<#
.SYNOPSIS
With this example script you can start the the RVTools export all to xlsx function for multiple vCenter servers.
The output xlsx files will be merged to one xlsx file which will be mailed
	
.DESCRIPTION
With this example script you can start the the RVTools export all to xlsx function for multiple vCenter servers.
The output xlsx files will be merged to one xlsx file which will be mailed

.NOTES
  Upgrade to:
    Send XLSX without merging
    Send E-mail to different personas according to day of the month
    Use Different Credentials
    
.EXAMPLE
 RVTools-ExportMultipleVCs -oneVC -ptbrFormat -SendMail:$false -MergeExcelFiles:$false

.CREATEDBY
    Juliano Alves de Brito Ribeiro (Find me at julianoalvesbr@live.com or https://github.com/julianoabr/)

.VERSION INFO
    0.1

.TO THINK
    Seria possível que a vida evoluísse aleatoriamente a partir de matéria inorgânica? Não de acordo com os matemáticos.

    Nos últimos 30 anos, um número de cientistas proeminentes têm tentado calcular as probabilidades de que um organismo de vida livre e unicelular, como uma bactéria, pode resultar da combinação aleatória de blocos de construção pré-existentes. 
    Harold Morowitz calculou a probabilidade como sendo uma chance em 10^100.000.000.000
    Sir Fred Hoyle calculou a probabilidade de apenas as proteínas de amebas surgindo por acaso como uma chance em 10^40.000.

    ... As probabilidades calculadas por Morowitz e Hoyle são estarrecedoras. 
    Essas probabilidades levaram Fred Hoyle a afirmar que a probabilidade de geração espontânea 'é a mesma que a de que um tornado varrendo um pátio de sucata poderia montar um Boeing 747 com o conteúdo encontrado'. 
    Os matemáticos dizem que qualquer evento com uma improbabilidade maior do que uma chance em 10^50 faz parte do reino da metafísica - ou seja, um milagre.1

    1. Mark Eastman, MD, Creation by Design, T.W.F.T. Publishers, 1996, 21-22.


.SOURCE
Based on Script of Rob de Veij

# =============================================================================================================
# Script:    RVToolsBatchMultipleVCs.ps1
# Version:   0.2
# Date:      October, 20
# By:        Rob de Veij
# URL:       https://www.robware.net/rvtools/
# =============================================================================================================

#>


<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function RVTools-ExportMultipleVCs
{
    [CmdletBinding()]
    Param
    (
        # Switch param for only one vCenter
        [Parameter(Mandatory=$false,
                   Position=0)]
        [switch]$oneVC,

        # Switch param for multiple vCenter
        [Parameter(Mandatory=$false,
                   Position=1)]
        [switch]$multipleVC,
        
        # Param to include Annotations - run slowly
        [Parameter(Mandatory=$false,
                    Position=2)]
        [System.Boolean]$excludeCustomAnnotations=$true,

        # Param to Merge All XLSX Files in One. 
        [Parameter(Mandatory=$false,
                   Position=3)]
        [System.Boolean]$MergeExcelFiles = $false,

        # Param to run on a machine with Windows PT-BR Format 
        [Parameter(Mandatory=$false,
                   Position=4)]
        [switch]$ptbrFormat,

        # Param to run on a machine with Windows PT-BR Format 
        [Parameter(Mandatory=$false,
                   Position=5)]
        [switch]$enusFormat,

        # Param to Send Mail
        [Parameter(Mandatory=$false,
                   Position=6)]
        [System.Boolean]$SendMail = $false
        
    )

#Clean vCenter Server List
$vcServerList = @()

# Save current directory
$SaveCurrentDir = (get-location).Path

# Set RVTools Path. This is default installation path
[System.String] $RVToolsPath = "$env:SystemDrive\Program Files (x86)\Robware\RVTools"


#Set vCenter List Path
[System.String] $vcListPath = "$env:SystemDrive\Temp"


$ActualDate = (Get-Date -Format 'ddMMyyyy_HHmm').ToString()

$ActualYear = (Get-Date -Format 'yyyy').ToString()

$ActualMonth = (Get-date -Format 'MMMM').ToString()

$ActualDay = (Get-date -Format 'dd').ToString()


if ($oneVC.IsPresent){
    
    [System.Boolean]$MergeExcelFiles = $false
    
    Write-Host "You choose to run this script to capture information about one vCenter Only" -ForegroundColor White -BackgroundColor DarkGreen

    $vcServerList = @()

    $vcServerList = Read-Host "Type vCenter FQDN or vCenter IP Address that you want to generate RVTools Information"

}


if ($multipleVC.IsPresent){
    
    Write-Host "You choose to run this script to capture information about multiple vCenters" -ForegroundColor White -BackgroundColor DarkGreen

    $vcServerList = @()
    
    if (Test-Path -Path "$vcListPath\vclist.txt"){
    
        $vcServerList = (Get-content -Path "$vcListPath\vcList.txt")
    
    }
    else{
    
        Write-Host "Path does not contains any file with name 'vclist.txt'. Please validate file" -ForegroundColor White -BackgroundColor Red
    
    
    }     
    

}#end of Multiple VC


#validate main folder to export rvtools
$tmpOutput = "$env:SystemDrive:\temp\export\rvtools"

#Folder Year to test, if not exist, create it
$yearPath = $tmpOutput + '\' + $ActualYear

$yearPathExists = Test-Path -Path $yearPath

if ($yearPathExists){

    Write-Output "Folder $actualYear already exists"
          
}
else{

    Set-Location $tmpOutput

    New-Item -ItemType Directory -Path . -Name "$ActualYear"
    
}

#Folder Month to test, if not exist, create it
$monthPath = $yearPath + '\' + $ActualMonth

$monthPathExists = Test-Path -Path $monthPath

if ($monthPathExists){

    Write-Output "Folder $actualMonth already exists"
          
}
else{

    Set-Location $yearPath

    New-Item -ItemType Directory -Path . -Name "$actualMonth"
    
}

#Folder Day to test, if not exist, create it
$dayPath = $monthPath + '\' + $ActualDay

$dayPathExists = Test-Path -Path $dayPath

if ($dayPathExists){

    Write-Output "Folder $actualDay already exists"
          
}
else{

    Set-Location $monthPath

    New-Item -ItemType Directory -Path . -Name "$actualDay"
    
}

#Change to RVTools directory path
Set-Location $RVToolsPath

foreach ($vcServer in $vcServerList)
{
    
    [System.String] $User = "user\domain"                                                    # or use -passthroughAuth
    
    #To encrypt use an executable file called RVToolsPasswordEncryption.exe in path: C:\Program Files (x86)\Robware\RVTools#
    [System.String] $EncryptedPassword = "_RVToolsPWD8qGUqaayzk/FtSI5f6KiDRcYoJGJ7DW3kvgeXjmpMb8="

    [System.String] $altUser = "Administrator@vsphere.local"

    [System.String] $altEncryptedPassword = "_RVToolsPWD8qGUqaayzk/FtSI5f6KiDRcYoJGJ7DW3kvgeXjmpMb8="

    [System.String] $altTwoEncryptedPassword = "_RVToolsPWD8qGUqaayzk/FtSI5f6KiDRcYoJGJ7DW3kvgeXjmpMb8="
        
    [System.String] $XlsxDir = "$dayPath"

    $ShortVCServer = $vcServer.Split('.')[0]

    [System.String]$XlsxFile = "RVTools_Export_All_$shortVCServer-$ActualDate.xlsx"

    # Start cli of RVTools
    Write-Host "Start export for vCenter Server: $VCServer" -ForegroundColor DarkBlue -BackgroundColor White
    
    #Condition to use another username to connect to vCenter
    if ($ShortVCServer -like 'nameOne' -or $ShortVCServer -like 'nameTwo'){
        
        if ($excludeCustomAnnotations)
            {
        
                $Arguments = "-u $altUser -p $altEncryptedPassword -s $VCServer -c ExportAll2xlsx -d $XlsxDir -f $XlsxFile -DBColumnNames -ExcludeCustomAnnotations"

            }#end of If
            else{
    
                $Arguments = "-u $altUser -p $altEncryptedPassword -s $VCServer -c ExportAll2xlsx -d $XlsxDir -f $XlsxFile -DBColumnNames"
    
            }#end of Else
    
    
    
    }#end of IF short vc server itausa and western union
    elseif($shortVCServer -match 'nameFour*'){
    
        if ($excludeCustomAnnotations)
            {
        
                $Arguments = "-u $altUser -p ""$altTwoEncryptedPassword"" -s $VCServer -c ExportAll2xlsx -d $XlsxDir -f $XlsxFile -DBColumnNames -ExcludeCustomAnnotations"

            }#end of If
            else{
    
                $Arguments = "-u $altUser -p ""$altTwoEncryptedPassword"" -s $VCServer -c ExportAll2xlsx -d $XlsxDir -f $XlsxFile -DBColumnNames"
    
            }#end of Else
          
    
    }#end of elseif short vc server itausa and western union
    else{
    
    if ($excludeCustomAnnotations)
            {
        
                $Arguments = "-u $User -p $EncryptedPassword -s $VCServer -c ExportAll2xlsx -d $XlsxDir -f $XlsxFile -DBColumnNames -ExcludeCustomAnnotations"

            }#end of If
            else{
    
                $Arguments = "-u $User -p $EncryptedPassword -s $VCServer -c ExportAll2xlsx -d $XlsxDir -f $XlsxFile -DBColumnNames"
    
            }#end of Else
        
    
    }#end of else vc server itausa and western union and datalog
       
    Write-Host $Arguments

    $Process = Start-Process -FilePath ".\RVTools.exe" -ArgumentList $Arguments -NoNewWindow -Wait -PassThru

    if($Process.ExitCode -eq -1)
    {
    
        Write-Host "Error: Export failed! RVTools returned exitcode -1, probably a connection error! Script is stopped" -ForegroundColor Red
    
    exit 1
    
    }

#WAIT 5 SECONDS TO CONTINUE
Start-Sleep -Seconds 5 -Verbose

}#end of ForEach

# -----------------------------------------------
# Merge xlsx files vCenter1 + vCenter2 + vCenter3
# -----------------------------------------------
if ($MergeExcelFiles){

Set-Location $XlsxDir

$fileDate = "$actualDay" + "$actualMonth" + "$actualYear" 

$fileList = @()

$fileList = Get-ChildItem -File | Where-Object -FilterScript {$PSItem.Name -like "RvTools*" -and $PSItem.Name -like "*$fileDate*"} | Select-Object -ExpandProperty FullName

$finalFileList = $fileList -join ';'

$OutputFile = "$tmpOutput\ConsolidateSpreadSheetvCenters.xlsx"

Set-Location $RVToolsPath

& .\RVToolsMergeExcelFiles.exe -inputfiles "$finalFileList" -outputfile $OutputFile -overwrite -verbose

}
else{

Write-Output "I will not merge the RVTOOLS Exported Vcenter Xlsx Files"

}


#IF SENDMAIL EQUALS TRUE
if ($SendMail){

    Clear-Host

    #WAIT 20 SECONDS TO SEND REPORTS
    Start-Sleep -Seconds 20

    Set-Location $tmpOutput

    if ($ptbrFormat.IsPresent){
        
        #PT-BR FORMAT
        $fileFromToday = $null

        $fileFromToday = (Get-Date -Format "ddMMyyyy").ToString()
    
    }
    if ($enusFormat.IsPresent){
    
        #EN-US FORMAT
         $fileFromToday = $null

        $fileFromToday = (Get-Date -Format "MMddyyyy").ToString()
    
    }

    $fileLocation = $tmpOutput


    #Attach the Files Generated
    $tmpAttachs = @()

    $fileAttachs = @()

    $tmpAttachs = Get-ChildItem | Where-Object -FilterScript {$_.Name -like "*$fileFromToday*" -and $_.LastWriteTime -gt ((Get-date).AddMinutes(-100))} | Select-Object -ExpandProperty Name

    #ADD ATTACHMENTS
    foreach ($attach in $tmpAttachs){
    
        $attachment = $fileLocation + '\' + $attach
    
        $fileAttachs += $attachment
    }

        #put html file to send in e-mail
        $tmpHTML = Get-Content "$env:systemdrive\temp\contentRVToolsAll.html"

        $finalHTML = $tmpHTML | Out-String

        $tmpTwoWaysEmail = (Get-Date -Format "dd").ToString()

    #SEND E-MAIL TO DIFFERENT PERSONS ACCORDING TO DAY OF MONTH
    if (($tmpTwoWaysEmail -eq "05") -or ($tmpTwoWaysEmail -eq "15") -or ($tmpTwoWaysEmail -eq "25")){

        ###########Define Variables########

        $fromaddress = "powershellrobot@yourdomain.com"
        $toaddress = ("yourteam@yourdomain.com","yourteam@yourdomain.com")
        #$toaddress = "youruser@yourdomain.com"
        #$bccaddress = "youruser@yourdomain.com"
        #$CCaddress = "youruser@yourdomain.com"
        $CCaddress = "youruser@yourdomain.com"
        $RVSubject = "[TEAM VMWARE] RVTools Export - VCenters Domain XYZ"
        $RVattachments = $fileAttachs
        $smtpserver = "yourSMTPServer.yourdomain.com"

        ####################################

        Send-MailMessage -SmtpServer $smtpserver -From $fromaddress -To $toaddress -Cc $CCaddress -Subject $RVSubject -Body $finalHTML -BodyAsHtml -Attachments $RVattachments -Priority Normal -Encoding UTF8 -Verbose

    }#END OF IF SEND ACCORDING TO DATE
    else{

        ###########Define Variables########

        $fromaddress = "powershellrobot@yourdomain.com"
        $toaddress = ("yourteam@yourdomain.com","yourteam@yourdomain.com", "anotherteam@yourdomain.com")
        #$toaddress = "youruser@yourdomain.com"
        #$bccaddress = "youruser@yourdomain.com"
        #$CCaddress = "youruser@yourdomain.com"
        $CCaddress = "youruser@yourdomain.com"
        $RVSubject = "[TEAM VMWARE] RVTools Export - VCenters Domain XYZ"
        $RVattachments = $fileAttachs
        $smtpserver = "yourSMTPServer.yourdomain.com"

        ####################################

        Send-MailMessage -SmtpServer $smtpserver -From $fromaddress -To $toaddress -Cc $CCaddress -Subject $RVSubject -Body $finalHTML -BodyAsHtml -Attachments $RVattachments -Priority Normal -Encoding UTF8 -Verbose


    }#END OF ELSE SEND ACCORDING TO DATE


}#END OF IF SENDMAIL
else{

Write-Host "I will not send e-mail" -ForegroundColor White -BackgroundColor Red


}#END OF ELSE SENDMAIL


# Back to starting dir
Set-Location $SaveCurrentDir


}#End of Function

