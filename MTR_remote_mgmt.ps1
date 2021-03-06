<#
 .SYNOPSIS  
     This script is intended to serve as a backup management tool to address basic and initial tasks on a Microsoft Teams Room system based on Windows (MTRoW)
  .NOTES
     File Name  : MTR_remote_mgmt.ps1
     Author     : https://github.com/rito2k
#>

#
# FUNCTIONS
#
function Show-Menu{     
     if ($MTR_ready){
          Write-Host "(Target MTR: $MTR_hostName)" -ForegroundColor Green
     }
     else {
          Write-Host "(No target MTR!)" -ForegroundColor Red
     }
     $menuOptions | ForEach-Object {write-host $_}
}
function selectOpt1 {
     Clear-Host
     Write-Host "Please select Option 1 first to target MTR device by providing credentials for remote connection." -ForegroundColor Magenta
}
function remote_logoff{
     param (
          [Parameter()]
          [string]$Computer,
          [ValidateNotNull()]
          [System.Management.Automation.Credential()]
          [System.Management.Automation.PSCredential]$cred
     )
     $scriptBlock = {
          $ErrorActionPreference = 'Stop'      
          try {
              ## Find all sessions matching the specified username
              if ($sessions = quser | Where-Object {$_ -match 'Skype'}){
                   ## Parse the session IDs from the output
                   $sessionIds = ($sessions -split ' +')[2]
                   Write-Host "Found $(@($sessionIds).Count) user login(s) on computer."
                   ## Loop through each session ID and pass each to the logoff command
                   $sessionIds | ForEach-Object {
                        Write-Host "Logging off session id [$($_)]..."
                        logoff $_
                     }
                }
          } catch {
              if ($_.Exception.Message -match 'No user exists') {
                  Write-Host "The user is not logged in."
              } else {
               Write-Warning $_.Exception.Message
              }
          }
     }
     invoke-command -ScriptBlock $scriptBlock -ComputerName $Computer -Credential $cred     
}
function resetUserPwd{    
    param (
          [Parameter()]
          [string]$Computer,
          [ValidateNotNull()]
          [System.Management.Automation.Credential()]
          [System.Management.Automation.PSCredential]$cred
     )
     $localUser = 'Admin'
     $localPwd = Read-Host -Prompt "Enter new password for $localUser (leave blank to cancel)" -AsSecureString
     if (!$localPwd -or ($localPwd.Length -eq 0)){
          return $false
     }
     else{
          $localPwd2 = Read-Host -Prompt "Re-enter new password for $localUser (leave blank to cancel)" -AsSecureString
          if (!$localPwd2 -or ($localPwd2.Length -eq 0)){
               return $false
          }
          else{
               $pwd1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($localPwd))
               $pwd2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($localPwd2))
               if ($pwd1 -ne $pwd2){
                    Write-Host "Passwords do not match, cancelling..." -ForegroundColor Yellow
                    return $false
               }
          }
     }
     try{
          Invoke-Command -ComputerName $Computer -ScriptBlock {$UserAccount = Get-LocalUser -Name $using:localUser; $UserAccount | Set-LocalUser -Password $using:localPwd} -Credential $cred
          return $true
     }
     catch {
          Write-Warning $_.Exception.Message
          return $false
     }
}
function checkMTRStatus{
     param (
          [Parameter()]
          [string]$Computer,
          [ValidateNotNull()]
          [System.Management.Automation.Credential()]
          [System.Management.Automation.PSCredential]$cred
     )
     if($cred -ne [System.Management.Automation.PSCredential]::Empty) {
          try{
               #Get System Info
               invoke-command {Write-Host "===== SYSTEM INFO =====" -ForegroundColor Blue;Get-WmiObject -Class Win32_ComputerSystem | Format-List PartOfDomain,Domain,Workgroup,Manufacturer,Model; Get-WmiObject -Class Win32_Bios | Format-List SerialNumber,SMBIOSBIOSVersion} -ComputerName $Computer -Credential $cred

               #Get Attached Devices
               invoke-command {Write-Host "===== VIDEO DEVICES =====" -ForegroundColor Blue;Get-WmiObject -Class Win32_PnPEntity | Where-Object {$_.PNPClass -eq "Image"} | Format-Table Name,Status,Present; Write-Host "===== AUDIO DEVICES =====" -ForegroundColor Blue; Get-WmiObject -Class Win32_PnPEntity | Where-Object {$_.PNPClass -eq "Media"} | Format-Table Name,Status,Present; Write-Host "===== DISPLAY DEVICES =====" -ForegroundColor Blue;Get-WmiObject -Class Win32_PnPEntity | Where-Object {$_.PNPClass -eq "Monitor"} | Format-Table Name,Status,Present} -ComputerName $Computer -credential $cred

               #Get App Status
               invoke-command {Write-Host "===== Teams App Status =====" -ForegroundColor Blue; $package = get-appxpackage -User Skype -Name Microsoft.SkypeRoomSystem; if ($null -eq $package) {Write-host "SkypeRoomSystems not installed."} else {write-host "Teams App version : " $package.Version}; $process = Get-Process -Name "Microsoft.SkypeRoomSystem" -ErrorAction SilentlyContinue; if ($null -eq $process) {write-host "App not running." -ForegroundColor Red} else {$process | format-list StartTime,Responding}} -ComputerName $Computer -Credential $cred

               #Get related scheduled tasks status
               invoke-command {Write-Host "===== Scheduled Tasks Status =====" -ForegroundColor Blue; get-ScheduledTask -TaskPath \Microsoft\Skype\ | format-table TaskName,State} -ComputerName $Computer -Credential $cred               
          }
          catch{
               Write-Warning $_.Exception.Message
          }          
     }
}
function RunDailyMaintenanceTask{
     param (
          [Parameter()]
          [string]$Computer,
          [ValidateNotNull()]
          [System.Management.Automation.Credential()]
          [System.Management.Automation.PSCredential]$cred
     )
     if($cred -ne [System.Management.Automation.PSCredential]::Empty) {
          try{
               #Run nightly maintenance scheduled task
               invoke-command {Start-ScheduledTask -TaskName "NightlyReboot" -TaskPath "\Microsoft\Skype\";Get-ScheduledTask -TaskName "NightlyReboot" | Select-Object TaskName,State} -ComputerName $Computer -Credential $cred
          }
          catch{
               Write-Warning $_.Exception.Message
          }          
     }
}
function rebootMTR{
     param (
          [Parameter()]
          [string]$Computer,
          [ValidateNotNull()]
          [System.Management.Automation.Credential()]
          [System.Management.Automation.PSCredential]$cred
     )
     if($cred -ne [System.Management.Automation.PSCredential]::Empty) {
          try{
               Write-Host "Restarting $computer..." -for Cyan
               invoke-command { Restart-Computer -force } -ComputerName $Computer -Credential $cred     
          }
          catch{
               Write-Warning $_.Exception.Message
          }
     }
}
function retrieveLogs{
     param (
          [Parameter()]
          [string]$Computer,
          [ValidateNotNull()]
          [System.Management.Automation.Credential()]
          [System.Management.Automation.PSCredential]$cred
     )
     if($cred -ne [System.Management.Automation.PSCredential]::Empty) {
          try{
               Write-Host "Collecting `'$Computer`' device logs..." -for Cyan
               $logFile = invoke-command {Powershell.exe -ExecutionPolicy Bypass -File C:\Rigel\x64\Scripts\Provisioning\ScriptLaunch.ps1 CollectSrsV2Logs.ps1; Get-ChildItem -Path C:\Rigel\*.zip | Sort-Object -Descending -Property LastWriteTime | Select-Object -First 1} -ComputerName $Computer -Credential $cred
               <# Debugging
               $logFile = invoke-command {Get-ChildItem -Path C:\Rigel\*.zip | Sort-Object -Descending -Property LastWriteTime | Select-Object -First 1} -ComputerName $Computer -Credential $cred
               #>
               if ($logFile){
                    $logFileName = $logFile.FullName
                    $localfile = $scriptPath+$logFile.Name
                    $MTR_session = new-pssession -ComputerName $Computer -Credential $cred
                    Write-Host "Downloading `'$Computer`' device logs..." -for Cyan
                    Copy-Item -Path $logFile.FullName -Destination $localfile -FromSession $MTR_session
                    Write-Host "Logs available in $localFile..." -for Cyan
                    Remove-PSSession $MTR_session
                    do{
                         $opt = (Read-host "Delete remote file `'$logFileName`' on `'$Computer`'? (y/n)").ToUpper()
                         if ($opt -eq "Y"){                              
                              invoke-command {remove-item -force $Using:logFileName} -ComputerName $Computer -Credential $cred
                              break
                         }
                     }until ("Y","N" -contains $opt)
               }
               else{
                    Write-Host "An unknown error occurred while collecting the files. Please try again." -ForegroundColor Red
               }
          }
          catch{
               Write-Warning $_.Exception.Message
               throw $_.Exception
          }
     }
}
function setTheme{
     param (
          [Parameter()]
          [string]$Computer,
          [ValidateNotNull()]
          [System.Management.Automation.Credential()]
          [System.Management.Automation.PSCredential]$cred
     )
     if($cred -ne [System.Management.Automation.PSCredential]::Empty){
          #Select one of the predefined Themes or set a custom one.
          $themes = @("Default","No Theme","Custom","Blue Wave","Digital Forest","Dreamcatcher","Limeade","Pixel Perfect","Purple Paradise","Roadmap","Sunset")
          $themes | ForEach-Object {"[$PSItem]"}
          do{               
               $themeName = Read-Host "Please enter theme name (leave blank to cancel)"
          }until (($themeName -eq "") -or ($themeName -in $themes))
          if ($themeName -eq ""){
               return
          }
          if ($themeName -eq "Custom"){
               do{
                    [bool]$fileOK = $false
                    Write-Host "Image should be exactly 3840X1080 pixels and must be one of the following file formats: jpg, jpeg, png, and bmp."
                    do{                    
                         $ImgLocalFile = Read-Host "Please enter full path to a valid background image file (leave blank to cancel)"
                    }until (($ImgLocalFile -eq "") -or ($ImgLocalFile -match '^.+\.(jpg|JPG|jpeg|JPEG|png|PNG|bmp|BMP)$'))
                    if ($ImgLocalFile -eq ""){
                         return
                    }
                    else{
                         $fileOK = Test-Path $ImgLocalFile -PathType Leaf
                    }
                    if (!$fileOK){
                         Write-Host "File '$ImgLocalFile' does not exist!" -ForegroundColor Red
                         continue
                    }
                    $ThemeImage = Split-Path $ImgLocalFile -Leaf -Resolve
               }until ($fileOK)
          }          
          $MTRAppPath = "C:\Users\Skype\AppData\Local\Packages\Microsoft.SkypeRoomSystem_8wekyb3d8bbwe\LocalState\"
          #$XmlRemoteFile = $MTRAppPath+"SkypeSettings.xml"
          try{
               $XmlLocalFile = "$PSScriptRoot\SkypeSettings.xml"
               #Create XML file structure
               $xmlfs = '<SkypeSettings>
               <Theming>
                    <ThemeName>$themeName</ThemeName>
                    <CustomThemeImageUrl>$themeImage</CustomThemeImageUrl>
                    <CustomThemeColor>
                         <RedComponent>1</RedComponent>
                         <GreenComponent>1</GreenComponent>
                         <BlueComponent>1</BlueComponent>
                    </CustomThemeColor>
               </Theming>
               </SkypeSettings>
               '
               #Interpret & replace variables values
               $xmlfs = $xmlfs.Replace('$themeName',$themeName)
               $xmlfs = $xmlfs.Replace('$themeImage',$themeImage)
               #Transform string to XML structure
               $xmlFile = [xml]$xmlfs
               #Save base RDC file
               $xmlFile.Save($XmlLocalFile)

               $MTR_session = new-pssession -ComputerName $Computer -Credential $cred
               Copy-Item -Path $XmlLocalFile -Destination $MTRAppPath -ToSession $MTR_session
               remove-item -force $XmlLocalFile
               if ($ImgLocalFile){
                    Copy-Item -Path $ImgLocalFile -Destination $MTRAppPath -ToSession $MTR_session
               }
               Remove-PSSession $MTR_session
               Write-Host "Please RESTART `'$Computer`' to apply new settings!" -for Cyan
          }
          catch{
               Write-Warning $_.Exception.Message
          }
     }
}
function IsValidEmail { 
     param([string]$Email)
     $Regex = '^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$'
 
    try {
         $obj = [mailaddress]$Email
         if($obj.Address -match $Regex){
             return $True
         }
         return $False
     }
     catch {
         return $False
     } 
 }
function setAppUserAccount{
     param (
          [Parameter()]
          [string]$Computer,
          [ValidateNotNull()]
          [System.Management.Automation.Credential()]
          [System.Management.Automation.PSCredential]$cred
     )
     #Disclaimer
     Write-Host "PLEASE NOTE:`nThis will only change the SkypeSignInAddress and associated Password value for the Teams App on the device.`nNo credentials validation are being performed!!`nPlease be sure you enter the correct credentials." -ForegroundColor Yellow
     if($cred -ne [System.Management.Automation.PSCredential]::Empty){
          do{               
               $localUser = Read-Host -Prompt "Please enter ressource account name (i.e. rito@contoso.com; Leave blank to cancel)"
               $isValid = isValidEmail $localUser
          }until ($isValid -or !$localUser -or ($localUser.Length -eq 0))
          if (!$isValid){
               return $false
          }

          $localPwd = Read-Host -Prompt "Enter password for $localUser (leave blank to cancel)" -AsSecureString
          if (!$localPwd -or ($localPwd.Length -eq 0)){
               return $false
          }
          else{
               $localPwd2 = Read-Host -Prompt "Re-enter new password for $localUser (leave blank to cancel)" -AsSecureString
               if (!$localPwd2 -or ($localPwd2.Length -eq 0)){
                    return $false
               }
               else{
                    $pwd1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($localPwd))
                    $pwd2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($localPwd2))
                    if ($pwd1 -ne $pwd2){
                         Write-Host "Passwords do not match, cancelling..." -ForegroundColor Yellow
                         return $false
                    }
               }
          }
          do{
               $yesNo = (Read-Host "Enable MA (Modern Auth)? (y/n)").ToUpper()               
          }until ("Y","N" -contains $yesNo)
          if ("Y" -eq $yesno) {$MAEnabled = "true"}
          else {$MAEnabled = "false"}

          #Select one of the predefined Meeting modes
          Write-Host "Please select option number for supported Meeting Mode:"
          $MeetingsModes = @("(1) Skype for Business (default) and Microsoft Teams","(2) Skype for Business and Microsoft Teams (default)","(3) Skype for Business only","(4) Microsoft Teams only")
          $MeetingsModes | ForEach-Object {write-host $_}
          do{               
               $MeetingsMode = Read-Host "Select option (leave blank to cancel)"
          }until (($MeetingsMode -eq "") -or ($MeetingsMode -in 1,2,3,4 ))
          if ($MeetingsMode -eq ""){
               return
          }
          Write-Host 'Selected Meeting Mode:' $MeetingsModes[$MeetingsMode-1] -ForegroundColor Cyan
          switch ($selection){
               '1'{$TeamsMeetingsEnabled="true";$SfbMeetingEnabled="true";$IsTeamsDefaultClient="false";break}
               '2'{$TeamsMeetingsEnabled="true";$SfbMeetingEnabled="true";$IsTeamsDefaultClient="true";break}
               '3'{$TeamsMeetingsEnabled="false";$SfbMeetingEnabled="true";$IsTeamsDefaultClient="false";break}
               '4'{$TeamsMeetingsEnabled="true";$SfbMeetingEnabled="false";$IsTeamsDefaultClient="true";break}
          }

          $MTRAppPath = "C:\Users\Skype\AppData\Local\Packages\Microsoft.SkypeRoomSystem_8wekyb3d8bbwe\LocalState\"
          #$XmlRemoteFile = $MTRAppPath+"SkypeSettings.xml"
          try{
               $XmlLocalFile = "$PSScriptRoot\SkypeSettings.xml"
               #Create XML file structure
               $xmlfs = '<SkypeSettings>
               <UserAccount>
                    <SkypeSignInAddress>$userName</SkypeSignInAddress>
                    <Password>$userPwd</Password>
                    <ModernAuthEnabled>$MAEnabled</ModernAuthEnabled>
               </UserAccount>
               <TeamsMeetingsEnabled>$TeamsMeetingsEnabled</TeamsMeetingsEnabled>
               <SfbMeetingEnabled>$SfbMeetingEnabled</SfbMeetingEnabled>
               <IsTeamsDefaultClient>$IsTeamsDefaultClient</IsTeamsDefaultClient>
               </SkypeSettings>
               '
               #Interpret & replace variables values
               $xmlfs = $xmlfs.Replace('$userName',$localUser)
               $xmlfs = $xmlfs.Replace('$userPwd',$pwd1)
               $xmlfs = $xmlfs.Replace('$MAEnabled',$MAEnabled)
               $xmlfs = $xmlfs.Replace('$TeamsMeetingsEnabled',$TeamsMeetingsEnabled)
               $xmlfs = $xmlfs.Replace('$SfbMeetingEnabled',$SfbMeetingEnabled)
               $xmlfs = $xmlfs.Replace('$IsTeamsDefaultClient',$IsTeamsDefaultClient)
               #Transform string to XML structure
               $xmlFile = [xml]$xmlfs
               #Save base RDC file
               $xmlFile.Save($XmlLocalFile)

               $MTR_session = new-pssession -ComputerName $Computer -Credential $cred
               Copy-Item -Path $XmlLocalFile -Destination $MTRAppPath -ToSession $MTR_session
               remove-item -force $XmlLocalFile               
               Remove-PSSession $MTR_session
               Write-Host "Please RESTART `'$Computer`' to apply new settings!" -for Cyan
          }
          catch{
               Write-Warning $_.Exception.Message
          }
     }
}

function downloadFile{
     #Base code --> https://gist.github.com/TheBigBear/68510c4e8891f43904d1
     param(
     [Parameter(Mandatory = $true,Position = 0)]
     [string]
     $Url,
     [Parameter(Mandatory = $false,Position = 1)]
     [string]
     [Alias('Folder')]
     $FolderPath
     )
     <# use as
          $url = 'https://go.microsoft.com/fwlink/?linkid=2151817'
          downloadFile $url -FolderPath "C:\temp\MTR"
     #>
     #Find out filename for the download
     try {
         # resolve short URLs
         $req = [System.Net.HttpWebRequest]::Create($Url)
         $req.Method = "HEAD"
         $response = $req.GetResponse()
         $fLength = $response.ContentLength/1MB
         $fUri = $response.ResponseUri
         $filename = [System.IO.Path]::GetFileName($fUri.LocalPath);
         $response.Close()
         # Download file
         $destination = (Get-Item -Path ".\" -Verbose).FullName
         if ($FolderPath) { $destination = $FolderPath }
         if ($destination.EndsWith('\')) {
          $destination += $filename
         } else {
          $destination += '\' + $filename
         }

         if (!(Test-Path -path $destination)) {
               Write-Host "File to be downloaded: $filename ($fLength MB)`nDestination: $destination" -ForegroundColor Yellow
               do{
               $opt = (Read-Host "Proceed? (y/n)").ToUpper()
                    if ($opt -eq "Y"){
                         Start-BitsTransfer -Source $fUri.AbsoluteUri -Destination $destination -Description "DOWNLOADING '$($fUri.AbsoluteUri)'($fLength MB) to `'$destination`'..."
                         Write-Host "File downloaded to `'$destination`'" -ForegroundColor Green 
                         #break                                             
                    }else{
                         return $false
                    }
               }until ("Y","N" -contains $opt)
         }
         else {
              Write-Host "File already exists, no download needed." -ForegroundColor DarkGreen
         }
         # CHECK if downloaded file is = size as estimated
         $locFileSize = (Get-Item $destination).length /1MB
         $remoteFileSize = $response.ContentLength/1MB
          if ($locFileSize -eq $remoteFileSize){
               Write-Host "File size matches!`nFile size: $locFileSize MB`nExpected file size: $remoteFileSize MB`n" -ForegroundColor Green
               return $destination
          }
          else{
               Write-Host "File size does not match!`nFile size: $locFileSize MB`nExpected file size: $remoteFileSize MB`nPlease remove `'$destination`'and retry." -ForegroundColor Red
          }
     }
     catch {
         Write-Host -ForegroundColor DarkRed $_.Exception.Message
     }
     return $false
 }
 function updateMTR{
     param (
          [Parameter()]
          [string]$Computer,
          [ValidateNotNull()]
          [System.Management.Automation.Credential()]
          [System.Management.Automation.PSCredential]$cred
     )
     if($cred -ne [System.Management.Automation.PSCredential]::Empty) {
          Write-Host "This procedure will first download and check the update locally, then transfer it to the MTR device and trigger a manual update." -ForegroundColor Yellow
          do{
               $yesNo = (Read-Host "Do you want to proceed? (y/n)").ToUpper()               
          }until ("Y","N" -contains $yesNo)
          if ("Y" -eq $yesno){
               try{                    
                    $destinationFolder = "C:\Rigel\"

                    if ($scriptFile = downloadFile "https://go.microsoft.com/fwlink/?linkid=2151817" -FolderPath $scriptPath){
                         Unblock-File -Path $scriptFile
                         $fileName = (Get-ChildItem -Path $scriptFile).Name
                         $remoteFileName = $destinationFolder+$fileName
                         $MTR_session = new-pssession -ComputerName $Computer -Credential $cred
                         Write-Host "Copying `'$scriptFile`' to `'$Computer`'..." -for Cyan                         
                         Copy-Item -Path $scriptFile -Destination $remoteFileName -ToSession $MTR_session -Force
                         Write-Host "File copied to `'$destinationFolder`' on `'$Computer`'..." -for Green
                         Write-Host "Applying update `'$fileName`'!! Please restart MTR when finished ;-)" -for Cyan
                         Invoke-command -ScriptBlock {cd $using:destinationFolder;PowerShell.exe -ExecutionPolicy Unrestricted -File $using:remoteFileName} -ComputerName $Computer -Credential $cred 
                         Remove-PSSession $MTR_session
                    }
                    else{
                         Write-Host "Aborting..." -ForegroundColor Red
                    }
               }
               catch{
                    Write-Warning $_.Exception.Message
                    throw $_.Exception
               }
          }          
     }
}
 
function connect2MTR{
     param (
          [string]$Computer,
          [REF]$funcCred
     )
     try{
          $funcCred.Value = Get-Credential -Message "Please enter password for Local Admin user on `'$Computer`'`r`n(Note: MTR factory password is 'sfb'. Please change ASAP!!!)" -user $MTR_AdminUser
          if (Test-WSMan $Computer -Credential $funcCred.Value -Authentication Negotiate -ErrorAction SilentlyContinue){
               Write-Host "$Computer successfully targeted!" -ForegroundColor Green
               return $true
          }
          else{
               Write-Host "$Computer could not get targeted!`r`nPlease confirm MTR is ON, reachable and remote Powershell is enabled (run Enable-PSRemoting locally on the MTR)." -ForegroundColor Red
               return $false
          }
     }
     catch{
          Write-Warning $_.Exception.Message
          return $false
     }
}
#
# INITIALIZE VARIABLES
#
[string]$MTR_hostName = $null
[string]$MTR_AdminUser = "Admin"
[bool]$MTR_ready = $false
#$ProgressPreference = 'SilentlyContinue'
[string]$scriptPath = "$PSScriptRoot\"
[System.Management.Automation.PSCredential] $global:creds = $null
$menuOptions = @(
"`n================ MTR REMOTE MANAGEMENT ================"
"1: Target MTR device."
"2: Change MTR 'Admin' local user password."
"3: Set MTR resource account (Teams App user account)."
"4: Check MTR status."
"5: Get MTR device logs."
"6: Set MTR theme image."
"7: Run nightly maintenance scheduled task."
"8: Update MTR App Version"
"9: Logoff MTR 'Skype' user."
"10: Restart MTR."
"Q: Press 'Q' to quit."
)

#
# MAIN
#
Clear-Host
do{
     Show-Menu
     $selection = (Read-Host "Please make a selection").ToUpper()
     if ('1' -eq $selection){          
          Write-Host 'Option #' $menuOptions[$selection] -ForegroundColor Cyan
          if ($MTR_hostName -eq ""){
               $MTR_hostName = Read-Host -Prompt "Please insert MTR resolvable HOSTNAME"
          }
          else{
               $prompt = Read-Host -Prompt "Please insert MTR resolvable HOSTNAME. Press Enter to keep default value if present [$MTR_hostName]"
               if ($prompt -ne '') {
                    $MTR_hostName = $prompt
               }
          }
          $MTR_hostName = $MTR_hostName.ToUpper()
          if ($MTR_hostName -ne ""){
               $MTR_ready = connect2MTR $MTR_hostName ([REF]$global:creds)
          }
     }
     else{
          if ($selection  -eq 'Q'){
                    exit
          }
          else{
               Write-Host 'Option #' $menuOptions[$selection] -ForegroundColor Cyan
               if ($MTR_ready){
                    switch ($selection){
                         '2'{
                              if (resetUserPwd $MTR_hostName $creds){
                                   $MTR_ready = $false
                                   write-host "Password changed. Please re-target MTR with updated credentials!" -ForegroundColor Yellow
                              }                              
                              break
                         }
                         '3'{setAppUserAccount $MTR_hostName $creds;break}
                         '4'{checkMTRStatus $MTR_hostName $creds;break}
                         '5'{retrieveLogs $MTR_hostName $creds;break}
                         '6'{setTheme $MTR_hostName $creds;break}
                         '7'{RunDailyMaintenanceTask $MTR_hostName $creds;break}
                         '8'{updateMTR $MTR_hostName $creds;break}
                         '9'{remote_logoff $MTR_hostName $creds;break}
                         '10'{rebootMTR $MTR_hostName $creds;$MTR_ready = $false;break}
                    }                    
               }
               else {selectOpt1}
          }
     }
}
until ($selection -eq 'Q')