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
     Write-Host "Please select Option 1 first to set MTR device and credentials for remote connection." -ForegroundColor Magenta
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
               invoke-command {Write-Host "SYSTEM INFO:";Get-WmiObject -Class Win32_ComputerSystem | Format-List PartOfDomain,Domain,Workgroup,Manufacturer,Model; Get-WmiObject -Class Win32_Bios | Format-List SerialNumber,SMBIOSBIOSVersion} -ComputerName $Computer -Credential $cred

               #Get Attached Devices
               invoke-command {Write-Host "VIDEO DEVICES:";Get-WmiObject -Class Win32_PnPEntity | Where-Object {$_.PNPClass -eq "Image"} | Format-Table Name,Status,Present; Write-Host "AUDIO DEVICES:"; Get-WmiObject -Class Win32_PnPEntity | Where-Object {$_.PNPClass -eq "Media"} | Format-Table Name,Status,Present; Write-Host "DISPLAY DEVICES:";Get-WmiObject -Class Win32_PnPEntity | Where-Object {$_.PNPClass -eq "Monitor"} | Format-Table Name,Status,Present} -ComputerName $Computer -credential $cred

               #Get App Status
               invoke-command { $package = get-appxpackage -User Skype -Name Microsoft.SkypeRoomSystem; if ($null -eq $package) {Write-host "SkypeRoomSystems not installed."} else {write-host "SkypeRoomSystem Version : " $package.Version}; $process = Get-Process -Name "Microsoft.SkypeRoomSystem" -ErrorAction SilentlyContinue; if ($null -eq $process) {write-host "App not running."} else {$process | format-list StartTime,Responding}} -ComputerName $Computer -Credential $cred
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
               #$logFile = invoke-command {Get-ChildItem -Path C:\Rigel\*.zip | Sort-Object -Descending -Property LastWriteTime | Select-Object -First 1} -ComputerName $Computer -Credential $cred
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
function setAppUserAccount{
     param (
          [Parameter()]
          [string]$Computer,
          [ValidateNotNull()]
          [System.Management.Automation.Credential()]
          [System.Management.Automation.PSCredential]$cred
     )
     if($cred -ne [System.Management.Automation.PSCredential]::Empty){
          $emailRegEx = '/^\S+@\S+\.\S+$/'
          do{
               $localUser = Read-Host -Prompt "Please enter ressource account name (i.e. rito@contoso.com; Leave blank to cancel)"
               if (!$localUser -or ($localUser.Length -eq 0)){
                    return $false
               }
          }until ($localUser -match $emailRegEx)
          

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
                
          $MTRAppPath = "C:\Users\Skype\AppData\Local\Packages\Microsoft.SkypeRoomSystem_8wekyb3d8bbwe\LocalState\"
          #$XmlRemoteFile = $MTRAppPath+"SkypeSettings.xml"
          try{
               $XmlLocalFile = "$PSScriptRoot\SkypeSettings.xml"
               #Create XML file structure
               $xmlfs = '<SkypeSettings>
               <UserAccount>
                    <SkypeSignInAddress>$userName</SkypeSignInAddress>
                    <ExchangeAddress>$userName</ExchangeAddress>
                    <DomainUsername>domain\username</DomainUsername>
                    <Password>$userPwd</Password>
                    <ConfigureDomain>domain1, domain2</ConfigureDomain>
                    <ModernAuthEnabled>$MAEnabled</ModernAuthEnabled>
               </UserAccount>
               </SkypeSettings>
               '
               #Interpret & replace variables values
               $xmlfs = $xmlfs.Replace('$userName',$userName)
               $xmlfs = $xmlfs.Replace('$userPwd',$userPwd)
               $xmlfs = $xmlfs.Replace('$MAEnabled',$MAEnabled)
               $xmlfs = $xmlfs.Replace('$MeetingsMode',$MeetingsMode)
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
function connect2MTR{
     param (
          [string]$Computer,
          [REF]$funcCred
     )
     try{
          $funcCred.Value = Get-Credential -Message "Please enter password for user `'$MTR_AdminUser`' on `'$Computer`'`r`n(Note: MTR factory password is 'sfb')" -user $MTR_AdminUser
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
$ProgressPreference = 'SilentlyContinue'
[string]$scriptPath = "$PSScriptRoot\"
[System.Management.Automation.PSCredential] $global:creds = $null
$menuOptions = @(
"`n================ MTR REMOTE MANAGEMENT ================"
"1: Target MTR device."
"2: Change MTR 'Admin' local user password."
"3: Check MTR status."
"4: Get MTR device logs."
"5: Set MTR theme image."
"6: Logoff MTR 'Skype' user."
"7: Restart MTR."
"Q: Press 'Q' to quit."
)
# 8: Run Agent Test Tool? Guess not
# 9: Set MTR ressource account credentials

#
# MAIN
#
Clear-Host
do{
     Show-Menu
     $selection = (Read-Host "Please make a selection").ToUpper()
     switch ($selection){
          '1' {
               Write-Host 'Option #' $menuOptions[$selection] -ForegroundColor Cyan
               if ($MTR_hostName -eq ""){
                    $MTR_hostName = Read-Host -Prompt "Please insert MTR resolvable HOSTNAME"
               }
               else{
                    $prompt = Read-Host -Prompt "Please insert MTR resolvable HOSTNAME. Press Enter to keep default value [$MTR_hostName]"
                    if ($prompt -ne '') {
                         $MTR_hostName = $prompt
                    }
               }
               $MTR_hostName = $MTR_hostName.ToUpper()
               if ($MTR_hostName -ne ""){
                    $MTR_ready = connect2MTR $MTR_hostName ([REF]$global:creds)
               }
               break
          }
          {'2','3','4','5','6','7' -contains $_} {
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
                         '3'{checkMTRStatus $MTR_hostName $creds;break}
                         '4'{retrieveLogs $MTR_hostName $creds;break}
                         '5'{setTheme $MTR_hostName $creds;break}
                         '6'{remote_logoff $MTR_hostName $creds;break}
                         '7'{rebootMTR $MTR_hostName $creds;$MTR_ready = $false;break}
                    }                    
               }
               else {                    
                    selectOpt1
               }
               break
          }
          'Q' {
               exit
          }           
     }
}
until ($input -eq 'q')