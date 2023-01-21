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
          Write-Host "[Target MTR: $MTR_hostName (v$MTR_version)]" -ForegroundColor Green
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
               if ($pwd1 -cne $pwd2){
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
               invoke-command {Write-Host "===== Teams App Status =====" -ForegroundColor Blue; $package = get-appxpackage -User Skype -Name Microsoft.SkypeRoomSystem; if ($null -eq $package) {Write-host "SkypeRoomSystems not installed."} else {write-host "Teams App version: " $package.Version;write-host "Teams App language: " (Get-WinUserLanguageList).LanguageTag}; $process = Get-Process -Name "Microsoft.SkypeRoomSystem" -ErrorAction SilentlyContinue; if ($null -eq $process) {write-host "App not running." -ForegroundColor Red} else {$process | format-list StartTime,Responding}} -ComputerName $Computer -Credential $cred

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
function setMTRLanguage{     
     param (
          [Parameter()]
          [string]$Computer,
          [ValidateNotNull()]
          [System.Management.Automation.Credential()]
          [System.Management.Automation.PSCredential]$cred
     )
     if($cred -ne [System.Management.Automation.PSCredential]::Empty) {
          #Get-Culture -listAvailable is only available for PowerShell 6.2+, so keeping the possible values list for retro-compatibility.
          $locales = @("aa","aa-DJ","aa-ER","aa-ET","af","af-NA","af-ZA","agq","agq-CM","ak","ak-GH","am","am-ET","ar","ar-001","ar-AE","ar-BH","ar-DJ","ar-DZ","ar-EG","ar-ER","ar-IL","ar-IQ","ar-JO","ar-KM","ar-KW","ar-LB","ar-LY","ar-MA","ar-MR","ar-OM","ar-PS","ar-QA","ar-SA","ar-SD","ar-SO","ar-SS","ar-SY","ar-TD","ar-TN","ar-YE","arn","arn-CL","as","as-IN","asa","asa-TZ","ast","ast-ES","az","az-Cyrl","az-Cyrl-AZ","az-Latn","az-Latn-AZ","ba","ba-RU","bas","bas-CM","be","be-BY","bem","bem-ZM","bez","bez-TZ","bg","bg-BG","bin","bin-NG","bm","bm-ML","bn","bn-BD","bn-IN","bo","bo-CN","bo-IN","br","br-FR","brx","brx-IN","bs","bs-Cyrl","bs-Cyrl-BA","bs-Latn","bs-Latn-BA","byn","byn-ER","ca","ca-AD","ca-ES","ca-FR","ca-IT","ccp","ccp-BD","ccp-IN","ce","ce-RU","ceb","ceb-PH","cgg","cgg-UG","chr","chr-US","ckb","ckb-IQ","ckb-IR","co","co-FR","cs","cs-CZ","cu","cu-RU","cy","cy-GB","da","da-DK","da-GL","dav","dav-KE","de","de-AT","de-BE","de-CH","de-DE","de-IT","de-LI","de-LU","dje","dje-NE","doi","doi-IN","dsb","dsb-DE","dua","dua-CM","dv","dv-MV","dyo","dyo-SN","dz","dz-BT","ebu","ebu-KE","ee","ee-GH","ee-TG","el","el-CY","el-GR","en","en-001","en-029","en-150","en-AE","en-AG","en-AI","en-AS","en-AT","en-AU","en-BB","en-BE","en-BI","en-BM","en-BS","en-BW","en-BZ","en-CA","en-CC","en-CH","en-CK","en-CM","en-CX","en-CY","en-DE","en-DK","en-DM","en-ER","en-FI","en-FJ","en-FK","en-FM","en-GB","en-GD","en-GG","en-GH","en-GI","en-GM","en-GU","en-GY","en-HK","en-ID","en-IE","en-IL","en-IM","en-IN","en-IO","en-JE","en-JM","en-KE","en-KI","en-KN","en-KY","en-LC","en-LR","en-LS","en-MG","en-MH","en-MO","en-MP","en-MS","en-MT","en-MU","en-MW","en-MY","en-NA","en-NF","en-NG","en-NL","en-NR","en-NU","en-NZ","en-PG","en-PH","en-PK","en-PN","en-PR","en-PW","en-RW","en-SB","en-SC","en-SD","en-SE","en-SG","en-SH","en-SI","en-SL","en-SS","en-SX","en-SZ","en-TC","en-TK","en-TO","en-TT","en-TV","en-TZ","en-UG","en-UM","en-US","en-US-POSIX","en-VC","en-VG","en-VI","en-VU","en-WS","en-ZA","en-ZM","en-ZW","eo","eo-001","es","es-419","es-AR","es-BO","es-BR","es-BZ","es-CL","es-CO","es-CR","es-CU","es-DO","es-EC","es-ES","es-GQ","es-GT","es-HN","es-MX","es-NI","es-PA","es-PE","es-PH","es-PR","es-PY","es-SV","es-US","es-UY","es-VE","et","et-EE","eu","eu-ES","ewo","ewo-CM","fa","fa-AF","fa-IR","ff","ff-Adlm","ff-Adlm-BF","ff-Adlm-CM","ff-Adlm-GH","ff-Adlm-GM","ff-Adlm-GN","ff-Adlm-GW","ff-Adlm-LR","ff-Adlm-MR","ff-Adlm-NE","ff-Adlm-NG","ff-Adlm-SL","ff-Adlm-SN","ff-Latn","ff-Latn-BF","ff-Latn-CM","ff-Latn-GH","ff-Latn-GM","ff-Latn-GN","ff-Latn-GW","ff-Latn-LR","ff-Latn-MR","ff-Latn-NE","ff-Latn-NG","ff-Latn-SL","ff-Latn-SN","fi","fi-FI","fil","fil-PH","fo","fo-DK","fo-FO","fr","fr-029","fr-BE","fr-BF","fr-BI","fr-BJ","fr-BL","fr-CA","fr-CD","fr-CF","fr-CG","fr-CH","fr-CI","fr-CM","fr-DJ","fr-DZ","fr-FR","fr-GA","fr-GF","fr-GN","fr-GP","fr-GQ","fr-HT","fr-KM","fr-LU","fr-MA","fr-MC","fr-MF","fr-MG","fr-ML","fr-MQ","fr-MR","fr-MU","fr-NC","fr-NE","fr-PF","fr-PM","fr-RE","fr-RW","fr-SC","fr-SN","fr-SY","fr-TD","fr-TG","fr-TN","fr-VU","fr-WF","fr-YT","fur","fur-IT","fy","fy-NL","ga","ga-GB","ga-IE","gd","gd-GB","gl","gl-ES","gn","gn-PY","gsw","gsw-CH","gsw-FR","gsw-LI","gu","gu-IN","guz","guz-KE","gv","gv-IM","ha","ha-GH","ha-NE","ha-NG","haw","haw-US","he","he-IL","hi","hi-IN","hr","hr-BA","hr-HR","hsb","hsb-DE","hu","hu-HU","hy","hy-AM","ia","ia-001","ibb","ibb-NG","id","id-ID","ig","ig-NG","ii","ii-CN","is","is-IS","it","it-CH","it-IT","it-SM","it-VA","iu","iu-CA","iu-Latn","iu-Latn-CA","ja","ja-JP","jgo","jgo-CM","jmc","jmc-TZ","jv","jv-ID","jv-Java","jv-Java-ID","ka","ka-GE","kab","kab-DZ","kam","kam-KE","kde","kde-TZ","kea","kea-CV","khq","khq-ML","ki","ki-KE","kk","kk-KZ","kkj","kkj-CM","kl","kl-GL","kln","kln-KE","km","km-KH","kn","kn-IN","ko","ko-KP","ko-KR","kok","kok-IN","kr","kr-Latn","kr-Latn-NG","ks","ks-Arab","ks-Arab-IN","ks-Deva","ks-Deva-IN","ksb","ksb-TZ","ksf","ksf-CM","ksh","ksh-DE","kw","kw-GB","ky","ky-KG","la","la-VA","lag","lag-TZ","lb","lb-LU","lg","lg-UG","lkt","lkt-US","ln","ln-AO","ln-CD","ln-CF","ln-CG","lo","lo-LA","lrc","lrc-IQ","lrc-IR","lt","lt-LT","lu","lu-CD","luo","luo-KE","luy","luy-KE","lv","lv-LV","mai","mai-IN","mas","mas-KE","mas-TZ","mer","mer-KE","mfe","mfe-MU","mg","mg-MG","mgh","mgh-MZ","mgo","mgo-CM","mi","mi-NZ","mk","mk-MK","ml","ml-IN","mn","mn-MN","mn-Mong","mn-Mong-CN","mn-Mong-MN","mni","mni-Beng","mni-Beng-IN","moh","moh-CA","mr","mr-IN","ms","ms-BN","ms-ID","ms-MY","ms-SG","mt","mt-MT","mua","mua-CM","my","my-MM","mzn","mzn-IR","naq","naq-NA","nb","nb-NO","nb-SJ","nd","nd-ZW","nds","nds-DE","nds-NL","ne","ne-IN","ne-NP","nl","nl-AW","nl-BE","nl-BQ","nl-CW","nl-NL","nl-SR","nl-SX","nmg","nmg-CM","nn","nn-NO","nnh","nnh-CM","nqo","nqo-GN","nr","nr-ZA","nso","nso-ZA","nus","nus-SS","nyn","nyn-UG","oc","oc-FR","om","om-ET","om-KE","or","or-IN","os","os-GE","os-RU","pa","pa-Arab","pa-Arab-PK","pa-Guru","pa-Guru-IN","pap","pap-029","pcm","pcm-NG","pl","pl-PL","prg","prg-001","ps","ps-AF","ps-PK","pt","pt-AO","pt-BR","pt-CH","pt-CV","pt-GQ","pt-GW","pt-LU","pt-MO","pt-MZ","pt-PT","pt-ST","pt-TL","qu","qu-BO","qu-EC","qu-PE","quc","quc-GT","rm","rm-CH","rn","rn-BI","ro","ro-MD","ro-RO","rof","rof-TZ","ru","ru-BY","ru-KG","ru-KZ","ru-MD","ru-RU","ru-UA","rw","rw-RW","rwk","rwk-TZ","sa","sa-IN","sah","sah-RU","saq","saq-KE","sat","sat-Olck","sat-Olck-IN","sbp","sbp-TZ","sd","sd-Arab","sd-Arab-PK","sd-Deva","sd-Deva-IN","se","se-FI","se-NO","se-SE","seh","seh-MZ","ses","ses-ML","sg","sg-CF","shi","shi-Latn","shi-Latn-MA","shi-Tfng","shi-Tfng-MA","si","si-LK","sk","sk-SK","sl","sl-SI","sma","sma-NO","sma-SE","smj","smj-NO","smj-SE","smn","smn-FI","sms","sms-FI","sn","sn-ZW","so","so-DJ","so-ET","so-KE","so-SO","sq","sq-AL","sq-MK","sq-XK","sr","sr-Cyrl","sr-Cyrl-BA","sr-Cyrl-ME","sr-Cyrl-RS","sr-Cyrl-XK","sr-Latn","sr-Latn-BA","sr-Latn-ME","sr-Latn-RS","sr-Latn-XK","ss","ss-SZ","ss-ZA","ssy","ssy-ER","st","st-LS","st-ZA","su","su-Latn","su-Latn-ID","sv","sv-AX","sv-FI","sv-SE","sw","sw-CD","sw-KE","sw-TZ","sw-UG","syr","syr-SY","ta","ta-IN","ta-LK","ta-MY","ta-SG","te","te-IN","teo","teo-KE","teo-UG","tg","tg-TJ","th","th-TH","ti","ti-ER","ti-ET","tig","tig-ER","tk","tk-TM","tn","tn-BW","tn-ZA","to","to-TO","tr","tr-CY","tr-TR","ts","ts-ZA","tt","tt-RU","twq","twq-NE","tzm","tzm-Arab","tzm-Arab-MA","tzm-DZ","tzm-MA","tzm-Tfng","tzm-Tfng-MA","ug","ug-CN","uk","uk-UA","ur","ur-IN","ur-PK","uz","uz-Arab","uz-Arab-AF","uz-Cyrl","uz-Cyrl-UZ","uz-Latn","uz-Latn-UZ","vai","vai-Latn","vai-Latn-LR","vai-Vaii","vai-Vaii-LR","ve","ve-ZA","vi","vi-VN","vo","vo-001","vun","vun-TZ","wae","wae-CH","wal","wal-ET","wo","wo-SN","xh","xh-ZA","xog","xog-UG","yav","yav-CM","yi","yi-001","yo","yo-BJ","yo-NG","zgh","zgh-MA","zh","zh-Hans","zh-Hans-CN","zh-Hans-HK","zh-Hans-MO","zh-Hans-SG","zh-Hant","zh-Hant-HK","zh-Hant-MO","zh-Hant-TW","zu","zu-ZA")
          do{
               $locale = Read-Host "Please enter a valid language (i.e. en-US, es-ES, de-DE, ...) or leave blank to cancel. Enter ? to list possible values"
               if ($locale -eq "?"){
                    Write-host $locales -for Magenta
               }
          }until (($locale -eq "") -or ($locale -in $locales))
          if ($locale -eq ""){
               return
          }
          $scriptBlock = {
               Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process
               Set-WinUserLanguageList $using:locale -Force
               Set-WinSystemLocale $using:locale
               C:\Rigel\x64\Scripts\Provisioning\ScriptLaunch.ps1 Applycurrentregionandlanguage.ps1
               }
          try{
               invoke-command -ScriptBlock $scriptBlock -ComputerName $Computer -Credential $cred
               Write-Host "Please RESTART `'$Computer`' to apply new settings!" -for Cyan
          }
          catch{
               Write-Warning $_.Exception.Message
               return
          }
     }
}
function get-MTR_version{     
     param (
          [Parameter()]
          [string]$Computer,
          [ValidateNotNull()]
          [System.Management.Automation.Credential()]
          [System.Management.Automation.PSCredential]$cred
     )
     if($cred -ne [System.Management.Automation.PSCredential]::Empty) {
          try{
               $ver = invoke-command { (get-appxpackage -User Skype -Name Microsoft.SkypeRoomSystem).Version } -ComputerName $Computer -Credential $cred
          }
          catch{
               Write-Warning $_.Exception.Message
          }
          return $ver
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
          if ($MTR_version.StartsWith("4.15.")){
               $themes = @("Default","Custom","No Theme","Blue Wave","Creative Conservatory","Digital Forest","Dreamcatcher","Into The Fold","Limeade","Pixel Perfect","Purple Paradise","Roadmap","Seaside Bliss","Summer Summit","Sunset","Vivid Flag Default")
          }
          else{
               $themes = @("Default","No Theme","Custom","Blue Wave","Digital Forest","Dreamcatcher","Limeade","Pixel Perfect","Purple Paradise","Roadmap","Sunset")
          }          
          $themes | ForEach-Object {write-host "[$PSItem]" -ForegroundColor Magenta}
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
                    if ($pwd1 -cne $pwd2){
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
     #Reference code --> https://gist.github.com/TheBigBear/68510c4e8891f43904d1
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
                         Invoke-command -ScriptBlock {Set-Location $using:destinationFolder;PowerShell.exe -ExecutionPolicy Unrestricted -File $using:remoteFileName} -ComputerName $Computer -Credential $cred 
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
[string]$MTR_version = "0.0.0.0"
#$ProgressPreference = 'SilentlyContinue'
[string]$scriptPath = "$PSScriptRoot\"
[System.Management.Automation.PSCredential] $global:creds = $null
$menuOptions = @(
"`n================ MTR REMOTE MANAGEMENT ================"
"1: TARGET MTR device."
"2: Change MTR 'Admin' local user PASSWORD."
"3: Set MTR resource ACCOUNT (Teams App user account)."
"4: Set MTR Teams App LANGUAGE."
"5: Check MTR STATUS."
"6: Get MTR device LOGS."
"7: Set MTR THEME image."
"8: Run nightly MAINTENANCE scheduled task."
"9: Update MTR App VERSION"
"10: LOGOFF MTR 'Skype' user."
"11: RESTART MTR."
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
               $MTR_version = get-MTR_version $MTR_hostName $creds
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
                         '4'{setMTRLanguage $MTR_hostName $creds;break}
                         '5'{checkMTRStatus $MTR_hostName $creds;break}
                         '6'{retrieveLogs $MTR_hostName $creds;break}
                         '7'{setTheme $MTR_hostName $creds;break}
                         '8'{RunDailyMaintenanceTask $MTR_hostName $creds;break}
                         '9'{updateMTR $MTR_hostName $creds;$MTR_version = get-MTR_version $MTR_hostName $creds;break}
                         '10'{remote_logoff $MTR_hostName $creds;break}
                         '11'{rebootMTR $MTR_hostName $creds;$MTR_ready = $false;break}
                    }                    
               }
               else {selectOpt1}
          }
     }
}
until ($selection -eq 'Q')