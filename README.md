# MTR_REMOTE_MGMT.ps1
This script is intended to serve as a backup management tool to address basic and initial tasks on a Microsoft Teams Room system based on Windows (MTRoW).

Please follow the recommended best practices by Microsoft to setup, manage and maintain this kind of devices:
https://docs.microsoft.com/en-us/microsoftteams/rooms/

This script can be helpful in the following scenarios:
* Initial setup and configuration tasks
* Unable to access or manage the device via Microsoft Teams Admin Center (TAC) or Microsoft Endpoint Manager Admin Center / Intune.
* Manage standalone devices

## Table of Contents

- [MTR_REMOTE_MGMT](#MTR_REMOTE_MGMT.ps1)
  - [Table of Contents](#table-of-contents)
  - [DISCLAIMER](#DISCLAIMER)
  - [Getting Started](#getting-started)
    - [Prerequisites](#prerequisites)    
  - [Usage](#usage)
  - [FAQ](#FAQ)
  - [License](#license)

## DISCLAIMER
This script requires to enable and allow remote access to the device, which may violate the security guidelines for devices in your environment. Only proceed if you know what you are doing, and feel comfortable to do so.

The script is provided AS IS without warranty of any kind. I do not take any responsability on it's use.

The entire risk arising out of the use or performance of the script and documentation remains with you. In no event shall I or anyone else involved in the creation, development or delivery of the script be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the scripts or documentation.

## Getting started

Meet the basic prerequisites and you should be ready to go!

### Prerequisites

* PowerShell (PS) Version 5.1+
* Remote Powershell enabled on remote MTR
* MTR device resolvable and accesible over the network

**On local management machine:**
* Run `$PSVersionTable` and check if PSVersion matches required version
* Define the trusted remote endpoint(s) where you will connect to:

    `Set-Item WSMan:\localhost\Client\TrustedHosts -Value "<MTR_IP>|<MTR_Name>" [-Force]`

**Locally, on remote MTR device:**
* Open an admin elevated Powershell session and run:

    `Enable-PSRemoting`

  **Optional:** Use the _-SkipNetworkProfileCheck_ parameter if network is not trusted (public) and still want to force PsRemoting, or directly define your network as private by executing:

    `Set-NetConnectionProfile -Name "<Network_Name>" -NetworkCategory Private`

* Maybe you also want or need to override the Local Security Policy to allow to connect to the MTR remotely from the network:

  Open the Local Security Policy by running _Secpol.msc_ and add the Administrators security group to Security Settings -> Local Policies -> User Rights Assignment -> Access this computer from the network.

From here on, you can enter a Powershell session on the remote device (1), or directly run PS commands (2) from your management server/workstation with the appropriate credentials:

    $cred = Get-Credential
    $MTR_device = "<MTR_FQDN_or_IP_Address>"

(1): 
    `Enter-PSSession -ComputerName $MTR_device -Credential $cred`

(2):
    `Invoke-command { <PS_scriptblock> } -ComputerName $MTR_device -Credential $cred`

## Usage

Invoke the script, select an option and follow the instructions:

**EXAMPLE:**
  `_.\MTR_REMOTE_MGMT.ps1_`

**MAIN MENU:**


    (No target MTR!)

    ================ MTR REMOTE MANAGEMENT ================
    1: Target MTR device.
    2: Change MTR 'Admin' local user password.
    3: Check MTR status.
    4: Get MTR device logs.
    5: Set MTR theme image.
    6: Logoff MTR 'Skype' user.
    7: Restart MTR.
    Q: Press 'Q' to quit.
    Please make a selection:


## FAQ

* Question: Can I use this tool to operate in batches on multiple MTR devices at the same time?
  Answer: No, this script is intended as a backup tool to operate on single MTR devices as of now.

## License

Distributed under the MIT License. See `LICENSE` for more information.
