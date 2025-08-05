param (
    [string]$site = "NB",
    [switch]$force_update 
)

########################################################################################## Util functions ############################################################################################################

function uninstall {
    param ($app_name)

    Write-Host "Searching for $app_name"
    winget uninstall --all --accept-source-agreements $app_name
}

function install {
    param ($app_name)

    # $curr_user = whoami.exe | Split-Path -Leaf

    # if ($curr_user.ToLower() -eq "admin") {
    #     Write-Host "Updating the manifest for $app_name"
    #     wingetcreate update $app_name
    # }
    
    Write-Host "Searching for $app_name"
    winget install -e --accept-package-agreements $app_name
}

function free_bstr {
    param ([System.IntPtr]$bstr)

    if ($bstr -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function compare_secure_string {
    param (
        [System.Security.SecureString]$SecStr1,
        [System.Security.SecureString]$SecStr2
    )

    $bstr1 = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecStr1)
    $bstr2 = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecStr2)


    $len1 = [Runtime.InteropServices.Marshal]::ReadInt32($bstr1, -4)
    $len2 = [Runtime.InteropServices.Marshal]::ReadInt32($bstr2, -4)

    if ($len1 -ne $len2) {
        free_bstr($bstr1)
        free_bstr($bstr2)
        return $false
    }

    for ($i = 0; $i -lt $len1; ++$i) {
        $b1 = [Runtime.InteropServices.Marshal]::ReadByte($bstr1, $i)
        $b2 = [Runtime.InteropServices.Marshal]::ReadByte($bstr2, $i)
        if ($b1 -ne $b2) {
            free_bstr($bstr1)
            free_bstr($bstr2)
            return $false
        }
    }

    free_bstr($bstr1)
    free_bstr($bstr2)
    return $true
}


function get_password {
    param ([string]$user)

    $password1 = Read-Host -AsSecureString -Prompt "Enter password for $user"
    $password2 = Read-Host -AsSecureString -Prompt "confirm password $user"

    $match = compare_secure_string $password1 $password2
    if (!$match) {
        Write-Host "Passwords do not match! Try again." -BackgroundColor "red"
        return get_password($user)
    }
    return $password1
}

function add_to_task_bar {
    param (
        $exe,
        $lnk
    )

    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($lnk)
    $Shortcut.TargetPath = $exe
    $Shortcut.Save()

    $Shell = New-Object -Verbose -ComObject Shell.Application
    $Shell.NameSpace((Split-Path $lnk)).ParseName((Split-Path $lnk -Leaf)).InvokeVerb("PinToTaskbar")
}

########################################################################################## Start Script ############################################################################################################

$status = Get-AppPackage *Microsoft.WindowsStore* | Select-Object status
if (!($status.status -eq 'Ok')) {
    Write-Host "Updating Microsoft store" -ForegroundColor "yellow" -BackgroundColor "blue"
    Write-Host
    Get-AppxPackage *WindowsStore* -AllUsers | ForEach-Object {Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\\AppXManifest.xml"}
}

$min_winget = 1.11
If ((Get-Command -Name "winget" -ErrorAction SilentlyContinue)) {
    $winget_v = (winget -v).SubString(1)
}
else {
    $winget_v = 0.0
}

If (($winget_v -lt $min_winget) -or $force_update) {
    Write-Host "Installing winget" -ForegroundColor "yellow" -BackgroundColor "blue"    
    Write-Host
    # Install winget
    Invoke-WebRequest -Uri https://aka.ms/getwinget -OutFile winget.msixbundle
    Add-AppxPackage winget.msixbundle
    Remove-Item winget.msixbundle
    Install-Module -Name Microsoft.WinGet.Client
}

Write-Host
Write-Host
Write-Host "Removing Uneeded Apps" -ForegroundColor "yellow" -BackgroundColor "blue"
Write-Host

uninstall("Microsoft Copilot")
uninstall("Microsoft 365 Copilot")
uninstall("Microsoft Edge Browser")
uninstall("Microsoft Edge DevTools Preview")
uninstall("Microsoft Beta")
uninstall("Microsoft Canary")
uninstall("Microsoft Dev")
uninstall("Microsoft WebDriver")
uninstall("Microsoft WebView2 Runtime")

Write-Host
Write-Host
Write-Host "Installing Packages" -ForegroundColor "yellow" -BackgroundColor "blue"
Write-Host

winget install wingetcreate

install("Google.Chrome")

Write-Host
Write-Host
Write-Host "Creating Shortcuts" -ForegroundColor "yellow" -BackgroundColor "blue"
Write-Host
$chrome_lnk = "Google Chrome.lnk"
add_to_task_bar "C:\Program Files\Google\Chrome\Application\chrome.exe"  "$Home\Desktop\$chrome_lnk"

Write-Host
Write-Host
Write-Host "Removing Shortcuts" -ForegroundColor "yellow" -BackgroundColor "blue"
Write-Host
$msedge_ln = "Microsoft Edge.lnk"
Remove-Item -ErrorAction SilentlyContinue "$Home\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\$msedge_ln"
Remove-Item -ErrorAction SilentlyContinue "$Home\Desktop\$msedge_ln"
Remove-Item -ErrorAction SilentlyContinue "C:\Users\Public\Desktop\$msedge_ln"
Remove-Item -ErrorAction SilentlyContinue "$Home\Desktop\$chrome_lnk"


$curr_user = whoami.exe | Split-Path -Leaf
if ($curr_user.ToLower() -eq "admin"){

    Write-Host
    Write-Host
    Write-Host "Changing timezone" -ForegroundColor "yellow" -BackgroundColor "blue"

    $timezone_pst_id = "Pacific Standard Time"
    $timezone_pst = [System.TimeZOneInfo]::FindSystemTimeZoneById($timezone_pst_id)
    $timezone_curr = [System.TimeZoneInfo]::Local
    
    $pst_offset = $timezone_pst.GetUtcOffset([System.DateTimeOffset]::UtcNow)
    $curr_offset = $timezone_curr.GetUtcOffset([System.DateTimeOffset]::UtcNow)
    $offset_curr_pst = $pst_offset - $curr_offset

    Set-TimeZone $timezone_pst_id
    Set-Date -Adjust $offset_curr_pst

    Remove-LocalUser -ErrorAction SilentlyContinue -Name "User"
    Remove-LocalUser -ErrorAction SilentlyContinue -Name "user"

    $new_pass_curr_user = get_password($curr_user)
    Set-LocalUser -Name $curr_user -Password $new_pass_curr_user

    $user = "user"
    Write-Host
    Write-Host
    Write-Host "Adding new default $user" -ForegroundColor "yellow" -BackgroundColor "blue"

    $password = get_password($user)
    New-LocalUser -Name $user -Password $password -FullName $user
    Add-LocalGroupMember -Group "Users" -Member $user

    .\printers.ps1 $site

    $assetnbr = Read-Host -Prompt "Enter the pivot asset tag number"
    $pcname = "pivot_$assetnbr"

    Rename-Computer -NewName $pcname -Restart
}

Write-Host

Get-WmiObject win32_bios
Get-WmiObject win32_computersystem

Read-Host

shutdown.exe /s /t 10