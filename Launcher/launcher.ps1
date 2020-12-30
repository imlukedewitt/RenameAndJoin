# Launcher script to open Rename And Join on a computer that hasn't already joined the domain.
# Continues to work after the computer has joined the domain, to avoid confusion.
# Version: 1.0

function RunAsAdmin
{
    # Checks if the script is running as admin. If it isn't, it relaunches the script as admin and exits the non-admin script
    # Passes any current parameters on to the elevated script (credentials)
    
    $adminaccess = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')
    if ($adminaccess) { return $true }
    $argList = " -ExecutionPolicy Bypass -NoLogo -NoProfile -WindowStyle Minimized -File $(Get-UNCPath -Quotes)"
    Start-Process Powershell.exe -Verb Runas -ArgumentList $argList 
    exit
}

function Get-UNCPath
{
    # Powershell can't use named network paths; it needs the full UNC path. This function checks the path of the running script and replaces it with a UNC path if necessary
    # This makes the script portable

    param ( [switch]$ScriptRoot, [switch]$Quotes )

    $path = $script:MyInvocation.MyCommand.Path
    if ($path.contains([io.path]::VolumeSeparatorChar))
    {
        $psDrive = Get-PSDrive -Name $path.Substring(0,1) -PSProvider FileSystem
        if ($psDrive.DisplayRoot) { $path = $path.Replace($psDrive.Name + [io.path]::VolumeSeparatorChar, $psDrive.DisplayRoot) }
    }

    if ($ScriptRoot) { $path = $path.Replace([io.path]::DirectorySeparatorChar + $script:MyInvocation.MyCommand.Name, "") }
    if ($Quotes) { $path = '"' + $path + '"' }
    return $path
}

if (!(RunAsAdmin)) { exit }

# Get credentials. They're converted to plaintext for self-elevation and ADSI operations
$cred = Get-Credential -Message "Please enter your employee username/password `n`r`n`r`n`r(Southern\Username)"
if (!$cred) { Write-Host "No username/Password entered, exiting"; pause; exit }
if ($cred.UserName -notlike "SOUTHERN\*") { $cred = [pscredential]::new("SOUTHERN\$($cred.UserName)",$cred.Password) }
$username = $cred.UserName
$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.Password))

# Map the J Drive if it isn't already present
$jDrive = Get-PSDrive -PSProvider FileSystem | Where-Object {$_.DisplayRoot -like "\\redacted"}
if (!$jDrive)
{
    Write-Host "mapping \\redacted"
    try {New-PSDrive -Name "RenameAndJoin" -PSProvider "FileSystem" -Root "\\redacted" -Credential $cred -Erroraction stop}
    catch {write-host "Unable to map network drive! Error message:`n`r`n`r$_"; pause; return}
}
else {Write-Host "\\redacted already mapped, proceeding`n"}

Write-Host "run script`n"
# Run Rename And Join with the credential parameters
. "\\redacted\Rename and Join\RenameAndJoin.ps1" -username "$username" -password "$password"