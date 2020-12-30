##: A GUI for renaming PCs, joining them to Active Directory, and moving them to the correct organizational unit
##: Made by Luke

# Form element syntax:
# l = label
# c = combobox
# t = textbox
# x = checkbox
# r = radio button
# b = button

# Optionally accept credentials to be passed from launcher script.
param
(
    [string]$username,
    [string]$password
)

function RunAsAdmin
{
    # Checks if the script is running as admin. If it isn't, it relaunches the script as admin and exits the non-admin script
    # Passes any current parameters on to the elevated script (credentials)
    
    $adminaccess = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')
    if ($adminaccess) { return $true }
    $argList = " -ExecutionPolicy Bypass -WindowStyle Minimized -NoProfile -NoLogo -File $(Get-UNCPath -Quotes)"
    # A better way to do this is passing all of the keys/values in MyInvocation.BoundParameters but I want all or nothing
    if ($username -and $password) { $argList += " -username `"$username`" -password `"$password`""}
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

function MainWindow
{
    function Get-IniConfig
    {
        # Parses the config file and returns an array of the desired settings section

        param
        (
            [Parameter(Position=0, Mandatory=$true)]
            [String]$Section
        )
        
        $startIndex = $config | Select-String -Pattern "\[$Section\]" | Select-Object -ExpandProperty LineNumber
        $sectionLength = $config[$startIndex..$config.Length] | Select-String -Pattern "\[*\]" | Select-Object -First 1 -ExpandProperty LineNumber
        if (!$sectionLength) { $sectionLength = $config.Length - $startIndex + 1}
        return $config[$startIndex..($startindex + $sectionLength - 2)] | Where-Object { $_.trim() -ne "" -and $_.trim() -notmatch "^;" }
    }

    function ClearFields
    {
        # Resets the form

        $cBuilding.SelectedIndex = 0
        $cBuilding.ForeColor     = 'DarkGray'
        $tRoom.Text              = 'Room Number'
        $tRoom.ForeColor         = 'DarkGray'
        $tCounter.Text           = '0'
        $tCounter.ForeColor      = 'DarkGray'
        $tAssetTag.Text          = [String](Get-WmiObject -Class Win32_SystemEnclosure).SMBIOSAssetTag
        $tAssetTag.ForeColor     = 'DarkGray'
        $rEmployee.Checked       = $true
        $rDesktop.Checked        = $true
        $StatusBar.Text          = ''
        $tNewName.Text           = ''
        $ltCounter.ForeColor     = $ltRoom.ForeColor = $lcBuilding.ForeColor = $ltAssetTag.ForeColor = 'Black'
    }

    function ValidateData
    {
        # Ensure form inputs are okay to use for a proper computer name. There are a lot of regex queries in here and I'm not going to explain them all.

        param( [switch]$highlightErrors )
        
        # Create blank error message. This gets overwritten so that it only contains a message for the foremost erroneous field 
        $errMsg = ''

        if ($tAssetTag.Text -eq 'Asset Tag') { $errMsg = "Please enter the machine's asset tag"; if ($highlightErrors) { $ltAssetTag.ForeColor = 'Red' } }
        elseif ($tAssetTag.Text -notmatch '^[0-9]{6}$') { $errMsg = 'Invalid asset tag'; $ltAssetTag.ForeColor = 'Red' }
        else { $ltAssetTag.ForeColor = 'Black' }

        if (($tCounter.Text -le 0) -and (!$rLab.checked)) { $errMsg = "Please enter the machine's counter number"; if ($highlightErrors) { $ltCounter.ForeColor = 'Red' } }
        else { $ltCounter.ForeColor = 'Black' }

        if ($tRoom.Text -eq 'Room Number') { $errMsg = 'Please enter a room number'; if ($highlightErrors) { $ltRoom.ForeColor = 'Red' } }
        elseif ($tRoom.Text -notmatch '^[0-9]{3}[a-zA-Z]{0,1}$') { $errMsg = 'Invalid room number'; $ltRoom.ForeColor = 'Red' }
        else
        {
            $ltRoom.ForeColor = 'Black'
        }

        if ($cBuilding.SelectedIndex -le 0) { $errMsg = 'Please select a building'; if ($highlightErrors) { $lcBuilding.ForeColor = 'Red' } }
        else { $lcBuilding.ForeColor = 'Black' }

        $StatusBar.Text = $errMsg
        # Check if there's an error and return true/false. If no error, generate a name.
        if ($errMsg -eq '') { GenerateName; return $true } else { $tNewName.Text = ''; return $false }
    }

    function GenerateName
    {
        # Generates a new computer name and OU based on the form fields

        $tempType      = ''
        $baseOU        = "OU=Campus Workstations,DC=mssu,DC=edu"

        # Adds the building code from the config file based on the selected building index
        $generatedName += $buildingConfig[$cBuilding.SelectedIndex - 1].Split(',')[0].Trim()

        # Add a room number and hyphon if necessary
        if ($generatedName.Length -eq 2)
        {
            $generatedName += $tRoom.Text.Trim().ToUpper()
            if ($tRoom.Text.Length -eq 3) { $generatedName += '-' }
        }

        if ($rDesktop.Checked)             { $tempType = 'D' }
        elseif ($rLaptop.Checked)          { $tempType = 'L' }
        elseif ($rTablet.Checked)          { $tempType = 'T' }
        elseif ($rPrinter.Checked)         { $tempType = 'P' }
        elseif ($rProjector.Checked)       { $tempType = 'J' }
        elseif ($rTV.Checked)              { $tempType = 'V' }
        elseif ($rCamera.Checked)          { $tempType = 'C' }
        elseif ($rPolycom.Checked)         { $tempType = 'Y' }

        if ($rEmployee.Checked)            { $generatedName += 'E' + $tempType + $tCounter.Text + $tAssetTag.Text ; $generatedOU = "OU=Employee,$baseOU" }
        elseif ($rClassroom.Checked)       { $generatedName += 'C' + $tempType + $tCounter.Text + $tAssetTag.Text ; $generatedOU = "OU=ClassRooms,$baseOU" }
        elseif ($rConference.Checked)      { $generatedName += 'F' + $tempType + $tCounter.Text + $tAssetTag.Text ; $generatedOU = "OU=ConferenceRooms,$baseOU" }
        elseif ($rKiosk.Checked)           { $generatedName += 'K' + $tempType + $tCounter.Text + $tAssetTag.Text ; $generatedOU = "OU=Kiosk,$baseOU" }
        elseif ($rFourwinds.Checked)       { $generatedName += 'W' + $tempType + $tCounter.Text + $tAssetTag.Text ; $generatedOU = "OU=FourwindsPlayers,DC=mssu,DC=edu" }
        elseif ($rSouthernWelcome.Checked) { $generatedName += 'S' + ("{0:00}" -f [Int]$tCounter.Text) + $tAssetTag.Text ; $generatedOU = "OU=Southern Welcome Laptops,OU=Employee,$baseOU"}
        elseif ($rITLoaner.Checked)        { $generatedName = "02122-ITL" + $tAssetTag.text ; $generatedOU = "OU=Employee,$baseOU" }
        elseif ($rLab.Checked)       
        {
            # If the PC is a lab computer, search AD for a corresponding Lab OU. If no match is found, just use the root Lab OU

            # Format the counter number to use a preceeding zero and add the asset tag
            $generatedName += 'L' + ("{0:00}" -f [Int]$tCounter.Text) + $tAssetTag.Text
            # Check the config file for a lab container name based on building/room
            $labContainer = ($labOUConfig | Where-Object { $_.split(',')[0] -like $cBuilding.SelectedIndex }).split(',')[1].trim() + $tRoom.text
            # Check AD to make sure container name exists. If it can't find a match, use default lab OU
            $OUToSearch = "OU=$labContainer,OU=Labs,$baseOU"
            $objSearcher.Filter = "(&(distinguishedName=$OUToSearch)(objectCategory=organizationalunit))"
            if ($objSearcher.FindOne()) { $generatedOU = $OUToSearch }
            else { $generatedOU = "OU=Labs,$baseOU" }
        }

        # update form fields with new name/OU. I'm using $script:[...] to avoid scope issues.
        $tNewName.Text = $script:generatedName = $generatedName
        $StatusBar.text = $script:generatedOU = $generatedOU
    }

    function RenameAndJoin
    {
        # Rather than outputting to a log file (like a good programmer), I'm writing status updates to the minimized Powershell window. This is mostly for debugging but I'm leaving it in the script for now.

        function CheckForConflictingObject
        {
            # Searches AD for a conflicting object, and tries to delete it if override is checked
            # The results of error popups are saved to a variable here so it doesn't interfere with returning true/false. Could also pipe to Out-Null.
            param ([string]$computerName)
            try
            {
                Write-Host "Checking for conflicting objects named $computerName"
                $objSearcher.Filter = "(&(objectCategory=Computer)(cn=$computerName))"
                $result = $objSearcher.FindOne()
                if ($result)
                {
                    Write-Host "existing object found, trying to delete" 
                    if ($xOverwrite.checked)
                    {
                        try {$result.getdirectoryentry().deleteobject(0) ; Write-Host "Success. Waiting for AD cache to update before continuing" ; Start-Sleep 20}
                        catch { $errorMsg = [system.Windows.Forms.MessageBox]::show("Could not overwrite AD object. Error message:`n`r`n`r$_",'Error') ; Write-Host "cancelling"; return $false}  
                    }
                    else
                    {
                        $errorMsg = [system.Windows.Forms.MessageBox]::show("Conflicting object found in Active Directory with name $computerName.`n`r`n`rTo overwrite this object, please check 'Overwrite existing AD object' and try again",'Error')
                        Write-Host "cancelling"
                        return $false
                    }                      
                }
            }
            catch { $errorMsg = [system.Windows.Forms.MessageBox]::show("Couldn't check for conflicting objects. Error message:`n`r`n`r$_",'Error') ;  return $false}
            return $true
        }
        
        # Double-check and make sure we're good to go
        if (!(ValidateData -highlightErrors)) { return }
        # Double check and make sure user is ready.
        $confirmation = [system.Windows.Forms.MessageBox]::show("Computer will reboot after renaming. Continue?",'Warning','YesNo')
        if ($confirmation -ne 'Yes') {return}

        Write-Host "`n`nrunning rename/join"
        $newName = $script:generatedName
        $newOU   = $script:generatedOU
        Write-Host "new name/ou: " $newName $newOU "`n"
        $joinedToDomain = (Get-WmiObject -Class Win32_ComputerSystem).PartOfDomain
        ## TODO: Everything gets messed up if you're joined to AD but you're signed into the local admin account. Fix another day.

        # Log
        "$username`n$(Get-Date)`nOld Name:  $($env:COMPUTERNAME)`nNew Name:  $newName`nJoin AD:   $($xJoinAD.Checked)`nOverwrite: $($xOverwrite.Checked)`n`n---------`n" | Out-File "$scriptRoot\log.log" -Append -Force
        
        if ($xJoinAD.Checked)
        {
            Write-Host "join AD" 
            if ($joinedToDomain) { [system.Windows.Forms.MessageBox]::show("This computer is already joined to Active Directory. `n`r`n`rPlease uncheck 'Join Active Directory' and try again",'Error') ; return }
            
            if (!(CheckForConflictingObject -computerName $env:COMPUTERNAME)) {return}
            if (!(CheckForConflictingObject -computerName $newName)) {return}

            Write-Host "trying to join" 
            try
            {
                if ($env:COMPUTERNAME -eq $newName) { Add-Computer -DomainName "mssu.edu" -OUPath $newOU -Credential $cred -Restart -ErrorAction Stop}
                else { Add-Computer -DomainName "mssu.edu" -OUPath $newOU -NewName $newName -Credential $cred -Restart -ErrorAction Stop}
            }
            catch { [system.Windows.Forms.MessageBox]::show("Could not join domain. Error message:`n`r`n`r$_",'Error') ; return }
        }
        else
        {
            Write-Host "rename computer only"
            if ($joinedToDomain)
            {
                Write-Host "checking OU"
                # Check the computer's current OU and see if it needs to be moved when it is renamed
                try
                {
                    $objSearcher.Filter = "(&(objectCategory=Computer)(cn=$($env:COMPUTERNAME)))"
                    $result = $objSearcher.FindOne()
                    $thisComputer = $result.properties["distinguishedName"]
                    if (!$thisComputer) { throw "Computer not found in AD" }
                    $currentOU = ($thisComputer.split(',',2))[1]
                    if ([string]$currentOU -notlike $newOU)
                    {
                        Write-Host "moving to correct OU"
                        $newOUObject = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$newOU", $username, $password
                        $thisComputer = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$thisComputer", $username, $password
                        $thisComputer.psbase.MoveTo($newOUObject)
                        Write-Host "success"
                    }
                    else { Write-Host "can stay in the same Ou" }
                }
                catch { [system.Windows.Forms.MessageBox]::show("Couldn't confirm/move object OU! This error can occur when the comptuter thinks it's in AD but it's actually not. If you're signed in as a local admin, try switching accounts. Cancelling.`n`r`n`r$_",'Error') ; return}

                Write-Host "calling Check"
                if (!(CheckForConflictingObject -computerName $newName)) { Write-Host "exiting"; return }

                # This app's official full name is 'Rename And Join And Move Or Rename And Move Or Just Move'
                if ($env:COMPUTERNAME -notlike $newName)
                {
                    Write-Host "renaming"
                    try { Rename-Computer -NewName $newName -Force -Restart -ErrorAction Stop }
                    catch { [system.Windows.Forms.MessageBox]::show("Could not rename computer. Error message:`n`r`n`r$_",'Error') ; return }
                }
            }
            else { Rename-Computer -NewName $newName -Force -Restart}
        }
    }

    function ToggleForm
    {
        # This function is used to the form while things are happening. I put more time into this than I really needed.
        # You can specify to Enable or Disable. If you don't specify an option it'll just toggle.
        # I never use toggle.

        [CmdletBinding(DefaultParameterSetName='Toggle')]
        param
        (
            [parameter(ParameterSetName='ForceEnable')]
            [switch]$Enabled,
            
            [parameter(ParameterSetName='ForceDisable')]
            [switch]$Disabled
        )

        $fields = @($lcBuilding, $cBuilding, $ltRoom, $ltCounter, $tRoom, $tCounter, $ltAssetTag, $tAssetTag, $grFunction, $grType, $ltNewName, $tNewName, $xJoinAD, $xOverwrite, $bRename, $bClear)      
        foreach ($field in $fields)
        {
            $targetValue =  if ($Enabled){$true} elseif ($Disabled){$false} else {!$field.Enabled}
            $field.Enabled = $targetValue
        }
    }

    # Get everything set up

    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $scriptRoot              = Get-UNCPath -ScriptRoot
    $config                  = Get-Content $scriptRoot\config.ini
    $buildingConfig          = Get-IniConfig 'Building Codes'
    $labOUConfig             = Get-IniConfig 'Lab OUs'
    $generatedName           = ''
    $generatedOU             = ''
    
    # Rather than installing the Powershell AD Module to every computer on campus, I'm using the built-in Active Directory Services Interfaces (ADSI)
    # I could have bundled the AD module and ran it from the J Drive but ADSI runs faster anyway.
    $objDomain               = New-Object System.DirectoryServices.DirectoryEntry -ArgumentList "LDAP://mssu.edu", $username, $password
    $objSearcher             = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot  = $objDomain
    $objSearcher.SearchScope = "subtree"
    $objSearcher.Filter = "(&(objectCategory=Computer))"
    # Do a quick search to make sure credentials are working. If it fails, exit.
    try   { $result = $objSearcher.FindOne() }
    catch { [system.Windows.Forms.MessageBox]::show("Invalid credentials",'Error') ; exit }

    #region BuildForm

    # I'm using collapsable regions throughout this section because there is a lot of text
    # This would have been easier if I used dropdown menus for everything but the UI nerds on the internet told me not to
    # Every time the user changes/selects/deselects a field value, it validates all feilds. It's a small form, so this isn't really a performance issue. The non-lazy way would be to write separate validation functions for each feild.
    # Winforms don't really support Placeholder Text so that's all being done manually when a field gains/loses focus. Not necessary but looks nice.

    #region Header

    $RenameAndJoin                  = New-Object system.Windows.Forms.Form
    $RenameAndJoin.ClientSize       = '400,670' # Can't make form much bigger due to 720p monitors (And fresh windows installs without graphics drivers)
    $RenameAndJoin.text             = "Rename And Join"
    $RenameAndJoin.StartPosition    = "CenterScreen"
    # $RenameAndJoin.TopMost          = $true # This breaks popups so don't use it.

    $StatusBar                      = New-Object Windows.Forms.StatusBar
    $StatusBar.Text                 = ''

    # Stupid trick you have to do with WinForms so no fields are focused until you press Tab once.
    $bDummyButton                   = New-Object System.Windows.Forms.Button
    $bDummyButton.Width             = 0
    $bDummyButton.TabIndex          = 0
    $bDummyButton.Add_LostFocus({$bDummyButton.TabStop = $false})

    $lHeader                        = New-Object System.Windows.Forms.Label
    $lHeader.text                   = "RENAME AND JOIN"
    $lHeader.AutoSize               = $true
    $lHeader.Location               = '40,35'
    $lHeader.Font                   = 'Century Gothic,25,style=Bold'

    #endregion Header

    #region UpperFields

    $lcBuilding                     = New-Object System.Windows.Forms.Label
    $lcBuilding.text                = "BUILDING"
    $lcBuilding.AutoSize            = $true
    $lcBuilding.Location            = New-Object System.Drawing.Point(50,105)
    $lcBuilding.Font                = 'Microsoft Sans Serif,7'
    $lcBuilding.Add_Click({$cBuilding.Focus()})

    $cBuilding                      = New-Object System.Windows.Forms.ComboBox
    $cBuilding.DropDownStyle        = 'DropDown'
    $cBuilding.Size                 = '300,20'
    $cBuilding.location             = '50,125'
    $cBuilding.Font                 = 'Microsoft Sans Serif,9'
    $cBuilding.ForeColor            = 'Darkgray'
    $aBuilding = ,'Building' + ($buildingConfig | ForEach-Object { $_.Split(',')[1].Trim() })
    $cBuilding.Items.AddRange($aBuilding)
    $cBuilding.SelectedIndex        = 0
    $cBuilding.AutoCompleteMode     = 'SuggestAppend'
    $cBuilding.AutoCompleteSource   = 'CustomSource'
    $cBuilding.AutoCompleteCustomSource.AddRange($aBuilding);
    $cBuilding.Add_GotFocus{ if($cBuilding.SelectedIndex -eq 0) { $cBuilding.Text = ""; $cBuilding.ForeColor = "Black" } }
    $cBuilding.Add_LostFocus{ if($cBuilding.SelectedIndex -eq 0 -or $cBuilding.SelectedIndex -eq -1) { $cBuilding.Text = "Building"; $cBuilding.ForeColor = "Darkgray" } }
    $cBuilding.Add_SelectedIndexChanged({ValidateData})

    $ltRoom                         = New-Object System.Windows.Forms.Label
    $ltRoom.text                    = "ROOM NUMBER"
    $ltRoom.AutoSize                = $true
    $ltRoom.Location                = '50,165'
    $ltRoom.Font                    = 'Microsoft Sans Serif,7'
    $ltRoom.Add_Click({$tRoom.Focus()})

    $tRoom                          = New-Object system.Windows.Forms.TextBox
    $tRoom.multiline                = $false
    $tRoom.location                 = '50,185'
    $tRoom.size                     = '100,20'
    $tRoom.Text                     = "Room Number"
    $tRoom.ForeColor                = 'DarkGray'
    $tRoom.Font                     = 'Microsoft Sans Serif,9'
    $tRoom.Add_GotFocus({if($tRoom.Text -eq 'Room Number') {$tRoom.Text = ''; $tRoom.ForeColor = 'Black'}})
    $tRoom.Add_LostFocus({if($tRoom.Text -eq ''){$tRoom.Text = 'Room Number'; $tRoom.ForeColor = 'Darkgray'} else { $tRoom.text = $tRoom.text.ToUpper().Trim()}; ValidateData})
    $tRoom.Add_TextChanged({ if($tRoom.Text.Length -ge 3) { ValidateData } })

    $ltCounter                      = New-Object System.Windows.Forms.Label
    $ltCounter.text                 = "COUNT"
    $ltCounter.AutoSize             = $true
    $ltCounter.Location             = '170,165'
    $ltCounter.Font                 = 'Microsoft Sans Serif,7'
    $ltCounter.Add_Click({$tCounter.Focus()})

    $tCounter                       = New-Object System.Windows.Forms.NumericUpDown
    $tCounter.Size                  = '60,20'
    $tCounter.location              = '170,185'
    $tCounter.ForeColor             = 'DarkGray'
    $tCounter.Font                  = 'Microsoft Sans Serif,9'
    $tCounter.Add_GotFocus({if($tCounter.Text -eq '0') {$tCounter.Text = ''; $tCounter.ForeColor = 'Black'}})
    $tCounter.Add_LostFocus({if($tCounter.Text -le '0'){$tCounter.Text = '0'; $tCounter.ForeColor = 'Darkgray'}})
    $tCounter.Add_TextChanged({ValidateData})

    $ltAssetTag                     = New-Object System.Windows.Forms.Label
    $ltAssetTag.text                = "ASSET TAG"
    $ltAssetTag.AutoSize            = $true
    $ltAssetTag.Location            = '250,165'
    $ltAssetTag.Font                = 'Microsoft Sans Serif,7'
    $ltAssetTag.Add_Click({$tAssetTag.Focus()})

    $tAssetTag                      = New-Object system.Windows.Forms.TextBox
    $tAssetTag.multiline            = $false
    $tAssetTag.Size                 = '100,20'
    $tAssetTag.location             = '250,185'
    $tAssetTag.Text                 = [String](Get-WmiObject -Class Win32_SystemEnclosure).SMBIOSAssetTag.Trim()
    $tAssetTag.Font                 = 'Microsoft Sans Serif,9'
    $tAssetTag.ForeColor            = 'Darkgray'
    $tAssetTag.Add_GotFocus({if($tAssetTag.Text -eq 'Asset Tag') {$tAssetTag.Text = ''}; $tAssetTag.ForeColor = 'Black'})
    $tAssetTag.Add_LostFocus({if($tAssetTag.Text -eq ''){$tAssetTag.Text = 'Asset Tag'; $tAssetTag.ForeColor = 'Darkgray'}})
    $tAssetTag.Add_TextChanged({ValidateData})

    #endregion UpperFeilds

    #region FunctionGroup

    $grFunction                     = New-Object System.Windows.Forms.GroupBox
    $grFunction.Location            = '50,230'
    $grFunction.Size                = '130,230'
    $grFunction.Text                = 'DEVICE FUNCTION'
    $grFunction.Font                = 'Microsoft Sans Serif,6.5'
    $grFunction.Add_Click({ValidateData})

    $rEmployee                      = New-Object System.Windows.Forms.RadioButton
    $rEmployee.Location             = '20,20'
    $rEmployee.AutoSize             = $true
    $rEmployee.Checked              = $true
    $rEmployee.Text                 = 'Employee'
    $rEmployee.Font                 = 'Microsoft Sans Serif,9'
    $rEmployee.Add_Click({ValidateData})

    $rLab                           = New-Object System.Windows.Forms.RadioButton
    $rLab.Location                  = '20,45'
    $rLab.AutoSize                  = $true
    $rLab.Checked                   = $false
    $rLab.Text                      = 'Lab'
    $rLab.Font                      = 'Microsoft Sans Serif,9'
    $rLab.Add_Click({ValidateData})

    $rClassroom                     = New-Object System.Windows.Forms.RadioButton
    $rClassroom.Location            = '20,70'
    $rClassroom.AutoSize            = $true
    $rClassroom.Checked             = $false
    $rClassroom.Text                = 'Classroom'
    $rClassroom.Font                = 'Microsoft Sans Serif,9'
    $rClassroom.Add_Click({ValidateData})

    $rConference                    = New-Object System.Windows.Forms.RadioButton
    $rConference.Location           = '20,95'
    $rConference.AutoSize           = $true
    $rConference.Checked            = $false
    $rConference.Text               = 'Conference'
    $rConference.Font               = 'Microsoft Sans Serif,9'
    $rConference.Add_Click({ValidateData})

    $rKiosk                         = New-Object System.Windows.Forms.RadioButton
    $rKiosk.Location                = '20,120'
    $rKiosk.AutoSize                = $true
    $rKiosk.Checked                 = $false
    $rKiosk.Text                    = 'Kiosk'
    $rKiosk.Font                    = 'Microsoft Sans Serif,9'
    $rKiosk.Add_Click({ValidateData})

    $rFourwinds                     = New-Object System.Windows.Forms.RadioButton
    $rFourwinds.Location            = '20,145'
    $rFourwinds.AutoSize            = $true
    $rFourwinds.Checked             = $false
    $rFourwinds.Text                = 'Fourwinds'
    $rFourwinds.Font                = 'Microsoft Sans Serif,9'
    $rFourwinds.Add_Click({ValidateData})

    $rSouthernWelcome               = New-Object System.Windows.Forms.RadioButton
    $rSouthernWelcome.Location      = '20,170'
    $rSouthernWelcome.AutoSize      = $true
    $rSouthernWelcome.Checked       = $false
    $rSouthernWelcome.Text          = 'S. Welcome'
    $rSouthernWelcome.Font          = 'Microsoft Sans Serif,9'
    $rSouthernWelcome.Add_Click({ValidateData})

    $rITLoaner                      = New-Object System.Windows.Forms.RadioButton
    $rITLoaner.Location             = '20,195'
    $rITLoaner.AutoSize             = $true
    $rITLoaner.Checked              = $false
    $rITLoaner.Text                 = 'IT Loaner'
    $rITLoaner.Font                 = 'Microsoft Sans Serif,9'
    $rITLoaner.Add_Click({ValidateData})

    $grFunction.Controls.AddRange(@($rEmployee, $rLab, $rClassroom, $rConference, $rKiosk, $rFourwinds, $rSouthernWelcome, $rITLoaner))
    $grFunction.Add_LostFocus({ValidateData})

    #endregion FunctionGroup

    #region TypeGroup

    $grType                         = New-Object System.Windows.Forms.GroupBox
    $grType.Location                = '220,230'
    $grType.Size                    = '130,230'
    $grType.Text                    = 'DEVICE TYPE'
    $grType.Font                    = 'Microsoft Sans Serif,6.5'
    $grType.Add_Click({ValidateData})

    $rDesktop                       = New-Object System.Windows.Forms.RadioButton
    $rDesktop.Location              = '20,20'
    $rDesktop.AutoSize              = $true
    $rDesktop.Checked               = $true
    $rDesktop.Text                  = 'Desktop'
    $rDesktop.Font                  = 'Microsoft Sans Serif,9'
    $rDesktop.Add_Click({ValidateData})

    $rLaptop                        = New-Object System.Windows.Forms.RadioButton
    $rLaptop.Location               = '20,45'
    $rLaptop.AutoSize               = $true
    $rLaptop.Checked                = $false
    $rLaptop.Text                   = 'Laptop'
    $rLaptop.Font                   = 'Microsoft Sans Serif,9'
    $rLaptop.Add_Click({ValidateData})

    $rTablet                        = New-Object System.Windows.Forms.RadioButton
    $rTablet.Location               = '20,70'
    $rTablet.AutoSize               = $true
    $rTablet.Checked                = $false
    $rTablet.Text                   = 'Tablet'
    $rTablet.Font                   = 'Microsoft Sans Serif,9'
    $rTablet.Add_Click({ValidateData})

    $rPrinter                       = New-Object System.Windows.Forms.RadioButton
    $rPrinter.Location              = '20,95'
    $rPrinter.AutoSize              = $true
    $rPrinter.Checked               = $false
    $rPrinter.Text                  = 'Printer'
    $rPrinter.Font                  = 'Microsoft Sans Serif,9'
    $rPrinter.Add_Click({ValidateData})

    $rProjector                     = New-Object System.Windows.Forms.RadioButton
    $rProjector.Location            = '20,120'
    $rProjector.AutoSize            = $true
    $rProjector.Checked             = $false
    $rProjector.Text                = 'Projector'
    $rProjector.Font                = 'Microsoft Sans Serif,9'
    $rProjector.Add_Click({ValidateData})

    $rTV                            = New-Object System.Windows.Forms.RadioButton
    $rTV.Location                   = '20,145'
    $rTV.AutoSize                   = $true
    $rTV.Checked                    = $false
    $rTV.Text                       = 'TV'
    $rTV.Font                       = 'Microsoft Sans Serif,9'
    $rTV.Add_Click({ValidateData})

    $rCamera                        = New-Object System.Windows.Forms.RadioButton
    $rCamera.Location               = '20,170'
    $rCamera.AutoSize               = $true
    $rCamera.Checked                = $false
    $rCamera.Text                   = 'Camera'
    $rCamera.Font                   = 'Microsoft Sans Serif,9'
    $rCamera.Add_Click({ValidateData})

    $rPolycom                       = New-Object System.Windows.Forms.RadioButton
    $rPolycom.Location              = '20,195'
    $rPolycom.AutoSize              = $true
    $rPolycom.Checked               = $false
    $rPolycom.Text                  = 'Polycom'
    $rPolycom.Font                  = 'Microsoft Sans Serif,9'
    $rPolycom.Add_Click({ValidateData})

    $grType.Controls.AddRange(@($rDesktop, $rLaptop, $rTablet, $rPrinter, $rProjector, $rTV, $rCamera, $rPolycom))

    #endregion TypeGroup

    #region LowerFields

    $xJoinAD                        = New-Object System.Windows.Forms.CheckBox
    $xJoinAD.Location               = '50,478'
    $xJoinAD.AutoSize               = $true
    $xJoinAD.Text                   = 'Join Active Directory'
    $xJoinAD.Checked                = !(Get-WmiObject -Class Win32_ComputerSystem).PartOfDomain
    $xJoinAD.Font                   = 'Microsoft Sans Serif,9'
    $xJoinAD.Add_CheckStateChanged({if($xJoinAD.Checked -eq $false) {$bRename.Text = "Rename"} else {$bRename.Text = "Rename And Join"}})

    $xOverwrite                     = New-Object System.Windows.Forms.CheckBox
    $xOverwrite.Location            = '50,503'
    $xOverwrite.AutoSize            = $true
    $xOverwrite.Text                = 'Overwrite existing AD object'
    $xOverwrite.Checked             = $false
    $xOverwrite.Font                = 'Microsoft Sans Serif,9'

    $ltCurrentName                  = New-Object System.Windows.Forms.Label
    $ltCurrentName.text             = "CURRENT MACHINE NAME"
    $ltCurrentName.AutoSize         = $true
    $ltCurrentName.Location         = '50,540'
    $ltCurrentName.Font             = 'Microsoft Sans Serif,7'
    $ltCurrentName.Add_Click({$tCurrentName.Focus()})

    $tCurrentName                   = New-Object System.Windows.Forms.TextBox
    $tCurrentName.size              = '130,10'
    $tCurrentName.location          = '50,560'
    $tCurrentName.Font              = 'Microsoft Sans Serif,9'
    $tCurrentName.Text              = $env:COMPUTERNAME
    $tCurrentName.ReadOnly          = $true
    $tCurrentName.TabStop           = $false
    $tCurrentName.Enabled           = $false
    $tCurrentName.TextAlign         = "Center"

    $ltNewName                      = New-Object System.Windows.Forms.Label
    $ltNewName.text                 = "NEW MACHINE NAME"
    $ltNewName.AutoSize             = $true
    $ltNewName.Location             = '220,540'
    $ltNewName.Font                 = 'Microsoft Sans Serif,7'
    $ltNewName.Add_Click({$tNewName.Focus()})

    $tNewName                       = New-Object System.Windows.Forms.TextBox
    $tNewName.size                  = '130,10'
    $tNewName.location              = '220,560'
    $tNewName.Font                  = 'Microsoft Sans Serif,9'
    $tNewName.Text                  = ''
    $tNewName.ReadOnly              = $true
    $tNewName.TabStop               = $false
    $tNewName.TextAlign             = "Center"

    $bRename                        = New-Object System.Windows.Forms.Button
    $bRename.Location               = '50,600'
    $bRename.Size                   = '130,20'
    $bRename.Text                   = 'Rename And Join'
    $bRename.Font                   = 'Microsoft Sans Serif,9'
    $bRename.Add_Click({ToggleForm -Disabled; RenameAndJoin; ToggleForm -Enabled})

    $bClear                         = New-Object System.Windows.Forms.Button
    $bClear.Location                = '220,600'
    $bClear.Size                    = '130,20'
    $bClear.Text                    = 'Reset Fields'
    $bClear.Font                    = 'Microsoft Sans Serif,9'
    $bClear.Add_Click({ClearFields})

    #endregion LowerFields

    $RenameAndJoin.Controls.AddRange(@($bDummyButton, $lHeader, $StatusBar, $lcBuilding, $cBuilding, $ltRoom, $ltCounter, $tRoom, $tCounter, $ltAssetTag, $tAssetTag, $grFunction, $grType, $ltNewName, $tNewName, $tCurrentName, $ltCurrentName,$xJoinAD, $xOverwrite, $bRename, $bClear))

    #endregion BuildForm

    $RenameAndJoin.Add_Shown({$RenameAndJoin.Activate()})
    $RenameAndJoin.ShowDialog() | Out-Null
}

# If the credentials weren't specified in a launcher script, get them now. Credentials are needed for AD actions on unjoined computer.
# I have to get plain text credentials because I'm going to pass them through Start-Process to self-elevate the script. I don't like this any more than you do. There are ways around this but all of them suck.
if (!$username -or !$password)
{
    $cred     = Get-Credential -Message "Please enter your employee username/password`r`n`r`n(Southern\Username)"
    $username = $cred.UserName
    if ($username -notlike "SOUTHERN\*") {$username = "SOUTHERN\" + $username}
    $password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.Password))        
}

$cred = New-Object System.Management.Automation.PSCredential ($username, (ConvertTo-SecureString $password -AsPlainText -Force))

# Ensure that the script is elevated and then launch the main function
if (RunAsAdmin) { MainWindow }
