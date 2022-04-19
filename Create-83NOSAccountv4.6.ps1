CLASSIFICATION: UNCLASSIFIED
<#
Author: Christopher Steele (1456084571)
Last Modified: 2022-01-28

The purpose of this script is to emulate a shitter DRA (hard to achieve, I know, but I did it) to streamline 
    the account creation process and to standardize account properties.

This script is best run with Area42 credentials.  Do not run in Powershell ISE because ISE will lock up after a while.
This script can be run with ACC or AFMC creds, but then you can only make accounts in the respective domain.
When running the script, a GUI window will pop up with fields to populate with information from the user's 2875.
This same GUI window allows you to select which accounts to create in which domains between ACC, AFMC, and AREA42.
Pressing cancel at any time, on almost any popup, will abort the script immediately.

The GUI itself is dynamically built based on the $fields variable;
While you will have to manually add how each field is handled when making the GUI, you will not have to readjust the window size.

Error handling is robust, so it will show you (red and bold) exactly what you did wrong, 1 field at a time.

Valdiates accounts made by running admin, as is done with NIPR DRA.

Too lazy to care about version logs
#>
    
#Escalate our window
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
if ($env:COMPUTERNAME -ne "MUHJW-8321HL" -and -not $myWindowsPrincipal.IsInRole($adminRole)) {
    $scriptpath = $MyInvocation.MyCommand.Definition
    $scriptpaths = "'$scriptPath'"
    Start-Process -FilePath PowerShell.exe -Verb runAs -ArgumentList "& $scriptPaths"
    exit
    }

#Set our accounts flags, so we can know what domains we can touch.
Remove-Variable accCreds,afmcCreds,42Creds,AfnoCreds,accRootCreds -EA silentlycontinue
if ((whoami) -match ".adm" -and (whoami) -notmatch "area42\\") {
    Write-Host ("Close and run this script using AREA42 creds to create accounts across domains.`nPress enter to just create accounts in the " + (whoami).split("\")[0].toUpper() + " domain.")
    Read-Host
    }
    
#Get what credentials we have
switch ((whoami).split("\")[0].toUpper()) {
    "ACCROOT" {$accRootCreds = $true;break}
    "ACC" {$accCreds = $true}
    "SAFMC" {$afmcCreds = $true}
    "AFNOAPPS" {$AfnoCreds = $true;break}
    "AREA42" {$Area42Creds = $true}
    "USAFE" {$usafeCreds = $true}
    "USAFEROOT" {$usafeRootCreds = $true}
    "AETC" {$aetcCreds = $true}
    "AETCROOT" {$aetcRootCreds = $true}
    "AMC-S" {$amcCreds = $true}
    "DS-S" {$amcRootCreds = $true}
    "AFSPC-S" {$afspcCreds = $true}
    "AFSPC-RT" {$afspcRootCreds = $true}
    "SPACAF-ROOT" {$pacafRootCreds = $true}
    }

#Script to make shell accounts for 83 NOS User and admin accounts
Import-Module activedirectory

function Generate-Password {
    #Password command forces 2 sym / 2 num / 2 UPPER / 2 lower and randomizes last 8 
    #Common error digits have been removed ex: oO0 ,. :;
    $CharsD = [Char[]]"123456789" 
    $CharsL = [Char[]]"abcdefghjkmnpqrstuvxyz"
    $CharsU = [Char[]]"ABCDEFGHJKLMNPQRSTUVXYZ"
    $CharsS = [Char[]]"!@#$%^&*+=?"
    $CharsA = [Char[]]"!@#$%^&*+=?ABCDEFGHJKLMNPQRSTUVXYZabcdefghjkmnpqrstuvxyz123456789"
    $Password = ""
    $Password += ($CharsD | Get-Random -Count 2) -join ""
    $Password += ($CharsL | Get-Random -Count 2) -join ""
    $Password += ($CharsU | Get-Random -Count 2) -join ""
    $Password += ($CharsS | Get-Random -Count 2) -join ""
    $Password += ($CharsA | Get-Random -Count (8..12 | Get-Random)) -join ""
    $Password = ($Password.ToCharArray()| Sort-Object {Get-Random}) -join ""
    Write-Output $Password
}

#Initialize some constants
#if (Get-ADOrganizationalUnit "OU=New Bases,DC=acc,DC=accroot,DC=ds,DC=af,DC=smil,DC=mil") {$UserOU = "OU=NOSC Users,OU=NOSC,OU=New Bases,DC=acc,DC=accroot,DC=ds,DC=af,DC=smil,DC=mil"}
#else {$UserOU = "OU=NOSC Users,OU=NOSC,OU=Bases,DC=acc,DC=accroot,DC=ds,DC=af,DC=smil,DC=mil"}
#$UserOU = "OU=Langley AFB Users,OU=Langley AFB,OU=Bases,DC=acc,DC=accroot,DC=ds,DC=af,DC=smil,DC=mil"
#$UserOU = "OU=NOSC Users,OU=NOSC,OU=Bases,DC=acc,DC=accroot,DC=ds,DC=af,DC=smil,DC=mil" #Only put DS users in the NOSC OU (maybe?)
$BasesOU = Get-ADOrganizationalUnit 5cebf52c-4ea5-4fc8-a89d-9c5a76e940b6 | select -ExpandProperty DistinguishedName
$UserOU = Get-ADOrganizationalUnit -SearchBase $basesOU -Filter {Name -eq "83 NOS Users"} | select -ExpandProperty DistinguishedName
$ACCAdminOU = "OU=83 NOS,OU=Administrative Accounts,OU=Administration,DC=acc,DC=accroot,DC=ds,DC=af,DC=smil,DC=mil"
$ACCRootAdminOU = "OU=83 NOS,OU=Administrative Accounts,OU=Administration,DC=accroot,DC=ds,DC=af,DC=smil,DC=mil"
$AFMCAdminOU = "OU=83 NOS,OU=Administrative Accounts,OU=Administration,DC=afmc,DC=ds,DC=af,DC=smil,DC=mil"
$42AdminOU = "OU=83 NOS,OU=Administrative Accounts,OU=Administration,DC=area42,DC=afnoapps,DC=usaf,DC=smil,DC=mil"
$USAFEAdminOU = "OU=83 NOS,OU=Administrative Accounts,OU=Administration,DC=usafe,DC=usaferoot,DC=ds,DC=af,DC=smil,DC=mil"
$USAFERootAdminOU = "OU=83 NOS,OU=Administrative Accounts,OU=Administration,DC=usaferoot,DC=ds,DC=af,DC=smil,DC=mil"
$AETCAdminOU = "OU=83 NOS,OU=Administrative Accounts,OU=Administration,DC=aetc,DC=aetcroot,DC=ds,DC=af,DC=smil,DC=mil"
$AETCRootAdminOU = "OU=83 NOS,OU=Administrative Accounts,OU=Administration,DC=aetcroot,DC=ds,DC=af,DC=smil,DC=mil"
$AMCAdminOU = "OU=83 NOS,OU=Administrative Accounts,OU=Administration,DC=amchub,DC=amc,DC=ds,DC=af,DC=smil,DC=mil"
$AMCRootAdminOU = "OU=83 NOS,OU=Administrative Accounts,OU=Administration,DC=amc,DC=ds,DC=af,DC=smil,DC=mil"
$AFSPCAdminOU = "OU=83 NOS,OU=Administrative Accounts,OU=Administration,DC=afspc-s,DC=afspc,DC=ds,DC=af,DC=smil,DC=mil"
$AFSPCRootAdminOU = "OU=Administrative Accounts,OU=AFSPC NOSC,OU=Admin,DC=afspc,DC=ds,DC=af,DC=smil,DC=mil"
$PACAFAdminOU = "OU=83 NOS,OU=Administrative Accounts,OU=Administration,DC=pacaf,DC=ds,DC=af,DC=smil,DC=mil"

$ACC = "acc.accroot.ds.af.smil.mil"
$ACCROOT = "accroot.ds.af.smil.mil"
$AFMC = "afmc.ds.af.smil.mil"
$AREA42 = "area42.afnoapps.usaf.smil.mil"
$AFNOAPPS = "afnoapps.usaf.smil.mil"
$USAFE = "usafe.usaferoot.ds.af.smil.mil"
$USAFEROOT = "usaferoot.ds.af.smil.mil"
$AETC = "aetc.aetcroot.ds.af.smil.mil"
$AETCROOT = "aetc.aetcroot.ds.af.smil.mil"
$AMCHUB = "amchub.amc.ds.af.smil.mil"
$AMC = "amc.ds.af.smil.mil"
$AFSPC = "afspc-s.afspc.ds.af.smil.mil"
$AFSPCROOT = "afspc.ds.af.smil.mil"
$PACAF = "pacaf.ds.af.smil.mil"

$ACCDN = "DC=acc,DC=accroot,DC=ds,DC=af,DC=smil,DC=mil"
$ACCROOTDN = "DC=accroot,DC=ds,DC=af,DC=smil,DC=mil"
$AFMCDN = "DC=afmc,DC=ds,DC=af,DC=smil,DC=mil"
$AREA42DN = "DC=area42,DC=afnoapps,DC=usaf,DC=smil,DC=mil"
$AFNOAPPSDN = "DC=afnoapps,DC=usaf,DC=smil,DC=mil"
$USAFEDN = "DC=usafe,DC=usaferoot,DC=ds,DC=af,DC=smil,DC=mil"
$USAFEROOTDN = "DC=usaferoot,DC=ds,DC=af,DC=smil,DC=mil"
$AETCDN = "DC=aetc,DC=aetcroot,DC=ds,DC=af,DC=smil,DC=mil"
$AETCROOTDN = "DC=aetcroot,DC=ds,DC=af,DC=smil,DC=mil"
$AMCDN = "DC=amchub,DC=amc,DC=ds,DC=af,DC=smil,DC=mil"
$AMCROOTDN = "DC=amc,DC=ds,DC=af,DC=smil,DC=mil"
$AFSPCDN = "DC=afspc-s,DC=afspc,DC=ds,DC=af,DC=smil,DC=mil"
$AFSPCROOTDN = "DC=afspc,DC=ds,DC=af,DC=smil,DC=mil"
$PACAFDN = "DC=pacaf,DC=ds,DC=af,DC=smil,DC=mil"

$ACCDC = Get-ADDomainController -Server $acc | Select -ExpandProperty hostname
$AdmDC = $ACCDC
if ($env:USERDNSDOMAIN -ne $ACC) {$AdmDC = Get-ADDomainController -Server $env:USERDNSDOMAIN | Select -ExpandProperty hostname}
$RunningCN = Get-ADUser -Server $AdmDC $env:USERNAME | select -ExpandProperty name

##Input data
##GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Initialize our style spacings
$string = 'T'
$font = [System.Windows.Forms.Label]::DefaultFont

$size = [System.Windows.Forms.TextRenderer]::MeasureText($string, $font)
$CharHeight = $size.Height
#$CharWidth = $size.Width #I dont know why this is wrong, but it is
$CharWidth = 7

$LabelBuffer = 3 #Pixels betwen label and its data entry field
$ComboBuffer = 5 #Pixels between data entry and next label
$TextBoxHeight = 20

$OffsetX = 10 #Pixels from left border to first element

#Loop for multiple users
do {
    $DateValidated = Get-Date -Format "yyyyMMdd"
    $OffsetY = 20 #Pixels from the bottom of the previous element (including top border) to next element
    
    #Generate our GUI
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Create 83 NOS account"
    $form.StartPosition = "CenterScreen"

    #Dynamically create GUI elements based on what forms of data to display/input
    $LabelHash = @{}
    $DataEntryHash = @{}

    $Fields = @"
Remedy Ticket Numer
EDIPI
E-mail
Last Name
First Name
Middle Initial
Generation Qualifier
Branch
MAJCOM
Organization
Office Symbol
DSN Number (###-####)
Personnel Type
Rank (Leave blank if N/A)
Address
Accounts to Create (CTRL+Click to choose multiple)
Groups
Error
"@.split("`n") | foreach {$_.trim()}
    :main foreach ($field in $Fields) {
        $LabelHash[$field] = New-Object System.Windows.Forms.Label

        $LabelHash[$field].Location = New-Object System.Drawing.Point($OffsetX,$OffsetY) 
        $LabelHash[$field].Size = New-Object System.Drawing.Size(280,$CharHeight) 
        $LabelHash[$field].Text = $Field + ":"
        #Keep a blank space for error messages incase the user is an idiot
        switch ($field) {
            "Error" {
                $LabelHash[$field].Text = ""
                $LabelHash[$field].ForeColor = "Red"
                $LabelHash[$field].Font = New-Object Drawing.Font($LabelHash[$field].Font.FontFamily, $LabelHash[$field].Font.Size, [Drawing.FontStyle]::Bold)
                break
                }
            "Groups" {
                $LabelHash[$Field].Text = "Separate windows will popup for pasting groups into."
                $LabelHash[$field].Font = New-Object Drawing.Font($LabelHash[$field].Font.FontFamily, $LabelHash[$field].Font.Size, [Drawing.FontStyle]::Bold)
                break
                }
            default {break}
            }
        $form.Controls.Add($LabelHash[$field]) 
        
        $OffsetY += $LabelBuffer + $CharHeight
        
        #We want to control what data can be selected for Personnel Type, and make sure only one item is selected.
        #Could have done a listbox without multi select on, but dropdowns are more compact
        if ($Field -match "Personnel Type") {
            $DataEntryHash[$Field] = new-object System.Windows.Forms.ComboBox
            $DataEntryHash[$Field].Location = new-object System.Drawing.Size($OffsetX,$OffsetY)
            $DataEntryHash[$Field].Size = new-object System.Drawing.Size(260,$CharHeight)
            $DataEntryHash[$Field].DropDownStyle = "DropDownList"

            $OffsetY += $TextBoxHeight + $ComboBuffer
            
            $DataEntryHash[$Field].Items.Add("Military") | Out-Null
            $DataEntryHash[$Field].Items.Add("Reservist") | Out-Null
            $DataEntryHash[$Field].Items.Add("Civilian") | Out-Null
            $DataEntryHash[$Field].Items.Add("Contractor") | Out-Null

            $Form.Controls.Add($DataEntryHash[$Field])
            }
        elseif ($Field -match "Generation Qualifier") {
            $DataEntryHash[$Field] = new-object System.Windows.Forms.ComboBox
            $DataEntryHash[$Field].Location = new-object System.Drawing.Size($OffsetX,$OffsetY)
            $DataEntryHash[$Field].Size = new-object System.Drawing.Size(260,$CharHeight)
            $DataEntryHash[$Field].DropDownStyle = "DropDownList"

            $OffsetY += $TextBoxHeight + $ComboBuffer
            
            $DataEntryHash[$Field].Items.Add("(none)") | Out-Null
            $DataEntryHash[$Field].Items.Add("Jr") | Out-Null
            $DataEntryHash[$Field].Items.Add("Sr") | Out-Null
            $DataEntryHash[$Field].Items.Add("II") | Out-Null
            $DataEntryHash[$Field].Items.Add("III") | Out-Null
            $DataEntryHash[$Field].Items.Add("IV") | Out-Null
            $DataEntryHash[$Field].Items.Add("V") | Out-Null

            #default to none
            $DataEntryHash[$Field].SelectedIndex = 0

            $Form.Controls.Add($DataEntryHash[$Field])
            }
        #Listbox to highlight accounts to create
        elseif ($Field -match "Accounts to Create") {
            $DataEntryHash[$Field] = New-Object System.Windows.Forms.ListBox 
            $DataEntryHash[$Field].Location = New-Object System.Drawing.Point($OffsetX,$OffsetY) 
            $DataEntryHash[$Field].Size = New-Object System.Drawing.Size(260,$CharHeight) 
            $DataEntryHash[$Field].SelectionMode = "MultiExtended"

            if ($Area42Creds -or $AfnoCreds -or $accRootCreds -or $accCreds) {
                [void] $DataEntryHash[$Field].Items.Add("ACC User")
                [void] $DataEntryHash[$Field].Items.Add("ACC Admin(s)")
                }
            if ($Area42Creds -or $AfnoCreds -or $accRootCreds) {
                [void] $DataEntryHash[$Field].Items.Add("ACCROOT Admin(s)")
                }
            if ($Area42Creds -or $AfnoCreds -or $afmcCreds) {
                [void] $DataEntryHash[$Field].Items.Add("AFMC Admin(s)")
                }
            if ($Area42Creds -or $AfnoCreds -or $usafeRootCreds -or $usafeCreds) {
                [void] $DataEntryHash[$Field].Items.Add("USAFE Admin(s)")
                }
            if ($Area42Creds -or $AfnoCreds -or $usafeRootCreds) {
                [void] $DataEntryHash[$Field].Items.Add("USAFERoot Admin(s)")
                }
            if ($Area42Creds -or $AfnoCreds -or $aetcRootCreds -or $aetcCreds) {
                [void] $DataEntryHash[$Field].Items.Add("AETC Admin(s)")
                }
            if ($Area42Creds -or $AfnoCreds -or $aetcRootCreds) {
                [void] $DataEntryHash[$Field].Items.Add("AETCRoot Admin(s)")
                }
            if ($Area42Creds -or $AfnoCreds-or $amcRootCreds -or $amcCreds ) {
                [void] $DataEntryHash[$Field].Items.Add("AMC Admin(s)")
                }
            if ($Area42Creds -or $AfnoCreds -or $amcRootCreds) {
                [void] $DataEntryHash[$Field].Items.Add("AMCRoot Admin(s)")
                }
            if ($Area42Creds -or $AfnoCreds -or $afspcRootCreds -or $afspcCreds) {
                [void] $DataEntryHash[$Field].Items.Add("AFSPC Admin(s)")
                }
            if ($Area42Creds -or $AfnoCreds -or $afspcRootCreds) {
                [void] $DataEntryHash[$Field].Items.Add("AFSPCRoot Admin(s)")
                }
            if ($Area42Creds -or $AfnoCreds -or $pacafRootCreds) {
                [void] $DataEntryHash[$Field].Items.Add("PACAF Admin(s)")
                }
            if ($Area42Creds -or $AfnoCreds) {
                [void] $DataEntryHash[$Field].Items.Add("AREA42 Admin(s)")
                }
            
            $DataEntryHash[$Field].Height = $CharHeight * ($DataEntryHash[$Field].Items.count + 1) + $LabelBuffer

            $form.Controls.Add($DataEntryHash[$Field]) 

            $OffsetY += $DataEntryHash[$Field].Height
            }
        #Every other item is manual entry, so just use Textboxes
        else {
            $DataEntryHash[$Field] = New-Object System.Windows.Forms.TextBox

            $DataEntryHash[$Field].Location = New-Object System.Drawing.Point($OffsetX,$OffsetY) 
            $DataEntryHash[$Field].Size = New-Object System.Drawing.Size(260,$TextBoxHeight)
        
            #Set any character limits for data entry, or prefill data.  
            #Error and Groups ends its iteration because we don't need data entry, just a line to display a message
            switch -Regex ($Field) {
                "EDIPI" {$DataEntryHash[$Field].MaxLength = 10}
                "Middle" {$DataEntryHash[$Field].MaxLength = 1}
                "Branch" {$DataEntryHash[$Field].Text = "USAF"} #Should never not be USAF for us
                "MAJCOM" {$DataEntryHash[$Field].Text = "ACC"} #Should never not be AFSPC for us
                "Organization" {$DataEntryHash[$Field].Text = "83 NOS"}
                "DSN" {$DataEntryHash[$Field].MaxLength = 12}
                "Address" {$DataEntryHash[$Field].Multiline = $true;$DataEntryHash[$Field].Size = New-Object System.Drawing.Size(260,$TextBoxHeight);$DataEntryHash[$Field].Text = "37 Elm St, Langley AFB Hampton, VA 23665"}
                "City" {$DataEntryHash[$Field].Text = "83 NOS"} #Used to be Langley AFB, but is now NOSC for the logon script.
                "Error|Groups" {$OffsetY += $ComboBuffer;continue main}
                }

            $form.Controls.Add($DataEntryHash[$Field]) 

            $OffsetY += $TextBoxHeight + $ComboBuffer
            }
        }

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(75,$OffsetY)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150,$OffsetY)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    #Dynamically decide how long our window will be
    #I honestly don't remember my logic behind the additions, but it works
    $form.Size = New-Object System.Drawing.Size(320,($OffsetY + $TextBoxHeight*4)) 

    #Make sure our GUI is up in yo face
    $form.Topmost = $True

    #Here we would prompt for input, but we moved it inside our data collection loop
    #if ($form.ShowDialog() -eq "Cancel") {exit}
    
    #This runs everytime showDialog runs, regardless of where it was added
    $form.Add_Shown({
        #Sets keyboard focus to the current field
        $DataEntryHash[$field].focus()
        #Makes our label Red and Bold
        $LabelHash[$field].ForeColor = "Red"
        $LabelHash[$field].Font = New-Object Drawing.Font($LabelHash[$field].Font.FontFamily, $LabelHash[$field].Font.Size, [Drawing.FontStyle]::Bold)
        })
    #This runs on form close
    $form.Add_FormClosing({
        #Now that we have the data we want, we can stop yelling
        if ($field -ne "Error") {
            $LabelHash[$field].ForeColor = "Black"
            $LabelHash[$field].Font = New-Object Drawing.Font($LabelHash[$field].Font.FontFamily, $LabelHash[$field].Font.Size, [Drawing.FontStyle]::Regular)
            }
        })

    #Get our data, including error checking
    #This giant loop is in case our user is stupid and changes data after we did our error checking on that field
    #Everytime the GUI gets brought back up, we will re-get all entered data
    do {
        #This is our exit flag.  Lets hope no one sullies it
        if ($form.ShowDialog() -eq "Cancel") {exit}
        $success = $true
        
        foreach ($field in $Fields) {
            #Error isn't a data field
            #Rank will be handled under Personnel Type
            #Groups is just a label, no data entry
            if ($field -eq "Error" -or $field -match "Rank" -or $field -eq "Groups") {continue}
            #Special case for our dropdowns
            elseif ($Field -match "Generation Qualifier") {
                $GenerationQualifier = $DataEntryHash[$Field].SelectedItem
                if ($GenerationQualifier -eq "(none)") {$GenerationQualifier = $null}
                }
            elseif ($Field -match "Personnel Type") {
                $PersonnelType = $DataEntryHash[$Field].Text
                
                #We set the error text each time anyway.  They won't notice unless the GUI comes back up asking for more/correct info
                $LabelHash["Error"].text = "Select the user's Personnel Type"
                
                #Loop until we get data
                if ($PersonnelType -eq "") {
                    #if ($form.ShowDialog() -eq "Cancel") {exit}
                    $success = $false
                    break
                    #$PersonnelType = $DataEntryHash[$Field].Text
                    }

                #Here we set our Abbreviations and PCC Codes
                #the condition variable is for our error checking loop
                #I dont care enough about military to code the ranks
                switch ($PersonnelType) {
                    "Military" {
                        $PCC = "A"
                        $payPlan = "ME"
                        $perTitle = $DataEntryHash["Rank (Leave blank if N/A)"].text
                        $DisplayRank = $perTitle
                        $condition = '$perTitle -eq ""'
                        break
                        }
                    "Reservist" {
                        $PCC = "V"
                        $payPlan = "ME"
                        $perTitle = $DataEntryHash["Rank (Leave blank if N/A)"].text
                        $DisplayRank = $perTitle
                        $condition = '$perTitle -eq ""'
                        break
                        }
                    "Civilian" {
                        $PCC = "C"
                        $payPlan = "GS"
                        $perTitle = $DataEntryHash["Rank (Leave blank if N/A)"].text
                        if ($perTitle -match "GS-[0-9]{2}.{0}$") {$payGrade = $perTitle.split("-")[1]}
                        $DisplayRank = "CIV"
                        $condition = '$perTitle -eq "" -and $perTitle -notmatch "GS-[0-9]{2}.{0}$"' #We need to know which GS rank the user is
                        break
                        }
                    "Contractor" {
                        $PCC = "E"
                        $payPlan = "99"
                        $payGrade = "00"
                        $perTitle = "CTR"
                        $DisplayRank = $perTitle
                        $condition = '$false'
                        break
                        }
                    }
                #Now we check for rank
                $field = "Rank (Leave blank if N/A)"
                $LabelHash["Error"].text = "Enter the user's Rank"
                
                if (Invoke-Expression $condition) {
                    #if ($form.ShowDialog() -eq "Cancel") {exit}
                    $success = $false
                    break
                    #$perTitle = $DataEntryHash["Rank (Leave blank if N/A)"].text
                    }
                }
            #We need to ensure there are actually accounts selected to create
            #Make them hit cancel to exit the script without doing anything
            elseif ($Field -match "Accounts to Create") {
                $LabelHash["Error"].text = "Select Accounts to create"
                
                if ($DataEntryHash[$Field].SelectedItems.count -eq 0) {
                    #Exit condition, in case we don't actually want accounts
                    #if ($form.ShowDialog() -eq "Cancel") {exit}
                    $success = $false
                    break
                    }
                }
            #Get the data from out textboxes
            else {
                #For sake of reducing lines and improving readability, we will use invoke-expression to assign our variables and define the error conditions
                switch -regex ($field) {
                    "Last" {
                        #$assignment = '$LastName = (Get-Culture).textinfo.totitlecase(($DataEntryHash[$field].text).toLower()).trim()'
                        $assignment = '$LastName = $DataEntryHash[$field].text.toUpper().trim()'
                        $errorCondition = '$Lastname -eq ""'
                        $LabelHash["Error"].text = "Enter the user's Last Name"
                        break
                        }
                    "First" {
                        #$assignment = '$FirstName = (Get-Culture).textinfo.totitlecase(($DataEntryHash[$field].text).toLower()).trim()'
                        $assignment = '$FirstName = $DataEntryHash[$field].text.toUpper().trim()'
                        $errorCondition = '$FirstName -eq ""'
                        $LabelHash["Error"].text = "Enter the user's First Name"
                        break
                        }
                    "Middle" {
                        $assignment = '$MiddleInitial = ($DataEntryHash[$field].text).toUpper().trim()'
                        $errorCondition = '!($MiddleInitial -match "\w" -or $MiddleInitial -eq "")'
                        $LabelHash["Error"].text = "Enter the user's proper Middle Initial (Leave blank if N/A)"
                        break
                        }
                    "^EDIPI$" {
                        $assignment = '$EDIPI = $DataEntryHash[$field].text'
                        $errorCondition = '$EDIPI -notmatch "1\d{9}"'
                        $LabelHash["Error"].text = "Enter the user's EDIPI"
                        break
                        }
                    "Branch" {
                        $assignment = '$Branch = ($DataEntryHash[$field].text).trim().ToUpper()'
                        $errorCondition = '$Branch -eq ""'
                        $LabelHash["Error"].text = "Enter the user's Branch"
                        break
                        }
                    "MAJCOM" {
                        $assignment = '$MAJCOM = ($DataEntryHash[$field].text).trim().ToUpper()'
                        $errorCondition = '$MAJCOM -eq ""'
                        $LabelHash["Error"].text = "Enter the user's MAJCOM"
                        break
                        }
                    "Organization" {
                        $assignment = '$Organization = ($DataEntryHash[$field].text).trim().ToUpper()'
                        $errorCondition = '$Organization -eq ""'
                        $LabelHash["Error"].text = "Enter the user's Organization"
                        break
                        }
                    "Office Symbol" {
                        $assignment = '$Office = ($DataEntryHash[$field].text).trim().toUpper()'
                        $errorCondition = '$Office -eq ""'
                        $LabelHash["Error"].text = "Enter the user's Office Symbol"
                        break
                        }
                    "Address" {
                        $assignment = '$Address = ($DataEntryHash[$field].text).trim()'
                        #We don't care what is or isn't in here
                        $errorCondition = '$false'
                        break
                        }
                    "City" {
                        $assignment = '$City = ($DataEntryHash[$field].text).trim()'
                        #Hard to enforce "[A-Z]? AFB" since not every place is an AFB 
                        $errorCondition = '$false'
                        break
                        }
                    "DSN Number" {
                        $assignment = '$telephone = $DataEntryHash[$field].text'
                        #I got Regex drunk and coded this, even though its not a required field.
                        #$errorCondition = '$telephone -notmatch "^(\d{3}-)??\d{3}-\d{4}.{0}$"'
                        #$LabelHash["Error"].text = "Enter a proper DSN Number for the user"

                        $errorCondition = '$false'
                        break
                        }
                    "E-mail" {
                        $assignment = '$mail = $DataEntryHash[$field].text'
                        $errorCondition = '$mail -notmatch "@mail\.smil\.mil$"'
                        $LabelHash["Error"].text = "Entered email address is not formatted properly"
                        break
                        }
                    "Remedy Ticket Numer" {
                        $assignment = '$INC = $DataEntryHash[$field].text'
                        $errorCondition = '$INC -notmatch "^INC(0{4})??2[0-9]{7}$"'
                        $LabelHash["Error"].text = "Entered ticket number does not appear to be correct"
                        break
                        }
                    "IA EDIPI" {
                        $assignment = '$IAEDIPI = $DataEntryHash[$field].text'
                        $errorCondition = '$IAEDIPI -notmatch "1\d{9}"'
                        $LabelHash["Error"].text = "Enter IA's EDIPI"
                        break
                        }
                    default {
                        Write-Host "Error: $field not accounted for"
                        break
                        }
                    }
                #Do our initial assignment
                Invoke-Expression $assignment
                
                #Loop until we get what we want, how we want
                if (Invoke-Expression $errorCondition) {
                    #Highlight our text so it can be easily overwritten
                    #Can't be thrown into Add_Shown because we also have a listbox and a dropdown.
                    $DataEntryHash[$field].SelectionStart = 0;
                    $DataEntryHash[$field].SelectionLength = $DataEntryHash[$field].Text.Length
                    
                    #if ($form.ShowDialog() -eq "Cancel") {1}
                    $success = $false
                    break
                    #Invoke-Expression $assignment
                    }
                }
            }
        } until ($success)
        
    #Since we removed a couple fields from our GUI, make sure we set them here
    if (!$Branch) {$Branch = "USAF"}
    if (!$MAJCOM) {$MAJCOM = "AFSPC"}
    if (!$City) {$City = "83 NOS"}
    if ($PersonnelType -eq "Reservist") {$Organization = "51 NOS"}

    function Make-Account {
        #With scoping, most of these variables can be seen from outside this function
        #param ($ou,$sam,$displayName,$UPN,$LastName,$FirstName,$MiddleInitial,$EDIPI,$MAJCOM,$branch,$Organization,$Office,$telephone,$PersonnelType,$perTitle,$PCC,$address,$city)
        #We only need to be fed the items that aren't the same across all accounts
        #param ($ou,$sam,$displayName,$CN,$UPN,$outputStr = @())
        param ($type)
        <# Testing purposes
        $OU = $UserOU
        $sam = $UserSAM
        $upn = ($EDIPI + "." + $PCC + "@smil.mil")
        #>

        #Find which domain we need to communicate with to make the account for
        switch ($type) {
            {$_ -eq "ACC Admin(s)" -or $_ -eq "ACC User"} {
                $netBIOSDomainName = "ACC";
                $server = Get-ADDomainController -Server $ACC | select -First 1 -ExpandProperty HostName
                $NetBiosParentDomainName = "ACCROOT"
                $ParentServer = Get-ADDomainController -server $ACCROOT | select -First 1 -ExpandProperty HostName
                $AdminOU = $ACCAdminOU
                }
            "ACCROOT Admin(s)" {
                $netBIOSDomainName = "ACCROOT";
                $server = Get-ADDomainController -Server $ACCROOT | select -First 1 -ExpandProperty HostName
                $NetBiosParentDomainName = "ACCROOT"
                $ParentServer = Get-ADDomainController -server $ACCROOT | select -First 1 -ExpandProperty HostName
                $AdminOU = $ACCRootAdminOU
                }
            "AFMC Admin(s)" {
                $netBIOSDomainName = "SAFMC";
                $server = Get-ADDomainController -Server $AFMC | select -First 1 -ExpandProperty HostName
                $NetBiosParentDomainName = $netBIOSDomainName
                $ParentServer = $server
                $AdminOU = $AFMCAdminOU
                }
            "USAFE Admin(s)" {
                $netBIOSDomainName = "USAFE";
                $server = Get-ADDomainController -Server $USAFE | select -First 1 -ExpandProperty HostName
                $NetBiosParentDomainName = "USAFEROOT"
                $ParentServer = Get-ADDomainController -server $USAFEROOT | select -First 1 -ExpandProperty HostName
                $AdminOU = $USAFEAdminOU
                }
            "USAFERoot Admin(s)" {
                $netBIOSDomainName = "USAFEROOT";
                $server = Get-ADDomainController -Server $USAFEROOT | select -First 1 -ExpandProperty HostName
                $NetBiosParentDomainName = "USAFEROOT"
                $ParentServer = Get-ADDomainController -server $USAFEROOT | select -First 1 -ExpandProperty HostName
                $AdminOU = $USAFEAdminOU
                }
            "AETC Admin(s)" {
                $netBIOSDomainName = "AETC";
                $server = Get-ADDomainController -Server $AETC | select -First 1 -ExpandProperty HostName
                $NetBiosParentDomainName = "AETCROOT"
                $ParentServer = Get-ADDomainController -server $AETCROOT | select -First 1 -ExpandProperty HostName
                $AdminOU = $AETCAdminOU
                }
            "AETCRoot Admin(s)" {
                $netBIOSDomainName = "AETCROOT";
                $server = Get-ADDomainController -Server $AETCROOT | select -First 1 -ExpandProperty HostName
                $NetBiosParentDomainName = "AETCROOT"
                $ParentServer = Get-ADDomainController -server $AETCROOT | select -First 1 -ExpandProperty HostName
                $AdminOU = $AETCAdminOU
                }
            "AMC Admin(s)" {
                $netBIOSDomainName = "AMC-S";
                $server = Get-ADDomainController -Server $AMCHUB | select -First 1 -ExpandProperty HostName
                $NetBiosParentDomainName = "DS-S"
                $ParentServer = Get-ADDomainController -server $AMC | select -First 1 -ExpandProperty HostName
                $AdminOU = $AMCAdminOU
                }
            "AMCRoot Admin(s)" {
                $netBIOSDomainName = "DS-S";
                $server = Get-ADDomainController -Server $AMC | select -First 1 -ExpandProperty HostName
                $NetBiosParentDomainName = "DS-S"
                $ParentServer = Get-ADDomainController -server $AMC | select -First 1 -ExpandProperty HostName
                $AdminOU = $AMCAdminOU
                }
            "AFSPC Admin(s)" {
                $netBIOSDomainName = "AFSPC-S";
                $server = Get-ADDomainController -Server $AFSPC | select -First 1 -ExpandProperty HostName
                $NetBiosParentDomainName = "AFSPC-RT"
                $ParentServer = Get-ADDomainController -server $AFSPCROOT | select -First 1 -ExpandProperty HostName
                $AdminOU = $AFSPCAdminOU
                }
            "AFSPCRoot Admin(s)" {
                $netBIOSDomainName = "AFSPC-RT";
                $server = Get-ADDomainController -Server $AFSPCROOT | select -First 1 -ExpandProperty HostName
                $NetBiosParentDomainName = "AFSPC-RT"
                $ParentServer = Get-ADDomainController -server $AFSPCROOT | select -First 1 -ExpandProperty HostName
                $AdminOU = $AFSPCAdminOU
                }
            "PACAF Admin(s)" {
                $netBIOSDomainName = "SPACAF-ROOT";
                $server = Get-ADDomainController -Server $PACAF | select -First 1 -ExpandProperty HostName
                $NetBiosParentDomainName = "SPACAF-ROOT"
                $ParentServer = Get-ADDomainController -server $PACAF | select -First 1 -ExpandProperty HostName
                $AdminOU = $PACAFAdminOU
                }
            "AREA42 Admin(s)" {
                $netBIOSDomainName = "AREA42";
                $server = Get-ADDomainController -Server $AREA42 | select -First 1 -ExpandProperty HostName
                $NetBiosParentDomainName = "AFNOAPPS"
                $ParentServer = Get-ADDomainController -server $AFNOAPPS | select -First 1 -ExpandProperty HostName
                $AdminOU = $42AdminOU
                }
            default {return}
            }

        if ($type -match "Admin") {
            #Select admin account types to create
            $OffsetY = 20 #Pixels from the bottom of the previous element (including top border) to next element
    
            #Generate our GUI
            $admSelectorForm = New-Object System.Windows.Forms.Form
            $admSelectorForm.Text = "Select ADM(s) for $netBIOSDomainName"
            $admSelectorForm.StartPosition = "CenterScreen" 
                
            $admSelectorLabel = New-Object System.Windows.Forms.Label
            $admSelectorLabel.Location = New-Object System.Drawing.Point($OffsetX,$OffsetY)
            $admSelectorLabel.Text = "Select ADM account(s) to make for `n$netBIOSDomainName\$CNBase"
            #I like to make things complicated, sue me
            $MaxWidth = [System.Windows.Forms.TextRenderer]::MeasureText("Select ADM account(s) to make for", $font).width
            $width2 = [System.Windows.Forms.TextRenderer]::MeasureText("$netBIOSDomainName\$CNBase.", $font).width
            if ($width2 -gt $MaxWidth) {$MaxWidth = $width2}
            $admSelectorLabel.Size = New-Object System.Drawing.Point($MaxWidth,($CharHeight * 2))
            $admSelectorForm.Controls.Add($admSelectorLabel)
                
            $OffsetY += $LabelBuffer + ($CharHeight * 2)
            
            $admSelector = New-Object System.Windows.Forms.ListBox 
            $admSelector.Location = New-Object System.Drawing.Point($OffsetX,$OffsetY) 
            $admSelector.Size = New-Object System.Drawing.Size(260,$CharHeight) 
            $admSelector.SelectionMode = "MultiExtended"

            #Admin Roles NOS could have
            "ACDEFW".ToCharArray() | foreach {
                [void]$admSelector.Items.Add("AD$_")
                }
            
            $admSelector.Height = ($CharHeight + $LabelBuffer) * $admSelector.Items.count

            $admSelectorForm.Controls.Add($admSelector) 

            $OffsetY += $admSelector.Height + $ComboBuffer

            $OKButton = New-Object System.Windows.Forms.Button
            $OKButton.Location = New-Object System.Drawing.Point(75,$OffsetY)
            $OKButton.Size = New-Object System.Drawing.Size(75,23)
            $OKButton.Text = "OK"
            $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $admSelectorForm.AcceptButton = $OKButton
            $admSelectorForm.Controls.Add($OKButton)

            $CancelButton = New-Object System.Windows.Forms.Button
            $CancelButton.Location = New-Object System.Drawing.Point(150,$OffsetY)
            $CancelButton.Size = New-Object System.Drawing.Size(75,23)
            $CancelButton.Text = "Cancel"
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $admSelectorForm.CancelButton = $CancelButton
            $admSelectorForm.Controls.Add($CancelButton)

            $admSelectorForm.Size = New-Object System.Drawing.Size(300,($OffsetY + $TextBoxHeight*4))  
            #This is our exit flag.  Lets hope no one sullies it

            if ($admSelectorForm.ShowDialog() -eq "Cancel") {exit}
            $success = $true

            #Build an array of accounts to make
            [array]$Accounts = $admSelector.SelectedItems | foreach {
                New-Object PSObject -Property @{
                    ou = $AdminOU
                    sam = $EDIPI + $PCC + "." + $_
                    displayName = $displayNameBase + " " + $_
                    CN = $CNBase + "." + $_
                    UPN = $EDIPI + "." + $_.replace("AD","ADM") + "@smil.mil"
                    sensitive = $true
                    EA3 = "ADM"
                    gigID = $null
                    userprincipalname = $EDIPI + ".ADM" + $PCC + $_[2] + "@smil.mil"
                    employeeType = "Z"
                    EA7 = "Acct Validated $DateValidated by $RunningCN"
                    }
                }
            }
        else {
            [array]$Accounts = New-Object PSObject -Property @{
                ou = $UserOU
                sam = $EDIPI + $PCC
                displayName = $displayNameBase
                CN = $CNBase
                UPN = $EDIPI + "." + $PCC + "@smil.mil"
                sensitive = $false
                EA3 = $null
                gigID = $EDIPI + $PCC
                employeeType = $PCC
                userprincipalname = $EDIPI + "." + $PCC + "@smil.mil"
                EA7 = $null
                }
            }

        :accounts foreach ($account in $Accounts) {
            $attribs = @{
                'sn' = $LastName
                'givenName' = $FirstName
                'initials' = $MiddleInitial
                'generationQualifier' = $GenerationQualifier
                'personalTitle' = $perTitle
                'payPlan'= $payPlan
                'payGrade' = $payGrade
                'company' = $branch
                'department' = $MAJCOM
                'o' = $Organization
                'physicalDeliveryOfficeName' = $Office
                'employeeId' = $EDIPI
                'description' = "Validated $INC"
                'l' = $City
                'st' = "VA"
                'postalCode' = 23665
                'c'= "US"
                'extensionAttribute4' = "US"
                'extensionAttribute10' = "ACC"
                'telephoneNumber' = $telephone
                'displayname' = $account.displayname
                'StreetAddress'= $address
                'scriptPath' = $scrTxt.Text
                'extensionAttribute3' = $account.EA3
                'extensionAttribute7' = $account.EA7
                'extensionAttribute8' = $INC
                'employeeType' = $account.employeeType
                'mail' = $mail
                'gigID' = $account.gigID
                }

            #Remove $null keys from the hashtable.
            [array]$attribs.keys | foreach {
                if (($attribs[$_] -eq $null) -or ($attribs[$_] -eq "")){
                    $attribs.remove($_)
                    }
                }

            #Check to see if account exists first
            Remove-Variable newuser -EA SilentlyContinue
            $newUser = Get-ADUser -Server $server -LDAPFilter "(UserPrincipalName=$($account.UPN))" -Properties * -EA SilentlyContinue
            if ($newUSer -eq $null) {$newUser = Get-ADUser -Server $server -LDAPFilter "(samaccountname=$($account.sam))" -Properties * -EA SilentlyContinue}
        
            #If account exists, strip it of groups and reconfigure it
            #Since we already made our user validate the sam of the user, we're safe to strip it
            #TODO we need to add a check before to see if the admin sam is free or in use as well
            if ($newUser) { 
                #preserve the sam
                $sam = $newUser.samaccountname
                #Confirm If we REALLY want to modify this account
                $Title = "Strip Account"
                $Message = "Are you absolutely sure you want to turn`n" + $newUser.distinguishedname + "`ninto a shell account?"
                $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Yes"
                $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No","No"
                $options = [System.Management.Automation.Host.ChoiceDescription[]]($No,$Yes)
            
                if ($host.ui.PromptForChoice($title,$message,$options,0) -eq 0) {
                    Write-Host "Skipping $sam then; do it manually."
                    #Add to string to paste into remedy
                    "$netBIOSDomainName\$sam already exists, but user prompted to skip configuring it.`n" | Write-Output
                    continue accounts     
                    }
    
                #Strip the groups
                foreach ($group in $newUser.memberof) {
                    #Remove-ADGroupMember -Server $ParentServer -Identity $group -Members $newuser.distinguishedname -Confirm:$false
                    Remove-ADGroupMember -Server $Server -Identity $group -Members $newuser.distinguishedname -Confirm:$false
                    }

                #Change its attributes            
                try {
                    #Set-ADAccountPassword -Server $server -Identity $newuser -NewPassword (ConvertTo-SecureString -AsPlainText (Generate-Password) -force)
                    #$newuser = Set-ADUser -Server $server -Identity $sam -AccountNotDelegated $account.sensitive -CannotChangePassword $false -City $City -Company $branch -Department $MAJCOM -DisplayName $displayName -EmployeeID $EDIPI -GivenName $FirstName -Initials $MiddleInitial -Office $Office -OfficePhone $telephone -Organization $Organization -SmartcardLogonRequired $true -Surname $LastName -PassThru -EA stop}
                    $newuser = Set-ADUser -Identity $sam -SamAccountName $account.sam -UserPrincipalName $account.UserPrincipalName -Replace $attribs -AccountNotDelegated $account.sensitive -CannotChangePassword $false -SmartcardLogonRequired $true -Server $server -PassThru -EA STOP
                    
                    #It seems like it will complain about the UPN if you're setting it to what it already is (not unique for forest)
                    if ($newuser.UserPrincipalName -ne $account.upn) {$newuser = Set-ADUser -Server $server -Identity $newuser -UserPrincipalName $account.upn -PassThru}
                    $newUser = Rename-ADObject -Server $server -Identity $newuser -NewName $account.cn -PassThru -EA stop
                    $newUSer = Move-ADObject -Server $server -Identity $newuser -TargetPath $account.ou -PassThru -EA stop
                
                    #Add to string to paste into remedy
                    "$netBIOSDomainName\$($account.UserPrincipalName) already exists and has been reconfigured." | Write-Output
                    }
                catch {
                    write-host "Error reconfiguring accounts: $_"
                    write-host $attribs
                    return "Error reconfiguring accounts: $_"
                    }
                }
            else {
                #Make fresh
                try {
                    #$newuser = New-ADUser -Server $server -Path $ou -Name $CN -SamAccountName $sam -AccountNotDelegated $sensitive -CannotChangePassword $false -City $City -Company $branch -Department $MAJCOM -DisplayName $displayName -EmployeeID $EDIPI -GivenName $FirstName -Initials $MiddleInitial -Office $Office -OfficePhone $telephone -Organization $Organization -SmartcardLogonRequired $true -StreetAddress $Address -Surname $LastName -UserPrincipalName $UPN -PassThru -EA stop
                    $newuser = New-ADUser -Server $server -Name $account.CN -SamAccountName $account.sam -UserPrincipalName $account.UserPrincipalName -Path $account.OU -AccountPassword (ConvertTo-SecureString -AsPlainText (Generate-Password) -force) -SmartcardLogonRequired $true -ChangePasswordAtLogon $true -Enabled $true -AccountNotDelegated $account.sensitive -OtherAttributes $attribs -PassThru -EA STOP #-WhatIf
                    }
                catch {
                    $msg = "Error creating new account: $($_.exception.message)"
                    write-host $msg
                    write-host $attribs
                    return $msg
                    }
                #Add to string to paste into remedy
                "$netBIOSDomainName\$($account.UserPrincipalName) has been created" | Write-Output
                }

            #GUI to add new groups
            #OffsetY is the only variable we change when building previous GUIs, so its the only one we have to reset
            $OffsetY = 20

            $Label = New-Object System.Windows.Forms.Label

            $Label.Location = New-Object System.Drawing.Point($OffsetX,$OffsetY) 
            $Label.Size = New-Object System.Drawing.Size(280,($CharHeight * 4)) 
            $Label.Text = "Paste in the list of groups to add to $netBIOSDomainName\$($account.sam).`nCan handle groups separated on each line, by commas, or by semi-colons."
    
            #Generate our GUI
            $form = New-Object System.Windows.Forms.Form 
            $form.Text = "Enter Groups"
            $form.StartPosition = "CenterScreen"

            $form.Controls.Add($Label) 

            $OffsetY += $LabelBuffer + ($CharHeight * 4)

            $DataEntry = New-Object System.Windows.Forms.RichTextBox
            $DataEntry.Location = New-Object System.Drawing.Point($OffsetX,$OffsetY) 
            $DataEntry.Size = New-Object System.Drawing.Size(260,($TextBoxHeight * 20))
            $DataEntry.Multiline = $true
            $DataEntry.ScrollBars = 3
            $DataEntry.ShortcutsEnabled = $true

            $form.Controls.Add($DataEntry) 

            $OffsetY += $TextBoxHeight*20 + $ComboBuffer

            $OKButton = New-Object System.Windows.Forms.Button
            $OKButton.Location = New-Object System.Drawing.Point(75,$OffsetY)
            $OKButton.Size = New-Object System.Drawing.Size(75,23)
            $OKButton.Text = "OK"
            $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.AcceptButton = $OKButton
            $form.Controls.Add($OKButton)

            $CancelButton = New-Object System.Windows.Forms.Button
            $CancelButton.Location = New-Object System.Drawing.Point(150,$OffsetY)
            $CancelButton.Size = New-Object System.Drawing.Size(75,23)
            $CancelButton.Text = "Skip"
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.CancelButton = $CancelButton
            $form.Controls.Add($CancelButton)

            $form.Size = New-Object System.Drawing.Size(300,($OffsetY + $TextBoxHeight*4)) 

            #Make sure our GUI is up in yo face
            $form.Topmost = $True

            if ($form.ShowDialog() -eq "OK") {
                $grpArray = $DataEntry.Text.split(',') | foreach {$_.trim().split(';') | foreach {$_.trim().split("`n") | foreach {$_.trim()}}} | Where {$_ -ne "" -and $_ -ne "Domain Users"}
                foreach ($grp in $grpArray ) {
                    $grp = $grp.trim()
                    $group = @()
                    $group += Get-ADGroup -Server $server -Filter {samaccountname -eq $grp -or CN -eq $grp -or DisplayName -eq $grp} -Properties CN,DisplayName
                    if ($parentserver -ne $server -and $group.count -eq 0) {$group += Get-ADGroup -Server $ParentServer -Filter {samaccountname -eq $grp -or CN -eq $grp -or DisplayName -eq $grp} -Properties CN,DisplayName}
                    switch ($group.count) {
                        0 {
                            #Add to string to paste into remedy
                            "Cannot add `"" + ($grp.trim()) + "`": Group does not exist in the $NetBiosParentDomainName Forest" | Write-Output
                            }
                        1 {
                            #We have a flaw.  We are assuming if we are in a parent domain, we are still in the same forest as the user.  I doubt this will be an issue though.
                            $groupDN = $group[0].DistinguishedName
                            if ($groupDN -like "*$ACCROOTDN" -and $groupDN -notlike "*$ACCDN") {
                                if ($accRootCreds -or $Area42Creds -or $AfnoCreds) {
                                    Add-ADGroupMember -Server $ParentServer $groupDN -Members $newUser
                                    "Account added to `"$($group[0].Name)`"" | Write-Output
                                    }
                                else {
                                    #$outputStr += "Use ACCROOT or AREA42 rights to add ACCROOT\$($newuser.distinguishedname.split(",")[0].split("=")[1]) to $grp manually." #This messes up for any CN with a comma in it
                                    "Use ACCROOT or AREA42 rights to add ACCROOT\$($newuser.distinguishedname.split("=")[1].TrimEnd(",OU").TrimStart("OU=")) to $grp manually." | Write-Output
                                    }
                                }
                            elseif ($groupDN -like "*$USAFEROOTDN" -and $groupDN -notlike "*$USAFEDN") {
                                if ($USAFERootCreds -or $Area42Creds -or $AfnoCreds) {
                                    Add-ADGroupMember -Server $ParentServer $groupDN -Members $newUser
                                    "Account added to `"$($group[0].Name)`"" | write-output
                                    }
                                else {
                                    #$outputStr += "Use ACCROOT or AREA42 rights to add USAFEROOT\$($newuser.distinguishedname.split(",")[0].split("=")[1]) to $grp manually." #This messes up for any CN with a comma in it
                                    "Use USAFEROOT or AREA42 rights to add USAFEROOT\$($newuser.distinguishedname.split("=")[1].TrimEnd(",OU").TrimStart("OU=")) to $grp manually." | Write-Output
                                    }
                                }
                            elseif ($groupDN -like "*$AETCROOTDN" -and $groupDN -notlike "*$AETCDN") {
                                if ($aetcRootCreds -or $Area42Creds -or $AfnoCreds) {
                                    Add-ADGroupMember -Server $ParentServer $groupDN -Members $newUser
                                    "Account added to `"$($group[0].Name)`"" | write-output
                                    }
                                else {
                                    #$outputStr += "Use ACCROOT or AREA42 rights to add USAFEROOT\$($newuser.distinguishedname.split(",")[0].split("=")[1]) to $grp manually." #This messes up for any CN with a comma in it
                                    "Use AMCROOT or AREA42 rights to add AETCROOT\$($newuser.distinguishedname.split("=")[1].TrimEnd(",OU").TrimStart("OU=")) to $grp manually." | Write-Output
                                    }
                                }
                            elseif ($groupDN -like "*$AMCROOTDN" -and $groupDN -notlike "*$AMCDN") {
                                if ($amcRootCreds -or $Area42Creds -or $AfnoCreds) {
                                    Add-ADGroupMember -Server $ParentServer $groupDN -Members $newUser
                                    "Account added to `"$($group[0].Name)`"" | write-output
                                    }
                                else {
                                    #$outputStr += "Use ACCROOT or AREA42 rights to add USAFEROOT\$($newuser.distinguishedname.split(",")[0].split("=")[1]) to $grp manually." #This messes up for any CN with a comma in it
                                    "Use AMCROOT or AREA42 rights to add AMC-S\$($newuser.distinguishedname.split("=")[1].TrimEnd(",OU").TrimStart("OU=")) to $grp manually." | Write-Output
                                    }
                                }
                            elseif ($groupDN -like "*$AFSPCROOTDN" -and $groupDN -notlike "*$AFSPCDN") {
                                if ($afspcRootCreds -or $Area42Creds -or $AfnoCreds) {
                                    Add-ADGroupMember -Server $ParentServer $groupDN -Members $newUser
                                    "Account added to `"$($group[0].Name)`"" | write-output
                                    }
                                else {
                                    #$outputStr += "Use ACCROOT or AREA42 rights to add USAFEROOT\$($newuser.distinguishedname.split(",")[0].split("=")[1]) to $grp manually." #This messes up for any CN with a comma in it
                                    "Use AMCROOT or AREA42 rights to add AMC-S\$($newuser.distinguishedname.split("=")[1].TrimEnd(",OU").TrimStart("OU=")) to $grp manually." | Write-Output
                                    }
                                }
                            elseif ($groupDN -like "*$AFNOAPPSDN" -and $groupDN -notlike "*$AREA42DN") {
                                if ($AfnoCreds) {
                                    Add-ADGroupMember -Server $ParentServer $groupDN -Members $newUser
                                    "Account added to `"$($group[0].Name)`"" | Write-Output
                                    }
                                else {
                                    #"Use AFNOAPPS rights to add AFNOAPPS\$($newuser.distinguishedname.split(",")[0].split("=")[1]) to $grp manually." #This messes up for any CN with a comma in it
                                    "Use AFNOAPPS rights to add AFNOAPPS\$($newuser.distinguishedname.split("=")[1].TrimEnd(",OU").TrimStart("OU=")) to $grp manually." | Write-Output
                                    }
                                }
                            else {
                                Add-ADGroupMember -Server $server $groupDN -Members $newUser
                                "Account added to `"$($group[0].Name)`"" | Write-Output
                                }
                            }
                        default {
                            $OffsetY = 20
                
                            $form_SelectAccount = New-Object System.Windows.Forms.Form 
                            $form_SelectAccount.Text = "Select Group"
                            $form_SelectAccount.StartPosition = "CenterScreen"
                
                            $Label = New-Object System.Windows.Forms.Label
                            $Label.Location = New-Object System.Drawing.Point($OffsetX,$OffsetY)
                            $Label.Text = "Select the proper Group (Hit cancel to add none)"
                            $form_SelectAccount.Controls.Add($Label)
                
                            $OffsetY += $LabelBuffer + $CharHeight

                            $ComboBox = new-object System.Windows.Forms.ComboBox
                            $ComboBox.Location = new-object System.Drawing.Size($OffsetX,$OffsetY)
                            $ComboBox.DropDownStyle = "DropDownList"

                            $OffsetY += $TextBoxHeight + $ComboBuffer
                
                            foreach ($str in $group) {
                                #If in ACCROOT forest
                                if ($str.DistinguishedName -like "*$ACCROOTDN") {
                                    #If not in ACC domain
                                    if ($str.DistinguishedName -notlike "*$ACCDN") {
                                        $domNTB = "ACCROOT"
                                        }
                                    else {$domNTB = "ACC"}
                                    }
                                #If in SAFMC forest
                                elseif ($str.DistinguishedName -like "*$AFMCDN") {$domNTB = "SAFMC"}
                                #If in AFNOAPPS forest
                                elseif ($str.DistinguishedName -like "*$AFNOAPPSDN") {
                                    #If not in AREA42 domain
                                    if ($str.DistinguishedName -notlike "*$AREA42DN") {
                                        $domNTB = "AFNOAPPS"
                                        }
                                    else {$domNTB = "AREA42"}
                                    }
                                #If in USAFEROOT forest
                                elseif ($str.DistinguishedName -like "*$USAFEROOTDN") {
                                    #If not in ACC domain
                                    if ($str.DistinguishedName -notlike "*$USAFEDN") {
                                        $domNTB = "USAFEROOT"
                                        }
                                    else {$domNTB = "USAFE"}
                                    }
                                #If in AETCROOT forest
                                elseif ($str.DistinguishedName -like "*$AETCROOTDN") {
                                    #If not in ACC domain
                                    if ($str.DistinguishedName -notlike "*$AETCDN") {
                                        $domNTB = "AETCROOT"
                                        }
                                    else {$domNTB = "AETC"}
                                    }
                                #If in AMC forest
                                elseif ($str.DistinguishedName -like "*$AMCROOTDN") {
                                    #If not in ACC domain
                                    if ($str.DistinguishedName -notlike "*$AMCDN") {
                                        $domNTB = "DS-S"
                                        }
                                    else {$domNTB = "AMC-S"}
                                    }
                                #If in AFSPC forest
                                elseif ($str.DistinguishedName -like "*$AFSPCROOTDN") {
                                    #If not in ACC domain
                                    if ($str.DistinguishedName -notlike "*$AFSPCDN") {
                                        $domNTB = "AFSPC-RT"
                                        }
                                    else {$domNTB = "AFSPC-S"}
                                    }
                                #If in PACAF forest
                                elseif ($str.DistinguishedName -like "*$PACAFDN") {$domNTB = "SPACAF-ROOT"}
                                else {Write-Host -ForegroundColor Red "ERROR: Unhandled domain.  "$($str.DistinguishedName)" will be skipped.";continue}

                                $ComboBox.Items.Add("$domNTB\$($str.CN)") | Out-Null
                                }

                            #Find the longest Entry to calculate label width
                            $MaxWidth = 0
                            foreach ($str in $ComboBox.Items) {
                                $StrPixelLength = [System.Windows.Forms.TextRenderer]::MeasureText($str, $font).width
                                if ($StrPixelLength -gt $MaxWidth) {$MaxWidth = $StrPixelLength}
                                }
                            $MaxWidth += $OffsetX * 3

                            $Label.Size = New-Object System.Drawing.Point($MaxWidth,$CharHeight) 
                            $ComboBox.Size = new-object System.Drawing.Size(($MaxWidth - $CharWidth * 2),$CharHeight)

                            $form_SelectAccount.Controls.Add($ComboBox)
                
                            $OKButton = New-Object System.Windows.Forms.Button
                            $OKButton.Location = New-Object System.Drawing.Point(75,$OffsetY)
                            $OKButton.Size = New-Object System.Drawing.Size(75,23)
                            $OKButton.Text = "OK"
                            $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                            $form_SelectAccount.AcceptButton = $OKButton
                            $form_SelectAccount.Controls.Add($OKButton)

                            $CancelButton = New-Object System.Windows.Forms.Button
                            $CancelButton.Location = New-Object System.Drawing.Point(150,$OffsetY)
                            $CancelButton.Size = New-Object System.Drawing.Size(75,23)
                            $CancelButton.Text = "Cancel"
                            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                            $form_SelectAccount.CancelButton = $CancelButton
                            $form_SelectAccount.Controls.Add($CancelButton)
                
                            #Dynamically decide how long our window will be
                            #I honestly don't remember my logic behind the additions, but it works
                            $form_SelectAccount.Size = New-Object System.Drawing.Size(($MaxWidth + ($OffsetX * 2)),($OffsetY + $TextBoxHeight*4)) 

                            #Make sure our GUI is up in yo face
                            $form_SelectAccount.Topmost = $True
                
                            if ($form_SelectAccount.ShowDialog() -eq "OK") {
                                if ($ComboBox.SelectedItem -like "ACCROOT\*") {
                                    if ($accRootCreds -or $Area42Creds) {
                                        Add-ADGroupMember -Server $ParentServer ($group[$ComboBox.SelectedIndex].DistinguishedName) -Members $newUser
                                        "Account added to `"$($group[0].Name)`"" | Write-Output
                                        }
                                    else {
                                        "Use ACCROOT or AREA42 rights to add $($newUser.samaccountname) to $($ComboBox.SelectedItem) manually." | Write-Output
                                        }
                                    }
                                elseif ($ComboBox.SelectedItem -like "USAFEROOT\*") {
                                    if ($usafeRootCreds -or $Area42Creds) {
                                        Add-ADGroupMember -Server $ParentServer ($group[$ComboBox.SelectedIndex].DistinguishedName) -Members $newUser | Write-Output
                                        "Account added to `"$($group[0].Name)`""
                                        }
                                    else {
                                        "Use USAFEROOT or AREA42 rights to add $($newUser.samaccountname) to $($ComboBox.SelectedItem) manually." | Write-Output
                                        }
                                    }
                                elseif ($ComboBox.SelectedItem -like "AFNOAPPS\*") {
                                    if ($AfnoCreds) {
                                        Add-ADGroupMember -Server $ParentServer ($group[$ComboBox.SelectedIndex].DistinguishedName) -Members $newUser
                                        "Account added to `"$($group[0].Name)`"" | Write-Output
                                        }
                                    else {
                                        "Use AFNOAPPS rights to add $($newUser.samaccountname) to $($ComboBox.SelectedItem) manually." | Write-Output
                                        }
                                    }
                                else {#Too lazy to check conditions for adding AFMC group to ACC account, etc.  Shouldnt happen ofc, but its best practice
                                    Add-ADGroupMember -Server $server ($group[$ComboBox.SelectedIndex].DistinguishedName) -Members $newUser
                                    "Account added to `"$($group[0].Name)`"" | Write-Output
                                    #$outputStr += "Added to `"$($ComboBox.SelectedItem)`""
                                    }
                                }
                            else {
                                #Add to string to paste into remedy
                                "Cannot add `"" + ($grp.trim()) + "`": Group does not exist in $NetBiosParentDomainName Forest`n" | Write-Output
                                } 
                            }
                        }
                    }
                }
            #This separates text so each account is its own block
            "" | Write-Output
            }
        }
        
    $UserSAM = $EDIPI + $PCC

    #Display Name is: Surname, GivenName Initials personalTitle Company Department o/physicalDeliveryOfficeName [ADM]
    $displayNameBase = "$LastName, $FirstName $MiddleInitial $GenerationQualifier $DisplayRank $Branch $MAJCOM $Organization/$Office".replace("   "," ").replace("  "," ") #In case no middle initial Or generationQualifier
    $CNBase = @($lastname,$Firstname,$MiddleInitial,$EDIPI,$PCC).where({($null -ne $_) -and ("" -ne $_)}) -join "."

    $AccountsSelected = $DataEntryHash["Accounts to Create (CTRL+Click to choose multiple)"].SelectedItems

    $outputStr = @()
    foreach ($account in $AccountsSelected) {
        $outputStr += Make-Account $account
        }

    $outputStr += "$mail is provisioned.  Routing to LRA for token."

    if ($outputStr -match "\.AD.") {$outputStr += "Please have Admin token OID marked Yes so it can be used with the AFPKI NPE portal."}
        
    Write-Host "Remedy notes copied to clipboard"
    $outputStr | clip.exe

    #Exit output
    $outputStr -match "1\d{9}" | Write-Host
    Read-Host -Prompt "Script finished for $EDIPI.`n`nPress Enter to continue."
        
    #Go again?
    $Title = "Go again?"
    $Message = "Process another user?"
    $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No","No"
    $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Yes"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($No,$Yes)
    $result = $host.ui.PromptForChoice($title,$message,$options,0)
    } until ($result -eq 0)
CLASSIFICATION: UNCLASSIFIED