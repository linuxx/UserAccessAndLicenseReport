
#download license map content
$strURL = "https://github.com/linuxx/UserAccessAndLicenseReport/raw/main/data/LicenseMap.txt"
Invoke-WebRequest -Uri $strURL -OutFile map.json

#convert to hash tables
$objLicenses = Get-Content -raw map.json | ConvertFrom-Json 
$objLicenseType = @{}
$objLicenses.psobject.properties | foreach{$objLicenseType[$_.Name]= $_.Value}

#connect to both services (this will prompt twice)
Write-Host $("Connecint to Exchange Online. Look for Password Prompt...") -ForegroundColor Cyan
Connect-ExchangeOnline
Write-Host $("Connecint to Microsoft Online for licenses. Look for Password Prompt...") -ForegroundColor Cyan
Connect-MsolService  

#get the mailboxes
$objMailboxes = Get-EXOMailbox | Select-Object UserPrincipalName, DisplayName, EmailAddresses
Write-Host $("Grabbed " + $objMailboxes.count + " mailboxes") -ForegroundColor Cyan

#create the output array
$objOutput = @()

#loop
foreach($objUser in $objMailboxes)
{
    Write-Host $("Processing: " + $objUser.UserPrincipalName) -ForegroundColor Cyan
    #our output object to add to the output array
    $objItem = New-Object psobject
    $objItem | Add-Member -type NoteProperty -Name 'DisplayName' -Value $objUser.DisplayName
    $objItem | Add-Member -type NoteProperty -Name 'PrimayEmail' -Value $objUser.UserPrincipalName

    $strEmails = "" #alias emails
    $strUsersWithAccess = "" #permissions to the mailbox

    #get aliases
    foreach($strMail in $objUser.EmailAddresses)
    {
        #if address not main mailbox
        if($strMail -cmatch 'smtp:') #case sensitive match
        {
            #if not onmicrosoft alias
            if($strMail -cnotmatch 'onmicrosoft.com') #case sensitive match
            {
                $strEmails += $strMail.Split(':')[1]
                $strEmails += ","
                #removes the smtp: from the address
            }
        }
    }
    if($strEmails.Length -gt 0) #if we have a trailing comma from above, remove it
    {
        $strEmails = $strEmails.Substring(0, $strEmails.Length - 1)
    }
    $objItem | Add-Member -type NoteProperty -Name 'EmailAliases' -Value $strEmails

    #permissions
    Write-Host $("Getting mailbox permissions...") -ForegroundColor Cyan
    $objPermissions = Get-EXOMailboxPermission -Identity $objUser.UserPrincipalName
    foreach($objPerm in $objPermissions)
    {
        #dont need SELF, etc.
        if($objPerm.User -match '@')
        {
            $strUsersWithAccess += $objPerm.User
            $strUsersWithAccess += ','
        }
    }
    if($strUsersWithAccess.Length -gt 0) #if we have a trailing comma from above, remove it
    {
        $strUsersWithAccess = $strUsersWithAccess.Substring(0, $strUsersWithAccess.Length - 1)
    }
    $objItem | Add-Member -type NoteProperty -Name 'UsersWithAccess' -Value $strUsersWithAccess

    #get licenses
    Write-Host $("Getting licenses...") -ForegroundColor Cyan
    $strLicense = ""
    $objUserLicenses = Get-MsolUser -UserPrincipalName $objUser.UserPrincipalName -ErrorAction SilentlyContinue | Select-Object Licenses
    foreach($objLicense in $objUserLicenses.Licenses)
    {
        #licenses are in the format of tenant:LICENSE
        $strLic = $objLicense.AccountSkuId.Split(":")
        $strLicense += $objLicenseType[$strLic[1]] #from the lookup table we download from github
        $strLicense += ','
    }
    if($strLicense.Length -gt 0) #if we have a trailing comma from above, remove it
    {
        $strLicense = $strLicense.Substring(0, $strLicense.Length - 1)
    }
    $objItem | Add-Member -type NoteProperty -Name 'Licenses' -Value $strLicense

    #add the item to the array
    $objOutput += $objItem
}

#prompt for save
Write-Host $("Done! Look for save dialog...") -ForegroundColor Cyan
$objSaveFileDialog = New-Object windows.forms.savefiledialog
$objSaveFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
$objSaveFileDialog.title = "Save User Data"
$objSaveFileDialog.filter = "CSV|*.csv|All Files|*.*" 
$objSaveFileDialog.ShowDialog()   

$objOutput | Export-Csv -NoTypeInformation -Path $objSaveFileDialog.FileName
