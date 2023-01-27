
#download license map content
$strURL = "https://raw.githubusercontent.com/linuxx/UserAccessAndLicenseReport/main/data/LicenseMap.txt"
$objLicenseMap = Invoke-WebRequest -Uri $strURL
$objLicenseType = $objLicenseMap.Content

#connect to both services
Connect-ExchangeOnline
Connect-MsolService  

$objMailboxes = Get-EXOMailbox | Select-Object UserPrincipalName, DisplayName, EmailAddresses

$objOutput = @()

foreach($objUser in $objMailboxes)
{
    $objItem = New-Object psobject
    $objItem | Add-Member -type NoteProperty -Name 'DisplayName' -Value $objUser.DisplayName
    $objItem | Add-Member -type NoteProperty -Name 'PrimayEmail' -Value $objUser.UserPrincipalName

    $strEmails = "" #alias emails
    $strUsersWithAccess = "" #permissions to the mailbox

    #get aliases
    foreach($strMail in $objUser.EmailAddresses)
    {
        if($strMail -cmatch 'smtp:') #case sensitive match
        {
            if($strMail -cnotmatch 'onmicrosoft.com') #case sensitive match
            {
                $strEmails += $strMail.Split(':')[1]
                $strEmails += ","

            }
        }
    }
    if($strEmails.Length -gt 0)
    {
        $strEmails = $strEmails.Substring(0, $strEmails.Length - 1)
    }
    $objItem | Add-Member -type NoteProperty -Name 'EmailAliases' -Value $strEmails

    #permissions
    $objPermissions = Get-EXOMailboxPermission -Identity $objUser.UserPrincipalName
    foreach($objPerm in $objPermissions)
    {
        if($objPerm.User -match '@')
        {
            $strUsersWithAccess += $objPerm.User
            $strUsersWithAccess += ','
        }
    }
    if($strUsersWithAccess.Length -gt 0)
    {
        $strUsersWithAccess = $strUsersWithAccess.Substring(0, $strUsersWithAccess.Length - 1)
    }
    $objItem | Add-Member -type NoteProperty -Name 'UsersWithAccess' -Value $strUsersWithAccess

    #get licenses
    $strLicense = ""
    $objUserLicenses = Get-MsolUser -UserPrincipalName $objUser.UserPrincipalName | Select-Object Licenses
    foreach($objLicense in $objUserLicenses.Licenses)
    {
        $strLic = $objLicense.AccountSkuId.Split(":")
        $strLicense += $objLicenseType[$strLic[1]]
        $strLicense += ','
    }
    if($strLicense.Length -gt 0)
    {
        $strLicense = $strLicense.Substring(0, $strLicense.Length - 1)
    }
    $objItem | Add-Member -type NoteProperty -Name 'Licenses' -Value $strLicense

    #add the item to the array
    $objOutput += $objItem
}

$objSaveFileDialog = New-Object windows.forms.savefiledialog
$objSaveFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
$objSaveFileDialog.title = "Save User Data"
$objSaveFileDialog.filter = "CSV|*.csv|All Files|*.*" 
$objSaveFileDialog.ShowDialog()   

$objOutput | Export-Csv -NoTypeInformation -Path $objSaveFileDialog.FileName
