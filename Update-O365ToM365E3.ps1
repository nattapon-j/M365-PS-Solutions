# Import the Microsoft Graph module
Import-Module Microsoft.Graph

# Connect to Microsoft 365
Connect-MgGraph -Scopes "User.ReadWrite.All"

# Define the SKUs
# Ref: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference
# $oldLicense = "ENTERPRISEPACK"  # Office 365 E3
$oldLicense = "Office_365_E3_(no_Teams)"  # Office 365 E3
# $newLicense = "SPE_E3"  # Microsoft 365 E3
$newLicense = "Microsoft_365_E3_(no_Teams)"  # Microsoft 365 E3

# Get all users
Write-Host "Fetching all users..." -ForegroundColor Yellow
# $users = Get-MgUser -All -ExpandProperty "AssignedLicenses"
$o365e3Sku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq $oldLicense
$users = Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq $($o365E3sku.SkuId) )" -ConsistencyLevel eventual -CountVariable e3licensedUserCount -All
Write-Host "Found $e3licensedUserCount Office 365 E3 licensed users."

# Iterate through users
foreach ($user in $users) {
    $userPrincipalName = $user.UserPrincipalName
    Write-Host "Processing $userPrincipalName..." -ForegroundColor Yellow

    Write-Host "$userPrincipalName has the old license $oldLicense. Updating to $newLicense..." -ForegroundColor Green
        
    $m365e3Sku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq $newLicense
    try {
        # Add the new license
        Set-MgUserLicense -UserId $user.Id -AddLicenses @{skuId = $m365e3Sku.SkuId } -RemoveLicenses @()
        
        # Remove the old license
        Set-MgUserLicense -UserId $user.Id -RemoveLicenses @($o365E3sku.SkuId) -AddLicenses @{}

        Write-Host "$userPrincipalName has been successfully updated to $newLicense." -ForegroundColor Cyan
    }
    catch {
        Write-Host "Failed to update $userPrincipalName : $_" -ForegroundColor Red
    }
}

Write-Host "License update process completed." -ForegroundColor Green

# Disconnect from Microsoft 365
Disconnect-MgGraph