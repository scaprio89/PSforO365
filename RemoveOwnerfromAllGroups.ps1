
#----------------------------------------------------------------------------------------------------------------------------
# Arrays for capturing the actions
$owned      = @()
$memberof   = @()

#----------------------------------------------------------------------------------------------------------------------------
# Get all of the Office 365 groups
$azgroups = Get-AzureADMSGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All:$true
Write-Output "$($azgroups.Count) Office 365 groups were found"

# Get info for departing user
$upn        = Read-Host "UserPrincipalName of user being removed from groups"
$AZuser     = Get-AzureADUser -ObjectId "cbb9020e-5a02-41f5-955c-ee43a0d54835"

# Get info for delegate
$delegate   = Read-Host "UserPrincipalName of user taking over group ownership"
$AZdelegate = Get-AzureADUser -ObjectId "877ed74d-a5f8-4e08-ae38-b49eeccf2a77"

# Check each group for the user
foreach ($group in $azgroups) {
    $members = (Get-AzureADGroupMember -ObjectId $group.id).UserPrincipalName
    If ($members -contains $upn) {
        Remove-AzureADGroupMember -ObjectId $group.Id -MemberId $AZuser.ObjectId 
        Write-Output "$upn was removed from $($group.DisplayName)"
        $memberof += $group

        $owners  = Get-AzureADGroupOwner -ObjectId $group.id
        foreach ($owner in $owners) {
            If ($upn -eq $owner.UserPrincipalName) {
                # Add a new owner to prevent orphaned
                If ($owners.count -lt 2) {
                Write-Output "$delegate was added as a new owner"
                Add-AzureADGroupOwner -ObjectId $group.Id -RefObjectId $AZdelegate.ObjectId
                }
                # Now we can remove the user
                Write-Output "$upn was removed as owner of $($group.DisplayName)"
                Remove-AzureADGroupOwner -ObjectId $group.Id -OwnerId $AZuser.ObjectId

                $owned += $group
            }
        }
    }
}

# Groups that the user owned:
Write-Output "$upn was removed as Owner of:"
$owned | Select-Object DisplayName, Id

#Groups that the user was a member of:
Write-Output "$upn was removed as Member of:"
$memberof | Select-Object DisplayName, Id