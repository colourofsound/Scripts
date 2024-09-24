<#
NAME
Automate-365-Licensing

DESCRIPTION
This script automates the application of licensing using group membership. Whilst licenses can be applied to groups, and this is inherited to child groups, there are scenarios where
this doesn't work - in this particular script Dynamics licensing couldn't be applied simply at group level, as this would lead to users having licenses that duplicated function and therefore wasted money.
Some licenses are mutually exclusive; but not all of them. The logic in this script is specific to my organisation but can be adapted to suit.

AUTHOR
Chris Walker @colourofsound
#>

# Connect to MS Graph with Ad- account
Connect-mgGraph -Scopes User.ReadWrite.All, Organization.Read.All

# Lets ask for which group we're applying this to

Write-Host "==========================================================================================================================================" -ForegroundColor Magenta
Write-Host "          Please Enter the ObjectID for the group you'd like to apply Dynamics Licenses to. These can be found in Entra or Intune.        " -ForegroundColor DarkYellow
Write-Host "==========================================================================================================================================" -ForegroundColor Magenta
$groupinput = Read-Host "Group Object ID"
Write-Host "==========================================================================================================================================" -ForegroundColor Magenta
Write-Host " "

# Lets list the licenses we need first. Running [Connect-MgGraph] followed by [Get-MgSubscribedSku] gives us a list of licenses. These are the ones we need and their SKU:
$salesent = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'DYN365_ENTERPRISE_SALES' # Sales Enterprise License
$salespro = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'D365_SALES_PRO' # Sales Pro License
$custent = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'DYN365_ENTERPRISE_CUSTOMER_SERVICE' # Customer Service Enterprise License
$custpro = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'DYN365_CUSTOMER_SERVICE_PRO' # Customer Service Pro License
#$powerappsprem = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'POWERAPPS_PER_USER' # PowerApps Premium License
$custproatt = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'D365_CUSTOMER_SERVICE_PRO_ATTACH' # Customer Service Pro Attach License

# Next we need arrays of groups so we can assign licenses based on group membership. Get Dynamics groups with [Get-MgGroup -Filter "startswith(displayname, 'Dynamics')"]
# Populate this array with groups that need a Sales Enterprise License
$salesentgroup = @(
"12345" # Group name for reference
"12345" # Group name for reference
"12345" # Group name for reference
)
# Populate this array with groups that need a Sales Pro License
$salesprogroup = @(
"12345" # Group name for reference
"12345" # Group name for reference
"12345" # Group name for reference
)
# Populate this array with groups that need a Customer Enterprise License
$custentgroup = @(
"12345" # Group name for reference
"12345" # Group name for reference
"12345" # Group name for reference
)
# Populate this array with groups that need a Customer Pro License
$custprogroup = @(
"12345" # Group name for reference
"12345" # Group name for reference
"12345" # Group name for reference
)
# Populate this variable with a group that needs PowerApps Premium
$powerappspremgroup = "12345" # Group name for reference

# Below are a number of functions that help obtain group memberships including functions to populate from parent and child groups recursively. These were written with the help of ChatGPT. They have been tested and are working.

# Function to get all groups a user is a direct member of
function Get-DirectUserGroups {
    param (
        [string]$userId
    )
    
    $directGroups = @()

    try {
        # Get all groups the user is directly a member of
        $userGroups = Get-MgUserMemberOf -UserId $userId -All
        foreach ($group in $userGroups) {
            if ($group.AdditionalProperties['@odata.type'] -eq "#microsoft.graph.group") {
                $directGroups += $group.Id
            }
        }
    } catch {
        Write-Error "Error retrieving direct groups for user $($userId): $_"
    }

    return $directGroups
}

# Function to get all parent groups iteratively
function Get-AllParentGroups {
    param (
        [string[]]$groupIds
    )

    $allGroups = @()
    $processedGroups = @()  # To track which groups we've already processed
    $queue = [System.Collections.Generic.Queue[string]]::new()

    # Enqueue initial group IDs
    foreach ($groupId in $groupIds) {
        $queue.Enqueue($groupId)
    }

    while ($queue.Count -gt 0) {
        $currentGroupId = $queue.Dequeue()

        if ($processedGroups -contains $currentGroupId) {
            continue  # Skip already processed groups
        }

        $processedGroups += $currentGroupId
        $allGroups += $currentGroupId

        try {
            # Get parent groups (i.e., groups that contain this group)
            $parentGroups = Get-MgGroupMemberOf -GroupId $currentGroupId -All
            #Write-Output "Processing group $currentGroupId Found parent groups $($parentGroups.Id -join ', ')"
        } catch {
            Write-Error "Error retrieving parent groups for group $($currentGroupId): $_"
            continue  # Skip to the next group
        }

        foreach ($parentGroup in $parentGroups) {
            if ($parentGroup.AdditionalProperties['@odata.type'] -eq "#microsoft.graph.group") {
                if (-not ($processedGroups -contains $parentGroup.Id)) {
                    $queue.Enqueue($parentGroup.Id)
                }
            }
        }
    }

    return $allGroups
}

# Function to get all groups a user belongs to, including inherited groups
function Get-AllUserGroups {
    param (
        [string]$userId
    )
    
    $allUserGroups = @()

    try {
        # Step 1: Get the direct groups the user is a member of
        $directGroups = Get-DirectUserGroups -userId $userId
        #Write-Output "Direct groups for user $($userId): $($directGroups -join ', ')"

        # Step 2: Get all parent groups (inherited groups) of the user's groups
        $inheritedGroups = Get-AllParentGroups -groupIds $directGroups
        #Write-Output "Inherited groups for user $($userId): $($inheritedGroups -join ', ')"

        # Combine direct and inherited groups, removing duplicates
        $allUserGroups = $directGroups + $inheritedGroups | Select-Object -Unique
    } catch {
        Write-Error "Error retrieving all groups for user $($userId): $_"
    }

    return $allUserGroups
}

# Function to get all members of a group, including nested groups
function Get-AllGroupMembers {
    param (
        [string]$groupId
    )
    
    $allMembers = @()

    # Get members of the group
    $groupMembers = Get-MgGroupMember -GroupId $groupId -All

    foreach ($member in $groupMembers) {
        # Check if the member is a group or a user using AdditionalProperties
        $memberType = $member.AdditionalProperties['@odata.type']
        
        # If the member is a group (nested group), recurse into it
        if ($memberType -eq "#microsoft.graph.group") {
            $nestedGroupMembers = Get-AllGroupMembers -groupId $member.Id
            $allMembers += $nestedGroupMembers
        } 
        # If it's a user, add the user's ID to the list
        elseif ($memberType -eq "#microsoft.graph.user") {
            $userDetails = @{
                Id = $member.Id
                DisplayName = (Get-MgUser -UserId $member.Id).DisplayName
                UserPrincipalName = (Get-MgUser -UserId $member.Id).UserPrincipalName
            }
            $allMembers += $userDetails
        }
    }

    return $allMembers
}

# Function to compare two sets of groups (arrays) including nested groups
function Compare-Groups {
    param (
        [string[]]$group1,
        [string[]]$group2
    )

    try {
        # Find if there is any common element
        $commonGroups = $group1 | Where-Object { $group2 -contains $_ }

        if ($commonGroups.Count -gt 0) {
            return $true
        } else {
            return $false
        }
    } catch {
        Write-Error "Error comparing groups: $_"
        return $false
    }
}

# Populate an array of active users from the master group provided at the beginning of the script
$usergroup = $groupinput

# Get members of the group
$users = Get-AllGroupMembers -groupId $usergroup

# Apply the licensing logic to each user in that group. Formatting and color has been used to help legability in logs or terminal output.
foreach ($user in $Users) {

    # Declare User
    Write-Host " "
    Write-Host "==========================================================================================================================================" -ForegroundColor Magenta 
    Write-Host "== Processing" $user.DisplayName "("$user.UserPrincipalName") ==" -ForegroundColor DarkYellow
    Write-Host "==========================================================================" -ForegroundColor White

    # Store group membership in a variable
    $groups = Get-AllUserGroups -UserId $user.id

    # Store existing licenses in variable
    $licenses = Get-MgUserLicenseDetail -UserId $user.id
    
    Write-Host "Assessing Sales License Requirements" -ForegroundColor Blue
    Write-Host "====================================" -ForegroundColor DarkGray

    # Does the user need a Sales License? Assess if they belong to the requisite groups:
    if(Compare-Groups -group1 $groups -group2 $salesentgroup){
        Write-Host "Qualifies for Sales Enterprise License" -ForegroundColor Green
        # Assign Sales Enterprise and Remove Sales Pro
        Set-MgUserLicense -UserId $user.id -AddLicenses @() -RemoveLicenses @($custpro.SkuId) -ErrorAction SilentlyContinue
        Set-MgUserLicense -UserId $user.id -AddLicenses @() -RemoveLicenses @($salespro.SkuId) -ErrorAction SilentlyContinue
        Set-MgUserLicense -UserId $user.id -AddLicenses @{SkuId = $salesent.SkuId} -RemoveLicenses @()
    }else{
        Write-Host "Doesn't qualify for Sales Enterprise License" -ForegroundColor Red
        if(Compare-Groups -group1 $groups -group2 $salesprogroup){
            Write-Host "Qualifies for Sales Pro License" -ForegroundColor Green
            # Assign Sales Pro; Remove Sales Enterprise
            Set-MgUserLicense -UserId $user.id -AddLicenses @() -RemoveLicenses @($custpro.SkuId) -ErrorAction SilentlyContinue
            Set-MgUserLicense -UserId $user.id -AddLicenses @() -RemoveLicenses @($salesent.SkuId) -ErrorAction SilentlyContinue
            Set-MgUserLicense -UserId $user.id -AddLicenses @{SkuId = $salespro.SkuId} -RemoveLicenses @()
        }else{
            Write-Host "Doesn't qualify for Sales Pro License" -ForegroundColor Red
        }
    }
    
    # Grab new license status
    $licenses = Get-MgUserLicenseDetail -UserId $user.id

    Write-Host " "
    Write-Host "Assessing Customer Service License Requirements" -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor DarkGray

    # Does the user need a Customer Service License? Assess if they belong to the requisite groups:
    if(Compare-Groups -group1 $groups -group2 $custentgroup){
        Write-Host "Qualifies for Customer Service Enterprise License" -ForegroundColor DarkGreen
        # Assign Customer Service Enterprise; Remove Customer Service Pro/Pro Attach
        Set-MgUserLicense -UserId $user.id -AddLicenses @() -RemoveLicenses @($custproatt.SkuId),@($custpro.SkuId) -ErrorAction SilentlyContinue
        Set-MgUserLicense -UserId $user.id -AddLicenses @{SkuId = $custent.SkuId} -RemoveLicenses @()
    }else{
        Write-Host "Doesn't qualify for Customer Service Enterprise License" -ForegroundColor DarkRed
        if(Compare-Groups -group1 $groups -group2 $custprogroup){          
            if (($licenses.SkuPartNumber -match "DYN365_ENTERPRISE_SALES") -or ($licenses.SkuPartNumber -match "D365_SALES_PRO")){
                Write-Host "Qualifies for Customer Service Pro Attach License" -ForegroundColor DarkGreen
                # Assign Customer Service Pro Attach; Remove Customer Service Pro
                Set-MgUserLicense -UserId $user.id -AddLicenses @() -RemoveLicenses @($custpro.SkuId) -ErrorAction SilentlyContinue
                Set-MgUserLicense -UserId $user.id -AddLicenses @{SkuId = $custproatt.SkuId} -RemoveLicenses @()
            }else{
                Write-Host "Qualifies for Customer Service Pro License" -ForegroundColor DarkGreen
                # Assign Customer Service Pro; Remove Pro Attach
                Set-MgUserLicense -UserId $user.id -AddLicenses @() -RemoveLicenses @($salespro.SkuId) -ErrorAction SilentlyContinue
                Set-MgUserLicense -UserId $user.id -AddLicenses @() -RemoveLicenses @($salesent.SkuId) -ErrorAction SilentlyContinue
                Set-MgUserLicense -UserId $user.id -AddLicenses @() -RemoveLicenses @($custproatt.SkuId) -ErrorAction SilentlyContinue
                Set-MgUserLicense -UserId $user.id -AddLicenses @{SkuId = $custpro.SkuId} -RemoveLicenses @()
            }
        }else{
            Write-Host "Doesn't qualify for Customer Service Pro Licenses" -ForegroundColor DarkRed
            if (($licenses.SkuPartNumber -match "DYN365_ENTERPRISE_SALES") -or ($licenses.SkuPartNumber -match "D365_SALES_PRO")){
                Write-Host "Qualifies for Customer Service Pro Attach License" -ForegroundColor DarkGreen
                # Assign Customer Service Pro Attach; Remove Customer Service Pro
                Set-MgUserLicense -UserId $user.id -AddLicenses @() -RemoveLicenses @($custpro.SkuId) -ErrorAction SilentlyContinue
                Set-MgUserLicense -UserId $user.id -AddLicenses @{SkuId = $custproatt.SkuId} -RemoveLicenses @()
            }else{
                Write-Host "Doesn't qualify for Customer Service Pro Attach License" -ForegroundColor DarkRed
            }
       }
    }

    Write-Host " "
    Write-Host "Assessing PowerApps Premium License Requirements" -ForegroundColor White
    Write-Host "================================================" -ForegroundColor DarkGray

    # What about a PowerApps Premium License?            
    if (($licenses.SkuPartNumber -match "DYN365_ENTERPRISE_SALES") -or ($licenses.SkuPartNumber -match "DYN365_ENTERPRISE_CUSTOMER_SERVICE")-or ($licenses.SkuPartNumber -match "POWERAPPS_PER_USER")){
        Write-Host "Doesn't qualify for PowerApps Premium - may be assigned manually" -ForegroundColor Red
    }else{
        Write-Host "Qualifies for Power Apps Premium" -ForegroundColor Green
        New-MgGroupMember -GroupId $powerappspremgroup -DirectoryObjectId $user.id 
    }
}


Write-Host " "
Write-Host "=====================================" -ForegroundColor DarkGray 
Write-Host "== All Done - Signing out of Graph ==" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor DarkGray

# Sign out of Microsoft Graph
Disconnect-MgGraph
