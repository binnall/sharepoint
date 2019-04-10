# Variable that define what sites to replicate permissions in
$csv_SiteList = "C:\!Jack\SharePoint\PowerShell\ReplicateHubPermissions\sites.csv"
$SiteListHeaders = 'Url'

# Group Config file
# h1 = Group Name, h2 = Group Owner
$csv_GroupsTocreate = "C:\!Jack\SharePoint\PowerShell\ReplicateHubPermissions\groups.csv"
$GroupToCreateHeaders = 'GroupName','GroupOwner','GroupRole','GroupMembers'

# Get the credentials for the user
$creds = Get-credential

# delete the groups?
$var_DeleteGroups = "Yes"

#uses a file path to read a csv and return the csv as an object
# csvPath = file location
# csvHeaders = headers for this particualr file
function ReadCsv{
    param(
        [string] $csvPath,
        [string] $csvHeaders
    )
    $csv = Import-Csv $csvPath -Header $csvHeaders
    return $csv
}

<#
.SYNOPSIS
Gets all the permission groups for a site

.DESCRIPTION
Gets all the permission groups for a site

.PARAMETER connection
connection = connection to SharePoint

.EXAMPLE
$connection = Connect-PnPOnline -url 'site url' -credentials get-credentials
GetGroups -connection $connection

.NOTES
General notes
#>
function GetGroups{
    param([SharePointPnP.PowerShell.Commands.Base.SPOnlineConnection] $connection)
    $groups = Get-PnPGroup -Connection $connection
    return $groups
}

<#
.SYNOPSIS
Deletes the groups in the site

.DESCRIPTION
Deletes all the groups in the site

.PARAMETER connection
SharePoint connection using Connect-PnPOnline

.PARAMETER groupTitle
Title of the group to delete

.EXAMPLE
DeleteGroup -connection $connection -groupTitle "XX Members"

.NOTES
General notes
#>
function DeleteGroup{
    param(
        [SharePointPnP.PowerShell.Commands.Base.SPOnlineConnection] $connection,
        [string] $groupTitle)
    Remove-PnPGroup -Identity $groupTitle -Connection $connection -Force
}

<#
.SYNOPSIS
Create the new groups

.DESCRIPTION
Long description

.PARAMETER connection
Parameter description

.PARAMETER groupName
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function CreateGroup{
    param(
        [SharePointPnP.PowerShell.Commands.Base.SPOnlineConnection] $connection,
        [string] $groupname,
        [string] $groupowneremail
    )
    New-PnPGroup -Title $groupname -Owner $groupowneremail -Connection $connection
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER connection
Parameter description

.PARAMETER groupName
Parameter description

.PARAMETER roleToAdd
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function SetGroupPermissions {
    param(
        [SharePointPnP.PowerShell.Commands.Base.SPOnlineConnection] $connection,
        [string] $groupName,
        [string] $roleToAdd
    )
    Set-PnPGroupPermissions -Identity $groupName -AddRole $roleToAdd -Connection $connection
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER connection
Parameter description

.PARAMETER groupName
Parameter description

.PARAMETER accountEmail
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function AddUserToGroup{
    param(
        [SharePointPnP.PowerShell.Commands.Base.SPOnlineConnection] $connection,
        [string] $groupName,
        [string] $accountEmail
    )
    Add-PnPUserToGroup -LoginName $accountEmail -Identity $groupName
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER members
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function UsersForGroup{
    param([string] $members)
    $users = $members.Split('|')
    return $users
}

$sites = ReadCsv -csvPath $csv_SiteList -csvHeaders $SiteListHeaders
$newgroups = Import-Csv $csv_GroupsTocreate -Header $GroupToCreateHeaders

# Loop around each site to remove existing permissions
foreach($site in $sites)
{
    Write-Host "connecting to: " $site.Url
    $connection = Connect-PnPOnline -Url $site.Url -Credentials $creds
    $groups = GetGroups -connection $connection
    # delete the current groups on the site?
    # Groups don't have to be deleted
    if ($var_DeleteGroups = "Yes") {
        foreach($group in $groups)
        {
            Write-Host $group.Title
            DeleteGroup -groupTitle $group.Title -connection $connection
        }
    }

    # Loop and create new permission groups, assign permission role and add accounts to the group      
    foreach($newgroup in $newgroups)
        {
            Write-Host "Creating Group: " $newgroup.GroupName
            
            CreateGroup -groupName $newgroup.GroupName -groupOwnerEmail $newgroup.GroupOwner -connection $connection
            SetGroupPermissions -groupName $newgroup.GroupName -roleToAdd $newgroup.GroupRole
            $members = UsersForGroup -members $newgroup.GroupMembers
            # call method to split out and add each user to group individually 
            foreach($member in $members){
                Write-Host "Adding: " $member
                AddUserToGroup -groupName $newgroup.GroupName -accountEmail $member -connection $connection
            }
    }
    Disconnect-PnPOnline -Connection $connection
}