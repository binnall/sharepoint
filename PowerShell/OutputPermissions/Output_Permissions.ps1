#sites
$csv_SiteList = "C:\!Jack\SharePoint\PowerShell\ReplicateHubPermissions\sites.csv"
$csv_siteheaders = 'Url'

#Store credentials
$creds = Get-Credential

# Path to create file in
$creation_path = "C:\!Jack\SharePoint\PowerShell\ReplicateHubPermissions"
$date = Get-Date
$date = $date.ToString("yyyymmddhhss")

$file_name = $date + "Permission Output.csv"

# Headers for the output csv
$headers = "Site Url, Group Name, Account"

# new line character
$ofs = "`n"

# output delim
$delim = ','

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
Short description

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
function GetGroupMembers{
    param(
            [SharePointPnP.PowerShell.Commands.Base.SPOnlineConnection] $connection,
            [string] $groupName
        )
        $members = Get-PnPGroupMembers -Identity $groupName -Connection $connection
        return $members
}

# function to get the site admins, will return a Client.User object
function GetSiteAdmins {
    Param([SharePointPnP.PowerShell.Commands.Base.SPOnlineConnection] $connection)
    $admins = Get-PnPSiteCollectionAdmin -Connection $connection
    return $admins
}

# will concat all users in a group to a string delimited by ;
function ReturnAllUsersString{
    param([Microsoft.SharePoint.Client.User] $users)
    $userlist = ""
    foreach($user in $users){
        if ($userlist -eq "") {
            $userlist = $user.Email
        }else{
            $userlist = $userlist + ";" + $user.Email
        }
    }
    return $userlist
}

# function to create a row in the output csv
function CreateRow{
    param(
        [string] $delim,
        [string] $siteurl,
        [string] $group,
        [string] $emailaccountlist
    )
    $row = $siteurl + $delim + $group + $delim + $emailaccountlist
    return $row
}

# Create some default content
$sites = Import-Csv -Path $csv_SiteList -Header $csv_siteheaders

$content = $headers + $ofs

New-Item -Path $creation_path -Name $file_name -ItemType File -Value $content

# itterate around each site and create a row
foreach($site in $sites)
{
    $connection = Connect-PnPOnline -Url $site.Url -Credentials $creds
    $groups = GetGroups -connection $connection
    $admins = GetSiteAdmins -connection $connection
    $adminslist = ReturnAllUsersString -users $admins
    $row = CreateRow -siteurl $site.Url -group "Admins" -emailaccountlist $adminslist -delim $delim
    $csv_path = $creation_path + '/' + $file_name
    $value = $row
    Add-Content -Path $csv_path -Value $value
    
    #foreach group in the site create a row in the csv so that it can be filtered / analaysed in seperate applications
    foreach($group in $groups)
    {       
        $members = GetGroupMembers -groupName $group.Title -connection $connection
        $users = ""
        foreach($member in $members)
        {
            if ($users -eq "")
            {
                $users = $member.Email
            }else {
                $users = $users + ';' + $member.Email
            }            
        }
        $row = CreateRow -siteurl $site.Url -group $group.Title -emailaccountlist $users -delim $delim
        $csv_path = $creation_path + '/' + $file_name
        Add-Content -Path $csv_path -Value $row
    }
}