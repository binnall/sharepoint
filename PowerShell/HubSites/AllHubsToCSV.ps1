# Credentials for the user
$creds = Get-Credential

# O365 Admin Url
$adminurl = "https://jeb03-admin.sharepoint.com"

# Get date as of now inc time and secs
$date = Get-Date
$date = $date.ToString("yyyymmddhhss")

# Path to create the file in
$creation_path = "C:\!Jack\SharePoint\PowerShell\Hub Sites"
# Create unique file name
$file_name = $date + "AllHubSites.csv"

# headers for the csv file
$headers = 'Site Title, Site Url'

# new line character
$ofs = "`n"

# output delim
$delim = ','

#Path of the csv
$csv_path = $creation_path + '/' + $file_name

# Csv header content
$content = $headers + $ofs

New-Item -Path $creation_path -Name $file_name -ItemType File -Value $content

Connect-PnPOnline -Url $adminurl -UseWebLogin

$hubs = Get-PnPHubSite

foreach($hub in $hubs)
{
    $row = $hub.Title + $delim + $hub.SiteUrl
    Add-Content -Path $csv_path -Value $row
}
