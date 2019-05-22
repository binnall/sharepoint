#Intro
 
# Script will take a csv file that contains url to SharePoint sites and analyse the site pages to see if any of the pages have hyperlinks.
# For every hyperlink in a page this gets output to a row in a csv that is delimited by a pipe

# Version that this script has been tested against.
#   Version     Name
#   -------     ----
#   3.3.1811.0  SharePointPnPPowerShellOnline

#Site CSV
#Each site in this list will have the script run against
$csv_SiteList = ".\WebPartsUsed\sites.csv"
$csv_siteheaders = 'Url'

# list of the web parts and there IDs used to match the href tags that are embeded in the canvas page content
$webpart_input_csv ='.\WebPartsUsed\webparts.csv'
$csv_webpartheaders = 'title', 'id'

#Date used in the file creation
$date = Get-Date
$date = $date.ToString("yyyymmddhhss")

#filename by using the date
$file_name = $date + 'WebPartsUsed.csv'

#Path to create the output file
$creation_path = ".\PowerShell\WebPartsUsed"

# The site pages list that this script will run against
$List = "SitePages"

# Headers for the output csv
$headers = "Site Title|Page Title|Page Url|Href Tag|Conversations Yammer|Her|Highlighted Content|Quick Links|Highlights Yammer|Image|News|Stream|Twitter|Events|Forms"

# new line character
$ofs = "`n"

# delimiter to use
$delim = '|'

# create object of all the sites
$sites = Import-Csv -Path $csv_SiteList -Header $csv_siteheaders

#webparts
$webparts = Import-CSV -Path $webpart_input_csv -header $csv_webpartheaders

#variable for the header
$csv_outputheader = $headers + $ofs

#complete file path
$csv_path = $creation_path + '/' + $file_name

# create output csv
New-Item -Path $creation_path -Name $file_name -ItemType File -Value $csv_outputheader

function CreateWebPartIDObj {
    [system.collections.ArrayList]$webpartArray = @()
    foreach($webpart in $webparts){
        $webpartobj = New-Object -TypeName webParts
        $webpartobj | Add-Member -MemberType NoteProperty -Name 'Title' -Value $webpart.title
        $webpartobj | Add-Member -MemberType NoteProperty -Name 'ID' -Value $webpart.id
        $webpartArray.Add($webpartobj)
        return $webpartArray
    }
}

# itterate around each site from the csv
foreach($site in $sites)
{
    # make the connection, get ome site information and create object that contains all the site pages
    $connection = Connect-PnPOnline -Url $site.Url -UseWebLogin
    $pnpsite = Get-PnPWeb -Connection $connection
    $site_title = $pnpsite.Title
    $pages = (Get-PnPListItem -List $List -Fields "CanvasContent1", "Title" -Connection $connection).FieldValues

    # itterate around each page in the stie to get the information from each page that will be used to build up the row and also conduct
    # the check to see if the canvas content has any href tags embeded
    foreach($page in $pages)
    {
        $page_title = $page.Get_Item("Title")
        $fileref = $page.Get_Item("FileRef")
        $canvascontent = $page.Get_Item("CanvasContent1")
        # check if the canvas has content 
        if ($canvascontent.Length -gt 0) 
        {
            $wpArray = CreateWebPartIDObj
            foreach($wp in $wpArray){
                $res = $canvascontent -match $wp.ID
            }
        }
    }
    $row = $site_title + $delim + $page_title + $delim + $fileref + $delim + 'conv yammer' + $delim + 'hero' + $delim + 'High content' + $delim + 'quick links' + $delim + 'high yammer' + $delim + 'image' + $delim + 'news' + $delim + 'stream' + $delim + 'twitter' + $delim + 'events' + $delim + 'forms'
    Add-Content -Path $csv_path -Value $row
    Disconnect-PnPOnline -Connection $connection
}

