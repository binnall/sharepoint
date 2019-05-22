# Version that this script has been tested against.
#   Version     Name
#   -------     ----
#   3.3.1811.0  SharePointPnPPowerShellOnline

#Site CSV
#Each site in this list will have the script run against the site pages library
$csv_SiteList = "C:\GIT\sharepoint\PowerShell\WebPartsUsed\sites.csv"
$csv_siteheaders = 'Url'

# list of the web parts and there IDs used to match the href tags that are embeded in the canvas page content
$webpart_input_csv = 'C:\GIT\sharepoint\PowerShell\WebPartsUsed\webparts.csv'
$csv_webpartheaders = 'Title', 'Id'

#Date used in the filename for the creation of the output file
$date = Get-Date
$date = $date.ToString("yyyymmddhhss")

#filename by using the date
$file_name = $date + 'WebPartsUsed.csv'

#Path to create the output file
$creation_path = "C:\GIT\sharepoint\PowerShell\WebPartsUsed"

# The site pages list that this script will run against
$List = "SitePages"

# Headers for the output csv
$headers = "Site Title|Page Title|Page Url|Conversations Yammer|Hero|Highlighted Content|Quick Links|Highlights Yammer|Image|News|Stream|Twitter|Events|Forms|People"

# new line character
$ofs = "`n"

#Default match result when a webpart ID is matched to content in canvascontent1
$matchResult = "yes"

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

# create an object that will store the 
function CreateWebPartIDObj {
    [system.collections.ArrayList]$webpartArray = @()
    foreach($webpart in $webparts){
        $webpartobjprop = @{
            Title = $webpart.title            
            ID = $webpart.id
        }
        $obj = New-Object -typeName psobject -Property $webpartobjprop
        $webpartArray.Add($obj) | Out-Null
    }
    return $webpartArray
}

# Create an object that will be used to store if a match has been found for the web parts being checked for 
function CreateMatchObj{
    $def = "no"
    $wpOuput = @{
        ConversationsYammer = $def
        Hero = $def
        HighlightedContent = $def
        QuickLinks = $def
        HighlighsYammer = $def
        Image = $def
        News = $def
        Stream = $def
        Twitter = $def
        Events = $def
        Forms = $def
        People = $def
    }
    return $wpOuput
}

function OutputToCSV
{
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
        $matchObj = CreateMatchObj
        # check if the canvas has content 
        if ($canvascontent.Length -gt 0) 
        {
            foreach($wp in $webparts){
                if($canvascontent -match $wp.ID){
                    switch ($wp.Title) {
                        "Conversations Yammer" {$matchObj.ConversationsYammer = $matchResult; break}
                        "Hero" {$matchObj.Hero = $matchResult; break}
                        "Highlighted Content" {$matchObj.HighlightedContent = $matchResult; break}
                        "Quick Links" {$matchObj.QuickLinks = $matchResult; break}
                        "Highlighs Yammer" {$matchObj.HighlighsYammer = $matchResult; break}
                        "Image" {$matchObj.Image = $matchResult; break}
                        "News" {$matchObj.News = $matchResult; break}
                        "Stream" {$matchObj.Stream = $matchResult; break}
                        "Twitter" {$matchObj.Twitter = $matchResult; break}
                        "Events" {$matchObj.Events = $matchResult; break}
                        "Forms" {$matchObj.Forms = $matchResult; break}
                        "People" {$matchObj.People = $matchResult; break}
                        Default {break}
                    }
                }
            }
        }
        #only run if a page title is valid. this will exclude any folders being output.
        if($page_title.Length -gt 0)
        {
            $row = $site_title + $delim + $page_title + $delim + $fileref + $delim + $matchObj.ConversationsYammer + $delim + $matchObj.Hero + $delim + $matchObj.HighlightedContent + $delim + $matchObj.QuickLinks + $delim + $matchObj.HighlighsYammer + $delim + $matchObj.Image + $delim + $matchObj.News + $delim + $matchObj.Stream + $delim + $matchObj.Twitter + $delim + $matchObj.Events + $delim + $matchObj.Forms + $delim + $matchObj.People
            Add-Content -Path $csv_path -Value $row
        }
    }
}

# itterate around each site from the csv
foreach($site in $sites)
{
    # make the connection, get ome site information and create object that contains all the site pages
    try
    {
        $connection = Connect-PnPOnline -Url $site.Url -UseWebLogin -ErrorAction Stop
        OutputToCSV
    }
    catch
    {
        Write-Host $_.Exception.Message
    }
    Disconnect-PnPOnline -Connection $connection
}

