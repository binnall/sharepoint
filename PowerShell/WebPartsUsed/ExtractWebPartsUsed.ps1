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

# Create an object that will be used to store if a match has been found for the web parts being checked for 
function CreateMatchObj{
    $def = "no" # default value set here. This could be updated to a bool if required 
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
    $pnpsite = Get-PnPWeb -Connection $connection # get the web object
    $site_title = $pnpsite.Title # set the title as this is used in the outout file
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
            # itterate around each webpart in the input file. If the ID is found 
            # in the canvas content for the current page then break into the swtich and update the default result to yes
            foreach($wp in $webparts){
                if($canvascontent -match $wp.ID) # when ID is located in the canvas content field it will return true
                {
                    switch ($wp.Title) { # this will switch on the title of the webpart that the ID has matched on
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
    # make the connection, and call the function to create the csv
    try
    {
        # create output csv
        New-Item -Path $creation_path -Name $file_name -ItemType File -Value $csv_outputheader -ErrorAction Stop
        $connection = Connect-PnPOnline -Url $site.Url -UseWebLogin -ErrorAction Stop
        OutputToCSV
    }
    catch # should catch an error if a connection to the site can't be made
    {
        Write-Host $_.Exception.Message
    }
    Disconnect-PnPOnline -Connection $connection
}