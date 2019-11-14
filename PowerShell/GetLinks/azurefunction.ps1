#Intro
 
# Script will take a csv file that contains url to SharePoint sites and analyse the site pages to see if any of the pages have hyperlinks.
# For every hyperlink in a page this gets output to a row in a csv that is delimited by a pipe

# will get the variables / credentials in the azure automation service
$itsite = Get-AutomationVariable -Name 'ITSite'
$cred = Get-AutomationPSCredential -Name 'admin'
$regex = Get-AutomationVariable -Name 'regex'

$list = "SitePages"

#connection for Internal IT site that contains a list of sites to run the script against

$internalITCon = Connect-PnPOnline -Url $itsite -Credentials $cred
# create object of all the sites
$sitesObj = (Get-PnPListItem -List 'sitesAutomation' -Fields "url" -Connection $internalITCon)

# itterate around each site from the csv
foreach($siteUrl in $sitesObj)
{
    Write-Host $siteUrl["url"]
    # make the connection, get ome site information and create object that contains all the site pages
    $connection = Connect-PnPOnline -Url $siteUrl["url"] -Credentials $cred
    $pnpsite = Get-PnPWeb -Connection $connection
    $site_title = $pnpsite.Title
    $pages = (Get-PnPListItem -List $list -Fields "CanvasContent1", "Title" -Connection $connection).FieldValues

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
            # hash table of the results that match the href regular expression
            $hrefmatches = ($canvascontent | select-string -pattern $regex -AllMatches).Matches.Value

            # itterate around each regular expression match and write it out into the output csv that is pipe delimited 
            foreach($hrefmatch in $hrefmatches)
            {
                $internalITCon2 = Connect-PnPOnline -Url "https://m365x630080.sharepoint.com/sites/InternalIT/" -Credentials $cred
                Add-PnPListItem -List "Migration Hyperlink Progress" -Values @{"Title" = "N/A"; "SiteTitle" = $site_title;"PageTitle" = $page_title; "Ref" = $fileref; "Match" = $hrefmatch} -Connection $internalITCon2
            }
        }
    }
    Disconnect-PnPOnline -Connection $connection
}