# Credentials
$creds = Get-Credential

# CSV location that contains the columns, site to create them in and lists to create them in
$columns_csv_locaiton = "\ColumnCreationtext.csv"

$headers = 'Site','List','InternalName', 'XML'

$columns = Import-Csv -Path $columns_csv_locaiton -Header $headers -Delimiter ','

# itterate the csv that contains all the columns
foreach($col in $columns)
{
    $pnpcon = Connect-PnPOnline -url $col.Site -Credentials $creds
    # if there is a field to create the first if statement will run
    if ($col.XML.Length -gt 0) {
        try {
            Add-PnPFieldFromXml -FieldXml $col.XML -Connection $pnpcon -ErrorAction Stop
    
            Add-PnPField -Field $col.InternalName -List $col.List -Connection $pnp -ErrorAction Stop
        }
        catch {
            write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
        }
    }elseif ($col.InternalName -gt 0) 
    {
        # this code should run only when a column is to be associated to a list, 
        # column already exists and just needs associating
        try {    
            Add-PnPField -Field $col.InternalName -List $col.List -Connection $pnp -ErrorAction Stop
        }
        catch {
            write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
        }
    }

    Disconnect-PnPOnline -Connection $pnpcon
}