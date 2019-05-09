# Description
* PowerShell script to create site columns and associate them to a list.
* Script uses a csv as an input file.
* All columns are created based on the XML field definition

# Csv info
A csv file is used for input

## Csv headers
Order of columns for csv:

Site, List, InternalName, XML

### Csv header desription
site = url to create column in

list = list to utilise the column created

InternalName = internal name of column so it can be added to the list

XML = Field XML based on the schema

## Examples of valid csv input
Example of valid csv input:

1) adding two text columns and 1 number column

https://jeb03.sharepoint.com/sites/CommsPnP,Site Pages,TextCol1,<Field ID="{}" Type="Text" Name="TextCol1" DisplayName="Text Col 1" StaticName="TextCol1" Group="Test" Required="FALSE"/>
https://jeb03.sharepoint.com/sites/CommsPnP,Site Pages,TextCol2,<Field ID="{}" Type="Text" Name="TextCol2" DisplayName="Text Col 2" StaticName="TextCol28" Group="Test" Required="FALSE"/>
https://jeb03.sharepoint.com/sites/CommsPnP,Site Pages,NumberCol1,<Field ID="{}" Type="Number" Name="NumberCol1" DisplayName="Number Col 1" StaticName="NumberCol1" Group="Test" Required="FALSE"/>

# References

https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/add-pnpfieldfromxml?view=sharepoint-ps
https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/add-pnpfield?view=sharepoint-ps

