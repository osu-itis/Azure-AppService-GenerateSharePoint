# Load ENVs
. C:\Users\carrk\GitHub\CodeSnippet-Azure-AutoLoadENVs\AutoLoadENVs.ps1
AutoLoadENVs

# Import the graph api token module
Import-Module C:\Users\carrk\GitHub\Azure-AppService-GenerateSharePoint\Modules\New-GraphAPIToken\New-GraphAPIToken.psm1

# Gathering a token and setting the headers
$GraphAPIToken = New-GraphAPIToken -ClientID $env:ClientID -ClientSecret $env:ClientSecret -TenantID $env:TenantID
$Headers = $GraphAPIToken.Headers
