# Load ENVs
. C:\Users\carrk\GitHub\CodeSnippet-Azure-AutoLoadENVs\AutoLoadENVs.ps1
AutoLoadENVs

# Import the graph api token module
Import-Module C:\Users\carrk\GitHub\Azure-AppService-GenerateSharePoint\Modules\New-GraphAPIToken\New-GraphAPIToken.psm1

# Gathering a token and setting the headers
$GraphAPIToken = New-GraphAPIToken -ClientID $env:ClientID -ClientSecret $env:ClientSecret -TenantID $env:TenantID
$Headers = $GraphAPIToken.Headers

# Setting the needed input values
$owner = "carrk@oregonstate.edu"
$displayName = "SharePoint Test"
$description = "Testing SharePoint"

# Creating the master 'values' variable that will contain all status information
$Values = [PSCustomObject]@{
    owner          = $owner
    displayName    = $displayName
    # Strip the special characters and spaces
    description    = $description
    
    # These will all be set later after other actions are taken
    ownerExists    = $null
    mailNickname   = $null
    template       = $null
    creationstatus = $null
    Sharepointdata = $null
}
Write-Host "Owner detected as $($Values.owner)"
Write-Host "Display name detected as $($Values.displayName)"

# Precheck - does the owner exist?
function DoesOwnerExist {
    <#
    .SYNOPSIS
    Check if the owner exists, return a true/false
    #>
    param (
        $Headers,
        $Owner
    )
    try {
        $ownerprecheck = Invoke-WebRequest -Method Get -Headers $Headers -Uri "https://graph.microsoft.com/v1.0/users/$($Values.owner)"
        return ($ownerprecheck.Content | ConvertFrom-Json).userprincipalname -eq $Values.owner
    }
    catch {
        return $false    
    }
}

$values.ownerExists = $(DoesOwnerExist -Headers $Headers -Owner $Values.owner)

Write-Host "Owner exists: $($Values.ownerExists)"




# Precheck - is the mail nickname already in use?

# Note we are adding the consistency level parameter to the headers so we can complete the group search against mail nicknames
# https://docs.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0
# formatting the uri into the needed format: https://graph.microsoft.com/v1.0/groups?$search="mailNickname:<MAILNICKNAME>"


function FindUnusedMailNickname {
    param (
        $displayName,
        $Headers
    )
    
    $mailNickname = $( $displayName -replace '`~!@#$%^&*()-_=+?|/\;:,.<>', '' ).Replace(' ', '')
    $test = Invoke-RestMethod -Method get -Headers $($Headers + @{ConsistencyLevel = 'eventual' }) -Uri $('https://graph.microsoft.com/v1.0/groups?$search=' + '"' + 'mailNickname:' + $mailNickname + '"')
    $NicknameExists = $test.value.mailnickname -contains $mailNickname
    if ($NicknameExists) {
        $X = 1
        do {
            $mailNickname = $( $displayName -replace '`~!@#$%^&*()-_=+?|/\;:,.<>', '' ).Replace(' ', '') + $X
            $test = Invoke-RestMethod -Method get -Headers $($Headers + @{ConsistencyLevel = 'eventual' }) -Uri $('https://graph.microsoft.com/v1.0/groups?$search=' + '"' + 'mailNickname:' + $mailNickname + '"')
            $NicknameExists = $test.value.mailnickname -contains $mailNickname
            $X++
        } until ($NicknameExists -eq $false)    
    }
    return $mailNickname
}

$Values.mailNickname = $(FindUnusedMailNickname -displayName $Values.displayName -Headers $Headers)



switch ($Values) {
    # Bad outcomes
    { $_.ownerExists -eq $false } { "Owner does not exist or could not be found"; continue }
    Default { "Could not determine the values status"; continue }

    # Good outcomes
    { $_.ownerExists -eq $true } { "Everything is ready to roll"; continue }
}


# https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-customizations-provisioning-sites
$values.template = [PSCustomObject]@{
    description          = $Values.description
    displayName          = $Values.displayName
    groupTypes           = @("Unified")
    mailEnabled          = $true
    mailNickname         = $Values.mailNickname
    securityEnabled      = $false
    Visibility           = "private"
    "owners@odata.bind"  = [array]@( $( [string]"https://graph.microsoft.com/v1.0/users/$($Values.owner)" ) )
    "members@odata.bind" = [array]@( $( [string]"https://graph.microsoft.com/v1.0/users/$($Values.owner)" ) )
}

$values.creationStatus = Invoke-WebRequest -Method Post -Headers $Headers -ContentType 'application/json' -Uri "https://graph.microsoft.com/v1.0/groups" -Body $($values.template | ConvertTo-Json)

Start-sleep -Seconds 60

$values.Sharepointdata = Invoke-RestMethod -Method get -Headers $Headers -Uri 'https://graph.microsoft.com/v1.0/sites?$search="sharepoint test"' | Select-Object -ExpandProperty value | Where-Object { $_.name -eq $Values.mailNickname }

Export-Clixml -InputObject $Values -Path .\values.cli.xml






# $results = Invoke-RestMethod -Method get -Headers $Headers -Uri https://graph.microsoft.com/v1.0/sites/
# $results2 = Invoke-RestMethod -Method get -Headers $Headers -Uri https://graph.microsoft.com/v1.0/sites/root
# $results3 = Invoke-RestMethod -Method get -Headers $Headers -Uri 'https://graph.microsoft.com/v1.0/sites?$search="sharepoint test"'

