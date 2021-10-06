using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

if ($Request) {
    $Request | Export-Clixml .\NewSharePointSite\request.cli.xml
}

if ($TriggerMetadata) {
    $TriggerMetadata | Export-Clixml .\NewSharePointSite\TriggerMetadata.cli.xml
}

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Load ENVs
. C:\Users\carrk\GitHub\CodeSnippet-Azure-AutoLoadENVs\AutoLoadENVs.ps1
AutoLoadENVs

# Import the graph api token module
Import-Module .\Modules\New-GraphAPIToken\New-GraphAPIToken.psm1

# Import the helper functions
Import-Module .\NewSharePointSite\HelperFunctions.psm1

# Gathering a token and setting the headers
$GraphAPIToken = New-GraphAPIToken -ClientID $env:ClientID -ClientSecret $env:ClientSecret -TenantID $env:TenantID
$Headers = $GraphAPIToken.Headers

# Importing the cached request
$Request = Import-Clixml -Path .\NewSharePointSite\request.cli.xml

# Creating the master 'values' variable that will contain all status information
$Values = [PSCustomObject]@{
    owner          = $Request.Body.owner
    displayName    = $Request.Body.displayName
    description    = $Request.Body.description
    ticketID       = $Request.Body.ticketID

    # These will all be set later after other actions are taken
    ownerExists    = $null
    mailNickname   = $null
    template       = $null
    creationstatus = $null
    Sharepointdata = $null
}

# Running precheck and gathering the needed information
switch ($values) {
    #Bad Outcomes
    { [string]::IsNullOrEmpty($_.owner) } {
        BadRequest -Body "Missing owner value"
        continue
    }
    { [string]::IsNullOrEmpty($_.displayName) } {
        BadRequest -Body "Missing displayName value"
        continue
    }
    { [string]::IsNullOrEmpty($_.description) } {
        BadRequest -Body "Missing description value"
        continue
    }
    { [string]::IsNullOrEmpty($_.ticketID) } {
        BadRequest -Body "Missing ticketID value"
        continue
    }
    Default {
        BadRequest -Body "Request is in an invalid format"
        continue
    }
    #Good Outcomes
    { ($_.owner -ne $null) -and ($_.displayName -ne $null) -and ($_.description -ne $null) -and ($_.ticketID -ne $null) } {
        # Writing to the Azure Functions log stream.
        Write-Host "Ticket ID detected as: $($Values.ticketID)"
        Write-Host "Owner detected as: $($Values.owner)"
        Write-Host "Display name detected as: $($Values.displayName)"

        # Precheck - does the owner exist?
        $values.ownerExists = $(DoesOwnerExist -Headers $Headers -Owner $Values.owner)

        # Write to the Azure Functions log stream.
        Write-Host "Owner exists: $($Values.ownerExists)"

        # Precheck - Find an unused mail nickname
        $Values.mailNickname = $(FindUnusedMailNickname -displayName $Values.displayName -Headers $Headers)

        # Write to the Azure functions log stream.
        Write-Host "Unused mail nickname: $($Values.mailNickname)"
    }
}

# Creating the template to be used when creating the new sharepoint group
switch ($Values) {
    # Bad outcomes
    Default {
        BadRequest -Body "Could not determine the values status"
        continue
    }
    { $_.ownerExists -eq $false } {
        BadRequest -Body "Owner does not exist or could not be found"
        continue
    }
    { [string]::IsNullOrEmpty($_.mailNickname) } {
        BadRequest -Body "No mail nickname"
        continue
    }
    # Good outcomes
    { $_.ownerExists -eq $true } {
        # Creating the template to be used for creating the new (sharepoint) group
        # https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-customizations-provisioning-sites
        $values.template = [PSCustomObject]@{
            description             = $Values.description
            displayName             = $Values.displayName
            groupTypes              = @("Unified")
            resourceBehaviorOptions = @("HideGroupInOutlook")
            mailEnabled             = $true
            mailNickname            = $Values.mailNickname
            securityEnabled         = $false
            Visibility              = "private"
            "owners@odata.bind"     = [array]@( $( [string]"https://graph.microsoft.com/v1.0/users/$($Values.owner)" ) )
            "members@odata.bind"    = [array]@( $( [string]"https://graph.microsoft.com/v1.0/users/$($Values.owner)" ) )
        }
    }
}

# If the template exists, make the new sharepoint group
switch ($values) {
    # Good Outcomes
    { $_.template -ne $null } {
        # Creating the sharepoint group and capturing the output
        Write-Host "Creating sharepoint group"
        $values.creationStatus = Invoke-WebRequest -Method Post -Headers $Headers -ContentType 'application/json' -Uri "https://graph.microsoft.com/v1.0/groups" -Body $($values.template | ConvertTo-Json)
        Write-Host "Creation status $($Values.creationstatus.StatusCode), $($Values.creationstatus.StatusDescription)"

        # Summary of the new sharepoint group
        $values.Sharepointdata = $values.creationstatus.Content | ConvertFrom-Json | Select-Object ID, Displayname, Description, @{name = "webUrl"; Expression = { "https://oregonstateuniversity.sharepoint.com/sites/" + $_.MailNickName } }, Mail, MailNickname, visibility, CreatedDateTime
        Write-Host "New sharepoint url: $($Values.Sharepointdata.webUrl)"

        # Responding with the good response, providing the summary of the sharepoint data
        GoodRequest -Body $Values.Sharepointdata
    }
    # Bad Outcomes
    { $_.template -eq $null } {
        BadRequest -Body "A sharepoint template could not be generated"
        Continue
    }
}

# Exporting the clixml with this run's values
Export-Clixml -InputObject $Values -Path .\NewSharePointSite\values.cli.xml
