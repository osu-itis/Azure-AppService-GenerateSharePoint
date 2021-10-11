function CheckNeededENVs {
    <#
    .SYNOPSIS
    Check that the needed ENVs are all avalable before proceeding
    #>
    param (
        # No parameters
    )

    if ([string]::IsNullOrEmpty($env:ClientID)) { Throw 'Could not find $env:ClientID' }
    if ([string]::IsNullOrEmpty($env:ClientSecret)) { Throw 'Could not find $env:ClientSecret' }
    if ([string]::IsNullOrEmpty($env:TenantId)) { Throw 'Could not find $env:TenantId' }
}

Function ConvertFormat {
    <#
    .SYNOPSIS
    Converts HTML encoding to standard formatting and removes any leading or trailing whitespace

    .PARAMETER InputText
    The input text to convert

    .EXAMPLE
    PS>$temp = "This+is+a+test%2fexample%0D%0A%0D%0AAnd+it+rocks"
    PS>convertformat -InputText $temp

    This is a test/example

    And it rocks
    "
    #>
    PARAM (
        [parameter(Mandatory = $true)][string]$InputText
    )

    Add-Type -AssemblyName System.Web

    $OutputText = [string]$(
        [System.Web.HttpUtility]::UrlDecode(
            $InputText
        )
    ).Trim()

    Return $OutputText
}

function ResolveOwner {
    <#
    .SYNOPSIS
    Resolve the email address to the UserPrincipalName
    #>
    param (
        [parameter(Mandatory = $true)][string]$Owner,
        [parameter(Mandatory = $true)]$Headers
    )

    # This formatting is intentional, the $filter needs to be single quoted due to the dollarsign, the single quotes need to be double quoted and the variables should not be single quoted so they are evaluated properly
    # Example of the output: https://graph.microsoft.com/v1.0/users/?$filter=mail eq 'email.address@oregonstate.edu' or userprincipalname eq 'email.address@oregonstate.edu'
    $Uri = "https://graph.microsoft.com/v1.0/users/" + '?$filter=mail eq' + " '" + $($Owner) + "' " + 'or userprincipalname eq' + " '" + $($Owner) + "' "

    try {
        # Getting the resolved owner
        $ResolvedOwner = Invoke-WebRequest -Method Get -Headers $Headers -Uri $Uri
    }
    catch {
        Write-Error -Message "Failed to identity the Owner" -ErrorAction Stop
    }

    # Returning the resolved owner's UPN
    Return ($ResolvedOwner.Content | ConvertFrom-Json).value.userprincipalname
}

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
        $ownerprecheck = Invoke-WebRequest -Method Get -Headers $Headers -Uri "https://graph.microsoft.com/v1.0/users/$($Owner)"
        return ($ownerprecheck.Content | ConvertFrom-Json).userprincipalname -eq $Owner
    }
    catch {
        return $false
    }
}

function FindUnusedMailNickname {
    <#
    .SYNOPSIS
    Find an unused mail nickname
    #>
    param (
        $displayName,
        $Headers
    )

    $mailNickname = $( $displayName -replace '`~!@#$%^&*()-_=+?|/\;:,.<>', '' ).Replace(' ', '')
    # Note we are adding the consistency level parameter to the headers so we can complete the group search against mail nicknames ( https://docs.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0 )
    # formatting the uri into the needed format: https://graph.microsoft.com/v1.0/groups?$search="mailNickname:<MAILNICKNAME>"
    $test = Invoke-RestMethod -Method get -Headers $($Headers + @{ConsistencyLevel = 'eventual' }) -Uri $('https://graph.microsoft.com/v1.0/groups?$search=' + '"' + 'mailNickname:' + $mailNickname + '"')
    $NicknameExists = $test.value.mailnickname -contains $mailNickname
    if ($NicknameExists) {
        $X = 1
        do {
            # Adding x to the nickname to try and find an unused nickname
            $mailNickname = $( $displayName -replace '`~!@#$%^&*()-_=+?|/\;:,.<>', '' ).Replace(' ', '') + $X
            $test = Invoke-RestMethod -Method get -Headers $($Headers + @{ConsistencyLevel = 'eventual' }) -Uri $('https://graph.microsoft.com/v1.0/groups?$search=' + '"' + 'mailNickname:' + $mailNickname + '"')
            $NicknameExists = $test.value.mailnickname -contains $mailNickname
            $X++
        } until ($NicknameExists -eq $false)
    }
    return $mailNickname
}

function BadRequest {
    param (
        [parameter(Mandatory = $true)]$Body
    )
    try {
        write-host "Sending BadRequest response: $Body"
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::BadRequest
                Body       = $($Body | ConvertTo-Json)
            })
    }
    catch {
        Write-Warning -Message $Body
    }
}

function GoodRequest {
    param (
        [parameter(Mandatory = $true)]$Body
    )
    try {
        write-host "Sending GoodRequest response"
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::OK
                Body       = $($Body | ConvertTo-Json)
            })
    }
    catch {
        Write-Host $Body
    }
}
