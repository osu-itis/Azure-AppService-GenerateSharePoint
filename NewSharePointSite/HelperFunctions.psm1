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

<#
.SYNOPSIS
Find an unused mail nickname
#>
function FindUnusedMailNickname {
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
        [parameter(Mandatory=$true)]$Body
    )
    try {
        write-host "Sending BadRequest response: $Body"
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body = $($Body|ConvertTo-Json)
        })
    }
    catch {
        Write-Warning -Message $Body
    }
}

function GoodRequest {
    param (
        [parameter(Mandatory=$true)]$Body
    )
    try {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body = $($Body|ConvertTo-Json)
        })
    }
    catch {
        Write-Host $Body
    }
}