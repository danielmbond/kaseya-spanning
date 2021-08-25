Clear-Host
# Remove o365 Spanning backup licenses from deleted accounts.
# Add Spanning License to all E5 users.

# Tested in PowerShell 7.1 and as of 8/3/2021 MSOnline module doesn't work. Use PowerShell 5.x

#region Initial setup.

# Install-Module MSOnline
# Install-Module Microsoft.PowerShell.SecretManagement -AllowPrerelease
# Install-Module Microsoft.PowerShell.SecretStore -AllowPrerelease
# Register-SecretVault -Name SecretStore -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
# Set-Secret -Name SpanningAPIKey -Secret APIKEY
# Set-Secret -Name SpanningLogin -Secret "EMAIL_ADDRESS"
# Set-Secret -Name 365Login -Secret "EMAIL_ADDRESS"
# Set-Secret -Name 365Password -Secret "PASSWORD"
#endregion

Import-Module MSOnline

Write-Host "Here we go."

$APIKEY = Get-Secret -Vault SecretStore -Name SpanningAPIKey -AsPlainText
$USERNAME = Get-Secret -Vault SecretStore -Name SpanningLogin -AsPlainText
$API_URL = "https://o365-api-us.spanningbackup.com/"
$API_USER_URL_ACTIVE = $API_URL + "external/users?inActiveDirectory=false&size=500"
$API_USER_URL_DELETED = $API_URL + "external/users?inActiveDirectory=false&size=500"
$API_ASSIGN_LICENSE = $API_URL + "external/users/assign"
$API_UNASSIGN_LICENSE = $API_URL + "external/users/unassign"
$spanningCred = $USERNAME + ":" + $APIKEY
$spanningCred64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($spanningCred))
$header = @{"Authorization" = "Basic " + $spanningCred64 }

$M365_PASSWORD = ConvertTo-SecureString (Get-Secret -Vault SecretStore -Name 365Password -AsPlainText) -AsPlainText -Force
$M365_USERNAME = Get-Secret -Vault SecretStore -Name 365Login -AsPlainText
[pscredential]$m365Cred = New-Object System.Management.Automation.PSCredential($M365_USERNAME, $M365_PASSWORD)

#region Fuctions
function add-spanning-licenses($addUsers) {
    if ($addUsers -and $addUsers.Count -gt 0) {
        $addUsers = $addUsers.Trim(",")
        $addUsersArray = $addUsers.Split(",")
        $bodyObj = New-Object -TypeName psobject
        $bodyObj | Add-Member -MemberType NoteProperty -Name userPrincipalNames -Value $addUsersArray -Force

        $body = $bodyObj | ConvertTo-Json

        try {
            Invoke-RestMethod -Method Post -Headers $header -ContentType "application/json" -uri $API_ASSIGN_LICENSE -Body $body
        } catch {Failure}
    
        $addUsers = ""
        Write-Host "$($addUsersArray.Count) users added to Spanning Licenses."
    } else {
        Write-Host 
    }
}

function add-spanning-licenses-to-e5s($licensedUsers = $licensedUsers, $e5s = $e5UPNs) {
    $addUsers = ""
    $count = 0
    foreach ($e5 in $e5s) {
        if ($licensedUsers.$e5 -ne $true) {
            $addUsers += "$($e5),"
            $count++
            if ($count % 100 -eq 0) { 
                add-spanning-licenses $addUsers
                $addUsers = ""
            }
        }
    }
    add-spanning-licenses $addUsers
    write-host "$count total Spanning Licenses added."
}

function Expand-UserProperties($user) {
    $userColumns = "DisplayName", "UserPrincipalName", "AccountEnabled", `
        "LastDirSyncTime", "Mail", `
        "MailNickName", "MSExchRecipientTypeDetails", "ObjectId", `
        "RefreshTokensValidFromDateTime", `
        "UserType", "WhenCreated"
    $userTemp = $user | Select-Object $userColumns

    return $userTemp
}

function Export-MsoUser-Object-To-CSV ($msoUsers=$e5s, $outfile=$null) {
    if ($null -eq $outfile) {
        $outfile = Get-Outfile
    }
    foreach ($user in $msoUsers) {
        $userTemp = Expand-UserProperties $user

        if ((Test-Path $outfile) -eq $false) {
            Write-Host $outfile
            $userTemp | ConvertTo-Csv -NoTypeInformation | Set-Content -Path $outfile
        }
        else {
            $userTemp | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Out-File -Append $outfile
        }
    }
}

function Get-Outfile($new=$false) {
    $datetime = (Get-Date -Format "yyyyMMdd-HHmmssK").Replace(":", "")
    $desktop = "$env:HOMEDRIVE$env:HOMEPATH\Desktop"
    $Global:outfile = "$desktop\output-$datetime.csv"
    if ($new -or $null -eq $Global:outfile) {
        return $Global:outfile
    } else {
        return $Global:outfile
    }
}

function Failure {
    $global:helpme = $body
    $global:helpmoref = $moref
    $global:result = $_.Exception.Response.GetResponseStream()
    $global:reader = New-Object System.IO.StreamReader($global:result)
    $global:responseBody = $global:reader.ReadToEnd();
    Write-Host -BackgroundColor:Black -ForegroundColor:Red "Status: A system exception was caught."
    Write-Host -BackgroundColor:Black -ForegroundColor:Red $global:responsebody
    Write-Host -BackgroundColor:Black -ForegroundColor:Red "The request body has been saved to `$global:helpme"
}

function Get-Accounts-With-Mailboxes ($accounts) {
#MSExchRecipientTypeDetails
    $accountsWithMailbox = @{}
    foreach ($account in $accounts) {
        if ($null -ne $account.MSExchRecipientTypeDetails) {
            $accountsWithMailbox.Add($account,$true)
        }
    }
    return ($accountsWithMailbox.Keys)
}

function Get-m365-Users($m365Cred = $m365Cred) {
    Import-Module MSOnline
    Connect-MsolService -Credential $m365Cred
    Write-Host "Getting users, this will take a few minutes."
    $msolUsers = Get-MsolUser -All -Verbose
    return $msolUsers
}

function get-spanning-users($API_USER_URL = $API_USER_URL_DELETED) {
    $firstRun = $true
    $usersHash = @{}
    while ($firstRun -or $users.nextLink) {
        if ($firstRun -eq $true) {
            $url = $API_USER_URL
        }
        else {
            $url = $users.nextLink
        }
        $firstRun = $false 
        Write-Host $url
        $users = Invoke-RestMethod -Method Get -Headers $header -ContentType "text/plain" -uri $url
        try {
            $usersHash.Add($users, "")
        }
        catch [System.Management.Automation.MethodInvocationException] {
            return $usersHash.keys.users
        }
        
    }
    return $usersHash.keys.users
}

function get-spanning-licensed-users-licenses($users = $users) {
    $licensedUsers = @{}
    $count = 0
    foreach ($user in $users) {
        if ($user.assigned -eq $true) {
            $licensedUsers.Add($user.userPrincipalName, $true)
            $count++
        }
    }
    Write-Host "$count licensed Spanning users."
    return $licensedUsers
}

function remove-spanning-deleted-users-licenses($users = $users) {
    $deletedUsers = ""
    $count = 0
    foreach ($user in $users) {
        if ($user.isDeleted -eq $true) {
            $deletedUsers += "$($user.userPrincipalName),"
            $count++
            Write-Host $count
            Write-Host $($deletedUsers.Count)
        }
    }
    if ($deletedUsers -and $deletedUsers.Count -gt 0) {
        $deletedUsers = $deletedUsers.Trim(",")
        $deletedUsersArray = $deletedUsers.Split(",")
        $bodyObj = New-Object -TypeName psobject
        $bodyObj | Add-Member -MemberType NoteProperty -Name userPrincipalNames -Value $deletedUsersArray -Force

        $body += $bodyObj | ConvertTo-Json

        Invoke-RestMethod -Method Post -Headers $header -ContentType "application/json" -uri $API_UNASSIGN_LICENSE -Body $body
        Write-Host "$count deleted users licenses removed."
    } else {
        Write-Host "There were no deleted users to remove licenses from."
    }
}
#endregion

# Get licened Spanning users NOT in AD.
if ($null -eq $users) {
    Write-Host "Getting Spanning users NOT in AD."
    $users = get-spanning-users
}
else {
    Write-Host "Using cached Spanning users. Set `$users = `$null to pull them down again, `
    it takes a while depending on the amount of users."
}

# Remove License from deleted Users.
if ($null -ne $users) {
    Write-Host "Removing licenses from users that have been deleted."
    remove-spanning-deleted-users-licenses $users
}

# Apply Licenses to E5 Users.
if ($null -eq $msolUsers) {
    $msolUsers = Get-m365-Users $m365Cred
}
else {
    Write-Host "Using cached user list. Set `$msolUsers to `$null to refresh the list."
}
$e5s = ($msolUsers | Where-Object `
    { ($_.licenses).AccountSkuId -match "EnterprisePremium" })
$e5sWithMailBoxes = Get-Accounts-With-Mailboxes $e5s
$e5UPNs = $e5sWithMailBoxes.UserPrincipalName
$e5sWithoutMailbox = ($e5s | Where-Object `
    { ($_.MSExchRecipientTypeDetails) -eq $null })#.UserPrincipalName
Export-MsoUser-Object-To-CSV $e5sWithoutMailbox

# Get licened Spanning users in AD.
$users = $null
Write-Host "Getting Spanning users in AD."
$users = get-spanning-users $API_USER_URL_ACTIVE


if ($null -eq $licensedUsers) {
    $licensedUsers = get-spanning-licensed-users-licenses $users
}
else {
    Write-Host "Using cached user list. Set `$licensedUsers to `$null to refresh the list."
}

# Add Spanning licenses to E5 users.
if ($null -ne $e5UPNs) {
    add-spanning-licenses-to-e5s $licensedUsers, $e5UPNs
}
