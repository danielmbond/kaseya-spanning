# Remove o365 Spanning backup licenses from deleted accounts.
# Add Spanning License to all E5 users.

# Tested in PowerShell 7.1 and as of 8/3/2021 MSOnline module doesn't work. Use PowerShell 5.x

# Initial setup.

# Install-Module MSOnline
# Install-Module Microsoft.PowerShell.SecretManagement -AllowPrerelease
# Install-Module Microsoft.PowerShell.SecretStore -AllowPrerelease
# Register-SecretVault -Name SecretStore -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
# Set-Secret -Name SpanningAPIKey -Secret APIKEY
# Set-Secret -Name SpanningLogin -Secret "EMAIL_ADDRESS"
# Set-Secret -Name 365Login -Secret "EMAIL_ADDRESS"
# Set-Secret -Name 365Password -Secret "PASSWORD"

Import-Module MSOnline

Write-Host "Here we go."

$APIKEY = Get-Secret -Vault SecretStore -Name SpanningAPIKey -AsPlainText
$USERNAME = Get-Secret -Vault SecretStore -Name SpanningLogin -AsPlainText
$API_URL = "https://o365-api-us.spanningbackup.com/"
$API_USER_URL = $API_URL + "external/users?size=1000"
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
    $addUsers = $addUsers.Trim(",")
    $addUsersArray = $addUsers.Split(",")
    $bodyObj = New-Object -TypeName psobject
    $bodyObj | Add-Member -MemberType NoteProperty -Name userPrincipalNames -Value $addUsersArray -Force

    $body = $bodyObj | ConvertTo-Json

    Invoke-RestMethod -Method Post -Headers $header -ContentType "application/json" -uri $API_ASSIGN_LICENSE -Body $body
    $addUsers = ""
    write-host "$($addUsersArray.Count) users added to Spanning Licenses."
}

function add-spanning-licenses-to-e5s($licensedUsers = $licensedUsers, $e5s = $e5s) {
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

function Get-m365-Users($m365Cred = $m365Cred) {
    Import-Module MSOnline
    Connect-MsolService -Credential $m365Cred
    Write-Host "Getting users, this will take a few minutes."
    $msolUsers = Get-MsolUser -All -Verbose
    return $msolUsers
}

function get-spanning-users($API_USER_URL = $API_USER_URL) {
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
            # Write-Host $($user.userPrincipalName)
            $deletedUsers += "$($user.userPrincipalName),"
            $count++
        }
    }
    
    $deletedUsers = $deletedUsers.Trim(",")
    $deletedUsersArray = $deletedUsers.Split(",")
    $bodyObj = New-Object -TypeName psobject
    $bodyObj | Add-Member -MemberType NoteProperty -Name userPrincipalNames -Value $deletedUsersArray -Force

    $body += $bodyObj | ConvertTo-Json

    Invoke-RestMethod -Method Post -Headers $header -ContentType "application/json" -uri $API_UNASSIGN_LICENSE -Body $body
    Write-Host "$count deleted users licenses removed."
}
#endregion

# Get all Spanning Users
if ($null -eq $users) {
    Write-Host "Getting users."
    $users = get-spanning-users
}
else {
    Write-Host "Using cached users. Set `$users = `$null to pull them down again, `
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
    { ($_.licenses).AccountSkuId -match "EnterprisePremium" }).UserPrincipalName

# Get licened Spanning users.
if ($null -eq $licensedUsers) {
    $licensedUsers = get-spanning-licensed-users-licenses $users
}
else {
    Write-Host "Using cached user list. Set `$licensedUsers to `$null to refresh the list."
}

# Add Spanning licenses to E5 users.
if ($null -ne $e5s) {
    add-spanning-licenses-to-e5s $licensedUsers, $e5s
}
