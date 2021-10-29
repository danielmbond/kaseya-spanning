Clear-Host
# Remove o365 Spanning backup licenses from deleted accounts.
# Add Spanning License to all E5 users.

# Tested in PowerShell 7.1 and as of 8/3/2021 MSOnline module doesn't work. Use PowerShell 5.x

#region Initial setup.

# Install-Module MSOnline
# Install-Module Microsoft.PowerShell.SecretManagement -AllowPrerelease
# Install-Module Microsoft.PowerShell.SecretStore -AllowPrerelease
# Register-SecretVault -Name SecretStore -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
# Set-Secret -Name SpanningAPIKey -Secret "APIKEY"
# Set-Secret -Name SpanningLogin -Secret "EMAIL_ADDRESS"
# Set-Secret -Name 365Login -Secret "EMAIL_ADDRESS"
# Set-Secret -Name 365Password -Secret "PASSWORD"
#endregion

Import-Module MSOnline

Write-Host "Here we go."

$APIKEY = Get-Secret -Vault SecretStore -Name SpanningAPIKey -AsPlainText
$USERNAME = Get-Secret -Vault SecretStore -Name SpanningLogin -AsPlainText
$API_URL = "https://o365-api-us.spanningbackup.com/"
$API_USER_URL_ACTIVE = $API_URL + "external/users?inActiveDirectory=true&size=500"
$API_USER_URL_DELETED = $API_URL + "external/users?inActiveDirectory=false&size=500"
$API_ASSIGN_LICENSE = $API_URL + "external/users/assign"
$API_UNASSIGN_LICENSE = $API_URL + "external/users/unassign"
$UNASSIGN_USER_FILE = 'no-ad-license.txt'
$DAYS_BEFORE_PURGE = 30

$spanningCred = $USERNAME + ":" + $APIKEY
$spanningCred64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($spanningCred))
$header = @{"Authorization" = "Basic " + $spanningCred64 }

$M365_PASSWORD = ConvertTo-SecureString (Get-Secret -Vault SecretStore -Name 365Password -AsPlainText) -AsPlainText -Force
$M365_USERNAME = Get-Secret -Vault SecretStore -Name 365Login -AsPlainText
[pscredential]$m365Cred = New-Object System.Management.Automation.PSCredential($M365_USERNAME, $M365_PASSWORD)

#region Fuctions
function add-spanning-licenses($addUsers, $dryrun = $false) {
    if ($addUsers -and $addUsers.Count -gt 0) {
        $addUsers = $addUsers.Trim(",")
        $addUsersArray = $addUsers.Split(",")
        $bodyObj = New-Object -TypeName psobject
        $bodyObj | Add-Member -MemberType NoteProperty -Name userPrincipalNames -Value $addUsersArray -Force

        $body = $bodyObj | ConvertTo-Json

        try {
            if ($dryrun -eq $false) {
                Invoke-RestMethod -Method Post -Headers $header -ContentType "application/json" -uri $API_ASSIGN_LICENSE -Body $body
            }
        }
        catch {}
    
        $addUsers = ""
        Write-Host "$($addUsersArray.Count) users added to Spanning Licenses."
    }
}

function add-spanning-licenses-to-m365LicensedUsers($licensedUsers = $licensedUsers, $m365LicensedUsers = $e5UPNs, $dryrun = $false) {
    $addUsers = ""
    $count = 0
    foreach ($m365LicensedUser in $m365LicensedUsers) {
        if ($licensedUsers.$m365LicensedUser -ne $true) {
            $addUsers += "$($m365LicensedUser),"
            $count++
            if ($count % 100 -eq 0) {
                add-spanning-licenses $addUsers $dryrun
                $addUsers = ""
            }
        }
    }
    add-spanning-licenses $addUsers $dryrun
    write-host "$count total Spanning Licenses added."
}

function Add-UsersToDelete($file = $UNASSIGN_USER_FILE, $users = $users) {
    $allUsers = [System.Collections.ArrayList]@()
    if ((Test-Path -LiteralPath $file)) {
        try {
            [System.Collections.ArrayList]$savedUsers = Get-Content -Raw $file | ConvertFrom-Json
            if ($savedUsers.Count -gt 0) {
                foreach ($savedUser in $savedUsers) {
                    $allUsers.Add($savedUser) | Out-Null
                }
            } else {
                $savedUsers = $null
            }
        }
        catch {}
    }

    if ($null -ne $users) {
        foreach ($user in $users) {
            Add-Member -InputObject $user -MemberType NoteProperty -Name dateDeleted -Value $(get-date).ToString()
            $allUsers.Add($user) | Out-Null
            if ($null -ne $savedUsers -and $savedUsers.userPrincipalName.Contains($user.userPrincipalName) -eq $true) {
                $allUsers.Remove($user)
                write-host "$($user.userPrincipalName) exists."
            } else {
                Write-Host "$($user.displayName) added."
            }
        }       
    }
    $allUsers | ConvertTo-Json | Out-File $file
    return $allUsers
}

function Expand-UserProperties($user) {
    $userColumns = "DisplayName", "UserPrincipalName", "AccountEnabled", `
        "LastDirSyncTime", "Mail", "MailNickName", "MSExchRecipientTypeDetails", "ObjectId", `
        "RefreshTokensValidFromDateTime", "UserType", "WhenCreated"
    $userTemp = $user | Select-Object $userColumns

    return $userTemp
}

function Export-MsoUser-Object-To-CSV ($msoUsers = $e5s, $outfile = $null) {
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

function Get-OldUsers([string]$file = $UNASSIGN_USER_FILE, $purgeAfter = $DAYS_BEFORE_PURGE) {
    if ((Test-Path -LiteralPath $file)) {
        $savedUsers = Get-Content -Raw $file | ConvertFrom-Json
        try {
            if ($savedUsers.Count -lt 114) {
                return $null
            }
        }
        catch {
            return $null
        }
        return $savedUsers
    }
    return $null
}

function Get-UsersToDelete($users = $users) {
    $pendingDeletes = Get-OldUsers
    if ($pendingDeletes) {

    }
    if ($users) {

    }
    if ($null -ne $users) {
        Add-Member -InputObject $tester -MemberType NoteProperty -Name dateDeleted -Value $(Get-Date)
        $newUsers.Userprincipalname.Contains("jmfxygs@jmfamily.com")
    }
}

function Get-Outfile($new = $false) {
    $datetime = (Get-Date -Format "yyyyMMdd-HHmmssK").Replace(":", "")
    $desktop = "$env:HOMEDRIVE$env:HOMEPATH\Desktop"
    $Global:outfile = "$desktop\output-$datetime.csv"
    if ($new -or $null -eq $Global:outfile) {
        return $Global:outfile
    }
    else {
        return $Global:outfile
    }
}

function Get-Accounts-With-Mailboxes ($accounts) {
    #MSExchRecipientTypeDetails
    $accountsWithMailbox = @{}
    foreach ($account in $accounts) {
        if ($null -ne $account.MSExchRecipientTypeDetails) {
            $accountsWithMailbox.Add($account, $true)
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

function get-timespan($start, $end) {
    $value = 1
    if ($null -eq $start -or $null -eq $end) {
        return $value
    }
    else {
        $value = New-TimeSpan -Start $start -End $end
    }
    return $value
}

function remove-spanning-deleted-users-licenses([System.Collections.ArrayList]$users = $users, [int]$purge = $DAYS_BEFORE_PURGE, [string]$file = $UNASSIGN_USER_FILE) {
    $deletedUsers = ""
    $count = 0
    $now = Get-Date
    $removeUsers = [System.Collections.ArrayList]@()
    # DEFECT need to account for single item strings
    foreach ($user in $users) {
        $dateDeleted = $user.dateDeleted
        $daySinceDeleted = get-timespan $now $dateDeleted
        if ($user.isDeleted -eq $true -and $daySinceDeleted -gt $purge -and $user.GetType() -ne [Int32]) {
            $deletedUsers += "$($user.userPrincipalName),"
            $removeUsers.Add($user) | Out-Null
            write-host $count $daySinceDeleted
            $count++
        }
    }

    if ($deletedUsers -and $deletedUsers.Count -gt 0) {
        $deletedUsers = $deletedUsers.Trim(",")
        $deletedUsersArray = $deletedUsers.Split(",")
        $bodyObj = New-Object -TypeName psobject
        $bodyObj | Add-Member -MemberType NoteProperty -Name userPrincipalNames -Value $deletedUsersArray -Force
        write-host $bodyObj
        $body += $bodyObj | ConvertTo-Json
        Write-Host $body
        try {
            Invoke-RestMethod -Method Post -Headers $header -ContentType "application/json" -uri $API_UNASSIGN_LICENSE -Body $body
        } catch {
        }
        Write-Host "$count deleted users licenses removed."
    }
    else {
        Write-Host "There were no deleted users to remove licenses from."
    }
    foreach($removeUser in $removeUsers) {
        $users.Remove($removeUser)
    }
    $users | ConvertTo-Json | Out-File $file
}
#remove-spanning-deleted-users-licenses $savedUsers
#endregion

# Get licened Spanning users NOT in AD.
Write-Host "Getting Spanning users NOT in AD."
if (!$go) {
    $usersToDelete = get-spanning-users
    $backupUser = $usersToDelete | Select-Object *
    $go = $true
}
else {
    Write-Host "Using cached users. Set `$go = `$false to pull down the list again."
    $usersToDelete = $backupUser | Select-Object *
}
[System.Collections.ArrayList]$savedUsers = Add-UsersToDelete $UNASSIGN_USER_FILE $usersToDelete
#$savedUsers

# Remove License from deleted Users.
if ($null -ne $savedUsers) {
    Write-Host "Removing licenses from users that have been deleted in the last $DAYS_BEFORE_PURGE days."
    remove-spanning-deleted-users-licenses $savedUsers
}

if ($null -eq $msolUsers) {
    try {
        $msolUsers = Get-m365-Users $m365Cred
    }
    catch { 
        write-host "Failed to connect ot microsoft, check username and password."
        break
    }
}
else {
    Write-Host "Using cached user list. Set `$msolUsers to `$null to refresh the list."
}

$e5UPNs = ($msolUsers | Where-Object `
    { ($_.licenses).AccountSkuId -match "EnterprisePremium" }).UserPrincipalName
$f3UPNs = ($msolUsers | Where-Object `
    { ($_.licenses).AccountSkuId -match "SPE_F1" }).UserPrincipalName

$upns = $e5UPNs + $f3UPNs

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
if ($null -ne $upns) {
    add-spanning-licenses-to-m365LicensedUsers $licensedUsers, $upns
    $e5UPNs = $null
    $f3UPNs = $null
    $upns = $null
}
#>
