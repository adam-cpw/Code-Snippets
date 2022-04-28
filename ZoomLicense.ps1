$licensedGroup = "563d7596-02fd-49b0-8efc-2ae111d079d9"
$basicGroup = "754979cf-a24d-4d42-9ff4-bf8e1e88a894"

. $PSScriptRoot\Utils.ps1

Install-TryModule "CredentialManager"
Install-TryModule "AzureAD"

# Get API key from Windows Credential Manager
$creds = Get-StoredCredential -Target 'AzureAD'

if (!$creds) {
    Write-Host "Error: No Azure Credential found in Windows Credential Manager" -ForegroundColor Red
    pause
    exit 1
}

try {
    Connect-AzureAD -Credential $creds
} catch {
    Write-Host "Error: Could not connect to Azure AD" -ForegroundColor Red
    pause
    exit 1
}

$upn = Read-Host -Prompt "Enter UPN (or email address) of user"

try {
    $user = Get-AzureADUser -ObjectID $upn
} catch {
    Write-Host "Error: Could not find user" -ForegroundColor Red
    pause
    exit 1
}

try {
    $groups = Get-AzureADUserMembership -ObjectId $user.ObjectId | Where-object { $_.ObjectType -eq "Group" } | Select-Object -ExpandProperty ObjectId

    $hasBasic = [bool]($groups -match $basicGroup)

    $hasLicensed = [bool]($groups -match $licensedGroup)
} catch {
    Write-Host "Error: Could not get user groups" -ForegroundColor Red
    pause
    exit 1
}

Write-Host "Current Zoom account status for $upn"
Write-Host "Basic: $hasBasic"
Write-Host "Licensed: $hasLicensed"

if ($hasBasic -and $hasLicensed) {
    Write-Host "Error: User is currently in both the Basic and Licensed groups. Not continuing" -ForegroundColor Red
    pause
    exit 1
}

$validResp = $false
while (!$validResp) {
    $resp = Read-Host "Enter B for Basic Group, L for Licensed Group, N for no license, or Q to quit"
        if ($resp -eq "B") {
            if ($hasBasic) {
                Write-Host "User is already in the Basic group" -ForegroundColor Yellow
            } else {
                # Remove from Licensed group
                if ($hasLicensed) {
                    Remove-AzureADGroupMember -MemberId $user.ObjectId -ObjectId $licensedGroup
                }
                # Add to Basic group
                Add-AzureADGroupMember -RefObjectId $user.ObjectId -ObjectId $basicGroup

                # We've done something!
                Write-Host "Assigned user to Basic group" -ForegroundColor Green
                $validResp = $true
            }
        }
        if ($resp -eq "L") {
            if ($hasLicensed) {
                Write-Host "User is already in the Licensed group" -ForegroundColor Yellow
            } else {
                # Remove from Basic group
                if ($hasBasic) {
                    Remove-AzureADGroupMember -MemberId $user.ObjectId -ObjectId $basicGroup
                }
                # Add to Licensed group
                Add-AzureADGroupMember -RefObjectId $user.ObjectId -ObjectId $licensedGroup

                # We've done something!
                Write-Host "Assigned user to Licenced group" -ForegroundColor Green
                $validResp = $true
            }
        }

        if ($resp -eq "N") {
            if ($hasBasic) {
                # Remove from Basic group
                Remove-AzureADGroupMember -MemberId $user.ObjectId -ObjectId $basicGroup
            }
            if ($hasLicensed) {
                # Remove from Licensed group
                Remove-AzureADGroupMember -MemberId $user.ObjectId -ObjectId $licensedGroup
            }

            # We've done something!
            Write-Host "Removed User from both groups" -ForegroundColor Green
            $validResp = $true
        }

        if ($resp -eq "Q") {
            Write-Host "Exiting" -ForegroundColor Green
            exit
        }

}

pause