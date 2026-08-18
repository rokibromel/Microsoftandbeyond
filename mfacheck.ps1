# Connect to Graph
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All", "Policy.Read.All"

# Ensure export folder exists
$exportPath = "C:\temp\MFAUsers.csv"
if (-not (Test-Path -Path "C:\temp")) {
    New-Item -ItemType Directory -Path "C:\temp" | Out-Null
}

# Get all users (excluding guests)
Write-Host "Retrieving Azure AD users..."
$Users = Get-MgUser -All | Where-Object { $_.UserType -ne "Guest" }

Write-Host "Processing $($Users.Count) users..."
$Report = @()

foreach ($User in $Users) {
    $mfaMethods = Get-MgUserAuthenticationMethod -UserId $User.Id

    $mfaEnabled = if ($mfaMethods.Count -gt 0) { "Enabled" } else { "Disabled" }

    # Determine default MFA method
    $defaultMethod = "Not enabled"
    $phoneNumber = $null

    foreach ($method in $mfaMethods) {
        switch ($method.'@odata.type') {
            # For phone-based methods
            "#microsoft.graph.phoneAuthenticationMethod" {
                $phoneMethod = Get-MgUserAuthenticationPhoneMethod -UserId $User.Id -PhoneAuthenticationMethodId $method.Id
                $phoneNumber = $phoneMethod.PhoneNumber
                if ($phoneMethod.PhoneType -eq "mobile") {
                    $defaultMethod = "Text or Call Authentication Phone"
                }
            }
            "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                $defaultMethod = "Microsoft Authenticator App"
            }
            "#microsoft.graph.fido2AuthenticationMethod" {
                $defaultMethod = "FIDO2 Security Key"
            }
            "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                $defaultMethod = "Windows Hello"
            }
        }
    }

    # Parse proxy addresses
    $primarySMTP = $User.ProxyAddresses | Where-Object { $_ -like "SMTP:*" } | ForEach-Object { $_ -replace "SMTP:", "" }
    $aliases     = $User.ProxyAddresses | Where-Object { $_ -like "smtp:*" } | ForEach-Object { $_ -replace "smtp:", "" }

    # Prepare report line
    $Report += [PSCustomObject]@{
        UserPrincipalName = $User.UserPrincipalName
        DisplayName       = $User.DisplayName
        MFAState          = $mfaEnabled
        MFADefaultMethod  = $defaultMethod
        MFAPhoneNumber    = $phoneNumber
        PrimarySMTP       = ($primarySMTP -join ',')
        Aliases           = ($aliases -join ',')
    }
}

# Output
$Report | Sort-Object UserPrincipalName | Out-GridView
$Report | Sort-Object UserPrincipalName | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
Write-Host "MFA Report exported to $exportPath"
