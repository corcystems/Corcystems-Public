<#
    .SYNOPSIS
    Get-AuthenticationMethods.ps1

    .DESCRIPTION
    Export users authentication methods report from Micrososoft Graph and know which MFA method
    is set as default for each user and what MFA methods are registered for each user.

    .LINK
    www.alitajran.com/get-mfa-status-entra/

    .NOTES
    Written by: ALI TAJRAN
    Website:    www.alitajran.com
    LinkedIn:   linkedin.com/in/alitajran

    .CHANGELOG
    V1.00, 10/12/2023 - Initial version
    V1.10, 11/04/2024 - Added parameters to script
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$CSVPath,
    [string]$GroupId = ""
)

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "AuditLog.Read.All", "Group.Read.All"

try {
    # If a GroupId is provided, fetch group members first
    if ($GroupId) {
        Write-Host "Fetching members for GroupId: $GroupId" -ForegroundColor Cyan
        $GroupMembers = Get-MgGroupMember -GroupId $GroupId -All
        $GroupMemberIds = $GroupMembers.Id
    }

    # Fetch user registration detail report from Microsoft Graph
    if ($GroupId) {
        $Users = Get-MgBetaReportAuthenticationMethodUserRegistrationDetail -All | Where-Object { $GroupMemberIds -contains $_.Id }
    }
    else {
        $Users = Get-MgBetaReportAuthenticationMethodUserRegistrationDetail -All
    }

    # Create custom PowerShell object and populate it with the desired properties
    $Report = foreach ($User in $Users) {
        [PSCustomObject]@{
            Id                                           = $User.Id
            UserPrincipalName                            = $User.UserPrincipalName
            UserDisplayName                              = $User.UserDisplayName
            IsAdmin                                      = $User.IsAdmin
            DefaultMfaMethod                             = $User.DefaultMfaMethod
            MethodsRegistered                            = $User.MethodsRegistered -join ','
            IsMfaCapable                                 = $User.IsMfaCapable
            IsMfaRegistered                              = $User.IsMfaRegistered
            IsPasswordlessCapable                        = $User.IsPasswordlessCapable
            IsSsprCapable                                = $User.IsSsprCapable
            IsSsprEnabled                                = $User.IsSsprEnabled
            IsSsprRegistered                             = $User.IsSsprRegistered
            IsSystemPreferredAuthenticationMethodEnabled = $User.IsSystemPreferredAuthenticationMethodEnabled
            LastUpdatedDateTime                          = $User.LastUpdatedDateTime
        }
    }

    # Output custom object to GridView
    $Report | Out-GridView -Title "Authentication Methods Report"

    # Export custom object to CSV file
    $Report | Export-Csv -Path $csvPath\mfaUsers.csv -NoTypeInformation -Encoding utf8

    Write-Host "Script completed. Report exported successfully to $csvPath" -ForegroundColor Green
}
catch {
    # Catch errors
    Write-Host "An error occurred: $_" -ForegroundColor Red
}