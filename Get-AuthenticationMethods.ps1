param (
    [Parameter(Mandatory = $true)]
    [string]$csvFileName,
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
    $Report | Export-Csv -Path $csvFileName -NoTypeInformation -Encoding utf8

    Write-Host "Script completed. Report exported successfully to $csvFileName" -ForegroundColor Green
}
catch {
    # Catch errors
    Write-Host "An error occurred: $_" -ForegroundColor Red
}
