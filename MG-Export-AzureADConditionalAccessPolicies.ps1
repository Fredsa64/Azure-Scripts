<#
	.SYNOPSIS  
	Creation date : 16/02/2024
	Author        : Frederic Satori - Inetum
	
	.DESCRIPTION  
	Fonction : Export Entra ID Conditional Access Policies with MS Graph PowerShell (MgGraph)
	
	.NOTES
	Modification history :
	
	+---------------+--------------------+---------+----------------------------------------------------------------------+
	|  Date         |   Author           | Version | Description                                                          |
	+---------------+--------------------+---------+----------------------------------------------------------------------+
	| 16/02/2024	| Frederic Satori    |  1.0    | Create Script                                                        |
	+---------------+--------------------+---------+----------------------------------------------------------------------+
	| 10/06/2024	| Frederic Satori    |  1.1    | Implementing CSV export                                              |
	+---------------+--------------------+---------+----------------------------------------------------------------------+
	
	.OUTPUT
	HTML + CSV files: list of CA Policies

#>

# Debug option
$DebugMode = $false

# Export files
$FilePrefix = "CAPolicy"
$DateTime = $((Get-Date).toString("yyMMdd-HHmm"))
$HTMLExportFile = ".\$FilePrefix-$DateTime.html"
$CSVExportFile = ".\$FilePrefix-$DateTime.csv"

##########################################################################
# Function: Convert Additonal Properties
function Convert-AdditionalProperties {
    [CmdletBinding()]
    param (
        $AdditionalProperties
    )

    if (($null -ne $AdditionalProperties) -and ($AdditionalProperties.Count -gt 0)) {
        $AdditonalPropertiesContent = [PSCustomObject]@()
        foreach ($Property in $AdditionalProperties) {
            $AdditonalPropertiesContent += $Property.Key + " = " + $Property.Value
        }
        $AdditonalPropertiesContent -join ", `r`n"
    }
    else {
        ""
    }
}
##########################################################################

##########################################################################
# Function: Convert CA Policy State
function Convert-State {
    [CmdletBinding()]
    param (
        $State
    )

    switch ($State.ToLower() ) {
        "enabled" { "On" }
        "disabled" { "Off" }
        "enabledforreportingbutnotenforced" { "Report" }
        Default { "" }
    }
}
##########################################################################

# Test Graph module
$GraphModule = Get-Module "Microsoft.Graph" -ListAvailable
If ($null -eq $GraphModule) {
    Write-Host "Microsoft.Graph Module not installed" -ForegroundColor Yellow
    Write-Host "Use: Install-Module -Name Microsoft.Graph -Scope CurrentUser" -ForegroundColor Yellow
    break
}

# Connect-MgGraph
$MgContext = Get-MgContext
If ($null -eq $MgContext) {
    Write-host "Connect-MgGraph"
    Connect-MgGraph -Scopes 'Policy.Read.All', 'Directory.Read.All', 'Application.Read.All'
}
else {
    Write-host "Disconnect-MgGraph"
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-host "Connect-MgGraph"
    Connect-MgGraph -Scopes 'Policy.Read.All', 'Directory.Read.All', 'Application.Read.All'
}

# Tenant Information
$TenantData = Get-MgOrganization
$TenantName = $TenantData.DisplayName
$Date = Get-Date

# Collect Conditional Access policies
Write-Host "Collecting CA Policies" -ForegroundColor Cyan
$CAPolicies = Get-MgIdentityConditionalAccessPolicy -All | Sort-Object DisplayName
if (-not $CAPolicies) {
    Write-Host "No CA policies found. Stopping script." -ForegroundColor Red
    return
}

# Extract CA policy information
$CACSV = [PSCustomObject]@()
$CAHTML = [PSCustomObject]@()
$UsersGroupsRoles = @()
$Apps = @()

Write-host "Extracting CA Policy data" -ForegroundColor Cyan
foreach ($Policy in $CAPolicies) {
    
    # Formating the lists of users/groups/roles
    $UsersGroupsRolesIncludeHTML = ""
    $UsersGroupsRolesExcludeHTML = ""
    $UsersGroupsRolesIncludeCSV = ""
    $UsersGroupsRolesExcludeCSV = ""

    #    Included users
    $UsersInclude = $null
    $UsersInclude = $Policy.Conditions.Users.IncludeUsers
    if ($UsersInclude.Count -gt 0) {
        if ($UsersGroupsRolesIncludeHTML -ne "") { $UsersGroupsRolesIncludeHTML = $UsersGroupsRolesIncludeHTML + "<br>" }
        $UsersGroupsRolesIncludeHTML = $UsersGroupsRolesIncludeHTML + "<b>Users:</b><br>" + ($UsersInclude -join "<br>")
        if ($UsersGroupsRolesIncludeCSV -ne "") { $UsersGroupsRolesIncludeCSV = $UsersGroupsRolesIncludeCSV + "`r`n" }
        $UsersGroupsRolesIncludeCSV = $UsersGroupsRolesIncludeCSV + "Users: `r`n" + ($UsersInclude -join ",")
    }

    #    Included groups
    $GroupsInclude = $null
    $GroupsInclude = $Policy.Conditions.Users.IncludeGroups
    if ($GroupsInclude.Count -gt 0) {
        if ($UsersGroupsRolesIncludeHTML -ne "") { $UsersGroupsRolesIncludeHTML = $UsersGroupsRolesIncludeHTML + "<br>" }
        $UsersGroupsRolesIncludeHTML = $UsersGroupsRolesIncludeHTML + "<b>Groups:</b><br>" + ($GroupsInclude -join "<br>")
        if ($UsersGroupsRolesIncludeCSV -ne "") { $UsersGroupsRolesIncludeCSV = $UsersGroupsRolesIncludeCSV + "`r`n" }
        $UsersGroupsRolesIncludeCSV = $UsersGroupsRolesIncludeCSV + "Groups: `r`n" + ($GroupsInclude -join ",")
    }

    #    Included roles
    $RolesInclude = $null
    $RolesInclude = $Policy.Conditions.Users.IncludeRoles
    if ($RolesInclude.Count -gt 0) {
        if ($UsersGroupsRolesIncludeHTML -ne "") { $UsersGroupsRolesIncludeHTML = $UsersGroupsRolesIncludeHTML + "<br>" }
        $UsersGroupsRolesIncludeHTML = $UsersGroupsRolesIncludeHTML + "<b>Roles:</b><br>" + ($RolesInclude -join "<br>")
        if ($UsersGroupsRolesIncludeCSV -ne "") { $UsersGroupsRolesIncludeCSV = $UsersGroupsRolesIncludeCSV + "`r`n" }
        $UsersGroupsRolesIncludeCSV = $UsersGroupsRolesIncludeCSV + "Roles: `r`n" + ($RolesInclude -join ",")
    }

    #    Excluded users
    $UsersExclude = $null
    $UsersExclude = $Policy.Conditions.Users.ExcludeUsers
    if ($UsersExclude.Count -gt 0) {
        if ($UsersGroupsRolesExcludeHTML -ne "") { $UsersGroupsRolesExcludeHTML = $UsersGroupsRolesExcludeHTML + "<br>" }
        $UsersGroupsRolesExcludeHTML = $UsersGroupsRolesExcludeHTML + "<b>Users:</b><br>" + ($UsersExclude -join "<br>")
        if ($UsersGroupsRolesExcludeCSV -ne "") { $UsersGroupsRolesExcludeCSV = $UsersGroupsRolesExcludeCSV + "`r`n" }
        $UsersGroupsRolesExcludeCSV = $UsersGroupsRolesExcludeCSV + "Users: `r`n" + ($UsersExclude -join ",")
    }

    #    Excluded groups
    $GroupsExclude = $null
    $GroupsExclude = $Policy.Conditions.Users.ExcludeGroups
    if ($GroupsExclude.Count -gt 0) {
        if ($UsersGroupsRolesExcludeHTML -ne "") { $UsersGroupsRolesExcludeHTML = $UsersGroupsRolesExcludeHTML + "<br>" }
        $UsersGroupsRolesExcludeHTML = $UsersGroupsRolesExcludeHTML + "<b>Groups:</b><br>" + ($GroupsExclude -join "<br>")
        if ($UsersGroupsRolesExcludeCSV -ne "") { $UsersGroupsRolesExcludeCSV = $UsersGroupsRolesExcludeCSV + "`r`n" }
        $UsersGroupsRolesExcludeCSV = $UsersGroupsRolesExcludeCSV + "Groups: `r`n" + ($GroupsExclude -join ",")
    }

    #    Excluded roles
    $RolesExclude = $null
    $RolesExclude = $Policy.Conditions.Users.ExcludeRoles
    if ($RolesExclude.Count -gt 0) {
        if ($UsersGroupsRolesExcludeHTML -ne "") { $UsersGroupsRolesExcludeHTML = $UsersGroupsRolesExcludeHTML + "<br>" }
        $UsersGroupsRolesExcludeHTML = $UsersGroupsRolesExcludeHTML + "<b>Roles:</b><br>" + ($RolesExclude -join "<br>")
        if ($UsersGroupsRolesExcludeCSV -ne "") { $UsersGroupsRolesExcludeCSV = $UsersGroupsRolesExcludeCSV + "`r`n" }
        $UsersGroupsRolesExcludeCSV = $UsersGroupsRolesExcludeCSV + "Roles: `r`n" + ($RolesExclude -join ",")
    }
 
    # Grouping the lists of users/groups/roles
    $UsersGroupsRolesInclude = $null
    $UsersGroupsRolesInclude = $UsersInclude
    $UsersGroupsRolesInclude += $GroupsInclude
    $UsersGroupsRolesInclude += $RolesInclude

    $UsersGroupsRolesExclude = $null
    $UsersGroupsRolesExclude = $UsersExclude
    $UsersGroupsRolesExclude += $GroupsExclude
    $UsersGroupsRolesExclude += $RolesExclude
 
    $UsersGroupsRoles += $UsersGroupsRolesInclude
    $UsersGroupsRoles += $UsersGroupsRolesExclude

    # Getting the lists of apps
    $AppsInclude = $null
    $AppsInclude = $Policy.Conditions.Applications.IncludeApplications

    $AppsExclude = $null
    $AppsExclude = $Policy.Conditions.Applications.ExcludeApplications

    $Apps += $AppsInclude
    $Apps += $AppsExclude

    # Adding CSV table data row
    $CACSV += New-Object PSObject -Property @{

        '1 Policy Info'                                                                                                    = "";
        '  1.1 Policy name'                                                                                                = $Policy.DisplayName;
        '  1.2 Policy ID'                                                                                                  = $Policy.ID;
        '  1.3 Description'                                                                                                = $Policy.Description;
        '  1.4 Creation date'                                                                                              = $Policy.CreatedDateTime;
        '  1.5 Modification date'                                                                                          = $Policy.ModifiedDateTime;
        '  1.6 State'                                                                                                      = (Convert-State $Policy.State);

        '2 Assignments'                                                                                                    = "";
        '  2.1 Users or workloads identities'                                                                              = "";
        '    2.1.1 Assignments/Users/Include'                                                                              = $UsersGroupsRolesIncludeCSV;
        '    2.1.2 Assignments/Users/Exclude'                                                                              = $UsersGroupsRolesExcludeCSV;
        '  2.2 Target resources'                                                                                           = "";
        '    2.2.1 Cloud apps'                                                                                             = "";
        '      2.2.1.1 Assignments/Target Resources/Cloud apps/Include'                                                    = ($AppsInclude -join ", `r`n");
        '      2.2.2.2 Assignments/Target Resources/Cloud apps/Exclude'                                                    = ($AppsExclude -join ", `r`n");
        '    2.2.2 Assignments/Target Resources/User Actions'                                                              = ($Policy.Conditions.Applications.IncludeUserActions -join ", `r`n");
        '    2.2.3 Assignments/Target Resources/Authentication context'                                                    = ($Policy.Conditions.Applications.IncludeAuthenticationContextClassReferences -join ", `r`n");
        '  2.3 Conditions'                                                                                                 = "";
        '    2.3.1 Assignments/Conditions/User risk'                                                                       = ($Policy.Conditions.UserRiskLevels -join ", `r`n");
        '    2.3.2 Assignments/Conditions/Sign-in risk'                                                                    = ($Policy.Conditions.SignInRiskLevels -join ", `r`n");
        '    2.3.3 Device platforms'                                                                                       = "";
        '      2.3.3.1 Assignments/Conditions/Device platforms/Include'                                                    = ($Policy.Conditions.Platforms.IncludePlatforms -join ", `r`n");
        '      2.3.3.2 Assignments/Conditions/Device platforms/Exclude'                                                    = ($Policy.Conditions.Platforms.ExcludePlatforms -join ", `r`n");
        '    2.3.4 Locations'                                                                                              = "";
        '      2.3.4.1 Assignments/Conditions/Locations/Include'                                                           = ($Policy.Conditions.Locations.Includelocations -join ", `r`n");
        '      2.3.4.2 Assignments/Conditions/Locations/Exclude'                                                           = ($Policy.Conditions.Locations.Excludelocations -join ", `r`n");
        '    2.3.5 Assignments/Conditions/Client apps'                                                                     = ($Policy.Conditions.ClientAppTypes -join ", `r`n");
        '    2.3.6 Filter for devices'                                                                                     = "";
        '      2.3.6.1 Assignments/Conditions/Filter for devices/Include'                                                  = ($Policy.Conditions.Devices.IncludeDevices -join ", `r`n");
        '      2.3.6.2 Assignments/Conditions/Filter for devices/Exclude'                                                  = ($Policy.Conditions.Devices.ExcludeDevices -join ", `r`n");
        '      2.3.6.3 Assignments/Conditions/Filter for devices/Rule syntax'                                              = ($Policy.Conditions.Devices.DeviceFilter.Rule -join ", `r`n");
     
        '3 Access controls'                                                                                                = "";
        '  3.1 Grant'                                                                                                      = "";
        '    3.1.1 Access controls/Grant/BuiltInControls'                                                                  = $($Policy.GrantControls.BuiltInControls);
        '    3.1.2 Access controls/Grant/TermsOfUse'                                                                       = $($Policy.GrantControls.TermsOfUse);
        '    3.1.3 Access controls/Grant/CustomControls'                                                                   = $($Policy.GrantControls.CustomAuthenticationFactors);
        '    3.1.4 Access controls/Grant/For multiple controls'                                                            = $Policy.GrantControls.Operator;
        '    3.1.5 Require authentication strength'                                                                        = "";
        '      3.1.5.1 Access controls/Grant/Require authentication strength/Authentication strength name'                 = $Policy.GrantControls.AuthenticationStrength.DisplayName;
        '      3.1.5.2 Access controls/Grant/Require authentication strength/Authentication strength policy type'          = $Policy.GrantControls.AuthenticationStrength.PolicyType;
        '      3.1.5.3 Access controls/Grant/Require authentication strength/Authentication strength description'          = $Policy.GrantControls.AuthenticationStrength.Description;
        '      3.1.5.4 Access controls/Grant/Require authentication strength/Authentication strength allowed combinations' = ($Policy.GrantControls.AuthenticationStrength.AllowedCombinations -join ", `r`n");
    
        '4 Session'                                                                                                        = "";
        '  4.1 Session/Additional properties'                                                                              = (Convert-AdditionalProperties $Policy.SessionControls.AdditionalProperties);
        '  4.2 Use app enforced restrictions'                                                                              = "";
        '    4.2.1 Session/Use app enforced restrictions/Enabled'                                                          = $Policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled;
        '    4.2.2 Session/Use app enforced restrictions/Additional properties'                                            = (Convert-AdditionalProperties $Policy.SessionControls.ApplicationEnforcedRestrictions.AdditionalProperties);
        '  4.3 Use Conditional Access App Control'                                                                         = "";
        '    4.3.1 Session/Use Conditional Access App Control/Enabled'                                                     = $Policy.SessionControls.CloudAppSecurity.IsEnabled;
        '    4.3.2 Session/Use Conditional Access App Control/Security type'                                               = $Policy.SessionControls.CloudAppSecurity.CloudAppSecurityType;
        '    4.3.3 Session/Use Conditional Access App Control/Additional properties'                                       = (Convert-AdditionalProperties $Policy.SessionControls.CloudAppSecurity.AdditionalProperties);
        '  4.4 Disable resilience defaults'                                                                                = "";
        '    4.4.1 Session/Disable resilience defaults/Enabled'                                                            = $Policy.SessionControls.DisableResilienceDefaults;
        '  4.5 Persistent browser session'                                                                                 = "";
        '    4.5.1 Session/Persistent browser session/Enabled'                                                             = $Policy.SessionControls.PersistentBrowser.IsEnabled;
        '    4.5.2 Session/Persistent browser session/Browser mode'                                                        = $Policy.SessionControls.PersistentBrowser.Mode;
        '    4.5.3 Session/Persistent browser session/Additional properties'                                               = (Convert-AdditionalProperties $Policy.SessionControls.PersistentBrowser.AdditionalProperties);
        '  4.6 Sign-in frequency'                                                                                          = "";
        '    4.6.1 Session/Sign-in frequency/Enabled'                                                                      = $Policy.SessionControls.SignInFrequency.IsEnabled;
        '    4.6.2 Session/Sign-in frequency/Authentication type'                                                          = $Policy.SessionControls.SignInFrequency.AuthenticationType;
        '    4.6.3 Session/Sign-in frequency/Interval'                                                                     = $Policy.SessionControls.SignInFrequency.FrequencyInterval;
        '    4.6.4 Session/Sign-in frequency/Frequency type'                                                               = $Policy.SessionControls.SignInFrequency.Type;
        '    4.6.5 Session/Sign-in frequency/Frequency value'                                                              = $Policy.SessionControls.SignInFrequency.Value;
        '    4.6.6 Session/Sign-in frequency/Additional properties'                                                        = (Convert-AdditionalProperties $Policy.SessionControls.SignInFrequency.AdditionalProperties)
        
    }

    # Adding HTML table data row
    $CAHTML += "
        <tr>
            <td>" + $Policy.DisplayName + "</td>
            <td>" + $Policy.ID + "</td>
            <td>" + $Policy.Description + "</td>
            <td>" + $Policy.CreatedDateTime + "</td>
            <td>" + $Policy.ModifiedDateTime + "</td>
            <td>" + (Convert-State $Policy.State) + "</td>
            <td>" + $UsersGroupsRolesIncludeHTML + "</td>
            <td>" + $UsersGroupsRolesExcludeHTML + "</td>
            <td>" + ($AppsInclude -join "<br>") + "</td>
            <td>" + ($AppsExclude -join "<br>") + "</td>
            <td>" + ($Policy.Conditions.Applications.IncludeUserActions -join "<br>") + "</td>
            <td>" + ($Policy.Conditions.Applications.IncludeAuthenticationContextClassReferences -join "<br>") + "</td>
            <td>" + ($Policy.Conditions.UserRiskLevels -join "<br>") + "</td>
            <td>" + ($Policy.Conditions.SignInRiskLevels -join "<br>") + "</td>
            <td>" + ($Policy.Conditions.Platforms.IncludePlatforms -join "<br>") + "</td>
            <td>" + ($Policy.Conditions.Platforms.ExcludePlatforms -join "<br>") + "</td>
            <td>" + ($Policy.Conditions.Locations.Includelocations -join "<br>") + "</td>
            <td>" + ($Policy.Conditions.Locations.Excludelocations -join "<br>") + "</td>
            <td>" + ($Policy.Conditions.ClientAppTypes -join "<br>") + "</td>
            <td>" + ($Policy.Conditions.Devices.IncludeDevices -join "<br>") + "</td>
            <td>" + ($Policy.Conditions.Devices.ExcludeDevices -join "<br>") + "</td>
            <td>" + ($Policy.Conditions.Devices.DeviceFilter.Rule -join "<br>") + "</td>
            <td>" + $($Policy.GrantControls.BuiltInControls) + "</td>
            <td>" + $($Policy.GrantControls.TermsOfUse) + "</td>
            <td>" + $($Policy.GrantControls.CustomAuthenticationFactors) + "</td>
            <td>" + $Policy.GrantControls.Operator + "</td>
            <td>" + $Policy.GrantControls.AuthenticationStrength.DisplayName + "</td>
            <td>" + $Policy.GrantControls.AuthenticationStrength.PolicyType + "</td>
            <td>" + $Policy.GrantControls.AuthenticationStrength.Description + "</td>
            <td>" + ($Policy.GrantControls.AuthenticationStrength.AllowedCombinations -join "<br>") + "</td>
            <td>" + $Policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled + "</td>
            <td>" + (Convert-AdditionalProperties $Policy.SessionControls.ApplicationEnforcedRestrictions.AdditionalProperties) + "</td>
            <td>" + $Policy.SessionControls.CloudAppSecurity.IsEnabled + "</td>
            <td>" + $Policy.SessionControls.CloudAppSecurity.CloudAppSecurityType + "</td>
            <td>" + (Convert-AdditionalProperties $Policy.SessionControls.CloudAppSecurity.AdditionalProperties) + "</td>
            <td>" + $Policy.SessionControls.DisableResilienceDefaults + "</td>
            <td>" + $Policy.SessionControls.PersistentBrowser.IsEnabled + "</td>
            <td>" + $Policy.SessionControls.PersistentBrowser.Mode + "</td>
            <td>" + (Convert-AdditionalProperties $Policy.SessionControls.PersistentBrowser.AdditionalProperties) + "</td>
            <td>" + $Policy.SessionControls.SignInFrequency.IsEnabled + "</td>
            <td>" + $Policy.SessionControls.SignInFrequency.AuthenticationType + "</td>
            <td>" + $Policy.SessionControls.SignInFrequency.FrequencyInterval + "</td>
            <td>" + $Policy.SessionControls.SignInFrequency.Type + "</td>
            <td>" + $Policy.SessionControls.SignInFrequency.Value + "</td>
            <td>" + (Convert-AdditionalProperties $Policy.SessionControls.SignInFrequency.AdditionalProperties) + "</td>
            <td>" + (Convert-AdditionalProperties $Policy.SessionControls.AdditionalProperties) + "</td>
        </tr>"
}

if ($DebugMode -ne $true) {

    # Swith user/group Guid to display names
    Write-host "Converting Entra ID Guids" -ForegroundColor Cyan
    # Filter out Objects
    $CAJson = $CACSV | ConvertTo-Json -Depth 4
    $ADSearch = $UsersGroupsRoles | Where-Object { $_ -ne 'All' -and $_ -ne 'GuestsOrExternalUsers' -and $_ -ne 'None' }
    $ADNames = @{}
    $ObjectList = [PSCustomObject]@()

    Get-MgDirectoryObjectById -ids $ADSearch | ForEach-Object {
        $ObjectId = $_.Id
        $ObjectName = $_.AdditionalProperties.displayName
        $ADNames.$ObjectId = $ObjectName
        $CAJson = $CAJson -replace "$ObjectId", "$ObjectName"
        $ObjectList += New-Object PSObject -Property @{
            Id   = $ObjectId;
            Name = $ObjectName
        }
    }
    # Switch Apps Guid with Display names
    $AllApps = Get-MgServicePrincipal -All
    $AllApps | Where-Object { $_.AppId -in $Apps } | ForEach-Object {
        $ObjectId = $_.AppId
        $ObjectName = $_.DisplayName
        $CAJson = $CAJson -replace "$ObjectId", "$ObjectName"
        $ObjectList += New-Object PSObject -Property @{
            Id   = $ObjectId;
            Name = $ObjectName
        }
    }
    # Switch named location Guid for Display Names
    Get-MgIdentityConditionalAccessNamedLocation | ForEach-Object {
        $ObjectId = $_.Id
        $ObjectName = $_.DisplayName
        $CAJson = $CAJson -replace "$ObjectId", "$ObjectName"
        $ObjectList += New-Object PSObject -Property @{
            Id   = $ObjectId;
            Name = $ObjectName
        }
    }
    # Switch Roles Guid to Names
    Get-MgDirectoryRoleTemplate | ForEach-Object {
        $ObjectId = $_.Id
        $ObjectName = $_.DisplayName
        $CAJson = $CAJson -replace "$ObjectId", "$ObjectName"
        $ObjectList += New-Object PSObject -Property @{
            Id   = $ObjectId;
            Name = $ObjectName
        }
    }

    foreach ($Object in $ObjectList) {
        $CAJson = $CAJson -replace $Object.Id, $Object.Name
    }
    $CACSV = $CAJson | ConvertFrom-Json

}

# Column Sorting Order
$Sort = `
    "  1.1 Policy name", `
    "  1.2 Policy ID", `
    "  1.3 Description", `
    "  1.4 Creation date", `
    "  1.5 Modification date", `
    "  1.6 State", `
    "    2.1.1 Assignments/Users/Include", `
    "    2.1.2 Assignments/Users/Exclude", `
    "      2.2.1.1 Assignments/Target Resources/Cloud apps/Include", `
    "      2.2.2.2 Assignments/Target Resources/Cloud apps/Exclude", `
    "    2.2.2 Assignments/Target Resources/User Actions", `
    "    2.2.3 Assignments/Target Resources/Authentication context", `
    "    2.3.1 Assignments/Conditions/User risk", `
    "    2.3.2 Assignments/Conditions/Sign-in risk", `
    "      2.3.3.1 Assignments/Conditions/Device platformsInclude", `
    "      2.3.3.2 Assignments/Conditions/Device platformsExclude", `
    "      2.3.4.1 Assignments/Conditions/Locations/Include", `
    "      2.3.4.2 Assignments/Conditions/Locations/Exclude", `
    "    2.3.5 Assignments/Conditions/Client apps", `
    "      2.3.6.1 Assignments/Conditions/Filter for devices/Include", `
    "      2.3.6.2 Assignments/Conditions/Filter for devices/Exclude", `
    "      2.3.6.3 Assignments/Conditions/Filter for devices/Rule syntax", `
    "    3.1.1 Access controls/Grant/BuiltInControls", `
    "    3.1.2 Access controls/Grant/TermsOfUse", `
    "    3.1.3 Access controls/Grant/CustomControls", `
    "    3.1.4 Access controls/Grant/For multiple controls", `
    "      3.1.5.1 Access controls/Grant/Require authentication strength/Authentication strength name", `
    "      3.1.5.2 Access controls/Grant/Require authentication strength/Authentication strength policy type", `
    "      3.1.5.3 Access controls/Grant/Require authentication strength/Authentication strength description", `
    "      3.1.5.4 Access controls/Grant/Require authentication strength/Authentication strength allowed combinations", `
    "  4.1 Session/Additional properties", `
    "    4.2.1 Session/Use app enforced restrictions/Enabled", `
    "    4.2.2 Session/Use app enforced restrictions/Additional properties", `
    "    4.3.1 Session/Use Conditional Access App Control/Enabled", `
    "    4.3.2 Session/Use Conditional Access App Control/Security type", `
    "    4.3.3 Session/Use Conditional Access App Control/Additional properties", `
    "    4.4.1 Session/Disable resilience defaults/Enabled", `
    "    4.5.1 Session/Persistent browser session/Enabled", `
    "    4.5.2 Session/Persistent browser session/Browser mode", `
    "    4.5.3 Session/Persistent browser session/Additional properties", `
    "    4.6.1 Session/Sign-in frequency/Enabled", `
    "    4.6.2 Session/Sign-in frequency/Authentication type", `
    "    4.6.3 Session/Sign-in frequency/Interval", `
    "    4.6.4 Session/Sign-in frequency/Frequency type", `
    "    4.6.5 Session/Sign-in frequency/Frequency value", `
    "    4.6.6 Session/Sign-in frequency/Additional properties"

# CSV Export
Write-host "Saving to CSV file: $CSVExportFile" -ForegroundColor Cyan
$CACSV | Select-Object $Sort | Sort-Object "  1.1 Policy name" | Export-CSV -Path $CSVExportFile -Delimiter ';' -NoTypeInformation -Encoding UTF8

# HTML Export
Write-host "Saving to HTML file: $HTMLExportFile" -ForegroundColor Cyan

# HTML script code
$JQuery = '<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
	<script>
	$(document).ready(function(){
		$("tr").click(function(){
		if(!$(this).hasClass("selected")){
			$(this).addClass("selected");
		} else {
			$(this).removeClass("selected");
		}

		});
		$("th").click(function(){
		if(!$(this).hasClass("colselected")){
			$(this).addClass("colselected");
		} else {
			$(this).removeClass("colselected");
		}

		});
	});
	</script>'

# HTML file header
$HTMLHeader = "<!DOCTYPE html>
<html>
<head>
    <base href='https://docs.microsoft.com/' target='_blank'>
	$JQuery
    <style>
	.header{
		position: sticky;
		top: 0px;
        left: 0px;
	}
	.title{
		display: block;
		font-size: 1em;
		margin-block-start: 0.67em;
		margin-block-end: 0.67em;
		margin-inline-start: 0px;
		margin-inline-end: 0px;
		font-weight: bold;
		font-family: Segoe UI;
	}
	table{
		border-collapse: collapse;
		margin: 25px 0;
		font-size: 0.9em;
		font-family: Segoe UI;
        box-shadow: 0 0 20px rgba(0, 0, 0, 0.15) ;
		text-align: center;
	}
	thead tr {
		background-color: #009879;
		color: #ffffff;
		text-align: left;
	}
	th, td {
		min-width: 150px;
		max-width: 1600px;
		padding: 12px 15px;
		border: 1px solid lightgray;
	}
	th {
		vertical-align: bottom;
	}
    td {
        vertical-align: top;
	}
	tbody tr {
		border-bottom: 1px solid #dddddd;
	    <!--background-color: #f3f3f3;-->
        text-align:left;
	}
	tbody tr:last-of-type {
		border-bottom: 2px solid #009879;
	}
	tr:hover{
	background-color: #FFF7C8!important;
	}
	.selected:not(th){
		background-color:#FFF7C8!important;
	}
	th{
        color: white;
		text-align:center;
		font-weight: bolder;
	}
	.colselected {
	    background-color: rgb(93, 236, 213)!important;
	}
	table tr th:first-child {
        color: white;
        white-space: pre-wrap;
		inset-inline-start: 0;
		font-weight: bolder;
		text-align: center;
	}
	table tr td:first-child {
        white-space: pre-wrap;
		inset-inline-start: 0;
		font-weight: bolder;
		text-align: left;
	}
	</style>
</head>
<body>
    <div class=""title"">Conditional Access Policies: $Tenantname - $Date </div>"

# HTML header background colors
$Blue1 = "#000099"
$Blue2 = "#666699"
$Blue3 = "#6666FF"
$Blue4 = "#9999FF"
$Green1 = "#009900"
$Green2 = "#669966"
$Green3 = "#66FF66"
$Green4 = "#99FF99"
$Red1 = "#990000"
$Red2 = "#996666"
$Red3 = "#FF6666"
$Red4 = "#FF9999"
$Gray1 = "#666666"
$Gray2 = "#999999"
$Gray3 = "#CCCCCC"
$Gray4 = "#333333"


# Building HTML table header rows
$HTMLTableHeaderRows = "
    <table>
        <!-- Table header rows -->
        <thead class=""header"">
            <!-- 1st header row -->
            <tr>
                <th colspan=""6"" color=""white"" bgcolor=""$Green1"">Policy Info</th>
                <th colspan=""16"" color=""white"" bgcolor=""$Red1"">Assignments</th>
                <th colspan=""24"" color=""white"" bgcolor=""$Blue1"">Access controls</th>
            </tr>
            <!-- 2nd header row -->
            <tr>
                <th rowspan=""3"" bgcolor=""$Green2"">Policy Name</th>
                <th rowspan=""3"" bgcolor=""$Green2"">Policy ID</th>
                <th rowspan=""3"" bgcolor=""$Green2"">Description</th>
                <th rowspan=""3"" bgcolor=""$Green2"">Creation date</th>
                <th rowspan=""3"" bgcolor=""$Green2"">Modification date</th>
                <th rowspan=""3"" bgcolor=""$Green2"">State</th>
                <th colspan=""2"" bgcolor=""$Red2"">Users or workloads identities</th>
                <th colspan=""4"" bgcolor=""$Red2"">Target resources</th>
                <th colspan=""10"" bgcolor=""$Red2"">Conditions</th>
    
                <th colspan=""8"" color=""white"" bgcolor=""$Blue2"">Grant</th>
                <th colspan=""16"" color=""white"" bgcolor=""$Blue2"">Session</th>

    
            </tr>
            <!-- 3rd header row -->
            <tr>
                <th rowspan=""2"" bgcolor=""$Red3"">Include (Users/Groups/Directory roles)</th>
                <th rowspan=""2"" bgcolor=""$Red3"">Exclude (Users/Groups/Directory roles)</th>
                <th colspan=""2"" bgcolor=""$Red3"">Cloud apps</th>
                <th rowspan=""2"" bgcolor=""$Red3"">User actions</th>
                <th rowspan=""2"" bgcolor=""$Red3"">Authentication context</th>
                <th rowspan=""2"" bgcolor=""$Red3"">User risk</th>
                <th rowspan=""2"" bgcolor=""$Red3"">Sign-in risk</th>
                <th colspan=""2"" bgcolor=""$Red3"">Device platforms</th>
                <th colspan=""2"" bgcolor=""$Red3"">Locations</th>
                <th rowspan=""2"" bgcolor=""$Red3"">Client apps</th>
                <th colspan=""3"" bgcolor=""$Red3"">Filter for devices</th>
                <th rowspan=""2"" bgcolor=""$Blue3"">BuiltInControls</th>
                <th rowspan=""2"" bgcolor=""$Blue3"">TermsOfUse</th>
                <th rowspan=""2"" bgcolor=""$Blue3"">CustomControls</th>
                <th rowspan=""2"" bgcolor=""$Blue3"">For multiple controls</th>
                <th colspan=""4""  bgcolor=""$Blue3"">Require authentication strength</th>
                <th colspan=""2""  bgcolor=""$Blue3"">Use app enforced restrictions</th>
                <th colspan=""3""  bgcolor=""$Blue3"">Use conditional access app control</th>
                <th colspan=""1""  bgcolor=""$Blue3"">Disable resilience defaults</th>
                <th colspan=""3""  bgcolor=""$Blue3"">Persistent browser session</th>
                <th colspan=""6""  bgcolor=""$Blue3"">Sign-in frequency</th>
                <th rowspan=""2"" bgcolor=""$Blue3"">Additional properties</th>
            </tr>
            <!-- 4th header row -->
            <tr>
                <th bgcolor=""$Red4"">Include</th>
                <th bgcolor=""$Red4"">Exclude</th>
                <th bgcolor=""$Red4"">Include</th>
                <th bgcolor=""$Red4"">Exclude</th>
                <th bgcolor=""$Red4"">Include</th>
                <th bgcolor=""$Red4"">Exclude</th>
                <th bgcolor=""$Red4"">Include</th>
                <th bgcolor=""$Red4"">Exclude</th>
                <th bgcolor=""$Red4"">Rule syntax</th>
                <th bgcolor=""$Blue4"">Authentication strength name</th>
                <th bgcolor=""$Blue4"">Authentication strength policy type</th>
                <th bgcolor=""$Blue4"">Authentication strength description</th>
                <th bgcolor=""$Blue4"">Authentication strength allowed combinations</th>
                <th bgcolor=""$Blue4"">Enabled</th>
                <th bgcolor=""$Blue4"">Additional properties</th>
                <th bgcolor=""$Blue4"">Security type</th>
                <th bgcolor=""$Blue4"">Enabled</th>
                <th bgcolor=""$Blue4"">Additional properties</th>
                <th bgcolor=""$Blue4"">Enabled</th>
                <th bgcolor=""$Blue4"">Enabled</th>
                <th bgcolor=""$Blue4"">Browser mode</th>
                <th bgcolor=""$Blue4"">Additional properties</th>
                <th bgcolor=""$Blue4"">Enabled</th>
                <th bgcolor=""$Blue4"">Authentication type</th>
                <th bgcolor=""$Blue4"">Interval</th>
                <th bgcolor=""$Blue4"">Frequency type</th>
                <th bgcolor=""$Blue4"">Frequency value</th>
                <th bgcolor=""$Blue4"">Additional properties</th>
            </tr>
        </thead
        <!-- Table data rows -->
        <tbody>"

$HTMLLastLine = "
        </tbody>
    </table>
</body>
</html>"

# Creating HTML code: file header
$HTMLCode = $HTMLHeader

# Adding header rows
$HTMLCode += $HTMLTableHeaderRows

# Adding data rows
foreach ($HTMLTableDataRow in $CAHTML) {
    $HTMLCode += $HTMLTableDataRow
}

# Closing HTML table
$HTMLCode += $HTMLLastLine

# Converting Entra Id object GUIDs to names
foreach ($Object in $ObjectList) {
    $HTMLCode = $HTMLCode -replace $Object.Id, $Object.Name
}

# Exporting HTML code to file
$HTMLCode | Out-File $HTMLExportFile

Write-host "Opening html export file" -ForegroundColor Cyan
start-process $HTMLExportFile
