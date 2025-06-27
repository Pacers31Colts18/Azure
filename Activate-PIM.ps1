function Activate-PIM {
    <#
.SYNOPSIS
    Activates PIM roles and groups
.DESCRIPTION
    Activates PIM roles and groups using a delegated permissions with an Enterprise Application in Azure. Required permissions:
    PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup
    PrivilegedEligibilitySchedule.Read.AzureADGroup
    RoleAssignmentSchedule.ReadWrite.Directory
    RoleEligibilitySchedule.Read.Directory

.PARAMETER Scope
    Choose between Group or Role for activation.

.PARAMETER Justification
    (Optional) Adds justification to the activation, if not will default to pre-defined text.

.PARAMETER Hours
    Defaults to 4 hours, can change if needed.

.EXAMPLE
    Activate-PIM -scope Role -justifcation "justification" -hours 8

.EXAMPLE
    Activate-PIM -scope Group

.NOTES
    Author: Joe Loveless
    Date: 6/27/2025
    
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("Role", "Group")]
        [string]$Scope,
        [string]$Justification,
        [ValidateNotNullOrEmpty()][string]$Hours = 4
    )

    #region Declarations
    $FunctionName = $MyInvocation.MyCommand.Name.ToString()
    $date = Get-Date -Format yyyyMMdd-HHmm
    if ($outputdir.Length -eq 0) { $outputdir = $pwd }
    $OutputFilePath = "$OutputDir\$FunctionName-$date.csv"
    $LogFilePath = "$OutputDir\$FunctionName-$date.log"
    $graphApiVersion = "beta"
    #endregion

    #region Connect to Graph
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Output "Connecting to Microsoft Graph..."
    try {
        Connect-JoeGraphAppDelegated
        #Connect-MgGraph -TenantId "$tenantID" -ClientId "$clientID"
        Write-Output "Connected to Microsoft Graph"
    }
    catch {
        Write-Error "Failed to connect: $($_.Exception.Message)"
        return
    }
    #endRegion

    #region Get User Info
    try {
        $user = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/$graphApiVersion/me" -Method GET
        Write-Output "User context confirmed: $($user.userPrincipalName)"
    }
    catch {
        Write-Error "Failed to validate user context: $($_.Exception.Message)"
        return
    }
    #endRegion

    #region Scope Handling
    if ($Scope -eq "Role") {
        Write-Host "Fetching eligible role assignments..." -ForegroundColor Cyan
        try {
            $uri = "https://graph.microsoft.com/$graphApiVersion/roleManagement/directory/roleEligibilitySchedules/filterByCurrentUser(on='principal')?`$expand=roleDefinition"
            $roles = (Invoke-MgGraphRequest -Uri $uri -Method GET).value
            if (-not $roles) {
                Write-Warning "No eligible roles found."
                return
            }
        }
        catch {
            Write-Error "Failed to fetch roles: $($_.Exception.Message)"
            return
        }

        $selectableRoles = $roles | ForEach-Object {
            $displayObj = [PSCustomObject]@{
                DisplayName = $_.RoleDefinition.DisplayName
            }
            $displayObj | Add-Member -NotePropertyName OriginalObject -NotePropertyValue $_ -Force
            $displayObj
        }

        $selected = $selectableRoles | Out-GridView -Title "Select a Role to Activate" -PassThru
        if (-not $selected) { Write-Warning "No role selected."; return }

        $selectedRole = $selected.OriginalObject
        if (-not $selectedRole) {
            Write-Warning "No role selected."
            return
        }

        if (-not $justification) {
            $justification = "Activating role: $($selectedRole.DisplayName)"
        }

        $duration = "PT" + $hours + "H"

        $params = @{
            Action           = "selfActivate"
            principalId      = $selectedRole.PrincipalId
            roleDefinitionId = $selectedRole.RoleDefinitionId
            directoryScopeId = $selectedRole.DirectoryScopeId
            justification    = $justification
            scheduleInfo     = @{
                startDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                expiration    = @{
                    type     = "AfterDuration"
                    duration = $duration
                }
            }
            ticketInfo       = @{
                ticketNumber = ""
                ticketSystem = ""
            }
        }

        try {
            $uri = "https://graph.microsoft.com/$graphApiVersion/roleManagement/directory/roleAssignmentScheduleRequests"
            Invoke-MgGraphRequest -Uri $uri -Method POST -Body $params -ContentType "application/json"
            Write-Output "Role activation request submitted."
        }
        catch {
            Write-Error "Activation failed: $($_.Exception.Message)"
            return
        }

        # Confirm activation
        Start-Sleep -Seconds 3
        try {
            $uri = "https://graph.microsoft.com/$graphApiVersion/roleManagement/directory/roleAssignmentScheduleInstances/filterByCurrentUser(on='principal')?`$expand=roleDefinition"
            $activeRoles = (Invoke-MgGraphRequest -Uri $uri -Method GET).value
            if ($activeRoles) {
                Write-Output "Currently active roles:"
                $activeRoles | ForEach-Object {
                    $expiration = if ($_.EndDateTime) { $_.EndDateTime.ToString("yyyy-MM-dd HH:mm:ss UTC") } else { "No Expiration" }
                    Write-Output "$($_.RoleDefinition.DisplayName) (Expires: $expiration)"
                }
            }
        }
        catch {
            Write-Error "Failed to verify active roles: $($_.Exception.Message)"
        }

    }
    elseif ($Scope -eq "Group") {
        Write-Host "Fetching eligible group assignments..." -ForegroundColor Cyan
        try {
            $uri = "https://graph.microsoft.com/$graphApiVersion/identityGovernance/privilegedAccess/group/eligibilitySchedules/filterByCurrentUser(on='principal')?`$expand=group"
            $groups = (Invoke-MgGraphRequest -Uri $uri -Method GET).value
            if (-not $groups) {
                Write-Warning "No eligible groups found."
                return
            }
        }
        catch {
            Write-Error "Failed to fetch groups: $($_.Exception.Message)"
            return
        }

        $selectableGroups = $groups | ForEach-Object {
            $displayObj = [PSCustomObject]@{
                DisplayName = $_.group.DisplayName
            }
            $displayObj | Add-Member -NotePropertyName OriginalObject -NotePropertyValue $_ -Force
            $displayObj
        }

        $selected = $selectableGroups | Out-GridView -Title "Select a Group to Activate" -PassThru
        if (-not $selected) {
            Write-Warning "No group selected."
            return
        }

        $selectedGroup = $selected.OriginalObject


        if (-not $justification) {
            $justification = "Activating group: $($selectedGroup.DisplayName)"
        }

        $duration = "PT" + $hours + "H"

        $params = @{
            accessId      = "member"
            principalId   = $selectedGroup.PrincipalId
            groupId       = $selectedGroup.GroupId
            action        = "AdminAssign"
            justification = $justification
            scheduleInfo  = @{
                startDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                expiration    = @{
                    type     = "AfterDuration"
                    duration = $duration
                }
            }
            ticketInfo    = @{
                ticketNumber = ""
                ticketSystem = ""
            }
        }

        try {
            $uri = "https://graph.microsoft.com/$graphApiVersion/identityGovernance/privilegedAccess/group/assignmentScheduleRequests"
            Invoke-MgGraphRequest -Uri $uri -Method POST -Body $params -ContentType "application/json"
            Write-Output "Group activation request submitted."
        }
        catch {
            Write-Error "Activation failed: $($_.Exception.Message)"
            return
        }

        # Confirm activation
        Start-Sleep -Seconds 3
        try {
            $uri = "https://graph.microsoft.com/$graphApiVersion/identityGovernance/privilegedAccess/group/assignmentScheduleInstances/filterByCurrentUser(on='principal')?`$expand=group"
            $activeGroups = (Invoke-MgGraphRequest -Uri $uri -Method GET).value
            if ($activeGroups) {
                Write-Output "Currently active groups:"
                $activeGroups | ForEach-Object {
                    $expiration = if ($_.EndDateTime) { $_.EndDateTime.ToString("yyyy-MM-dd HH:mm:ss UTC") } else { "No Expiration" }
                    Write-Output "$($_.Group.DisplayName) (Expires: $expiration)"
                }
            }
        }
        catch {
            Write-Error "Failed to verify active groups: $($_.Exception.Message)"
        }
    }
    #endregion
}
