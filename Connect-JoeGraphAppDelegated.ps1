Function Connect-JoeGraphAppDelegated {
    
    <#
    .SYNOPSIS
    Connect to the enterprise app registration used to access Graph API.
    .DESCRIPTION
    This function will attempt to use an existing session and perform authentication only if needed.
    .EXAMPLE
    Connect-g46GraphAppDelegated
    #>
    
    #Check for existing sessions
    $MgContextResult = Get-MgContext -ErrorAction SilentlyContinue
    if($null -eq $MgContextResult) {
        try {
            #Connect to graph using Delegated permissions w/Interactive Logon
            if (Connect-MgGraph -TenantId "a1aaf7b2-e297-4cff-9c6f-c27ca192602a" -ClientId "529881d6-a106-4bcf-bee9-0b8f16ac837f") {
            }
        }
            Catch {
                Write-Error $_
            }
        }
    }