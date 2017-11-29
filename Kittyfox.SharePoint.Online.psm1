# .PSObject.TypeNames.Insert(0,'Kittyfox.SharePoint.Online.ListItem')

trap { throw }

try {
    #Add references to SharePoint client assemblies and authenticate to Office 365 site
    Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.Runtime.dll"
} catch {
    Write-Error "Could not load required assemblies for SharePoint Online CSOM functionality. Please install the SharePoint Online Management Shell."
    Write-Error "Download URL: http://go.microsoft.com/fwlink/p/?LinkId=255251"
    exit
}

[Microsoft.SharePoint.Client.SharePointOnlineCredentials] $SPOnline_Credentials = $null

Function Set-SPOCredentials
{
    <#
        .SYNOPSIS
            Set-SPOCredentials stores credential information for use with the Kittyfox.SharePoint.Online module commandlets.
        .DESCRIPTION
            Set-SPOCredentials is an initialization commandlet that prepares your script environment for using the
            remaining commandlets in the Kittyfox.SharePoint.Online module 
        .PARAMETER Credentials
            The Credentials parameter allows the passage of Username and Passwords through a PSCredential object, as gathered
            from another commandlet like Get-Credentials
        .PARAMETER Username
            The Username parameter accepts the username of the target account to be used with the SPO Module
        .PARAMETER Password
            The Password parameter is a secure string version of the password to be used with the 
        .EXAMPLE
            Set-SPOCredentials -Credentials (Get-Credentials)

            This example takes the output from the Get-Credentials dialog prompt and feeds it into the Set-SPOCredentials Commandlet.
        .EXAMPLE
            Set-SPOCredentials -Username 'scadmin@mytenant.onmicrosoft.com' -Password ("Passw0rd" | ConvertTo-SecureString -AsPlaintext -Force)

            This example takes uses the username 'scadmin@mytenant.onmicrosoft.com' and converted password "Passw0rd" to store for connecting to SPO.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(ParameterSetName="CredentialObj", Mandatory=$true, Position=1)]
        [System.Management.Automation.PSCredential]
        $Credentials,

        [Parameter(ParameterSetName="RawCreds", Mandatory=$true, Position=1)]
        [string]
        $Username,

        [Parameter(ParameterSetName="RawCreds", Mandatory=$true, Position=2)]
        [SecureString]
        $Password
    )

    if (@(Get-Variable -Scope Script -Name SPOnline_Credentials -ErrorAction SilentlyContinue).Count -gt 0) {
        Write-Verbose "Previously set credentials detected, will be overwritten."
    }

    switch ($PSCmdlet.ParameterSetName) {
        "CredentialObj" { 
            $Username = $Credentials.UserName
            $Password = $Credentials.Password 
        }
    }

    $Script:SPOnline_Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username,$Password)
}

Function ConvertTo-SPOCredentials {
    <#
        .SYNOPSIS
            ConvertTo-SPOCredentials converts a standard PSCredential object into a SharePointOnline Credentials Object.
        .DESCRIPTION
        .PARAMETER Credentials
            The Credentials parameter allows the passage of Username and Passwords through a PSCredential object, as gathered
            from another commandlet like Get-Credentials
        .PARAMETER Username
            The Username parameter accepts the username of the target account to be used with the SPO Module
        .PARAMETER Password
            The Password parameter is a secure string version of the password to be used with the 
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(ParameterSetName="CredentialObj", Mandatory=$true, Position=1)]
        [System.Management.Automation.PSCredential]
        $Credentials,

        [Parameter(ParameterSetName="RawCreds", Mandatory=$true, Position=1)]
        [string]
        $Username,

        [Parameter(ParameterSetName="RawCreds", Mandatory=$true, Position=2)]
        [SecureString]
        $Password
    )

    switch ($PSCmdlet.ParameterSetName) {
        "CredentialObj" { 
            $Username = $Credentials.UserName
            $Password = $Credentials.Password 
        }
    }

    Write-Output (New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username,$Password))
}

# SharePoint CSOM Commandlets
#
Function Get-SPOSiteTemplate
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .PARAMETER Language
        .EXAMPLE
            Get-SPOWebTemplate -URL 'https://mytenant.sharepoint.com/teams/SiteCol'

            This example returns all available web templates (global and site collection-specific) within the given site collection.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$false, Position=2)]
        [string]
        $Name = "*",

        [Parameter(Mandatory=$false)]
        [string]
        $Language = "1033",

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials
    )

    BEGIN {
        if ($Credential -eq $null) {
            throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
        }

        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
        $Context.Credentials = $Credential
    }

    PROCESS {
        $Templates = $Context.Site.GetWebTemplates($Language,"0")
        $Context.Load($Templates)
        $Context.ExecuteQuery()

        Write-Output $Templates
    }

    END {
    }
}

Function New-SPOWeb
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .PARAMETER Title
        .PARAMETER Description
        .PARAMETER Template
        .PARAMETER Language
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string]
        $Name,

        [Parameter(Mandatory=$false)]
        [string]
        $Title = $Name,

        [Parameter(Mandatory=$false)]
        [string]
        $Description = "",

        [Parameter(Mandatory=$false)]
        [string]
        $Template = "STS#0",

        [Parameter(Mandatory=$false)]
        [string]
        $Language = "1033",

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials
    )

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $WCI = New-Object Microsoft.SharePoint.Client.WebCreationInformation

    $WCI.WebTemplate = $Template
    $WCI.Description = $Description
    $WCI.Title = $Title
    $WCI.Url = $Name
    $WCI.Language = $Language
    $SubWeb = $Context.Web.Webs.Add($WCI)

    try {
        $Context.ExecuteQuery()
        Write-Verbose "Site created successfully at $Url/$Name."
    } catch {
        throw "Failed to create subsite $Url/$($Name): $($_.Exception.Message)"
    }
}

Function Get-SPOWeb
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER IncludeParent
        .PARAMETER Recurse
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$false)]
        [switch]
        $IncludeParent,

        [Parameter(Mandatory=$false)]
        [switch]
        $Recurse,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials
    )

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $Context.Load($Context.Web)
    $Context.Load($Context.Web.Webs)
    $Context.ExecuteQuery()
    $OutWebs = 0
    $InWebs = @()

    if ($IncludeParent -eq $true) {
        $InWebs += $Context.Web
    }

    $Context.Web.Webs | foreach { $InWebs += $_ }

    foreach ($Web in $InWebs) {
        Write-Debug "Processing subsite $($Web.Url)."

        $Context.Load($Web.AssociatedOwnerGroup)
        $Context.Load($Web.AssociatedMemberGroup)
        $Context.Load($Web.AssociatedVisitorGroup)
        try { $Context.ExecuteQuery() }
        catch {
            throw "Error loading associated groups for site '$($Url)': $($_.Exception.Message)"
        }

        $WebObj = $Web | Select Title,Id,Url,ServerRelativeUrl,Description,WebTemplate,SiteLogoUrl,SiteLogoDescription,RecycleBinEnabled,QuickLaunchEnabledCreated
        Add-Member -InputObject $WebObj -MemberType NoteProperty -Name AssociatedOwnerGroup -Value $Web.AssociatedOwnerGroup.Title
        Add-Member -InputObject $WebObj -MemberType NoteProperty -Name AssociatedMemberGroup -Value $Web.AssociatedMemberGroup.Title
        Add-Member -InputObject $WebObj -MemberType NoteProperty -Name AssociatedVisitorGroup -Value $Web.AssociatedVisitorGroup.Title
        Write-Output $WebObj

        if ($Recurse) {
            Get-SPOWeb -Url $Web.Url -Recurse
        }
    }
}

Function Set-SPOWeb
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER AssociatedOwnerGroup
        .PARAMETER AssociatedMemberGroup
        .PARAMETER AssociatedVisitorGroup
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$false)]
        [string]
        $AssociatedOwnerGroup,

        [Parameter(Mandatory=$false)]
        [string]
        $AssociatedMemberGroup,

        [Parameter(Mandatory=$false)]
        [string]
        $AssociatedVisitorGroup,

        [Parameter(Mandatory=$false)]
        [Hashtable[]]
        $AddGlobalNav,

        [Parameter(Mandatory=$false)]
        [Hashtable[]]
        $AddQuickLaunch,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $Context.Load($Context.Web)
    $Context.Load($Context.Web.SiteGroups)
    $Context.ExecuteQuery()

    $GroupList = @()
    $GroupList += New-Object PSObject -Property @{
        Name = 'AssociatedOwnerGroup'
        Value = $AssociatedOwnerGroup
    }
    $GroupList += New-Object PSObject -Property @{
        Name = 'AssociatedMemberGroup'
        Value = $AssociatedMemberGroup
    }
    $GroupList += New-Object PSObject -Property @{
        Name = 'AssociatedVisitorGroup'
        Value = $AssociatedVisitorGroup
    }

    foreach ($AssocGroup in $GroupList) {
        if (-not [string]::IsNullOrEmpty($AssocGroup.Value)) {
            Write-Verbose "Set request for $($AssocGroup.Name) for the site $Url" 
            $Context.Load($Context.Web.$($AssocGroup.Name))
            $Context.ExecuteQuery()

            $Group = $Context.Web.SiteGroups.GetByName($AssocGroup.Value)
            $Context.Load($Group)
            try { 
                $Context.ExecuteQuery() 
            } catch {
                Write-Error "Could not locate requested group $($AssocGroup.Value) in the collection $($Url): $($_.Exception.Message)"
            }

            $Context.Web.$($AssocGroup.Name) = $Group
            $Context.Web.Update()

            try { 
                $Context.ExecuteQuery()
                Write-Verbose "Set request completed for property $($AssocGroup.Name) on site $Url."                 
            } catch {
                Write-Error "Could not set requested $($AssocGroup.Name) to $($AssocGroup.Value) in the collection $($Url): $($_.Exception.Message)"
            }
        }
    }

    if ($PSBoundParameters.ContainsKey("AddGlobalNav")) {
        Helper-AddNavigation -Url $Url -NavDataCol $AddGlobalNav -NavType Global
    }

    if ($PSBoundParameters.ContainsKey("AddQuickLaunch")) {
        Helper-AddNavigation -Url $Url -NavDataCol $AddQuickLaunch -NavType Quick
    }
}

Function Remove-SPOWeb
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER AssociatedGroups
        .EXAMPLE
    #>
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$false)]
        [switch]
        $AssociatedGroups,

        [Parameter(Mandatory=$false)]
        [switch]
        $Force,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $Context.Load($Context.Web)
    try { $Context.ExecuteQuery() }
    catch {
        throw "Error loading site '$($Url)': $($_.Exception.Message)"
    }

    if ($AssociatedGroups -eq $True)
    {
        $Context.Load($Context.Web.AssociatedOwnerGroup)
        $Context.Load($Context.Web.AssociatedMemberGroup)
        $Context.Load($Context.Web.AssociatedVisitorGroup)

        try { $Context.ExecuteQuery() }
        catch {
            throw "Error loading associated groups for site '$($Url)': $($_.Exception.Message)"
        }
        
        foreach ($AssocGroup in @($Context.Web.AssociatedOwnerGroup, $Context.Web.AssociatedMemberGroup, $Context.Web.AssociatedVisitorGroup)) {
            if ($AssocGroup -ne $null -and -not [string]::IsNullOrEmpty($AssocGroup.Title)) {
                Remove-SPOGroup -Url $Url -Name $AssocGroup.Title
            }
        }
    }

    if ($PSCmdlet.ShouldProcess($Url, "Delete SharePoint Online site"))
    {
        if ($Force -eq $true -or $PSCmdlet.ShouldContinue("Are you sure you want to delete the SharePoint Online site at $($Url)?", "Confirm SPO Site deletion")) {
            $Context.Web.DeleteObject()
            try { 
                $Context.ExecuteQuery() 
                Write-Verbose "Successfully deleted site '$Url'."
            } catch {
                Write-Error "Error deleting site '$($Url)': $($_.Exception.Message)"
            }
        }
    }
}

Function Get-SPOUser
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .PARAMETER Description
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2, ValueFromPipeline=$True)]
        [AllowNull()][AllowEmptyString()]
        [string[]]
        $User,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    Begin {
        if ($Credential -eq $null) {
            throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
        }

        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
        $Context.Credentials = $Credential

        $Context.Load($Context.Site.RootWeb.SiteUsers)
        $Context.ExecuteQuery()
        $SiteUsers = $Context.Site.RootWeb.SiteUsers
    }

    Process {
        foreach ($UserItem in $User) {
            if ([string]::IsNullOrEmpty($UserItem)) {
                continue
            }

            $LookupId = [int] 0
            $OutputItem = $null

            if ([int]::TryParse($UserItem, [ref] $LookupId)) {
                $UserById = $SiteUsers | Where { $_.Id -eq $LookupId }

                if ($UserById -eq $null) {
                    Write-Error "User specified by ID '$UserItem' but not found on site $Url."
                } else {
                    $OutputItem = $UserById
                }
            } else {
                Write-Debug "Performing lookup of user logon name '$($UserItem)'"
                [array] $Precheck = $SiteUsers | Where { $_.LoginName -like "*|$UserItem" }

                if ($Precheck.Count -eq 1) {
                    Write-Debug "Found user in SiteUsers, ID = $($Precheck.Id)"
                    $OutputItem = $Precheck
                } else {
                    $LookupUser = $Context.Web.EnsureUser($UserItem)
                    $Context.Load($LookupUser)
                    try { 
                        $Context.ExecuteQuery()
                        Write-Debug "Found User on EnsureUser, ID = $($LookupUser.Id)'"
                        $OutputItem = $LookupUser
                    } catch {
                        Write-Error "Could not lookup user '$($UserItem)' for site '$($Url)': $($_.Exception.Message)"
                    }
                }
            }

            if ($OutputItem -ne $null) {
                Write-Output $OutputItem | Select Id,Title,LoginName,@{N='LookupValue'; E={$UserItem}}
            }
        }
    }

    End {
    }
}

Function New-SPOGroup
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .PARAMETER Description
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string]
        $Name,

        [Parameter(Mandatory=$false, Position=3)]
        [string]
        $Description,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $SiteName = $Url.Substring($Url.LastIndexOf("/")+1, $Url.Length-$Url.LastIndexOf("/")-1)

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $Context.Load($Context.Web)
    $Context.Load($Context.Web.SiteGroups)
    $Context.ExecuteQuery()

    
    Write-Verbose "Creating group named $Name with $Permission Perms for $Url."
    $GroupDef = New-Object Microsoft.SharePoint.Client.GroupCreationInformation
    $GroupDef.Description = $Description
    $GroupDef.Title = $Name
    $Group = $Context.Web.SiteGroups.Add($GroupDef)
    # $Context.Load($Group)

    try {
        $Context.ExecuteQuery()
        Write-Verbose "Group $Name created successfully at site $Url."
    } catch {
        Write-Error "Failed to create group $($Name): $($_.Exception.Message)"
    }
}

Function Get-SPOGroup
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$false)]
        [string]
        $Name = "*",

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $Context.Load($Context.Web)
    $Context.Load($Context.Web.SiteGroups)
    $Context.ExecuteQuery()

    return $Context.Web.SiteGroups | Select Title,Id,LoginName,Description,OwnerTitle | where { $_.Title -like $Name }
}

Function Test-SPOGroup
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$false)]
        [string]
        $Name = "*",

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $Context.Load($Context.Web)
    $Context.Load($Context.Web.SiteGroups)
    $Context.ExecuteQuery()

    return ($Context.Web.SiteGroups.Title -contains $Name)
}

Function Set-SPOGroup
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .PARAMETER Owner
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string]
        $Name,

        [Parameter(Mandatory=$false)]
        [string]
        $Owner,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    Write-Verbose "Modifying settings for $Name at $Url."

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $Context.Load($Context.Web)
    $Context.Load($Context.Web.SiteGroups)
    $Context.ExecuteQuery()

    $Group = $Context.Web.SiteGroups.GetByName($Name)
    $Context.Load($Group)
    try {
        $Context.ExecuteQuery()
    } catch {
        throw "Could not locate Group $Name at $Url to modify."
    }

    if (-not [string]::IsNullOrEmpty($Owner))
    {
        Write-Verbose "Ownership change request processing for $Name."

        #Test if this is a site group
        $OwnerObj = $Context.Web.SiteGroups.GetByName($Owner)
        $Context.Load($OwnerObj)
        try {
            $Context.ExecuteQuery() 
        } catch {
            # Non-critical failure. Just didn't find the requested item as a group. 
            $OwnerObj = $null 
        }

        # Retry with User as owner
        if ($OwnerObj -eq $null)
        {
            $OwnerObj = $Context.Web.EnsureUser($Owner)
            $Context.Load($OwnerObj)
            try { 
                $Context.ExecuteQuery() 
            } catch {
                throw "Owner not found as group or user. Please check ownership and try again."
            }
        }
            
        $Group.Owner = $OwnerObj
        $Group.Update()

        try {
            $Context.ExecuteQuery()
            Write-Verbose "Group '$Name' ownership changed successfully to '$Owner'."
        } catch {
            Write-Error "Error setting ownership for group $($Name): $($_.Exception.Message)"
        }
    }
}

Function Remove-SPOGroup
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .EXAMPLE
    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string]
        $Name,

        [Parameter(Mandatory=$false)]
        [switch]
        $Force,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    Write-Verbose "Modifying settings for $Name at $Url."

    if ($Credential -eq $null) {
        $PSCmdlet.ThrowTerminatingError("No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets.")
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential
    Write-Debug "Credentials acquired and context created for $Url."

    $Context.Load($Context.Web)
    $Context.Load($Context.Web.SiteGroups)
    $Context.ExecuteQuery()
    Write-Debug "Completed initial load of site and site groups."

    $Group = $Context.Web.SiteGroups.GetByName($Name)
    $Context.Load($Group)

    try {
        $Context.ExecuteQuery()
        Write-Verbose "Successfully located target group $Name in $Url."
    } catch {
        throw "Could not locate Group $Name at $Url to remove: $($_.Exception.Message)"
    }

    if ($pscmdlet.ShouldProcess($Name, "Delete SharePoint Online Group at $Url")) {
        if ($Force -eq $true -or $pscmdlet.ShouldContinue("Are you sure you want to delete the SharePoint Online Group $Name at $($Url)?", "Confirm SPO Group deletion")) {
            try { 
                $Context.Web.SiteGroups.Remove($Group)
                $Context.ExecuteQuery() 
                Write-Verbose "Successfully deleted site group '$Name' at '$Url'."
            } catch {
                Write-Error "Error deleting site group '$Name' at '$($Url)': $($_.Exception.Message)"
            }
        }
    }
}

Function Add-SPOUserToGroup
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER User
        .PARAMETER Group
        .EXAMPLE
    #>
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Low")]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string[]]
        $User,

        [Parameter(Mandatory=$true, Position=3)]
        [string]
        $Group,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    Write-Verbose "Adding user $User to group $Group for $Url."

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $UserList = @() + $User
    if ($UserList.Count -eq 0) {
        Write-Warning "Cannot add users to group $Name at $($Url): No users specified"
        return
    }
    Write-Debug "Given $($UserList.Count) users to add to $Group at $Url. User List: $($UserList -join ', ')"

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $Context.Load($Context.Web)
    $Context.Load($Context.Web.SiteGroups)
    $Context.ExecuteQuery()

    $GroupObj = $Context.Web.SiteGroups.GetByName($Group)
    $Context.Load($GroupObj)

    foreach ($U in $UserList) {
        Write-Debug "Attempting to validate user $U to add to group $Group at $Url."
        $UserObj = $Context.Web.EnsureUser($U)
        $Context.Load($UserObj)

        $UserToAdd = $GroupObj.Users.AddUser($UserObj)
        $Context.Load($UserToAdd)
    }

    try {
        Write-Debug "Committing user changes to group $Group at $Url."
        $Context.ExecuteQuery()
        Write-Verbose "Successfully added $($UserList.Count) user$(if ($UserList -gt 1) { "s" }) to SharePoint Group '$Group' at $Url."
    } catch {
        Write-Error "Failed to add users to Group $($Group): $($ex.Exception.Message)"
    }
}

Function Remove-SPOUserFromGroup
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER User
        .PARAMETER Group
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string[]]
        $User,

        [Parameter(Mandatory=$true, Position=3)]
        [string]
        $Group,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    Write-Verbose "Adding user $User to group $Group for $Url."

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $UserList = @() + $User
    if ($UserList.Count -eq 0) {
        Write-Warning "Cannot remove users from group $Name at $($Url): No users specified"
        return
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $Context.Load($Context.Web)
#    $Context.Load($Context.Web.SiteGroups)
    $GroupObj = $Context.Web.SiteGroups.GetByName($Group)
    $Context.Load($GroupObj)
    $Context.ExecuteQuery()

    foreach ($U in $UserList) {
        Write-Debug "Attempting to validate user $U to remove from group $Group at $Url."
        $UserObj = $Context.Web.EnsureUser($U)
        $Context.Load($UserObj)

        $GroupObj.Users.Remove($UserObj)
    }

    $GroupObj.Update()

    try {
        $Context.ExecuteQuery()
        Write-Verbose "Successfully removed $($UserList.Count) user$(if ($UserList -gt 1) { "s" }) from SharePoint Group '$Group' at $Url."
    } catch {
        Write-Error "Failed to remove user(s) from Group $($Group): $($ex.Exception.Message)"
    }
}

Function Get-SPOGroupMembers
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string]
        $Name,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $Context.Load($Context.Web)
    $Context.Load($Context.Web.SiteGroups)
    $Context.ExecuteQuery()

    $Group = $Context.Web.SiteGroups.GetByName($Name)
    if ($Group -eq $null) { 
        Write-Warning "A group by the name of $Name was not found at $Url"
        return
    } 

    $Context.Load($Group.Users)
    try { $Context.ExecuteQuery() }
    catch {
        throw "Error accessing membership of $Name at $($Url): $($_.Exception.Message)"
    }

    Write-Output ($Group.users | Select Id,Title,Email,LoginName,PrincipalType,IsSiteAdmin)
}

Function Set-SPOSitePermission
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .PARAMETER Permission
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string]
        $Name,

        [ValidateSet('Full Control', 'Edit', 'Read')]
        [string]
        $Permission,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    Write-Verbose "Modifying site permissions $Url, adding $Name with $Permission rights."

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $Context.Load($Context.Web)
    $Context.Load($Context.Web.SiteGroups)
    $Context.Load($Context.Web.RoleDefinitions)
    $Context.Load($Context.Web.RoleAssignments)
    $Context.ExecuteQuery()

    $Principal = $Context.Web.SiteGroups.GetByName($Name)
    $Context.Load($Principal)
    try { $Context.ExecuteQuery() }
    catch { $Principal = $null }

    # Retry with User as owner
    if ($Principal -eq $null)
    {
        $Principal = $Context.Web.EnsureUser($Owner)
        $Context.Load($Principal)
        try { $Context.ExecuteQuery() }
        catch
        {
            Write-Error "Principal not found as group or user. Please check the name and try again."
            return
        }
    }

    $RoleDef = $Context.Web.RoleDefinitions.GetByName($Permission)
    $Binding = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($Context)
    $Binding.Add($RoleDef)
    $Assignment = $Context.Web.RoleAssignments.Add($Principal, $Binding)

    try {
        $Context.ExecuteQuery()
        Write-Verbose "Successfully granted '$Permission' rights to '$Name'"
    } catch {
        Write-Error "Failed to create group $($Name): $($_.Exception.Message)"
    }
}

Function New-SPOList
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string]
        $Name,

        [Parameter(Mandatory=$false)]
        [string]
        $Description = "",

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials
    )

    Begin {
        if ($Credential -eq $null) {
            throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
        }

        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
        $Context.Credentials = $Credential

        $Context.Load($Context.Web)
        $Context.Load($Context.Web.Lists)
        $Context.ExecuteQuery()
    }

    Process {
        $NewListInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
        $NewListInfo.Title = $Name
        $NewListInfo.Description = $Description
        $NewListInfo.TemplateType = [Microsoft.SharePoint.Client.ListTemplateType]::GenericList

        $NewList = $Context.Web.Lists.Add($NewListInfo)
        $Context.Load($NewList)

        $Context.ExecuteQuery()
    }

    End {
    }
}

Function Get-SPOList
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$false, Position=2)]
        [string]
        $Name = "*",

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $Context.Load($Context.Web)
    $Context.Load($Context.Web.Lists)
    $Context.ExecuteQuery()

    $SiteLists = $Context.Web.Lists | 
                Select Title,Description,Created,Id,ItemCount,Hidden,@{N="Type";E={$_.BaseType.ToString()}},@{N="UnderlyingObject";E={$_}} |
                where { $_.Title -like $Name }
  
    foreach ($List in $SiteLists) {
        $List.PSObject.TypeNames.Insert(0,'Kittyfox.SharePoint.Online.List') 
    }
    
    return $SiteLists 
}

Function Get-SPOListFields
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .PARAMETER Hidden
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(ParameterSetName="ByContext", Mandatory=$true, Position=1)]
        [Microsoft.SharePoint.Client.ClientContext]
        $Context,

        [Parameter(ParameterSetName="ByList", Mandatory=$true, Position=1)]
        [Microsoft.SharePoint.Client.List]
        $List,

        [Parameter(ParameterSetName="ByUrl", Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string]
        $Name,

        [Parameter(Mandatory=$false)]
        [switch]
        $Hidden,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials
    )

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    if ($PSCmdlet.ParameterSetName -eq 'ByUrl') {
        Write-Debug "Provided URL $Url Only, generating our own context."

        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
        $Context.Credentials = $Credential

        $Context.Load($Context.Web)
        $Context.Load($Context.Web.Lists)
        #$Context.ExecuteQuery()
    }

    $FieldsNeedsInit = $true
    if ($PSCmdlet.ParameterSetName -ne 'ByList') {
        Write-Debug "Passed in Context or Url, fetching List."
        $List = $Context.Web.Lists.GetByTitle($Name)
    } else {
        Write-Debug "Passed parameter by list, checking if Fields needs initialization."
        try { 
            $List.Fields.Count | Out-Null 
            $FieldsNeedsInit = $false
            Write-Debug "Fields is READY FOR USE."
        } catch { 
            Write-Debug "Fields needs initializing."
        }
    }
    
    if ($FieldsNeedsInit -eq $true) {
        $Context.Load($List)
        $Context.Load($List.Fields)
        try { $Context.ExecuteQuery() } 
        catch { 
            Write-Error "Error fetching $List from $($Url): $($_.Exception.Message)"
            return
        }
    }

    $OutArray = $List.Fields | Select Title,InternalName,Id,TypeAsString,Scope,Hidden,CanBeDeleted,ReadOnlyField,Required,@{N='UnderlyingObject'; E={$_}}

    if ($Hidden -eq $false) {
        $OutArray = $OutArray | where { $_.Hidden -eq $false }
    }

    $OutArray | foreach { $_.PSObject.TypeNames.Insert(0,'Kittyfox.SharePoint.Online.ListField') }

    Write-Output $OutArray
}

Function Get-SPOListItems
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .PARAMETER Name
        .PARAMETER Filter
        .PARAMETER Recurse
        .EXAMPLE
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string]
        $Name,

        [Parameter(Mandatory=$false)]
        [hashtable]
        $Filter,

        [Parameter(Mandatory=$false)]
        [string]
        $Folder,
        
        [Parameter(Mandatory=$false)]
        [ValidateScript({[int]::TryParse($_, [ref] $null) -or $_ -eq 'All'})]
        [string]
        $Limit,

        [Parameter(Mandatory=$false)]
        [switch]
        $Recurse,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials
    )

    $ShowLimit = 0

    if ($PSBoundParameters.ContainsKey('Filter')) {
        $ViewLimit = 20000
    } else {
        $ViewLimit = 5000
    }

    if ($Credential -eq $null) {
        throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
    }

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $Context.Credentials = $Credential

    $Context.Load($Context.Web)
    $Context.Load($Context.Web.Lists)
    #$Context.ExecuteQuery()

    $List = $Context.Web.Lists.GetByTitle($Name)
    $Context.Load($List)
    $Context.Load($List.Fields)
    try { $Context.ExecuteQuery() } 
    catch { 
        throw "Error fetching $($List.Name) from $($Url): $($_.Exception.Message)"
    }

    $StdFields = 'ID', 'Modified', 'Created', 'Author', 'Editor'
    $SelectedFields = Get-SPOListFields -List $List -Name $Name | where { $_.ReadOnlyField -eq $false -or $_.InternalName -in $StdFields } | Select InternalName,TypeAsString

    $CamlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery

    if ($Recurse -eq $true) {
        $Scope = " Scope=`"Recursive`""
    }

    # Check to see if we have to clamp to maximum view limit. If we do clamp, and the user didn't already specify 'All', warn them.
    if ($Limit -eq 'All' -or $Filter -ne $null) {
        $ShowLimit = $List.ItemCount
        $ShowAll = $true
        Write-Debug "Caller requested Limit = 'All' or specified Filter, configuring for Show All behavior."
    } elseif ([string]::IsNullOrEmpty($Limit)) {
        $ShowLimit = $List.ItemCount
    } else {
        $ShowLimit = [Convert]::ToInt32($Limit)
    }

    Write-Debug "Comparison check: ShowLimit - $($ShowLimit.GetType().Name) - $ShowLimit; ViewLimit - $($ViewLimit.GetType().Name) - $ViewLimit"

    if ($ShowLimit -gt $ViewLimit) {
        if ($ShowAll -ne $true) {
            Write-Warning "List $Name has item count ($($List.ItemCount)) greater than maximum allowable view of $ViewLimit. Use -Limit 'All' to retrieve all items or use -Filter to limit the results returned."
        }
        $RowLimit = "<RowLimit>$ViewLimit</RowLimit>"
        # $RowLimit2 = " RowLimit=$ViewLimit"
        Write-Debug "Clamped display limit to maximum view limit of $ViewLimit"
    } elseif (-not [string]::IsNullOrEmpty($ShowLimit)) {
        $RowLimit = "<RowLimit>$ShowLimit</RowLimit>"
        # $RowLimit2 = " RowLimit=$ViewLimit"
        Write-Debug "Clamped display limit to requested limit of $ShowLimit"
    }

    $QueryXml = "<View$Scope>{0}$RowLimit</View>"

    if ($Filter -ne $null) {
        $QueryXml = $QueryXml -f "<Query><Where>{0}</Where></Query>"
        if ($Filter.Count -gt 1) {
            $QueryXml = $QueryXml -f "<And>{0}</And>"
        }

        Write-Debug "Processing $($Filter.Count) filter terms."

        $QueryXmlTerms = "{0}"

        foreach ($Term in @($Filter.Keys)) {
            $FieldDefType = $($SelectedFields | where { $_.InternalName -eq $Term } | Select -Expand TypeAsString)
            $outInt = [int] 0
            $outDate = [datetime]::Now

            $FieldType = "Text"

            if ([int]::TryParse($Filter[$Term], [ref] $outInt)) { 
                if ($FieldDefType -eq 'Lookup') {
                    $FieldType = "Lookup"
                    $FieldRefLookup = "LookupId=`"TRUE`""
                } Else {
                    $FieldType = "Integer"
                    if ($FieldDefType -eq 'User') {
                        $FieldRefLookup = "LookupId=`"TRUE`""
                    }
                }
            }
            if ([datetime]::TryParse($Filter[$Term], [ref] $outDate)) { $FieldType = "Date" }

            if ($FieldType -eq 'Text' -and $Filter[$Term].Contains('*')) {
                $QueryXmlTerms = $QueryXmlTerms -f "<And>{0}</And>"
                $FilterValue = $Filter[$Term]
                $FilterParts = $FilterValue.Split('*')
                $Op = "<BeginsWith>{0}</BeginsWith>{1}"

                if ($FilterValue.Substring(0, 1) -eq '*') {
                    $Op = "<Contains>{0}</Contains>{1}"
                }

                foreach ($Part in $FilterParts) {
                    if ($Part.Length -gt 0) {
                        $Conditional = $Op -f "<FieldRef Name=`"$Term`" $FieldRefLookup/><Value Type=`"$FieldType`"><![CDATA[$($Part)]]></Value>", "{0}"
                        $QueryXmlTerms = $QueryXmlTerms -f $Conditional
                        $Op = "<Contains>{0}</Contains>{1}"
                    }
                }
            } else {
                $QueryXmlTerms = $QueryXmlTerms -f "<Eq><FieldRef Name=`"$Term`" $FieldRefLookup/><Value Type=`"$FieldType`"><![CDATA[$($Filter[$Term])]]></Value></Eq>{0}"
            }
        }
        $QueryXml = $QueryXml -f ($QueryXmlTerms -f [string]::Empty)
    } else {
        $QueryXml = $QueryXml -f [string]::Empty
    }

    Write-Debug "QueryXml String: $QueryXml"
    $CamlQuery.ViewXml = $QueryXml

    $GatheredCount = 0

    do {
        if ($ListPosition -ne $null) {
            Write-Debug "Iterating through additional pages of data."
            $CamlQuery.ListItemCollectionPosition = $ListPosition
        }

        $Items = $List.GetItems($CamlQuery)
        $Context.Load($Items)
        try { $Context.ExecuteQuery() } 
        catch { 
            throw "Error fetching items list $($List.Name) a $($Url): $($_.Exception.Message)"
        }

        foreach ($Item in $Items) {
            $ItemObj = New-Object -TypeName PSObject 

            foreach ($Field in $SelectedFields) {
                switch ($Field.TypeAsString) {
                    "URL" {
                        $UrlObj = New-Object -Type PSObject -Property @{Description=$Item[$Field.InternalName].Description; Url=$Item[$Field.InternalName].Url}
                        Add-Member -InputObject $ItemObj -NotePropertyName $Field.InternalName -NotePropertyValue $UrlObj
                    }
                    "Lookup" {
                        Add-Member -InputObject $ItemObj -Name $Field.InternalName -MemberType NoteProperty -Value $Item[$Field.InternalName].Email
                    }
                    "User" {
                        $UserObj = New-Object -Type PSObject -Property @{ID=$Item[$Field.InternalName].LookupId; DisplayName=$Item[$Field.InternalName].LookupValue; Email=$Item[$Field.InternalName].Email}
                        Add-Member -InputObject $ItemObj -Name $Field.InternalName -MemberType NoteProperty -Value $UserObj
                    }
                    "Calculated" {
                        $CalcField = $Item[$Field.InternalName] -as [Microsoft.SharePoint.SPFieldCalculated]
                        $CalcValue = $CalcField.GetFieldValueAsText($Item[$Field.InternalName])
                        Add-Member -InputObject $ItemObj -Name $Field.InternalName -MemberType NoteProperty -Value $CalcValue
                    }
                    default {
                        Add-Member -InputObject $ItemObj -NotePropertyName $Field.InternalName -NotePropertyValue $Item[$Field.InternalName]
                    }
                }
            }

            $ItemObj.PSObject.TypeNames.Insert(0,'Kittyfox.SharePoint.Online.ListItem')

            Write-Output $ItemObj
        }
        $GatheredCount += $Items.Count
        $ListPosition = $Items.ListItemCollectionPosition
    } while ($GatheredCount -lt $ShowLimit -and $ShowAll -eq $True -and $ListPosition -ne $null)             
}

Function New-SPOListItem
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .EXAMPLE
    #>
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Low")]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string]
        $Name,

        [Parameter(Mandatory=$true, Position=3)]
        [hashtable[]]
        $Property,

        [Parameter(Mandatory=$false)]
        [switch]
        $PassThru,

        [Parameter(Mandatory=$false)]
        [switch]
        $Batch,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials
    )

    Begin {
        if ($Credential -eq $null) {
            throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
        }

        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
        $Context.Credentials = $Credential

        $Context.Load($Context.Web)
        $Context.Load($Context.Web.Lists)
        # $Context.ExecuteQuery()

        $List = $Context.Web.Lists.GetByTitle($Name)
        $Context.Load($List)
        $Context.Load($List.Fields)
        try { $Context.ExecuteQuery() } 
        catch { 
            throw "Error fetching $List from $($Url): $($_.Exception.Message)"
        }

        $EditableFields = Get-SPOListFields -List $List -Name $Name | where { $_.ReadOnlyField -eq $false } | Select InternalName,Required,TypeAsString
        $RequiredFields = $EditableFields | where { $_.Required -eq $true } | Select -Expand InternalName
        $RowList = @()
    }

    Process {
        foreach ($PropertySet in $Property) {
            $AllRequiredSupplied = $true
            foreach ($ReqTerm in $RequiredFields) {
                if ($ReqTerm -notin $PropertySet.Keys) {
                    $AllRequiredSupplied = $false
                }
            }

            if ($AllRequiredSupplied -eq $false) {
                throw "All required fields not supplied on input. Required fields for $Name are: $($RequiredFields -join ', ')"
            }

            # Ruddy Hacks. 
            # 
            # Process the data and, if we need to do lookups, we MUST do this before we start creating
            # the list object.
            $FormattedData = New-Object System.Collections.Hashtable
            foreach ($Term in $PropertySet.Keys) {
                $IsRequired = $RequiredFields -contains $Term

                if ($Term -notin $EditableFields.InternalName) {
                    throw "Specified field not present in $Name as an editable field: $Term"
                }

                $FieldType = $EditableFields | where { $_.InternalName -eq $Term } | Select -Expand TypeAsString
        
                try {
                    $Value = Helper-PrepareListItemData -Url $Url -Term $Term -FieldType $FieldType -Value $PropertySet[$Term] -ErrorAction Stop
                } catch {
                    Write-Error $_.Exception.Message
                }

                if ([string]::IsNullOrEmpty($Value) -and $IsRequired) {
                    throw "Required field $Term has no value after processing, aborting."
                }

                $FormattedData.Add($Term, $Value)
            }

            # Now that we have formatted all the data, performed any column or user lookups we had to
            # NOW we can create the List Item Object and force the object's creation
            #
            # To do this in a single pass means we only get partial data (as the pending object will
            # be only partially created in the list), or we get errors on the lookup data to be
            # inserted. So stupid.
            $ListItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
            $ListItem = $List.AddItem($ListItemInfo)

            foreach ($Term in $Property.Keys) {
                $ListItem[$Term] = $FormattedData[$Term]
            }

            $ListItem.Update()
            $RowList += $ListItem
        }
    }

    End {
        $Context.Load($List)

        try {
            if ($PSCmdlet.ShouldProcess($Url, "New List Item(s) for list $Name")) {
                $Context.ExecuteQuery()
                Write-Verbose "$($RowList.Count) new list items created on list $Name."
            
                if ($PassThru) { 
                    foreach ($ListItem in $RowList) { 
                        Write-Output $ListItem.Id 
                    } 
                }
            }
        } catch {
            Write-Error "Failed to add new item to list $Name at $($Url): $($_.Exception.Message)"
        }
    }
}

Function Update-SPOListItem
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .EXAMPLE
    #>
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Medium")]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,

        [Parameter(Mandatory=$true, Position=2)]
        [string]
        $Name,

        [Parameter(Mandatory=$true, Position=3, ValueFromPipelineByPropertyName)]
        [int[]]
        $ItemId,

        [Parameter(Mandatory=$true, Position=4, ValueFromPipelineByPropertyName)]
        [hashtable[]]
        $Property,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials
    )

    Begin {
        if ($Credential -eq $null) {
            throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
        }

        if ($ItemId.Count -ne $Property.Count) {
            throw "Parameter Mismatch. Must have the same number of ItemIds as PropertySets to apply."
        }

        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
        $Context.Credentials = $Credential

        $Context.Load($Context.Web)
        $Context.Load($Context.Web.Lists)
        $Context.ExecuteQuery()

        $List = $Context.Web.Lists.GetByTitle($Name)
        $Context.Load($List)
        $Context.Load($List.Fields)

        try { $Context.ExecuteQuery() } 
        catch { 
            throw "Error fetching $List from $($Url): $($_.Exception.Message)"
        }

        $EditableFields = Get-SPOListFields -List $List -Name $Name | where { $_.ReadOnlyField -eq $false } | Select InternalName,Required,TypeAsString
        $Index = 0
    }

    Process {
        while ($Index -lt $ItemId.Count) {
            $CurrentIndex = $Index++

            # Ruddy Hacks. 
            # 
            # Process the data and, if we need to do lookups, we MUST do this before we start creating
            # the list object.
            $FormattedData = New-Object System.Collections.Hashtable
            foreach ($Term in $Property[$CurrentIndex].Keys) {
                if ($Term -notin $EditableFields.InternalName) {
                    Write-Error "Specified field not present in $Name as an editable field: $Term"
                    continue
                }

                $FieldType = $EditableFields | where { $_.InternalName -eq $Term } | Select -Expand TypeAsString

                $Value = Helper-PrepareListItemData -Url $Url -Term $Term -FieldType $FieldType -Value $Property[$CurrentIndex][$Term]
                $FormattedData.Add($Term, $Value)
            }

            # Now that we have formatted all the data, performed any column or user lookups we had to
            # NOW we can reacquire the List Item Object and make the necessary changes.
            #
            # To do this in a single pass means we only get partial data (as the pending object will
            # be only partially created in the list), or we get errors on the lookup data to be
            # inserted. So stupid.
            $ListItem = $List.GetItemById($ItemId[$CurrentIndex])
            $Context.Load($ListItem)

            #try { $Context.ExecuteQuery() } 
            #catch { 
            #    Write-Error "Error fetching item ID '$($ItemId[$CurrentIndex])' from $List at $($Url): $($_.Exception.Message)"
            #    continue
            #}

            foreach ($Term in $Property[$CurrentIndex].Keys) {
                $ListItem[$Term] = $FormattedData[$Term]
            }

            $ListItem.Update()
        }
    }

    End {
        $Context.Load($List)

        try {
            if ($PSCmdlet.ShouldProcess($Url, "Update list item ID '$($ItemId -join ', ')' in list $Name")) {
                $Context.ExecuteQuery()
                Write-Verbose "Successfully updated list item with ID '$($ItemId -join ', ')'" 
            }
        } catch {
            Write-Error "Failed to update item ID '$($ItemId -join ', ')' on list $Name at $($Url): $($_.Exception.Message)"
        }
    }
}

Function Remove-SPOListItem
{
    <#
        .SYNOPSIS
        .DESCRIPTION
        .PARAMETER Url
        .EXAMPLE
    #>
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Medium")]
    Param(
        [Parameter(Mandatory=$true, Position=1)]
        [string]
        $Url,
            
        [Parameter(Mandatory=$true, Position=2)]
        [string]
        $Name,

        [Parameter(Mandatory=$true, Position=3)]
        [string[]]
        $ItemId,

        [Parameter(Mandatory=$false)]
        [switch]
        $Force,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials
    )

    Begin {
        if ($Credential -eq $null) {
            throw "No SharePoint Online credentials detected. Please use Set-SPOCredentials and provide either a PSCredential object or Username and Password before using the SPO Commandlets."
        }

        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
        $Context.Credentials = $Credential

        $Context.Load($Context.Web)
        $Context.Load($Context.Web.Lists)
        $Context.ExecuteQuery()

        $List = $Context.Web.Lists.GetByTitle($Name)
        $Context.Load($List)
        try { $Context.ExecuteQuery() } 
        catch { 
            Write-Error "Error fetching $List from $($Url): $($_.Exception.Message)"
            return
        }
    }

    Process {
        $Counter = 0


        $ItemsFromList = @()

        foreach ($Id in $ItemId) {
            $ListItem = $List.GetItemById($Id)
            $Context.Load($ListItem)
            $ItemsFromList += $ListItem
        }

        try { $Context.ExecuteQuery() } 
        catch 
        { 
            Write-Error "Error fetching item ID '$ItemId' from $List at $($Url): $($_.Exception.Message)"
            return
        }

        if ($Force -eq $true) {
            foreach ($ListItem in $ItemsFromList) {
                $ListItem.DeleteObject()
            }
            $RemoveMsg = "Delete"
            $Local:ConfirmPreference = 'Medium'
        } else {
            foreach ($ListItem in $ItemsFromList) {
                $RecycleObj = $ListItem.Recycle()
            }
            $RemoveMsg = "Recycle"
        }
        $Context.Load($List)
        try {
            if ($PSCmdlet.ShouldProcess($Url, "$RemoveMsg list item ID '$ItemID' in list $Name")) {
                $Context.ExecuteQuery()
                Write-Verbose "Successfully $($RemoveMsg.ToLower())d item(s) from list." 
            }
        } catch {
            Write-Error "Failed to $($RemoveMsg.ToLower()) item(s) from list $Name at $($Url): $($_.Exception.Message)"
        }
    }

    End {
    }
}

Function Helper-PrepareListItemData
{
    [CmdletBinding()]
    Param(
        [string]
        $Url,
        [string]
        $Term, 
        [string]
        $FieldType, 
        [object]
        $Value,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials

    )

    $UserCtx = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $UserCtx.Credentials = $Credential

    switch ($FieldType)
    {
        "Url" {
            if ($Value.GetType().Name -ne 'Hashtable') { 
                throw "Specified URL field '$Term' missing required data type (hashtable of Url and Description)."
            }

            foreach ($UrlTerm in @('Url', 'Description')) {
                if ($UrlTerm -notin $Value.Keys) {
                    throw "Specified URL field '$UrlTerm' missing in data supplied for field '$Term'."
                }    
            }
            $UrlObj = New-Object Microsoft.SharePoint.Client.FieldUrlValue
            $UrlObj.Url = $Value["Url"]
            $UrlObj.Description = $Value["Description"]
            Write-Output ([Microsoft.SharePoint.Client.FieldUrlValue] $UrlObj)
        }
        "User" {
            $LookupId = [int] 0

            $UserObj = New-Object Microsoft.SharePoint.Client.FieldUserValue
            if ([int]::TryParse($Value, [ref] $LookupId)) {
                Write-Debug "Setting LookupId for term '$Term' to '$LookupId'"
                $UserObj.LookupId = $LookupId
            } else {
                Write-Debug "Performing lookup of user logon name '$($Value)'"
                $LookupUser = $UserCtx.Web.EnsureUser($Value)
                $UserCtx.Load($LookupUser)
                try { 
                    $UserCtx.ExecuteQuery()
                } catch {
                    Write-Error "Could not lookup user '$($Value)' to add as value for field '$($Term)': $($_.Exception.Message)"
                    return
                }
                Write-Debug "Found User, setting LookupId to '$($LookupUser.Id)'"
                $UserObj.LookupId = [int]$LookupUser.Id
            }
            Write-Debug "Setting User value for term '$Term' to '$($UserObj.LookupId)'"
            Write-Output ([Microsoft.SharePoint.Client.FieldUserValue] $UserObj)
        }
        "Lookup" {
            $LookupId = [int] 0

            $LookupObj = New-Object Microsoft.SharePoint.Client.FieldLookupValue
            if ([int]::TryParse($Value, [ref] $LookupId)) {
                Write-Debug "Setting LookupId for term '$Term' to '$LookupId'"
                $LookupObj.LookupId = $LookupId
            } else {
                Write-Debug "Setting LookupValue for term '$Term' to '$($Value)'"
                $LookupObj.LookupValue = $Value
            }
            Write-Debug "Setting Lookup value for term '$Term' to '$($Value)'"
            Write-Output ([Microsoft.SharePoint.Client.FieldLookupValue] $LookupObj)
        }
        "Text" {
            Write-Debug "Setting Text Value for term '$Term' to '$($Value)'"
            Write-Output ([string]$Value)
        }
        default {
            Write-Debug "Setting Other Value for term '$Term' to '$($Value)'"
            Write-Output $Value
        }
    }
}

Function Helper-AddNavigation
{
    [CmdletBinding()]
    Param(
        [string]
        $Url,

        [hashtable[]]
        $NavDataCol,

        [ValidateSet('Global', 'Quick')]
        [string]
        $NavType,

        [Parameter(Mandatory=$false)]
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]
        $Credential = $Script:SPOnline_Credentials
    )

    Begin {
        $WebCtx = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
        $WebCtx.Credentials = $Credential

        switch ($NavType) {
            "Global" {
                $NavigationCol = $WebCtx.Web.Navigation.TopNavigationBar
            }
            "Quick" {
                $NavigationCol = $WebCtx.Web.Navigation.QuickLaunch
            }
        }
    }

    Process {
        foreach ($NavData in $NavDataCol) {
            foreach ($UrlTerm in @('Url', 'Description')) {
                if ($NavData.GetType() -ne [Hashtable] -or $UrlTerm -notin $NavData.Keys) {
                    throw "Specified URL field '$UrlTerm' not present in supplied url."
                }    
            }

            $Navnode = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
            $NavNode.Url = $NavData.Url
            $NavNode.Title = $NavData.Description
            $NavNode.AsLastNode = $true
            $NavObj = $NavigationCol.Add($NavNode)
        }

        $WebCtx.Load($NavigationCol)
        try { 
            $WebCtx.ExecuteQuery()
            Write-Verbose "Successfully added navigation nodes to $NavType nav bar on site $($Url)."
        } catch {
            Write-Error "Could not add requested navigation nodes on site $($Url): $($_.Exception.Message)"
        }
    }

    End {
    }
}

Export-ModuleMember -Alias * -Function Get-*, New-*, Update-*, Remove-*, Add-*, Set-*, Write-*, Test-*, ConvertTo-*