function Update-InstalledModule
{
    <#
    .SYNOPSIS
    Updates all locally installed PowerShell modules to the latest available versions.

    .DESCRIPTION
    Searches for all installed PowerShell modules and updates them to the latest online versions from PSGallery.

    Updates 'PackageManagement' and 'PowerShellGet' first if needed, as they are prerequisites for updating other modules.

    By default, installs to the CurrentUser scope. Use the -AllUsers switch to install to the AllUsers scope (requires elevation).

    The script presents a formatted list of all modules with their current and available versions.

    Differences with PowerShell's built-in 'Update-Module' command:
        - 'Update-Module' will not update 'PackageManagement' and 'PowerShellGet'.
        - It shows no data while operating unless you use '-Verbose', which displays too much information.

    The Az and Microsoft.Graph meta-modules are excluded to avoid reinstalling all submodules. Individual submodules (e.g., Az.Accounts, Microsoft.Graph.Users) are updated individually.

    .PARAMETER AllUsers
    Install updates to the AllUsers scope. Requires running as administrator.

    .EXAMPLE
    Update-InstalledModule
    Updates all locally installed modules to the CurrentUser scope.

    .EXAMPLE
    Update-InstalledModule -AllUsers
    Updates all locally installed modules to the AllUsers scope (requires elevation).

    .NOTES
    Filename:       Update-InstalledModule.psm1
    Contributors:   Kieran Walsh
    Created:        2021-01-09
    Last Updated:   2026-03-12
    Version:        2.00.00
    ProjectUri:     https://github.com/kieranwalsh/Update-InstalledModule
    #>

    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$AllUsers
    )

    if ($PSVersionTable.PSVersion -lt [version]'5.0.0')
    {
        Write-Warning -Message "This function requires PowerShell 5.0 or newer. You are running $($PSVersionTable.PSVersion)."
        return
    }

    if ($ExecutionContext.SessionState.LanguageMode -eq 'ConstrainedLanguage')
    {
        Write-Warning -Message 'Constrained Language mode is enabled. The function cannot continue.'
        return
    }

    if ($AllUsers)
    {
        $CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
        $IsAdmin = (New-Object -TypeName 'Security.Principal.WindowsPrincipal' -ArgumentList $CurrentUser).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
        if (-not $IsAdmin)
        {
            Write-Warning -Message 'The -AllUsers switch requires running as administrator.'
            return
        }
        $InstallScope = 'AllUsers'
    } else
    {
        $InstallScope = 'CurrentUser'
    }

    $StartTime = Get-Date

    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $PositiveSymbol = [char]0x2713
    $NegativeSymbol = [char]0x2717

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $NewSessionRequired = $false

    try
    {
        $RegisteredRepositories = Get-PSRepository -ErrorAction 'Stop' -WarningAction 'Stop'
    } catch
    {
        Write-Warning -Message "Unable to query 'PSGallery' online. Check your proxy/firewall settings."
        return
    }

    if ($RegisteredRepositories -notmatch 'PSGallery')
    {
        try
        {
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy 'Trusted' -ErrorAction 'Stop'
        } catch
        {
            Write-Warning -Message 'Unable to set the PSRepository.'
            return
        }
    }

    #Region PackageManagement check
    Write-Output "Checking which version of 'PackageManagement' is installed locally."
    $PackageManagement = (Get-Module -ListAvailable -Name 'PackageManagement' | Sort-Object -Property 'Version' -Descending)[0]
    Write-Output "Found 'PackageManagement' version '$($PackageManagement.Version)'."

    if ([version]$PackageManagement.Version -lt [version]'1.4.7')
    {
        Write-Output "An updated version is required. Attempting to install 'NuGet'."
        try
        {
            $NugetInstall = Install-PackageProvider -Name 'NuGet' -Force -Scope $InstallScope -ErrorAction 'Stop'
            Write-Output "Successfully installed 'NuGet' version '$($NugetInstall.Version)'."
        } catch
        {
            if ($_.Exception.Message -match 'No match was found for the specified search criteria for the provider')
            {
                Write-Warning -Message 'Unable to find packages online. Check proxy settings.'
            } else
            {
                Write-Warning -Message "Failed to install NuGet: $($_.Exception.Message)"
            }
            return
        }

        Write-Output "Searching for a newer version of 'PackageManagement'."
        try
        {
            $OnlinePackageManagement = Find-Module -Name 'PackageManagement' -Repository 'PSGallery' -ErrorAction 'Stop'
        } catch
        {
            Write-Warning -Message "Unable to find 'PackageManagement' online. Check your internet connection."
            return
        }

        try
        {
            $OnlinePackageManagement | Install-Module -Force -SkipPublisherCheck -Scope $InstallScope -ErrorAction 'Stop'
            Write-Output "Successfully installed 'PackageManagement' version '$($OnlinePackageManagement.Version)'."
        } catch
        {
            Write-Warning -Message "Failed to install 'PackageManagement'."
            return
        }

        Write-Output "Close PowerShell and re-open it to use the new 'PackageManagement' module."
        $NewSessionRequired = $true
    }
    #EndRegion PackageManagement check

    #Region PowerShellGet check
    Write-Output "Checking which version of 'PowerShellGet' is installed locally."
    $PowerShellGet = (Get-Module -ListAvailable -Name 'PowerShellGet' | Sort-Object -Property 'Version' -Descending)[0]
    Write-Output "Found 'PowerShellGet' version '$($PowerShellGet.Version)'."

    if ([version]$PowerShellGet.Version -lt [version]'1.6.0')
    {
        $OnlinePSGet = Find-Module -Name 'PowerShellGet' -Repository 'PSGallery'
        Write-Output "Version '$($OnlinePSGet.Version)' found online, will attempt to update."
        try
        {
            $OnlinePSGet | Install-Module -Force -SkipPublisherCheck -Scope $InstallScope -ErrorAction 'Stop'
            Write-Output "Successfully installed 'PowerShellGet' version '$($OnlinePSGet.Version)'. Close PowerShell and re-open it to use the new module."
            $NewSessionRequired = $true
        } catch
        {
            Write-Warning -Message "Failed to install the latest version of 'PowerShellGet'."
            return
        }
    }
    #EndRegion PowerShellGet check

    if ($NewSessionRequired)
    {
        return
    }

    #Region Update modules
    $FailedModules = [System.Collections.Generic.List[object]]::new()
    Write-Output 'Searching for all locally installed modules.'

    $InstalledModules = Get-InstalledModule |
    Where-Object -FilterScript {
        $_.Name -notmatch '^(PackageManagement|PowerShellGet|Az|Microsoft\.Graph)$'
    } |
    Group-Object -Property 'Name' |
    ForEach-Object { $_.Group | Select-Object -First 1 } |
    Sort-Object -Property 'Name'

    if (-not $InstalledModules)
    {
        Write-Output 'No installed modules found.'
        return
    }

    $SuccessfulUpdates = 0
    Write-Output "Found $($InstalledModules.Count) module(s). Checking for newer versions online."

    $MaxNameWidth = ($InstalledModules.Name | Sort-Object -Property 'Length' -Descending | Select-Object -First 1).Length + 3
    $MaxVersionWidth = [Math]::Max(
        ($InstalledModules.Version | ForEach-Object { "$_".Length } | Sort-Object -Descending | Select-Object -First 1) + 3,
        12
    )

    foreach ($InstalledModule in $InstalledModules)
    {
        Write-Host -Object ("{0,-$MaxNameWidth}" -f $InstalledModule.Name) -NoNewline

        try
        {
            $Module = Get-InstalledModule -Name $InstalledModule.Name |
            Sort-Object -Property { [version](($_.Version -split '-')[0]) } -Descending |
            Select-Object -First 1
            $LatestAvailableOnline = Find-Module -Name $InstalledModule.Name -ErrorAction 'Stop'
        } catch
        {
            $FailedModules.Add($InstalledModule)
            Write-Host -Object $NegativeSymbol -ForegroundColor 'Red'
            continue
        }

        Write-Host -Object ("{0,-$MaxVersionWidth} " -f $Module.Version) -NoNewline

        $InstalledVersionString = ($Module.Version -replace '[a-z]*', '').Replace('-', '.') -replace '\.$', ''
        $OnlineVersionString = ($LatestAvailableOnline.Version -replace '[a-z]*', '').Replace('-', '.') -replace '\.$', ''
        $InstalledVersion = [version]$InstalledVersionString
        $OnlineVersion = [version]$OnlineVersionString
        # Normalize to 4-part versions to avoid -1 Revision mismatch (e.g. 6.3.1 vs 6.3.1.0)
        if ($InstalledVersion.Revision -eq -1)
        {
            $InstalledVersion = [version]"$InstalledVersionString.0"
        }
        if ($OnlineVersion.Revision -eq -1)
        {
            $OnlineVersion = [version]"$OnlineVersionString.0"
        }

        if ($InstalledVersion -ge $OnlineVersion)
        {
            Write-Host -Object $PositiveSymbol -ForegroundColor 'Green'
        } else
        {
            $PublishedDate = Get-Date($LatestAvailableOnline.PublishedDate) -Format 'yyyy-MM-dd'
            Write-Host -Object "Online: '$($LatestAvailableOnline.Version)' (Published $PublishedDate) - updating... " -ForegroundColor 'Yellow' -NoNewline

            $UpdateParams = @{
                Name          = $Module.Name
                AcceptLicense = $true
                Force         = $true
                Scope         = $InstallScope
                ErrorAction   = 'Stop'
            }

            try
            {
                Update-Module @UpdateParams
                Write-Host -Object $PositiveSymbol -ForegroundColor 'Green'
                $SuccessfulUpdates++
            } catch
            {
                Write-Host -Object "$NegativeSymbol Update failed, attempting reinstall. " -ForegroundColor 'Yellow' -NoNewline
                try
                {
                    $InstallParams = @{
                        Name               = $Module.Name
                        AcceptLicense      = $true
                        AllowClobber       = $true
                        Force              = $true
                        Scope              = $InstallScope
                        SkipPublisherCheck = $true
                        ErrorAction        = 'Stop'
                    }
                    Install-Module @InstallParams 3>$null
                    Write-Host -Object $PositiveSymbol -ForegroundColor 'Green'
                    $SuccessfulUpdates++
                } catch
                {
                    Write-Host -Object $NegativeSymbol -ForegroundColor 'Red'
                }
            }
        }
    }
    #EndRegion Update modules

    #Region PowerShellGet final check
    $OnlinePSGet = Find-Module -Name 'PowerShellGet' -Repository 'PSGallery' -ErrorAction 'SilentlyContinue'
    if ($OnlinePSGet)
    {
        $LocalPSGetVersion = [version](($PowerShellGet.Version -split '-')[0])
        $OnlinePSGetVersion = [version](($OnlinePSGet.Version -split '-')[0])
        if ($LocalPSGetVersion -lt $OnlinePSGetVersion)
        {
            Write-Output "A newer version of 'PowerShellGet' is available online, attempting to update."
            try
            {
                $OnlinePSGet | Install-Module -Force -SkipPublisherCheck -Scope $InstallScope -ErrorAction 'Stop'
                Write-Output "Successfully installed 'PowerShellGet' version '$($OnlinePSGet.Version)'. Close PowerShell and re-open to use the new module."
            } catch
            {
                Write-Warning -Message "Failed to install the latest version of 'PowerShellGet'."
            }
        }
    }
    #EndRegion PowerShellGet final check

    #Region Summary
    if ($FailedModules.Count -gt 0)
    {
        Write-Output 'Unable to find these modules online:'
        $FailedModules | ForEach-Object { Write-Output "  $($_.Name)" }
    }

    $EndTime = Get-Date
    $TimeTaken = ''
    $TakenSpan = New-TimeSpan -Start $StartTime -End $EndTime

    if ($TakenSpan.Hours)
    {
        $TimeTaken += "$($TakenSpan.Hours) hours, $($TakenSpan.Minutes) minutes, "
    } elseif ($TakenSpan.Minutes)
    {
        $TimeTaken += "$($TakenSpan.Minutes) minutes, "
    }
    $TimeTaken += "$($TakenSpan.Seconds) seconds"

    if ($SuccessfulUpdates)
    {
        Write-Output "Successfully updated $SuccessfulUpdates module(s) in $TimeTaken."
    } else
    {
        Write-Output "Completed in $TimeTaken. No updates were required."
    }
    #EndRegion Summary
}
