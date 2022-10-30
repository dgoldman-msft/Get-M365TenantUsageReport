function New-LoggingDirectory {
    <#
        .SYNOPSIS
            Create directories

        .DESCRIPTION
            Create the root and all subfolder needed for logging

        .PARAMETER LoggingPath
            Logging Path

        .PARAMETER SubFolder
            Switch to indicated we are creating a subfolder

        .PARAMETER SubFolderName
            Subfolder Name

        .EXAMPLE
            PS C:\New-LoggingDirectory -SubFolder SubFolderName

        .NOTES
            Internal function
    #>

    [OutputType('System.IO.Folder')]
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]
        $LoggingPath,

        [switch]
        $SubFolder,

        [string]
        $SubFolderName
    )

    begin {
        if (-NOT($SubFolder)) {
            Write-Verbose "Creating directory: $($LoggingPath)"
        }
        else {
            Write-Verbose "Creating directory: $LoggingPath\$SubFolderName)"
        }
    }

    process {
        try {
            # Leaving this here in case the root directory gets deleted between executions so we will re-create it again
            if (-NOT(Test-Path -Path $LoggingPath)) {
                if (New-Item -Path $LoggingPath -ItemType Directory -ErrorAction Stop) {
                    Write-Verbose "$LoggingPath directory created!"
                }
                else {
                    Write-Verbose "$($LoggingPath) already exists!"
                }
            }
            if ($SubFolder) {
                if (-NOT(Test-Path -Path $LoggingPath\$SubFolderName)) {
                    if (New-Item -Path $LoggingPath\$SubFolderName -ItemType Directory -ErrorAction Stop) {
                        Write-Verbose "$LoggingPath\$SubFolderName directory created!"
                    }
                    else {
                        Write-Verbose "$($SubFolderName) already exists!"
                    }
                }
            }
        }
        catch {
            Write-Output "Error: $_"
            return
        }
    }
}

function Get-M365TenantUsageReport {
    <#
        .SYNOPSIS
            Export M365 usage reports

        .DESCRIPTION
            Connect using Graph API and export M365 usage reports from a global, GCC or DoD tenant

        .PARAMETER Endpoint
            Endpoint to connect to

        .PARAMETER Format
            Format to save in (csv or json)

        .PARAMETER LoggingPath
            Logging path

        .PARAMETER LengthOfTime
            Days of logs to search for. 7 (default), 30, 90 or 180 days

        .PARAMETER ResourceType
            Graph namespace to retrieve

        .PARAMETER ShowModuleInfoInVerbose
            Used to troubleshoot module install and import

        .EXAMPLE
            PS C:\Get-M365TenantUsageReport

            Retrieves the getOffice365ActiveUserDetail M365 usage report (the default report)

        .EXAMPLE
            PS C:\GetUr

            Retrieves the default M365 usage report via alias

        .EXAMPLE
            PS C:\Get-M365TenantUsageReport -Endpoint Commercial

            Retrieves the default M365 usage report from a commercial endpoint

        .EXAMPLE
            PS C:\Get-M365TenantUsageReport -ResourceType getOffice365ActiveUserDetail -Format CSV

            Retrieves the getOffice365ActiveUserDetail M365 usage report and saves it in csv format

        .EXAMPLE
            PS C:\Get-M365TenantUsageReport -ResourceType getOffice365ActiveUserDetail -Format JSON

            Retrieves the getOffice365ActiveUserDetail M365 usage report and saves it it json format

        .EXAMPLE
            PS C:\Get-M365TenantUsageReport -ResourceType getOffice365ActiveUserDetail -LengthOfTime D7

            Retrieves the getOffice365ActiveUserDetail M365 usage report for the last 7 days (the default). You can select 7, 30, 90 or 180 days

        .NOTES
            https://learn.microsoft.com/en-us/graph/filter-query-parameter
            https://learn.microsoft.com/en-us/powershell/microsoftgraph/get-started?view=graph-powershell-1.0
            https://learn.microsoft.com/en-us/graph/api/resources/intune-shared-devicemanagement?view=graph-rest-beta
   #>

    [OutputType('PSCustomObject')]
    [CmdletBinding()]
    [Alias('GetUr')]
    param(
        [ValidateSet('Global', 'GCC', 'DOD')]
        [parameter(Position = 0)]
        [string]
        $Endpoint = 'Global',

        [ValidateSet('CSV', 'JSON')]
        [parameter(Position = 1)]
        [string]
        $Format = "CSV",

        [parameter(Position = 2)]
        $LoggingPath = "$env:Temp\ExportedM365Reports",

        [ValidateSet('D7', 'D30', 'D90', 'D180')]
        [parameter(Position = 3)]
        [string]
        $LengthOfTime = 'D7',

        [ValidateSet('getTeamsDeviceUsageUserDetail', 'getTeamsDeviceUsageUserCounts', 'getTeamsDeviceUsageDistributionUserCounts', 'getTeamsUserActivityUserDetail', `
                'getTeamsUserActivityCounts', 'getEmailActivityUserDetail', 'getEmailActivityCounts', 'getEmailAppUsageUserDetail', 'getEmailAppUsageAppsUserCounts', `
                'getEmailAppUsageUserCounts', 'getEmailAppUsageVersionsUserCounts', 'getMailboxUsageDetail', 'getMailboxUsageMailboxCounts', 'getMailboxUsageQuotaStatusMailboxCounts', `
                'getMailboxUsageStorage', 'getOffice365ActiveUserDetail', 'getOffice365ActiveUserCounts', 'getOffice365ServicesUserCounts', 'getM365AppUserDetail', 'getM365AppUserCounts', `
                'getM365AppPlatformUserCounts', 'getOffice365GroupsActivityDetail', 'getOffice365GroupsActivityCounts', 'getOffice365GroupsActivityGroupCounts', 'getOffice365GroupsActivityStorage', `
                'getOffice365GroupsActivityFileCounts', 'getOneDriveActivityUserDetail', 'getOneDriveActivityUserCounts', 'getOneDriveActivityFileCounts', 'getOneDriveUsageAccountDetail', `
                'getOneDriveUsageAccountCounts', 'getOneDriveUsageFileCounts', 'getOneDriveUsageStorage', 'getSharePointActivityUserDetail', 'getSharePointActivityFileCounts', 'getSharePointActivityUserCounts', `
                'getSharePointActivityPages', 'getSharePointSiteUsageDetail', 'getSharePointSiteUsageFileCounts', 'getSharePointSiteUsageSiteCounts', 'getSharePointSiteUsageStorage', `
                'getSharePointSiteUsagePages', 'getYammerActivityUserDetail', 'getYammerActivityCounts', 'getYammerActivityUserCounts', 'getYammerDeviceUsageUserDetail', 'getYammerDeviceUsageDistributionUserCounts', `
                'getYammerDeviceUsageUserCounts', 'getYammerGroupsActivityDetail', 'getYammerGroupsActivityGroupCounts', 'getYammerGroupsActivityCounts' )]
        [parameter(Position = 4)]
        [string]
        $ResourceType = 'getOffice365ActiveUserDetail',

        [switch]
        $ShowModuleInfoInVerbose
    )

    begin {
        Write-Output "Retrieving M365 $($ResourceType) Report"
        $parameters = $PSBoundParameters
        $modules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Reports")
        $successful = $false
    }

    process {
        if ($PSVersionTable.PSEdition -ne 'Core') {
            Write-Output "You need to run this script using PowerShell core due to dependencies."
            return
        }

        try {
            # Create root directory
            New-LoggingDirectory -LoggingPath $LoggingPath
        }
        catch {
            Write-Output "Error: $_"
            return
        }

        try {
            foreach ($module in $modules) {
                if ($found = Get-Module -Name $module -ListAvailable | Sort-Object Version | Select-Object -First 1) {
                    if (Import-Module -Name $found -ErrorAction Stop -Verbose:$ShowModuleInfoInVerbose -PassThru) {
                        Write-Verbose "$found imported!"
                        $successful = $true
                    }
                    else {
                        Throw "Error importing $($found). Please Run Export-IntunePolicy -Verbose -ShowModuleInfoInVerbose"
                    }
                }
                else {
                    Write-Output "$module not found! Installing module $($module) from the PowerShell Gallery"
                    if (Install-Module -Name $module -Repository PSGallery -Force -Verbose:$ShowModuleInfoInVerbose -PassThru) {
                        Write-Verbose "$module installed successfully! Importing $($module)"
                        if (Import-Module -Name $module -ErrorAction Stop -Verbose:$ShowModuleInfoInVerbose -PassThru) {
                            Write-Verbose "$module imported successfully!"
                            $successful = $true
                        }
                        else {
                            Throw "Error importing $($found). Please Run Export-IntunePolicy -Verbose -ShowModuleInfoInVerbose"
                        }
                    }
                }
            }
        }
        catch {
            Write-Output "Error: $_"
            return
        }

        try {
            if ($successful) {
                Select-MgProfile -Name "beta" -ErrorAction Stop
                Write-Verbose "Using MGProfile (Beta)"
                If ($Endpoint -eq 'Global') { Connect-MgGraph -Scopes "User.Read.All", "Reports.Read.All" -Environment Global -ForceRefresh -ErrorAction Stop }
                if ($Endpoint -eq 'GCC') { Connect-MgGraph -Scopes "User.Read.All", "Reports.Read.All" -Environment USGov -ForceRefresh -ErrorAction Stop }
                if ($Endpoint -eq 'Dod') { Connect-MgGraph -Scopes "User.Read.All", "Reports.Read.All" -Environment USGovDoD -ForceRefresh -ErrorAction Stop }
            }
            else {
                Write-Output "Error: Unable to connect to the Graph endpoint. $_"
                return
            }
        }
        catch {
            Write-Output "Error: $_"
            return
        }

        try {
            switch ($Format) {
                'CSV' { $searchFormat = 'text/csv' }
                'JSON' { $searchFormat = 'application/json' }
            }
            switch ($Endpoint) {
                'Global' {
                    $uri = "https://graph.microsoft.com/beta/reports/$ResourceType(period='$LengthOfTime')?`$format=$searchFormat"
                    continue
                }
                'GCC' {
                    $uri = "https://graph.microsoft.us/beta/reports/$ResourceType(period='$LengthOfTime')?`$format=$searchFormat"
                    continue
                }
                'DoD' {
                    $uri = "https://dod-graph.microsoft.us/beta/reports/$ResourceType(period='$LengthOfTime')?`$format=$searchFormat"
                    continue
                }
            }

            Write-Output "Querying Graph uri: $($uri)"
            switch ($Format) {
                'CSV' {
                    New-LoggingDirectory -LoggingPath $LoggingPath -SubFolder $ResourceType
                    Invoke-MgGraphRequest -Method GET -Uri $uri -OutputFilePath (Join-Path -Path $LoggingPath\$ResourceType -ChildPath ($ResourceType + ".csv")) -StatusCodeVariable statusCode

                    if ($statusCode -eq 200) {
                        Write-Verbose "Return status code: $($statusCode)"
                        Write-Verbose "Saving $(Join-Path -Path $LoggingPath\$ResourceType -ChildPath ($ResourceType + ".csv"))"
                        $successful = $true
                        continue
                    }
                    else {
                        Throw "Invoke-MgGraphRequest Error: $($statusCode)"
                    }
                }
                'JSON' {
                    New-LoggingDirectory -LoggingPath $LoggingPath -SubFolder $ResourceType
                    if ($statusCode -eq 200) {
                        $report = Invoke-MgGraphRequest -Method GET -Uri $uri -StatusCodeVariable statusCode
                        Write-Verbose "Return status code: $($statusCode)"
                        [PSCustomObject]$report | ConvertTo-Json -Depth 10 | Set-Content (Join-Path -Path $LoggingPath\$ResourceType -ChildPath $($ResourceType + ".json")) -ErrorAction Stop -Encoding UTF8
                        Write-Verbose "Saving $(Join-Path -Path $LoggingPath\$ResourceType  -ChildPath $($ResourceType + ".json"))"
                        continue
                    }
                    else {
                        Throw "Invoke-MgGraphRequest Error: $($statusCode)"
                    }
                }
            }
        }
        catch {
            Write-Output "Error: $_"
        }
    }

    end {
        if (($report.Count -gt 0) -and ($parameters.ContainsKey('SaveResultsToJSON')) -or ($successful)) {
            Write-Output "`nResults exported to: $($LoggingPath)`nCompleted!"
        }
        else {
            $null = Disconnect-MgGraph
            Write-Output "Completed!"
        }
    }
}