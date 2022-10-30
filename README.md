# Get-M365TenantUsageReport

Export M365 usage reports from a commercial, GCC or DoD tenant

## Getting Started with Get-M365TenantUsageReport

You must be running PowerShell 7 for this script to work due to dependencies.

Running this script you agree to install Microsoft.Graph PowerShell modules and consent to permissions on your system so you can connect to GraphAPI to export Intune policy information

### DESCRIPTION

Connect using Graph API (Beta) and export an M365 usage report. These are the reports that can be exported:

    getTeamsDeviceUsageUserDetail
    getTeamsDeviceUsageUserCounts
    getTeamsDeviceUsageDistributionUserCounts
    getTeamsUserActivityUserDetail
    getTeamsUserActivityCounts
    getEmailActivityUserDetail
    getEmailActivityCounts
    getEmailAppUsageUserDetail
    getEmailAppUsageAppsUserCounts
    getEmailAppUsageUserCounts
    getEmailAppUsageVersionsUserCounts
    getMailboxUsageDetail
    getMailboxUsageMailboxCounts
    getMailboxUsageQuotaStatusMailboxCounts
    getMailboxUsageStorage
    getOffice365ActiveUserDetail
    getOffice365ActiveUserCounts
    getOffice365ServicesUserCounts
    getM365AppUserDetail
    getM365AppUserCounts
    getM365AppPlatformUserCounts
    getOffice365GroupsActivityDetail
    getOffice365GroupsActivityCounts
    getOffice365GroupsActivityGroupCounts
    getOffice365GroupsActivityStorage
    getOffice365GroupsActivityFileCounts
    getOneDriveActivityUserDetail
    getOneDriveActivityUserCounts
    getOneDriveActivityFileCounts
    getOneDriveUsageAccountDetail
    getOneDriveUsageAccountCounts
    getOneDriveUsageFileCounts
    getOneDriveUsageStorage
    getSharePointActivityUserDetail
    getSharePointActivityFileCounts
    getSharePointActivityUserCounts
    getSharePointActivityPages
    getSharePointSiteUsageDetail
    getSharePointSiteUsageFileCounts
    getSharePointSiteUsageSiteCounts
    getSharePointSiteUsageStorage
    getSharePointSiteUsagePages
    getYammerActivityUserDetail
    getYammerActivityCounts
    getYammerActivityUserCounts
    getYammerDeviceUsageUserDetail
    getYammerDeviceUsageDistributionUserCounts
    getYammerDeviceUsageUserCounts
    getYammerGroupsActivityDetail
    getYammerGroupsActivityGroupCounts
    getYammerGroupsActivityCounts

### Examples

- EXAMPLE 1: PS C:\Get-M365TenantUsageReport

    Retrieves the getOffice365ActiveUserDetail M365 usage report (the default report)

- EXAMPLE 2: PS C:\GetUr

    Retrieves the default M365 usage report via alias

- EXAMPLE 3: PS C:\Get-M365TenantUsageReport -Endpoint Commercial

    Retrieves the default M365 usage report from a commercial endpoint

- EXAMPLE 4: PS C:\Get-M365TenantUsageReport -ResourceType getOffice365ActiveUserDetail -Format CSV

    Retrieves the getOffice365ActiveUserDetail M365 usage report and saves it in csv format

- EXAMPLE 5: PS C:\Get-M365TenantUsageReport -ResourceType getOffice365ActiveUserDetail -Format JSON

    Retrieves the getOffice365ActiveUserDetail M365 usage report and saves it it json format

- EXAMPLE 6: PS C:\Get-M365TenantUsageReport -ResourceType getOffice365ActiveUserDetail -LengthOfTime D7

    Retrieves the getOffice365ActiveUserDetail M365 usage report for the last 7 days (the default). You can select 7, 30, 90 or 180 days

### Note on file export

All policies will be exported in csv or json to "$env:Temp\ExportedIntunePolicies". This path can be changed if necessary.
