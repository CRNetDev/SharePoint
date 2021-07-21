<#
.NOTES
    Name: Export_SPO_permission_report.ps1
    Author: Bela Bana | https://github.com/belabana
    Story: I requested to get a comprehensive permission report of the customer's exported SharePoint Online sites,
           and needed to find a way to bypass HTTP 429 error given by Azure API due to the high number of requests.
           Referece: https://docs.microsoft.com/en-us/troubleshoot/azure/general/request-throttling-http-403
    Classification: Public
    Disclaimer: Author does not take responsibility for any unexpected outcome that could arise from using this script.
                Please always test it in a virtual lab or UAT environment before executing it in production environment.
        
.SYNOPSIS
    Import the list of site collections and query SharePoint Online to get and export the permissions for each site.

.DESCRIPTION
    This script will help service desk to collect information about active permissions in association with each site collection.
    It requires and works from an up to date list of sites provided in a CSV file. Reference: C:\temp\SPO_Sites.csv
    It uses the given URL and administrator credentials to connect to and query the targeted SharePoint Online tenancy.
    The permission report will be exported to a CSV file. Path: “C:\temp”
    Transcript and stopwatch are also added to calculate runtime and write output to a .log file. Path: “C:\temp”
    It should be executed with the credentials of a user with full SharePoint Online access and from a computer with Microsoft.Online.SharePoint.PowerShell module.
    To avoid throttling issue with Azure, you might need to increase the time to sleep value in the for loop. Line 85.

.PARAMETER SPOAdminURL
    The hyperlink pointing to the customer's SharePoint Online Admin Portal.
    Reference: "https://yourcompanydomain-admin.sharepoint.com"

.PARAMETER SPOAdminCredential
    Credentials of a user with full SharePoint Online Administrator access.
#>

param(
    [parameter(Mandatory=$true)]
    [System.String] $SPOAdminURL,

    [Parameter(Mandatory=$true)]
    [System.Management.Automation.PSCredential] $SPOAdminCredential
)
begin {
$TranscriptPath = "C:\temp\Export_SPO_permission_report_$(Get-Date -Format yyyy-MM-dd-HH-mm).log"
Start-Transcript -Path $TranscriptPath
Write-Host -ForegroundColor Yellow "Script started: "(Get-Date -Format "dddd MM/dd/yyyy HH:mm")
#Variables
$file = "C:\temp\SPO_Sites.csv"
$Report = @()
#Confirm if SPO_Sites.csv file exists
Write-Host -ForegroundColor Yellow "Checking if SPO_Sites.csv file exists.."
if (Test-Path -Path $file -PathType Leaf) {
    Write-Host -ForegroundColor Green "File [$file] was found."
}
else {
    Write-Host -ForegroundColor Red "The file [$file] was not found. Terminating process."
    Exit
}
#Verify SharePoint Online URL and request confirmation
Write-Warning "Would you like to connect to the following system:`n$SPOAdminURL" -WarningAction Inquire
}
process {
#Start a timer to measure runtime
$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
#Importing sites from SPO_Sites.csv file
Write-Host -ForegroundColor Yellow "Importing sites from file.."
$Sitelist = Import-CSV $file
#Connect to SharePoint Online
try {
    Connect-SPOService -Url $SPOAdminURL -Credential $SPOAdminCredential
    Write-Host -ForegroundColor Green "Connected to SharePoint Online."
}
catch {
    Write-Host -ForegroundColor Red "Unable to connect to SharePoint Online. Terminating process."
    Write-Error -Message "$_" -ErrorAction Stop
    return;
}

#Get security groups for each site
ForEach ($Site in $Sitelist) {
    try {
        Write-Host "Processing site called " -nonewline; Write-Host $Site.Title -ForegroundColor Green
        Start-Sleep -Seconds 1
        $SiteSecGroups = Get-SPOSiteGroup -Site $Site.URL
        #Write-Host "Getting security groups for " -nonewline; Write-Host $Site.Title -ForegroundColor Green

        #Get and extend the report with users and permissions found in each security group
        ForEach ($SecGroup in $SiteSecGroups) {
            Start-Sleep -Seconds 5
            $Report += New-Object PSObject -Property @{
                "Site URL" = $Site.URL
                "Security Group" = $SecGroup.Title
                "Permissions" = $SecGroup.Roles -join ","
                "Users" = $SecGroup.Users -join ","
            }
        }
    }
    catch {
        Write-Host -ForegroundColor Red "Unable to query the given site. Skipping to the next one."
        Write-Error -Message "$_"
    }
}
#Export the data to CSV
$Report | Export-CSV C:\temp\Export_SPO_permission_report_$(Get-Date -Format yyyy-MM-dd-HH-mm).csv -NoTypeInformation
Write-host -ForegroundColor Green "Permission report has been exported Path: C:\temp"
Write-Host -ForegroundColor Yellow "Script ended: "(Get-Date -Format "dddd MM/dd/yyyy HH:mm")
Write-Host -ForegroundColor Yellow "Script completed in" $Stopwatch.Elapsed.Hours "hours and" $Stopwatch.Elapsed.Minutes "minutes and" $Stopwatch.Elapsed.Seconds "seconds."
Stop-Transcript
}