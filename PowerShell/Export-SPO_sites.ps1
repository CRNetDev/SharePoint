<#
.NOTES
    Name: Export_SPO_sites.ps1
    Author: Bela Bana | https://github.com/belabana
    Story: As part of a project, I needed to export all SharePoint Online site collections of the customer.
    Classification: Public
    Disclaimer: Author does not take responsibility for any unexpected outcome that could arise from using this script.
                Please always test it in a virtual lab or UAT environment before executing it in production environment.
        
.SYNOPSIS
    Query SharePoint Online to get and export the customer's SharePoint Online site collections to a CSV file.

.DESCRIPTION
    This script will help service desk to build a full list of the customer's SharePoint Online site collections. 
    It uses the provided URL and administrator credentials to connect to and query the targeted SharePoint Online tenancy.
    The requested details will be exported to a CSV file. Path: “C:\temp”
    Transcript and stopwatch are also added to calculate runtime and write output to a .log file. Path: “C:\temp”
    It should be executed with full SharePoint Online access and from a computer with Microsoft.Online.SharePoint.PowerShell module.

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
$TranscriptPath = "C:\temp\Export_SPO_sites_$(Get-Date -Format yyyy-MM-dd-HH-mm).log"
Start-Transcript -Path $TranscriptPath
Write-Host -ForegroundColor Yellow "Script started: "(Get-Date -Format "dddd MM/dd/yyyy HH:mm")
#Variables
$Report = @()
#Verify SharePoint Online URL and request confirmation
Write-Warning "Would you like to connect to the following system:`n$SPOAdminURL" -WarningAction Inquire
}
process {
#Start a timer to measure runtime
$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
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
#Get and export all SharePoint Online sites with URL
$ErrorActionPreference = 'Continue'
$Sites = Get-SPOSite -Limit ALL 
$Sites | Select Title,URL | Export-CSV C:\temp\SPO_Sites_$(Get-Date -Format yyyy-MM-dd-HH-mm).csv -NoTypeInformation

#Export the data to a CSV file
$Report | Export-CSV C:\temp\SPO_Sites.csv -NoTypeInformation
Write-host -ForegroundColor Green "SharePoint Online site collections have been exported. Path: C:\temp"
Write-Host -ForegroundColor Yellow "Script ended: "(Get-Date -Format "dddd MM/dd/yyyy HH:mm")
$StopWatch.Stop()
Write-Host -ForegroundColor Yellow "Script completed in" $Stopwatch.Elapsed.Hours "hours and" $Stopwatch.Elapsed.Minutes "minutes and" $Stopwatch.Elapsed.Seconds "seconds."
Stop-Transcript
}