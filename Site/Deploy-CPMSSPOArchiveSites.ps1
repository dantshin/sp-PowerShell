<#
	.SYNOPSIS
		Deploy-CPMSSPOArchiveSites creates a set of SPO archival sites from CSV input

	.DESCRIPTION
		Deploy-CPMSSPOArchiveSites creates a set of SPO archival sites from CSV input file for CPMS sites that are living inside legacy Intranet

	.PARAMETER AdminUrl
		AdminUrl of SharePoint for the tenant
	.PARAMETER RootsiteUrl
		RootsiteUrl of the tenant's SharePoint instance
	.PARAMETER InputCSVFile
		InputCSVFile containing the list of sites to create

	.EXAMPLE
		Deploy-CPMSSPOArchiveSites -AdminUrl https://<tenant_name>-admin.sharepoint.com -RootsiteUrl https://<tenant_name>.sharepoint.com -InputCSVFile .\CSVFile.csv

	.INPUTS
		System.String System.String System.String

	.OUTPUTS
		System.String

	.NOTES
		AUTHOR: Daniel Tshin
		AUTHOR: Copilot
LASTEDIT: $(Get-Date)
# Written by: Daniel Tshin (daniel.tshin@undp.org) on 24 January 2024
#
# version: 1.1
#
# Revision history
# # 24 Jan 2024 Creation
# # 25 Jan 2024 Refactoring code to use Jobs for concurrency

.FUNCTIONALITY
		Deploy-CPMSSPOArchiveSites creates a set of SPO archival sites for CPMS sites that are living inside legacy Intranet

	.LINK
		about_functions_advanced

	.LINK
		about_comment_based_help

#>
#region command line parameters ###############################################
# Input Parameters
param(
	[Parameter(Mandatory = $true)][string]$AdminUrl = $(throw "Please specify the Admin URL for the tenant - https://<tenant_name>-admin.sharepoint.com"),
    [Parameter(Mandatory = $true)][string]$RootsiteUrl = $(throw "Please specify the Root Site URL for the tenant - https://<tenant_name>.sharepoint.com"),
    [Parameter(Mandatory = $true)][string]$InputCSVFile = $(throw "Please specify the input CSV file containing the list of sites to create")
);
#endregion ####################################################################
#Config Parameters
#$AdminUrl = "https://<tenant_name>-admin.sharepoint.com/"
#$RootsiteUrl = "https://<tenant_name>.sharepoint.com/"

#CSV file structure
#SiteName,SitePath,LegacySiteURL,siteUrl,SiteOwner,SecondarySiteOwner1,SecondarySiteOwner2
#CPMS Archival site for corporate,/sites/corporate,https://intranet/sites/corporate,corporate,siteowner@contoso.com,siteowner2@contoso.com,"Site Owner Group Name"
#CPMS Archival site for H10,/sites/H10,https://intranet/sites/H10,H10
$timeZone = 10 #"UTCMINUS0500_EASTERN_TIME_US_AND_CANADA"
$siteTemplateId = "BDR#0" #Document Center
# Need to trim trailing slashes in $RootsiteUrl
#Import CSV
$sites = Import-Csv -Path $InputCSVFile

#Connect to SharePoint Admin Center
Connect-PnPOnline -Url $AdminUrl -Interactive

function NewSPOSite {
	param(
		$site, $newSiteUrl, $siteTemplateId, $timeZone
	)
	try {

		#Create sites
		New-PnPTenantSite -Template $siteTemplateId `
			-Title $site.SiteName `
			-Url $newSiteUrl `
			-Owner $site.SiteOwner `
			-TimeZone $timeZone `
			-Wait
		#Add Site Collection admins
		Set-PnPTenantSite `
			-Identity $newSiteUrl `
			-Owners @($site.SecondarySiteOwner1,$site.SecondarySiteOwner2) `
			-DenyAddAndCustomizePages:$false

		Write-Host "Created SPO site: $newSiteUrl" -ForegroundColor Green
	}
    catch [System.Net.WebException], [System.IO.IOException] {
        Write-Host "Unable to create site $newSiteUrl for $($site.siteUrl)" -ForegroundColor Red
    }
}

$counter = 0

foreach ($site in $sites) {
	$newSiteUrl = $RootsiteUrl + "/teams/" + $site.siteUrl + "_archive"
	NewSPOSite -site $site -newSiteUrl $newSiteUrl -siteTemplateId $siteTemplateId -timeZone $timeZone

	$counter++
	Write-Progress -Activity "Processing SPO Site Creation" -Status "Processed: $counter of $($sites.Count) sites" -PercentComplete (($counter / $sites.Count) * 100)
}
