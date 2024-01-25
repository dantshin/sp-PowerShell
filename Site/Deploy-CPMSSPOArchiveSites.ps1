<#
	.SYNOPSIS
		Deploy-CPMSSPOArchiveSites creates a set of SPO archival sites from CSV input

	.DESCRIPTION
		Deploy-CPMSSPOArchiveSites creates a set of SPO archival sites from CSV input file for CPMS sites that are living inside legacy Intranet

	.PARAMETER
		AdminUrl of SharePoint for the tenant
		RootsiteUrl of the tenant's SharePoint instance
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
# version: 1.0
#
# Revision history
# # 24 Jan 2024 Creation

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
#SiteName,SitePath,LegacySiteURL,SiteURL,SiteOwner,SecondarySiteOwner1,SecondarySiteOwner2
#CPMS Archival site for corporate,/sites/corporate,https://intranet/sites/corporate,corporate,siteowner@contoso.com,siteowner2@contoso.com,"Site Owner Group Name"
#CPMS Archival site for H10,/sites/H10,https://intranet/sites/H10,H10
$TimeZone = 10 #"UTCMINUS0500_EASTERN_TIME_US_AND_CANADA"
$SiteTemplateId = "BDR#0" #Document Center
# Need to trim trailing slashes in $RootsiteUrl
#Import CSV
$sites = Import-Csv -Path $InputCSVFile

#Connect to SharePoint Admin Center
Connect-PnPOnline -Url $AdminUrl -Interactive

$counter = 0
#Create sites
foreach ($site in $sites) {
    try {
		$siteURL = $RootsiteUrl + "/teams/" + $site.SiteURL + "_archive"
        New-PnPTenantSite -Template $SiteTemplateId -Title $site.SiteName -Url $siteURL -Owner $site.SiteOwner -TimeZone $TimeZone -Wait
		#Add Site Collection admins
		Set-PnPTenantSite -Identity $siteURL -Owners @($site.SecondarySiteOwner1,$site.SecondarySiteOwner2) -DenyAddAndCustomizePages:$false
		Write-Host "Created SPO site: $siteURL" -ForegroundColor Green
    }
    catch [System.Net.WebException], [System.IO.IOException] {
        Write-Host "Unable to create site $($site.SiteURL)" -ForegroundColor Red
    }
	$counter++
	Write-Progress -Activity "Processing SPO Site Creation" -Status "Processed: $counter of $($sites.Count) sites" -PercentComplete (($counter / $sites.Count) * 100)
}
