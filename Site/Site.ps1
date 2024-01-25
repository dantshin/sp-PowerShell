function Get-WebInfo {
<#
	.SYNOPSIS
		Get information about SharePoint Online web.

	.DESCRIPTION
		The Get-WebInfo function gets information of a SharePoint Online web including Title, URL, WebTemplate, List Count, Subwebs Count, MasterPageURL, InheritsMasterURL, InheritsCustomMasterURL, Owner, Size, Language, Locale, ParentWebId, SiteLogoPath.

	.PARAMETER SiteUrl
		Site url of the web to inspect.

	.EXAMPLE
		Get-WebInfo -SiteUrl "https://tenant.sharepoint.com/sites/MySite"

	.NOTES
		AUTHOR: Daniel Tshin, daniel.tshin@undp.org
		AUTHOR: Giuseppe Campanelli, suprememilanfan@gmail.com
#>
	param([string] $SiteUrl)

	try {
		Connect-PnPOnline -Url $SiteUrl -Interactive
		$web = Get-PnPWeb -Includes WebTemplate, Lists, SiteLogoUrl, Language, ParentWeb, RegionalSettings, Configuration
		$site = Get-PnPSite -Includes Owner, Usage
		$theme = Get-PnPTheme
		$feature = Get-PnPFeature -Scope Site -Identity "7c637b23-06c4-472d-9a9a-7c175762c5c4"
		if ($feature.DefinitionId -eq $null) {
			$lockdownMode = "Disabled"
		} else {
			$lockdownMode = "Enabled"
		}

		$obj = New-Object PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "Title" -Value $web.Title
		$obj | Add-Member -MemberType NoteProperty -Name "URL" -Value $web.Url
		$obj | Add-Member -MemberType NoteProperty -Name "Template" -Value "$($web.WebTemplate)#$($web.Configuration)"
		$obj | Add-Member -MemberType NoteProperty -Name "Number of Lists/Libraries" -Value $web.Lists.Count
		$obj | Add-Member -MemberType NoteProperty -Name "Site Logo Path" -Value $web.SiteLogoUrl
		$obj | Add-Member -MemberType NoteProperty -Name "Language" -Value $web.Language
		$obj | Add-Member -MemberType NoteProperty -Name "Locale" -Value $web.RegionalSettings.LocaleId
		$obj | Add-Member -MemberType NoteProperty -Name "Theme" -Value $theme.Name
		$obj | Add-Member -MemberType NoteProperty -Name "Owner" -Value $site.Owner.Title
		$obj | Add-Member -MemberType NoteProperty -Name "Size (% Used)" -Value $site.Usage.StoragePercentageUsed
		$obj | Add-Member -MemberType NoteProperty -Name "Limited-access user permission lockdown mode" -Value $lockdownMode

		$obj
	} catch {
		$exceptionName = $_.Exception.GetType().Name
		Write-Host -ForegroundColor "red" "[${exceptionName}] Cannot access: ${SiteUrl}"
	} finally {
		Disconnect-PnPOnline
	}
}

function Get-SiteInfo {
<#
	.SYNOPSIS
		Get information about SharePoint Online site.

	.DESCRIPTION
		The Get-SiteInfo function gets information of SharePoint Online sites including Title, URL, WebTemplate, List Count, Subwebs Count, MasterPageURL, InheritsMasterURL, InheritsCustomMasterURL, Owner, Size, Language, Locale, ParentWebId, SiteLogoPath.

	.PARAMETER SiteUrl
		Site url of the webs to inspect.

	.PARAMETER CSVFilePath
		CSV file path containing Site URLs of the webs to inspect.

	.PARAMETER ColumnName
		Column name containing Site URLs.

	.EXAMPLE
		Get-SiteInfo -SiteUrl "https://tenant.sharepoint.com/sites/MySite"

	.EXAMPLE
		Get-SiteInfo -CSVFilePath "C:/CSVFolder/MyCSVFile.csv" -ColumnName "Urls"

	.OUTPUTS
		PSObject

	.NOTES
		AUTHOR: Daniel Tshin, daniel.tshin@undp.org
		AUTHOR: Giuseppe Campanelli, suprememilanfan@gmail.com
#>
	param(
		[Parameter(Mandatory, ParameterSetName="Url")][string]$SiteUrl,
		[Parameter(Mandatory, ParameterSetName="CSVFile")][string]$CSVFilePath,
		[Parameter(Mandatory, ParameterSetName="CSVFile")][string]$ColumnName
	)

	$ErrorActionPreference = "Stop"

	if ($SiteUrl) {
		Write-Host "Getting Site Info for ${SiteUrl}"
		Get-WebInfo($SiteUrl)
	} else {
		$results = @()
		$csv = Import-Csv $CSVFilePath

		$csv.$ColumnName | ForEach-Object {
			Write-Host "Getting Site Info for ${_}"
			$results += Get-WebInfo($_)
		}

		$results
	}
}

function Set-WebLogo {
<#
	.SYNOPSIS
		Set the logo of a web.

	.DESCRIPTION
		The Set-WebLogo function sets the logo of a web.

	.PARAMETER WebUrl
		Specifies the WebUrl where the site logo is to be updated.

	.PARAMETER SiteLogoPath
		Specifies the path of the site logo that should be applied. This string should only contain the path, not the full path which includes the host.

	.EXAMPLE
		Set-WebLogo -WebUrl "https://tenant.sharepoint.com/teams/MyTeam" -SiteLogoPath "/SiteAssets/SiteLogo.png"

	.NOTES
		AUTHOR: Daniel Tshin, daniel.tshin@undp.org
		AUTHOR: Giuseppe Campanelli, suprememilanfan@gmail.com
#>
	param(
		[parameter(Mandatory)][string]$WebUrl,
		[parameter(Mandatory)][string]$SiteLogoPath
	)

	$web = Get-PnPWeb -Identity $WebUrl
	Write-Output ("Setting site logo on subweb: {0}" -f $web.ServerRelativeUrl)
	Set-PnPWeb -SiteLogoUrl $SiteLogoPath
}

function Set-SiteLogo {
<#
	.SYNOPSIS
		Set the logo of a site, and optionally, all subwebs of a site collection

	.DESCRIPTION
		The Set-SiteLogo function sets the site logo of a site collection, and optionally, all the subwebs in that site collection.

	.PARAMETER SiteUrl
		Specifies the SiteUrl where the site logo is to be updated.

	.PARAMETER Recursion
		Optional parameter, specifies whether the site logo is to be applied on all of the subwebs. Default is false.

	.PARAMETER SiteLogoPath
		Specifies the path of the site logo that should be applied. This string should only contain the path, not the full path which includes the host.

	.EXAMPLE
		Set-SiteLogo -SiteUrl "https://tenant.sharepoint.com/teams/MyTeam" -Recursion $true -SiteLogoPath "/SiteAssets/SiteLogo.png"

	.EXAMPLE
		Set-SiteLogo -SiteUrl "https://tenant.sharepoint.com/teams/MyTeam" -SiteLogoPath "/SiteAssets/SiteLogo.png"

	.LINK
		Script inspired by Jason Lee: http://www.jrjlee.com/2016/12/applying-logo-to-every-site-in.html

	.NOTES
		AUTHOR: Daniel Tshin, daniel.tshin@undp.org
		AUTHOR: Giuseppe Campanelli, suprememilanfan@gmail.com
#>
	param(
		[Parameter(Mandatory)][string]$SiteUrl = $(throw "Please specify the URL for the site where you want to change the site logo."),
		[Parameter(Mandatory=$false)][string]$Recursion,
		[Parameter(Mandatory)][string]$SiteLogoPath = $(throw "Please specify the path of the site logo. it should start with '/' and should not include the host ('https://tenant.sharepoint.com/').")
	)

	try {
		Connect-PnPOnline -Url $SiteUrl -Interactive
		$rootWeb = Get-PnPWeb

		# Set logo on root web
		Set-PnPWeb -SiteLogoUrl $SiteLogoPath
		Write-OutPut ("Setting site logo on web: {0}" -f $rootWeb.ServerRelativeUrl)

		if ($Recursion -eq $true) {
			$subwebs = Get-PnPSubWeb -Recurse
			foreach ($subweb in $subwebs) {
				Set-WebLogo($subweb.ServerRelativeUrl, $SiteLogoPath)
			}
		}

	} catch {
		Write-Host -f Red "Error in Script: " $_.Exception.Message
	} finally {
		Disconnect-PnPOnline
	}
}
