<#
	.SYNOPSIS
	Cmdlet to create a new SharePoint/Microsoft List from a schema definition file

	.DESCRIPTION
	This cmdlet, New-SPListFromSchema creates a new SharePoint/Microsoft List from a CSV file that contains the schema definition.

	.PARAMETER siteUrl
	Specifies the destination SPWeb URL where the list is to be created.

	.PARAMETER listTitle
	Specifies the Display Name of the list to be created

	.PARAMETER listUrl
	Specifies the URL of the list to be created - this will be https://{tenant}.sharepoint.com/{listUrl}. Note: the listUrl does NOT include "Lists/", which is the standard for creating lists.

	.PARAMETER csvFilePath
	Specifies the path for the CSV schema definition file

	.INPUTS
	None. You cannot pipe objects to New-SPListFromSchema

	.OUTPUTS
	New-SPListFromSchema returne a PSObject of the output

	.EXAMPLE
	C:\PS> New-SPListFromSchema -siteUrl "https://tenant.sharepoint.com/teams/someSite" -listTitle "My New List" -listUrl "lists/NewListUrl" -csvFilePath .\ListDef-Schema.csv

	.LINK
	https://www.sharepointdiary.com/2015/01/create-new-custom-list-in-sharepoint-using-powershell.html
	https://www.sharepointdiary.com/2015/12/how-to-add-column-to-sharepoint-list-using-powershell.html


	.LINK
	Set-Item

	.NOTES
AUTHOR: Daniel Tshin, Shyam Minda-Shankar
LASTEDIT: $(Get-Date)
# Written by: Daniel Tshin (daniel.tshin@undp.org) and Shyam Minda-Shankar (shyam.minda-shankar@undp.org) on 11 November 2022
#
# version: 1.1
#
# Revision history
# # 11 Nov 2022 Creation
# # 13 Nov 2022 Minor updates for readability
# # 03 Feb 2023 adding comments for readability

#>

#region command line parameters ###############################################
# Input Parameters
# $args: URL of the web hosting the list or library, the name of the list or library to inspect. Separated by space

param (
	[Parameter(Mandatory = $true)][string]$siteUrl = $(Read-Host -Prompt "Please specify the URL of the site where you want to create this list"),
	[Parameter(Mandatory = $true)][string]$listTitle = $(Read-Host -Prompt "Please give the Display Name of the list you wish to create"),
	[Parameter(Mandatory = $true)][string]$listUrl = $(Read-Host -Prompt "Please specify the URL part of the list - there should be NO SPACES"),
	[Parameter(Mandatory = $true)][string]$csvFilePath = $(Read-Host -Prompt "Please specify the CSV file of the source list definition")
)

#endregion ####################################################################

$csvfile = Import-CSV $csvFilePath -Delimiter ","
$counter = 1
$connection = Connect-PnPOnline -Url $siteUrl -Interactive
$site = Get-PnPSite
$list = New-PnPList -Title $listTitle -Url $listUrl -Template GenericList

try {
	foreach ($filerow in $csvfile) {
		switch ($filerow.Type) {
			# This catches when "Choice" is selected with "MultipleSelections" specified (or not) - Technically, Type:Multichoice should be the one used, but the table can be ambiguous
			"Choice" {
				$choices = $filerow.Choices.split(",")
				if ($filerow.MultipleSelections -eq "true")
				{
					$addfield = Add-PnPField -List $list -Type "MultiChoice" -DisplayName $filerow.DisplayFieldName -InternalName $filerow.InternalFieldName -Choices $choices -AddToDefaultView
				}
				else {
					$addfield = Add-PnPField -List $list -Type $filerow.Type -DisplayName $filerow.DisplayFieldName -InternalName $filerow.InternalFieldName -Choices $choices -AddToDefaultView
				}
				Break;
			}
			# This catches when "Multichoice" is specified
			"Multichoice" {
				$addfield = Add-PnPField -List $list -Type "MultiChoice" -DisplayName $filerow.DisplayFieldName -InternalName $filerow.InternalFieldName -Choices $choices -AddToDefaultView
			}
			"User" {
				$addfield = Add-PnPField -List $list -Type $filerow.Type -DisplayName $filerow.DisplayFieldName -InternalName $filerow.InternalFieldName -AddToDefaultView
				if ($filerow.MultipleSelections -eq "true")
				{
					$tempField = Get-PnPField -List $list -Identity $filerow.DisplayFieldName
					[XML]$SchemaXml = $tempField.SchemaXml
					$SchemaXml.field.SetAttribute("Mult","TRUE")
					$OuterXML = $SchemaXml.OuterXml.Replace('Field Type="User"','Field Type="UserMulti"')
					Set-PnPField -List $list -Identity $addfield.Id -Values @{SchemaXml=$OuterXML} #multiple user

					#Set-PnPField -List $list -Identity $addfield.Id -Values @{"SelectionMode"=0} #set field to people only (no groups)
				}
			}
			"Lookup" {
				# see https://sharepoint.stackexchange.com/questions/255727/configure-lookup-field-using-sharepoint-pnp-powershell
			}
			"Invalid" {
				# Not used, so do not create a field/column
			}
			# All other cases (Integer, Text, Note, DateTime, Counter, Boolean, Number, Currency, URL, Computed, Threading, Guid, GridChoice, Calculated, File, Attachments, Recurrence, CrossProjectLink, ModStat, Error, ContentTypeId, PageSeparator, ThreadIndex, WorkflowStatus, AllDayEvent, WorkflowEventType, Geolocation, OutcomeChoice, Location, Thumbnail, MaxItems) are handled by Default
			Default {
				$addfield = Add-PnPField -List $list -Type $filerow.Type -DisplayName $filerow.DisplayFieldName -InternalName $filerow.InternalFieldName -AddToDefaultView
			}
		}
		$getfield = Get-PnPField -List $list -Identity $filerow.InternalFieldName
		$getfield.Description = $filerow.Description
		$getfield.Update()
		Write-Host "Added list field/column:" $filerow.InternalFieldName "from CSV row" $counter
		$counter++
	}
}
catch {
	Write-host -ForegroundColor "red" -BackgroundColor "black" "Error in the function"
	Write-Host $_.Exception.Message $filerow.InternalFieldName "- at row: " $counter
}
finally {
}
