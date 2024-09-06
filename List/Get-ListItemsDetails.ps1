<#
	.SYNOPSIS
		Get-ListItemsDetails returns details of list items for a given SharePoint list.

	.DESCRIPTION
		Get-ListItemsDetails outputs the details of list items in a specific SharePoint list including Attachments details

	.PARAMETER siteURL
		Specify the Url of the SharePoint site where the list lives
	.PARAMETER listName
		Specify the Display name of the List. Use the DisplayName, not the internal listname

	.INPUTS
		System.String System.String

	.OUTPUTS
		System.PSObject

	.EXAMPLE
		Get-ListItemsDetails -siteUrl SiteURL -listName "Display Name Of List"


	.LINK
		The name of a related topic. The value appears on the line below the ".LINK" keyword and must be preceded by a comment symbol # or included in the comment block.
		http://www.fabrikam.com/Add-Extension.html

	.LINK
		Get-Item

	.NOTES
AUTHOR: Daniel Tshin, Salaudeen Rajack
LASTEDIT: $(Get-Date)
# Written by: Daniel Tshin (daniel.tshin@undp.org) on 16 August 2024
# Inspired by https://www.sharepointdiary.com/2019/05/sharepoint-online-get-attachments-from-list-using-powershell.html and https://www.sharepointdiary.com/2017/01/sharepoint-online-download-attachments-from-list-using-powershell.html by Salaudeen Rajack
# version: 1.0
#
# Revision history
# # 16 Aug 2024 Initial creation

#>


#region command line parameters ###############################################
# Input Parameters
# $args: URL of the web hosting the list or library, the name of the list or library to inspect. Separated by space
param(
    [Parameter(Mandatory=$true)][string]$siteURL = $(throw "Please specify the URL for the source site for the list"),
    [Parameter(Mandatory=$true)][string]$listName = $(throw "Please specify the Display Name of the list")
	);
#endregion ####################################################################

#Connect to SharePoint Online
Connect-PnPOnline -Url $siteURL -Interactive

#Get the Lists
#$List = Get-PnPList -Identity $listName

$Resultset = @()
try
{
    #Get All List Items with Attachments
    $SPQuery = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='Attachments' /><Value Type='Boolean'>1</Value></Eq></Where></Query></View>"
    $ListItems = Get-PnPListItem -List $listName -PageSize 200 -Query $SPQuery

    #Loop through each item in the list
    ForEach ($ListItem in $ListItems)
    {
        #Get All Attachments of the List Item
        $Attachments = Get-PnPProperty -ClientObject $ListItem -Property "AttachmentFiles"
		$AttachmentResultset = @()

		# If there are attachments, then collect them and get the info.
		if ($Attachments.Count -gt 0) {
			#Collect Attachment Properties
			ForEach ($Attachment in $Attachments)
			{
				$AttachmentResultset += New-Object PSObject -Property ([ordered]@{
					AttachmentName = $Attachment.FileName
					ServerRelativeUrl = $Attachment.ServerRelativeUrl
					CreatedBy = $ListItem.FieldValues.Author.LookupValue
				})
			}
		}


		$Resultset += New-Object PSObject -Property ([ordered]@{
			ItemId = $ListItem.Id
			ItemTitle = $ListItem.FieldValues.Title
			TotalAttachments = $Attachments.Count
			CreatedBy = $ListItem.FieldValues.Author.LookupValue
			CreatedDate = $ListItem.FieldValues.Created_x0020_Date
			ModifiedBy = $listItem.FieldValues.Editor.LookupValue
			ModifiedDate = $ListItem.FieldValues.Last_x0020_Modified
			AttachmentFileNames = $AttachmentResultset
		})

		$counter++
		Write-Progress -Activity "Getting List Attachment inventory" -Status "Processed: $counter of $($ListItems.Count)" -PercentComplete (($counter / $ListItems.Count) * 100)
    }
}
catch
{
	Write-Host "Unable to generate inventory of attachments in $listName. " $_.Exception.Message -ForegroundColor Red
}

$Resultset
