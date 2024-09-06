<#
	.SYNOPSIS
		Get-ListAttachmentInventory returns details of all the attachments in a given SharePoint list.

	.DESCRIPTION
		Get-ListAttachmentInventory outputs the inventory of attachments in a SharePoint list including the AttachmentName and ServerRelativeUrl of the attachment.

	.PARAMETER siteURL
		Specify the Url of the SharePoint site where the list lives
	.PARAMETER listName
		Specify the Display name of the List. Use the DisplayName, not the internal listname

	.INPUTS
		System.String System.String

	.OUTPUTS
		System.PSObject

	.EXAMPLE
		Get-ListAttachmentInventory -siteUrl SiteURL -listName "Display Name Of List"


	.LINK
		The name of a related topic. The value appears on the line below the ".LINK" keyword and must be preceded by a comment symbol # or included in the comment block.
		http://www.fabrikam.com/Add-Extension.html

	.LINK
		Set-Item

	.NOTES
AUTHOR: Daniel Tshin, Salaudeen Rajack
LASTEDIT: $(Get-Date)
# Written by: Daniel Tshin (daniel.tshin@undp.org) on 15 August 2024
# Inspired by https://www.sharepointdiary.com/2019/05/sharepoint-online-get-attachments-from-list-using-powershell.html and https://www.sharepointdiary.com/2017/01/sharepoint-online-download-attachments-from-list-using-powershell.html by Salaudeen Rajack
# version: 1.0
#
# Revision history
# # 15 Aug 2024 Initial creation

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
    $ListItems = Get-PnPListItem -List $listName -PageSize 100 -Query $SPQuery

    #Loop through each item in the list
    ForEach ($ListItem in $ListItems)
    {
         #Get All Attachments of the List Item
        $Attachments = Get-PnPProperty -ClientObject $ListItem -Property "AttachmentFiles"

        #Collect Attachment Properties
        ForEach($Attachment in $Attachments)
        {
            $Resultset += New-Object PSObject -Property ([ordered]@{
                ItemID = $ListItem.ID
				ItemTitle = $ListItem.FieldValues.Title
				TotalAttachments = $Attachments.Count
				AttachmentName = $Attachment.FileName
                ServerRelativeUrl = $Attachment.ServerRelativeUrl
                CreatedBy = $ListItem.FieldValues.Author.LookupValue
            })

        }
		$counter++
		Write-Progress -Activity "Getting List Attachment inventory" -Status "Processed: $counter of $($ListItems.Count)" -PercentComplete (($counter / $ListItems.Count) * 100)
    }
}
catch
{
	Write-Host "Unable to generate inventory of attachments in $listName. " $_.Exception.Message -ForegroundColor Red
}

$Resultset
