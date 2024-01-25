<#
	.SYNOPSIS
	Cmdlet to get permissions of all list items of a given SPO list/library

	.DESCRIPTION
	This PowerShell cmdlet, Get-LibraryPermissions gets all of the permissions groups and users in a SharePoint

	.PARAMETER siteUrl
	Specifies the Source SPWeb URL where the list/library lives.

	.PARAMETER listName
	Specifies the SPList Title. Use the DisplayName, not the internal listname

	.PARAMETER scanLevel
	Specifies the level of granularity in the scan. Values are: list, folder, item

	.PARAMETER includeInheritedPermissions
	Indicates if the script should include inherited permissions in the report

	.INPUTS
	None. You cannot pipe objects to Get-LibraryPermissions.

	.OUTPUTS
	Get-LibraryPermissions returns a PSObject of the output

	.EXAMPLE
	C:\PS> Get-LibraryPermissions -siteUrl https://undp.sharepoint.com/sites/SiteName -listName "Title of the List or Library"

	.EXAMPLE
	C:\PS> Get-LibraryPermissions -siteUrl https://undp.sharepoint.com/sites/SiteName -listName "Title of the List or Library" -scanLevel folder

	.LINK
	Get Document Library Permissions and Export to CSV using PnP PowerShell
	https://www.sharepointdiary.com/2019/02/sharepoint-online-pnp-powershell-to-export-document-library-permissions.html
	https://365basics.com/powershell-sharepoint-permission-report-for-all-lists-and-libraries-within-every-site-collection-and-subsite/

	.LINK
	Set-Item

	.NOTES
AUTHOR: Daniel Tshin
LASTEDIT: $(Get-Date)
# Written by: Daniel Tshin (daniel.tshin@undp.org) on 09 January 2023
#
# version: 1.0
#
# Revision history
# # 09 Jan 2023 Creation

#>


#region command line parameters ###############################################
# Input Parameters
# $args: URL of the web hosting the list or library, the name of the list or library to inspect. Separated by space
param(
	[Parameter(Mandatory = $true)][string]$webUrl = $(Read-Host -Prompt "Please specify the URL for the source site (SPWeb) for the list"),
	[Parameter(Mandatory = $true)][string]$listName = $(Read-Host -Prompt "Please specify the Title (Display Name) of the list"),
	[Parameter(Mandatory = $false)][string]$scanLevel = $(Read-Host -Prompt "Please indicate the level of granularity for this scan. Values: list, folder, item"),
	[Parameter(Mandatory = $false)][switch]$includeInheritedPermissions = $(Read-Host -Prompt "Please indicate if you want to include inherited permissions in the scan")
);
#endregion ####################################################################


Function Get-Permissions([Microsoft.SharePoint.Client.SecurableObject]$Object) {
	#Determine the type of the object
	Switch ($Object.TypedObject.ToString()) {
		"Microsoft.SharePoint.Client.ListItem" {
			If ($Object.FileSystemObjectType -eq "Folder") {
				$ObjectType = "Folder"
				#Get the URL of the Folder
				$Folder = Get-PnPProperty -ClientObject $Object -Property Folder
				$ObjectTitle = $Object.Folder.Name
				$ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $Object.Folder.ServerRelativeUrl)
			}
			Else { #File or List Item
				#Get the URL of the Object
				Get-PnPProperty -ClientObject $Object -Property File, ParentList
				If ($Object.File.Name -ne $Null) {
					$ObjectType = "File"
					$ObjectTitle = $Object.File.Name
					$ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $Object.File.ServerRelativeUrl)
				}
				else {
					$ObjectType = "List Item"
					$ObjectTitle = $Object["Title"]
					#Get the URL of the List Item
					$DefaultDisplayFormUrl = Get-PnPProperty -ClientObject $Object.ParentList -Property DefaultDisplayFormUrl
					$ObjectURL = $("{0}{1}?ID={2}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $DefaultDisplayFormUrl, $Object.ID)
				}
			}
		}
		Default {
			$ObjectType = "List or Library"
			$ObjectTitle = $Object.Title
			#Get the URL of the List or Library
			$RootFolder = Get-PnPProperty -ClientObject $Object -Property RootFolder
			$ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $RootFolder.ServerRelativeUrl)
		}
	}

	#Get permissions assigned to the object
	Get-PnPProperty -ClientObject $Object -Property HasUniqueRoleAssignments, RoleAssignments

	#Check if Object has unique permissions
	$HasUniquePermissions = $Object.HasUniqueRoleAssignments

	#Loop through each permission assigned and extract details
	$PermissionCollection = @()
	Foreach ($RoleAssignment in $Object.RoleAssignments) {
		#Get the Permission Levels assigned and Member
		Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member

		#Get the Principal Type: User, SP Group, AD Group
		$PermissionType = $RoleAssignment.Member.PrincipalType

		#Get the Permission Levels assigned
		$PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name

		#Remove Limited Access
		$PermissionLevels = ($PermissionLevels | Where-Object { $_ -ne "Limited Access" }) -join ","

		#Leave Principals with no Permissions
		If ($PermissionLevels.Length -eq 0) { Continue }

		#Get SharePoint group members
		If ($PermissionType -eq "SharePointGroup") {
			#Get Group Members
			$GroupMembers = Get-PnPGroupMember -Identity $RoleAssignment.Member.LoginName

			#Leave Empty Groups
			If ($GroupMembers.count -eq 0) { Continue }
			$GroupUsers = ($GroupMembers | Select-Object -ExpandProperty Title) -join "; "

			#Add the Data to Object
			$Permissions = New-Object PSObject
			$Permissions | Add-Member NoteProperty Object($ObjectType)
			$Permissions | Add-Member NoteProperty Title($ObjectTitle)
			$Permissions | Add-Member NoteProperty URL($ObjectURL)
			$Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
			$Permissions | Add-Member NoteProperty GrantedThrough("SharePoint Group: $($RoleAssignment.Member.LoginName)")
			$Permissions | Add-Member NoteProperty Users($GroupUsers)
			$Permissions | Add-Member NoteProperty Type($PermissionType)
			$Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
			$PermissionCollection += $Permissions
		}
		Else {
			#Add the Data to Object
			$Permissions = New-Object PSObject
			$Permissions | Add-Member NoteProperty Object($ObjectType)
			$Permissions | Add-Member NoteProperty Title($ObjectTitle)
			$Permissions | Add-Member NoteProperty URL($ObjectURL)
			$Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
			$Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
			$Permissions | Add-Member NoteProperty Users($RoleAssignment.Member.Title)
			$Permissions | Add-Member NoteProperty Type($PermissionType)
			$Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
			$PermissionCollection += $Permissions
		}
	}
	#Export Permissions to CSV File
	$PermissionCollection #| Export-CSV $ReportFile -NoTypeInformation -Append
}

#Function to get sharepoint online list permissions report
Try {
	#Function to Get Permissions of All List Folders of a given List
	Function Get-ListFoldersPermission([Microsoft.SharePoint.Client.List]$List) {
		Write-host -f Yellow "`t `t Getting Permissions of List Folders in the List:"$List.Title

		#Get All Items from List in batches
		$Folders = Get-PnPListItem -List $List -PageSize 500 | Where-Object {$_.FileSystemObjectType -eq "Folder"}

		$ItemCounter = 0
		#Loop through each List item
		ForEach ($Folder in $Folders) {
			#Get Objects with Unique Permissions or Inherited Permissions based on 'includeInheritedPermissions' switch
			If ($includeInheritedPermissions) {
				Get-Permissions -Object $Folder
			}
			Else {
				#Check if List Item has unique permissions
				$HasUniquePermissions = Get-PnPProperty -ClientObject $Folder -Property HasUniqueRoleAssignments
				If ($HasUniquePermissions -eq $True) {
					#Call the function to generate Permission report
					Get-Permissions -Object $Folder
				}
			}
			$ItemCounter++
			Write-Progress -PercentComplete ($ItemCounter / ($Folders.Count) * 100) -Activity "Processing Folders $ItemCounter of $($Folders.Count)" -Status "Searching Unique Permissions in List Folders of '$($List.Title)'"
		}
	}

	#Function to Get Permissions of All List Items of a given List
	Function Get-ListItemsPermission([Microsoft.SharePoint.Client.List]$List) {
		Write-host -f Yellow "`t `t Getting Permissions of List Items in the List:"$List.Title

		#Get All Items from List in batches
		$ListItems = Get-PnPListItem -List $List -PageSize 500

		$ItemCounter = 0
		#Loop through each List item
		ForEach ($ListItem in $ListItems) {
			#Get Objects with Unique Permissions or Inherited Permissions based on 'includeInheritedPermissions' switch
			If ($includeInheritedPermissions) {
				Get-Permissions -Object $ListItem
			}
			Else {
				#Check if List Item has unique permissions
				$HasUniquePermissions = Get-PnPProperty -ClientObject $ListItem -Property HasUniqueRoleAssignments
				If ($HasUniquePermissions -eq $True) {
					#Call the function to generate Permission report
					Get-Permissions -Object $ListItem
				}
			}
			$ItemCounter++
			Write-Progress -PercentComplete ($ItemCounter / ($List.ItemCount) * 100) -Activity "Processing Items $ItemCounter of $($List.ItemCount)" -Status "Searching Unique Permissions in List Items of '$($List.Title)'"
		}
	}

#---------------------------------------------------------------------------------------
### Start of script ###

	# Connect to SPO Site Collection
	Connect-PnPOnline -Url $webUrl -Interactive
	# End SPO

	# Get the SPWeb object
	$Web = Get-PnPWeb

	#Get the List
	$List = Get-PnpList -Identity $ListName -Includes RoleAssignments

	Write-host -f Yellow "Getting Permissions of the List '$ListName'..."
	#Get List Permissions
	Get-Permissions -Object $List

	#Get Folder or Item Level Permissions if 'scanLevel' switch is specified
	switch ($scanLevel) {
		"folder" {
			Get-ListFoldersPermission -List $List
		}

		"item" {
			#Get List Items Permissions
			Get-ListItemsPermission -List $List
		}


	}

	#If ($scanLevel -eq "list")
	Write-host -f Green "`t List Permission Report Generated Successfully!"
}
Catch {
	write-host -f Red "Error Generating List Permission Report!" $_.Exception.Message
}