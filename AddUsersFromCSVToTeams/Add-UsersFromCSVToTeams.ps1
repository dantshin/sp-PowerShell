<#
	.SYNOPSIS
		Add-UsersFromCSVToTeams grabs a list of users in an CSV file and adds them to a MS Team

	.DESCRIPTION
		Add-UsersFromCSVToTeams takes a CSV file and imports the users and adds them as members to a MS Team. This cmdlet uses the PowerShell Module for Microsoft Teams.

	.PARAMETER UserImportFile
		Filename of the CSV file to import. The CSV file should be formatted with

	.PARAMETER GroupID
		GUID of the MS Team. This can be obtained from the cmdlet Get-Team

	.EXAMPLE
		Add-UsersFromCSVToTeams -GroupID c7e2b692-2eb6-4dc3-a5e5-3a7a6a9c89b7 -UserImportFile C:\ImportFile.csv

	.EXAMPLE
		Add-UsersFromCSVToTeams

	.INPUTS
		groupID
		userImportFile

	.OUTPUTS
		PSObject

	.FUNCTIONALITY
		Add-UsersFromCSVToTeams is used to add/import users as members of a MS Team from a CSV file  This cmdlet uses the PowerShell Module for Microsoft Teams. Output can be processed or exported.

	.LINK
		about_functions_advanced

	.LINK
		https://www.jijitechnologies.com/blogs/create-teams-microsoft-teams-powershell

	.NOTES
		AUTHOR: Daniel Tshin
LASTEDIT: $(Get-Date)
# Written by: Daniel Tshin (daniel.tshin@undp.org) on 03 December 2018
#
# version: 1.0
#
# Revision history
# # 03 Dec 2018 Creation
# # 05 Dec 2018 Simplified function, removed group and channel creation
# # This script doesn't work properly - the external user doesn't get added to the AzureAD properly.
#>


#region command line parameters ###############################################
# Input Parameters
# $args: URL of the web hosting the list or library, the name of the list or library to inspect. Separated by space
param(
    [Parameter(Mandatory=$true)][string]$groupID = $(throw "Please specify the GroupID (GUID) for the MS Teams"),
	[Parameter(Mandatory=$true)][string]$userImportFile = $(throw "Please specify the CSV file containing the list of users to add to the MS Teams"),
	[Parameter(Mandatory=$true)][string]$redirectURL = $(throw "Please specify the URL to redirect the user when the invited user has accepted the invitation")
	);
#endregion ####################################################################

<#function Add-Users
{
    param(
			$userImportFile
		  )
#>
    Process
    {
		Import-Module MicrosoftTeams
		Import-Module AzureAD
        #$cred = Get-Credential
        #$username = $cred.UserName
		Connect-MicrosoftTeams #-Credential $cred
		Connect-AzureAD

        try {
			$getTeam = Get-Team | Where-Object {$_.GroupID -eq $groupID}
			if ($getTeam -ne $null)
			{
				$users = Import-Csv -Path $userImportFile
				foreach($user in $users)
				{
					$displayName = $user.Firstname + " " + $user.Lastname
					Write-Host "Adding User: " $displayName " Email: " $user.Email

					if ($user.Email -ne )
					# Investigate using New-AzureADUser -UserType Guest
					New-AzureADMSInvitation -InvitedUserDisplayName $displayName -InvitedUserEmailAddress $user.Email -SendInvitationMessage $true -InviteRedirectUrl $redirectURL
					Start-Sleep -Seconds 5
					Add-TeamUser -GroupId $groupID -User $user.Email
					Write-Host "Added User: " + $user.Email
				}
			}
		}
        Catch
            {
				Write-host -ForegroundColor "red" -BackgroundColor "black" "Error adding user: " $displayName ", Email: " $user.Email
				Write-Host $_.Exception.Message
			}
	}
<#}#>


#Create-NewTeam -ImportPath $userImportFile
#Add-Users -ImportPath $userImportFile