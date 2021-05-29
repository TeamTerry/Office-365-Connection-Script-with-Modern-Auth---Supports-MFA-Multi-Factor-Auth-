###################################################################################################################
###                                                                                                             ###
###  	Script by Terry Munro -                                                                                 ###
###     Technical Blog -               http://365admin.com.au                                                   ###
###     Webpage -                      https://www.linkedin.com/in/terry-munro/                                 ###
###     TechNet Gallery Scripts -      http://tinyurl.com/TerryMunroTechNet                                     ###
###                                                                                                             ###
###     TechNet Download link -        https://gallery.technet.microsoft.com/Office-365-Connection-47e03052     ###
###                                                                                                             ###
###     Version 1.1 - 20/04/2017                                                                                ### 
###     Version 1.2 - 28/04/2017 - Added Skype For Business MFA                                                 ###
###     Version 1.3 - 01/07/2017 - Added variable for tenant name and UPN                                       ###
###     Version 2.0 - 22/07/2017 - Major upgrade with Windows Form GUI                                          ###
###     Version 2.5 - 02/10/2017 - Added Compliance Center and edit to allow window to remain open for use      ###
###     Version 3.0 - 11/07/2020 - Major upgrade - include Exchange Online v2 - Teams - Install Modules         ###
###                                                                                                             ###
###################################################################################################################


####  Notes for Usage  #####################################################################################################################
#                                                                                                                                          #
#  Ensure you update the variable script with your tenant name                                                                             #
#  The tenant name is used in the SharePoint Online section for SharePoint connection URL                                                  # 
#                                                                                                                                          #
#  Special thanks to Bozford for notification on Compliance Center support and help in improving the script                                #
#                                                                                                                                          #
#  Thanks to Scine for the Exchange Online component -                                                                                     #
#  https://github.com/Scine/Powershell/blob/master/Connect%20To%20Powershell%20with%20or%20without%202%20form%20factor%20auth%20enabled    #
#                                                                                                                                          #
#  Thanks to Steven Winston-Brown for guidance on getting Skype for Business PowerShell MFA working                                        #
#  - - https://www.linkedin.com/in/steve-winston-brown/                                                                                    #
#                                                                                                                                          #
#  Support Guides -   http://www.365admin.com.au/2017/07/all-mfa-multi-factor-authentication.html                                          #
#   - Pre-Requisites -                                                                                                                     #
#                                                                                                                                          #
#   - - Configure your PC for Office 365 Admin inculding MFA -                                                                             #
#   - - - http://www.365admin.com.au/2017/01/how-to-configure-your-desktop-pc-for.html                                                     #
#                                                                                                                                          #
#   - - How to enable MFA (Multi-Factor Authentication) for Office 365 administrators                                                      #
#   - - - http://www.365admin.com.au/2017/07/how-to-enable-mfa-multi-factor.html                                                           #
#                                                                                                                                          #
#   - - How to connect to Office 365 via PowerShell with MFA - Multi-Factor Authentication                                                 #
#   - - -http://www.365admin.com.au/2017/07/how-to-connect-to-office-365-via.html                                                          # 
#                                                                                                                                          #
#                                                                                                                                          #
############################################################################################################################################

$Tenant = "<tenant>"
$UPN = "<user>@<tenant>.onmicrosoft.com"

#----------------------------------------------
# Generated Form Function
#----------------------------------------------
function Show-ConnectWithModernAuth-v3-0-Final-Keep-Open_psf {

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$formConnectToOffice365Us = New-Object 'System.Windows.Forms.Form'
	$textbox2 = New-Object 'System.Windows.Forms.TextBox'
	$buttonUpdateModules = New-Object 'System.Windows.Forms.Button'
	$buttonTechnicalBlog = New-Object 'System.Windows.Forms.Button'
	$buttonInstallModules = New-Object 'System.Windows.Forms.Button'
	$textbox1 = New-Object 'System.Windows.Forms.TextBox'
	$buttonSupportURLs = New-Object 'System.Windows.Forms.Button'
	$buttonConnectToAzureInform = New-Object 'System.Windows.Forms.Button'
	$buttonConnectToAzureADV2 = New-Object 'System.Windows.Forms.Button'
	$buttonConnectToAzureADV1 = New-Object 'System.Windows.Forms.Button'
	$buttonConnectTeams = New-Object 'System.Windows.Forms.Button'
	$buttonConnectToSharePointO = New-Object 'System.Windows.Forms.Button'
	$buttonConnectToExchangeOnl = New-Object 'System.Windows.Forms.Button'
	$buttonOK = New-Object 'System.Windows.Forms.Button'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	$formConnectToOffice365Us_Load={
		#TODO: Initialize Form Controls here
		
	}
	
	$buttonConnectToExchangeOnl_Click={
		#TODO: Place custom script here
		
		Write-Host "Running the script to Connect to Exchange Online Version 1"
		
		Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName | ?{ $_ -notmatch "_none_" } | select -First 1)
		$EXOSession = New-ExoPSSession -UserPrincipalName $UPN
		Import-PSSession $EXOSession -AllowClobber
		
		Write-Host "Completed running the script to Connect to Exchange Online V1 - Run the cmdlet - Get-Mailbox - to test connection"
		
	}
	
	$buttonConnectToSharePointO_Click={
		#TODO: Place custom script here
		
		Write-Host "Running the script to Connect to SharePoint Online"
		
		Connect-SPOService -Url https://$($Tenant)-admin.sharepoint.com
		
		Write-Host "Completed running the script to Connect to SharePoint Online - Run the cmdlet - Get-SPOTenant - to test connection"
		
	}
	
	$buttonConnectTeams_Click={
		#TODO: Place custom script here
		
		Write-Host "Running the script to Connect to Microsoft Teams"
		
		Import-Module MicrosoftTeams
		Connect-MicrosoftTeams
		
		Write-Host "Completed running the script to Connect to Microsoft Teams - Run the cmdlet - Get-TeamsApp - to test connection"
		
	}
	
	$buttonConnectToAzureADV1_Click={
		#TODO: Place custom script here
		
		Write-Host "Running the script to Connect to Azure Active Directory v1"
		
		Connect-MsolService
		
		Write-Host "Completed running the script to Azure Active Directory v1 - Run the cmdlet - Get-MSOLUser - to test connection"
		
	}
	
	$buttonConnectToAzureADV2_Click={
		#TODO: Place custom script here
		
		Write-Host "Running the script to Connect to Azure Active Directory v2"
		
		Connect-AzureAD
		
		Write-Host "Completed running the script to Azure Active Directory v2 - Run the cmdlet - Get-AzureADUser - to test connection"
		
	}
	
	
	$buttonConnectToAzureInform_Click={
		#TODO: Place custom script here
		
		Write-Host "Running the script to Connect to Exchange Online V2"
		
		Connect-ExchangeOnline
		
		Write-Host "Completed running the script to connect to Exchange Online V2 - Run the cmdlet - Get-EXOMailbox - to test connection"
		
	}
	
	$buttonInstallModules_Click={
		#TODO: Place custom script here
		
		Write-Host "Installing the Azure AD v1 Module"
		Install-Module MSOnline
		Write-Host "Finished installing the Azure AD v1 Module"
		
		Write-Host "Installing the Azure AD v2 Module"
		Install-Module AzureAD
		Write-Host "Finished installing the Azure AD v2 Module"
		
		Write-Host "Installing the SharePoint Online Module"
		Install-Module Microsoft.Online.SharePoint.PowerShell
		Write-Host "Finished installing the SharePoint Online Module"
		
		Write-Host "Installing the Microsoft Teams Module"
		Install-Module MicrosoftTeams
		Write-Host "Finished installing the Microsoft Teams Module"
		
		Write-Host "Installing the Exchange Online v2 Module"
		Install-Module ExchangeOnlineManagement
		Write-Host "Finished installing the Exchange Online v2 Module"
		
		Write-Host "Installing the Exchange Online v1 Module"
		start microsoft-edge:https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application
		Write-Host "Finished installing the Exchange Online v1 Module"
		
		Write-Host "Completed installing Office 365 Modules"
	}
	
	$buttonTechnicalBlog_Click={
		#TODO: Place custom script here
		
		Start-Process -FilePath http://365admin.com.au
		
	}
	
	$buttonSupportURLs_Click={
		#TODO: Place custom script here
		
		Start-Process -FilePath http://www.365admin.com.au/2017/07/all-mfa-multi-factor-authentication.html
		
	}
	
	$buttonUpdateModules_Click={
		#TODO: Place custom script here
		
		Write-Host "Start updating the Azure AD v1 Module"
		Update-Module MSOnline
		Write-Host "Completed updating the Azure AD v1 module"
		
		Write-Host "Start updating the Azure AD v2 Module"
		Update-Module AzureAD
		Write-Host "Completed updating the Azure AD v2 module"
		
		Write-Host "Start updating the SharePoint Online Module"
		Update-Module Microsoft.Online.SharePoint.PowerShell
		Write-Host "Completed updating the SharePoint Online module"
		
		Write-Host "Start updating the Microsoft Teams Module"
		Update-Module MicrosoftTeams
		Write-Host "Completed updating the Microsoft Teams module"
		
		Write-Host "Start Updating Exchange Online V1 module"
		start microsoft-edge:https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application
		Write-Host "Completed Updating Exchange Online V1 module"
		
		Write-Host "Start Updating Exchange Online V2 module"
		Import-Module ExchangeOnlineManagement; Get-Module ExchangeOnlineManagement
		Update-Module -Name ExchangeOnlineManagement
		Write-Host "Completed Updating Exchange Online V2 module"
		
		Write-Host "Completed updating Office 365 PowerShell Modules"
	}
	
	$textbox1_TextChanged={
		#TODO: Place custom script here
		
	}
	
	$textbox2_TextChanged={
		#TODO: Place custom script here
		
	}
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$formConnectToOffice365Us.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$textbox2.remove_TextChanged($textbox2_TextChanged)
			$buttonUpdateModules.remove_Click($buttonUpdateModules_Click)
			$buttonTechnicalBlog.remove_Click($buttonTechnicalBlog_Click)
			$buttonInstallModules.remove_Click($buttonInstallModules_Click)
			$textbox1.remove_TextChanged($textbox1_TextChanged)
			$buttonSupportURLs.remove_Click($buttonSupportURLs_Click)
			$buttonConnectToAzureInform.remove_Click($buttonConnectToAzureInform_Click)
			$buttonConnectToAzureADV2.remove_Click($buttonConnectToAzureADV2_Click)
			$buttonConnectToAzureADV1.remove_Click($buttonConnectToAzureADV1_Click)
			$buttonConnectTeams.remove_Click($buttonConnectTeams_Click)
			$buttonConnectToSharePointO.remove_Click($buttonConnectToSharePointO_Click)
			$buttonConnectToExchangeOnl.remove_Click($buttonConnectToExchangeOnl_Click)
			$formConnectToOffice365Us.remove_Load($formConnectToOffice365Us_Load)
			$formConnectToOffice365Us.remove_Load($Form_StateCorrection_Load)
			$formConnectToOffice365Us.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$formConnectToOffice365Us.SuspendLayout()
	#
	# formConnectToOffice365Us
	#
	$formConnectToOffice365Us.Controls.Add($textbox2)
	$formConnectToOffice365Us.Controls.Add($buttonUpdateModules)
	$formConnectToOffice365Us.Controls.Add($buttonTechnicalBlog)
	$formConnectToOffice365Us.Controls.Add($buttonInstallModules)
	$formConnectToOffice365Us.Controls.Add($textbox1)
	$formConnectToOffice365Us.Controls.Add($buttonSupportURLs)
	$formConnectToOffice365Us.Controls.Add($buttonConnectToAzureInform)
	$formConnectToOffice365Us.Controls.Add($buttonConnectToAzureADV2)
	$formConnectToOffice365Us.Controls.Add($buttonConnectToAzureADV1)
	$formConnectToOffice365Us.Controls.Add($buttonConnectTeams)
	$formConnectToOffice365Us.Controls.Add($buttonConnectToSharePointO)
	$formConnectToOffice365Us.Controls.Add($buttonConnectToExchangeOnl)
	$formConnectToOffice365Us.Controls.Add($buttonOK)
	$formConnectToOffice365Us.AcceptButton = $buttonOK
	$formConnectToOffice365Us.AutoScaleDimensions = '6, 13'
	$formConnectToOffice365Us.AutoScaleMode = 'Font'
	$formConnectToOffice365Us.BackColor = 'Window'
	$formConnectToOffice365Us.ClientSize = '783, 483'
	$formConnectToOffice365Us.FormBorderStyle = 'FixedDialog'
	$formConnectToOffice365Us.MaximizeBox = $False
	$formConnectToOffice365Us.MinimizeBox = $False
	$formConnectToOffice365Us.Name = 'formConnectToOffice365Us'
	$formConnectToOffice365Us.StartPosition = 'CenterScreen'
	$formConnectToOffice365Us.Text = 'Connect to Office 365 using Modern Auth v3.0 - By Terry Munro - 365admin.com.au'
	$formConnectToOffice365Us.add_Load($formConnectToOffice365Us_Load)
	#
	# textbox2
	#
	$textbox2.BackColor = 'Chartreuse'
	$textbox2.Font = 'Microsoft Sans Serif, 11pt'
	$textbox2.Location = '28, 372'
	$textbox2.Name = 'textbox2'
	$textbox2.Size = '722, 24'
	$textbox2.TabIndex = 15
	$textbox2.Text = 'After connecting to your services - 
Click the OK button 
to close this GUI and 
start your PowerShell session'
	$textbox2.TextAlign = 'Center'
	$textbox2.add_TextChanged($textbox2_TextChanged)
	#
	# buttonUpdateModules
	#
	$buttonUpdateModules.BackColor = 'Window'
	$buttonUpdateModules.Location = '589, 231'
	$buttonUpdateModules.Name = 'buttonUpdateModules'
	$buttonUpdateModules.Size = '130, 43'
	$buttonUpdateModules.TabIndex = 14
	$buttonUpdateModules.Text = 'Update Modules'
	$buttonUpdateModules.UseVisualStyleBackColor = $False
	$buttonUpdateModules.add_Click($buttonUpdateModules_Click)
	#
	# buttonTechnicalBlog
	#
	$buttonTechnicalBlog.BackColor = 'Window'
	$buttonTechnicalBlog.Location = '589, 302'
	$buttonTechnicalBlog.Name = 'buttonTechnicalBlog'
	$buttonTechnicalBlog.Size = '130, 43'
	$buttonTechnicalBlog.TabIndex = 12
	$buttonTechnicalBlog.Text = 'Technical Blog'
	$buttonTechnicalBlog.UseVisualStyleBackColor = $False
	$buttonTechnicalBlog.add_Click($buttonTechnicalBlog_Click)
	#
	# buttonInstallModules
	#
	$buttonInstallModules.BackColor = 'Window'
	$buttonInstallModules.Location = '589, 149'
	$buttonInstallModules.Name = 'buttonInstallModules'
	$buttonInstallModules.Size = '130, 43'
	$buttonInstallModules.TabIndex = 11
	$buttonInstallModules.Text = 'Install Modules'
	$buttonInstallModules.UseVisualStyleBackColor = $False
	$buttonInstallModules.add_Click($buttonInstallModules_Click)
	#
	# textbox1
	#
	$textbox1.BackColor = 'Window'
	$textbox1.Location = '560, 33'
	$textbox1.Name = 'textbox1'
	$textbox1.Size = '190, 20'
	$textbox1.TabIndex = 9
	$textbox1.Text = 'Support Links'
	$textbox1.TextAlign = 'Center'
	$textbox1.add_TextChanged($textbox1_TextChanged)
	#
	# buttonSupportURLs
	#
	$buttonSupportURLs.BackColor = 'Window'
	$buttonSupportURLs.Location = '589, 72'
	$buttonSupportURLs.Name = 'buttonSupportURLs'
	$buttonSupportURLs.Size = '130, 43'
	$buttonSupportURLs.TabIndex = 8
	$buttonSupportURLs.Text = 'Support URLs'
	$buttonSupportURLs.UseVisualStyleBackColor = $False
	$buttonSupportURLs.add_Click($buttonSupportURLs_Click)
	#
	# buttonConnectToAzureInform
	#
	$buttonConnectToAzureInform.BackColor = 'Control'
	$buttonConnectToAzureInform.Font = 'Microsoft Sans Serif, 11.25pt'
	$buttonConnectToAzureInform.Location = '287, 33'
	$buttonConnectToAzureInform.Name = 'buttonConnectToAzureInform'
	$buttonConnectToAzureInform.Size = '190, 82'
	$buttonConnectToAzureInform.TabIndex = 7
	$buttonConnectToAzureInform.Text = 'Connect to Exchange Online V2'
	$buttonConnectToAzureInform.UseVisualStyleBackColor = $False
	$buttonConnectToAzureInform.add_Click($buttonConnectToAzureInform_Click)
	#
	# buttonConnectToAzureADV2
	#
	$buttonConnectToAzureADV2.BackColor = 'Control'
	$buttonConnectToAzureADV2.Font = 'Microsoft Sans Serif, 11.25pt'
	$buttonConnectToAzureADV2.Location = '287, 149'
	$buttonConnectToAzureADV2.Name = 'buttonConnectToAzureADV2'
	$buttonConnectToAzureADV2.Size = '190, 82'
	$buttonConnectToAzureADV2.TabIndex = 5
	$buttonConnectToAzureADV2.Text = 'Connect to Azure AD v2'
	$buttonConnectToAzureADV2.UseVisualStyleBackColor = $False
	$buttonConnectToAzureADV2.add_Click($buttonConnectToAzureADV2_Click)
	#
	# buttonConnectToAzureADV1
	#
	$buttonConnectToAzureADV1.BackColor = 'Control'
	$buttonConnectToAzureADV1.Font = 'Microsoft Sans Serif, 11.25pt'
	$buttonConnectToAzureADV1.Location = '53, 149'
	$buttonConnectToAzureADV1.Name = 'buttonConnectToAzureADV1'
	$buttonConnectToAzureADV1.Size = '190, 82'
	$buttonConnectToAzureADV1.TabIndex = 4
	$buttonConnectToAzureADV1.Text = 'Connect to Azure AD v1'
	$buttonConnectToAzureADV1.UseVisualStyleBackColor = $False
	$buttonConnectToAzureADV1.add_Click($buttonConnectToAzureADV1_Click)
	#
	# buttonConnectTeams
	#
	$buttonConnectTeams.BackColor = 'Control'
	$buttonConnectTeams.Font = 'Microsoft Sans Serif, 11.25pt'
	$buttonConnectTeams.Location = '53, 263'
	$buttonConnectTeams.Name = 'buttonConnectTeams'
	$buttonConnectTeams.Size = '190, 82'
	$buttonConnectTeams.TabIndex = 3
	$buttonConnectTeams.Text = 'Connect to Microsoft Teams'
	$buttonConnectTeams.UseVisualStyleBackColor = $False
	$buttonConnectTeams.add_Click($buttonConnectTeams_Click)
	#
	# buttonConnectToSharePointO
	#
	$buttonConnectToSharePointO.BackColor = 'Control'
	$buttonConnectToSharePointO.Font = 'Microsoft Sans Serif, 11.25pt'
	$buttonConnectToSharePointO.Location = '287, 263'
	$buttonConnectToSharePointO.Name = 'buttonConnectToSharePointO'
	$buttonConnectToSharePointO.Size = '190, 82'
	$buttonConnectToSharePointO.TabIndex = 2
	$buttonConnectToSharePointO.Text = 'Connect to SharePoint Online'
	$buttonConnectToSharePointO.UseVisualStyleBackColor = $False
	$buttonConnectToSharePointO.add_Click($buttonConnectToSharePointO_Click)
	#
	# buttonConnectToExchangeOnl
	#
	$buttonConnectToExchangeOnl.BackColor = 'Control'
	$buttonConnectToExchangeOnl.Font = 'Microsoft Sans Serif, 11.25pt'
	$buttonConnectToExchangeOnl.Location = '53, 33'
	$buttonConnectToExchangeOnl.Name = 'buttonConnectToExchangeOnl'
	$buttonConnectToExchangeOnl.Size = '190, 82'
	$buttonConnectToExchangeOnl.TabIndex = 1
	$buttonConnectToExchangeOnl.Text = 'Connect to Exchange Online V1'
	$buttonConnectToExchangeOnl.UseVisualStyleBackColor = $False
	$buttonConnectToExchangeOnl.add_Click($buttonConnectToExchangeOnl_Click)
	#
	# buttonOK
	#
	$buttonOK.Anchor = 'Bottom, Right'
	$buttonOK.DialogResult = 'OK'
	$buttonOK.Location = '339, 417'
	$buttonOK.Name = 'buttonOK'
	$buttonOK.Size = '89, 35'
	$buttonOK.TabIndex = 0
	$buttonOK.Text = '&OK'
	$buttonOK.UseVisualStyleBackColor = $True
	$formConnectToOffice365Us.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $formConnectToOffice365Us.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$formConnectToOffice365Us.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$formConnectToOffice365Us.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $formConnectToOffice365Us.ShowDialog()

} #End Function

#Call the form
Show-ConnectWithModernAuth-v3-0-Final-Keep-Open_psf | Out-Null
