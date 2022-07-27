# Script:        Get SharePoint Online List Data
#
# Description:   Ran from Developer machine in PowerShell ISE, uses local credentials to access a SharePoint Online Site,
#                and retrieve data in a given List for export to .csv/.xlsx. Intention is for this data to be regularly pulled to refresh
#                a Power BI Report based upon said data.
#
# Last Modified: 7/6/2022
#--------------------------------------------------
# PRE-REQUISITE(S): 
#--------------------------------------------------
#    1. Have Access to the SharePoint Site List in question
#    2. Install-Module -Name PnP.PowerShell
#--------------------------------------------------
# DEVELOPER NOTE(S):
#--------------------------------------------------
#    PnP PowerShell Documentation: https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html#interactive-login-for-multi-factor-authentication
#--------------------------------------------------
# CONFIGURATION PARAMETER(S):
#--------------------------------------------------
$siteURL = "https://{yourDomain}.sharepoint.com/sites/{nameOfSite}"
$listName = "{nameOfList}"
$exportFileLocation = "\\fileServer\ITGroup\DBTeam\SharePointExportedData\"
$CSVPath = $exportFileLocation + "Exported Data.csv"
#$ExcelPath = $exportFileLocation + "Exported Data.xlsx"

# EDIT on Line 62 as needed for fields to be renamed upon Export.
$propertiesToExport = @("ID","Title","field_1","field_2","field_3"); 
#--------------------------------------------------
# BEGIN SCRIPT
#--------------------------------------------------
try 
{
	# Check if Module is installed on current machine
    if (Get-Module -ListAvailable -Name PnP.PowerShell) {
        $dataForExport = @()
        cls
    
        # Connect-PnPOnline: https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html
        # INFO - The following prompts for credentials, but Fails with either 403 Forbidden or non-matching name/password in Microsoft account system. Will not work if Multi-Factor Authentication (MFA) is enabled.
        #    Connect-PnPOnline -Url $SiteURL -Credentials (Get-Credential) # 
        # INFO - Non-Admin: The following returns "The sign-in name or password does not match one in the Microsoft account system - assuming also not working due to Multi-Factor Authentication.
        #        Admin: AADSTS53000: Device is not in required device state: compliant. Conditional Access policy requires a compliant device, and the device is not compliant. The user must enroll their device with an approved MDM provider like Intune.
        #    $creds = (New-Object System.Management.Automation.PSCredential "<<UserName>>",(ConvertTo-SecureString "<<Password>>" -AsPlainText -Force))
        #    Connect-PnPOnline -Url $siteURL -Credentials $creds
        # Ex. 9 - Creates a brief popup window that checks for login cookies. Can use -ForceAuthentication to reset and force a new login.
        #         NOTE: This is required for systems utilizing Multi-Factor Authentication.
        Connect-PnPOnline -Url $siteURL -UseWebLogin 

        $counter = 0
   
        #$ListItems = Get-PnPListItem -List $ListName -Fields $SelectedFields -PageSize 2000 # PageSize: The number of items to retrieve per page request
        #$ListItems = (Get-PnPListItem -List $ListName -Fields $SelectedFields).FieldValues
        $listItemIDTitleGUID = Get-PnPListItem -List $listName #  = Id, Title, GUID

        foreach ($item in $listItemIDTitleGUID)
        {
            $counter++
            Write-Progress -PercentComplete ($counter / $($listItemIDTitleGUID.Count)  * 100) -Activity "Exporting List Items..." -Status  "Exporting Item $counter of $($listItemIDTitleGUID.Count)"

            # Get all list items specificed properties by ID using PnP cmdlet - https://pnp.github.io/powershell/cmdlets/Get-PnPListItem.html
            $listItems = (Get-PnPListItem -List $listName -Id $item.Id -Fields $propertiesToExport).FieldValues
    
            # Add the Name / Value pairing of each property for export to the export data array. Adding and removing some properties by name to specify the actual field designation upon Export.
            $listItems | foreach {
                $itemForExport = New-Object PSObject
                foreach ($field in $propertiesToExport) {
                    if ($field -eq "field_1") {
                        $itemForExport | Add-Member -MemberType NoteProperty -name "field_1" -value $_[$field]
                        $itemForExport | Add-Member -MemberType NoteProperty -Name "ProjectSponsor" -Value $itemForExport.field_1 -Force
                        $itemForExport.PSObject.Properties.Remove("field_1")
                    } else {
                        if ($field -eq "field_2") {
                            $itemForExport | Add-Member -MemberType NoteProperty -name "field_2" -value $_[$field]
                            $itemForExport | Add-Member -MemberType NoteProperty -Name "InitiativeKey" -Value $itemForExport.field_2 -Force
                            $itemForExport.PSObject.Properties.Remove("field_2")
                        } else {
                            if ($field -eq "field_3") {
                                $itemForExport | Add-Member -MemberType NoteProperty -name "field_3" -value $_[$field]
                                $itemForExport | Add-Member -MemberType NoteProperty -Name "AreaOfFocus" -Value $itemForExport.field_3 -Force
                                $itemForExport.PSObject.Properties.Remove("field_3")
                            } else {
                                $itemForExport | Add-Member -MemberType NoteProperty -name $field -value $_[$field]

                            }
                        }
                    }
                }
		
              # Add the object with the above updated properties to the export Array
              $dataForExport += $itemForExport
          }
      }

      #Export the result Array to csv/xlsx files
      $dataForExport | Export-CSV $CSVPath -NoTypeInformation
			#$dataForExport | Export-Excel -Path $ExcelPath

        # Disconnect the current connection and clear token cache.
        Disconnect-PnPOnline
    } else {
        cls
        Write-Host "In order to execute this script, the Cmdlet Module {PnP.PowerShell} must be installed. Please open an Admin shell and run the following command: Install-Module -Name PnP.PowerShell"
    }
} catch {
    Write-Host $_.Exception.Message
    $Error.Clear()
}
#--------------------------------------------------
# END SCRIPT
#--------------------------------------------------
