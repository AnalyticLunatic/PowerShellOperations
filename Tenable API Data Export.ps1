# Script:             Tenable Data Export
#
# Background:		  Tenable is a software service in my organization which scans all network connected Assets (laptops, desktops, servers, etc.)
#                     to see if any known or reported Vulnerabilities have been found require patching/updates.
#
# Description:        Allows the authenticated person to call Tenable API and trigger a fresh export of latest Tenable scanned Assets 
#                     and identified Vulnerabilites. These exported datasets are then saved to a local file for Power BI to process. 
#                     First all Assets data is exported, followed by all Non-Info (Low, Medium, High, Critial) Vulnerabilities data, 
#                     and then {Info}rmational Vulnerabilities.
#
#					  Note: Each new Export of data gets its own unique {uuid}, which can be used in API to check status of export and retrieve exported data from.
#
# Documentation:      https://developer.tenable.com/reference/exports-assets-request-export
#                     https://developer.tenable.com/reference/exports-vulns-request-export
#                     - There are many Filters available to specify what to return.
#--------------------------------------------------------------------------------------------------------------------
# Instructions:
#--------------------------------------------------------------------------------------------------------------------
# 1. Go to your Tenable portal and create a Secret Key for the API to utilize. This value should resemble 'accessKey=xxxxx;secretKey=yyyyyy'.
#
# 2. Enter the newly created API key into the $apiKey variable below. 
#
# 3. (Optional) Adjust the data chunk size for how many Assets/Vulnerabilities records to include in each data chunk.
#               Assets Per Chunk Range:          100 - 10,000
#               Vulnerabilities per Chunk Range:  50 - 5,000
#
# 4. (Optional) Adjust API body parameter filters for what data should be returned. 
#               Note: The customBody variables for Vulnerabilites include the chunk size "num_assets" as part of the body.
#
# 5. Edit the $exportFileLocation variable to where the data should be output to.
#--------------------------------------------------------------------------------------------------------------------
# Setup Variables:
#--------------------------------------------------------------------------------------------------------------------
$apiKey = 'accessKey=xxxxx;secretKey=yyyyy'

$assetsChunkSize = '{"chunk_size": 10000}' 
#$vulnsChunkSize = '{"chunk_size": 5000}'   

# Handling {Informational} vulnerabilities separately for Data Modeling purposes.
$customBodyNonInfo = '{"filters":{"severity":["low","medium","high","critical"]},"num_assets":5000}'

$customBodyInfoOnly = '{"filters":{"severity":["info"]},"num_assets":5000}'  

$exportFileLocation = 'C:\{YourUser}\Downloads\Data\Tenable\'

$headers=@{}
$headers.Add("Accept", "application/json")
$headers.Add("Content-Type", "application/json")
$headers.Add("X-ApiKeys", "$($apiKey)")
#--------------------------------------------------------------------------------------------------------------------
# Section 1 - Export Assets:
#--------------------------------------------------------------------------------------------------------------------
Clear-Host
Start-Transcript -OutputDirectory "$($exportFileLocation)TenableExportLogs"

Write-Host "1.  $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Triggering Tenable export of Assets."

$response = Invoke-WebRequest -Uri 'https://cloud.tenable.com/assets/export' -Method POST -Headers $headers -ContentType 'application/json' -Body $assetsChunkSize -UseBasicParsing

$temp = $response.Content | ConvertFrom-Json
$assetExportuuid = $temp.export_uuid

Write-Host "2.  $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Checking status of Tenable Assets export {uuid: $($assetExportuuid)}. Loop checks again every 10sec."
Do {
    Start-Sleep -s 10
    $assetStatusCheck = Invoke-WebRequest -Uri "https://cloud.tenable.com/assets/export/$($assetExportuuid)/status" -Method GET -Headers $headers -UseBasicParsing
        #$assetStatusCheck.RawContent - Ex. Note: {"status":"FINISHED","chunks_available":[1,2,3,4,5,6,7,8,9,10,11,12]}
    $tempAssetStatus = $assetStatusCheck.Content | ConvertFrom-Json
    Write-Host "    $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - $($tempAssetStatus.status)"
} Until ($tempAssetStatus.status -eq "FINISHED")

Write-Host "3.  $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Downloading data chunk {1} of {$($tempAssetStatus.chunks_available)} for Tenable Assets export {uuid: $($assetExportuuid)}."
$assetsData = Invoke-RestMethod -Uri "https://cloud.tenable.com/assets/export/$($assetExportuuid)/chunks/1" -Method GET -Headers $headers -UseBasicParsing

# Remove properties we do not need in Power BI data model from being written to file.
#foreach ($obj in $assetsData) {
#    $obj.PSObject.Properties.Remove('id')
#    $obj.PSObject.Properties.Remove('has_agent')
#    $obj.PSObject.Properties.Remove('has_plugin_results')
#    $obj.PSObject.Properties.Remove('created_at')
#    $obj.PSObject.Properties.Remove('terminated_at')
#    $obj.PSObject.Properties.Remove('terminated_by')
#    $obj.PSObject.Properties.Remove('updated_at')
#    $obj.PSObject.Properties.Remove('deleted_at')
#    $obj.PSObject.Properties.Remove('deleted_by')
#    $obj.PSObject.Properties.Remove('first_seen')
#    $obj.PSObject.Properties.Remove('last_seen')
#    $obj.PSObject.Properties.Remove('first_scan_time')
#    $obj.PSObject.Properties.Remove('last_scan_time')
#    $obj.PSObject.Properties.Remove('last_authenticated_scan_date')
#    $obj.PSObject.Properties.Remove('last_licensed_scan_date')
#    $obj.PSObject.Properties.Remove('last_scan_id')
#    $obj.PSObject.Properties.Remove('last_schedule_id')
#    $obj.PSObject.Properties.Remove('azure_vm_id')
#    $obj.PSObject.Properties.Remove('azure_resource_id')
#    $obj.PSObject.Properties.Remove('gcp_project_id')
#    $obj.PSObject.Properties.Remove('gcp_zone')
#    $obj.PSObject.Properties.Remove('gcp_instance_id')
#    $obj.PSObject.Properties.Remove('aws_ec2_isntance_ami_id')
#    $obj.PSObject.Properties.Remove('aws_ec2_instance_id')
#    $obj.PSObject.Properties.Remove('agent_uuid')
#    $obj.PSObject.Properties.Remove('bios_uuid')
#    $obj.PSObject.Properties.Remove('network_id')
#    $obj.PSObject.Properties.Remove('network_name')
#    $obj.PSObject.Properties.Remove('aws_owner_id')
#    $obj.PSObject.Properties.Remove('aws_availability_zone')
#    $obj.PSObject.Properties.Remove('aws_region')
#    $obj.PSObject.Properties.Remove('aws_vpc_id')
#    $obj.PSObject.Properties.Remove('aws_ec2_instance_group_name')
#    $obj.PSObject.Properties.Remove('aws_ec2_instance_state_name')
#    $obj.PSObject.Properties.Remove('aws_ec2_instance_type')
#    $obj.PSObject.Properties.Remove('aws_subnet_id')
#    $obj.PSObject.Properties.Remove('aws_ec2_product_code')
#    $obj.PSObject.Properties.Remove('aws_ec2_name')
#    $obj.PSObject.Properties.Remove('mcafee_epo_guid')
#    $obj.PSObject.Properties.Remove('mcafee_epo_agent_guid')
#    $obj.PSObject.Properties.Remove('servicenow_sysid')
#    $obj.PSObject.Properties.Remove('bigfix_asset_id')
#    $obj.PSObject.Properties.Remove('agent_names')
#    $obj.PSObject.Properties.Remove('installed_software')
#    $obj.PSObject.Properties.Remove('ipv4s')
#    $obj.PSObject.Properties.Remove('ipv6s')
#    $obj.PSObject.Properties.Remove('fqdns')
#    $obj.PSObject.Properties.Remove('mac_addresses')
#    $obj.PSObject.Properties.Remove('netbios_names')
#    $obj.PSObject.Properties.Remove('operating_systems')
#    $obj.PSObject.Properties.Remove('system_types')
#    $obj.PSObject.Properties.Remove('hostnames')
#    $obj.PSObject.Properties.Remove('ssh_fingerprints')
#    $obj.PSObject.Properties.Remove('qualys_asset_ids')
#    $obj.PSObject.Properties.Remove('qualys_host_ids')
#    $obj.PSObject.Properties.Remove('manufacturer_tps_ids')
#    $obj.PSObject.Properties.Remove('symantec_ep_hardware_keys')
#    $obj.PSObject.Properties.Remove('sources')
#    $obj.PSObject.Properties.Remove('tags')
#    $obj.PSObject.Properties.Remove('network_interfaces')
#    $obj.PSObject.Properties.Remove('acr_score')
#    $obj.PSObject.Properties.Remove('exposure_score')
#}

Write-Host "4.  $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Saving Tenable chunk {1} of {$($tempAssetStatus.chunks_available)} exported Assets data to ($($exportFileLocation)Tenable_Data_Assets_All.json) for Power BI Refresh."
$assetsData | ConvertTo-Json | Out-File "$($exportFileLocation)Tenable_Data_Assets_All.json"

#Write-Host "Begin saving backup copy of json file with timestamp: $(Get-Date -Format 'MM/dd/yyyy HH:mm')"
#    b. Create a copy of current json file with date-time stamp.
#$assetsData.Content | ConvertTo-Json | Out-File "$($exportFileLocation)Tenable_Data_$(Get-Date -Format 'MMddyyyy_HHmm').json"

Write-Host "5.  $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Tenable assets {uuid: $($assetExportuuid)} exported from data chunk."
#--------------------------------------------------------------------------------------------------------------------
# Section 2 - Export Vulnerabilites {Low, Medium, High, Critical}
#--------------------------------------------------------------------------------------------------------------------
Write-Host "6.  $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Triggering Tenable export of Non-Info Vulnerabilities."

$customBody = '{"filters":{"severity":["low","medium","high","critical"]},"num_assets":5000}' 

$exportNonInfoVulns = Invoke-WebRequest -Uri 'https://cloud.tenable.com/vulns/export' -Method POST -Headers $headers -ContentType 'application/json' -Body $customBodyNonInfo -UseBasicParsing
$temp2 = $exportNonInfoVulns.Content | ConvertFrom-Json
$vulnExportuuid = $temp2.export_uuid

Write-Host "7.  $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Checking status of Tenable Non-Info Vulnerabilities export {uuid: $($vulnExportuuid)}. Loop checks again every 1min."    
Do {
    Start-Sleep -s 60
    $vulnNonInfoStatusCheck = Invoke-WebRequest -Uri "https://cloud.tenable.com/vulns/export/$($vulnExportuuid)/status" -Method GET -Headers $headers -UseBasicParsing
        #$vulnNonInfoStatusCheck.RawContent - Ex. Note: {"status":"FINISHED","chunks_available":[1,2,3,4,5,6,7,8,9,10,11,12]}
    $tempVulnNonInfoStatus = $vulnNonInfoStatusCheck.Content | ConvertFrom-Json
    Write-Host "    $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - $($tempVulnNonInfoStatus.status)"
    #Write-Host $statusCheck.RawContent
} Until ($tempVulnNonInfoStatus.status -eq "FINISHED")

# Create a For loop to get each chunk of data?

Write-Host "8.  $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Downloading data [chunk] 1 of {$($tempVulnNonInfoStatus.total_chunks)} for Tenable Non-Info Vulnerabilities export {uuid: $($vulnExportuuid)}."
$vulnNonInfoData = Invoke-RestMethod -Uri "https://cloud.tenable.com/vulns/export/$($vulnExportuuid)/chunks/1" -Method GET -Headers $headers -UseBasicParsing

# Remove properties we do not need in Power BI data model from being written to file.
#foreach ($obj in $vulnNonInfoData) {
#    $obj.PSObject.Properties.Remove('asset')
#    $obj.PSObject.Properties.Remove('output')
#    $obj.PSObject.Properties.Remove('plugin')
#    $obj.PSObject.Properties.Remove('port')
#    $obj.PSObject.Properties.Remove('scan')
#    $obj.PSObject.Properties.Remove('severity')
#    $obj.PSObject.Properties.Remove('severity_id')
#    $obj.PSObject.Properties.Remove('severity_default_id')
#    $obj.PSObject.Properties.Remove('severity_modification_type')
#    $obj.PSObject.Properties.Remove('first_found')
#    $obj.PSObject.Properties.Remove('last_found')
#    $obj.PSObject.Properties.Remove('state')
#    $obj.PSObject.Properties.Remove('indexed')
#}

Write-Host "9.  $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Saving Tenable [chunk] {1} of {$($tempVulnNonInfoStatus.total_chunks)} exported Non-Info Vulnerabilities data to ($($exportFileLocation)Tenable_Data_Vulnerabilities_NonInfo.json) for Power BI Refresh."
$vulnNonInfoData | ConvertTo-Json | Out-File "$($exportFileLocation)Tenable_Data_Vulnerabilities_NonInfo.json"

Write-Host "10. $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Tenable Non-Info Vulnerabilities {uuid: $($vulnExportuuid)} exported from data chunk."
#--------------------------------------------------------------------------------------------------------------------
# Section 3 - Export Informational Vulnerabilites
#--------------------------------------------------------------------------------------------------------------------
Write-Host "11. $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Triggering Tenable export of {Info}rmational Vulnerabilities."

$exportInfoVulns = Invoke-WebRequest -Uri 'https://cloud.tenable.com/vulns/export' -Method POST -Headers $headers -ContentType 'application/json' -Body $customBodyInfoOnly -UseBasicParsing
$temp3 = $exportInfoVulns.Content | ConvertFrom-Json
$vulnInfoExportuuid = $temp3.export_uuid

Write-Host "12. $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Checking status of Tenable {Info}rmational Vulnerabilities export {uuid: $($vulnInfoExportuuid)}."       
Do {
    Start-Sleep -s 30
    $vulnInfoStatusCheck = Invoke-WebRequest -Uri "https://cloud.tenable.com/vulns/export/$($vulnInfoExportuuid)/status" -Method GET -Headers $headers -UseBasicParsing
        #$vulnInfoStatusCheck.RawContent - Ex. Note: {"status":"FINISHED","chunks_available":[1,2,3,4,5,6,7,8,9,10,11,12]}
    $tempVunlInfoStatusCheck = $vulnInfoStatusCheck.Content | ConvertFrom-Json
    Write-Host "    $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - $($tempVunlInfoStatusCheck.status)"
} Until ($tempVunlInfoStatusCheck.status -eq "FINISHED")

# Create a For loop to get each chunk of data?

Write-Host "13. $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Downloading data [chunk] 1 of {$($tempVunlInfoStatusCheck.total_chunks)} for Tenable Non-Info Vulnerabilities export {uuid: $($vulnExportuuid)}."
$vulnInfoData = Invoke-RestMethod -Uri "https://cloud.tenable.com/vulns/export/$($vulnInfoExportuuid)/chunks/1" -Method GET -Headers $headers -UseBasicParsing

# Remove properties we do not need in Power BI data model from being written to file.
#foreach ($obj in $vulnInfoData) {
#    $obj.PSObject.Properties.Remove('asset')
#    $obj.PSObject.Properties.Remove('output')
#    $obj.PSObject.Properties.Remove('plugin')
#    $obj.PSObject.Properties.Remove('port')
#    $obj.PSObject.Properties.Remove('scan')
#    $obj.PSObject.Properties.Remove('severity')
#    $obj.PSObject.Properties.Remove('severity_id')
#    $obj.PSObject.Properties.Remove('severity_default_id')
#    $obj.PSObject.Properties.Remove('severity_modification_type')
#    $obj.PSObject.Properties.Remove('first_found')
#    $obj.PSObject.Properties.Remove('last_found')
#    $obj.PSObject.Properties.Remove('state')
#    $obj.PSObject.Properties.Remove('indexed')
#}

Write-Host "14. $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Saving Tenable [chunk] {1} of {$($tempVunlInfoStatusCheck.total_chunks)} exported {Info}rmational Vulnerabilities data to ($($exportFileLocation)Tenable_Data_Vulnerabilities_Informational.json) for Power BI Refresh."
$vulnInfoData | ConvertTo-Json | Out-File "$($exportFileLocation)Tenable_Data_Vulnerabilities_InformationalOnly.json"

Write-Host "15. $(Get-Date -Format 'MM/dd/yyyy HH:mm tt') - Tenable {Info}rmational Vulnerabilities {uuid: $($vulnInfoExportuuid)} exported from data chunk."

Stop-Transcript 
