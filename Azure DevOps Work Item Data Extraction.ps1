# Script:        Get DevOps WorkItem Data
# Last Modified: 4/10/2022 3:30 PM
#
# Note: If copying from Repo, add in your Personal-Access-Token (PAT) and sepcify export file location.
#--------------------------------------------------------------------------------
# 1. Use WIQL to get a queried for list of Work Items matching certain conditions
#--------------------------------------------------------------------------------
$token = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"

$organization = "AnalyticArchives"
$project = "ProDev"
$team = "ProDev Team"

$exportFileLocation = 'C:\Users\{YourUser}\Downloads'

$codedToken = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($token)"))
$AzureDevOpsAuthenicationHeader = @{Authorization = 'Basic ' + $codedToken}

$uri = "https://dev.azure.com/$organization/$project/$team/_apis/wit/wiql?api-version=6.0"

$body = @{
    #NOTE: Ignore the specified columns, only returns ID and URL by design
    "query" = "Select [System.Id], 
                      [System.Title], 
                      [System.State],
                      [Custom.FirstDetected] 
                      From WorkItems 
                      Where [System.WorkItemType] = 'Vulnerability' 
                      AND [State] = 'Done'"
} | ConvertTo-Json 

$response = Invoke-RestMethod -Method Post -Uri $uri -Headers $AzureDevOpsAuthenicationHeader -Body $body -ContentType 'application/json'

#--------------------------------------------------------------------------------
# 2. Use returned list of Work Item ID's to get field Data
#--------------------------------------------------------------------------------
$exportData = @()
$tempID = 0
foreach ($item in $response.workItems) {
    $tempID = $item.id
    $url = "https://dev.azure.com/$organization/$project/_apis/wit/workitems/$($tempID)?`$expand=All&api-version=5.0"
    $response = Invoke-RestMethod -Uri $url -Headers $AzureDevOpsAuthenicationHeader -Method Get -ContentType application/json 

    foreach ($field in $response) {

        $customObject = new-object PSObject -property @{

            "WorkItemID" = $field.fields.'System.Id'
            "URL" = $field.url
            "WorkItemType" = $field.fields.'System.WorkItemType'
            "Title" = $field.fields.'System.Title'
            "IterationPath" = $field.fields.'System.IterationPath'
            "State" = $field.fields.'System.State' 
            "AssignedTo" = $field.fields.'System.AssignedTo'.displayName
            "CoAssigned" = $field.fields.'Custom.CoAssigned'.displayName
            "FirstDetected" = $field.fields.'Custom.FirstDetected'
            "DueDate" = $field.fields.'Microsoft.VSTS.Scheduling.DueDate'
            "Priority" = $field.fields.'Microsoft.VSTS.Common.Priority'
            "Effort" = $field.fields.'Microsoft.VSTS.Scheduling.Effort'
            "ChangedDate" = $field.fields.'System.ChangedDate'
            "ChangedBy" = $field.fields.'System.ChangedBy'.displayName
            "ParentWorkItem" = $field.fields.'System.Parent'
            "RelatedWorkItems" = $field.relations
        } 
    $exportData += $customObject      
    }
}

$exportData | Select-Object -Property WorkItemID,
    URL,
    WorkItemType,
    Title,
    IterationPath,
    State,
    AssignedTo,
    CoAssigned,
    FirstDetected,
    DueDate,
    Priority,
    Effort,
    ChangedDate, 
    ChangedBy,
    ParentWorkItem,
    RelatedWorkItems

$exportData | ConvertTo-Json | Out-File "$($exportFileLocation)\DEVOPS_EXPORT.json"
