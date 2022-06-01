[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $CollectionUrl,
    [Parameter(Mandatory = $true)]
    [string]
    $Project,
    [Parameter(Mandatory = $true)]
    [string]
    $Token
)

$ErrorActionPreference = "Stop"

function FormatQuery([string] $projectName, [object] $mapping){
    $queryFormat = "SELECT [System.Id] FROM workitemLinks WHERE ( [Source].[System.TeamProject] = '{0}' AND [Source].[System.WorkItemType] <> '' AND [Source].[System.State] <> '' ) AND ( {1} ) AND ( [Target].[System.TeamProject] = '{0}' AND [Target].[System.WorkItemType] <> '' ) ORDER BY [System.Id] MODE (MustContain)"
    $queryFilter = $mapping | % {
        "[System.Links.LinkType] = '{0}'" -f $_.oldLinkType
    } | Join-String -Separator " OR "

    return $queryFormat -f $projectName, $queryFilter
}

function ReplaceLinks ($workItemRelation, [object] $mapping){
    Write-Host "Fetching work item with id '$($workItemRelation.source.id)'..." -ForegroundColor White

    $sourceWorkItem = Invoke-RestMethod -Uri "$($workItemApi -f $workItemRelation.source.id)&`$expand=All" -Method Get -Headers $authenicationHeader -ContentType "application/json"
    
    $relationCount = $sourceWorkItem.relations.Length
    for ($i = 0; $i -lt $relationCount; $i++) {
        if ($sourceWorkItem.relations[$i].rel -eq $mapping.oldLinkType){
            break;
        }
    }

    $removeLinkOperation = @(@{
        "op" = "remove"
        "path" = "/relations/{0}" -f $i
    }) | ConvertTo-Json -AsArray
    
    Invoke-RestMethod -Uri "$($workItemApi -f $workItemRelation.source.id)" -Method Patch -Body $removeLinkOperation -Headers $authenicationHeader -ContentType "application/json-patch+json" | Out-Null

    Write-Host "    Removed link type '$($mapping.oldLinkType)'" -ForegroundColor Yellow

    $addLinkOperation = @(
        @{
            "op" = "add"
            "path" = "/relations/-"
            "value" = @{
                "rel"= $mapping.newLinkType
                "url"= $workItemRelation.target.url
                "attributes"= @{
                    "comment"= "Changing link type from '{0}' to '{1}'" -f $mapping.oldLinkType, $mapping.newLinkType
                }
            }
        }
    ) 
    
    if($mapping.tags){
        $addLinkOperation += 
        @{
            "op"= "add"
            "path"= "/fields/System.Tags"
            "value"= $mapping.tags
        }
    }

    Invoke-RestMethod -Uri "$($workItemApi -f $workItemRelation.source.id)" -Method Patch -Body $($addLinkOperation | ConvertTo-Json -Depth 3 -AsArray) -Headers $authenicationHeader -ContentType "application/json-patch+json" | Out-Null

    Write-Host "    Added link '$($mapping.newLinkType)'" -ForegroundColor Yellow
}

$authenicationHeader = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($Token)")) }

$wiqlApi = "$CollectionUrl/_apis/wit/wiql?api-version=6.0"
$workItemApi = "$CollectionUrl/_apis/wit/workitems/{0}?api-version=6.0"

$linkReplacementMapping = @(
    @{
        "oldLinkType" = "System.LinkTypes.Hierarchy-Forward"
        "newLinkType" = "System.LinkTypes.Related"
        "tags" = "Tag1;Tag2"
    }
)

$wiql = @{
    "query" = FormatQuery -projectName $Project -mapping $linkReplacementMapping
} | ConvertTo-Json -Depth 1

Write-Host "Fetching work items with matching links..." -ForegroundColor Cyan

$response = Invoke-RestMethod -Uri $wiqlApi -Method Post -Headers $authenicationHeader -Body $wiql -ContentType "application/json"

if($response.workItemRelations.Length -gt 0){

    Write-Host "Found $($response.workItemRelations.Length) work items." -ForegroundColor Green

    foreach ($relation in $response.workItemRelations) {
        $mapping = $linkReplacementMapping | where {$_.oldLinkType -eq $relation.rel}
        
        if($mapping){
            ReplaceLinks -workItemRelation $relation -mapping $mapping
        }
    }

    Write-Host "Done." -ForegroundColor Green
}else{
    Write-Host "No work items found." -ForegroundColor Red
    Write-Host "Done." -ForegroundColor Red
}

