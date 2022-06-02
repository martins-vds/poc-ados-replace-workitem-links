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

function FormatQuery([string] $projectName, [object] $mapping) {
    $queryFormat = "SELECT [System.Id] FROM workitemLinks WHERE ( [Source].[System.TeamProject] = '{0}' AND [Source].[System.WorkItemType] <> '' AND [Source].[System.State] <> '' ) AND ( {1} ) AND ( [Target].[System.TeamProject] = '{0}' AND [Target].[System.WorkItemType] <> '' ) ORDER BY [System.Id] MODE (MustContain)"
    $queryFilter = $mapping | ForEach-Object {
        "[System.Links.LinkType] = '{0}'" -f $_.oldLinkType
    } | Join-String -Separator " OR "

    return $queryFormat -f $projectName, $queryFilter
}

function ReplaceLinks ($workItemRelation, [object] $mapping) {    
    $sourceWorkItem = Invoke-RestMethod -Uri "$($workItemApi -f $workItemRelation.source.id)&`$expand=All" -Method Get -Headers $authenicationHeader -ContentType "application/json"
    
    $i = 0
    $found = $false
    $relationCount = $sourceWorkItem.relations.Length

    while ($i -lt $relationCount -and $found -eq $false) {
        if ($sourceWorkItem.relations[$i].rel -eq $mapping.oldLinkType) {
            $found = $true
        }
        
        if (!$found) {
            $i++
        }
    }

    if (!$found) {
        Write-Host "    Link already replaced." -ForegroundColor Yellow
        return $false
    }

    $addLinkOperation = @(
        @{
            "op"    = "add"
            "path"  = "/relations/-"
            "value" = @{
                "rel"        = $mapping.newLinkType
                "url"        = $workItemRelation.target.url
                "attributes" = @{
                    "comment" = "Changing link type from '{0}' to '{1}'" -f $mapping.oldLinkType, $mapping.newLinkType
                }
            }
        }
    ) 
    
    Write-Host "    Adding link '$($mapping.newLinkType)' to work item with id '$($workItemRelation.target.id)'..." -ForegroundColor White

    try {
        $null = Invoke-RestMethod -Uri "$($workItemApi -f $workItemRelation.source.id)" -Method Patch -Body $($addLinkOperation | ConvertTo-Json -Depth 3 -AsArray) -Headers $authenicationHeader -ContentType "application/json-patch+json"
    }
    catch {
        if ($_.Exception.ErrorDetails){
            $exceptionDetails = $_.Exception.ErrorDetails | ConvertFrom-Json -AsHashtable
            
            if($exceptionDetails.typeName -eq "Microsoft.TeamFoundation.WorkItemTracking.Server.WorkItemLinksLimitExceededException"){
                Write-Host "    Failed to add link. Reason: work item with id '$($workItemRelation.target.id)' will exceed the 1000 link limit." -ForegroundColor Red
            }else{
                throw
            }
        }
    }

    $removeLinkOperation = @(@{
            "op"   = "remove"
            "path" = "/relations/{0}" -f $i
        }) | ConvertTo-Json -AsArray
    
    
    Write-Host "    Removing link type '$($mapping.oldLinkType)'..." -ForegroundColor White

    $null = Invoke-RestMethod -Uri "$($workItemApi -f $workItemRelation.source.id)" -Method Patch -Body $removeLinkOperation -Headers $authenicationHeader -ContentType "application/json-patch+json"

    return $true
}

$authenicationHeader = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($Token)")) }

$wiqlApi = "$CollectionUrl/_apis/wit/wiql?api-version=6.0&`$top={0}&`$skip={1}"
$workItemApi = "$CollectionUrl/_apis/wit/workitems/{0}?api-version=6.0"

$linkReplacementMapping = @(
    @{
        "oldLinkType" = "System.LinkTypes.Hierarchy-Forward"
        "newLinkType" = "System.LinkTypes.Related"
    }
)

$wiql = @{
    "query" = FormatQuery -projectName $Project -mapping $linkReplacementMapping
} | ConvertTo-Json -Depth 1

$summaryLogName = "replacement-$Project-$(Get-Date -UFormat "%Y-%m-%d_%H-%m-%S")"
$summaryLogFile = "$summaryLogName.log"

$workItemsProcessed = 0
$processingTime = Measure-Command {
    $top = 100
    $skip = 0
    do {
        $summaryLog = @()

        $response = Invoke-RestMethod -Uri $($wiqlApi -f $top, $skip) -Method Post -Headers $authenicationHeader -Body $wiql -ContentType "application/json"
    
        if ($response.workItemRelations.Length -gt 0) {
            foreach ($relation in $response.workItemRelations) {
                if ($relation.rel) {
                    $mapping = $linkReplacementMapping | Where-Object { $_.oldLinkType -eq $relation.rel }
            
                    if ($mapping) {
                        Write-Host "Replacing links on work item with id '$($relation.source.id)'..." -ForegroundColor White
                        if (ReplaceLinks -workItemRelation $relation -mapping $mapping) {
                            $summaryLog += @{
                                Project     = $Project
                                SourceId    = $relation.source.Id
                                TargetId    = $relation.target.id
                                OldLinkType = $mapping.oldLinkType
                                NewLinkType = $mapping.newLinkType
                            }
                            $workItemsProcessed++
                        }
                    }
                    else {
                        Write-Host "Skipped work item with id '$($relation.source.id)'. No mapping found for link type '$($relation.rel)'" -ForegroundColor Yellow
                    }
                }
            }
    
        }
        else {
            if ($skip -eq 0) {
                Write-Host "No work items found." -ForegroundColor Yellow
            }    
        }

        if (Test-Path -Path $summaryLogFile) {
            $summaryLog | Export-Csv -Path $summaryLogFile -UseQuotes AsNeeded -Append
        }
        else {
            $summaryLog | Export-Csv -Path $summaryLogFile -UseQuotes AsNeeded -Force
        }

        $skip += $top
    } until (
        $response.workItemRelations.Length -le 0
    )
}

Write-Host "Processed $workItemsProcessed work items in: $("{0:dd}d:{0:hh}h:{0:mm}m:{0:ss}s" -f $processingTime)" -ForegroundColor DarkMagenta
Write-Host "Summary log can be found at $($(Get-ChildItem -Path $summaryLogFile).FullName)" -ForegroundColor DarkCyan
Write-Host "Done." -ForegroundColor Green
