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
    $Token,
    [Parameter(Mandatory = $true)]
    [ValidateScript({
            if ( -Not ($_ | Test-Path -PathType leaf) ) {
                throw "File '$_' does not exist."
            }
            return $true
        })]
    [System.IO.FileInfo]
    $MappingsFile
)

$ErrorActionPreference = "Stop"

function ParseMappingsFile() {    
    return @(Get-Content -Path $MappingsFile -Raw -Encoding utf8 | ConvertFrom-Json)
}

function FormatQuery([string] $projectName, [object] $mappings) {
    $queryFormat = "SELECT [System.Id] FROM workitemLinks WHERE ( [Source].[System.TeamProject] = '{0}' AND [Source].[System.WorkItemType] <> '' AND [Source].[System.State] <> '' ) AND ( {1} ) AND ( [Target].[System.TeamProject] = '{0}' AND [Target].[System.WorkItemType] <> '' ) ORDER BY [System.Id] MODE (MustContain)"
    $queryFilter = $mappings | ForEach-Object {
        "[System.Links.LinkType] = '{0}'" -f $_.oldLinkType
    } | Join-String -Separator " OR "

    return $queryFormat -f $projectName, $queryFilter
}

function ReplaceLinks ($workItemRelation, [object] $mapping) {    
    $sourceWorkItem = Invoke-RestMethod -Uri "$($workItemApi -f $workItemRelation.source.id)&`$expand=All" -Method Get -Headers $authenicationHeader -ContentType "application/json"
    
    $i = 0
    $found = $false
    $relationCount = $sourceWorkItem.relations.Length

    for ($i = 0; $i -lt $relationCount; $i++) {
        if ($sourceWorkItem.relations[$i].rel -eq $mapping.oldLinkType) {
            Write-Host "    Replacing old link to work item with id '$($workItemRelation.target.id)' with new link '$($mapping.newLinkType)'..." -ForegroundColor White
            
            $found = $true

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

            $null = Invoke-RestMethod -Uri "$($workItemApi -f $workItemRelation.source.id)" -Method Patch -Body $($addLinkOperation | ConvertTo-Json -Depth 3 -AsArray) -Headers $authenicationHeader -ContentType "application/json-patch+json"

            $removeLinkOperation = @(@{
                    "op"   = "remove"
                    "path" = "/relations/{0}" -f $i
                }) | ConvertTo-Json -AsArray
    
            $null = Invoke-RestMethod -Uri "$($workItemApi -f $workItemRelation.source.id)" -Method Patch -Body $removeLinkOperation -Headers $authenicationHeader -ContentType "application/json-patch+json"
        }
    }

    if (!$found){
        Write-Host "    Link already replaced." -ForegroundColor Yellow
    }

    return $found
}

$authenicationHeader = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($Token)")) }

$wiqlApi = "$CollectionUrl/_apis/wit/wiql?api-version=6.0&`$top={0}"
$workItemApi = "$CollectionUrl/_apis/wit/workitems/{0}?api-version=6.0"

$linkReplacementMappings = ParseMappingsFile

$wiql = @{
    "query" = FormatQuery -projectName $Project -mappings $linkReplacementMappings
} | ConvertTo-Json -Depth 1

$summaryLogName = "replacement-$Project-$(Get-Date -UFormat "%Y-%m-%d_%H-%m-%S")"
$summaryLogFile = "$summaryLogName.log"

$workItemsProcessed = 0
$processingTime = Measure-Command {
    $top = 100
    do {
        $summaryLog = @()

        $response = Invoke-RestMethod -Uri $($wiqlApi -f $top) -Method Post -Headers $authenicationHeader -Body $wiql -ContentType "application/json"
    
        if ($response.workItemRelations.Length -gt 0) {
            foreach ($relation in $response.workItemRelations) {
                if ($relation.rel) {
                    $mapping = $linkReplacementMappings | Where-Object { $_.oldLinkType -eq $relation.rel }
            
                    if ($mapping -and $mapping -isnot [array]) {
                        Write-Host "Replacing links on work item with id '$($relation.source.id)'..." -ForegroundColor White
                        try {
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
                        catch {
                            if ($_.ErrorDetails.Message) {
                                $exceptionDetails = $_.ErrorDetails.Message | ConvertFrom-Json
                                
                                Write-Host "    Failed to replace links. Reason: '$($exceptionDetails.message)'" -ForegroundColor Red
                            }
                            else {
                                Write-Host "    Failed to replace links. Reason: '$($_.Exception.ToString())'" -ForegroundColor Red
                            }
                        }
                    }
                    else {
                        if ($mapping -is [array]) {
                            Write-Host "Skipped work item with id '$($relation.source.id)'. Reason: Multiple mappings for link type '$($relation.rel)' found." -ForegroundColor Yellow
                        }
                        else {
                            Write-Host "Skipped work item with id '$($relation.source.id)'. Reason: No mapping found for link type '$($relation.rel)'." -ForegroundColor Yellow
                        }
                    }
                }
            }
    
        }
        else {
            if ($response.workItemRelations.Length -eq 0) {
                Write-Host "No work items found." -ForegroundColor Yellow
            }    
        }

        if (Test-Path -Path $summaryLogFile) {
            $summaryLog | Export-Csv -Path $summaryLogFile -UseQuotes AsNeeded -Append
        }
        else {
            $summaryLog | Export-Csv -Path $summaryLogFile -UseQuotes AsNeeded -Force
        }

    } until (
        $response.workItemRelations.Length -le 0
    )
}

Write-Host "Processed $workItemsProcessed work items in: $("{0:dd}d:{0:hh}h:{0:mm}m:{0:ss}s" -f $processingTime)" -ForegroundColor DarkMagenta
Write-Host "Summary log can be found at $($(Get-ChildItem -Path $summaryLogFile).FullName)" -ForegroundColor DarkCyan
Write-Host "Done." -ForegroundColor Green
