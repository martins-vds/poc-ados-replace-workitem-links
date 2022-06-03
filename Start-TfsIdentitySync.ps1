[CmdletBinding()]
param (
    [Parameter(Mandatory = $True)]
    [string]
    $ServerUrl
)

$headers = @{
    SOAPAction = "http://microsoft.com/webservices/QueueJobs"
}

$body =
@'
<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
        <soap:Body>
        <QueueJobs xmlns="http://microsoft.com/webservices/">
            <jobIds>
                <guid>544dd581-f72a-45a9-8de0-8cd3a5f29dfe</guid>
            </jobIds>
        </QueueJobs>
    </soap:Body>
</soap:Envelope>
'@

$credential = Get-Credential

Invoke-RestMethod -Uri "$Serverl/TeamFoundation/Administration/v3.0/JobService.asmx" -Method Post -Headers $headers -Body $body -Credential $credential -ContentType "text/xml"