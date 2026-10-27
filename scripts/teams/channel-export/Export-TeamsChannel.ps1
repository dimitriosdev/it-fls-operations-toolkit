<#
.SYNOPSIS
Exports Microsoft Teams channel threads and replies to JSON.

.DESCRIPTION
Retrieves all messages and replies from a Teams channel,
filters by date range, converts HTML to plain text,
handles Graph pagination and retry logic, and exports
the result to a JSON file.

.REQUIREMENTS
Microsoft.Graph.Authentication

Authentication must be performed before execution:

Connect-MgGraph -Scopes "ChannelMessage.Read.All"

.CONFIGURATION
Values are loaded from the .env file.

Example .env:

TEAM_ID=
CHANNEL_ID=
FROM_DATE=2026-08-01
TO_DATE=2026-10-30
OUTPUT_FILE=teams-export.json
LOG_FILE=export.log
MAX_RETRIES=5
RETRY_DELAY_SECONDS=5
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --------------------------------------------------
# Load .env
# --------------------------------------------------

$envFile = ".env"

if (-not (Test-Path $envFile)) {
    throw ".env file not found."
}

Get-Content $envFile | ForEach-Object {

    $line = $_.Trim()

    if ([string]::IsNullOrWhiteSpace($line)) {
        return
    }

    if ($line.StartsWith("#")) {
        return
    }

    if ($line -match '^([^=]+)=(.*)$') {

        $name = $matches[1].Trim()
        $value = $matches[2].Trim()

        [System.Environment]::SetEnvironmentVariable(
            $name,
            $value,
            "Process"
        )
    }
}

# --------------------------------------------------
# Configuration
# --------------------------------------------------

$TeamId = $env:TEAM_ID
$ChannelId = $env:CHANNEL_ID

$FromDate = [datetime]$env:FROM_DATE
$ToDate = ([datetime]$env:TO_DATE).Date.AddDays(1).AddSeconds(-1)

$OutputFile = $env:OUTPUT_FILE
$LogFile = $env:LOG_FILE

$MaxRetries = if ($env:MAX_RETRIES) {
    [int]$env:MAX_RETRIES
}
else {
    5
}

$RetryDelaySeconds = if ($env:RETRY_DELAY_SECONDS) {
    [int]$env:RETRY_DELAY_SECONDS
}
else {
    5
}

# --------------------------------------------------
# Validation
# --------------------------------------------------

$RequiredVariables = @(
    "TEAM_ID",
    "CHANNEL_ID",
    "FROM_DATE",
    "TO_DATE",
    "OUTPUT_FILE"
)

foreach ($Variable in $RequiredVariables) {

    $Value = [Environment]::GetEnvironmentVariable($Variable)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "$Variable is missing from .env"
    }
}

# --------------------------------------------------
# Logging
# --------------------------------------------------

function Write-Log {

    param(
        [string]$Message
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $Entry = "[$Timestamp] $Message"

    Write-Host $Entry

    if (-not [string]::IsNullOrWhiteSpace($LogFile)) {
        Add-Content -Path $LogFile -Value $Entry
    }
}

# --------------------------------------------------
# HTML Cleanup
# --------------------------------------------------

function Convert-HtmlToPlainText {

    param(
        [string]$Html
    )

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }

    $Text = $Html

    $Text = $Text -replace '<br\s*/?>', "`n"
    $Text = $Text -replace '</p>', "`n"
    $Text = $Text -replace '</div>', "`n"

    $Text = [regex]::Replace(
        $Text,
        '<[^>]+>',
        ''
    )

    $Text = [System.Net.WebUtility]::HtmlDecode($Text)

    return $Text.Trim()
}

# --------------------------------------------------
# Retry Handler
# --------------------------------------------------

function Invoke-WithRetry {

    param(
        [scriptblock]$Action
    )

    $Attempt = 0

    while ($Attempt -lt $MaxRetries) {

        try {

            return & $Action

        }
        catch {

            $Attempt++

            $Message = $_.Exception.Message

            Write-Log "Attempt $Attempt failed."

            Write-Log $Message

            $StatusCode = $null

            try {
                $StatusCode = $_.Exception.Response.StatusCode.value__
            }
            catch {
            }

            if ($StatusCode -eq 429) {

                Write-Log "Graph throttling detected. Waiting 30 seconds."

                Start-Sleep -Seconds 30
            }
            else {

                if ($Attempt -lt $MaxRetries) {

                    Start-Sleep -Seconds $RetryDelaySeconds
                }
            }

            if ($Attempt -ge $MaxRetries) {
                throw
            }
        }
    }
}

# --------------------------------------------------
# Authentication Check
# --------------------------------------------------

$Context = Get-MgContext

if (-not $Context) {

    throw @"

Not connected to Microsoft Graph.

Run:

Connect-MgGraph -Scopes "ChannelMessage.Read.All"

"@
}

Write-Log "Connected account: $($Context.Account)"

function Get-NextLink {

    param(
        $Response
    )

    if ($null -eq $Response) {
        return $null
    }

    if ($Response -is [System.Collections.IDictionary]) {
        return $Response["@odata.nextLink"]
    }

    $NextLinkProperty = $Response.PSObject.Properties["@odata.nextLink"]

    if ($NextLinkProperty) {
        return $NextLinkProperty.Value
    }

    return $null
}

# --------------------------------------------------
# Get All Messages
# --------------------------------------------------

Write-Log "Loading channel messages..."

$Messages = @()

$MessageUri =
"https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/messages"

do {

    $Response = Invoke-WithRetry {

        Invoke-MgGraphRequest `
            -Method GET `
            -Uri $MessageUri
    }

    if ($Response.value) {

        $Messages += $Response.value

        Write-Log "Collected messages: $($Messages.Count)"
    }

    $MessageUri = Get-NextLink -Response $Response

}
while ($MessageUri)

Write-Log "Total messages retrieved: $($Messages.Count)"

# --------------------------------------------------
# Reply Retrieval
# --------------------------------------------------

function Get-AllReplies {

    param(
        [string]$MessageId
    )

    $Replies = @()

    $ReplyUri =
    "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$ChannelId/messages/$MessageId/replies"

    do {

        $Response = Invoke-WithRetry {

            Invoke-MgGraphRequest `
                -Method GET `
                -Uri $ReplyUri
        }

        if ($Response.value) {
            $Replies += $Response.value
        }

        $ReplyUri = Get-NextLink -Response $Response

    }
    while ($ReplyUri)

    return $Replies
}

# --------------------------------------------------
# Process Export
# --------------------------------------------------

$Export = @()

$Current = 0
$Total = $Messages.Count

foreach ($Message in $Messages) {

    $Current++

    Write-Progress `
        -Activity "Exporting Teams Messages" `
        -Status "$Current of $Total" `
        -PercentComplete (($Current / $Total) * 100)

    $Created = [datetime]$Message.createdDateTime

    if ($Created -lt $FromDate -or $Created -gt $ToDate) {
        continue
    }

    Write-Log "Processing thread: $($Message.id)"

    try {

        $Replies = @(Get-AllReplies -MessageId $Message.id)
    }
    catch {

        Write-Log "Unable to retrieve replies for $($Message.id)"

        $Replies = @()
    }

    $Export += [PSCustomObject]@{

        MessageId = $Message.id

        Author = $Message.from.user.displayName

        CreatedDateTime = $Message.createdDateTime

        LastModifiedDateTime = $Message.lastModifiedDateTime

        PlainText = Convert-HtmlToPlainText `
            $Message.body.content

        HtmlContent = $Message.body.content

        ReplyCount = $Replies.Count

        Replies = @(
            $Replies |
            Sort-Object createdDateTime |
            ForEach-Object {

                [PSCustomObject]@{

                    ReplyId = $_.id

                    Author = $_.from.user.displayName

                    CreatedDateTime = $_.createdDateTime

                    LastModifiedDateTime = $_.lastModifiedDateTime

                    PlainText = Convert-HtmlToPlainText `
                        $_.body.content

                    HtmlContent = $_.body.content
                }
            }
        )
    }
}

# --------------------------------------------------
# Final Payload
# --------------------------------------------------

$Payload = [PSCustomObject]@{

    ExportDateTime = (Get-Date).ToString("o")

    DateRange = @{
        From = $FromDate.ToString("o")
        To   = $ToDate.ToString("o")
    }

    TotalThreads = $Export.Count

    Threads = $Export
}

# --------------------------------------------------
# Save JSON
# --------------------------------------------------

Write-Log "Writing JSON output..."

$Payload |
    ConvertTo-Json -Depth 100 |
    Set-Content `
        -Path $OutputFile `
        -Encoding UTF8

Write-Log "Export complete."
Write-Log "Threads exported: $($Export.Count)"
Write-Log "Output file: $OutputFile"