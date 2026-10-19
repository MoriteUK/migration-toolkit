<#
.SYNOPSIS
    Sends a test message to a recipient and watches the sender's mailbox for NDRs.
    If an NDR comes back, waits (default 3 minutes) and resends, repeating until a
    test message is delivered without an NDR (or MaxAttempts is reached).

.DESCRIPTION
    Uses Microsoft Graph (Microsoft.Graph.Users.Actions + Microsoft.Graph.Mail).
    Each attempt carries a unique tracking token in the subject and body, so only
    NDRs for the current attempt are counted (late NDRs from earlier attempts are ignored).

    An attempt is considered successful when no NDR for it arrives within
    -NdrWaitMinutes of sending.

.PARAMETER SenderMailbox
    Mailbox the test is sent from. NDRs are checked in this mailbox.

.PARAMETER Recipient
    Address to send the test message to.

.PARAMETER RetryDelayMinutes
    How long to wait after an NDR before resending. Default 3.

.PARAMETER NdrWaitMinutes
    How long to watch for an NDR after each send before declaring success. Default 3.

.PARAMETER PollSeconds
    How often to check the mailbox during the NDR wait window. Default 30.

.PARAMETER MaxAttempts
    Stop after this many attempts. 0 = unlimited. Default 0.

.PARAMETER TenantId / ClientId / CertificateThumbprint
    Optional app-only auth. If omitted, connects interactively with delegated
    Mail.Send + Mail.Read scopes (unless already connected).

.EXAMPLE
    .\Test-MailFlowNDR.ps1 -SenderMailbox admin@contoso.com -Recipient user@fabrikam.com

.EXAMPLE
    .\Test-MailFlowNDR.ps1 -SenderMailbox svc-mailtest@contoso.com -Recipient user@fabrikam.com `
        -TenantId contoso.onmicrosoft.com -ClientId <appId> -CertificateThumbprint <thumb> -MaxAttempts 20
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SenderMailbox,
    [Parameter(Mandatory)][string]$Recipient,
    [int]$RetryDelayMinutes = 3,
    [int]$NdrWaitMinutes = 3,
    [int]$PollSeconds = 30,
    [int]$MaxAttempts = 0,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$LogPath = (Join-Path $PSScriptRoot ("MailFlowNDR_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date)))
)

$ErrorActionPreference = 'Stop'

# Hide WARNING: output from Graph / module cmdlets
$WarningPreference = 'SilentlyContinue'
$PSDefaultParameterValues['*:WarningAction'] = 'SilentlyContinue'

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','OK')][string]$Level = 'INFO')
    $line = "{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}" -f (Get-Date), $Level, $Message
    $colour = @{ INFO = 'Gray'; WARN = 'Yellow'; ERROR = 'Red'; OK = 'Green' }[$Level]
    Write-Host $line -ForegroundColor $colour
    Add-Content -Path $LogPath -Value $line
}

#region Connect
foreach ($mod in 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Users.Actions', 'Microsoft.Graph.Mail') {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        throw "Module $mod not found. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    }
    Import-Module $mod -ErrorAction Stop
}

$ctx = Get-MgContext
if ($TenantId -and $ClientId -and $CertificateThumbprint) {
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
}
elseif (-not $ctx -or ($ctx.Scopes -notcontains 'Mail.Send' -and $ctx.AuthType -eq 'Delegated')) {
    Connect-MgGraph -Scopes 'Mail.Send', 'Mail.Read' -NoWelcome
}
$ctx = Get-MgContext
$who = if ($ctx.Account) { $ctx.Account } else { "app $($ctx.ClientId)" }
Write-Log "Connected to Graph as $who"
#endregion

# Patterns that identify a non-delivery report
$ndrSubjectPattern = '^(Undeliverable|Undelivered|Delivery Status Notification \(Failure\)|Delivery has failed|Mail delivery failed|Returned mail|Failure Notice|Non[- ]?delivery)'
$ndrSenderPattern  = '^(postmaster@|mailer-daemon@|MicrosoftExchange[0-9a-f]{32}@)'

function Send-TestMessage {
    param([int]$Attempt, [string]$Token)
    $subject = "Mail flow test [$Token] attempt $Attempt"
    $body = @"
This is an automated mail flow test.

Tracking token : $Token
Attempt        : $Attempt
Sent (UTC)     : $((Get-Date).ToUniversalTime().ToString('u'))
From           : $SenderMailbox
To             : $Recipient

No action is required.
"@
    $params = @{
        message = @{
            subject      = $subject
            body         = @{ contentType = 'Text'; content = $body }
            toRecipients = @(@{ emailAddress = @{ address = $Recipient } })
        }
        saveToSentItems = $true
    }
    Send-MgUserMail -UserId $SenderMailbox -BodyParameter $params
    return $subject
}

function Get-NdrForToken {
    param([datetime]$SinceUtc, [string]$Token)
    $since = $SinceUtc.AddSeconds(-30).ToString('yyyy-MM-ddTHH:mm:ssZ')
    # Searches all folders (NDRs occasionally land in Junk)
    $msgs = Get-MgUserMessage -UserId $SenderMailbox -All `
        -Filter "receivedDateTime ge $since" `
        -Property 'id,subject,from,receivedDateTime,bodyPreview' -ErrorAction Stop

    $msgs | Where-Object {
        $fromAddr = $_.From.EmailAddress.Address
        ($_.Subject -match $ndrSubjectPattern -or $fromAddr -match $ndrSenderPattern) -and
        ($_.Subject -like "*$Token*" -or $_.BodyPreview -like "*$Token*")
    }
}

Write-Log "Starting mail flow test: $SenderMailbox -> $Recipient (NDR wait ${NdrWaitMinutes}m, retry delay ${RetryDelayMinutes}m, max attempts $(if ($MaxAttempts) { $MaxAttempts } else { 'unlimited' }))"
Write-Log "Logging to $LogPath"

$attempt = 0
while ($true) {
    $attempt++
    if ($MaxAttempts -gt 0 -and $attempt -gt $MaxAttempts) {
        Write-Log "Reached MaxAttempts ($MaxAttempts) without a clean delivery. Giving up." 'ERROR'
        exit 1
    }

    $token   = ([guid]::NewGuid().ToString('N')).Substring(0, 12)
    $sentUtc = (Get-Date).ToUniversalTime()
    try {
        $subject = Send-TestMessage -Attempt $attempt -Token $token
        Write-Log "Attempt ${attempt}: sent '$subject'"
    }
    catch {
        Write-Log "Attempt ${attempt}: send failed - $($_.Exception.Message). Retrying in $RetryDelayMinutes minute(s)." 'ERROR'
        Start-Sleep -Seconds ($RetryDelayMinutes * 60)
        continue
    }

    # Watch for an NDR for this attempt
    $deadline = (Get-Date).AddMinutes($NdrWaitMinutes)
    $ndr = $null
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $PollSeconds
        try {
            $ndr = Get-NdrForToken -SinceUtc $sentUtc -Token $token | Select-Object -First 1
        }
        catch {
            Write-Log "Error checking for NDRs: $($_.Exception.Message)" 'WARN'
        }
        if ($ndr) { break }
        Write-Host "  ...no NDR yet ($([int]($deadline - (Get-Date)).TotalSeconds)s left in window)" -ForegroundColor DarkGray
    }

    if (-not $ndr) {
        Write-Log "Attempt ${attempt}: no NDR received within $NdrWaitMinutes minute(s). Mail flow to $Recipient looks healthy." 'OK'
        exit 0
    }

    Write-Log "Attempt ${attempt}: NDR received at $($ndr.ReceivedDateTime) from $($ndr.From.EmailAddress.Address) - '$($ndr.Subject)'" 'WARN'
    if ($ndr.BodyPreview) {
        Write-Log ("  NDR preview: " + ($ndr.BodyPreview -replace '\s+', ' ').Substring(0, [Math]::Min(300, ($ndr.BodyPreview -replace '\s+', ' ').Length))) 'WARN'
    }
    Write-Log "Waiting $RetryDelayMinutes minute(s) before resending..."
    Start-Sleep -Seconds ($RetryDelayMinutes * 60)
}
