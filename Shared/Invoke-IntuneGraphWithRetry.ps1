function Invoke-IntuneGraphWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [int]$MaxRetries = 5,
        [int]$BaseDelaySec = 2,
        [scriptblock]$OnRetry
    )

    $attempt = 0
    while ($true) {
        try { return (& $ScriptBlock) }
        catch {
            $attempt++
            $message = $_.ToString()
            $statusCode = $null
            try { if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode } } catch { }
            if (-not $statusCode) {
                if ($message -match '429|TooManyRequests') { $statusCode = 429 }
                elseif ($message -match '503|ServiceUnavail') { $statusCode = 503 }
                elseif ($message -match '502|BadGateway') { $statusCode = 502 }
                elseif ($message -match '500|InternalServer') { $statusCode = 500 }
            }
            if ($attempt -gt $MaxRetries -or $statusCode -notin @(429, 500, 502, 503)) { throw }
            $delaySeconds = $BaseDelaySec * [math]::Pow(2, $attempt - 1)
            if ($statusCode -eq 429) {
                try {
                    $retryAfter = $_.Exception.Response.Headers.GetValues('Retry-After') | Select-Object -First 1
                    if ($retryAfter) { $delaySeconds = [int]$retryAfter }
                } catch { }
                if ($delaySeconds -lt 30) { $delaySeconds = 30 }
            }
            if ($OnRetry) { & $OnRetry $attempt $delaySeconds $message }
            Start-Sleep -Seconds $delaySeconds
        }
    }
}
