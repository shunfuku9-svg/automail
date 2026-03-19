param(
  [string]$CsvPath = ".\recipients.csv",
  [string]$TemplatePath = ".\message.txt",
  [string]$PromptPath = ".\ai-prompt.txt",
  [string]$PreviewPath = ".\preview.txt",
  [string]$Subject = "Consultation Request",
  [string]$Model = "claude-sonnet-4-20250514",
  [switch]$DryRun,
  [switch]$GenerateOnly,
  [switch]$UseClaude
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$SentMarker = "送信済み"

try {
  [Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
  [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
  $OutputEncoding = [Console]::OutputEncoding
} catch {
}

function Test-ValidEmail {
  param([string]$Email)
  if ([string]::IsNullOrWhiteSpace($Email)) { return $false }
  return $Email -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'
}

function Replace-TemplateTokens {
  param([string]$Template, [pscustomobject]$Row)
  $body = $Template
  foreach ($property in $Row.PSObject.Properties) {
    $token = "{{{0}}}" -f $property.Name
    $body = $body.Replace($token, [string]$property.Value)
  }
  return $body
}

function Get-Domains {
  param([string]$InputValue)
  if ([string]::IsNullOrWhiteSpace($InputValue)) { return @() }
  $tokens = $InputValue.ToLower().Split(@(",", " ", ";", "`r", "`n"), [System.StringSplitOptions]::RemoveEmptyEntries)
  $domains = New-Object System.Collections.Generic.List[string]
  foreach ($tokenValue in $tokens) {
    $token = $tokenValue.Trim()
    if ($token.Contains("@")) { $token = $token.Split("@")[-1] }
    if ($token -match '^https?://') {
      try { $token = ([System.Uri]$token).Host } catch {}
    }
    if ($token.StartsWith("www.")) { $token = $token.Substring(4) }
    if (-not $token.Contains(".")) { $token = "$token.com" }
    if ($token -match '^[a-z0-9.-]+\.[a-z]{2,}$' -and -not $domains.Contains($token)) {
      $domains.Add($token)
    }
  }
  return @($domains)
}

function Get-EmailCandidates {
  param([string]$RomajiName, [string]$DomainInput)
  $parts = @($RomajiName.ToLower().Split(@(" ", "　"), [System.StringSplitOptions]::RemoveEmptyEntries) |
    ForEach-Object { ($_ -replace '[^a-z]', '').Trim() } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  if ($parts.Count -eq 0) { return @() }

  $first = $parts[0]
  $last = if ($parts.Count -ge 2) { $parts[$parts.Count - 1] } else { "" }
  $firstInitial = if ($first.Length -gt 0) { $first.Substring(0, 1) } else { "" }
  $lastInitial = if ($last.Length -gt 0) { $last.Substring(0, 1) } else { "" }

  $patterns = New-Object System.Collections.Generic.List[string]
  if (-not [string]::IsNullOrWhiteSpace($last)) {
    @(
      "$first.$last", "$first$last", "${first}_$last", "${first}-$last",
      "$firstInitial$last", "$firstInitial.$last", "${firstInitial}_$last", "${firstInitial}-$last",
      "$first$lastInitial", "$first.$lastInitial", "${first}_$lastInitial", "${first}-$lastInitial",
      "$last.$first", "$last$first", "${last}_$first", "${last}-$first",
      "$last$firstInitial", "$last.$firstInitial", "$first", "$last",
      "$firstInitial$lastInitial", "$lastInitial$firstInitial"
    ) | ForEach-Object {
      if (-not [string]::IsNullOrWhiteSpace($_) -and -not $patterns.Contains($_)) { $patterns.Add($_) }
    }
  } else {
    @($first, $firstInitial) | ForEach-Object {
      if (-not [string]::IsNullOrWhiteSpace($_) -and -not $patterns.Contains($_)) { $patterns.Add($_) }
    }
  }

  $emails = New-Object System.Collections.Generic.List[string]
  foreach ($pattern in $patterns) {
    foreach ($domain in (Get-Domains -InputValue $DomainInput)) {
      $email = "$pattern@$domain"
      if (-not $emails.Contains($email)) { $emails.Add($email) }
    }
  }
  return @($emails)
}

function Invoke-ClaudeEmailDraft {
  param(
    [string]$ApiKey,
    [string]$ModelName,
    [string]$PromptTemplate,
    [string]$BaseDraft,
    [pscustomobject]$Row
  )

  $rowJson = $Row | ConvertTo-Json -Depth 5 -Compress
  $userPrompt = @"
Base draft:
$BaseDraft

Recipient data JSON:
$rowJson
"@

  $payload = @{
    model = $ModelName
    max_tokens = 1200
    system = $PromptTemplate
    messages = @(
      @{
        role = "user"
        content = @(
          @{
            type = "text"
            text = $userPrompt
          }
        )
      }
    )
  } | ConvertTo-Json -Depth 8

  $headers = @{
    "x-api-key" = $ApiKey
    "anthropic-version" = "2023-06-01"
    "content-type" = "application/json"
  }

  $response = Invoke-RestMethod `
    -Method Post `
    -Uri "https://api.anthropic.com/v1/messages" `
    -Headers $headers `
    -Body $payload

  $texts = New-Object System.Collections.Generic.List[string]
  foreach ($content in $response.content) {
    if ($content.type -eq "text" -and -not [string]::IsNullOrWhiteSpace([string]$content.text)) {
      $texts.Add([string]$content.text)
    }
  }

  $text = ($texts -join "`n").Trim()
  if ([string]::IsNullOrWhiteSpace($text)) {
    throw "Claude response did not include generated text."
  }
  return $text
}

if (-not (Test-Path -LiteralPath $CsvPath)) { throw "CSV file not found: $CsvPath" }
if (-not (Test-Path -LiteralPath $TemplatePath)) { throw "Template file not found: $TemplatePath" }
if ($UseClaude -and -not (Test-Path -LiteralPath $PromptPath)) { throw "Prompt file not found: $PromptPath" }

$rows = @(Import-Csv -LiteralPath $CsvPath)
if ($rows.Count -eq 0) { throw "CSV has no rows." }

$template = Get-Content -LiteralPath $TemplatePath -Raw -Encoding UTF8
$promptTemplate = if ($UseClaude) { Get-Content -LiteralPath $PromptPath -Raw -Encoding UTF8 } else { "" }

$apiKey = $env:ANTHROPIC_API_KEY
if ($UseClaude -and [string]::IsNullOrWhiteSpace($apiKey)) {
  throw "ANTHROPIC_API_KEY is not set."
}

$from = $null
$credential = $null
if (-not $DryRun -and -not $GenerateOnly) {
  $from = Read-Host "From Gmail address"
  $appPassword = Read-Host "Gmail App Password" -AsSecureString
  $credential = New-Object System.Management.Automation.PSCredential($from, $appPassword)
}

$sent = 0
$skipped = 0
$previewBlocks = New-Object System.Collections.Generic.List[string]

foreach ($row in $rows) {
  $status = [string]$row.status
  $emailsRaw = [string]$row.emails
  $generatedEmails = Get-EmailCandidates -RomajiName ([string]$row.romaji_name) -DomainInput ([string]$row.domain)

  if ($status -eq "sent" -or $status -eq $SentMarker) {
    $skipped++
    continue
  }

  $emails = @($emailsRaw -split "[`r`n;]+" | ForEach-Object { $_.Trim() } | Where-Object { Test-ValidEmail $_ })
  if ($emails.Count -eq 0) { $emails = $generatedEmails }
  if ($emails.Count -eq 0) {
    Write-Host "SKIP: no valid email -> $($row.kanji_name)"
    $skipped++
    continue
  }

  if ($GenerateOnly) {
    Write-Host "GENERATED: $($row.kanji_name) / $($emails -join ', ')"
    continue
  }

  $baseBody = Replace-TemplateTokens -Template $template -Row $row
  $body = if ($UseClaude) {
    Invoke-ClaudeEmailDraft -ApiKey $apiKey -ModelName $Model -PromptTemplate $promptTemplate -BaseDraft $baseBody -Row $row
  } else {
    $baseBody
  }

  if ($DryRun) {
    Write-Host "DRY RUN: $($row.kanji_name) / $($emails -join ', ')"
    $previewBlock = @"
NAME: $($row.kanji_name)
EMAILS: $($emails -join ', ')

$body

----------------------------------------
"@
    $previewBlocks.Add($previewBlock)
    continue
  }

  foreach ($email in $emails) {
    Send-MailMessage `
      -SmtpServer "smtp.gmail.com" `
      -Port 587 `
      -UseSsl `
      -Credential $credential `
      -From $from `
      -To $email `
      -Subject $Subject `
      -Body $body `
      -Encoding ([System.Text.Encoding]::UTF8)
  }

  Write-Host "SENT: $($row.kanji_name) / $($emails -join ', ')"
  $row.status = $SentMarker
  if ($row.PSObject.Properties.Name -contains "sent_at") {
    $row.sent_at = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  }
  $sent++
}

if (-not $DryRun -and -not $GenerateOnly) {
  $rows | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
}

if ($DryRun -and $previewBlocks.Count -gt 0) {
  $previewText = ($previewBlocks -join "`r`n")
  [System.IO.File]::WriteAllText((Resolve-Path -LiteralPath ".").Path + "\" + [System.IO.Path]::GetFileName($PreviewPath), $previewText, [System.Text.UTF8Encoding]::new($false))
  Write-Host "Preview saved: $PreviewPath"
}

Write-Host ""
Write-Host "Done"
Write-Host "Sent: $sent"
Write-Host "Skipped: $skipped"
