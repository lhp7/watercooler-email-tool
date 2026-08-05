param(
    [string]$TargetMailbox = ""
)

$ErrorActionPreference = "Stop"
$packageDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifestPath = Join-Path $packageDirectory "drafts.json"
$logPath = Join-Path $packageDirectory "outlook-draft-import-log.csv"
$draftKeyProperty = "WaterCoolerDraftKey"
$batchProperty = "WaterCoolerBatchId"
$olFolderDrafts = 16
$olMailItem = "IPM.Note"
$olFormatHTML = 2
$olByValue = 1
$olTo = 1
$olCc = 2
$olBcc = 3

function Release-ComObjectSafely {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Add-OutlookRecipient {
    param(
        [object]$Message,
        [string]$Address,
        [int]$RecipientType
    )
    if ([string]::IsNullOrWhiteSpace($Address)) { return }
    $recipient = $null
    try {
        $recipient = $Message.Recipients.Add($Address.Trim())
        $recipient.Type = $RecipientType
        [void]$recipient.Resolve()
        if (-not $recipient.Resolved) {
            throw "Outlook could not resolve recipient address: $Address"
        }
    }
    finally {
        Release-ComObjectSafely $recipient
    }
}

if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    throw "drafts.json was not found. Keep every extracted file in the same folder."
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]::IsNullOrWhiteSpace($TargetMailbox)) {
    $TargetMailbox = [string]$manifest.target_mailbox
}
if ([string]::IsNullOrWhiteSpace($TargetMailbox)) {
    throw "No target Outlook mailbox is configured."
}

$drafts = @($manifest.drafts)
if ($drafts.Count -eq 0) {
    throw "This package does not contain any drafts."
}

foreach ($draft in $drafts) {
    if ([string]::IsNullOrWhiteSpace([string]$draft.to)) {
        throw "A packaged draft is missing its recipient email address."
    }
    $attachmentPath = Join-Path $packageDirectory (([string]$draft.attachment) -replace '/', [IO.Path]::DirectorySeparatorChar)
    if (-not (Test-Path -LiteralPath $attachmentPath -PathType Leaf)) {
        throw "Missing PDF attachment: $attachmentPath"
    }
}

Write-Host ""
Write-Host "Water Cooler Outlook Draft Importer" -ForegroundColor Cyan
Write-Host "Mailbox: $TargetMailbox"
Write-Host "Drafts in this package: $($drafts.Count)"
Write-Host ""
Write-Host "This creates drafts only. It never sends email." -ForegroundColor Yellow
$confirmation = Read-Host "Type CREATE to continue"
if ($confirmation -cne "CREATE") {
    Write-Host "Cancelled. Nothing was changed."
    exit 0
}

$outlook = $null
$session = $null
$account = $null
$store = $null
$draftsFolder = $null
$items = $null
$existingKeys = @{}
$results = [System.Collections.Generic.List[object]]::new()

try {
    $outlook = New-Object -ComObject Outlook.Application
    $session = $outlook.Session

    foreach ($candidate in @($session.Accounts)) {
        if ([string]$candidate.SmtpAddress -ieq $TargetMailbox) {
            $account = $candidate
            break
        }
        Release-ComObjectSafely $candidate
    }

    if ($null -eq $account) {
        $available = @($session.Accounts | ForEach-Object { [string]$_.SmtpAddress }) -join ", "
        throw "Outlook is not signed into $TargetMailbox. Available accounts: $available"
    }

    $store = $account.DeliveryStore
    $draftsFolder = $store.GetDefaultFolder($olFolderDrafts)
    $items = $draftsFolder.Items

    for ($index = 1; $index -le $items.Count; $index++) {
        $existingItem = $null
        $existingProperty = $null
        try {
            $existingItem = $items.Item($index)
            $existingProperty = $existingItem.UserProperties.Find($draftKeyProperty)
            if ($null -ne $existingProperty -and -not [string]::IsNullOrWhiteSpace([string]$existingProperty.Value)) {
                $existingKeys[[string]$existingProperty.Value] = $true
            }
        }
        catch {
            # Ignore unrelated Outlook items that do not expose user properties.
        }
        finally {
            Release-ComObjectSafely $existingProperty
            Release-ComObjectSafely $existingItem
        }
    }

    $created = 0
    $skipped = 0
    $failed = 0
    foreach ($draft in $drafts) {
        $draftKey = [string]$draft.draft_key
        if ($existingKeys.ContainsKey($draftKey)) {
            $skipped++
            $results.Add([pscustomobject]@{
                Recipient = [string]$draft.to
                Subject = [string]$draft.subject
                Status = "Skipped - already created"
                Detail = ""
            })
            continue
        }

        $message = $null
        $pdfAttachment = $null
        $logoAttachment = $null
        $draftKeyUserProperty = $null
        $batchUserProperty = $null
        try {
            $message = $draftsFolder.Items.Add($olMailItem)
            $message.Subject = [string]$draft.subject
            $message.BodyFormat = $olFormatHTML
            $message.HTMLBody = [string]$draft.html_body
            $message.SendUsingAccount = $account

            Add-OutlookRecipient -Message $message -Address ([string]$draft.to) -RecipientType $olTo
            foreach ($address in @($draft.cc)) {
                Add-OutlookRecipient -Message $message -Address ([string]$address) -RecipientType $olCc
            }
            foreach ($address in @($draft.bcc)) {
                Add-OutlookRecipient -Message $message -Address ([string]$address) -RecipientType $olBcc
            }

            $attachmentPath = Join-Path $packageDirectory (([string]$draft.attachment) -replace '/', [IO.Path]::DirectorySeparatorChar)
            $pdfAttachment = $message.Attachments.Add(
                $attachmentPath,
                $olByValue,
                1,
                [string]$draft.attachment_display_name
            )

            $logoPath = Join-Path $packageDirectory "signature_logo.png"
            if (Test-Path -LiteralPath $logoPath -PathType Leaf) {
                $logoAttachment = $message.Attachments.Add($logoPath, $olByValue, 0, "The Dallas Foundation")
                $logoAttachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                    "water-cooler-signature-logo"
                )
                $logoAttachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x370E001F",
                    "image/png"
                )
                $logoAttachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
                    $true
                )
            }

            $draftKeyUserProperty = $message.UserProperties.Add($draftKeyProperty, 1, $true)
            $draftKeyUserProperty.Value = $draftKey
            $batchUserProperty = $message.UserProperties.Add($batchProperty, 1, $true)
            $batchUserProperty.Value = [string]$manifest.batch_id

            $message.Save()
            $created++
            $existingKeys[$draftKey] = $true
            $results.Add([pscustomobject]@{
                Recipient = [string]$draft.to
                Subject = [string]$draft.subject
                Status = "Draft created"
                Detail = ""
            })
        }
        catch {
            $failed++
            $results.Add([pscustomobject]@{
                Recipient = [string]$draft.to
                Subject = [string]$draft.subject
                Status = "Failed"
                Detail = $_.Exception.Message
            })
        }
        finally {
            Release-ComObjectSafely $batchUserProperty
            Release-ComObjectSafely $draftKeyUserProperty
            Release-ComObjectSafely $logoAttachment
            Release-ComObjectSafely $pdfAttachment
            Release-ComObjectSafely $message
        }
    }

    $results | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "Finished." -ForegroundColor Green
    Write-Host "Created: $created"
    Write-Host "Already existed: $skipped"
    Write-Host "Failed: $failed"
    Write-Host "Log: $logPath"
    Write-Host ""
    Write-Host "Open Drafts under $TargetMailbox in Outlook. The drafts will sync to Outlook on the web." -ForegroundColor Cyan
}
finally {
    Release-ComObjectSafely $items
    Release-ComObjectSafely $draftsFolder
    Release-ComObjectSafely $store
    Release-ComObjectSafely $account
    Release-ComObjectSafely $session
    Release-ComObjectSafely $outlook
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
