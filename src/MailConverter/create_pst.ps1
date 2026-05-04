# EML to PST converter using PowerShell - Simple version
# Usage: .\create_pst.ps1 -PstPath "C:\output.pst" -EmlDir "C:\eml_files"

param(
    [Parameter(Mandatory=$true)]
    [string]$PstPath,

    [Parameter(Mandatory=$true)]
    [string]$EmlDir
)

$ErrorActionPreference = "Continue"

# Add Outlook COM
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

Write-Host "Outlook connected" -ForegroundColor Green

# Delete old PST if exists
if (Test-Path $PstPath) {
    Remove-Item $PstPath -Force
}

$PstName = [System.IO.Path]::GetFileNameWithoutExtension($PstPath)

# Create PST
Write-Host "Creating PST: $PstPath" -ForegroundColor Yellow
try {
    $namespace.Stores.AddPstStore($PstPath, $PstName)
    Write-Host "Created with AddPstStore"
} catch {
    try {
        $namespace.AddStore($PstPath)
        Write-Host "Created with AddStore"
    } catch {
        Write-Host "Failed to create PST: $_" -ForegroundColor Red
        exit 1
    }
}

Start-Sleep -Seconds 2

# Find PST folder
$pstFolder = $null
for ($i = 1; $i -le $namespace.Stores.Count; $i++) {
    $store = $namespace.Stores.Item($i)
    if ($store.FilePath -and $store.FilePath -match [regex]::Escape((Split-Path $PstPath -Leaf))) {
        foreach ($folder in $namespace.Folders) {
            if ($folder.Store -and $folder.Store.FilePath -match [regex]::Escape((Split-Path $PstPath -Leaf))) {
                $pstFolder = $folder
                break
            }
        }
    }
}

if (-not $pstFolder) {
    $pstFolder = $namespace.Folders | Where-Object { $_.Name -eq $PstName } | Select-Object -First 1
}

Write-Host "PST folder found: $($pstFolder.Name)" -ForegroundColor Green

# Get inbox for temp creation
$inbox = $namespace.GetDefaultFolder(6)

# Scan EML files
$emlFiles = Get-ChildItem -Path $EmlDir -Recurse -Filter "*.eml"
Write-Host "Found $($emlFiles.Count) EML files" -ForegroundColor Cyan

$totalEmails = 0

foreach ($emlFile in $emlFiles) {
    try {
        # Try to open EML using Outlook
        $mail = $null

        # Method 1: Try CreateItemFromTemplate
        try {
            $mail = $outlook.CreateItemFromTemplate($emlFile.FullName)
            Write-Host "Used CreateItemFromTemplate for: $($emlFile.Name)" -ForegroundColor Gray
        } catch {
            # Method 2: Create new mail and parse headers manually
            $mail = $inbox.Items.Add(0)

            # Read file content
            $content = Get-Content -Path $emlFile.FullName -Raw -Encoding UTF8

            # Extract headers using regex
            if ($content -match "(?im)^Subject:\s*(.+)$") { $mail.Subject = $matches[1].Trim() }
            if ($content -match "(?im)^From:\s*(.+)$") { $mail.Sender = $matches[1].Trim() }
            if ($content -match "(?im)^To:\s*(.+)$") { $mail.To = $matches[1].Trim() }
            if ($content -match "(?im)^Cc:\s*(.+)$") { $mail.CC = $matches[1].Trim() }

            # Extract body - find first blank line after headers
            if ($content -match "(?s)^[^\r\n]*\r?\n\r?\n(.+)$") {
                $body = $matches[1]
                if ($body -match "<html>|<body>") {
                    $mail.HTMLBody = $body
                } else {
                    $mail.Body = $body
                }
            }
        }

        if ($mail) {
            # Save and copy/move
            $mail.Save()
            $copied = $mail.Copy()
            $copied.Move($pstFolder)

            $totalEmails++

            if ($totalEmails % 10 -eq 0) {
                Write-Host "Progress: $totalEmails" -ForegroundColor Cyan
            }
        }

    } catch {
        Write-Host "Error: $($emlFile.Name) - $_" -ForegroundColor Red
        continue
    }
}

Write-Host "Total emails imported: $totalEmails" -ForegroundColor Green
Write-Host "SUCCESS!" -ForegroundColor Green
