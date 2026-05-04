[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ProgressPreference = 'SilentlyContinue'

# Download AzCopy from GitHub releases
$url = "https://github.com/Azure/azure-storage-azcopy/releases/download/v10.25.1/azcopy_windows_amd64_10.25.1.zip"
$output = "E:\foxmailToPstfileProject\src\MailConverter\bin\Debug\net48\azcopy.zip"

Write-Host "Downloading AzCopy..."
Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing -TimeoutSec 120

if (Test-Path $output) {
    $size = (Get-Item $output).Length
    Write-Host "Downloaded: $size bytes"

    # Extract
    $extractDir = "E:\foxmailToPstfileProject\src\MailConverter\bin\Debug\net48\azcopy"
    New-Item -ItemType Directory -Force -Path $extractDir | Out-Null

    Write-Host "Extracting..."
    Expand-Archive -Path $output -DestinationPath $extractDir -Force

    # List contents
    Get-ChildItem $extractDir -Recurse | Select-Object FullName

    # Cleanup zip
    Remove-Item $output -Force

    Write-Host "Done! AzCopy extracted to: $extractDir"
} else {
    Write-Host "Download failed!"
}
