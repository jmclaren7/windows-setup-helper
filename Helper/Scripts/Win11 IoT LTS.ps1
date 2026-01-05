$progressPreference = 'silentlyContinue'
Write-Host "Installing WinGet PowerShell module from PSGallery..."
Install-PackageProvider -Name NuGet -Force | Out-Null
Install-Module -Name Microsoft.WinGet.Client -Force -Repository PSGallery | Out-Null
Write-Host "Using Repair-WinGetPackageManager cmdlet to bootstrap WinGet..."
Repair-WinGetPackageManager
Write-Host "Installing applications via WinGet..."
$Apps = "Microsoft.WindowsTerminal,Microsoft.PowerShell"
$Apps -split ',' | ForEach-Object {"installing: $_"; winget install --scope Machine --accept-package-agreements --accept-source-agreements --silent -e --id $_.Trim()}