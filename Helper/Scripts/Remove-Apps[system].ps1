param(
    [switch]$List
)
$classicApps = @(
    #"Microsoft OneDrive"
)
$storeApps = @(
    #"Microsoft.WindowsStore"
    "Clipchamp.Clipchamp"
    "Microsoft.3DBuilder"
    "Microsoft.549981C3F5F10" #Cortana
    "Microsoft.BingFinance"
    "Microsoft.BingFoodAndDrink"
    "Microsoft.BingHealthAndFitness"
    "Microsoft.BingNews"
    "Microsoft.BingSearch"
    "Microsoft.BingSports"
    "Microsoft.BingTranslator"
    "Microsoft.BingTravel"
    "Microsoft.BingWeather"
    "Microsoft.CommsPhone"
    "Microsoft.Edge.GameAssist"
    "Microsoft.GamingApp"
    #"Microsoft.GetHelp"
    "Microsoft.Getstarted"
    "Microsoft.Microsoft3DViewer"
    "Microsoft.MicrosoftOfficeHub"
    "Microsoft.MicrosoftSolitaireCollection"
    "Microsoft.MicrosoftStickyNotes"
    "Microsoft.MixedReality.Portal"
    #"Microsoft.MSPaint"
    "Microsoft.Office.OneNote"
    "Microsoft.OutlookForWindows"
    #"Microsoft.Paint"
    "Microsoft.People"
    "Microsoft.PowerAutomateDesktop"
    #"Microsoft.ScreenSketch" #Snip & Sketch
    "Microsoft.SkypeApp"
    "Microsoft.Todos"
    "Microsoft.Windows.DevHome"
    "Microsoft.WindowsAlarms"
    #"Microsoft.WindowsCalculator"
    "Microsoft.WindowsCamera"
    "microsoft.windowscommunicationsapps" #Mail
    #"Microsoft.Copilot"#Testing
    "Microsoft.WindowsFeedbackHub"
    "Microsoft.WindowsMaps"
    #"Microsoft.WindowsNotepad"
    "Microsoft.WindowsPhone"
    #"Microsoft.Windows.Photos"#Testing
    "Microsoft.WindowsSoundRecorder"
    "Microsoft.Xbox.TCUI"
    "Microsoft.XboxApp"
    "Microsoft.XboxGameOverlay"
    "Microsoft.XboxGamingOverlay"
    "Microsoft.XboxSpeechToTextOverlay"
    "Microsoft.YourPhone"
    "Microsoft.ZuneMusic" #Groove Music
    "Microsoft.ZuneVideo" #Movies & TV
    "MicrosoftCorporationII.MicrosoftFamily"
    #"MicrosoftCorporationII.QuickAssist"
    #"MSTeams"
)
# Combine Get-AppXProvisionedPackage and Get-AppxPackage results and then filter for the apps we want to remove
$installedApps = @(Get-AppXProvisionedPackage -Online | Select-Object -ExpandProperty DisplayName) + @(Get-AppxPackage -AllUsers | Select-Object -ExpandProperty Name)

if ($List) {
    $installedApps | Sort-Object
    exit
}

$appsToRemove = $apps | Where-Object { $installedApps -contains $_ }


# Special case for OneDrive
# If OneDrive is in the $classicApps list do this
if ($classicApps -contains "Microsoft OneDrive") {
    Write-Host "Uninstalling OneDrive"
    try {
        $oneDriveExe = Join-Path $env:WinDir 'System32\OneDriveSetup.exe'
        Write-Host "    Uninstalling OneDrive using $oneDriveExe"
        $proc = Start-Process -FilePath $oneDriveExe -ArgumentList '/uninstall' -NoNewWindow -Wait -PassThru -ErrorAction Stop
        if ($proc.ExitCode -eq 0) {
            Write-Host "    OneDrive uninstalled" -ForegroundColor Green
        }
        else {
            Write-Host "    OneDrive uninstall exited with code $($proc.ExitCode)" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "    Failed to uninstall OneDrive: $($_.Exception.Message -replace '\s+', ' ')" -ForegroundColor Red
    }
}

# Remove classic apps
foreach ($appName in $classicApps) {
    $Paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $App = Get-ItemProperty $Paths -EA SilentlyContinue | Where-Object { $_.DisplayName -match $appName } | Select-Object -First 1
    If ($App) { 
        cmd /c $App.UninstallString /passive /norestart 
    } else { 
        Write-Output "$appName is not installed." 
    }
}

# Remove Store Apps
foreach ($app in $appsToRemove) {
    Write-Host "Removing $app"

    # Remove AppxPackage
    try {
        $package = Get-AppxPackage -Name $app -AllUsers -ErrorAction Stop
    }
    catch {
        Write-Host "    Error retrieving AppxPackage: $($_.Exception.Message -replace '\s+', ' ')" -ForegroundColor Yellow
    }

    if ($package) {
        try {
            $package | Remove-AppxPackage -AllUsers -ErrorAction Stop
            Write-Host "    Removed AppxPackage" -ForegroundColor Green
        }
        catch {
            Write-Host "    Failed to remove AppxPackage: $($_.Exception.Message -replace '\s+', ' ')" -ForegroundColor Red
        }
    }
    else {
        Write-Host "    AppxPackage not found"
    }

    # Remove AppxProvisionedPackage
    try {
        $provisioned = Get-AppXProvisionedPackage -Online -ErrorAction Stop | Where-Object DisplayName -Match $app
    }
    catch {
        Write-Host "    Error retrieving AppxProvisionedPackage: $($_.Exception.Message -replace '\s+', ' ')" -ForegroundColor Yellow
    }
    
    if ($provisioned) {
        try {
            $provisioned | Remove-AppxProvisionedPackage -Online -ErrorAction Stop
            Write-Host "    Removed AppxProvisionedPackage" -ForegroundColor Green
        }
        catch {
            Write-Host "    Failed to remove AppxProvisionedPackage: $($_.Exception.Message -replace '\s+', ' ')" -ForegroundColor Red
        }
    }
    else {
        Write-Host "    AppxProvisionedPackage not found"
    }
}