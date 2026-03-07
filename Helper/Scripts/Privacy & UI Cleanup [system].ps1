# This is an example of a script that will run in the specialize pass because it has [system] in the file name (running as system, after setup but before logon)
# The specific changes in this example, try to reduce bloat or UI clutter

$DefaultUserHive = "HKLM\default-user"
reg.exe load $DefaultUserHive "C:\Users\Default\NTUSER.DAT"

$cdmPath = "$DefaultUserHive\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
$cdmValues = @{
    ContentDeliveryAllowed          = 0
    FeatureManagementEnabled        = 0
    OEMPreInstalledAppsEnabled      = 0
    PreInstalledAppsEnabled         = 0
    PreInstalledAppsEverEnabled     = 0
    SilentInstalledAppsEnabled      = 0
    SoftLandingEnabled              = 0
    SubscribedContentEnabled        = 0
    "SubscribedContent-310093Enabled" = 0
    "SubscribedContent-338387Enabled" = 0
    "SubscribedContent-338388Enabled" = 0
    "SubscribedContent-338389Enabled" = 0
    "SubscribedContent-338393Enabled" = 0
    "SubscribedContent-353698Enabled" = 0
    SystemPaneSuggestionsEnabled    = 0
}
foreach ($name in $cdmValues.Keys) { Set-ItemProperty -Path $cdmPath -Name $name -Value $cdmValues[$name] -Type DWord -Force }

$explorerPath = "$DefaultUserHive\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$cdmValues = @{
    ShowSyncProviderNotifications = 0
    ShowMeTheWindowsWelcomeExperience = 0
    ShowWindowsTips = 0
    ShowTaskViewButton = 0
    TaskbarDa = 0
    ShowCopilotButton = 0
    HideFileExt = 0
}
foreach ($name in $cdmValues.Keys) { Set-ItemProperty -Path $explorerPath -Name $name -Value $cdmValues[$name] -Type DWord -Force }

$runPath = "$DefaultUserHive\Software\Microsoft\Windows\CurrentVersion\Run"
Remove-ItemProperty -Path $runPath -Name OneDriveSetup -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $runPath -Name OneDrive -ErrorAction SilentlyContinue

reg.exe unload $DefaultUserHive

$cloudPath = "HKLM:\Software\Policies\Microsoft\Windows\CloudContent"
New-Item -Path $cloudPath -Force | Out-Null
Set-ItemProperty -Path $cloudPath -Name DisableWindowsConsumerFeatures -Value 1 -Type DWord -Force