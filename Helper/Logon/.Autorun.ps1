$MyInvocation.MyCommand.Path | Split-Path | Push-Location

# Log function
Function _Log {
    Param ([string]$LogString)
    $Message = "$((Get-Date).toString("yyyy-MM-dd HH:mm:ss"))> $LogString"
    $Message
    Add-content "$($MyInvocation.MyCommand.Name).log" -value $Message
}

_Log("Processing logon scripts...")
Start-Sleep 2

# Safety check to make sure we're running on a new OS install, if not don't run anything or make changes
$os = Get-WmiObject -Class Win32_OperatingSystem
$installAge = [DateTime]::Now - $os.ConvertToDateTime($os.InstallDate)
$Execute = $true 
if ($installAge.Days -gt 1) { 
    $Execute = $false
    _Log("System is not new, skipping execution...")
}

# Set the file types to run
$FileTypes = ('*.bat', '*.ps1', '*.exe', '*.reg', '*.cmd')

# Get a list of files from the working directory
$Run = Get-ChildItem -Include $FileTypes -Exclude '.*' -Recurse -Depth 0

# Filter out system files unless specified
if ($args[0] -eq 'system') {
    $Run = $Run | Where-Object { $_.Name.Substring($_.Name.Length - 9) -like '*.sys.*' }
}
else {
    $Run = $Run | Where-Object { $_.Name.Substring($_.Name.Length - 9) -notlike '*.sys.*' }
}

# Run each file in the list
$Run | ForEach-Object {
    _Log($_.Name)
    
    Start-Sleep 2

    # Wait for Windows Installer to be available
    for ($i = 0; $i -lt 12; $i++) {
        try {
            $Mutex = [System.Threading.Mutex]::OpenExisting("Global\_MSIExecute")
            $Mutex.Dispose()
            _Log("Windows Installer is busy, waiting...")
        }
        catch {
            Start-Sleep 2
            Break
        }
		
        Start-Sleep 4
    }

    # Run the file
    if ($Execute) { 
        if ($_.Extension -eq ".reg") {
            $Proc = Start-Process reg.exe -ArgumentList "import `"$($_.FullName)`"" -PassThru
        }
        elseif ($_.Extension -eq ".ps1") {
            $Proc = Start-Process powershell.exe -ArgumentList "-file `"$($_.FullName)`"" -PassThru
        }
        else {
            $Proc = Start-Process $_.FullName -PassThru
        }

        $Proc | Wait-Process -Timeout 20 -ErrorAction SilentlyContinue
        if ($Proc.HasExited -eq $False) {
            _Log("Process still running, continuing anyway...")
        }
    }
}

_Log("Complete, displaying prompt...")

Start-Sleep 2

# Display a message box, if the user doesn't click 'cancel' the scripts will be removed automatically in 2 minutes
# 1=Ok/Cancel 32=Question Mark 4096=Always on top (Undocumented?)
$Wscript_Shell = New-Object -ComObject "Wscript.Shell"
$MsgBox = $Wscript_Shell.Popup("Logon scripts finished, removing scripts in 2 minutes, press 'ok' to remove now or 'cancel' to stop automatic removal.", 120, "Setup Helper", 1 + 32 + 4096)

if ($MsgBox -eq 1 -or $MsgBox -eq -1) { 
    _Log("Deleting script folder")

    # Change working directory so the folder can be deleted
    Set-Location ..

    If ($Execute) { 
        Remove-Item -LiteralPath $(Split-Path -Parent $MyInvocation.MyCommand.Definition) -Recurse -Force -ErrorAction SilentlyContinue 

        # Drop to powershell prompt
        Powershell.exe -NoLogo
    }

   
}

