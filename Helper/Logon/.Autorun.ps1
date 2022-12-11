$MyInvocation.MyCommand.Path | Split-Path | Push-Location

"Processing logon scripts..."

Start-Sleep 2

$Like = @(
    '*.bat',
    '*.ps1',
    '*.exe',
    '*.reg',
    '*.cmd')
$NotLike = @(
    '.*')

$Run = @()

foreach($File in Get-ChildItem){
    $Add = $False

    foreach($pattern in $Like){
        if($File.Name -like $pattern){ 
            $Add = $True
            Break
        }
    }

    foreach($pattern in $NotLike){
        if($File.Name -like $pattern){ 
            $Add = $False 
            Break
        }
    }

    If ($Add){ 
        $Run += $File.FullName 
    }
}

$Run | ForEach-Object {
    $Item = Get-ChildItem $_
    $Item.Name

    if($Item.Extension -eq ".reg"){
        Start-Process reg.exe -ArgumentList "import `"$($Item.FullName)`""
    }elseif($Item.Extension -eq ".ps1"){
        Start-Process powershell.exe -ArgumentList "-file `"$($Item.FullName)`""
    }else{
        Start-Process $Item.FullName
    }

    Start-Sleep 1
}

"Complete, displaying prompt..."

Start-Sleep 2

$Wscript_Shell = New-Object -ComObject "Wscript.Shell"
$MsgBox = $Wscript_Shell.Popup("Logon scripts finished, removing scripts in 60 seconds, press 'yes' to remove now or 'no' to cancel automatic removal",60,"Setup Helper",4+32)

switch  ($MsgBox) {
    {"6", "-1" -contains $_} { 
        "Deleting script folder"
        #Remove-Item -Force -Confirm:$false * 
		Set-Location ..
		Remove-Item -LiteralPath $(Split-Path -Parent $MyInvocation.MyCommand.Definition) -Recurse -Force
    }
    default { 
        "No or timeout, leaving scripts in place"
        Exit 
    }
}