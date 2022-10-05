$MyInvocation.MyCommand.Path | Split-Path | Push-Location

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
        Start-Process reg -ArgumentList "import `"$($Item.FullName)`""
    }else{
        Start-Process $Item.FullName
    }

    Start-Sleep 1
}

"Complete, displaying prompt..."

Start-Sleep 2

$Wscript_Shell = New-Object -ComObject "Wscript.Shell"
$MsgBox = $Wscript_Shell.Popup("Logon scripts finished, removing scripts in 60 seconds, press ok to remove now",60,"Setup Helper",4+32)
$MsgBox
switch  ($MsgBox) {
    {"6", "-1" -contains $_} { 
        "Yes"
        Remove-Item -Force -Confirm:$false * 
    }
    default { 
        "Default"
        Exit 
    }
}