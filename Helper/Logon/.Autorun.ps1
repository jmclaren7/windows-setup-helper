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
$MsgBox = $Wscript_Shell.Popup("Logon scripts finished, removing scripts in 2 minutes, press 'yes' to remove now or 'no' to cancel automatic removal.", 120, "Setup Helper", 1+32)

If ($MsgBox -ne 1 -AND $MsgBox -ne -1) { Exit }

"Deleting script folder"
Set-Location ..
Remove-Item -LiteralPath $(Split-Path -Parent $MyInvocation.MyCommand.Definition) -Recurse -Force

if(-Not $?){
    Powershell.exe -NoLogo
}