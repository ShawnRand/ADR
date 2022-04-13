$Master = Split-Path $MyInvocation.MyCommand.Definition -Leaf
$ScriptPath = Split-Path $MyInvocation.MyCommand.Definition
Get-ChildItem "$ScriptPath\*.ps1" | Where{$_.Name -ne $Master} | ForEach-Object { & $_.FullName }