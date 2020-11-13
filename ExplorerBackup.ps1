#This program saves paths of open explorers.
#and batch file will be created in $OutputFolder.
#The Created batch file reopens saved explorer paths.


$OutputFolder = (split-path -parent $MyInvocation.MyCommand.Definition) + "\"
#$OutputFolder = $PSScriptRoot + "\" Not executable before Powershell Version 2

$shell=New-Object -ComObject shell.application
$Explorer = $shell.windows()|Where-Object{$_.fullname -like "*Explorer.EXE"}
$Output = New-Item -path ($OutputFolder+"BackupExplorer_"+(get-date -Format "yyyyMMddHHmmss")+".bat")

if ($Explorer.Count -eq 0){
  Write-Output "nothing"
  Remove-Item $Output
}else{
  Write-Output "------------------saved paths------------------"
  foreach($Fullpath in $Explorer){
    Add-Content -value ("explorer """+$Fullpath.locationURL + """") -path $Output
    Write-Output $Fullpath.locationURL
  } 
  Write-Output "------------------create batch------------------"
  Write-Output $Output.FullName `r`n
}
pause
