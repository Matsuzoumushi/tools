#This code saves the path of the open explorer.
#The Created batch file reopens saved explorer path.
#Batch file will be created in $OutputFolder.

$OutputFolder =".\"

$shell=New-Object -ComObject shell.application
$Explorer = $shell.windows()|Where-Object{$_.fullname -like "*Explorer.EXE"}
$Output = New-Item -path ($OutputFolder+"BackupExplorer_"+(get-date -Format "yyyyMMddHHmmss")+".bat")

if ($Explorer.Count -eq 0){
Write-Output "nothing"
}else{
Write-Output "------------------saved path------------------"
foreach($Fullpath in $Explorer){
Add-Content -value ("explorer """+$Fullpath.locationURL + """") -path $Output
Write-Output $Fullpath.locationURL
} 
}
pause