#asking for PC Name to Query at beginning
Write-host "Please enter the PC Name to be queried" -ForegroundColor Yellow
$PCNAME = Read-Host "Please enter the PC Name"
Write-Host "Thank You, Proceeding with the script" -ForegroundColor Green 

#Executing turning on the Remote Protocol on the PC
Write-Host "Turning on Remote Protocol" -ForegroundColor Yellow
PsExec.exe /accepteula \\$PCNAME -h -s powershell.exe Enable-PSRemoting -Force
Write-Host "Turning on Remote Protocol Completed, Proceeding with the remainder of the script" -ForegroundColor Green

#command will force gpupdate using PsExec
Write-Host "Running GPupdate Forced" -ForegroundColor Yellow
PsExec.exe /accepteula \\$PCNAME gpupdate /force
Write-Host "GPupdate Forced Completed" -ForegroundColor Green

#turning off remote desktop protocol
Write-Host "Turning off Remote Protocol" -ForegroundColor Yellow
PsExec.exe /accepteula \\$PCNAME -h -s powershell.exe Disable-PSRemoting -Force
Write-Host "Turning off Remote Protocol Completed" -ForegroundColor Green 