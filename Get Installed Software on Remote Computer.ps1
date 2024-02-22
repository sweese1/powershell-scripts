#asking for PC Name to Query at beginning
Write-host "Please enter the PC Name to be queried" -ForegroundColor Yellow
$PCNAME = Read-Host "Please enter the PC Name"

#starting Transcript and logs
Write-Host "Testing Path for Log Files" -ForegroundColor Yellow

#Testing Path for the following TXT file of installed Software
IF(!(test-path -path "C:\IT"))
{New-Item -Path "C:\IT" -itemType Directory -Force} ##Create the path if it doesn't exist
Write-Host "Testing Path for Installed Software File, Please be Patient" -ForegroundColor Yellow

#Creating path for File and Logging
IF(!(test-path -path "C:\IT\$PCNAME"))
{New-Item -Path "C:\IT\$PCNAME" -itemType Directory -Force} ##Create the path if it doesn't exist
Write-Host "Folder Path has been created" -ForegroundColor Green

#Testing Path for log files
IF(!(test-path -path "C:\IT\$PCNAME"))
{New-Item -Path "C:\IT\$PCNAME" -itemType Directory -Force} ##Create the path if it doesn't exist
Write-Host "Testing Path for Installed Software File, Please be Patient" -ForegroundColor Yellow

Write-Host "All File Paths have been created successfully" -ForegroundColor Green

Start-Sleep 3

Start-Transcript -Path "C:\IT\$PCNAME\log.txt"

#Defining Export Path via Variable
$ExportCSV = "C:\IT\$PCNAME\InstalledSoftware_$PCNAME.txt"
Write-Host "Path is now created in C:\IT" -ForegroundColor Green

Start-Sleep 3

#Executing turning on the Remote Protocol on the PC
Write-Host "Turning on Remote Protocol" -ForegroundColor Yellow
PsExec.exe /accepteula \\$PCNAME -h -s powershell.exe Enable-PSRemoting -Force
Write-Host "Turning on Remote Protocol Completed, Proceeding with the remainder of the script" -ForegroundColor Green

Write-Host "Running the Software Report Please be Patient!" -ForegroundColor Green

Get-CimInstance Win32_Product -ComputerName $PCNAME | ft name,version,vendor,packagename > $ExportCSV

#turning of remote desktop protocol
Write-Host "Turning off Remote Protocol" -ForegroundColor Yellow
PsExec.exe /accepteula \\$PCNAME -h -s powershell.exe Disable-PSRemoting -Force
Write-Host "Turning off Remote Protocol Completed" -ForegroundColor Green 

#Open output file after execution 
    Write-Host `n"Script executed successfully"
    if((Test-Path -Path $ExportCSV) -eq "True")
    {
        Write-Host `n" Detailed report available in:" -NoNewline -ForegroundColor Green
		Write-Host $ExportCSV 
        $Prompt = New-Object -ComObject wscript.shell  
        $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)  
        If ($UserInput -eq 6)  
        {  
            Invoke-Item "$ExportCSV" 
            
        } 
        }
    Else
    {
        Write-Host `n"No group found" -ForegroundColor Red
        
    }

    Stop-Transcript