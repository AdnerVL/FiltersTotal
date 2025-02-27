@echo off
echo Generating and running filter calculation script for Excel...

:: Clear any existing temp_script.ps1
if exist temp_script.ps1 del temp_script.ps1

:: Write the PowerShell script line by line
echo # Check if ImportExcel module is installed, install if not >> temp_script.ps1
echo if (-not (Get-Module -ListAvailable -Name ImportExcel)) { >> temp_script.ps1
echo     Write-Host "Installing ImportExcel module (requires admin rights)..." >> temp_script.ps1
echo     Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force >> temp_script.ps1
echo     Install-Module -Name ImportExcel -Scope CurrentUser -Force >> temp_script.ps1
echo } >> temp_script.ps1
echo. >> temp_script.ps1
echo # Find any .xlsx file in the current directory >> temp_script.ps1
echo $excelFile = Get-ChildItem -Path "." -Filter "*.xlsx" ^| Select-Object -First 1 >> temp_script.ps1
echo if (-not $excelFile) { >> temp_script.ps1
echo     Write-Host "No .xlsx file found in the directory!" >> temp_script.ps1
echo     exit >> temp_script.ps1
echo } >> temp_script.ps1
echo Write-Host "Processing file: $($excelFile.Name)" >> temp_script.ps1
echo. >> temp_script.ps1
echo # Import the Excel file without headers, starting at B2 >> temp_script.ps1
echo $filters = Import-Excel -Path $excelFile.FullName -WorksheetName "Sheet1" -NoHeader -StartRow 2 >> temp_script.ps1
echo. >> temp_script.ps1
echo # Rename columns: P2 is Size (Column B), P3 is Quantity (Column C) >> temp_script.ps1
echo $filters = $filters ^| Select-Object @{Name='Size';Expression={$_.P2}}, @{Name='Quantity';Expression={[int]$_.P3}} >> temp_script.ps1
echo. >> temp_script.ps1
echo # Group by Size and sum the Quantities >> temp_script.ps1
echo $results = $filters ^| Where-Object { $_.Size -ne $null } ^| Group-Object -Property Size ^| ForEach-Object { >> temp_script.ps1
echo     [PSCustomObject]@{ >> temp_script.ps1
echo         Size = $_.Name >> temp_script.ps1
echo         TotalQuantity = ($_.Group ^| Measure-Object -Property Quantity -Sum).Sum >> temp_script.ps1
echo     } >> temp_script.ps1
echo } >> temp_script.ps1
echo. >> temp_script.ps1
echo # Display the results on screen >> temp_script.ps1
echo $results ^| Format-Table -AutoSize >> temp_script.ps1
echo # Export results to a CSV file >> temp_script.ps1
echo $results ^| Export-Csv -Path "filter_totals.csv" -NoTypeInformation >> temp_script.ps1
echo Write-Host "Results saved to filter_totals.csv in the current directory." >> temp_script.ps1
echo Write-Host "Done! These are the total quantities for each filter size." >> temp_script.ps1

:: Run the script
powershell.exe -ExecutionPolicy Bypass -File "temp_script.ps1"

:: Clean up
del temp_script.ps1

echo Process complete!
pause