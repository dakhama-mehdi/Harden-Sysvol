PowerShell Execution Errors
When running HardenSysvol, you may encounter execution errors due to PowerShell's security settings that block scripts from running. To bypass these restrictions, use one of the following options:

Option 1: Use the -executionpolicy Flag
Run the following command to invoke HardenSysvol with the Bypass execution policy:

```powershell
powershell.exe -executionpolicy bypass invoke-hardensysvol
````

Option 2: Set the Execution Policy for the Current Session
Alternatively, open PowerShell and run this command to set the execution policy to Bypass for the session:

powershell.exe -ExecutionPolicy Bypass
Then, run the HardenSysvol script:

Invoke-HardenSysvol
