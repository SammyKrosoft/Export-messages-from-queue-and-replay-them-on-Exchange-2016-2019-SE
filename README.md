This PowerShell script is designed to export all messages from the Poison Message Queue on an Exchange server and save them as .eml files in a specified directory.

This script:
- Iterates through all messages in the Poison queue on MyServer.
- Logs their subject and size.
- Exports each message as an .eml file to D:\MsgBackups. Note that the directory is stored in the $ExportDirectory variable.

> **Use case**: Typically used by Exchange administrators to inspect or recover messages stuck in the Poison queue, which usually contains problematic or potentially harmful messages that couldn’t be processed.

```PowerShell
$count=0
$ExportDirectory = "D:\MsgBackups"
$ServerName = "MyServer"

Get-Message -Queue "$($ServerName)\Poison" -ResultSize Unlimited | ForEach-Object {
  $count++
  Write-host "Processing message $($count) $($_.Subject.ToString()) and size is $($_.Size.ToString())" -foregroundcolor Green
  $Temp = $ExportDirectory + $_.InternetMessageID + ".eml"
  $Temp = $Temp.Replace("<", "_").Replace(">", "_")
  Export-Message $_.Identity | AssembleMessage -Path $Temp
}
```
