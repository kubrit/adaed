command = "powershell.exe -executionpolicy bypass -WindowStyle Hidden -file \\example.domain\files\repo\scripts\ADAED\adaed.ps1"
set shell = CreateObject("WScript.Shell")
shell.Run command,0