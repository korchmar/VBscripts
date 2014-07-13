strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")

	Set objShell = Wscript.CreateObject("Wscript.Shell")
  Set objEnv = objShell.Environment("Process")

On Error Resume Next

  do while  1 > 0


 WScript.sleep 2000

Set colProcesses3 = objWMIService.ExecQuery( _
    "Select ProcessId From Win32_Process " _
    & "Where Name = 'IEDriverServer.exe'")
If colProcesses3.Count >0  Then  

Set colProcesses2 = objWMIService.ExecQuery( _
    "Select ProcessId From Win32_Process " _
    & "Where Name = 'iexplore.exe'")
  If colProcesses2.Count < 2  Then 
  WScript.sleep 2000
  Set colProcesses2 = objWMIService.ExecQuery( _
    "Select ProcessId From Win32_Process " _
    & "Where Name = 'iexplore.exe'")
   If colProcesses2.Count < 2  Then 
     For Each objProcess in colProcesses3
      objProcess.Terminate()
     next
   End if

  End If

End if

loop