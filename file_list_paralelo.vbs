strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Wscript.Echo date()
	
'DISK FREE
Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='c:'")
Wscript.Echo Round((objLogicalDisk.FreeSpace/objLogicalDisk.Size)* 100,2) & "% Free in C|<br>"
'Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='d:'")
'Wscript.Echo Round((objLogicalDisk.FreeSpace/objLogicalDisk.Size)* 100,2) & "% Free in D"

'FILES	
Set colFiles = objWMIService. _
    ExecQuery("Select * from CIM_DataFile where Drive = 'z:' and Path = '\\'")
For Each objFile in colFiles
	Wscript.Echo mid(objFile.LastModified,1,12) & "->"
    Wscript.Echo mid(objFile.Name,4,35) & "->"
	Wscript.Echo Round(((objFile.FileSize/1024)/1024)/1024,2) & " GB|<br>"
Next
