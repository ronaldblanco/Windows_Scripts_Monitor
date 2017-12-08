strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

'DISK FREE
Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='c:'")
Wscript.Echo Round((objLogicalDisk.FreeSpace/objLogicalDisk.Size)* 100,2) & "% Free in C"
Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='d:'")
Wscript.Echo Round((objLogicalDisk.FreeSpace/objLogicalDisk.Size)* 100,2) & "% Free in D"

'FILES	
Set colFiles = objWMIService. _
    ExecQuery("Select * from CIM_DataFile where Drive = 'd:' and Path = '\\MBA3 - SERVIDOR v17.1\\Base de Datos\\BACKUPS\\'")
For Each objFile in colFiles
    Wscript.Echo mid(objFile.Name,48,35)
	Wscript.Echo mid(objFile.CreationDate,1,12)
	Wscript.Echo Round(((objFile.FileSize/1024)/1024)/1024,2) & " GB"
Next
