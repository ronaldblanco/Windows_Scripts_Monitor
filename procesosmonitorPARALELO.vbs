'#############################################################################################
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Name: Include Code from another file
' By: Greg Upton
' Date: 06/14/09
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Include("\\Server\Share\File") ' Path to code file

Sub Include(sInstFile)
	Dim f, s, oFSO
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	If oFSO.FileExists(sInstFile) Then
		Set f = oFSO.OpenTextFile(sInstFile)
		s = f.ReadAll
		f.Close
		ExecuteGlobal s
	End If
	On Error Goto 0
	Set f = Nothing
	Set oFSO = Nothing
End Sub
'####################################################################################################

Include("env.vbs") ' Path to code file

strComputer = computer

' Store the arguments in a variable:
Set objArgs = Wscript.Arguments

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process where Name='4D Server.exe'")

Const ForReading=1
Const ForWriting=2
Set objFSO = CreateObject("Scripting.FileSystemObject")
'folder = "\\newmba\Backup_Servers\00temp\SQL_STATE_share\"
filePath = folder & "PROCESSSTATUSPARALELO.asp"

Set myTemp= objFSO.OpenTextFile(filePath & ".tmp1", ForWriting, True)

myTemp.WriteLine "PARALELO|"

If colProcessList.Count = 0 Then
    myTemp.WriteLine "Process is not running.|"
Else
    myTemp.WriteLine "Process is running.|"
End If

For Each objProcess in colProcessList
    colProperties = objProcess.GetOwner(strNameOfUser,strUserDomain)
	myTemp.WriteLine strUserDomain & "\" & strNameOfUser & ".|<br>"
Next

For Each objProcess in colProcessList
  if  objProcess.Name = void then
	myTemp.WriteLine "Process is stoped|"
  end if
  myTemp.WriteLine "Name: " & objProcess.Name &"|"
  myTemp.WriteLine "ID: " & objProcess.ProcessID &"|"
  myTemp.WriteLine "RAM Used: " & Round(((objProcess.WorkingSetSize /1024)/1024)/1024,2) & " GB" &"|<br>"
Next

'DISK FREE
Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='c:'")
myTemp.WriteLine Round((objLogicalDisk.FreeSpace/objLogicalDisk.Size)* 100,2) & "% Free in C|<br>"

'FILES	
Set colFiles = objWMIService. _
    ExecQuery("Select * from CIM_DataFile where Drive = 'z:' and Path = '\\'")
For Each objFile in colFiles
	mydate = CStr(CInt(mid(mid(objFile.LastModified,1,12),5,2)) & "/"& CInt(mid(mid(objFile.LastModified,1,12),7,2)) & "/"& mid(mid(objFile.LastModified,1,12),1,4) & " " & mid(mid(objFile.LastModified,1,12),9,2) & ":" & mid(mid(objFile.LastModified,1,12),11,2))
	if mid(mydate,1,9) = mid(CStr(now()),1,9) then
		myTemp.WriteLine CInt(mid(mid(objFile.LastModified,1,12),5,2)) & "/"& CInt(mid(mid(objFile.LastModified,1,12),7,2)) & "/"& mid(mid(objFile.LastModified,1,12),1,4) & " " & mid(mid(objFile.LastModified,1,12),9,2) & ":" & mid(mid(objFile.LastModified,1,12),11,2) & "->"
		myTemp.WriteLine mid(objFile.Name,4,35) & "->"
		myTemp.WriteLine Round(((objFile.FileSize/1024)/1024)/1024,2) & " GB|<br>"
	end if
Next

myTemp.WriteLine now()&"|"

myTemp.Close
If (objFSO.FileExists(filePath)) Then
  objFSO.DeleteFile(filePath)
end if
objFSO.MoveFile filePath&".tmp1", filePath

wscript.quit (0)	
'' SIG '' Begin signature block
'' SIG '' MIID8wYJKoZIhvcNAQcCoIID5DCCA+ACAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFI0MmvJ8XuhO
'' SIG '' Ws8ZqIffzsV/Y6j7oIICGzCCAhcwggGAoAMCAQICEIig
'' SIG '' ohjssziCQLgC4RnUrh0wDQYJKoZIhvcNAQEFBQAwGDEW
'' SIG '' MBQGA1UEAxMNUm9uYWxkIEJsYW5jbzAeFw0xNzAxMDEw
'' SIG '' NTAwMDBaFw0yMzAxMDEwNTAwMDBaMBgxFjAUBgNVBAMT
'' SIG '' DVJvbmFsZCBCbGFuY28wgZ8wDQYJKoZIhvcNAQEBBQAD
'' SIG '' gY0AMIGJAoGBANuG7UVFH9ARXhRqrHl+Z8XGqD7Are2u
'' SIG '' x2ksTzsQtVsfF4c/smclBWvYy/tgRx3N41mYLJtlYmGt
'' SIG '' LuqSa3rgJNf2MCsGzCm2+L/pXet1DnJ8sKTiED/9ZXtF
'' SIG '' pdSd/i3jyxVTlLj9hNaEvTqx6QnQQ9qe0ww4EDLAWbu3
'' SIG '' tqR/IooTAgMBAAGjYjBgMBMGA1UdJQQMMAoGCCsGAQUF
'' SIG '' BwMDMEkGA1UdAQRCMECAEOsjZ+Gf5VyZWesnccFYdw+h
'' SIG '' GjAYMRYwFAYDVQQDEw1Sb25hbGQgQmxhbmNvghCIoKIY
'' SIG '' 7LM4gkC4AuEZ1K4dMA0GCSqGSIb3DQEBBQUAA4GBAH5t
'' SIG '' bFvwi+ZCAcxrs8KhHGxE5O4jpJDgiVmkY4YnHtdZeRU1
'' SIG '' 8TKMZjCp10mwb8MH3+C109JGzrnNNNY1bHnKbBxJG9nl
'' SIG '' GpX0d9EYCt6aUMUB4Z5WAQ3aWTMGOiJ89B/Gal7Rb9pt
'' SIG '' M4DbmYYnaExGVWrU1Hc17euRGp/sbQcd18czMYIBRDCC
'' SIG '' AUACAQEwLDAYMRYwFAYDVQQDEw1Sb25hbGQgQmxhbmNv
'' SIG '' AhCIoKIY7LM4gkC4AuEZ1K4dMAkGBSsOAwIaBQCgcDAQ
'' SIG '' BgorBgEEAYI3AgEMMQIwADAZBgkqhkiG9w0BCQMxDAYK
'' SIG '' KwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYB
'' SIG '' BAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUSdX4iIjokpfd
'' SIG '' 8Pz7YxAmOWzJgj0wDQYJKoZIhvcNAQEBBQAEgYB/rNJc
'' SIG '' 2ocXZih6nrJ9ApuK8F0hkW1onrqHx/WhDX17Cb61VqIb
'' SIG '' fqo+F8VVXxGRBAXWiGt4z5IaDLIU4lTGOjP9iuDe+aRd
'' SIG '' 67zbfGFJXD/dAXKJuA/keK0wjRkljFhO2OEn6CPzO3Hg
'' SIG '' YdTO58lbbyc0ZtgcFo5n8gRTH0q5MIWztA==
'' SIG '' End signature block
