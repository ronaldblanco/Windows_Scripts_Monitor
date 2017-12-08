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

'Store the arguments in a variable:
Set objArgs = Wscript.Arguments
'proccessName = objArgs(0)

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Const ForReading=1
Const ForWriting=2
Set objFSO = CreateObject("Scripting.FileSystemObject")
'folder = "\\newmba\Backup_Servers\00temp\SQL_STATE_share\"
filePath = folder & "PROCESSSTATUSMDMS.asp"

Set myTemp= objFSO.OpenTextFile(filePath & ".tmp1", ForWriting, True)

'Server
myTemp.WriteLine "MonitoreDMS|"

'DISK FREE
Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='c:'")
myTemp.WriteLine Round((objLogicalDisk.FreeSpace/objLogicalDisk.Size)* 100,2) & "% Free in C|"
Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='d:'")
myTemp.WriteLine Round((objLogicalDisk.FreeSpace/objLogicalDisk.Size)* 100,2) & "% Free in D|<br>"

'FILES
'E:\0MonitoreDMS_Backup
Set colFiles = objWMIService. _
    ExecQuery("Select * from CIM_DataFile where Drive = 'e:' and Path = '\\0MonitoreDMS_Backup\\'")
For Each objFile in colFiles
	mydate = CStr(CInt(mid(mid(objFile.LastModified,1,12),5,2)) & "/"& CInt(mid(mid(objFile.LastModified,1,12),7,2)) & "/"& mid(mid(objFile.LastModified,1,12),1,4) & " " & mid(mid(objFile.LastModified,1,12),9,2) & ":" & mid(mid(objFile.LastModified,1,12),11,2))
	if mid(objFile.Name,24,35) = "monitoredms_backup.bkf" or mid(objFile.Name,24,35) = "sevastian_pc_asci.bkf" then
	    myTemp.WriteLine mydate & "->"
		myTemp.WriteLine mid(objFile.Name,24,35) & "->"
		myTemp.WriteLine Round(((objFile.FileSize/1024)/1024)/1024,2) & " GB|<br>"
	end if
Next

'FILES
'E:\0Apiserver_Backup
Set colFiles = objWMIService. _
    ExecQuery("Select * from CIM_DataFile where Drive = 'e:' and Path = '\\0Apiserver_Backup\\'")
For Each objFile in colFiles
	mydate = CStr(CInt(mid(mid(objFile.LastModified,1,12),5,2)) & "/"& CInt(mid(mid(objFile.LastModified,1,12),7,2)) & "/"& mid(mid(objFile.LastModified,1,12),1,4) & " " & mid(mid(objFile.LastModified,1,12),9,2) & ":" & mid(mid(objFile.LastModified,1,12),11,2))
	if mid(objFile.Name,22,35) = "apiserver_backup.bkf" then
		myTemp.WriteLine mydate & "->"
		myTemp.WriteLine mid(objFile.Name,22,35) & "->"
		myTemp.WriteLine Round(((objFile.FileSize/1024)/1024)/1024,2) & " GB|<br>"
	end if
Next

'Final Date
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
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFOaTgn3nYSVu
'' SIG '' 8NuXb2ab2HgKXZf9oIICGzCCAhcwggGAoAMCAQICEIig
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
'' SIG '' BAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUQedRtS0KNCxe
'' SIG '' k2UmYLxi/9VI7jIwDQYJKoZIhvcNAQEBBQAEgYCbbfER
'' SIG '' SdAGwwVkI9JXc6tN1g0TtTBwHciWLFi4CPdcgTpreM0z
'' SIG '' Qy9gLf97o9C7PiHLNw8LUqts8jbUpLlODsMB7+OQVp3i
'' SIG '' e7JGTWbXkzhQYlDkfHE4GX0fF1IQuL5tGQUW8ckVAyw7
'' SIG '' 0TTSzw11yewLVbZuTqsAokRqPwtfyZpPzg==
'' SIG '' End signature block
