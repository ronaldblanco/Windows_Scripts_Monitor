'To run a VBScript from within another VBScript:
'32 bit

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

folderl = folder

Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

myloop = 0
do

' netstat.cmd must be on the local server to aboit publisher problems.
Return = objShell.Run("C:\Users\Administrator\Desktop\" & "netstat.cmd " & folderl & "shareSQL_STATE\",0)
WScript.Sleep 150
objShell.Run wscriptroute & folderl & "netstat_MORE.vbs " & folderl & "shareSQL_STATE\ 19812"
WScript.Sleep 50
objShell.Run wscriptroute & folderl & "netstat.vbs " & folderl & "shareSQL_STATE\ 19812"
WScript.Sleep 6750
objShell.Run wscriptroute & folderl & "procesosmonitor.vbs " & " "
objShell.Run wscriptroute & folderl & "procesosmonitorClient.vbs " & " "

myloop = myloop + 1
if myloop = 100 then myloop = 1
loop until myloop = 999999
'' SIG '' Begin signature block
'' SIG '' MIID8wYJKoZIhvcNAQcCoIID5DCCA+ACAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFKfrGYv21JL+
'' SIG '' EDyQfO2Byf5Kcm48oIICGzCCAhcwggGAoAMCAQICEIig
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
'' SIG '' BAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU4eamNFm9OKNX
'' SIG '' t5UIWUW5FlJXcMcwDQYJKoZIhvcNAQEBBQAEgYC4RiWH
'' SIG '' PrUW1Ukl1QcCSLfxQB3B88DjZkx1HRhL/dNsXZ13fTs9
'' SIG '' tq0fCWXFdZg+iKB32DAPzlJ9sn7HzlMsjc7Co7OIXMrX
'' SIG '' ZJWgT0/uLFN6B+/EzeRCRGtm27QD8ren+OvCVX6rnTSY
'' SIG '' tve7QcwGoyaeNO6E80BOr1aePo7wHLh/lw==
'' SIG '' End signature block
