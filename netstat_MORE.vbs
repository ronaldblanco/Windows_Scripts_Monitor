' Store the arguments in a variable:
Set objArgs = Wscript.Arguments
folder = objArgs(0)
port = objArgs(1)

Const ForReading=1
Const ForWriting=2
Set objFSO = CreateObject("Scripting.FileSystemObject")
filePath = folder & "netstat.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")
If (objFSO.FileExists(filePath)) Then
   Set myFile = objFSO.OpenTextFile(filePath, ForReading, True)
   Set myTemp= objFSO.CreateTextFile(filePath & ".tmp", ForWriting, True)
Else
   Set myFile = objFSO.CreateTextFile(filePath, ForReading, True)
   Set myTemp= objFSO.CreateTextFile(filePath & ".tmp", ForWriting, True)
End If

mycount = 0
Do While Not myFile.AtEndofStream
 myLine = myFile.ReadLine
    
 if (InStr(myLine, port) and InStr(myLine, "ESTABLISHED"))then
	myTemp.WriteLine "<br>" & myLine &"|"
 end if
 
Loop

myFile.Close
myTemp.Close
If (objFSO.FileExists(filePath & ".more")) Then
  objFSO.DeleteFile(filePath & ".more")
end if
objFSO.MoveFile filePath&".tmp", filePath & ".more"
'' SIG '' Begin signature block
'' SIG '' MIID8wYJKoZIhvcNAQcCoIID5DCCA+ACAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFAXeH8AWmiik
'' SIG '' uAHWETnVk2hiuOz6oIICGzCCAhcwggGAoAMCAQICEIig
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
'' SIG '' BAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUmcatX3mJKuNh
'' SIG '' a1LwuIWiZRYK/D4wDQYJKoZIhvcNAQEBBQAEgYBOVQd6
'' SIG '' PMG4HyBTizGcxodBvq6OSsjDSRDFX7mKAg3xHfC1voly
'' SIG '' v4dzypoaf27d7PXOyeMnCRWY+gUH4wgvaXacz+RzTC5w
'' SIG '' DkVxiE+pqDKRw8U+Mk/2MyvAhkHSNatWv4uVWhw6djHK
'' SIG '' Gir6RvvjW4lVOPWtfamlAJtfiWIg1NgS6A==
'' SIG '' End signature block
