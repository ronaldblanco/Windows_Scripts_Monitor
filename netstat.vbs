' Store the arguments in a variable:
Set objArgs = Wscript.Arguments
folder = objArgs(0)
port = objArgs(1)

Const ForReading=1
Const ForWriting=2

filePath = folder & "netstat.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")
If (objFSO.FileExists(filePath)) Then
   Set myFile = objFSO.OpenTextFile(filePath, ForReading, True)
   Set myTemp= objFSO.CreateTextFile(filePath & ".tmp1", ForWriting, True)
Else
   Set myFile = objFSO.CreateTextFile(filePath, ForReading, True)
   Set myTemp= objFSO.CreateTextFile(filePath & ".tmp1", ForWriting, True)
End If

mycount = 0
Do While Not myFile.AtEndofStream
 myLine = myFile.ReadLine
   
 if (InStr(myLine, port) and InStr(myLine, "ESTABLISHED"))then
	mycount = mycount + 1
 end if

Loop
myTemp.WriteLine mycount


myFile.Close
myTemp.Close
If (objFSO.FileExists(filePath&".1")) Then
  objFSO.DeleteFile(filePath&".1")
end if
objFSO.MoveFile filePath&".tmp1", filePath&".1"
'' SIG '' Begin signature block
'' SIG '' MIID8wYJKoZIhvcNAQcCoIID5DCCA+ACAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFLYza+19Mjhf
'' SIG '' U26s5YtCuQ2uBv/ZoIICGzCCAhcwggGAoAMCAQICEIig
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
'' SIG '' BAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUhy3XUM8vmvpz
'' SIG '' t6Gjr8B3DdYeebgwDQYJKoZIhvcNAQEBBQAEgYAEczSw
'' SIG '' 86TosgbSWzb4DWVU7KHU1fKTo7nkRKt7wG7KrZZ14PmN
'' SIG '' dqVAP+jdhpvVmRVMwLbV7aRyo61beuXMReSLJnG7QOTV
'' SIG '' 4HAKcjvWdN75TUM6JSMb99lZdxhy+ycYU96JXA21uYve
'' SIG '' 9lh0vFu7C661bbzt5TiWKNO66N9sBA+SFg==
'' SIG '' End signature block
