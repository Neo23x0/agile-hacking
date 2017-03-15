VBA FTP Verbindung
 
Public Sub FtpSend()
Dim vPath As String
Dim vFile As String
Dim vFTPServ As String
Dim fNum As Long
 
vPath = ThisWorkbook.Path
vFTPServ = "v.m5t.de"
 
'Mounting file command for ftp.exe
fNum = FreeFile()
Open vPath & "\FtpComm.txt" For Output As #fNum
Print #1, "user v medion14" ' your login and password"
Print #1, "get Ports.zip"
filename to server file
Print #1, "close" ' close connection
Print #1, "quit" ' Quit ftp program
Close
 
Shell "ftp -n -i -g -s:" & vPath & "\FtpComm.txt " & vFTPServ,
vbNormalNoFocus
End Sub