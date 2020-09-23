Attribute VB_Name = "FExist"
'©Copyright HardStream Software
'This module is written by HardStream Software

'Check if a file exists
Public Function FileExist(File As String) As Boolean
Dim FS
Set FS = CreateObject("Scripting.FileSystemObject")
FileExist = FS.FileExists(File)
End Function

'Check if a direction exists
Public Function DirExists(Dir As String) As Boolean
Dim FS
Set FS = CreateObject("Scripting.FileSystemObject")
DirExists = FS.folderexists(Dir)
End Function

'Check if a drive exists
Public Function DriveExist(Drive As String)
Dim FS, D
Set FS = CreateObject("Scripting.FileSystemObject")
If FS.DriveExists(Drive) = True Then
Set D = FS.GetDrive(Drive)
DriveExist = 1
If D.IsReady = True Then
DriveExist = 2
Exit Function
End If
Else
DriveExist = 0
End If
End Function
