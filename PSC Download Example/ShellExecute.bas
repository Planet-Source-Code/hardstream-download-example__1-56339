Attribute VB_Name = "ShellExecute"
'®Copyright HardStream Software
'This module is written by HardStream Software

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub EasyExecute(Filename)
ShellExecute 0, "open", Filename, "", vbNull, SW_SHOWNORMAL
End Sub
