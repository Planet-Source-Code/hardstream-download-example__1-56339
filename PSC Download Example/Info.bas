Attribute VB_Name = "Info"
'ęCopyright HardStream Software
'This module is written by HardStream Software

Function DesktopPath() As String
Dim WShell
Set WShell = CreateObject("wscript.shell")
DesktopPath = WShell.specialfolders("desktop")
End Function
