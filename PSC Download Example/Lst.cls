VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This Class is NOT written by HardStream Software
'
'Programmer: Jason Hensley
'
'Email: elitecobra@hotmail.com
'
'Website: http://www.vbcodesource.com
'
'Version: 2.2 Final
'
'Stipped from the MusicModule 2.2 by www.vbcodesource.com
'
'
Option Explicit
'Api to access the registry
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Const HKEY_CLASSES_ROOT = &H80000000
'Api to move form without titlebar
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long

'Open a list
Public Sub OpenList(TheList As String, Listbox As ComboBox)
On Error GoTo TrapIt
Dim test As String
If TheList = "" Then Exit Sub
Open TheList For Input As 1
While Not EOF(1)
Line Input #1, test
Listbox.AddItem RTrim(test)
Wend
Close 1
Exit Sub
TrapIt:
Exit Sub
End Sub

'Used to find the last backslash of the file path
Public Function GetLastSlash(Text As String) As String
On Error GoTo TrapIt
Dim i, Pos As Integer
Dim LastSlash As Integer
For i = 1 To Len(Text)
Pos = InStr(i, Text, "/", vbTextCompare)
If Pos <> 0 Then LastSlash = Pos
Next i
GetLastSlash = Right(Text, Len(Text) - LastSlash)
TrapIt:
Exit Function
End Function

'Used for getting the Right/Left of a string
Public Function RightLeft(Source As String, Token As String) As String
On Error GoTo TrapIt
Dim i As Long
RightLeft = ""
For i = Len(Source) To 1 Step -1
If Mid(Source, i, 1) = Token Then
RightLeft = Left(Source, i - 1)
Exit Function
End If
Next i
Exit Function
TrapIt:
Exit Function
End Function

'Save a list
Public Sub SavePlaylist(TheList As String, Listbox As ComboBox)
On Error GoTo TrapIt
Dim i As Integer
Dim a As String
Open TheList For Output As #1
For i = 0 To Listbox.ListCount - 1
a$ = Listbox.List(i)
Print #1, a$
Next
Close 1
Exit Sub
TrapIt:
Exit Sub
End Sub
