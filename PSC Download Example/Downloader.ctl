VERSION 5.00
Begin VB.UserControl Downloader 
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2385
   ScaleWidth      =   3480
   ToolboxBitmap   =   "Downloader.ctx":0000
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "Downloader.ctx":0312
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Downloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'I found this code somewhere on the internet
'I don't know where I found it, sorry!
'I thought it might be useful to VB coders...
'This usercontrol can download multiple files at the same time!

Option Explicit
Event DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
Event DownloadError(SaveFile As String)
Event DownloadComplete(MaxBytes As Long, SaveFile As String)

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
On Error Resume Next
Dim F() As Byte, Fn As Long
If AsyncProp.BytesMax <> 0 Then
Fn = FreeFile
F = AsyncProp.Value
Open AsyncProp.PropertyName For Binary Access Write As #Fn
Put #Fn, , F
Close #Fn
Else
RaiseEvent DownloadError(AsyncProp.PropertyName)
End If
RaiseEvent DownloadComplete(CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
On Error Resume Next
If AsyncProp.BytesMax <> 0 Then
RaiseEvent DownloadProgress(CLng(AsyncProp.BytesRead), CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)
End If
End Sub

Private Sub UserControl_Resize()
UserControl.Width = ScaleX(32, vbPixels, vbTwips)
UserControl.Height = ScaleY(32, vbPixels, vbTwips)
End Sub

Public Sub BeginDownload(URL As String, SaveFile As String)
On Error GoTo ErrorBeginDownload
UserControl.AsyncRead URL, vbAsyncTypeByteArray, SaveFile, vbAsyncReadForceUpdate
Exit Sub
ErrorBeginDownload:
RaiseEvent DownloadError(SaveFile)
Exit Sub
End Sub
