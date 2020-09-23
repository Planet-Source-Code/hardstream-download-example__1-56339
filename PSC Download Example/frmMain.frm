VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Download example"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Run when file is downloaded"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   2415
   End
   Begin Project1.Downloader DL 
      Left            =   5520
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   6015
   End
   Begin Project1.Sep Sep1 
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   3195
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   53
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download"
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar Prog 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Text            =   "http://home.wanadoo.nl/j.terluun/index.htm"
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "0 / 0"
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   2880
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "To local:"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "©Copyright HardStream Software"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   3240
      Width           =   2355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remote URL:"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Get acces to Lst class
Public Lst As New Lst

Sub SaveList()
'Save Combo1 to a file
Lst.SavePlaylist App.Path & "\Lst.Cbo", Combo1
End Sub

Private Sub Combo1_Change()
Text1.Text = DesktopPath & "\" & Lst.GetLastSlash(Combo1.Text)
End Sub

Private Sub Combo1_Click()
Text1.Text = DesktopPath & "\" & Lst.GetLastSlash(Combo1.Text)
End Sub

Private Sub Combo1_Scroll()
Text1.Text = DesktopPath & "\" & Lst.GetLastSlash(Combo1.Text)
End Sub

Private Sub Command1_Click()
CheckExistance
DL.BeginDownload Combo1.Text, Text1.Text
End Sub

'Check if an item already exists in Combo1
Function CheckExistance() As Boolean
On Error GoTo FndErr
Dim Found As Boolean
Dim URL As String
URL = Combo1.Text

Combo1.ListIndex = 0

'Check if the text in Combo1 is the same as URL
Check:
If Combo1.Text = URL Then
'URL is already in the list
CheckExistance = True
GoTo FndErr
Else
'Combo1.text and URL are not the same
If Not Combo1.ListIndex = Combo1.ListCount - 1 Then
'The current item isn't the last item, go to the next item
Combo1.ListIndex = Combo1.ListIndex + 1
'Check if the text in Combo1 is the same as URL
GoTo Check
Else
'The selected item is the last item in Combo1, URL isn't found in Combo1
CheckExistance = False
Combo1.AddItem URL
'Select last item in Combo1
Combo1.ListIndex = Combo1.ListCount - 1
SaveList
GoTo FndErr
End If
End If

'FndErr, if an error is found or the URL is found
FndErr:
Exit Function
End Function

Private Sub DL_DownloadComplete(MaxBytes As Long, SaveFile As String)
Prog.Value = 0

If Check1.Value = 1 Then EasyExecute SaveFile Else Exit Sub
End Sub

Private Sub DL_DownloadError(SaveFile As String)
MsgBox "Error while downloading to " & SaveFile & "." & vbCrLf & "Downloading will be stopped", vbCritical, "Download error"
End Sub

Private Sub DL_DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
Prog.Max = MaxBytes
Prog.Value = CurBytes
Label4.Caption = CurBytes & " / " & MaxBytes
End Sub

Private Sub Form_Initialize()
InitXP
End Sub

Private Sub Form_Load()
Text1.Text = DesktopPath & "\" & Lst.GetLastSlash(Combo1.Text)

If FileExist(App.Path & "\Lst.Cbo") = True Then
Lst.OpenList App.Path & "\Lst.Cbo", Combo1
GoTo Done
Else
Combo1.AddItem "http://home.wanadoo.nl/j.terluun/index.htm"
Combo1.AddItem "http://home.wanadoo.nl/j.terluun/Back.bmp"
Combo1.AddItem "http://home.wanadoo.nl/j.terluun/FolderDialog.ocx"
End If
Combo1.ListIndex = 0

CheckExistance

Done:
Exit Sub
End Sub
