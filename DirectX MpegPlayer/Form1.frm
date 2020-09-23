VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectX MPEG Player"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Progress"
      Height          =   1935
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "00:00:00"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "00:00:00"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "00:00:00"
         Top             =   840
         Width           =   735
      End
      Begin VB.HScrollBar H1 
         Height          =   255
         LargeChange     =   10
         Left            =   240
         Max             =   0
         TabIndex        =   17
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Lenght :"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Time Used :"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "TimeLeft :"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "0%"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "50%"
         Height          =   195
         Left            =   2280
         TabIndex        =   19
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   195
         Left            =   4320
         TabIndex        =   18
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Balance"
      Height          =   735
      Left            =   5160
      TabIndex        =   7
      Top             =   2160
      Width           =   2055
      Begin VB.HScrollBar H2 
         Height          =   255
         LargeChange     =   100
         Left            =   120
         Max             =   5000
         Min             =   -5000
         SmallChange     =   10
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Left         Center        Right"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1875
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Speed"
      Height          =   1935
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Width           =   975
      Begin VB.VScrollBar V2 
         Height          =   1575
         LargeChange     =   40
         Left            =   120
         Max             =   10
         Min             =   200
         SmallChange     =   3
         TabIndex        =   6
         Top             =   240
         Value           =   100
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "<50%"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   1560
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   960
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "200%"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Volume"
      Height          =   1935
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   975
      Begin VB.VScrollBar V1 
         Height          =   1575
         LargeChange     =   2000
         Left            =   120
         Max             =   -10000
         SmallChange     =   100
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "0%"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   1560
         Width           =   210
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "50%"
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   960
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   195
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3240
      Top             =   2160
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   240
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   3840
      Picture         =   "Form1.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   2040
      MaskColor       =   &H000000FF&
      Picture         =   "Form1.frx":11AE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'visit : http:\\besoftwaredeveloper.flappie.nl

Dim Audio As IBasicAudio
Dim FileName As String
Dim MediaControl As IMediaControl
Dim MediaPosition As IMediaPosition

Private Sub Command1_Click()
On Error Resume Next
LoadFile
' tells the player to start making nois
MediaControl.Run
MediaPosition.CurrentPosition = 0
End Sub

Private Sub Command2_Click()
On Error Resume Next
' tells the player to stop playing
MediaControl.Stop
End Sub

Private Sub Command3_Click()
LoadFile True
End Sub

Function LoadFile(Optional ReLoad As Boolean)
On Error GoTo 1
If FileName = "" Or ReLoad Then
CommonDialog1.Filter = "*.wav, *.mp2, *.mp3, *.bmp, *.jpg, *.mpg, *.avi, *.mid|*.wav;*.mp2;*.mp3;*.bmp;*.jpg;*.mpg;*.avi;*.mid"
CommonDialog1.ShowOpen
If Dir(CommonDialog1.FileName) = "" Or CommonDialog1.FileName = "" Then
MsgBox "Please select a file", vbInformation + vbOKOnly, "Error"
Exit Function
End If
FileName = CommonDialog1.FileName

'if controls are loaded unload them and reload it :P
CleanUp
'making the controls
Set MediaControl = New FilgraphManager
Set Audio = MediaControl
Set MediaPosition = MediaControl

MediaControl.RenderFile FileName ' find a rendering for it. and load it up
H1.Max = MediaPosition.Duration ' setting the scrollbar max time
Text3.Text = SecToTime(MediaPosition.Duration) 'shows the lenght of the file
End If
Exit Function
1:
MsgBox "Render error no rendering device found", vbExclamation + vbOKOnly, "Error"
CleanUp
End Function

Private Sub Form_Unload(Cancel As Integer)
CleanUp
End Sub

Private Sub H1_Scroll()
On Error Resume Next
' set the possition you scrolled to
MediaPosition.CurrentPosition = H1.Value
End Sub

Private Sub H2_Change()
On Error Resume Next
'setting the balance
Audio.Balance = H2.Value
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
' setting the current posion to the slider
H1.Value = MediaPosition.CurrentPosition
Text1.Text = SecToTime(MediaPosition.Duration - MediaPosition.CurrentPosition)
Text2.Text = SecToTime(MediaPosition.CurrentPosition)
End Sub

Function CleanUp()
'unloading the objects
Set MediaControl = Nothing
Set Audio = Nothing
Set MediaPosition = Nothing
End Function

Private Sub V1_Change()
On Error Resume Next
'setting the volume
Audio.Volume = V1.Value
End Sub

Private Sub V2_Change()
On Error Resume Next
' set playing speed
MediaPosition.Rate = V2.Value / 100
End Sub

Private Function SecToTime(NewSec As Double) As String
'creates seconds to hours , min , sec

On Error Resume Next
Dim Secx, Minx, Hourx
NewSec = Int(NewSec)
If NewSec < 1 Then SecToTime = "00:00:00": Exit Function
Secx = NewSec - Int(NewSec / 60) * 60
Minx = Int((NewSec - Int(NewSec / 3600) * 3600) / 60)
Hourx = Int(NewSec / 3600)
If Int(Hourx) > 24 Then
SecToTime = "24:59:59"
Else
SecToTime = Format(Str(Hourx) & ":" & Str(Minx) & ":" & Str(Secx), "hh:mm:ss")
End If
End Function
