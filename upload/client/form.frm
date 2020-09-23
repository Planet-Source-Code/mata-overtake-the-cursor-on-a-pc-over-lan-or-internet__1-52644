VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmHaupt 
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2040
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   2040
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fraWichCon 
      Caption         =   "Choose Connection"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
      Begin VB.OptionButton optInternet 
         Caption         =   "Internet"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optLAN 
         Caption         =   "LAN"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame fraTimeout 
      Caption         =   "Set Timeout"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
      Begin VB.HScrollBar hscTimeout 
         Height          =   255
         Left            =   120
         Max             =   1000
         Min             =   10
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Value           =   20
         Width           =   1095
      End
      Begin VB.Label lblTimeout 
         Alignment       =   1  'Rechts
         Caption         =   "20"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Timer tmrConnect 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1920
      Top             =   840
   End
   Begin MSWinsockLib.Winsock wskCursor 
      Index           =   0
      Left            =   1920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraIP 
      Caption         =   "Enter IP or Name"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start!"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Timer tmrCursor 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   480
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect!"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   1920
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Zentriert
      Caption         =   "Not Connected"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   1815
   End
End
Attribute VB_Name = "frmHaupt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type PointAPI
        X As Long
        Y As Long
End Type
Private CaseOpt As Boolean
Private i As Long
Private vIndex As Integer
Public ip As String
Private CursorPos() As String
Private Connection As Boolean
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Private Sub cmdConnect_Click()
Dim i As Integer
Dim port As Integer
Dim SleepTime As Integer
If Connection = False Then
    If optInternet = True Then
        SleepTime = 500
        CaseOpt = True
    ElseIf optLAN = True Then
        SleepTime = 50
        CaseOpt = False
    End If
    port = 5000
    frmHaupt.ip = txtIP.Text
    For i = 0 To 21
        port = port + 1
        wskCursor(i).Connect ip, port
        Sleep (SleepTime)
    Next i
    lblPosition.Caption = "Connecting ..."
    cmdConnect.Enabled = False
    tmrConnect.Enabled = True
    If CaseOpt = True Then
        optLAN.Enabled = False
        optInternet.Enabled = False
    Else
        optInternet.Enabled = False
        optLAN.Enabled = False
    End If
ElseIf Connection = True Then
    For i = 0 To 21
        wskCursor(i).Close
    Next i
    optLAN.Enabled = True
    optInternet.Enabled = True
    cmdStart.Enabled = False
    Connection = False
    cmdConnect.Caption = "Connect!"
    lblPosition.Caption = "Not Connected"
End If
End Sub

Private Sub cmdStart_Click()
tmrCursor.Interval = lblTimeout.Caption
cmdConnect.Enabled = False
cmdStart.Enabled = False
hscTimeout.Enabled = False
Load frmSetCursor
frmSetCursor.Visible = True
End Sub
Private Sub Form_Load()
Dim i As Integer
For i = 1 To 21
        Load wskCursor(i)
    Next i
Connection = False
Unload frmSetCursor
End Sub
Private Sub hscTimeout_Change()
lblTimeout.Caption = hscTimeout.Value
End Sub

Private Sub hscTimeout_Scroll()
lblTimeout.Caption = hscTimeout.Value
End Sub

Private Sub optInternet_Click()
hscTimeout.Value = 200
End Sub

Private Sub optLAN_Click()
hscTimeout.Value = 20
End Sub

Private Sub tmrCursor_Timer()
vIndex = vIndex + 1
If vIndex = 21 Then vIndex = 0
Dim CursorPos As String
frmSetCursor.lblPosition.Caption = RecordMouse
CursorPos = RecordMouse
wskCursor(vIndex).SendData (CursorPos)
End Sub
Private Function RecordMouse() As String
Dim pos As PointAPI
Dim cursor As Long
cursor = GetCursorPos(pos)
RecordMouse = pos.X & " " & pos.Y
End Function

Private Sub tmrConnect_Timer()
Dim i As Integer
MsgBox "Couldn't connect to " & ip, vbExclamation, "Error"
cmdConnect.Enabled = True
lblPosition.Caption = ""
For i = 0 To 21
wskCursor(i).Close
Next i
tmrConnect.Enabled = False
optLAN.Enabled = True
optInternet.Enabled = True
End Sub

Private Sub wskCursor_Close(index As Integer)
Dim i As Integer
If Connection = True Then
    MsgBox "Lost connection to " & ip, vbExclamation, "Error"
    For i = 0 To 21
        wskCursor(i).Close
    Next i
    Connection = False
    cmdConnect.Caption = "Connect!"
    cmdConnect.Enabled = True
    cmdStart.Enabled = False
    lblPosition.Caption = "Not Connected"
    tmrCursor.Enabled = False
    optLAN.Enabled = True
    optInternet.Enabled = True
    If frmSetCursor.Visible = True Then Unload (frmSetCursor)
End If
End Sub

Private Sub wskCursor_Connect(index As Integer)
If index = 10 Then
vIndex = -1
tmrConnect.Enabled = False
cmdConnect.Enabled = True
cmdStart.Enabled = True
cmdConnect.Caption = "Disconnect!"
lblPosition.Caption = "Connected"
Connection = True
MsgBox "Connected to " & ip & "." & vbCrLf & "Press 'Start' to take over the cursor.", vbInformation, "Connected"
End If
End Sub


