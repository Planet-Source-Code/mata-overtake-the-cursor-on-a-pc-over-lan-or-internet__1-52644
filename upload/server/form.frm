VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSetCursor 
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleWidth      =   1950
   Begin MSWinsockLib.Winsock wskClick 
      Index           =   0
      Left            =   1800
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskCursor 
      Index           =   0
      Left            =   1800
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdWait 
      Caption         =   "Wait for Connection!"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Zentriert
      Caption         =   "Not Connected"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "frmSetCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Connection As Boolean
Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
Const MOUSEEVENTF_RIGHTDOWN = &H8
Const MOUSEEVENTF_RIGHTUP = &H10


Private Sub cmdWait_Click()
Dim i As Integer
Dim port As Integer
If Connection = False Then
    cmdWait.Caption = "Stop Connection!"
    Connection = True
    lblStatus.Caption = "Waiting ..."
    port = 5000
    For i = 0 To 21
        port = port + 1
        wskCursor(i).LocalPort = port
        wskCursor(i).Listen
    Next i
    wskClick(0).LocalPort = 6000
    wskClick(0).Listen
    wskClick(1).LocalPort = 6001
    wskClick(1).Listen
    Exit Sub
ElseIf Connection = True Then
    Connection = False
    cmdWait.Caption = "Wait for Connection!"
    lblStatus.Caption = "Not Connected"
    For i = 0 To 21
    wskCursor(i).Close
    Next i
    wskClick(0).Close
    wskClick(1).Close
    Exit Sub
End If
End Sub




Private Sub Form_Load()
Dim i As Integer
For i = 1 To 21
        Load wskCursor(i)
    Next i
Load wskClick(1)
End Sub

Private Sub wskClick_Close(index As Integer)
wskClick(0).Close
wskClick(0).LocalPort = 6000
wskClick(0).Listen
wskClick(1).Close
wskClick(1).LocalPort = 6001
wskClick(1).Listen
lblStatus.Caption = "Connected"
End Sub

Private Sub wskClick_ConnectionRequest(index As Integer, ByVal requestID As Long)
wskClick(index).Close
wskClick(index).Accept requestID
End Sub

Private Sub wskClick_DataArrival(index As Integer, ByVal bytesTotal As Long)
Dim strData As String
wskClick(index).GetData strData
If strData = "RIGHT" Then
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
ElseIf strData = "LEFT" Then
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End If
End Sub


Private Sub wskCursor_Close(index As Integer)
Dim i As Integer
Dim port As Integer
If Connection = True Then
    For i = 0 To 21
        wskCursor(i).Close
    Next i
port = 5000
    For i = 0 To 21
        port = port + 1
        wskCursor(i).LocalPort = port
        wskCursor(i).Listen
    Next i
Connection = True
cmdWait.Caption = "Stop Connection!"
lblStatus = "Waiting ..."
End If
End Sub

Private Sub wskCursor_ConnectionRequest(index As Integer, ByVal requestID As Long)
    wskCursor(index).Close
    wskCursor(index).Accept requestID
    If index = 0 Then
    lblStatus.Caption = "Connected"
    End If
End Sub

Private Sub wskCursor_DataArrival(index As Integer, ByVal bytesTotal As Long)

    SetCursor (index)
End Sub
Private Function SetCursor(index As Integer) As String
Dim strData As String
Dim position() As String
Dim PosX As Long
Dim PosY As Long
wskCursor(index).GetData strData
position = Split(strData)
    PosX = CLng(position(o))
    PosY = CLng(position(1))
lblStatus.Caption = PosX & " " & PosY
SetCursorPos PosX, PosY
End Function

