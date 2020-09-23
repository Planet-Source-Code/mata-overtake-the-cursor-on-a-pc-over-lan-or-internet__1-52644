VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSetCursor 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Kreuz
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximiert
   Begin VB.Timer tmrConnect 
      Interval        =   5000
      Left            =   0
      Top             =   720
   End
   Begin MSWinsockLib.Winsock wskClick 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit!"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmSetCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private vClick As Integer

Private Sub cmdExit_Click()
Unload Me
frmHaupt.tmrCursor.Enabled = False
frmHaupt.cmdStart.Enabled = True
frmHaupt.hscTimeout.Enabled = True
frmHaupt.cmdConnect.Enabled = True
End Sub


Private Sub Form_Load()
Load wskClick(1)
vClick = 1
wskClick(0).Connect frmHaupt.ip, 6000
Sleep (100)
wskClick(1).Connect frmHaupt.ip, 6001
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If tmrConnect.Enabled = False Then
    Dim strData As String
    vClick = vClick * -1
    If vClick = 1 Then
        If Button = vbRightButton Then
            strData = "RIGHT"
            wskClick(0).SendData (strData)
        ElseIf Button = vbLeftButton Then
            strData = "LEFT"
            wskClick(0).SendData (strData)
        End If
    Else
        If Button = vbRightButton Then
            strData = "RIGHT"
            wskClick(1).SendData (strData)
        ElseIf Button = vbLeftButton Then
            strData = "LEFT"
            wskClick(1).SendData (strData)
        End If
    End If
Else
    MsgBox "Wait until you are connected!", vbExclamation, "Wait"
End If
End Sub
Private Sub lblPosition_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If tmrConnect.Enabled = False Then
    Dim strData As String
    vClick = vClick * -1
    If vClick = 1 Then
        If Button = vbRightButton Then
            strData = "RIGHT"
            wskClick(0).SendData (strData)
        ElseIf Button = vbLeftButton Then
            strData = "LEFT"
            wskClick(0).SendData (strData)
        End If
    Else
        If Button = vbRightButton Then
            strData = "RIGHT"
            wskClick(1).SendData (strData)
        ElseIf Button = vbLeftButton Then
            strData = "LEFT"
            wskClick(1).SendData (strData)
        End If
    End If
Else
    MsgBox "Wait until you are connected!", vbExclamation, "Wait"
End If
End Sub

Private Sub tmrConnect_Timer()
MsgBox "Couldn't connect to " & frmHaupt.ip, vbExclamation, "Error"
End Sub

Private Sub wskClick_Close(index As Integer)
If index = 1 Then
MsgBox "Lost connection to " & frmHaupt.ip, vbExclamation, "Error"
Unload Me
frmHaupt.tmrCursor.Enabled = False
frmHaupt.cmdStart.Enabled = True
frmHaupt.hscTimeout.Enabled = True
frmHaupt.cmdConnect.Enabled = True
End If
End Sub

Private Sub wskClick_Connect(index As Integer)
If index = 1 Then
tmrConnect.Enabled = False
MsgBox "OK, let's go." & vbCrLf & "Just Click.", vbInformation, "Let's Go"
frmHaupt.tmrCursor.Enabled = True
End If
End Sub
