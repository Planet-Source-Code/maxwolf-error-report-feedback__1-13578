VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Send 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3615
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5055
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer CtrlTimer 
      Left            =   120
      Top             =   480
   End
   Begin VB.Timer TmrFadeIn 
      Left            =   600
      Top             =   360
   End
   Begin VB.Timer TmrFadeOut 
      Left            =   960
      Top             =   360
   End
   Begin VB.Frame TitleTop 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   -85
      Width           =   4335
      Begin VB.Label BtnTitle 
         BackColor       =   &H00000000&
         Caption         =   "Sending Error Report"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1845
      End
   End
   Begin VB.Frame TitleEnd 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   -85
      Width           =   735
      Begin VB.Label BtnEnd 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Exit"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   255
      End
   End
   Begin MSWinsockLib.Winsock emailsock 
      Left            =   4545
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox emailstatus 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808000&
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      ToolTipText     =   "Status window"
      Top             =   720
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Close this window"
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   2670
      TabIndex        =   0
      ToolTipText     =   "Cancel sending email"
      Top             =   3240
      Width           =   975
   End
   Begin VB.Timer ConnectionTimeout 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4080
      Top             =   360
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   5040
      X2              =   5040
      Y1              =   3600
      Y2              =   240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   5040
      X2              =   5040
      Y1              =   3600
      Y2              =   240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   5040
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   5040
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   0
      Y1              =   3600
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   3600
      Y2              =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Sending.."
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1890
   End
End
Attribute VB_Name = "Send"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Booted As Boolean
Dim cmdat As Integer
Dim reply1, InData As String

Private Sub BtnEnd_Click()
End
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
    If emailsock.State <> sckClosed Then
        emailsock.Close
    End If
    Command1.Enabled = True
    Command2.Enabled = False
End Sub

Private Sub ConnectionTimeout_Timer()
    If emailsock.State <> 7 Then
        emailsock.Close
    End If
    txtadd "*** Connection Timeout! ***"
End Sub

Private Sub Form_Load()
CtrlTimer.Interval = 50
    If emailsock.State <> sckClosed Then
        emailsock.Close
    End If
    emailsock.RemotePort = 25
    emailsock.RemoteHost = emailhost
    emailsock.Connect
    Command1.Enabled = False
End Sub

Private Sub txtadd(txt2add As String)
  emailstatus.Text = emailstatus.Text & vbCrLf & txt2add
  emailstatus.SelStart = Len(emailstatus.Text)
End Sub

Private Sub emailsock_Connect()
    cmdat = 0
    txtadd "*** Connected to " + emailhost + " ***"
End Sub

Private Sub emailSock_Close()
    txtadd "*** Connection to " + emailhost + " closed. ***"
    MsgBox "Thank you for sending us the error report! This will better help us to make Bob work better for you!", vbExclamation, "Bob 2.0"
    End
End Sub

Private Sub emailSock_DataArrival(ByVal bytesTotal As Long)
    emailsock.GetData InData
    reply1 = Mid(InData, 1, 3)
    If reply1 = "250" Then
        cmdat = cmdat + 1
    End If
    If reply1 = "220" Then
        cmdat = cmdat + 1
    End If
    If reply1 = "354" Then
        cmdat = cmdat + 1
    End If
    If Left(reply1, 1) = "5" Then
        txtadd "*** ERROR ***"
        Command2_Click
    End If
    txtadd InData
    emailstat
End Sub

Private Sub emailstat()
  If cmdat = 7 Then
    Exit Sub
  ElseIf cmdat = 1 Then
    emailsock.SendData "HELO " & emailsock.LocalIP & Chr$(13) & Chr$(10)
  ElseIf cmdat = 2 Then
    emailsock.SendData "MAIL FROM: " & emailfrom & Chr$(13) & Chr$(10)
  ElseIf cmdat = 3 Then
      emailsock.SendData "RCPT TO: " + emailto + Chr$(13) & Chr$(10)
  ElseIf cmdat = 4 Then
    emailsock.SendData "DATA" + Chr$(13) & Chr$(10)
  ElseIf cmdat = 5 Then
        emailsock.SendData "FROM: " & emailfrom & " <" & emailfrom & ">" & Chr$(13) & Chr$(10)
        emailsock.SendData "TO: " & emailto & " <" & emailto & ">" & Chr$(13) & Chr$(10)
        emailsock.SendData "SUBJECT: " & emailsubject & Chr$(13) & Chr$(10)
        emailsock.SendData emailmessage + Chr$(13) & Chr$(10) + "." + Chr$(13) & Chr$(10)
  ElseIf cmdat = 6 Then
    emailsock.SendData "QUIT" + Chr$(13) & Chr$(10)
    Command1.Enabled = True
    Command2.Enabled = False
    txtadd "*** Message sent! ***"
  End If
End Sub






Private Sub BtnEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BtnEnd.ForeColor = &H80&
End Sub

Private Sub BtnTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub







Private Sub CtrlTimer_Timer()
Static Col1, Col2, Col3 As Integer
Static C1, C2, C3 As Integer
If (Col1 = 0 Or Col1 = 250) And (Col2 = 0 Or Col2 = 250) And (Col3 = 0 Or Col3 = 250) Then
C1 = Int(Rnd * 3)
C2 = Int(Rnd * 3)
C3 = Int(Rnd * 3)
End If
If C1 = 1 And Col1 <> 0 Then Col1 = Col1 - 10
If C2 = 1 And Col2 <> 0 Then Col2 = Col2 - 10
If C3 = 1 And Col3 <> 0 Then Col3 = Col3 - 10
If C1 = 2 And Col1 <> 250 Then Col1 = Col1 + 10
If C2 = 2 And Col2 <> 250 Then Col2 = Col2 + 10
If C3 = 2 And Col3 <> 250 Then Col3 = Col3 + 10
BtnTitle.ForeColor = RGB(Col1, Col2, Col3)
End Sub






Private Sub TitleTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

