VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   5775
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer CtrlTimer 
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame TitleTop 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   -85
      Width           =   4335
      Begin VB.Label BtnTitle 
         BackColor       =   &H00000000&
         Caption         =   "Error Reporter"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1845
      End
   End
   Begin VB.Frame TitleAction 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   -85
      Width           =   735
      Begin VB.Label BtnUp 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Up"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label BtnDown 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   " Down"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.Frame TitleEnd 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   -85
      Width           =   735
      Begin VB.Label BtnEnd 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Exit"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.TextBox emailhosttxt 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "pop-server.columbus.rr.com"
      ToolTipText     =   "SMTP email server"
      Top             =   420
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.TextBox emailfromtxt 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "bob@bob.net"
      ToolTipText     =   "Email from address (Doesn't have to exist, but some SMTP servers require the domain to exist.)"
      Top             =   765
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.TextBox emailtotxt 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "maxwolf@columbus.rr.com"
      ToolTipText     =   "Recipient's address"
      Top             =   1125
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.TextBox emailsubjecttxt 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Error Report"
      ToolTipText     =   "Subject"
      Top             =   1485
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.TextBox emailmessagetxt 
      Height          =   3030
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "Main.frx":0682
      ToolTipText     =   "Message body"
      Top             =   405
      Width           =   5445
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   0
      Y1              =   4080
      Y2              =   240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   4080
      Y2              =   240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   5760
      X2              =   0
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   5760
      X2              =   0
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   5760
      X2              =   5760
      Y1              =   4080
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   5760
      X2              =   5760
      Y1              =   4080
      Y2              =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuSend 
         Caption         =   "&Send"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu sepr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
Call mnuSend_Click
End Sub

Private Sub Command2_Click()
Call mnuClear_Click
End Sub

'The Fake Email source code is provided 'As is'
'with no warrenties whatsoever.
'
'By using it, you agree you will not hold www.fakeemail.org
'responsible for anything you do or happens from using
'it.
'
'You may only use this code for educational purposes,
'any other uses for this code is strictly prohibited
'
'www.fakeemail.org


Private Sub mnuClear_Click()
    emailhosttxt.Text = ""
    emailtotxt.Text = ""
    emailfromtxt.Text = ""
    emailsubjecttxt.Text = ""
    emailmessagetxt.Text = ""
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuSend_Click()
    
    emailhost = emailhosttxt.Text
    emailfrom = emailfromtxt.Text
    emailto = emailtotxt.Text
    emailsubject = emailsubjecttxt.Text
    emailmessage = emailmessagetxt.Text
    If emailhost = "" Then
        MsgBox "SMTP Server required!"
    ElseIf emailfrom = "" Then
        MsgBox "From address required! "
    ElseIf emailto = "" Then
        MsgBox "To address required!"
    ElseIf emailmessage = "" Then
        MsgBox "Can't send a blank message!"
    Else
        Me.Hide
        emailhosttxt.Text = ""
        emailfromtxt.Text = ""
        emailtotxt.Text = ""
        emailsubjecttxt.Text = ""
        emailmessagetxt.Text = ""
        Unload Main
        Send.Show 1
    End If
End Sub



Private Sub BtnDown_Click()
Me.Height = 4125
BtnDown.Visible = False
BtnUp.Visible = True
End Sub

Private Sub BtnDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BtnDown.ForeColor = &H80&
End Sub

Private Sub BtnEnd_Click()
End
End Sub

Private Sub BtnEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BtnEnd.ForeColor = &H80&
End Sub

Private Sub BtnTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub BtnUp_Click()
Me.Height = 315
BtnUp.Visible = False
BtnDown.Visible = True
End Sub

Private Sub BtnUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BtnUp.ForeColor = &H80&
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

Private Sub Form_Load()
CtrlTimer.Interval = 50
BtnUp.Visible = True
End Sub




Private Sub TitleAction_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BtnUp.ForeColor = &H8000&
BtnDown.ForeColor = &H8000&
End Sub

Private Sub TitleEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BtnEnd.ForeColor = &H8000&
End Sub

Private Sub TitleTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub



