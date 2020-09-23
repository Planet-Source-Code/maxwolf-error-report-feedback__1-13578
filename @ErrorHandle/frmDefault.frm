VERSION 5.00
Begin VB.Form frmError 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5775
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmDefault.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   3500
      Left            =   4680
      Top             =   360
   End
   Begin VB.TextBox DirPath 
      Height          =   285
      Left            =   4800
      TabIndex        =   10
      Text            =   "Program Path"
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Whine About It"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Bug Report"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Error:"
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   5535
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   $"frmDefault.frx":08CA
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5295
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer CtrlTimer 
      Left            =   4320
      Top             =   360
   End
   Begin VB.Frame TitleEnd 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   -85
      Width           =   735
      Begin VB.Label BtnEnd 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Exit"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame TitleTop 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   -85
      Width           =   5055
      Begin VB.Label BtnTitle 
         BackColor       =   &H00000000&
         Caption         =   "Error Reporter"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1845
      End
   End
   Begin VB.Label Beg 
      BackColor       =   &H00000000&
      Caption         =   """Please Don't Hate Me, Hate My Programmer!"""
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   480
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   120
      Picture         =   "frmDefault.frx":0971
      Top             =   360
      Width           =   525
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
Call PutValue("Error", "Path", App.Path, App.Path & "/plugin.ini")
Call PutValue("Error", "Version", App.Revision, App.Path & "/plugin.ini")
DirPath = GetValue("Error", "Path", App.Path & "/plugin.ini")
End Sub

Private Sub BtnEnd_Click()
End
End Sub

Private Sub BtnEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BtnEnd.ForeColor = &H80&
End Sub



Private Sub Command1_Click()
Unload Me
Main.Show 1
End Sub

Private Sub Command2_Click()
MsgBox "OhHhh, you hurt our feelings...", vbCritical, "It hurts..."
End Sub

Private Sub Command3_Click()
End
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
End Sub



Private Sub Timer1_Timer()
BobBeg Beg, "I am so sorry...", 1
BobBeg Beg, "Please, don't kill me :(", 1
BobBeg Beg, "I was only doing what you told me!", 1
BobBeg Beg, "Tell my master!", 1
BobBeg Beg, "He will fix it! Make it better!", 1
BobBeg Beg, "Please don't kill me...", 1
BobBeg Beg, "...Please...", 1
BobBeg Beg, "Make a nice error report!", 1
BobBeg Beg, "Send it to the boss!", 1
BobBeg Beg, "He will know what to do!!!", 1
BobBeg Beg, "He knows everything!", 1
BobBeg Beg, "", 1
BobBeg Beg, "", 1
BobBeg Beg, "", 1
BobBeg Beg, "", 1
BobBeg Beg, "", 1
BobBeg Beg, "", 1
BobBeg Beg, "Please don't kill me...", 1
BobBeg Beg, "Or eat me, you fatass...", 1
End Sub

Private Sub TitleEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BtnEnd.ForeColor = &H8000&
End Sub

Private Sub TitleTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub



Private Sub BobBeg(X As Control, ByVal GREETEDperson As String, FinalEffects As Integer)
Dim GetLetters As Integer
Dim SingleLetter As String
Dim xSize As Integer

X.Caption = ""                                                  'Clear Caption
For GetLetters% = 1 To Len(GREETEDperson$)                      'Count Characters
    SingleLetter$ = Mid$(GREETEDperson$, GetLetters%, 1)        'Get Single Character
    X.Caption = X.Caption + SingleLetter$                       'Add Each Character
    TimeOut 0.025                                               'Hold On!
Next GetLetters%

X.ForeColor = &HFFFF00: TimeOut 0.05                                     'Lighten Color
X.ForeColor = &HC0C000: TimeOut 0.05                                      'Darken Color
X.ForeColor = &H808000: TimeOut 0.05                                  'Fade Color
TimeOut 0.4                                                     'Hold On!
End Sub
