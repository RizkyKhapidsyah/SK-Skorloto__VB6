VERSION 5.00
Begin VB.Form frmHelp 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "http://www.singlix.com/download/skorloto.html"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   Icon            =   "HELP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "HELP.frx":0442
   MousePointer    =   99  'Custom
   ScaleHeight     =   4395
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   240
      Top             =   1920
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      Picture         =   "HELP.frx":074C
      ScaleHeight     =   615
      ScaleWidth      =   855
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "HELP.frx":098C
      ScaleHeight     =   495
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   3775
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00A0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00808000&
      Caption         =   " http://www.singlix.com/download/skorloto.html"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   275
      Left            =   0
      TabIndex        =   7
      Top             =   10
      Width           =   5300
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808000&
      Height          =   4400
      Left            =   0
      Top             =   0
      Width           =   5300
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   2535
      Left            =   1080
      TabIndex        =   6
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "© Erdogan Tan - 2001"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1275
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ADVANCED ARCHITECTURE"
      ForeColor       =   &H00007070&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SKORLOTO 417 Version 2.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   4000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SKORLOTO 417 Version 2.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1125
      TabIndex        =   2
      Top             =   510
      Width           =   4000
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Kayma As Integer

Private Sub Command1_Click()
Unload frmHelp
End Sub

Private Sub Form_Activate()
If M <> 0 Then
   Beep
   If M < 1 Then
      frmIMsg.Show
   Else
      frmQMsg.Show
   End If
Else
   F = 4
End If
End Sub

Private Sub Form_DblClick()
Unload frmHelp
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_ESCAPE Then
   Unload frmHelp
Else
   If KeyCode = KEY_F10 Then
      Beep
      MsgTxt = "BUILD: 2001.7.21" + Chr$(13) + Chr$(13) + "417@singlix.com"
      frmIMsg.Show
   End If
End If
End Sub

Private Sub Form_Load()
Label5.Caption = F10Text
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
Label10.BackColor = &H808000
Shape1.BorderColor = &H808000
Label6.ForeColor = QBColor(0)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.BackColor = QBColor(12)
Shape1.BorderColor = QBColor(12)
MousePosX = X
MousePosY = Y
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
If Button = 1 Then
   frmHelp.Left = frmHelp.Left + (X - MousePosX)
   frmHelp.Top = frmHelp.Top + (Y - MousePosY)
Else
   Label10.BackColor = &H808000
   Shape1.BorderColor = &H808000
   Label6.ForeColor = QBColor(0)
End If
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
Label10.BackColor = &H808000
Shape1.BorderColor = &H808000
Label6.ForeColor = QBColor(0)
End Sub

Private Sub Label6_Click()
Unload frmHelp
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 99
If Button = 1 Then
   Label6.ForeColor = QBColor(12)
Else
   Label6.ForeColor = QBColor(9)
End If
End Sub

Private Sub Timer1_Timer()
If Kayma < 490 Then
   Picture2.Top = Picture2.Top + 5
   Kayma = Kayma + 1
Else
   If Picture2.Top > 990 Then
      Picture2.Top = Picture2.Top - 5
   Else
      Kayma = 0
   End If
End If
If 3 * Rnd > 2 Then Label4.ForeColor = QBColor(Int(16 * Rnd))
End Sub
