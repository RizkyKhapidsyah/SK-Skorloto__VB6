VERSION 5.00
Begin VB.Form frmQMsg 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "QMSG.frx":0000
   MousePointer    =   99  'Custom
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   1995
   ScaleMode       =   0  'User
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   345
      Index           =   1
      Left            =   2595
      Top             =   1380
      Width           =   1005
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2700
      TabIndex        =   3
      Top             =   1425
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "QMSG.frx":030A
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "  SKORLOTO 417 Version 2.1e"
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
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   10
      Width           =   4400
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   2000
      Left            =   0
      Top             =   0
      Width           =   4400
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "YES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   905
      TabIndex        =   1
      Top             =   1425
      Width           =   795
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   345
      Index           =   0
      Left            =   800
      Top             =   1380
      Width           =   1005
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   455
      Width           =   3525
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0F8FF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   0
      Top             =   1440
      Width           =   4400
   End
End
Attribute VB_Name = "frmQMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim H As Integer

Public Sub Form_Activate()
M = 1
If Q = 11 Then
   Label1.Caption = "Do you want to load balls from file SKORLOTO.DAT ?"
Else
   If Q = 12 Then
      Label1.Caption = "Do you want to save balls into SKORLOTO.DAT ?"
   Else
      If Q = 21 Then
         Label1.Caption = "Do you want to load banko balls from file SKORLOTO.DAT ?"
      Else
         If Q = 22 Then
            Label1.Caption = "Do you want to save banko balls into SKORLOTO.DAT ?"
         End If
      End If
   End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_UP Or KeyCode = KEY_DOWN Or KeyCode = 37 Or KeyCode = 39 Then
   If Label9(0).Visible = True Then
      If Label9(1).ForeColor = QBColor(9) Then
         Label9(1).ForeColor = QBColor(0)
         Label9(0).ForeColor = QBColor(9)
      Else
         Label9(0).ForeColor = QBColor(0)
         Label9(1).ForeColor = QBColor(9)
      End If
   End If
Else
   If KeyCode = 13 Then
      If Label9(1).ForeColor = QBColor(9) Then
         Label9_Click (1)
      Else
         If Label9(0).Visible = True And Label9(0).ForeColor = QBColor(9) Then Label9_Click (0)
      End If
   Else
      If KeyCode = 27 Then Label9_Click (1)
   End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
Label9(0).ForeColor = QBColor(0)
Label9(1).ForeColor = QBColor(0)
Shape2(0).BorderColor = QBColor(9)
Shape2(1).BorderColor = QBColor(9)
End Sub

Private Sub Form_Unload(Cancel As Integer)
M = 0
If Q = 11 Then
   If H = True Then Q = 0
   frmBalls.Show
Else
  If Q = 12 Then
     frmBalls.Show
  Else
     If Q = 21 Then
        If H = True Then Q = 0
        frmBanko.Show
     Else
        If Q = 22 Then
           frmBanko.Show
        End If
     End If
  End If
End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePosX = X
MousePosY = Y
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
If Button = 1 Then
   frmQMsg.Left = frmQMsg.Left + (X - MousePosX)
   frmQMsg.Top = frmQMsg.Top + (Y - MousePosY)
Else
   Label9(0).ForeColor = QBColor(0)
   Label9(1).ForeColor = QBColor(0)
   Shape2(0).BorderColor = QBColor(9)
   Shape2(1).BorderColor = QBColor(9)
End If
End Sub

Public Sub Label9_Click(Index As Integer)
Beep
If Index = 1 Then
   Q = 0
Else
   If Q = 11 Then
      Command3_1
   Else
     If Q = 12 Then
        Command4_1
     Else
        If Q = 21 Then
           Command3_2
        Else
           If Q = 22 Then
              Command4_2
           End If
        End If
     End If
   End If
   If H = True Then frmIMsg.Show
End If
Unload frmQMsg
End Sub

Private Sub Label9_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9(Index).ForeColor = QBColor(12)
Shape2(Index).BorderColor = QBColor(9)
End Sub

Private Sub Label9_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 99
Label9(Index).ForeColor = QBColor(9)
Shape2(Index).BorderColor = QBColor(12)
End Sub

Private Sub Command3_1()
FileNo = FreeFile
Open "SKORLOTO.DAT" For Random As #FileNo Len = 2
Get #FileNo, 1, X
If Y < (2 * X) Then
   H = True
   MsgTxt = "'SKORLOTO.DAT' has not correct size and data !"
Else
   H = False
   If X > 0 Then
      Bag(0) = X
      If Bag(0) < 50 Then
         For X = 1 To Bag(0)
             Get #FileNo, X + 1, Y
             Bag(X) = Y
         Next X
      Else
         Bag(0) = 0
      End If
   Else
      Bag(0) = 0
   End If
   
End If
Close #FileNo
End Sub

Private Sub Command4_1()
H = True
FileNo = FreeFile
Open "SKORLOTO.DAT" For Random As #FileNo Len = 2
For X = 0 To Bag(0)
    Y = Bag(X)
    Put #FileNo, X + 1, Y
Next X
If Bag(0) < 49 Then
   For X = Bag(0) + 1 To 49
       Y = 0
       Put #FileNo, X + 1, Y
   Next X
End If
Close #FileNo
MsgTxt = "Selected balls have been saved into SKORLOTO.DAT..."
End Sub



Private Sub Command3_2()
Beep
FileNo = FreeFile
Open "SKORLOTO.DAT" For Random As #FileNo Len = 2
Get #FileNo, 51, Y
If X < (2 * (Y + 51)) Then
   H = True
   MsgTxt = "'SKORLOTO.DAT' has not banko balls data !"
Else
   H = False
   If Y < 6 And Y > 0 Then
      Banko(0) = Y
      For X = 1 To Y
          Get #FileNo, 51 + X, Y
          Banko(X) = Y
      Next X
   Else
      Banko(0) = 0
   End If
End If
Close #FileNo
End Sub

Private Sub Command4_2()
H = True
FileNo = FreeFile
Open "SKORLOTO.DAT" For Random As #FileNo Len = 2
If FileLen("SKORLOTO.DAT") < 2 Then
   Y = 0
   For X = 1 To 50
       Put #FileNo, X, Y
   Next X
Else
   Get #FileNo, 1, Y
   For X = Y + 1 To 51
       Y = 0
       Put #FileNo, X, Y
   Next X
End If
For X = 0 To Banko(0)
    Y = Banko(X)
    Put #FileNo, 51 + X, Y
Next X
Close #FileNo

MsgTxt = "Selected balls have been saved into SKORLOTO.DAT..."
End Sub
