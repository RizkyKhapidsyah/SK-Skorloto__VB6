VERSION 5.00
Begin VB.Form frmIMsg 
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
   MouseIcon       =   "IMSG.frx":0000
   MousePointer    =   99  'Custom
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   1995
   ScaleMode       =   0  'User
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "IMSG.frx":030A
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   1425
      Width           =   795
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   345
      Left            =   1695
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
Attribute VB_Name = "frmIMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Form_Activate()
M = -1
Label1.Caption = MsgTxt
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 27 Then Label9_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
Label9.ForeColor = QBColor(0)
Shape2.BorderColor = QBColor(9)
End Sub

Private Sub Form_Unload(Cancel As Integer)
M = 0
Select Case F
       Case 0
            frmLoto.Show
       Case 1
            frmBalls.Show
       Case 2
            frmBanko.Show
       Case 3
            frmList.Show
       Case 4
            frmHelp.Show
End Select
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
   frmIMsg.Left = frmIMsg.Left + (X - MousePosX)
   frmIMsg.Top = frmIMsg.Top + (Y - MousePosY)
Else
   Label9.ForeColor = QBColor(0)
   Shape2.BorderColor = QBColor(9)
End If
End Sub

Public Sub Label9_Click()
Unload frmIMsg
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = QBColor(12)
Shape2.BorderColor = QBColor(9)
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 99
Label9.ForeColor = QBColor(9)
Shape2.BorderColor = QBColor(12)
End Sub
