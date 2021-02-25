VERSION 5.00
Begin VB.Form frmLoto 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "SKORLOTO 417 Version 2.0"
   ClientHeight    =   2400
   ClientLeft      =   1740
   ClientTop       =   2205
   ClientWidth     =   4830
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "SKORLOTO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "SKORLOTO.frx":0442
   MousePointer    =   99  'Custom
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2400
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2040
      Top             =   2760
   End
   Begin VB.Shape Shape7 
      Height          =   2400
      Left            =   0
      Top             =   0
      Width           =   4830
   End
   Begin VB.Shape Shape10 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFB000&
      FillStyle       =   0  'Solid
      Height          =   275
      Left            =   0
      Top             =   2000
      Width           =   4830
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   14
      ToolTipText     =   "Click here to get program details"
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   13
      ToolTipText     =   "Click here to get 8 random columns"
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Banko"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   12
      ToolTipText     =   "Click here to open banko ball selection window"
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   11
      ToolTipText     =   "Click here to open ball selection window"
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Click here to sort balls in ascending order"
      Top             =   1680
      Width           =   495
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00F0E0D0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   1640
      Width           =   4830
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "© Erdogan Tan - 2001"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A1A1A1&
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
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
      Index           =   1
      Left            =   4125
      TabIndex        =   8
      ToolTipText     =   "Click here to exit"
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Column"
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
      Index           =   0
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Click here to get next 6/49 balls"
      Top             =   1320
      Width           =   855
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFF0E0&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   0
      Top             =   1245
      Width           =   4830
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF8000&
      Caption         =   " SKORLOTO 417 Version 2.1e"
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
      TabIndex        =   6
      Top             =   10
      Width           =   4830
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   100
      Picture         =   "SKORLOTO.frx":074C
      Top             =   375
      Width           =   810
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   315
      Left            =   4245
      Top             =   435
      Width           =   405
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4245
      TabIndex        =   5
      Top             =   435
      Width           =   405
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   315
      Left            =   3705
      Top             =   435
      Width           =   405
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3705
      TabIndex        =   4
      Top             =   435
      Width           =   405
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   315
      Left            =   3150
      Top             =   435
      Width           =   405
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3150
      TabIndex        =   3
      Top             =   435
      Width           =   405
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   315
      Left            =   2595
      Top             =   435
      Width           =   405
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2595
      TabIndex        =   2
      Top             =   435
      Width           =   405
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   315
      Left            =   2055
      Top             =   435
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2055
      TabIndex        =   1
      Top             =   435
      Width           =   405
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   315
      Left            =   1500
      Top             =   435
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Top             =   435
      Width           =   405
   End
End
Attribute VB_Name = "frmLoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Zaman As Integer
Dim TimeOut As Integer

Private Sub Form_Activate()
If M <> 0 Then
   Beep
   If M < 1 Then
      frmIMsg.Show
   Else
      frmQMsg.Show
   End If
Else
   F = 0
   F10Text = "FUNCTION KEYS : " + Chr$(13) + Chr$(13) + "F1=  New 6/49 Column" + Chr$(13) + "F2=  Bag (Balls will be included by columns)" + Chr$(13) + "F3=  Banko Balls (will be used by every column)" + Chr$(13) + "F4=  List (8 columns)" + Chr$(13) + "F5=  Sort balls in ascending order" + Chr$(13) + Chr$(13) + "F10= Help" + Chr$(13) + Chr$(13) + "ESC= Exit/Return/Close"
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_ESCAPE Then
   End
Else
   If KeyCode = KEY_F1 Then
      Timer1.Enabled = True
   Else
      If KeyCode = KEY_F2 Then
         frmBalls.Show
      Else
         If KeyCode = KEY_F3 Then
            frmBanko.Show
         Else
            If KeyCode = KEY_F4 Then
               If Bag(0) < 6 Then
                  Beep
                  MsgTxt = "You have to select 6 balls at least !" + Chr$(13) + Chr$(13) + "Press F2 to enter ball selection menu..."
                  frmIMsg.Show
               Else
                  frmList.Show
               End If
            Else
               If KeyCode = KEY_F5 Then
                  AscendingSort
               Else
                  If KeyCode = KEY_F10 Then frmHelp.Show
               End If
            End If
         End If
      End If
   End If
End If
End Sub

Private Sub ShowBalls()
Beep
If TimeOut = True Then
   Timer1.Enabled = False
   TimeOut = False
Else
   Exit Sub
End If
Randomize
Numara(1) = Label1.Caption
Numara(2) = Label2.Caption
Numara(3) = Label3.Caption
Numara(4) = Label4.Caption
Numara(5) = Label5.Caption
Numara(6) = Label6.Caption
For X = 1 To 5
    For Y = X + 1 To 6
Tekrar:
        If Numara(Y) = Numara(X) Then
           Numara(Y) = GetRandomSelectedBalls(Y)
           X = 1
           GoTo Tekrar
        End If
    Next Y
Next X
Label1.Caption = Numara(1)
Label2.Caption = Numara(2)
Label3.Caption = Numara(3)
Label4.Caption = Numara(4)
Label5.Caption = Numara(5)
Label6.Caption = Numara(6)
End Sub

Private Sub Form_Load()
M = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
Label10.BackColor = &HFF8000
For I = 0 To Label9.Count - 1
    Label9(I).ForeColor = QBColor(0)
    Label9(I).FontUnderline = False
Next I
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.BackColor = QBColor(12)
MousePosX = X
MousePosY = Y
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
If Button = 1 Then
   frmLoto.Left = frmLoto.Left + (X - MousePosX)
   frmLoto.Top = frmLoto.Top + (Y - MousePosY)
Else
   Label10.BackColor = &HFF8000
   For I = 0 To Label9.Count - 1
       Label9(I).ForeColor = QBColor(0)
       Label9(I).FontUnderline = False
   Next I
End If
End Sub

Private Sub Label9_Click(Index As Integer)
If M <> 0 Then
   Beep
   If M < 1 Then
      frmIMsg.Show
   Else
      frmQMsg.Show
   End If
Else
   Select Case Index
          Case 0
               Timer1.Enabled = True
          Case 1
               End
          Case 2
               AscendingSort
          Case 3
               frmBalls.Show
          Case 4
               frmBanko.Show
          Case 5
               If Bag(0) < 6 Then
                  Beep
                  MsgTxt = "You have to select 6 balls at least !" + Chr$(13) + Chr$(13) + "Press F2 to enter ball selection menu..."
                  frmIMsg.Show
               Else
                  frmList.Show
               End If
          Case 6
               frmHelp.Show
   End Select
End If
End Sub

Private Sub Label9_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9(Index).ForeColor = QBColor(12)
End Sub

Private Sub Label9_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 99
If Button <> 1 Then
   Label9(Index).ForeColor = QBColor(9)
   Label9(Index).FontUnderline = True
End If
End Sub

Private Sub Label9_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9(Index).ForeColor = QBColor(1)
End Sub

Private Sub Timer1_Timer()
GetRandomSelectedBalls (6)
If Bag(0) < 6 Then
   Beep
   Timer1.Enabled = False
   MsgTxt = "You have to select 6 balls at least !" + Chr$(13) + Chr$(13) + "Press F2 to enter ball selection menu..."
   frmIMsg.Show
   Exit Sub
End If
If Zaman > (Int(24 * Rnd) + 1) Then
   TimeOut = True
   Zaman = 0
End If
Label1.Caption = GetRandomSelectedBalls(1)
Label2.Caption = GetRandomSelectedBalls(2)
Label3.Caption = GetRandomSelectedBalls(3)
Label4.Caption = GetRandomSelectedBalls(4)
Label5.Caption = GetRandomSelectedBalls(5)
Label6.Caption = GetRandomSelectedBalls(6)
Zaman = Zaman + 1
ShowBalls
End Sub

Private Sub AscendingSort()
Rem This code is used for sorting 6/49 lotto numbers in ascending order
Rem © Erdogan Tan 16 June 2001
Rem You can use this code for sorting numbers in ascending or descending order
Rem for example, number of teams (X) may be 18 and maximum point (Y) may be
Rem 34*3 or 102 [ SKORTOTO 417 program -> Football teams points chart ]
Rem This code is a part of "SKORLOTO 417 v1.2e" program and
Rem "skorloto_src.zip", a Visual Basic 5.0 (6.0) project
Rem Visit http://www.singlix.com/download/skorloto.html for details
ReDim Z(49) As Integer
For X = 1 To 6
    Z(Numara(X)) = True
Next X
Y = 0
For X = 1 To 6
    For Y = Y + 1 To 49
        If Z(Y) = True Then
           Numara(X) = Y
           Exit For
        End If
    Next Y
Next X
Beep
Label1.Caption = Numara(1)
Label2.Caption = Numara(2)
Label3.Caption = Numara(3)
Label4.Caption = Numara(4)
Label5.Caption = Numara(5)
Label6.Caption = Numara(6)
Rem above labels are used for showing a 6/49 lotto column
Rem before this subroutine, Label1 to Label6 had non-sorted
Rem values between 1 and 49
Rem Example:
Rem 37 (on Label1), 1 (on Label2), 13, 8, 24, 41 (non-sorted)
Rem 1 (on Label1), 8 (on Label2), 13, 24, 37, 41 (sorted)
End Sub
