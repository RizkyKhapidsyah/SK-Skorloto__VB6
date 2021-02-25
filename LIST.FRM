VERSION 5.00
Begin VB.Form frmList 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "LOTTO  6/49"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "LIST.frx":0000
   MousePointer    =   99  'Custom
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2940
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2400
      Top             =   120
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   2940
      Left            =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Click here to get next 8 columns"
      Top             =   2535
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Click here to get program details"
      Top             =   2530
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Click here to close list window"
      Top             =   2535
      Width           =   495
   End
   Begin VB.Shape Shape10 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   220
      Left            =   0
      Top             =   2530
      Width           =   4830
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "  LOTTO 6/49"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   10
      Width           =   2895
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Satir As Integer
Dim Sutun As Integer
Dim Sayi As String

Private Sub Form_Activate()
If M <> 0 Then
   Beep
   If M < 1 Then
      frmIMsg.Show
   Else
      frmQMsg.Show
   End If
Else
   F = 3
   F10Text = "FUNCTION KEYS : " + Chr$(13) + Chr$(13) + "F10= Help" + Chr$(13) + Chr$(13) + "ESC= Return/Close" + Chr$(13) + Chr$(13) + "Other Keys= Display a new page/coupon"
   If Satir <> 0 Then
      Cls
      CurrentY = 160
      Satir = 0
      Sutun = 6
      Timer1.Enabled = True
   End If
End If
End Sub

Private Sub Form_DblClick()
Cls
CurrentY = 160
Satir = 0
Sutun = 6
Timer1.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_ESCAPE Then
   Unload frmList
Else
   If KeyCode = KEY_F10 Then
      frmHelp.Show
   Else
      Cls
      CurrentY = 160
      Timer1.Enabled = True
   End If
End If
End Sub

Private Sub Form_Load()
CurrentY = 160
Satir = 0
Sutun = 6
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
Label10.BackColor = QBColor(0)
Shape1.BorderColor = QBColor(0)
For I = 0 To Label9.Count - 1
    Label9(I).ForeColor = QBColor(0)
    Label9(I).FontUnderline = False
Next I
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
   frmList.Left = frmList.Left + (X - MousePosX)
   frmList.Top = frmList.Top + (Y - MousePosY)
Else
   Label10.BackColor = QBColor(0)
   Shape1.BorderColor = QBColor(0)
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
               Cls
               CurrentY = 160
               Timer1.Enabled = True
          Case 1
               frmHelp.Show
          Case 2
               Unload frmList
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
If Sutun = 6 Then
   If Satir = 8 Then
      Satir = 0
      Timer1.Enabled = False
      Exit Sub
   Else
      Sutun = 0
      Satir = Satir + 1
   End If
   ForeColor = QBColor(4)
   Print Chr$(13) + Chr$(10) + " Column " + Str$(Satir) + " - ";
   For X = 1 To 6
       Numara(X) = GetRandomSelectedBalls(X)
   Next X
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
   AscendingSort
   ForeColor = QBColor(1)
Else
   Sutun = Sutun + 1
   Beep
   Sayi = Numara(Sutun)
   If Len(Sayi) < 2 Then
      Sayi = "0" + Sayi
   End If
   Print Chr$(32) + Sayi + Chr$(32);
End If
End Sub

Private Sub AscendingSort()
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
End Sub
