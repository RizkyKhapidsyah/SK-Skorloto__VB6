VERSION 5.00
Begin VB.Form frmBanko 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFF4&
   BorderStyle     =   0  'None
   Caption         =   "Banko"
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "BANKO.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   4980
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1200
      TabIndex        =   11
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1200
      TabIndex        =   10
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1200
      TabIndex        =   9
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   8
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.OptionButton Option6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF4&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   6
      Top             =   440
      UseMaskColor    =   -1  'True
      Width           =   400
   End
   Begin VB.OptionButton Option5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF4&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   440
      UseMaskColor    =   -1  'True
      Width           =   400
   End
   Begin VB.OptionButton Option4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF4&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   440
      UseMaskColor    =   -1  'True
      Width           =   400
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF4&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   440
      UseMaskColor    =   -1  'True
      Width           =   400
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF4&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   440
      UseMaskColor    =   -1  'True
      Width           =   400
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF4&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   440
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   400
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INFO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   4
      Left            =   3840
      TabIndex        =   23
      ToolTipText     =   "Click here to get program details"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   375
      Index           =   3
      Left            =   3840
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   375
      Index           =   2
      Left            =   2400
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   375
      Index           =   1
      Left            =   3840
      Top             =   840
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   375
      Index           =   0
      Left            =   2400
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   22
      ToolTipText     =   "Click here to close banko ball selection window"
      Top             =   885
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   21
      ToolTipText     =   "Click here to save banko balls into 'SKORLOTO.DAT'"
      Top             =   1485
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   20
      ToolTipText     =   "Click here to load banko balls from 'SKORLOTO.DAT'"
      Top             =   1485
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   19
      ToolTipText     =   "Click here to use selected banko balls"
      Top             =   885
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00007000&
      Height          =   4980
      Left            =   0
      Top             =   0
      Width           =   5250
   End
   Begin VB.Label Label10 
      BackColor       =   &H00007000&
      Caption         =   "  Banko"
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
      Height          =   270
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   5250
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "© Erdogan Tan"
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   2400
      TabIndex        =   17
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Banko 5 ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Banko 4 ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Banko 3 ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Banko 2 ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Banko 1 ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Banko Balls :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   440
      Width           =   2075
   End
End
Attribute VB_Name = "frmBanko"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Beep
If Option2.Value = True Then
   Banko(1) = Val(Text1(0).Text)
   If Banko(1) = 0 Then
      Banko(0) = 0
   Else
      Banko(0) = 1
   End If
Else
   If Option3.Value = True Then
      Banko(1) = Val(Text1(0).Text)
      Banko(2) = Val(Text1(1).Text)
      If Banko(1) > 0 Then
         Banko(0) = 1
         If Banko(2) = Banko(1) Then Banko(2) = 0
      Else
         Banko(0) = 0
         Banko(1) = Banko(2)
      End If
      If Banko(2) > 0 Then Banko(0) = Banko(0) + 1
   Else
      If Option4.Value = True Then
         Banko(0) = 0
         Banko(1) = Val(Text1(0).Text)
         Banko(2) = Val(Text1(1).Text)
         Banko(3) = Val(Text1(2).Text)
         If Banko(1) > 0 Then
            Banko(0) = 1
            If Banko(2) = Banko(1) Then Banko(2) = 0
            If Banko(3) = Banko(1) Then Banko(3) = 0
         End If
         If Banko(2) > 0 Then
            Banko(0) = Banko(0) + 1
            If Banko(3) = Banko(2) Then Banko(3) = 0
         End If
         If Banko(3) > 0 Then Banko(0) = Banko(0) + 1
         If Banko(0) <> 3 And Banko(0) > 0 Then
            If Banko(1) = 0 Then
               If Banko(2) > 0 Then
                  Banko(1) = Banko(2)
                  Banko(2) = Banko(3)
               Else
                  Banko(1) = Banko(3)
               End If
            Else
               If Banko(2) = 0 Then Banko(2) = Banko(3)
            End If
         End If
      Else
         If Option5.Value = True Then
            Banko(0) = 0
            Banko(1) = Val(Text1(0).Text)
            Banko(2) = Val(Text1(1).Text)
            Banko(3) = Val(Text1(2).Text)
            Banko(4) = Val(Text1(3).Text)
            If Banko(1) > 0 Then
               Banko(0) = Banko(0) + 1
               If Banko(2) = Banko(1) Then Banko(2) = 0
               If Banko(3) = Banko(1) Then Banko(3) = 0
               If Banko(4) = Banko(1) Then Banko(4) = 0
            End If
            If Banko(2) > 0 Then
               Banko(0) = Banko(0) + 1
               If Banko(3) = Banko(2) Then Banko(3) = 0
               If Banko(4) = Banko(2) Then Banko(4) = 0
            End If
            If Banko(3) > 0 Then
               Banko(0) = Banko(0) + 1
               If Banko(4) = Banko(3) Then Banko(4) = 0
            End If
            If Banko(4) > 0 Then Banko(0) = Banko(0) + 1
            If Banko(0) < 4 And Banko(0) > 0 Then
               Y = 0
               For X = 1 To 4
                   If Banko(X) > 0 Then
                      Y = Y + 1
                      Banko(Y) = Banko(X)
                   End If
               Next X
            End If
         Else
            If Option6.Value = True Then
               Banko(0) = 0
               Banko(1) = Val(Text1(0).Text)
               Banko(2) = Val(Text1(1).Text)
               Banko(3) = Val(Text1(2).Text)
               Banko(4) = Val(Text1(3).Text)
               Banko(5) = Val(Text1(4).Text)
               If Banko(1) > 0 Then
                  Banko(0) = Banko(0) + 1
                  If Banko(2) = Banko(1) Then Banko(2) = 0
                  If Banko(3) = Banko(1) Then Banko(3) = 0
                  If Banko(4) = Banko(1) Then Banko(4) = 0
                  If Banko(5) = Banko(1) Then Banko(5) = 0
               End If
               If Banko(2) > 0 Then
                  Banko(0) = Banko(0) + 1
                  If Banko(3) = Banko(2) Then Banko(3) = 0
                  If Banko(4) = Banko(2) Then Banko(4) = 0
                  If Banko(5) = Banko(2) Then Banko(5) = 0
               End If
               If Banko(3) > 0 Then
                  Banko(0) = Banko(0) + 1
                  If Banko(4) = Banko(3) Then Banko(4) = 0
                  If Banko(5) = Banko(3) Then Banko(5) = 0
               End If
               If Banko(4) > 0 Then
                  Banko(0) = Banko(0) + 1
                  If Banko(5) = Banko(4) Then Banko(5) = 0
               End If
               If Banko(5) > 0 Then Banko(0) = Banko(0) + 1
               If Banko(0) < 5 And Banko(0) > 0 Then
                  Y = 0
                  For X = 1 To 5
                      If Banko(X) > 0 Then
                         Y = Y + 1
                         Banko(Y) = Banko(X)
                      End If
                  Next X
               End If
            Else
               Banko(0) = 0
            End If
         End If
      End If
   End If
End If
For Y = 0 To 4
    Text1(Y).Text = ""
Next Y
BankoGoster
End Sub

Private Sub Command2_Click()
Unload frmBanko
End Sub

Private Sub Command3_Click()
Beep
If Dir("SKORLOTO.DAT") = "SKORLOTO.DAT" Then
   X = FileLen("SKORLOTO.DAT")
   If X > 101 Then
      Q = 21
      frmQMsg.Show
   End If
End If
End Sub

Private Sub Command4_Click()
Command1_Click
Q = 22
frmQMsg.Show
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
   F = 2
   F10Text = "FUNCTION KEYS : " + Chr$(13) + Chr$(13) + "F1=  Apply (Use requested balls as banko/forever)" + Chr$(13) + "F2=  Load (Load banko balls from 'SKORLOTO.DAT')" + Chr$(13) + "F3=  Save (Save banko balls in 'SKORLOTO.DAT')" + Chr$(13) + Chr$(13) + "F4-F9= Number of Banko options" + Chr$(13) + Chr$(13) + "F10= Help" + Chr$(13) + Chr$(13) + "ESC= Return/Close"
   If Q = 21 Then
      For Y = 0 To 4
          Text1(Y).Text = ""
      Next Y
      BankoGoster
   End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_ESCAPE Then
   Unload frmBanko
Else
   If KeyCode = KEY_F1 Then
      Command1_Click
   Else
      If KeyCode = KEY_F2 Then
         Command3_Click
      Else
         If KeyCode = KEY_F3 Then
            Command4_Click
         Else
            If KeyCode = KEY_F4 Then
               Option1.Value = True
               Option1_Click
            Else
               If KeyCode = KEY_F5 Then
                  Option2.Value = True
                  Option2_Click
               Else
                  If KeyCode = KEY_F6 Then
                     Option3.Value = True
                     Option3_Click
                  Else
                     If KeyCode = KEY_F7 Then
                        Option4.Value = True
                        Option4_Click
                     Else
                        If KeyCode = KEY_F8 Then
                           Option5.Value = True
                           Option5_Click
                        Else
                           If KeyCode = KEY_F9 Then
                              Option6.Value = True
                              Option6_Click
                           Else
                              If KeyCode = KEY_F10 Then frmHelp.Show
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
Label10.BackColor = &H7000
Shape1.BorderColor = &H7000
For I = 0 To Label9.Count - 1
    Label9(I).ForeColor = QBColor(0)
Next I
End Sub

Private Sub Form_Resize()
Shape1.Height = frmBanko.Height
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
   frmBanko.Left = frmBanko.Left + (X - MousePosX)
   frmBanko.Top = frmBanko.Top + (Y - MousePosY)
Else
   Label10.BackColor = &H7000
   Shape1.BorderColor = &H7000
   For I = 0 To Label9.Count - 1
       Label9(I).ForeColor = QBColor(0)
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
               Command1_Click
          Case 1
               Command2_Click
          Case 2
               Command3_Click
          Case 3
               Command4_Click
          Case 4
               frmHelp.Show
   End Select
End If
End Sub

Private Sub Label9_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9(Index).ForeColor = QBColor(12)
End Sub

Private Sub Label9_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 99
Label9(Index).ForeColor = QBColor(9)
End Sub

Private Sub Option1_Click()
Option2.ForeColor = QBColor(9)
Option3.ForeColor = QBColor(9)
Option4.ForeColor = QBColor(9)
Option5.ForeColor = QBColor(9)
Option6.ForeColor = QBColor(9)
Option1.ForeColor = QBColor(12)
Label2(0).Visible = False
Text1(0).Visible = False
Label2(1).Visible = False
Text1(1).Visible = False
Label2(2).Visible = False
Text1(2).Visible = False
Label2(3).Visible = False
Text1(3).Visible = False
Label2(4).Visible = False
Text1(4).Visible = False
frmBanko.Height = 2500
End Sub

Private Sub Option2_Click()
Option1.ForeColor = QBColor(9)
Option3.ForeColor = QBColor(9)
Option4.ForeColor = QBColor(9)
Option5.ForeColor = QBColor(9)
Option6.ForeColor = QBColor(9)
Option2.ForeColor = QBColor(12)
Label2(0).Visible = True
Text1(0).Visible = True
Label2(1).Visible = False
Text1(1).Visible = False
Label2(2).Visible = False
Text1(2).Visible = False
Label2(3).Visible = False
Text1(3).Visible = False
Label2(4).Visible = False
Text1(4).Visible = False
frmBanko.Height = 2500
Text1(0).SetFocus
End Sub

Private Sub Option3_Click()
Option1.ForeColor = QBColor(9)
Option2.ForeColor = QBColor(9)
Option4.ForeColor = QBColor(9)
Option5.ForeColor = QBColor(9)
Option6.ForeColor = QBColor(9)
Option3.ForeColor = QBColor(12)
Label2(0).Visible = True
Text1(0).Visible = True
Label2(1).Visible = True
Text1(1).Visible = True
Label2(2).Visible = False
Text1(2).Visible = False
Label2(3).Visible = False
Text1(3).Visible = False
Label2(4).Visible = False
Text1(4).Visible = False
frmBanko.Height = 2500
Text1(0).SetFocus
End Sub

Private Sub Option4_Click()
Option1.ForeColor = QBColor(9)
Option2.ForeColor = QBColor(9)
Option3.ForeColor = QBColor(9)
Option5.ForeColor = QBColor(9)
Option6.ForeColor = QBColor(9)
Option4.ForeColor = QBColor(12)
Label2(0).Visible = True
Text1(0).Visible = True
Label2(1).Visible = True
Text1(1).Visible = True
Label2(2).Visible = True
Text1(2).Visible = True
Label2(3).Visible = False
Text1(3).Visible = False
Label2(4).Visible = False
Text1(4).Visible = False
frmBanko.Height = 3350
Text1(0).SetFocus
End Sub

Private Sub Option5_Click()
Option1.ForeColor = QBColor(9)
Option2.ForeColor = QBColor(9)
Option3.ForeColor = QBColor(9)
Option4.ForeColor = QBColor(9)
Option6.ForeColor = QBColor(9)
Option5.ForeColor = QBColor(12)
Label2(0).Visible = True
Text1(0).Visible = True
Label2(1).Visible = True
Text1(1).Visible = True
Label2(2).Visible = True
Text1(2).Visible = True
Label2(3).Visible = True
Text1(3).Visible = True
Label2(4).Visible = False
Text1(4).Visible = False
frmBanko.Height = 4175
Text1(0).SetFocus
End Sub

Private Sub Option6_Click()
Option1.ForeColor = QBColor(9)
Option2.ForeColor = QBColor(9)
Option3.ForeColor = QBColor(9)
Option4.ForeColor = QBColor(9)
Option5.ForeColor = QBColor(9)
Option6.ForeColor = QBColor(12)
Label2(0).Visible = True
Text1(0).Visible = True
Label2(1).Visible = True
Text1(1).Visible = True
Label2(2).Visible = True
Text1(2).Visible = True
Label2(3).Visible = True
Text1(3).Visible = True
Label2(4).Visible = True
Text1(4).Visible = True
frmBanko.Height = 5000
Text1(0).SetFocus
End Sub

Private Sub Text1_Change(Index As Integer)
Dim Textval As Integer
Textval = Val(Text1(Index).Text)
If Textval = 0 Or Textval > 49 Then
   Text1(Index).Text = ""
Else
   Text1(Index).Text = Textval
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   If Index < 4 Then
      If Text1(Index + 1).Visible = True Then
         Text1(Index + 1).SetFocus
      End If
   End If
End If
End Sub

Private Sub BankoGoster()
Select Case (Banko(0))
       Case 0
             Option1.Value = True
             Option1_Click
       Case 1
             Option2.Value = True
             Option2_Click
             Text1(0).Text = Banko(1)
       Case 2
             Option3.Value = True
             Option3_Click
             Text1(0).Text = Banko(1)
             Text1(1).Text = Banko(2)
       Case 3
             Option4.Value = True
             Option4_Click
             Text1(0).Text = Banko(1)
             Text1(1).Text = Banko(2)
             Text1(2).Text = Banko(3)
       Case 4
             Option5.Value = True
             Option5_Click
             Text1(0).Text = Banko(1)
             Text1(1).Text = Banko(2)
             Text1(2).Text = Banko(3)
             Text1(3).Text = Banko(4)
       Case 5
             Option6.Value = True
             Option6_Click
             Text1(0).Text = Banko(1)
             Text1(1).Text = Banko(2)
             Text1(2).Text = Banko(3)
             Text1(3).Text = Banko(4)
             Text1(4).Text = Banko(5)
       Case Else
             Option1_Click
End Select
End Sub
