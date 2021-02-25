Attribute VB_Name = "Main"
Option Explicit

Global Const KEY_UP = &H26
Global Const KEY_DOWN = &H28
Global Const KEY_PRIOR = &H21
Global Const KEY_NEXT = &H22
Global Const KEY_HOME = &H24
Global Const KEY_END = &H23
Global Const KEY_F1 = &H70
Global Const KEY_F2 = &H71
Global Const KEY_F3 = &H72
Global Const KEY_F4 = &H73
Global Const KEY_F5 = &H74
Global Const KEY_F6 = &H75
Global Const KEY_F7 = &H76
Global Const KEY_F8 = &H77
Global Const KEY_F9 = &H78
Global Const KEY_F10 = &H79
Global Const KEY_F11 = &H7A
Global Const KEY_F12 = &H7B
Global Const KEY_ESCAPE = &H1B

Global Banko(5) As Integer
Global Bag(49) As Integer
Global Numara(6) As Integer
Global X As Integer
Global Y As Integer
Global I As Integer
Global F As Integer
Global M As Integer
Global Q As Integer
Global FileNo As Integer
Global MousePosX As Single
Global MousePosY As Single
Global F10Text As String
Global MsgTxt As String

Public Function GetRandomSelectedBalls(BallCell As Integer) As Integer
Dim R As Integer
If Bag(0) = 0 Then
   If Banko(0) > 0 Then
      If Banko(0) >= BallCell Then
         GetRandomSelectedBalls = Banko(BallCell)
         Exit Function
      End If
      For R = 0 To 48
          For X = 1 To Banko(0)
              If Banko(X) = R + 1 Then GoTo GetNextBall
          Next X
          Bag(0) = Bag(0) + 1
          Bag(Bag(0)) = R + 1
GetNextBall:
      Next R
   Else
      For R = 0 To 48
          Bag(0) = Bag(0) + 1
          Bag(Bag(0)) = R + 1
      Next R
   End If
End If
If Bag(0) >= 6 Then
   If Banko(0) < BallCell Then
      R = Int(Bag(0) * Rnd) + 1
      GetRandomSelectedBalls = Bag(R)
   Else
      GetRandomSelectedBalls = Banko(BallCell)
   End If
End If
End Function
