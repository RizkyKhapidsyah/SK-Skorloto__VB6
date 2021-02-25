VERSION 5.00
Begin VB.Form frmBalls 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "6/49 Balls"
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "BALLS.frx":0000
   MousePointer    =   99  'Custom
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2805
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "49"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   48
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   48
      Top             =   1800
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "48"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   47
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   47
      Top             =   1800
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "47"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   46
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   46
      Top             =   1800
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "46"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   45
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   45
      Top             =   1800
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "45"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   44
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   44
      Top             =   1800
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "44"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   43
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   43
      Top             =   1800
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "43"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   42
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   42
      Top             =   1800
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "42"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   41
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   41
      Top             =   1800
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "41"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   40
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   40
      Top             =   1800
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   39
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   39
      Top             =   1800
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "39"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   38
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   38
      Top             =   1440
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   37
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   37
      Top             =   1440
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "37"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   36
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   36
      Top             =   1440
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "36"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   35
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   35
      Top             =   1440
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   34
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   34
      Top             =   1440
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "34"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   33
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   33
      Top             =   1440
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "33"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   32
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   32
      Top             =   1440
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   31
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   31
      Top             =   1440
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   30
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   30
      Top             =   1440
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   29
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   29
      Top             =   1440
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "29"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   28
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   28
      Top             =   1080
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "28"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   27
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   27
      Top             =   1080
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "27"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   26
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   26
      Top             =   1080
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "26"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   25
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   25
      Top             =   1080
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   24
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   24
      Top             =   1080
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   23
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   23
      Top             =   1080
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   22
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   22
      Top             =   1080
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   21
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   21
      Top             =   1080
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   20
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   20
      Top             =   1080
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   19
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   19
      Top             =   1080
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   18
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   18
      Top             =   720
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   17
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   17
      Top             =   720
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   16
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   16
      Top             =   720
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   15
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   15
      Top             =   720
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   14
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   14
      Top             =   720
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   13
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   13
      Top             =   720
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   12
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   12
      Top             =   720
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   11
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   11
      Top             =   720
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   10
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   10
      Top             =   720
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   9
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   9
      Top             =   720
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "09"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   8
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   8
      Top             =   360
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "08"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   7
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   7
      Top             =   360
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "07"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   6
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   6
      Top             =   360
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "06"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   5
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   360
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "05"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   4
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   360
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "04"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   3
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   360
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   360
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   360
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   360
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F0FFFF&
      BorderColor     =   &H000040C0&
      Height          =   2805
      Left            =   0
      Top             =   0
      Width           =   7260
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   55
      ToolTipText     =   "Click here to use selected balls"
      Top             =   2275
      Width           =   735
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
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   54
      ToolTipText     =   "Click here to load selected balls from 'SKORLOTO.DAT'"
      Top             =   2275
      Width           =   735
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
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   53
      ToolTipText     =   "Click here to save selected balls into 'SKORLOTO.DAT'"
      Top             =   2275
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Info"
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
      Index           =   3
      Left            =   4800
      TabIndex        =   52
      ToolTipText     =   "Click here to get program details"
      Top             =   2275
      Width           =   495
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
      Height          =   255
      Index           =   4
      Left            =   6360
      TabIndex        =   51
      ToolTipText     =   "Click here to close ball selection window"
      Top             =   2275
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   2240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   375
      Index           =   1
      Left            =   1560
      Top             =   2240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   375
      Index           =   2
      Left            =   3120
      Top             =   2240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   375
      Index           =   3
      Left            =   4680
      Top             =   2240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   375
      Index           =   4
      Left            =   6360
      Top             =   2240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C8F4FF&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   0
      Top             =   2275
      Width           =   7335
   End
   Begin VB.Label Label10 
      BackColor       =   &H000040C0&
      Caption         =   "  6/49 Balls"
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
      TabIndex        =   50
      Top             =   10
      Width           =   7260
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F0FFFF&
      Caption         =   "Balls"
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
      Left            =   120
      TabIndex        =   49
      Top             =   420
      Width           =   615
   End
End
Attribute VB_Name = "frmBalls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click(Index As Integer)
If Check1(Index).Value = 1 Then
   Check1(Index).ForeColor = QBColor(12)
Else
   Check1(Index).ForeColor = QBColor(8)
End If
End Sub

Private Sub Command1_Click()
Beep
Bag(0) = 0
Y = 0
For X = 0 To 48
    If Check1(X).Value = 1 Then
       Y = Y + 1
       Bag(0) = Y
       Bag(Y) = X + 1
    End If
Next X
End Sub

Private Sub Command2_Click()
Unload frmBalls
End Sub

Private Sub Command3_Click()
Beep
If Dir("SKORLOTO.DAT") = "SKORLOTO.DAT" Then
   Y = FileLen("SKORLOTO.DAT")
   If Y > 2 Then
      Q = 11
      frmQMsg.Show
   End If
End If
End Sub

Private Sub Command4_Click()
Command1_Click
Q = 12
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
   F = 1
   F10Text = "FUNCTION KEYS : " + Chr$(13) + Chr$(13) + "F1=  Apply (Put checked balls into the bag)" + Chr$(13) + "F2=  Load (Load balls from 'SKORLOTO.DAT')" + Chr$(13) + "F3=  Save (Save checked balls in 'SKORLOTO.DAT')" + Chr$(13) + Chr$(13) + "F10= Help" + Chr$(13) + Chr$(13) + "ESC= Return/Close"
   If Q = 11 Then
      For X = 0 To 48
          Check1(X).Value = 0
      Next X
      For X = 1 To Bag(0)
          If Bag(X) > 0 And Bag(X) < 50 Then
             Check1(Bag(X) - 1).Value = 1
          End If
      Next X
   End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_ESCAPE Then
   Unload frmBalls
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
            If KeyCode = KEY_F10 Then frmHelp.Show
         End If
      End If
   End If
End If
End Sub

Private Sub Form_Load()
If Bag(0) > 0 Then
   For X = 1 To Bag(0)
       Check1(Bag(X) - 1).Value = 1
   Next X
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 0
Label10.BackColor = &H40C0&
Shape1.BorderColor = &H40C0&
For I = 0 To Shape2.Count - 1
    Shape2(I).Visible = False
    Label9(I).ForeColor = QBColor(0)
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
   frmBalls.Left = frmBalls.Left + (X - MousePosX)
   frmBalls.Top = frmBalls.Top + (Y - MousePosY)
Else
   Label10.BackColor = &H40C0&
   Shape1.BorderColor = &H40C0&
   For I = 0 To Shape2.Count - 1
       Shape2(I).Visible = False
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
               Command3_Click
          Case 2
               Command4_Click
          Case 3
               frmHelp.Show
          Case 4
               Command2_Click
   End Select
End If
End Sub

Private Sub Label9_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9(Index).ForeColor = QBColor(12)
End Sub

Private Sub Label9_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 99
Shape2(Index).Visible = True
End Sub

