VERSION 5.00
Begin VB.Form frmComputerQuiz 
   Caption         =   "Computer Quiz"
   ClientHeight    =   11715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   ScaleHeight     =   11715
   ScaleWidth      =   14010
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCorrect1 
      Height          =   975
      Left            =   10560
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   34
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   1080
         X2              =   600
         Y1              =   240
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   240
         X2              =   600
         Y1              =   480
         Y2              =   720
      End
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit answers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3840
      TabIndex        =   33
      Top             =   10080
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear answers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1920
      TabIndex        =   32
      Top             =   10080
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Back to main"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   31
      Top             =   10080
      Width           =   1575
   End
   Begin VB.PictureBox picWrong5 
      Height          =   1215
      Left            =   12240
      ScaleHeight     =   1155
      ScaleWidth      =   1395
      TabIndex        =   29
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Line Line21 
         BorderWidth     =   3
         X1              =   1200
         X2              =   240
         Y1              =   240
         Y2              =   960
      End
      Begin VB.Line Line20 
         BorderWidth     =   3
         X1              =   120
         X2              =   1080
         Y1              =   240
         Y2              =   960
      End
   End
   Begin VB.PictureBox picCorrect5 
      Height          =   1095
      Left            =   10680
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   28
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Line Line19 
         BorderWidth     =   3
         X1              =   1080
         X2              =   480
         Y1              =   240
         Y2              =   960
      End
      Begin VB.Line Line18 
         BorderWidth     =   3
         X1              =   0
         X2              =   480
         Y1              =   600
         Y2              =   960
      End
   End
   Begin VB.PictureBox picWrong4 
      Height          =   1215
      Left            =   12480
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   27
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
      Begin VB.Line Line17 
         BorderWidth     =   3
         X1              =   1080
         X2              =   240
         Y1              =   240
         Y2              =   1080
      End
      Begin VB.Line Line16 
         BorderWidth     =   3
         X1              =   240
         X2              =   1080
         Y1              =   120
         Y2              =   1080
      End
   End
   Begin VB.PictureBox picCorrect4 
      Height          =   1215
      Left            =   10800
      ScaleHeight     =   1155
      ScaleWidth      =   1395
      TabIndex        =   26
      Top             =   6360
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Line Line15 
         BorderWidth     =   3
         X1              =   1080
         X2              =   480
         Y1              =   360
         Y2              =   960
      End
      Begin VB.Line Line14 
         BorderWidth     =   3
         X1              =   120
         X2              =   480
         Y1              =   600
         Y2              =   960
      End
   End
   Begin VB.PictureBox picWrong3 
      Height          =   1095
      Left            =   12240
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   25
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Line Line13 
         BorderWidth     =   3
         X1              =   1080
         X2              =   120
         Y1              =   240
         Y2              =   840
      End
      Begin VB.Line Line12 
         BorderWidth     =   3
         X1              =   120
         X2              =   1080
         Y1              =   240
         Y2              =   840
      End
   End
   Begin VB.PictureBox picCorrect3 
      Height          =   1095
      Left            =   10560
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
      Begin VB.Line Line11 
         BorderWidth     =   3
         X1              =   1200
         X2              =   600
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Line Line10 
         BorderWidth     =   3
         X1              =   120
         X2              =   600
         Y1              =   480
         Y2              =   840
      End
   End
   Begin VB.PictureBox picWrong2 
      Height          =   855
      Left            =   12360
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Line Line9 
         BorderWidth     =   3
         X1              =   1080
         X2              =   240
         Y1              =   120
         Y2              =   600
      End
      Begin VB.Line Line8 
         BorderWidth     =   3
         X1              =   120
         X2              =   960
         Y1              =   120
         Y2              =   720
      End
   End
   Begin VB.PictureBox picCorrect2 
      Height          =   975
      Left            =   10680
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   22
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
      Begin VB.Line Line5 
         BorderWidth     =   3
         X1              =   1080
         X2              =   600
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   240
         X2              =   600
         Y1              =   480
         Y2              =   720
      End
   End
   Begin VB.PictureBox picWrong1 
      Height          =   855
      Left            =   12240
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   21
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Line Line7 
         BorderWidth     =   3
         X1              =   1080
         X2              =   240
         Y1              =   360
         Y2              =   600
      End
      Begin VB.Line Line6 
         BorderWidth     =   3
         X1              =   240
         X2              =   1080
         Y1              =   240
         Y2              =   720
      End
   End
   Begin VB.Frame fraAnswer5 
      Height          =   1335
      Left            =   6240
      TabIndex        =   10
      Top             =   7920
      Width           =   3975
      Begin VB.OptionButton optFalse5 
         Caption         =   "False"
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optTrue5 
         Caption         =   "True"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame fraAnswer4 
      Height          =   1335
      Left            =   6240
      TabIndex        =   9
      Top             =   5880
      Width           =   3975
      Begin VB.OptionButton optFalse4 
         Caption         =   "False"
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optTrue4 
         Caption         =   "True"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame fraAnswer3 
      Height          =   1215
      Left            =   6240
      TabIndex        =   8
      Top             =   4200
      Width           =   3975
      Begin VB.OptionButton optFalse3 
         Caption         =   "False"
         Height          =   195
         Left            =   2040
         TabIndex        =   16
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton optTrue3 
         Caption         =   "True"
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame fraAnswer2 
      Height          =   1215
      Left            =   6240
      TabIndex        =   7
      Top             =   2400
      Width           =   3975
      Begin VB.OptionButton optFalse2 
         Caption         =   "False"
         Height          =   195
         Left            =   2160
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optTrue2 
         Caption         =   "True"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame fraAnswer1 
      Height          =   975
      Left            =   6240
      TabIndex        =   6
      Top             =   1200
      Width           =   3975
      Begin VB.OptionButton optFalse1 
         Caption         =   "False"
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optTrue1 
         Caption         =   "True"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Line Line2 
      X1              =   5760
      X2              =   5880
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblRecap 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6240
      TabIndex        =   30
      Top             =   10200
      Width           =   7455
   End
   Begin VB.Label lblTitle 
      Caption         =   "Computer science quiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label lblQuestion5 
      Caption         =   "The first computer was made in 1946."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      TabIndex        =   4
      Top             =   7920
      Width           =   4095
   End
   Begin VB.Label lblQuestion4 
      Caption         =   "Graphic cards are never removable."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      TabIndex        =   3
      Top             =   6000
      Width           =   4095
   End
   Begin VB.Label lblQuestion3 
      Caption         =   "You can store RAM on a USB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      TabIndex        =   2
      Top             =   4320
      Width           =   4095
   End
   Begin VB.Label lblQuestion2 
      Caption         =   "Transistors use silicon."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      TabIndex        =   1
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label lblQuestion1 
      Caption         =   "Charles Babbage originated the idea for the first computer."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   4095
   End
End
Attribute VB_Name = "frmComputerquiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdReturn_Click()
    Unload frmComputerquiz
End Sub

Private Sub Form_Load()
optTrue1 = False
optTrue2 = False
optTrue3 = False
optTrue4 = False
optTrue5 = False

optFalse1 = False
optFalse2 = False
optFalse3 = False
optFalse4 = False
optFalse5 = False

Dim Score As Integer

Score = 0
End Sub

Private Sub cmdClear_Click()
optTrue1 = False
optTrue2 = False
optTrue3 = False
optTrue4 = False
optTrue5 = False

optFalse1 = False
optFalse2 = False
optFalse3 = False
optFalse4 = False
optFalse5 = False

picCorrect1.Visible = False
picCorrect2.Visible = False
picCorrect3.Visible = False
picCorrect4.Visible = False
picCorrect5.Visible = False

picWrong1.Visible = False
picWrong2.Visible = False
picWrong3.Visible = False
picWrong4.Visible = False
picWrong5.Visible = False

lblRecap.Caption = ""

Score = 0
End Sub

Private Sub cmdSubmit_Click()
If optTrue1 = True Then
    picCorrect1.Visible = True
    Score = Score + 1
Else: picWrong1.Visible = True
End If

If optTrue2 = True Then
    picCorrect2.Visible = True
    Score = Score + 1
Else: picWrong2.Visible = True
End If

If optTrue3 = True Then
    picWrong3.Visible = True
Else: picCorrect3.Visible = True
    Score = Score + 1
End If

If optTrue4 = True Then
    picWrong4.Visible = True
Else: picCorrect4.Visible = True
    Score = Score + 1
End If

If optTrue5 = True Then
    picCorrect5.Visible = True
    Score = Score + 1
Else: picWrong5.Visible = True
End If

lblRecap.Caption = "You scored " & Score * 20 & "%"
End Sub
