VERSION 5.00
Begin VB.Form frmGuessingGame 
   Caption         =   "Guessing Game"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9345
   FillColor       =   &H80000012&
   ForeColor       =   &H80000016&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   9345
   Begin VB.TextBox txtTry 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   12
      Text            =   "0"
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdExt 
      Caption         =   "Back to main"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   10
      Top             =   6240
      Width           =   3015
   End
   Begin VB.CommandButton cmdGss 
      Caption         =   "Guess"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   9
      Top             =   5160
      Width           =   2175
   End
   Begin VB.HScrollBar hsbNum 
      Height          =   855
      Left            =   360
      Max             =   100
      Min             =   1
      TabIndex        =   8
      Top             =   4080
      Value           =   1
      Width           =   8055
   End
   Begin VB.TextBox txtNum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   5
      Top             =   3000
      Width           =   3375
   End
   Begin VB.CommandButton cmdGam 
      Caption         =   "Start game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   8535
   End
   Begin VB.TextBox txtMax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblRcp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   14
      Top             =   6120
      Width           =   4095
   End
   Begin VB.Label lblTtl 
      Caption         =   "Set range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblTry 
      Caption         =   "Guess #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   11
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lblHigh 
      Caption         =   "Too high"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblLow 
      Caption         =   "Too low"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblMax 
      Caption         =   "Max:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblMin 
      Caption         =   "Min:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmGuessingGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExt_Click()
    Unload frmGuessingGame
End Sub

Private Sub Form_Load()
    Randomize Timer
    ' Declaring variables
    Dim number As Integer
    Dim guesses As Integer
End Sub

Private Sub cmdGam_Click()
    lblRcp.Caption = ""
    ' Creating random number
    number = Int(Rnd * (Val(txtMax) - Val(txtMin)) + Val(txtMin))
    ' Resetting guesses
    guesses = 0
    ' Setting range on scrollbar
    hsbNum.Min = Val(txtMin)
    hsbNum.Max = Val(txtMax)
    hsbNum.Value = Val(txtMin)
End Sub

Private Sub cmdGss_Click()
    ' Output depending on how high or low guess is
    If Val(txtNum.Text) < number Then
        lblLow.Visible = True
        lblHigh.Visible = False
    ElseIf Val(txtNum.Text) > number Then
        lblHigh.Visible = True
        lblLow.Visible = False
    ElseIf Val(txtNum.Text) = number Then
        lblHigh.Visible = False
        lblLow.Visible = False
        lblRcp.Caption = "Congratulations, you guessed the number"
    End If
    ' Adding a guess to the counter
    guesses = guesses + 1
    txtTry.Text = guesses
End Sub

Private Sub hsbNum_Change()
    ' Displaying number from scrollbar
    txtNum.Text = hsbNum.Value
End Sub
