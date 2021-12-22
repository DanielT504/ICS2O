VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4080
      TabIndex        =   5
      Top             =   7920
      Width           =   5415
   End
   Begin VB.CommandButton cmdLogicGates 
      Caption         =   "Logic Gates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7200
      TabIndex        =   4
      Top             =   4680
      Width           =   5415
   End
   Begin VB.CommandButton cmdDecimalConversion 
      Caption         =   "Decimal Conversion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1680
      TabIndex        =   3
      Top             =   4680
      Width           =   4575
   End
   Begin VB.CommandButton cmdComputerQuiz 
      Caption         =   "Computer Quiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7440
      TabIndex        =   2
      Top             =   1680
      Width           =   5055
   End
   Begin VB.CommandButton cmdGuessinggame 
      Caption         =   "Guessing game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label lblTitle 
      Caption         =   "Daniel Thero Summative June 2016"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdComputerQuiz_Click()
    frmComputerquiz.Show vbModal
End Sub

Private Sub cmdDecimalConversion_Click()
    frmConversion.Show vbModal
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdGuessinggame_Click()
    frmGuessingGame.Show vbModal
End Sub

Private Sub cmdLogicGates_Click()
    frmLogicGates.Show vbModal
End Sub
