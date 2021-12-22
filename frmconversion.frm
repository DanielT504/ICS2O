VERSION 5.00
Begin VB.Form frmConversion 
   Caption         =   "Decimal Conversion"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   14790
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1560
      TabIndex        =   7
      Top             =   11160
      Width           =   5295
   End
   Begin VB.TextBox txtBinary 
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
      Left            =   6480
      TabIndex        =   6
      Top             =   7440
      Width           =   4095
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4320
      TabIndex        =   4
      Top             =   4440
      Width           =   4455
   End
   Begin VB.TextBox txtDecimal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7560
      TabIndex        =   2
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label lbl2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   8
      Top             =   8880
      Width           =   975
   End
   Begin VB.Label lblBinary 
      Caption         =   "Binary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      TabIndex        =   5
      Top             =   7800
      Width           =   3495
   End
   Begin VB.Label lbl10 
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11160
      TabIndex        =   3
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblDecimal 
      Caption         =   "Decimal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1560
      TabIndex        =   1
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label lblTitle 
      Caption         =   "Decimal number conversion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConvert_Click()
number = Val(txtDecimal.Text)

If number >= 128 Then
    number = number - 128
    num1 = 1
Else: num1 = 0
End If

If number >= 64 Then
    number = number - 64
    num2 = 1
Else: num2 = 0
End If

If number >= 32 Then
    number = number - 32
    num3 = 1
Else: num3 = 0
End If

If number >= 16 Then
    number = number - 16
    num4 = 1
Else: num4 = 0
End If

If number >= 8 Then
    number = number - 8
    num5 = 1
Else: num5 = 0
End If


If number >= 4 Then
    number = number - 4
    num6 = 1
Else: num6 = 0
End If

If number >= 2 Then
    number = number - 2
    num7 = 1
Else: num7 = 0
End If

If number >= 1 Then
    number = number - 1
    num8 = 1
Else: num8 = 0
End If

txtBinary.Text = num1 & num2 & num3 & num4 & " " & num5 & num6 & num7 & num8
End Sub

Private Sub cmdReturn_Click()
    Unload frmConversion
End Sub

Private Sub Form_Load()
Dim num1 As Integer
Dim num2 As Integer
Dim num3 As Integer
Dim num4 As Integer
Dim num5 As Integer
Dim num6 As Integer
Dim num7 As Integer
Dim num8 As Integer
Dim number As Integer

End Sub
