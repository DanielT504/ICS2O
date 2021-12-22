VERSION 5.00
Begin VB.Form frmLogicGates 
   Caption         =   "Logic Gates"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   14790
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh1 
      Caption         =   "Refresh intermiediate outputs"
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
      Left            =   4680
      TabIndex        =   38
      Top             =   5160
      Width           =   3615
   End
   Begin VB.PictureBox picXOR2 
      Height          =   855
      Left            =   6120
      Picture         =   "frmLogcGates.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   37
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picXNOR2 
      Height          =   855
      Left            =   6120
      Picture         =   "frmLogcGates.frx":2F22
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   36
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picOR2 
      Height          =   855
      Left            =   6120
      Picture         =   "frmLogcGates.frx":5E44
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   35
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picNOR2 
      Height          =   855
      Left            =   6120
      Picture         =   "frmLogcGates.frx":8D66
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   34
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picNAND2 
      Height          =   855
      Left            =   6120
      Picture         =   "frmLogcGates.frx":BC88
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   33
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picXOR3 
      Height          =   855
      Left            =   11160
      Picture         =   "frmLogcGates.frx":EBAA
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   32
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picXNOR3 
      Height          =   855
      Left            =   11160
      Picture         =   "frmLogcGates.frx":11ACC
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   31
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picOR3 
      Height          =   855
      Left            =   11160
      Picture         =   "frmLogcGates.frx":149EE
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   30
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picNOR3 
      Height          =   855
      Left            =   11160
      Picture         =   "frmLogcGates.frx":17910
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   29
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picNAND3 
      Height          =   855
      Left            =   11160
      Picture         =   "frmLogcGates.frx":1A832
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   28
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picXOR1 
      Height          =   855
      Left            =   6000
      Picture         =   "frmLogcGates.frx":1D754
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   27
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picXNOR1 
      Height          =   855
      Left            =   6000
      Picture         =   "frmLogcGates.frx":20676
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picOR1 
      Height          =   855
      Left            =   6000
      Picture         =   "frmLogcGates.frx":23598
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   25
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picNOR1 
      Height          =   855
      Left            =   6000
      Picture         =   "frmLogcGates.frx":264BA
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   24
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picNAND1 
      Height          =   855
      Left            =   6000
      Picture         =   "frmLogcGates.frx":293DC
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   23
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbGate3 
      Height          =   315
      ItemData        =   "frmLogcGates.frx":2C2FE
      Left            =   11160
      List            =   "frmLogcGates.frx":2C314
      TabIndex        =   14
      Text            =   "Choose gate"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.ComboBox cmbGate2 
      Height          =   315
      ItemData        =   "frmLogcGates.frx":2C337
      Left            =   6120
      List            =   "frmLogcGates.frx":2C34D
      TabIndex        =   13
      Text            =   "Choose gate"
      Top             =   8160
      Width           =   1335
   End
   Begin VB.ComboBox cmbGate1 
      Height          =   315
      ItemData        =   "frmLogcGates.frx":2C370
      Left            =   6000
      List            =   "frmLogcGates.frx":2C386
      TabIndex        =   12
      Text            =   "Choose gate"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to main form"
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
      Left            =   9000
      TabIndex        =   11
      Top             =   10200
      Width           =   4335
   End
   Begin VB.PictureBox picAND3 
      Height          =   855
      Left            =   11160
      Picture         =   "frmLogcGates.frx":2C3A9
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picAND2 
      Height          =   855
      Left            =   6120
      Picture         =   "frmLogcGates.frx":2F2CB
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picAND1 
      Height          =   855
      Left            =   6000
      Picture         =   "frmLogcGates.frx":321ED
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraInput4 
      Caption         =   "Input 4"
      Height          =   1815
      Left            =   360
      TabIndex        =   4
      Top             =   8880
      Width           =   2415
      Begin VB.OptionButton optFalse4 
         Caption         =   "False"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton optTrue4 
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame fraInput3 
      Caption         =   "Input 3"
      Height          =   1695
      Left            =   360
      TabIndex        =   3
      Top             =   6600
      Width           =   2415
      Begin VB.OptionButton optFalse3 
         Caption         =   "False"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton optTrue3 
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraInput2 
      Caption         =   "Input 2"
      Height          =   1935
      Left            =   360
      TabIndex        =   2
      Top             =   4080
      Width           =   2295
      Begin VB.OptionButton optFalse2 
         Caption         =   "False"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optTrue2 
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame fraInput1 
      Caption         =   "Input 1"
      Height          =   1815
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
      Begin VB.OptionButton optFalse1 
         Caption         =   "False"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optTrue1 
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label lblIntermediateOutput2 
      Height          =   615
      Left            =   8760
      TabIndex        =   10
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label lblIntermediateOutput1 
      Height          =   615
      Left            =   8880
      TabIndex        =   9
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblOutput 
      Height          =   855
      Left            =   12720
      TabIndex        =   8
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Line Line6 
      X1              =   7320
      X2              =   11160
      Y1              =   7680
      Y2              =   5760
   End
   Begin VB.Line Line5 
      X1              =   7200
      X2              =   11280
      Y1              =   3600
      Y2              =   5640
   End
   Begin VB.Line Line4 
      X1              =   2880
      X2              =   6120
      Y1              =   9840
      Y2              =   7800
   End
   Begin VB.Line Line3 
      X1              =   2880
      X2              =   6120
      Y1              =   7440
      Y2              =   7560
   End
   Begin VB.Line Line2 
      X1              =   2760
      X2              =   6000
      Y1              =   5280
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   6000
      Y1              =   2760
      Y2              =   3480
   End
   Begin VB.Label lblTitle 
      Caption         =   "Logic gates"
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
      Left            =   5400
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmLogicGates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdRefresh1_Click()

If cmbGate1.Text = "AND" Then
    picAND1.Visible = True
    If optTrue1 = True And optTrue2 = True Then
        lblIntermediateOutput1.Caption = "TRUE"
    ElseIf optTrue1 = True And optTrue2 = False Then
        lblIntermediateOutput1.Caption = "FALSE"
    ElseIf optTrue1 = False And optTrue2 = True Then
        lblIntermediateOutput1.Caption = "FALSE"
    ElseIf optTrue1 = False And optTrue2 = False Then
        lblIntermediateOutput1.Caption = "FALSE"
    End If
ElseIf cmbGate1.Text = "NAND" Then
    picNAND1.Visible = True
    If optTrue1 = True And optTrue2 = True Then
        lblIntermediateOutput1.Caption = "FALSE"
    ElseIf optTrue1 = True And optTrue2 = False Then
        lblIntermediateOutput1.Caption = "TRUE"
    ElseIf optTrue1 = False And optTrue2 = True Then
        lblIntermediateOutput1.Caption = "TRUE"
    ElseIf optTrue1 = False And optTrue2 = False Then
        lblIntermediateOutput1.Caption = "TRUE"
    End If
ElseIf cmbGate1.Text = "OR" Then
    picOR1.Visible = True
    If optTrue1 = True And optTrue2 = True Then
        lblIntermediateOutput1.Caption = "TRUE"
    ElseIf optTrue1 = True And optTrue2 = False Then
        lblIntermediateOutput1.Caption = "TRUE"
    ElseIf optTrue1 = False And optTrue2 = True Then
        lblIntermediateOutput1.Caption = "TRUE"
    ElseIf optTrue1 = False And optTrue2 = False Then
        lblIntermediateOutput1.Caption = "FALSE"
    End If
ElseIf cmbGate1.Text = "NOR" Then
    picNOR1.Visible = True
    If optTrue1 = True And optTrue2 = True Then
        lblIntermediateOutput1.Caption = "TRUE"
    ElseIf optTrue1 = True And optTrue2 = False Then
        lblIntermediateOutput1.Caption = "FALSE"
    ElseIf optTrue1 = False And optTrue2 = True Then
        lblIntermediateOutput1.Caption = "FALSE"
    ElseIf optTrue1 = False And optTrue2 = False Then
        lblIntermediateOutput1.Caption = "FALSE"
    End If
ElseIf cmbGate1.Text = "XOR" Then
    picXOR1.Visible = True
    If optTrue1 = True And optTrue2 = True Then
        lblIntermediateOutput1.Caption = "FALSE"
    ElseIf optTrue1 = True And optTrue2 = False Then
        lblIntermediateOutput1.Caption = "TRUE"
    ElseIf optTrue1 = False And optTrue2 = True Then
        lblIntermediateOutput1.Caption = "TRUE"
    ElseIf optTrue1 = False And optTrue2 = False Then
        lblIntermediateOutput1.Caption = "FALSE"
    End If
ElseIf cmbGate1.Text = "XNOR" Then
    picXNOR1.Visible = True
    If optTrue1 = True And optTrue2 = True Then
        lblIntermediateOutput1.Caption = "TRUE"
    ElseIf optTrue1 = True And optTrue2 = False Then
        lblIntermediateOutput1.Caption = "FALSE"
    ElseIf optTrue1 = False And optTrue2 = True Then
        lblIntermediateOutput1.Caption = "FALSE"
    ElseIf optTrue1 = False And optTrue2 = False Then
        lblIntermediateOutput1.Caption = "TRUE"
    End If
End If

If cmbGate2.Text = "AND" Then
    picAND2.Visible = True
    If optTrue2 = True And optTrue3 = True Then
        lblIntermediateOutput2.Caption = "TRUE"
    ElseIf optTrue2 = True And optTrue3 = False Then
        lblIntermediateOutput2.Caption = "FALSE"
    ElseIf optTrue2 = False And optTrue3 = True Then
        lblIntermediateOutput2.Caption = "FALSE"
    ElseIf optTrue2 = False And optTrue3 = False Then
        lblIntermediateOutput2.Caption = "FALSE"
    End If
ElseIf cmbGate2.Text = "NAND" Then
    picNAND2.Visible = True
    If optTrue2 = True And optTrue3 = True Then
        lblIntermediateOutput2.Caption = "FALSE"
    ElseIf optTrue2 = True And optTrue3 = False Then
        lblIntermediateOutput2.Caption = "TRUE"
    ElseIf optTrue2 = False And optTrue3 = True Then
        lblIntermediateOutput2.Caption = "TRUE"
    ElseIf optTrue2 = False And optTrue3 = False Then
        lblIntermediateOutput2.Caption = "TRUE"
    End If
ElseIf cmbGate2.Text = "OR" Then
    picOR2.Visible = True
    If optTrue2 = True And optTrue3 = True Then
        lblIntermediateOutput2.Caption = "TRUE"
    ElseIf optTrue2 = True And optTrue3 = False Then
        lblIntermediateOutput2.Caption = "TRUE"
    ElseIf optTrue2 = False And optTrue3 = True Then
        lblIntermediateOutput2.Caption = "TRUE"
    ElseIf optTrue2 = False And optTrue3 = False Then
        lblIntermediateOutput2.Caption = "FALSE"
    End If
ElseIf cmbGate2.Text = "NOR" Then
    picNOR2.Visible = True
    If optTrue2 = True And optTrue3 = True Then
        lblIntermediateOutput2.Caption = "FALSE"
    ElseIf optTrue2 = True And optTrue3 = False Then
        lblIntermediateOutput2.Caption = "FALSE"
    ElseIf optTrue2 = False And optTrue3 = True Then
        lblIntermediateOutput2.Caption = "FALSE"
    ElseIf optTrue2 = False And optTrue3 = False Then
        lblIntermediateOutput2.Caption = "TRUE"
    End If
ElseIf cmbGate2.Text = "XOR" Then
    picXOR2.Visible = True
    If optTrue2 = True And optTrue3 = True Then
        lblIntermediateOutput2.Caption = "FALSE"
    ElseIf optTrue2 = True And optTrue3 = False Then
        lblIntermediateOutput2.Caption = "TRUE"
    ElseIf optTrue2 = False And optTrue3 = True Then
        lblIntermediateOutput2.Caption = "TRUE"
    ElseIf optTrue2 = False And optTrue3 = False Then
        lblIntermediateOutput2.Caption = "FALSE"
    End If
ElseIf cmbGate2.Text = "XNOR" Then
    picXNOR2.Visible = True
    If optTrue2 = True And optTrue3 = True Then
        lblIntermediateOutput2.Caption = "TRUE"
    ElseIf optTrue2 = True And optTrue3 = False Then
        lblIntermediateOutput2.Caption = "FALSE"
    ElseIf optTrue2 = False And optTrue3 = True Then
        lblIntermediateOutput2.Caption = "FALSE"
    ElseIf optTrue2 = False And optTrue3 = False Then
        lblIntermediateOutput2.Caption = "TRUE"
    End If
End If

If cmbGate3.Text = "AND" Then
    picAND3.Visible = True
    If lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "FALSE"
    End If
ElseIf cmbGate3.Text = "NAND" Then
    picNAND3.Visible = True
    If lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "TRUE"
    End If
ElseIf cmbGate3.Text = "OR" Then
    picOR3.Visible = True
    If lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "FALSE"
    End If
ElseIf cmbGate3.Text = "NOR" Then
    picNOR3.Visible = True
    If lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "TRUE"
    End If
ElseIf cmbGate3.Text = "XOR" Then
    picXOR3.Visible = True
    If lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "FALSE"
    End If
ElseIf cmbGate3.Text = "XNOR" Then
    picXNOR3.Visible = True
    If lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "TRUE"
    End If
End If
End Sub

Private Sub cmdRefresh2_Click()
If cmbGate3.Text = "AND" Then
    picAND3.Visible = True
    If lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "FALSE"
    End If
ElseIf cmbGate3.Text = "NAND" Then
    picNAND3.Visible = True
    If lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "TRUE"
    End If
ElseIf cmbGate3.Text = "OR" Then
    picOR3.Visible = True
    If lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "FALSE"
    End If
ElseIf cmbGate3.Text = "NOR" Then
    picNOR3.Visible = True
    If lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "TRUE"
    End If
ElseIf cmbGate3.Text = "XOR" Then
    picXOR3.Visible = True
    If lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "FALSE"
    End If
ElseIf cmbGate3.Text = "XNOR" Then
    picXNOR3.Visible = True
    If lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "TRUE"
    ElseIf lblIntermediateOutput1.Caption = "TRUE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "TRUE" Then
        lblOutput.Caption = "FALSE"
    ElseIf lblIntermediateOutput1.Caption = "FALSE" And lblIntermediateOutput2.Caption = "FALSE" Then
        lblOutput.Caption = "TRUE"
    End If
End If
End Sub

Private Sub cmdReturn_Click()
    Unload frmLogicGates
End Sub

Private Sub Form_Load()
Dim input1 As Boolean
Dim input2 As Boolean
Dim input3 As Boolean
Dim input4 As Boolean
Dim output1 As Boolean
Dim output2 As Boolean

input1 = True
input2 = True
input3 = True
input4 = True
End Sub
