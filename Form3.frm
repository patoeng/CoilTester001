VERSION 5.00
Begin VB.Form FrmDebug 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calibration Mode"
   ClientHeight    =   3570
   ClientLeft      =   6735
   ClientTop       =   4200
   ClientWidth     =   6510
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "ECG"
      Height          =   3135
      Index           =   1
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   2895
      Begin VB.CommandButton cmdECG 
         Caption         =   "Read ECG"
         Height          =   495
         Left            =   840
         TabIndex        =   13
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtF 
         Height          =   495
         Left            =   480
         TabIndex        =   11
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtD 
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtA 
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "RCL"
      Height          =   3135
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.CommandButton cmdRCL 
         Caption         =   "Read RCL"
         Height          =   495
         Left            =   960
         TabIndex        =   12
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtL 
         Height          =   495
         Left            =   600
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtR 
         Height          =   495
         Left            =   600
         TabIndex        =   7
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   255
      End
   End
End
Attribute VB_Name = "FrmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdECG_Click()
    frmMain.cmdTriggerECG_Click
End Sub

Private Sub cmdRCL_Click()
    frmMain.cmdTriggerRCL_Click
End Sub

Private Sub Form_Activate()
    DebugMode = True
End Sub

Private Sub Form_Load()
    txtR.Text = ""
    txtL.Text = ""
    txtA.Text = ""
    txtD.Text = ""
    txtF.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DebugMode = False
End Sub
