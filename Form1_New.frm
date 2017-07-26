VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   19875
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   19875
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   480
      TabIndex        =   225
      Text            =   "Text4"
      Top             =   9120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00008000&
      Caption         =   "OUTPUT"
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   8160
      TabIndex        =   218
      Top             =   9000
      Width           =   11895
      Begin VB.TextBox txtPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   1440
         TabIndex        =   221
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtFail 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   5520
         TabIndex        =   220
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   9720
         TabIndex        =   219
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   8280
         TabIndex        =   224
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAIL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   4560
         TabIndex        =   223
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   240
         TabIndex        =   222
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Timer tmrdelayPost 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5520
      Top             =   9000
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   4680
      TabIndex        =   217
      Top             =   9000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer tmrDelaySendStatus 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5520
      Top             =   9360
   End
   Begin VB.Frame Frame9 
      Caption         =   "RCL dan ECG"
      Height          =   7575
      Left            =   9000
      TabIndex        =   122
      Top             =   1560
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Frame Frame8 
         BackColor       =   &H00008000&
         Caption         =   "Instrument Status"
         ForeColor       =   &H8000000E&
         Height          =   1095
         Left            =   5280
         TabIndex        =   210
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "ECG Test 2"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   214
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "ECG Test 1"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   213
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "RCL Test 2"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   212
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "RCL Test 1"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   211
            Top             =   360
            Width           =   1095
         End
         Begin VB.Shape shpRCLTest1 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   1320
            Shape           =   3  'Circle
            Top             =   360
            Width           =   255
         End
         Begin VB.Shape shpRCLTest2 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   1320
            Shape           =   3  'Circle
            Top             =   720
            Width           =   255
         End
         Begin VB.Shape shpECGTest1 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   1320
            Shape           =   3  'Circle
            Top             =   1080
            Width           =   255
         End
         Begin VB.Shape shpECGTest2 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   1320
            Shape           =   3  'Circle
            Top             =   1440
            Width           =   255
         End
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H0000FF00&
         Height          =   285
         Left            =   4560
         TabIndex        =   209
         Text            =   "2"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   3840
         TabIndex        =   208
         Text            =   "2"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   3840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   207
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Index           =   16
         Left            =   3840
         TabIndex        =   205
         Top             =   1680
         Visible         =   0   'False
         Width           =   2895
         Begin VB.Timer tmrReadPLC 
            Interval        =   300
            Left            =   600
            Top             =   240
         End
         Begin MSWinsockLib.Winsock Winsock1 
            Left            =   120
            Top             =   240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin VB.Label lblStatusPLC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   120
            TabIndex        =   206
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00008000&
         ForeColor       =   &H8000000E&
         Height          =   1455
         Left            =   3840
         TabIndex        =   197
         Top             =   240
         Visible         =   0   'False
         Width           =   4815
         Begin VB.CommandButton cmdConnect 
            BackColor       =   &H8000000E&
            Caption         =   "Connect"
            Height          =   615
            Left            =   2640
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Form1_New.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   203
            Top             =   1200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton Command16 
            Caption         =   "TEST"
            Height          =   495
            Left            =   1560
            TabIndex        =   202
            Top             =   840
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Test Send DATA to Laser"
            Height          =   495
            Left            =   240
            TabIndex        =   201
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtLaserID 
            Height          =   285
            Left            =   2040
            TabIndex        =   200
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtClearPosition 
            Height          =   285
            Left            =   1440
            TabIndex        =   199
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtCavity 
            Height          =   285
            Left            =   240
            TabIndex        =   198
            Text            =   "1"
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Timer TMRDELAYECG 
            Enabled         =   0   'False
            Interval        =   300
            Left            =   2760
            Top             =   600
         End
         Begin VB.Timer tmrStatusECG1 
            Interval        =   500
            Left            =   3240
            Top             =   600
         End
         Begin VB.Timer tmrStatusECG2 
            Interval        =   500
            Left            =   3720
            Top             =   600
         End
         Begin VB.Timer tmrStatusRCL2 
            Interval        =   500
            Left            =   3720
            Top             =   240
         End
         Begin VB.Timer tmrStatusRCL1 
            Interval        =   500
            Left            =   3240
            Top             =   240
         End
         Begin VB.Timer TMRDELAY 
            Enabled         =   0   'False
            Interval        =   300
            Left            =   2760
            Top             =   240
         End
         Begin VB.Timer Timer4 
            Enabled         =   0   'False
            Interval        =   1500
            Left            =   4200
            Top             =   240
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Cavity"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   240
            TabIndex        =   204
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Timer Timer5 
         Interval        =   500
         Left            =   2760
         Top             =   1320
      End
      Begin VB.TextBox txtIDPost2 
         Height          =   375
         Left            =   1440
         TabIndex        =   167
         Text            =   "txtIDPost2"
         Top             =   840
         Width           =   1215
      End
      Begin VB.Timer tmrCylPost2 
         Interval        =   500
         Left            =   2760
         Top             =   960
      End
      Begin VB.TextBox txtIDPost3 
         Height          =   375
         Left            =   1440
         TabIndex        =   166
         Text            =   "txtIDPost3"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Timer tmrCylPost3 
         Interval        =   500
         Left            =   2760
         Top             =   3240
      End
      Begin VB.CommandButton cmdTriggerECG 
         Caption         =   "Read ECG"
         Height          =   495
         Left            =   1440
         TabIndex        =   132
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdTriggerRCL 
         Caption         =   "Read RCL"
         Height          =   495
         Left            =   1440
         TabIndex        =   131
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdPos2 
         Caption         =   "Pos 2"
         Height          =   495
         Left            =   120
         TabIndex        =   130
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdPos3 
         Caption         =   "Pos 3"
         Height          =   495
         Left            =   120
         TabIndex        =   129
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Height          =   495
         Left            =   120
         TabIndex        =   126
         Top             =   720
         Width           =   1215
         Begin VB.OptionButton optGenapPos2 
            Caption         =   "Option1"
            Height          =   255
            Left            =   720
            TabIndex        =   128
            Top             =   120
            Width           =   255
         End
         Begin VB.OptionButton optGanjilPos2 
            Caption         =   "Option1"
            Height          =   255
            Left            =   240
            TabIndex        =   127
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.Frame Frame4 
         Height          =   495
         Left            =   120
         TabIndex        =   123
         Top             =   2520
         Width           =   1215
         Begin VB.OptionButton optGenapPos3 
            Caption         =   "Option1"
            Height          =   195
            Left            =   720
            TabIndex        =   125
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton optGanjilPos3 
            Caption         =   "Option2"
            Height          =   195
            Left            =   240
            TabIndex        =   124
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Label Label28 
         Caption         =   "mw20 dan 21 =1"
         Height          =   255
         Left            =   120
         TabIndex        =   171
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "mw10 dan 11 =1"
         Height          =   255
         Left            =   120
         TabIndex        =   170
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "IDPost2"
         Height          =   255
         Left            =   1440
         TabIndex        =   169
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "IDPost3"
         Height          =   255
         Left            =   1440
         TabIndex        =   168
         Top             =   2400
         Width           =   855
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Rotate"
      Height          =   1095
      Left            =   14280
      TabIndex        =   136
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
      Begin VB.TextBox txtWeek 
         Height          =   285
         Left            =   120
         TabIndex        =   142
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtRun 
         Height          =   285
         Left            =   1200
         TabIndex        =   141
         Text            =   "0"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtDis 
         Height          =   285
         Left            =   1200
         TabIndex        =   140
         Text            =   "1"
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtCav4UpdateRCL 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   139
         Text            =   "1"
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtCav4UpdateECG 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   480
         TabIndex        =   138
         Text            =   "1"
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtClear 
         Height          =   285
         Left            =   840
         TabIndex        =   137
         Text            =   "2"
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Frame Frame18 
      BackColor       =   &H00008000&
      Caption         =   "Control"
      ForeColor       =   &H8000000E&
      Height          =   1455
      Left            =   840
      TabIndex        =   195
      Top             =   7320
      Width           =   3015
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H8000000E&
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   216
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H8000000E&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   215
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H8000000E&
         Height          =   615
         Left            =   2040
         Picture         =   "Form1_New.frx":1C02
         Style           =   1  'Graphical
         TabIndex        =   196
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00008000&
      Caption         =   "Machine Status"
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   840
      TabIndex        =   117
      Top             =   5280
      Width           =   3015
      Begin VB.Shape shpLowPressure 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         Shape           =   3  'Circle
         Top             =   1560
         Width           =   255
      End
      Begin VB.Shape shpLaserOn 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         Shape           =   3  'Circle
         Top             =   1155
         Width           =   255
      End
      Begin VB.Shape shpRun 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         Shape           =   3  'Circle
         Top             =   765
         Width           =   255
      End
      Begin VB.Shape shpDevConnect 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         Shape           =   3  'Circle
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblPressure 
         BackStyle       =   0  'Transparent
         Caption         =   "Low Pressure"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   121
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblLaser 
         BackStyle       =   0  'Transparent
         Caption         =   "Laser Off"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   120
         Top             =   1155
         Width           =   735
      End
      Begin VB.Label lblRun 
         BackStyle       =   0  'Transparent
         Caption         =   "Stop"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   119
         Top             =   765
         Width           =   495
      End
      Begin VB.Label lblDevConnect 
         BackStyle       =   0  'Transparent
         Caption         =   "Device Connect"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   118
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00008000&
      Caption         =   "Select Reference"
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   840
      TabIndex        =   190
      Top             =   4080
      Width           =   3015
      Begin VB.TextBox txtRef 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   191
         Text            =   "txtRef"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtECGRef 
         Height          =   375
         Left            =   240
         TabIndex        =   192
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00008000&
      ForeColor       =   &H8000000E&
      Height          =   3255
      Left            =   4440
      TabIndex        =   172
      Top             =   2040
      Width           =   3015
      Begin VB.TextBox txtSpecRmin 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   183
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtSpecRmax 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   182
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtSpecA 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   181
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtSpecD 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   180
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtSpecF 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   179
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtMeasR 
         Height          =   375
         Left            =   1920
         TabIndex        =   178
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtMeasL 
         Height          =   375
         Left            =   1920
         TabIndex        =   177
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtMeasA 
         Height          =   375
         Left            =   1920
         TabIndex        =   176
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtMeasD 
         Height          =   375
         Left            =   1920
         TabIndex        =   175
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtMeasF 
         Height          =   375
         Left            =   1920
         TabIndex        =   174
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtSpecRnom 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   173
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "R Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   189
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "R Nom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   188
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "R Max"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   187
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Area Max"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   186
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Dif_A Max"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   185
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Flutt_Max"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   184
         Top             =   2760
         Width           =   735
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Laser Decision"
      Height          =   1695
      Left            =   10440
      TabIndex        =   145
      Top             =   6360
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox txtDecGanjil 
         Height          =   285
         Left            =   840
         TabIndex        =   153
         Text            =   "Decision Ganjil"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtCavGanjil 
         Height          =   285
         Left            =   840
         TabIndex        =   152
         Text            =   "Cavity Ganjil"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCavGenap 
         Height          =   285
         Left            =   2160
         TabIndex        =   151
         Text            =   "Cavity Genap"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtDecGenap 
         Height          =   285
         Left            =   2160
         TabIndex        =   150
         Text            =   "Decision Genap"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtLaserGanjil 
         Height          =   285
         Left            =   840
         TabIndex        =   149
         Text            =   "IndexLaserGanjil"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtLaserGenap 
         Height          =   285
         Left            =   2160
         TabIndex        =   148
         Text            =   "IndexLaserGenap"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtMarkGanjil 
         Height          =   285
         Left            =   840
         TabIndex        =   147
         Text            =   "Marking Ganjil"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtMarkGenap 
         Height          =   285
         Left            =   2160
         TabIndex        =   146
         Text            =   "Marking Genap"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Decision"
         Height          =   255
         Left            =   120
         TabIndex        =   155
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "laser"
         Height          =   255
         Left            =   120
         TabIndex        =   154
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Timer"
      Height          =   1215
      Left            =   14280
      TabIndex        =   143
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   840
         Top             =   720
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   480
         Top             =   720
      End
      Begin VB.Timer tmrUpdate 
         Interval        =   10000
         Left            =   120
         Top             =   240
      End
      Begin VB.Timer tmrSend2RCLInstrument 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   480
         Top             =   240
      End
      Begin VB.Timer tmrIndikator 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   840
         Top             =   240
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Laser dan Decision"
      Height          =   1335
      Left            =   14280
      TabIndex        =   133
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
      Begin VB.CommandButton cmdUpdatetxtPcLaser 
         Caption         =   "update.txt"
         Height          =   375
         Left            =   120
         TabIndex        =   135
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdateLaser 
         Caption         =   "Update Laser"
         Height          =   495
         Left            =   120
         TabIndex        =   134
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Input instrument"
      Height          =   1575
      Left            =   14280
      TabIndex        =   114
      Top             =   6120
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox txtBoolComm 
         Height          =   285
         Left            =   120
         TabIndex        =   144
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtInputRCL 
         Height          =   285
         Left            =   120
         TabIndex        =   116
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtInputECG 
         Height          =   285
         Left            =   120
         TabIndex        =   115
         Top             =   240
         Width           =   3255
      End
      Begin MSCommLib.MSComm RCLInstrument 
         Left            =   2400
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   5
         DTREnable       =   -1  'True
         RThreshold      =   1
         RTSEnable       =   -1  'True
         BaudRate        =   19200
      End
      Begin MSCommLib.MSComm ECGInstrument 
         Left            =   1800
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   4
         DTREnable       =   -1  'True
         RThreshold      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1080
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   120
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Laser Template"
      Height          =   2895
      Left            =   15960
      TabIndex        =   108
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
      Begin VB.TextBox txtLaserTemplate 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   113
         Text            =   "Text1"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtLaserTemplate 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   112
         Text            =   "Text1"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtLaserTemplate 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   111
         Text            =   "Text1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtLaserTemplate 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   110
         Text            =   "Text1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtLaserTemplate 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   109
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   16560
      TabIndex        =   107
      Top             =   5640
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   91160578
      CurrentDate     =   41127
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Caption         =   "Laser Marking"
      ForeColor       =   &H8000000E&
      Height          =   3135
      Left            =   4440
      TabIndex        =   96
      Top             =   5640
      Width           =   3015
      Begin VB.TextBox txtLaser2 
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   106
         Text            =   "Text1"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtLaser2 
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   105
         Text            =   "Text1"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtLaser2 
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   104
         Text            =   "Text1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtLaser2 
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   103
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtLaser2 
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   102
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtLaser1 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   101
         Text            =   "Text1"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtLaser1 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   100
         Text            =   "Text1"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtLaser1 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   99
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtLaser1 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   98
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtLaser1 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   97
         Text            =   "Text1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblCavity 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   193
         Top             =   2760
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 16"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   15
      Left            =   18600
      TabIndex        =   90
      Top             =   5520
      Width           =   1455
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   15
         Left            =   240
         TabIndex        =   95
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   15
         Left            =   240
         TabIndex        =   94
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   15
         Left            =   240
         TabIndex        =   93
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   15
         Left            =   240
         TabIndex        =   92
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   15
         Left            =   240
         TabIndex        =   91
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   15
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   15
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 15"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   14
      Left            =   17160
      TabIndex        =   84
      Top             =   5520
      Width           =   1455
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   14
         Left            =   240
         TabIndex        =   89
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   14
         Left            =   240
         TabIndex        =   88
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   14
         Left            =   240
         TabIndex        =   87
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   14
         Left            =   240
         TabIndex        =   86
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   14
         Left            =   240
         TabIndex        =   85
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   14
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   14
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 14"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   13
      Left            =   15600
      TabIndex        =   78
      Top             =   5520
      Width           =   1455
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   83
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   82
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   81
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   80
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   79
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   13
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   13
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 13"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   12
      Left            =   14160
      TabIndex        =   72
      Top             =   5520
      Width           =   1455
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   77
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   76
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   75
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   74
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   73
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   12
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   12
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 12"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   11
      Left            =   12600
      TabIndex        =   66
      Top             =   5520
      Width           =   1455
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   11
         Left            =   240
         TabIndex        =   71
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   11
         Left            =   240
         TabIndex        =   70
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   11
         Left            =   240
         TabIndex        =   69
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   11
         Left            =   240
         TabIndex        =   68
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   11
         Left            =   240
         TabIndex        =   67
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   11
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   11
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 11"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   10
      Left            =   11160
      TabIndex        =   60
      Top             =   5520
      Width           =   1455
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   65
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   64
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   63
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   62
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   61
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 10"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   9
      Left            =   9600
      TabIndex        =   54
      Top             =   5520
      Width           =   1455
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   59
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   58
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   57
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   56
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   55
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 9"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   8
      Left            =   8160
      TabIndex        =   48
      Top             =   5520
      Width           =   1455
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   53
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   52
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   51
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   50
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   49
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 8"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   7
      Left            =   18600
      TabIndex        =   42
      Top             =   2040
      Width           =   1455
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   47
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   46
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   45
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   44
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   43
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 7"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   6
      Left            =   17160
      TabIndex        =   36
      Top             =   2040
      Width           =   1455
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   41
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   40
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   39
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   38
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   37
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 6"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   5
      Left            =   15600
      TabIndex        =   30
      Top             =   2040
      Width           =   1455
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   35
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   34
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   33
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   32
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   31
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 5"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   4
      Left            =   14160
      TabIndex        =   24
      Top             =   2040
      Width           =   1455
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   29
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   28
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   27
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   26
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   25
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 4"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   3
      Left            =   12600
      TabIndex        =   18
      Top             =   2040
      Width           =   1455
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   23
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 3"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   2
      Left            =   11160
      TabIndex        =   12
      Top             =   2040
      Width           =   1455
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 2"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   1
      Left            =   9600
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Cavity 1"
      ForeColor       =   &H8000000E&
      Height          =   3255
      Index           =   0
      Left            =   8160
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
      Begin VB.TextBox txtR 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Text            =   "txtR"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtL 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Text            =   "txtL"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Text            =   "txtA"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtD 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Text            =   "txtD"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtF 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Text            =   "txtF"
         Top             =   2280
         Width           =   855
      End
      Begin VB.Shape shpECG 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   720
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape shpRCL 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Image Image1 
      Height          =   2025
      Left            =   840
      Picture         =   "Form1_New.frx":3898
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Image RightBar 
      Height          =   9990
      Left            =   21480
      Picture         =   "Form1_New.frx":804E2
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   165
   End
   Begin VB.Image LeftBar 
      Height          =   9990
      Left            =   0
      Picture         =   "Form1_New.frx":80628
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   165
   End
   Begin VB.Image TopRight 
      Height          =   1380
      Left            =   21240
      Picture         =   "Form1_New.frx":80763
      Top             =   0
      Width           =   390
   End
   Begin VB.Image TopLeft 
      Height          =   1380
      Left            =   0
      Picture         =   "Form1_New.frx":82465
      Top             =   0
      Width           =   390
   End
   Begin VB.Image BottomBar 
      Height          =   1380
      Left            =   360
      Picture         =   "Form1_New.frx":84167
      Stretch         =   -1  'True
      Top             =   11280
      Width           =   21030
   End
   Begin VB.Image BottomRight 
      Height          =   1380
      Left            =   21360
      Picture         =   "Form1_New.frx":8437C
      Top             =   11280
      Width           =   390
   End
   Begin VB.Image BottomLeft 
      Height          =   1380
      Left            =   0
      Picture         =   "Form1_New.frx":847B5
      Top             =   11280
      Width           =   390
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "COIL TESTER TESYS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7560
      TabIndex        =   194
      Top             =   360
      Width           =   5250
   End
   Begin VB.Image TopBar 
      Height          =   1380
      Left            =   360
      Picture         =   "Form1_New.frx":84C11
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20910
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7800
      TabIndex        =   165
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7800
      TabIndex        =   164
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7800
      TabIndex        =   163
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7800
      TabIndex        =   162
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7800
      TabIndex        =   161
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7800
      TabIndex        =   160
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7800
      TabIndex        =   159
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7800
      TabIndex        =   158
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7800
      TabIndex        =   157
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7800
      TabIndex        =   156
      Top             =   2400
      Width           =   375
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mMode 
      Caption         =   "Mode"
      Begin VB.Menu mPLC_TCPIP 
         Caption         =   "PLC TCP/IP"
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug"
      End
   End
   Begin VB.Menu mLogFile 
      Caption         =   "Log File"
      Begin VB.Menu mOpenLogFile 
         Caption         =   "Open Log File"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ind As Boolean
Dim laser, decision, countCom As Boolean


'Private Sub cmdDataLog_Click()
'    On Error GoTo ErrHandler
'    Dim sFile As String
'    Dim xlsApp As Object
'    Dim xlsWB1 As Object
'
'    With CommonDialog1
'        .DialogTitle = "Open data Log"
'        .CancelError = False
'        .Filter = "LogFile (*.csv)|*.csv"
'        .ShowOpen
'        If Len(.FileName) = 0 Then
'            Exit Sub
'        End If
'        sFile = .FileName
'    End With
'    Set xlsApp = CreateObject("Excel.Application")
'    xlsApp.Visible = True
'    Set xlsWB1 = xlsApp.Workbooks.Open(sFile)
'    Exit Sub
'
'ErrHandler:
'            MsgBox "There is a problem while opening the xls document. " & _
'            " Please ensure it is present!", vbCritical, "Error"
'End Sub





Private Sub cmdStart_Click(Index As Integer)
    Select Case Index
    Case 0
        holdread = True
        Call cmdTriggerECG_Click
    Case 1
        holdread = False
    End Select
End Sub

Public Sub cmdTriggerECG_Click()
     firstload = False
    If ECGInstrument.PortOpen = True Then   'ECGInstrument = MSComm1 lama
        txtInputECG.Text = ""
        Text3.Text = "1"
        ECGInstrument.Output = "S" & vbCrLf
        ECGtriger = True
    Else
        MsgBox "Open Device First!!!", vbOKOnly + vbExclamation
    End If
End Sub

Public Sub cmdTriggerRCL_Click()
    If RCLInstrument.PortOpen = True Then   'RCLInstrument = MSComm2 lama
        Text2.Text = "2"
        RCLInstrument.Output = Chr$(27) + "8" 'go to trigger
        RCLInstrument.Output = Chr$(27) + "7" 'asking status register
        txtInputRCL.Text = ""
        tmrSend2RCLInstrument.Enabled = True
    Else
        MsgBox "Open Device First!!!", vbOKOnly + vbExclamation
    End If

End Sub



 Sub SaveDatalog(Station As Integer, CavityName As Integer, RCL As String, ECG As String, strResult As String)
        If CavityName = 0 Then Exit Sub
        Dim FolderName As String
        Dim FileName As String
        Dim strID As String
    
        FolderName = App.Path & "\DataLog\"
        FileName = Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".csv"
        strID = Day(Date) & Month(Date) & Year(Date)
        On Error Resume Next
        If Dir(FolderName) = "" Then
            MkDir (FolderName)
        End If
        If Dir(FolderName & FileName) = "" Then
            Open FolderName & FileName For Append As #1
            Print #1, "Station,CavityNumber,Reff,R,L,A,D,F,RCL,ECG,Final Result,Datetime"
            '15-3-2013
        Else
            Open FolderName & FileName For Append As #1
        End If

            Print #1, Station _
            & "," & CavityName _
            & "," & txtRef.Text _
            & "," & DatatxtR _
            & "," & DatatxtL _
            & "," & txtMeasA.Text _
            & "," & txtMeasD.Text _
            & "," & txtMeasF.Text _
            & "," & RCL _
            & "," & ECG _
            & "," & strResult _
            & "," & Now
        Close #1
        Call UpdateOutputQTY(strID, strResult)

    
End Sub




Public Sub cmdUpdatetxtPcLaser_Click()
    'Exit Sub 'sementara lewati
    On Error Resume Next
    If Left(UCase(txtRef.Text), 1) = "U" Then
        Open "\\192.168.0.10\DATA\update.txt" For Output As #1
        Print #1, "1"
        Print #1, "U7_Art"
        Close #1

        's = Shell("\\192.168.0.10\Data\tmp.txt", vbNormalFocus)

    Else
        Open "\\192.168.0.10\DATA\update.txt" For Output As #1
        Print #1, "1"
        Print #1, "Common_Art"
        Close #1
        's = Shell("\\192.168.0.10\Data\tmp.txt", vbNormalFocus)
    End If
End Sub






Private Sub Command15_Click()
    Update_Datalog_Laser
End Sub

Private Sub Command16_Click()
    cmdUpdatetxtPcLaser_Click
End Sub



Private Sub ECGInstrument_OnComm()
    Dim AA, DD, FF As Integer
    Dim AT, DT, FT As Integer
    Dim AE, DE, FE As Integer
    Dim Result As String
    
    'ECGInstrument.Output = "J"
    txtInputECG.Text = txtInputECG & ECGInstrument.Input
    txtMeasA.Text = Trim(Mid(txtInputECG.Text, 3, 6))               'Trim(Mid(txtInputECG.Text, 7, 6))
    txtMeasD.Text = Trim(Mid(txtInputECG.Text, 13, 5))              'Trim(Mid(txtInputECG.Text, 19, 5))
    txtMeasF.Text = Trim(Mid(txtInputECG.Text, 22, 5))              'Trim(Mid(txtInputECG.Text, 30, 5))
    'Text4.Text = txtInputECG.Text
    
    FrmDebug.txtA.Text = Trim(Mid(txtInputECG.Text, 3, 6))          'Trim(Mid(txtInputECG.Text, 7, 6))
    FrmDebug.txtD.Text = Trim(Mid(txtInputECG.Text, 13, 5))         'Trim(Mid(txtInputECG.Text, 19, 6))
    FrmDebug.txtF.Text = Trim(Mid(txtInputECG.Text, 220, 5))         'Trim(Mid(txtInputECG.Text, 30, 6))
    
    If DebugMode = False Then
        If firstload = False And ECGtriger = True Then
            ECGtriger = False 'add 11/03 13
        End If
    End If
  
End Sub


Sub OpenPortRCL_ECG()
    If RCLInstrument.PortOpen = True Then
        RCLInstrument.PortOpen = False
    End If
    RCLInstrument.PortOpen = True
    Text2.Text = "2"
    RCLInstrument.Output = Chr(27) & "2" 'go to remote
    
    'ecg
    If ECGInstrument.PortOpen = True Then
        ECGInstrument.PortOpen = False
    End If
    ECGInstrument.PortOpen = True
    Text3.Text = "2"
    ECGInstrument.Output = txtECGRef.Text & vbCrLf
    Timer4.Enabled = True
End Sub

Private Sub cmdPos2_Click()
Dim i, j As Integer

i = Val(txtIDPost2.Text)
If optGenapPos2.Value = True Then
    txtR(i * 2 - 2).Text = txtMeasR.Text
    txtL(i * 2 - 2).Text = txtMeasL.Text
ElseIf optGanjilPos2.Value = True Then
    txtR(i * 2 - 1).Text = txtMeasR.Text
    txtL(i * 2 - 1).Text = txtMeasL.Text
End If


'''update shape RCL cavity
For j = 0 To 15
    If Val(txtR(j).Text) < Val(txtSpecRmin.Text) Or Val(txtR(j).Text) > Val(txtSpecRmax.Text) Then
        shpRCL(j).BackColor = vbRed
    ElseIf txtR(j).Text = "" Then
        shpRCL(j).BackColor = vbRed
    Else
        shpRCL(j).BackColor = vbGreen
    End If
Next j



'Check RCL Pass/Fail
If Val(txtMeasR.Text) < Val(txtSpecRmin.Text) Or Val(txtMeasR.Text) > Val(txtSpecRmax.Text) Then
        WriteModbus 1070, 0 'NG
Else
        WriteModbus 1070, 1 'OK
End If
'RCL Finish

'Do Until Modbuswait = False
'    DoEvents
'Loop

WriteModbus 1071, 1
TMRDELAY.Enabled = True


End Sub

Private Sub cmdPos3_Click()

    Dim Q As Integer
    Dim i As Integer
    Dim j As Integer
    Dim A, D, F As String
    Dim CavityName As Integer
    Dim Result As String
    Dim strRCL As String
    Dim strECG As String
    Dim strResult As String
    
    
    Q = Val(txtIDPost3.Text)
    
    '1
    If optGenapPos3.Value = True Then
        txtA(Q * 2 - 2).Text = txtMeasA.Text
        txtD(Q * 2 - 2).Text = txtMeasD.Text
        txtF(Q * 2 - 2).Text = txtMeasF.Text
        DatatxtR = txtR(Q * 2 - 2).Text
        DatatxtL = txtL(Q * 2 - 2).Text
        CavityName = 1
    '2
    ElseIf optGanjilPos3.Value = True Then
        txtA(Q * 2 - 1).Text = txtMeasA.Text
        txtD(Q * 2 - 1).Text = txtMeasD.Text
        txtF(Q * 2 - 1).Text = txtMeasF.Text
        DatatxtR = txtR(Q * 2 - 1).Text
        DatatxtL = txtL(Q * 2 - 1).Text
        CavityName = 2
    End If
    
    
    'update shape ECG cavity
    For j = 0 To 15
        If Val(txtA(j).Text) > Val(txtSpecA.Text) Or Val(txtD(j).Text) > Val(txtSpecD.Text) Or Val(txtF(j).Text) > Val(txtSpecF.Text) Then
           shpECG(j).BackColor = vbRed
        ElseIf txtA(j).Text = "" Or txtD(j).Text = "" Or txtF(j).Text = "" Then
            shpECG(j).BackColor = vbRed
        Else
            shpECG(j).BackColor = vbGreen
        End If
    Next j

    If Val(txtMeasA.Text) > Val(txtSpecA.Text) Or Val(txtMeasD.Text) > Val(txtSpecD.Text) Or Val(txtMeasF.Text) > Val(txtSpecF.Text) Then
       WriteModbus 1080, 0 'FAIL
    Else
        WriteModbus 1080, 1 'PASS
    End If
    
    
        Select Case Q
        Case 1
            i = 0
        Case 2
            i = 2
        Case 3
            i = 4
        Case 4
            i = 6
        Case 5
            i = 8
        Case 6
            i = 10
        Case 7
            i = 12
        Case 8
            i = 14
    End Select
    
    
    'Read RCL and ECG Status
    If CavityName = 1 Then
        If shpRCL(i).BackColor = vbGreen Then
           strRCL = "OK"
        Else
           strRCL = "NG"
        End If
        
        If shpECG(i).BackColor = vbGreen Then
           strECG = "OK"
        Else
           strECG = "NG"
        End If
        
        'final Result 1
        If strRCL = "OK" And strECG = "OK" Then
            strResult = "PASS"
        Else
            strResult = "FAIL"
        End If
        '
    
    ElseIf CavityName = 2 Then
        If shpRCL(i + 1).BackColor = vbGreen Then
           strRCL = "OK"
        Else
           strRCL = "NG"
        End If
        
        If shpECG(i + 1).BackColor = vbGreen Then
           strECG = "OK"
        Else
           strECG = "NG"
        End If
                
        'final Result 2
        If strRCL = "OK" And strECG = "OK" Then
            strResult = "PASS"
        Else
            strResult = "FAIL"
        End If

    End If
    
    
      

    'ECG Finish
'    Do Until Modbuswait = False
'        DoEvents
'    Loop
    
    Call SaveDatalog(Q, CavityName, strRCL, strECG, strResult)
    'List1.AddItem ("TEST")

    
    WriteModbus 1081, 1
    TMRDELAYECG.Enabled = True
End Sub

Public Sub cmdReset_Click()
    FirstInisial
End Sub

Private Sub Form_Load()
    FirstInisial
End Sub


Sub FirstInisial()
Dim i As Integer
countCom = False
firstload = True
Dim ClearAll As Integer

For ClearAll = 0 To 15
    txtR(ClearAll).Text = ""
    txtL(ClearAll).Text = ""
    txtA(ClearAll).Text = ""
    txtD(ClearAll).Text = ""
    txtF(ClearAll).Text = ""
Next

txtSpecRmin.Text = ""
txtSpecRnom.Text = ""
txtSpecRmax.Text = ""
txtSpecA.Text = ""
txtSpecD.Text = ""
txtSpecF.Text = ""
txtRef.Text = ""
lblCavity.Caption = "-"

'StatusBar1.Panels.Item(1).Text = "Sistem Standby"
'StatusBar1.Panels.Item(2).Text = "Database Not Connected"


    ind = False
    
    cycle = 0
    cav = Val(0)
    cav1 = Val(0)
    
    For i = 0 To 4
        txtLaser1(i) = ""
        txtLaser2(i) = ""
        txtLaserTemplate(i) = ""
    Next i
    
    SettingPLC
    OpenPortRCL_ECG
   
    
    txtRef.Text = ""
    holdread = True
    DebugMode = False
    ECGtriger = False
    stsDelay = True
    Call ClearshpRCL_ECG

End Sub

Sub ClearshpRCL_ECG()
    Dim Clrshape As Integer
    For Clrshape = 0 To 15
        shpRCL(Clrshape).BackColor = &H8000&
        shpECG(Clrshape).BackColor = &H8000&
    Next
End Sub


Sub waktu()
MonthView1.Value = Now()

txtWeek.Text = "8B" & Right((MonthView1.Year), 2) & Format(MonthView1.Week, "00") & (Val(MonthView1.DayOfWeek) - 1) & "3"

'8B120922 - 8B(fix) + 12(var, 2012) + 09(var, week 9) + 2 (var, hari ke 2) + 2(fix,tester 2)
End Sub

Sub connect_database()
Dim DBHis
Dim i As Integer
Dim rs As ADODB.Recordset
Dim A As String

If (txtRef.Text <> "") Then
    Set DBHis = New ADODB.Connection
                              
    DBHis.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\db_tes.mdb;Persist Security Info=False"
    DBHis.Open
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * from coilref where PrintCoil like '" + txtRef.Text + "%'", DBHis, adOpenKeyset, adLockOptimistic
    With rs
    'rs.MoveFirst
    If .EOF = True Then
        Beep
        .Close
        MsgBox "Database not exist", vbOKOnly + vbExclamation, "Warning"
        txtRef.SetFocus
        Exit Sub
    Else
        While (.EOF = False)
            
                'resistance
                txtSpecRmin.Text = .Fields("ResistansiMini")
                txtSpecRnom.Text = .Fields("ResistansiNominal")
                txtSpecRmax.Text = .Fields("Resistansi Maxi")
                'ECG
                txtSpecA.Text = .Fields("Area")
                txtSpecD.Text = .Fields("Dif_A")
                txtSpecF.Text = .Fields("Flutt")
                'laser_marking
                txtLaserTemplate(0).Text = .Fields("PrintCoil")
                txtLaserTemplate(1).Text = .Fields("PrintTeg")
                txtLaserTemplate(2).Text = .Fields("PrintFrekuensi")
                If IsNull(.Fields("PrintFrekuensi2")) Then
                    txtLaserTemplate(3).Text = ""
                Else
                txtLaserTemplate(3).Text = .Fields("PrintFrekuensi2")
                End If
                
                .Close
                Exit Sub
            
        Wend
        
    End If
    End With
ElseIf (txtRef.Text = "") Then
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub



Private Sub mnuDebug_Click()
    FrmDebug.Show 1
End Sub

Private Sub mOpenLogFile_Click()    'new dari menu olf lama
    On Error GoTo ErrHandler
    Dim sFile As String
    Dim xlsApp As Object
    Dim xlsWB1 As Object
    
    With CommonDialog1
        .DialogTitle = "Open data Log"
        .CancelError = False
        .Filter = "LogFile (*.csv)|*.csv"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.Visible = True
    Set xlsWB1 = xlsApp.Workbooks.Open(sFile)
    Exit Sub
    
ErrHandler:
            MsgBox "There is a problem while opening the xls document. " & _
            " Please ensure it is present!", vbCritical, "Error"
End Sub

Private Sub mPLC_TCPIP_Click()
    Form1.Show
End Sub

Private Sub mExit_Click()
    End
End Sub

Private Sub RCLInstrument_OnComm()
Dim Result As String


If Val(Text2.Text) = 1 Then
    txtInputRCL.Text = txtInputRCL.Text + RCLInstrument.Input
    Result = txtInputRCL.Text
    Debug.Print Result
    p = InStr(txtInputRCL.Text, "R")
    Q = InStr(txtInputRCL.Text, "L")
    
    If Val(p) > 0 Then
        txtMeasR.Text = Mid(txtInputRCL.Text, Val(p + 1), 6)
        FrmDebug.txtR.Text = Mid(txtInputRCL.Text, Val(p + 1), 6)
    End If
    
    If Val(Q) > 0 Then
        txtMeasL.Text = Mid(txtInputRCL.Text, Val(Q + 1), 6)
        FrmDebug.txtL.Text = Mid(txtInputRCL.Text, Val(Q + 1), 6)
    End If
    
    If DebugMode = False Then cmdPos2_Click
    
End If
End Sub




Private Sub Timer4_Timer()
    RCLInstrument.Output = Chr(27) + "2"
    ECGInstrument.Output = "QJ" & vbCrLf
    Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
    txtIDPost2.Text = Form1.Text4(2)
    txtIDPost3.Text = Form1.Text4(3)
End Sub


Private Sub tmrCylPost2_Timer()
    
    If Val(Form1.Text4(10).Text) = 1 Then
        optGenapPos2.Value = True
    ElseIf Val(Form1.Text4(11).Text) = 1 Then
        optGanjilPos2.Value = True
    Else
        optGanjilPos2.Value = False
        optGenapPos2.Value = False
    End If

End Sub

Private Sub tmrCylPost3_Timer()
    If Val(Form1.Text4(20).Text) = 1 Then
    optGenapPos3.Value = True
    ElseIf Val(Form1.Text4(21).Text) = 1 Then
    optGanjilPos3.Value = True
    Else
    optGanjilPos3.Value = False
    optGenapPos3.Value = False
    End If
End Sub

Private Sub TMRDELAY_Timer()
tunggu = tunggu + 1
    If tunggu = 1 Then
        WriteModbus 1071, 0 'Netral
        TMRDELAY.Enabled = False
        tunggu = 0
        goNextStep = True
    End If
End Sub

Private Sub TMRDELAYECG_Timer()
tunggu2 = tunggu2 + 1
    If tunggu2 = 1 Then
        WriteModbus 1081, 0
        TMRDELAYECG.Enabled = False
        tunggu2 = 0
    End If
End Sub

Private Sub tmrdelayPost_Timer()
    tmrdelayPost.Enabled = False
    Call cmdPos3_Click
End Sub

Private Sub tmrDelaySendStatus_Timer()
    Delay = False
    tmrDelaySendStatus.Enabled = False
End Sub

Private Sub tmrIndikator_Timer()
    If ind = False Then
        shpDevConnect.BackColor = vbRed
        
        If Val(txtRun.Text) = 1 Then
            shpRun.BackColor = vbYellow
            lblRun.Caption = "Run"
        ElseIf Val(txtRun.Text) = 0 Then
            shpRun.BackColor = vbRed
            lblRun.Caption = "Stop"
        End If
    
        If Val(Form1.Text4(101).Text) = 1 Then
            shpLaserOn.BackColor = vbRed
            lblLaser.Caption = "Laser On"
        ElseIf Val(Form1.Text4(101).Text) = 0 Then
            shpLaserOn.BackColor = vbRed
            lblLaser.Caption = "Laser Off"
        End If
        
        If Val(Form1.Text4(102).Text) = 1 Then
            shpLowPressure.BackColor = vbGreen
            lblPressure.Caption = "Low Pressure"
        ElseIf Val(Form1.Text4(102).Text) = 0 Then
            shpLowPressure.BackColor = vbGreen
            lblPressure.Caption = "Pressure OK"
        End If
        
        If Val(Form1.Text4(54).Text) = 1 Then
            shpRCLTest1.BackColor = vbRed
        ElseIf Val(Form1.Text4(54).Text) = 0 Then   'sensor cylinder 1/kanan
            shpRCLTest1.BackColor = vbRed
        End If
        
        If Val(Form1.Text4(55).Text) = 1 Then
            shpRCLTest2.BackColor = vbGreen
        ElseIf Val(Form1.Text4(55).Text) = 0 Then   'sensor cylinder 1/kiri
            shpRCLTest2.BackColor = vbRed
        End If
        
        If Val(Form1.Text4(56).Text) = 1 Then
            shpECGTest1.BackColor = vbRed
        ElseIf Val(Form1.Text4(56).Text) = 0 Then   'sensor cylinder 2/kanan
            shpECGTest1.BackColor = vbRed
        End If
        
        If Val(Form1.Text4(57).Text) = 1 Then
            shpECGTest2.BackColor = vbGreen
        ElseIf Val(Form1.Text4(57).Text) = 0 Then   'sensor cylinder 2/kiri
            shpECGTest2.BackColor = vbRed
        End If
    Else
        shpDevConnect.BackColor = vbGreen
    
        If Val(txtRun.Text) = 1 Then
            shpRun.BackColor = vbGreen
            lblRun.Caption = "Run"
        ElseIf Val(txtRun.Text) = 0 Then
            shpRun.BackColor = vbRed
            lblRun.Caption = "Stop"
        End If
    
        If Val(Form1.Text4(101).Text) = 1 Then
            shpLaserOn.BackColor = vbGreen
            lblPressure.Caption = "Laser On"
        ElseIf Val(Form1.Text4(101).Text) = 0 Then
            shpLowPressure.BackColor = vbRed
            lblPressure.Caption = "Laser Off"
        End If
        
        If Val(Form1.Text4(102).Text) = 1 Then
            shpLaserOn.BackColor = vbRed
            lblPressure.Caption = "Low Pressure"
        ElseIf Val(Form1.Text4(102).Text) = 0 Then
            shpLowPressure.BackColor = vbGreen
            lblPressure.Caption = "Pressure OK"
        End If
        
        If Val(Form1.Text4(54).Text) = 1 Then
            shpRCLTest1.BackColor = vbGreen
        ElseIf Val(Form1.Text4(54).Text) = 0 Then
            shpRCLTest1.BackColor = vbRed
        End If
        
        If Val(Form1.Text4(55).Text) = 1 Then
            shpRCLTest2.BackColor = vbRed
        ElseIf Val(Form1.Text4(55).Text) = 0 Then
            shpRCLTest2.BackColor = vbRed
        End If
        
        If Val(Form1.Text4(56).Text) = 1 Then
            shpECGTest1.BackColor = vbGreen
        ElseIf Val(Form1.Text4(56).Text) = 0 Then
            shpECGTest1.BackColor = vbRed
        End If
        
        If Val(Form1.Text4(57).Text) = 1 Then
            shpECGTest2.BackColor = vbRed
        ElseIf Val(Form1.Text4(57).Text) = 0 Then
            shpECGTest2.BackColor = vbRed
        End If
    End If
    
    ind = Not ind
End Sub

Private Sub tmrSend2RCLInstrument_Timer()   'pengganti timer 2 lama
    Text2.Text = "7"
    'RCLInstrument.Output = Chr(27) + "7"
    Text2.Text = "1"
    RCLInstrument.Output = "RESI?;INDU?" + Chr(10)
    'RCLInstrument.Output = "RESI? ;INDU? " + Chr(10)
    'RCLInstrument.Output = Chr(99) & Chr(111) & Chr(109) & Chr(63) & Chr(10) ' "RESI?" + Chr(10)
    'RCLInstrument.Output = Chr(27) + "7"
    'RCLInstrument.Output = Chr(27) + "8"
    tmrSend2RCLInstrument.Enabled = False
End Sub



Private Sub tmrStatusECG1_Timer()
   If Val(Form1.Text4(20).Text) = 1 Then
        Call frmMain.cmdTriggerECG_Click
        tmrStatusECG1.Enabled = False
        tmrStatusECG2.Enabled = True
   End If
End Sub

Private Sub tmrStatusECG2_Timer()
   If Val(Form1.Text4(21).Text) = 1 Then
        Call frmMain.cmdTriggerECG_Click
        tmrStatusECG1.Enabled = True
        tmrStatusECG2.Enabled = False
   End If
End Sub

Private Sub tmrStatusRCL1_Timer()
   'RCL
    If Val(Form1.Text4(10).Text) = 1 Then
        Call frmMain.cmdTriggerRCL_Click
        tmrStatusRCL1.Enabled = False
        tmrStatusRCL2.Enabled = True
    End If
End Sub
Private Sub tmrStatusRCL2_Timer()
     If Val(Form1.Text4(11).Text) = 1 Then
        Call frmMain.cmdTriggerRCL_Click
        tmrStatusRCL2.Enabled = False
        tmrStatusRCL1.Enabled = True
     End If
End Sub


Private Sub tmrUpdate_Timer()
    Call waktu
End Sub

Private Sub txtBoolComm_Change()
    If txtBoolComm.Text = "true" Then
        txtInputECG.Text = ""
        Text2.Text = "1"
        ECGInstrument.Output = "S" & vbCrLf
    End If
End Sub




Private Sub txtClearPosition_Change()
    ClearTxtNext
End Sub



Public Sub ClearTxtNext()
        Dim iClear, jClear As Integer
        Dim iClear2, jClear2 As Integer
    If Val(txtClearPosition.Text) <> 0 Then
        If Val(txtClearPosition.Text) <= 6 Then
            iClear = Val(txtClearPosition.Text) * 2 - 2 + 2
            jClear = Val(txtClearPosition.Text) * 2 - 1 + 2
            iClear2 = Val(txtClearPosition.Text) * 2 - 2 + 4
            jClear2 = Val(txtClearPosition.Text) * 2 - 1 + 4
        ElseIf Val(txtClearPosition.Text) = 7 Then
            iClear = Val(txtClearPosition.Text) * 2 - 2 + 2
            jClear = Val(txtClearPosition.Text) * 2 - 1 + 2
            iClear2 = (Val(txtClearPosition.Text) - 6) * 2 - 2
            jClear2 = (Val(txtClearPosition.Text) - 6) * 2 - 1
        ElseIf Val(txtClearPosition.Text) = 8 Then
            iClear = Val(Val(txtClearPosition.Text) - 7) * 2 - 2
            jClear = Val(Val(txtClearPosition.Text) - 7) * 2 - 1
            iClear2 = (Val(txtClearPosition.Text) - 6) * 2 - 2
            jClear2 = (Val(txtClearPosition.Text) - 6) * 2 - 1
        End If
            
            txtR(iClear).Text = ""
            txtL(iClear).Text = ""
            txtA(iClear).Text = ""
            txtD(iClear).Text = ""
            txtF(iClear).Text = ""
            txtR(jClear).Text = ""
            txtL(jClear).Text = ""
            txtA(jClear).Text = ""
            txtD(jClear).Text = ""
            txtF(jClear).Text = ""

            txtR(iClear2).Text = ""
            txtL(iClear2).Text = ""
            txtA(iClear2).Text = ""
            txtD(iClear2).Text = ""
            txtF(iClear2).Text = ""
            txtR(jClear2).Text = ""
            txtL(jClear2).Text = ""
            txtA(jClear2).Text = ""
            txtD(jClear2).Text = ""
            txtF(jClear2).Text = ""
      End If
End Sub


Private Sub TxtLaserID_Change()
    Dim i, j As Integer
    Dim CurIDLaser As Integer
    CurIDLaser = Val(TxtLaserID.Text)
    Dim StatusCavity1, StatusCavity2 As String
    Dim idx As Integer
    
    '  1       2       3       4       5       6       7       8
    '0   1   2   3   4   5   6   7   8   9   10  11  12  13  14  15
    Select Case CurIDLaser
        Case 1
            i = 0
        Case 2
            i = 2
        Case 3
            i = 4
        Case 4
            i = 6
        Case 5
            i = 8
        Case 6
            i = 10
        Case 7
            i = 12
        Case 8
            i = 14
    End Select
    
    'Cavity1
    If shpRCL(i).BackColor = vbGreen And shpECG(i).BackColor = vbGreen Then
        'Pass
        Call connect_database
        txtLaserTemplate(4).Text = txtWeek.Text
        txtLaser1(4).Text = txtWeek.Text
        For idx = 0 To 3
            txtLaser1(idx).Text = txtLaserTemplate(idx).Text
        Next
        
    ElseIf shpRCL(i).BackColor = vbRed And shpECG(i).BackColor = vbGreen Then
          'FailRCL
          txtLaser1(0).Text = " FAIL 22"
          For idx = 1 To 4
            txtLaser1(idx).Text = ""
          Next
          
    ElseIf shpRCL(i).BackColor = vbGreen And shpECG(i).BackColor = vbRed Then
        'FailECG
          txtLaser1(0).Text = " FAIL 32"
          For idx = 1 To 4
            txtLaser1(idx).Text = ""
          Next
    Else
        'FailECG
          txtLaser1(0).Text = " FAIL 32"
          For idx = 1 To 4
            txtLaser1(idx).Text = ""
          Next
    End If
    

    Dim intry As Integer
    intry = Val(i) + 1
    'Cavity2
    If shpRCL(i + 1).BackColor = vbGreen And shpECG(i + 1).BackColor = vbGreen Then
        'LOAD db TO REFRESH
        Call connect_database
        txtLaserTemplate(4).Text = txtWeek.Text
        txtLaser2(4).Text = txtWeek.Text
        For idx = 0 To 3
            txtLaser2(idx).Text = txtLaserTemplate(idx).Text
        Next
    ElseIf shpRCL(i + 1).BackColor = vbRed And shpECG(i + 1).BackColor = vbGreen Then
          'FailRCL
          txtLaser2(0).Text = " FAIL 22"
          For idx = 1 To 4
            txtLaser2(idx).Text = ""
          Next
          
    ElseIf shpRCL(i + 1).BackColor = vbGreen And shpECG(i + 1).BackColor = vbRed Then
        'FailECG
          txtLaser2(0).Text = " FAIL 32"
          For idx = 1 To 4
            txtLaser2(idx).Text = ""
          Next
    Else
        'FailECG
          txtLaser2(0).Text = " FAIL 32"
          For idx = 1 To 4
            txtLaser2(idx).Text = ""
          Next
    End If
    
    'Disable Sementara
    Call Update_Datalog_Laser 'Send Data To PC Laser
    
'    tmrDelaySendStatus.Enabled = True
'    Do Until stsDelay = False
'        DoEvents
'    Loop
'    stsDelay = True
    
    Call cmdUpdatetxtPcLaser_Click
    WriteModbus 1030, 1 'Signal to PLC DataLaser Update
    lblCavity.Caption = "CAVITY " & TxtLaserID.Text

End Sub
Sub UpdateOutputQTY(strIDdate As String, strResult As String)
    Dim DBHis
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim LastQtyPass, LastQtyFail As String
        Set DBHis = New ADODB.Connection
        DBHis.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\db_tes.mdb;Persist Security Info=False"
        DBHis.Open
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * from tblOutput where IDdate like '" + strIDdate + "%'", DBHis, adOpenKeyset, adLockOptimistic
        With rs
            If .EOF = True Then
                .AddNew
                    .Fields("IDdate") = strIDdate
                    If strResult = "PASS" Then
                        .Fields("QtyPass") = 1
                        .Fields("QtyFail") = 0
                        txtPass.Text = 1
                        txtFail.Text = 0
                    Else
                        .Fields("QtyPass") = 0
                        .Fields("QtyFail") = 1
                        txtPass.Text = 0
                        txtFail.Text = 1
                    End If
                .Update
            Else
                LastQtyPass = .Fields("QtyPass")
                LastQtyFail = .Fields("QtyFail")
                If strResult = "PASS" Then
                    .Fields("QtyPass") = Val(LastQtyPass) + 1
                        txtPass.Text = Val(LastQtyPass) + 1
                        txtFail.Text = LastQtyFail
                Else
                    .Fields("QtyFail") = Val(LastQtyFail) + 1
                        txtPass.Text = LastQtyPass
                        txtFail.Text = Val(LastQtyFail) + 1
                End If
                .Update
            End If
            .Close
        End With
        txtTotal = Val(txtPass.Text) + Val(txtFail.Text)
End Sub

Private Sub txtMeasF_Change()
    If txtMeasF.Text <> "" Or Val(txtMeasF.Text) <> 0 Then
        tmrdelayPost.Enabled = True
    End If
End Sub

Private Sub txtRef_Change()
Dim i As Integer

refer = UCase(txtRef.Text)

txtECGRef.Text = "B" & "," & refer
If refer <> "" Then
    ECGInstrument.Output = Trim(txtECGRef.Text) & vbCrLf
    ECGInstrument.Output = "U" & vbCrLf
End If



Call connect_database
Call waktu

txtLaserTemplate(4).Text = txtWeek.Text
txtLaser1(4).Text = txtWeek.Text
txtLaser2(4).Text = txtWeek.Text

            
    For i = 0 To 3
        txtLaser1(i).Text = txtLaserTemplate(i).Text
        txtLaser2(i).Text = txtLaserTemplate(i).Text
    Next i
    
    
    '///'add Dian
    If txtLaser1(0).Text <> "" Then Update_Datalog_Laser
    If txtRef.Text <> "" Then
        cmdStart(0).Enabled = True
    Else
        cmdStart(0).Enabled = False
    End If
    
    Call Update_Datalog_Laser 'Send Data To PC Laser
    Call cmdUpdatetxtPcLaser_Click
    
    Timer4.Enabled = True '140313

    '/////
End Sub

'Region "PLC Execution"
        Private Sub tmrReadPLC_Timer()
            If Winsock1.State <> 7 Then
                lblStatusPLC.Caption = "PLC Not Connecting.."
                'StatusBar1.Panels.Item(2).Text = "PLC Not Connecting.."
                tmrIndikator.Enabled = False
                shpDevConnect.BackColor = vbRed
                Winsock_Connect
            Else
                lblStatusPLC.Caption = "PLC Connected.."
                
                ReadModbus (1001)
                tmrIndikator.Enabled = True
            End If
        End Sub
        Private Sub Winsock1_DataArrival(ByVal datalength As Long)
            DataArrival (datalength)
        End Sub
        Sub SettingPLC()
            Settings.IP.Address = "192.168.0.13"
            'Settings.IP.Address = "127.0.0.1"
            Settings.IP.Port = "502"
        End Sub
        
'            Do Until Modbuswait = False
'                DoEvents
'            Loop
'End Region


