VERSION 5.00
Begin VB.Form man 
   BackColor       =   &H00008000&
   Caption         =   "Manual Mode"
   ClientHeight    =   8025
   ClientLeft      =   3435
   ClientTop       =   2190
   ClientWidth     =   9975
   LinkTopic       =   "Form2"
   ScaleHeight     =   8025
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "RCL"
      Height          =   375
      Left            =   480
      TabIndex        =   92
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00008000&
      Caption         =   "Rotate"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   91
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00008000&
      Caption         =   "Sinkron"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   90
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2760
      TabIndex        =   89
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Confirm"
      Height          =   375
      Left            =   2760
      TabIndex        =   88
      Top             =   960
      Width           =   855
   End
   Begin VB.OptionButton Check5 
      BackColor       =   &H00008000&
      Caption         =   "Output"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   87
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton Check4 
      BackColor       =   &H00008000&
      Caption         =   "Input"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   86
      Top             =   600
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   720
      Top             =   5160
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00008000&
      Caption         =   "manual ON/OFF"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   85
      Top             =   120
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   720
      Top             =   4560
   End
   Begin VB.Frame Frame4 
      Caption         =   "Unloading Station"
      Height          =   3495
      Index           =   6
      Left            =   5040
      TabIndex        =   57
      Top             =   4440
      Width           =   3495
      Begin VB.CheckBox Check2 
         Caption         =   "cyl. R"
         Height          =   255
         Left            =   2280
         TabIndex        =   82
         Top             =   2760
         Width           =   735
      End
      Begin VB.CheckBox Check54 
         Caption         =   "cyl. N"
         Height          =   375
         Index           =   6
         Left            =   2280
         TabIndex        =   63
         Top             =   2280
         Width           =   735
      End
      Begin VB.CheckBox Check53 
         Caption         =   "cyl. M"
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   62
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox Check52 
         Caption         =   "cyl. P"
         Height          =   195
         Index           =   6
         Left            =   2280
         TabIndex        =   61
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox Check51 
         Caption         =   "cyl. O"
         Height          =   195
         Index           =   6
         Left            =   2280
         TabIndex        =   60
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox Check50 
         Caption         =   "cyl. H"
         Height          =   195
         Index           =   6
         Left            =   2280
         TabIndex        =   59
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox Check49 
         Caption         =   "cyl. G"
         Height          =   195
         Index           =   6
         Left            =   2280
         TabIndex        =   58
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label71 
         Caption         =   "i0.2.29"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   84
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label71 
         Caption         =   "i0.2.28"
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   83
         Top             =   2760
         Width           =   495
      End
      Begin VB.Shape Shape62 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   1
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   3120
         Width           =   135
      End
      Begin VB.Shape Shape62 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   0
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label Label71 
         Caption         =   "i0.2.14"
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   79
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label70 
         Caption         =   "i0.2.15"
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   78
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label69 
         Caption         =   "i0.2.10"
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   77
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label68 
         Caption         =   "i0.2.11"
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   76
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label67 
         Caption         =   "i0.2.20"
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   75
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label66 
         Caption         =   "i0.2.21"
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   74
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label65 
         Caption         =   "i0.2.18"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   73
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label64 
         Caption         =   "i0.2.19"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   72
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label63 
         Caption         =   "i0.2.17"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   71
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label62 
         Caption         =   "i0.2.16"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   70
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label61 
         Caption         =   "i0.2.12"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   69
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label60 
         Caption         =   "i0.2.13"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   68
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label59 
         Caption         =   "i0.1.17"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   67
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label58 
         Caption         =   "i0.1.16"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   66
         Top             =   600
         Width           =   495
      End
      Begin VB.Shape Shape62 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape61 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   135
      End
      Begin VB.Shape Shape60 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   1680
         Width           =   135
      End
      Begin VB.Shape Shape59 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape Shape58 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   120
         Shape           =   3  'Circle
         Top             =   3120
         Width           =   135
      End
      Begin VB.Shape Shape57 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   135
      End
      Begin VB.Shape Shape56 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape55 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   960
         Width           =   135
      End
      Begin VB.Shape Shape54 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   135
      End
      Begin VB.Shape Shape53 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1680
         Width           =   135
      End
      Begin VB.Shape Shape52 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   600
         Width           =   135
      End
      Begin VB.Shape Shape51 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape Shape50 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   120
         Shape           =   3  'Circle
         Top             =   960
         Width           =   135
      End
      Begin VB.Shape Shape49 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   6
         Left            =   120
         Shape           =   3  'Circle
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label57 
         Caption         =   "Output"
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   65
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label56 
         Caption         =   "Input"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   64
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Laser Station"
      Height          =   3495
      Index           =   5
      Left            =   1680
      TabIndex        =   38
      Top             =   4440
      Width           =   3255
      Begin VB.CheckBox Check1 
         Caption         =   "cyl. Q"
         Height          =   255
         Left            =   2400
         TabIndex        =   81
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox Check52 
         Caption         =   "cyl. L"
         Height          =   195
         Index           =   5
         Left            =   2400
         TabIndex        =   42
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox Check51 
         Caption         =   "cyl. K"
         Height          =   195
         Index           =   5
         Left            =   2400
         TabIndex        =   41
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox Check50 
         Caption         =   "cyl. J"
         Height          =   195
         Index           =   5
         Left            =   2400
         TabIndex        =   40
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox Check49 
         Caption         =   "cyl. I"
         Height          =   195
         Index           =   5
         Left            =   2400
         TabIndex        =   39
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label69 
         Caption         =   "i0.1.10"
         Height          =   375
         Index           =   5
         Left            =   1320
         TabIndex        =   56
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label68 
         Caption         =   "i0.1.11"
         Height          =   375
         Index           =   5
         Left            =   1320
         TabIndex        =   55
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label67 
         Caption         =   "i0.1.9"
         Height          =   375
         Index           =   5
         Left            =   1320
         TabIndex        =   54
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label66 
         Caption         =   "i0.1.8"
         Height          =   375
         Index           =   5
         Left            =   1320
         TabIndex        =   53
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label65 
         Caption         =   "i0.1.6"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   52
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label64 
         Caption         =   "i0.1.7"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   51
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label63 
         Caption         =   "i0.1.12"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   50
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label62 
         Caption         =   "i0.1.13"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   49
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label61 
         Caption         =   "i0.1.4"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   48
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label60 
         Caption         =   "i0.1.5"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   47
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label59 
         Caption         =   "i0.1.19"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   46
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label58 
         Caption         =   "i0.1.18"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   45
         Top             =   600
         Width           =   495
      End
      Begin VB.Shape Shape60 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   1680
         Width           =   135
      End
      Begin VB.Shape Shape59 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape Shape58 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   120
         Shape           =   3  'Circle
         Top             =   3120
         Width           =   135
      End
      Begin VB.Shape Shape57 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   135
      End
      Begin VB.Shape Shape56 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape55 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   960
         Width           =   135
      End
      Begin VB.Shape Shape54 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   135
      End
      Begin VB.Shape Shape53 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1680
         Width           =   135
      End
      Begin VB.Shape Shape52 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   600
         Width           =   135
      End
      Begin VB.Shape Shape51 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape Shape50 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   120
         Shape           =   3  'Circle
         Top             =   960
         Width           =   135
      End
      Begin VB.Shape Shape49 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   5
         Left            =   120
         Shape           =   3  'Circle
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label57 
         Caption         =   "Output"
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label56 
         Caption         =   "Input"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   43
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Greasing Station"
      Height          =   2895
      Index           =   4
      Left            =   7440
      TabIndex        =   27
      Top             =   1440
      Width           =   2295
      Begin VB.CheckBox Check50 
         Caption         =   "cyl. F"
         Height          =   195
         Index           =   4
         Left            =   1320
         TabIndex        =   29
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox Check49 
         Caption         =   "cyl. E"
         Height          =   195
         Index           =   4
         Left            =   1320
         TabIndex        =   28
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label63 
         Caption         =   "i0.1.2"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   37
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label62 
         Caption         =   "i0.1.3"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   36
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label61 
         Caption         =   "i0.1.0"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   35
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label60 
         Caption         =   "i0.1.1"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   34
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label59 
         Caption         =   "i0.1.21"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   33
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label58 
         Caption         =   "i0.1.20"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   32
         Top             =   600
         Width           =   615
      End
      Begin VB.Shape Shape56 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   4
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape54 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   4
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   135
      End
      Begin VB.Shape Shape53 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   4
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1680
         Width           =   135
      End
      Begin VB.Shape Shape51 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   4
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape Shape50 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   4
         Left            =   120
         Shape           =   3  'Circle
         Top             =   960
         Width           =   135
      End
      Begin VB.Shape Shape49 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   4
         Left            =   120
         Shape           =   3  'Circle
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label57 
         Caption         =   "Output"
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label56 
         Caption         =   "Input"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "RCL Station"
      Height          =   2895
      Index           =   1
      Left            =   2640
      TabIndex        =   16
      Top             =   1440
      Width           =   2295
      Begin VB.CheckBox Check50 
         Caption         =   "cyl. B"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   18
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox Check49 
         Caption         =   "cyl. A"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   17
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Output"
         Height          =   255
         Left            =   1200
         TabIndex        =   80
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label63 
         Caption         =   "i0.2.5"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   26
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label62 
         Caption         =   "i0.2.4"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   25
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label61 
         Caption         =   "i0.2.7"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   24
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label60 
         Caption         =   "i0.2.6"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   23
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label59 
         Caption         =   "i0.1.25"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   22
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label58 
         Caption         =   "i0.1.24"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   21
         Top             =   600
         Width           =   495
      End
      Begin VB.Shape Shape56 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   1
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Shape Shape54 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   1
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   135
      End
      Begin VB.Shape Shape53 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   1
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1680
         Width           =   135
      End
      Begin VB.Shape Shape51 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   1
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape Shape50 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   1
         Left            =   120
         Shape           =   3  'Circle
         Top             =   960
         Width           =   135
      End
      Begin VB.Shape Shape49 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   1
         Left            =   120
         Shape           =   3  'Circle
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label57 
         Caption         =   "Output"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label56 
         Caption         =   "Input"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Loading Station"
      Height          =   2895
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   2295
      Begin VB.Label Label59 
         Caption         =   "i0.1.27"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label58 
         Caption         =   "i0.1.26"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.Shape Shape50 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   0
         Left            =   120
         Shape           =   3  'Circle
         Top             =   960
         Width           =   135
      End
      Begin VB.Shape Shape49 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   0
         Left            =   120
         Shape           =   3  'Circle
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label56 
         Caption         =   "Input"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ECG Station"
      Height          =   2895
      Index           =   2
      Left            =   5040
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
      Begin VB.CheckBox Check49 
         Caption         =   "cyl. C"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   3
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox Check50 
         Caption         =   "cyl. D"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   2
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label56 
         Caption         =   "Input"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label57 
         Caption         =   "Output"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Shape Shape49 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   2
         Left            =   120
         Shape           =   3  'Circle
         Top             =   600
         Width           =   135
      End
      Begin VB.Shape Shape50 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   2
         Left            =   120
         Shape           =   3  'Circle
         Top             =   960
         Width           =   135
      End
      Begin VB.Shape Shape51 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   2
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape Shape53 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   2
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1680
         Width           =   135
      End
      Begin VB.Shape Shape54 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   2
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   135
      End
      Begin VB.Shape Shape56 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   2
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label58 
         Caption         =   "i0.1.22"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label59 
         Caption         =   "i0.1.23"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label60 
         Caption         =   "i0.2.3"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label61 
         Caption         =   "i0.2.2"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label62 
         Caption         =   "i0.2.0"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label63 
         Caption         =   "i0.2.1"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Top             =   2400
         Width           =   495
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Home"
      Height          =   975
      Left            =   8760
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   5520
      Picture         =   "Form2.frx":04E0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3060
   End
End
Attribute VB_Name = "man"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'mas.Show
frmMain.Show
man.Hide

Check3.Value = 0
Check4.Value = False
Check5.Value = False

With Form1

.Text4(81).Text = "1"
.Text4(80).Text = "0"
'.CmdWrite_Click
Timer1.Enabled = False

.CmdRead = False


End With
End Sub
Private Sub Command2_Click()
With Form1

'Rcl
If Check6.Value = 1 Then
.Text4(111).Text = "1"
Else
.Text4(111).Text = "0"
End If

If Check3.Value = 1 Then
Timer1.Enabled = False
.CmdWrite = True
.Text4(80).Text = "1"
.Text4(81).Text = "0"
End If

If Check3.Value = 0 Then
Timer1.Enabled = False
.CmdWrite = False
.Text4(80).Text = "0"
.Text4(81).Text = "1"
End If

If Check4.Value = True Then
Timer1.Enabled = True
.CmdWrite = False
.CmdRead = True
End If

If Check5.Value = True Then
Timer1.Enabled = False
.CmdRead = False
.CmdWrite = True

Timer1.Enabled = False

If Check7.Value = 1 Then
.Text4(100).Text = "1"
Else
.Text4(100).Text = "0"
End If


'Rcl
If Check49(1).Value = 1 Then
.Text4(60).Text = "1"
Else
.Text4(60).Text = "0"
End If

If Check50(1).Value = 1 Then
.Text4(61).Text = "1"
Else
.Text4(61).Text = "0"
End If

'ecg

If Check49(2).Value = 1 Then
.Text4(62).Text = "1"
Else
.Text4(62).Text = "0"
End If

If Check50(2).Value = 1 Then
.Text4(63).Text = "1"
Else
.Text4(63).Text = "0"
End If

'greasing
If Check49(4).Value = 1 Then
.Text4(64).Text = "1"
Else
.Text4(64).Text = "0"
End If

If Check50(4).Value = 1 Then
.Text4(65).Text = "1"
Else
.Text4(65).Text = "0"
End If

'laser
If Check49(5).Value = 1 Then
.Text4(66).Text = "1"
Else
.Text4(66).Text = "0"
End If

If Check50(5).Value = 1 Then
.Text4(67).Text = "1"
Else
.Text4(67).Text = "0"
End If

If Check51(5).Value = 1 Then
.Text4(69).Text = "1"
Else
.Text4(69).Text = "0"
End If

If Check52(5).Value = 1 Then
.Text4(70).Text = "1"
Else
.Text4(70).Text = "0"
End If

If Check1.Value = 1 Then
.Text4(68).Text = "1"
Else
.Text4(68).Text = "0"
End If

'unloading
If Check49(6).Value = 1 Then
.Text4(71).Text = "1"
Else
.Text4(71).Text = "0"
End If

If Check50(6).Value = 1 Then
.Text4(72).Text = "1"
Else
.Text4(72).Text = "0"
End If

If Check51(6).Value = 1 Then
.Text4(73).Text = "1"
Else
.Text4(73).Text = "0"
End If

If Check52(6).Value = 1 Then
.Text4(74).Text = "1"
Else
.Text4(74).Text = "0"
End If

If Check53(6).Value = 1 Then
.Text4(75).Text = "1"
Else
.Text4(75).Text = "0"
End If

If Check54(6).Value = 1 Then
.Text4(76).Text = "1"
Else
.Text4(76).Text = "0"
End If

If Check2.Value = 1 Then
.Text4(77).Text = "1"
Else
.Text4(77).Text = "0"
End If


End If

End With
End Sub

Private Sub Command3_Click()
'If mas.MSComm2.PortOpen = True Then
'
'
'MSComm2.Output = Chr(27) + "8" 'go to trigger
'MSComm2.Output = Chr(27) + "7" 'asking status register
'
'End If
End Sub

Private Sub Form_Load()
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error GoTo lanjut
With Form1

If .Text4(80) = "1" Then
Timer1.Enabled = True
.CmdWrite = False
End If

If .CmdWrite = False Then
.CmdRead = True
End If

''loading station
'If .Text4(47).Text = 1 Then
'Shape49(0).BackColor = vbGreen
'Else
'Shape49(0).BackColor = vbRed
'End If
'
'If .Text4(48).Text = 1 Then
'Shape50(0).BackColor = vbGreen
'Else
'Shape50(0).BackColor = vbRed
'End If
'
''rcl station
'If .Text4(5).Text = 1 Then
'Shape49(1).BackColor = vbGreen
'Else
'Shape49(1).BackColor = vbRed
'End If
'
'If .Text4(6).Text = 1 Then
'Shape50(1).BackColor = vbGreen
'Else
'Shape50(1).BackColor = vbRed
'End If
'
'If .Text4(1).Text = 1 Then
'Shape51(1).BackColor = vbGreen
'Else
'Shape51(1).BackColor = vbRed
'End If
'
'If .Text4(2).Text = 1 Then
'Shape53(1).BackColor = vbGreen
'Else
'Shape53(1).BackColor = vbRed
'End If
'
'If .Text4(3).Text = 1 Then
'Shape54(1).BackColor = vbGreen
'Else
'Shape54(1).BackColor = vbRed
'End If
'
'If .Text4(4).Text = 1 Then
'Shape56(1).BackColor = vbGreen
'Else
'Shape56(1).BackColor = vbRed
'End If
'
'
''ecg station
'If .Text4(11).Text = 1 Then
'Shape49(2).BackColor = vbGreen
'Else
'Shape49(2).BackColor = vbRed
'End If
'
'If .Text4(12).Text = 1 Then
'Shape50(2).BackColor = vbGreen
'Else
'Shape50(2).BackColor = vbRed
'End If
'
'If .Text4(7).Text = 1 Then
'Shape51(2).BackColor = vbGreen
'Else
'Shape51(2).BackColor = vbRed
'End If
'
'If .Text4(8).Text = 1 Then
'Shape53(2).BackColor = vbGreen
'Else
'Shape53(2).BackColor = vbRed
'End If
'
'If .Text4(9).Text = 1 Then
'Shape54(2).BackColor = vbGreen
'Else
'Shape54(2).BackColor = vbRed
'End If
'
'If .Text4(10).Text = 1 Then
'Shape56(2).BackColor = vbGreen
'Else
'Shape56(2).BackColor = vbRed
'End If
'
'
''greasing station
'If .Text4(17).Text = 1 Then
'Shape49(4).BackColor = vbGreen
'Else
'Shape49(4).BackColor = vbRed
'End If
'
'If .Text4(18).Text = 1 Then
'Shape50(4).BackColor = vbGreen
'Else
'Shape50(4).BackColor = vbRed
'End If
'
'If .Text4(13).Text = 1 Then
'Shape51(4).BackColor = vbGreen
'Else
'Shape51(4).BackColor = vbRed
'End If
'
'If .Text4(14).Text = 1 Then
'Shape53(4).BackColor = vbGreen
'Else
'Shape53(4).BackColor = vbRed
'End If
'
'If .Text4(15).Text = 1 Then
'Shape54(4).BackColor = vbGreen
'Else
'Shape54(4).BackColor = vbRed
'End If
'
'If .Text4(16).Text = 1 Then
'Shape56(4).BackColor = vbGreen
'Else
'Shape56(4).BackColor = vbRed
'End If
'
''laser station
'If .Text4(29).Text = 1 Then
'Shape49(5).BackColor = vbGreen
'Else
'Shape49(5).BackColor = vbRed
'End If
'
'If .Text4(30).Text = 1 Then
'Shape50(5).BackColor = vbGreen
'Else
'Shape50(5).BackColor = vbRed
'End If
'
'If .Text4(19).Text = 1 Then
'Shape51(5).BackColor = vbGreen
'Else
'Shape51(5).BackColor = vbRed
'End If
'
'If .Text4(20).Text = 1 Then
'Shape53(5).BackColor = vbGreen
'Else
'Shape53(5).BackColor = vbRed
'End If
'
'If .Text4(21).Text = 1 Then
'Shape54(5).BackColor = vbGreen
'Else
'Shape54(5).BackColor = vbRed
'End If
'
'If .Text4(22).Text = 1 Then
'Shape56(5).BackColor = vbGreen
'Else
'Shape56(5).BackColor = vbRed
'End If
'
'If .Text4(25).Text = 1 Then
'Shape57(5).BackColor = vbGreen
'Else
'Shape57(5).BackColor = vbRed
'End If
'
'If .Text4(26).Text = 1 Then
'Shape58(5).BackColor = vbGreen
'Else
'Shape58(5).BackColor = vbRed
'End If
'
'If .Text4(27).Text = 1 Then
'Shape52(5).BackColor = vbGreen
'Else
'Shape52(5).BackColor = vbRed
'End If
'
'If .Text4(28).Text = 1 Then
'Shape55(5).BackColor = vbGreen
'Else
'Shape55(5).BackColor = vbRed
'End If
'
'If .Text4(23).Text = 1 Then
'Shape59(5).BackColor = vbGreen
'Else
'Shape59(5).BackColor = vbRed
'End If
'
'If .Text4(24).Text = 1 Then
'Shape60(5).BackColor = vbGreen
'Else
'Shape60(5).BackColor = vbRed
'End If
'
'
''unloading station
'If .Text4(45).Text = 1 Then
'Shape49(6).BackColor = vbGreen
'Else
'Shape49(6).BackColor = vbRed
'End If
'
'If .Text4(46).Text = 1 Then
'Shape50(6).BackColor = vbGreen
'Else
'Shape50(6).BackColor = vbRed
'End If
'
'If .Text4(31).Text = 1 Then
'Shape51(6).BackColor = vbGreen
'Else
'Shape51(6).BackColor = vbRed
'End If
'
'If .Text4(32).Text = 1 Then
'Shape53(6).BackColor = vbGreen
'Else
'Shape53(6).BackColor = vbRed
'End If
'
'If .Text4(33).Text = 1 Then
'Shape54(6).BackColor = vbGreen
'Else
'Shape54(6).BackColor = vbRed
'End If
'
'If .Text4(34).Text = 1 Then
'Shape56(6).BackColor = vbGreen
'Else
'Shape56(6).BackColor = vbRed
'End If
'
'If .Text4(35).Text = 1 Then
'Shape57(6).BackColor = vbGreen
'Else
'Shape57(6).BackColor = vbRed
'End If
'
'If .Text4(36).Text = 1 Then
'Shape58(6).BackColor = vbGreen
'Else
'Shape58(6).BackColor = vbRed
'End If
'
'If .Text4(37).Text = 1 Then
'Shape52(6).BackColor = vbGreen
'Else
'Shape52(6).BackColor = vbRed
'End If
'
'If .Text4(38).Text = 1 Then
'Shape55(6).BackColor = vbGreen
'Else
'Shape55(6).BackColor = vbRed
'End If
'
'If .Text4(39).Text = 1 Then
'Shape59(6).BackColor = vbGreen
'Else
'Shape59(6).BackColor = vbRed
'End If
'
'If .Text4(40).Text = 1 Then
'Shape60(6).BackColor = vbGreen
'Else
'Shape60(6).BackColor = vbRed
'End If
'
'If .Text4(41).Text = 1 Then
'Shape61(6).BackColor = vbGreen
'Else
'Shape61(6).BackColor = vbRed
'End If
'
'If .Text4(42).Text = 1 Then
'Shape62(6).BackColor = vbGreen
'Else
'Shape62(6).BackColor = vbRed
'End If
'
'If .Text4(43).Text = 1 Then
'Shape62(0).BackColor = vbGreen
'Else
'Shape62(0).BackColor = vbRed
'End If
'
'If .Text4(44).Text = 1 Then
'Shape62(1).BackColor = vbGreen
'Else
'Shape62(1).BackColor = vbRed
'End If

End With
lanjut:
End Sub

Private Sub Timer2_Timer()
'Rcl
If Check49(1).Value = 1 Then
Text1.Text = Check49(1).Caption
Else
Text1.Text = ""
End If
If Check50(1).Value = 1 Then
Text1.Text = Check50(1).Caption
End If

'ecg

If Check49(2).Value = 1 Then
Text1.Text = Check49(2).Caption
End If

If Check50(2).Value = 1 Then
Text1.Text = Check50(2).Caption
End If

'greasing
If Check49(4).Value = 1 Then
Text1.Text = Check49(4).Caption
End If

If Check50(4).Value = 1 Then
Text1.Text = Check50(4).Caption
End If

'laser
If Check49(5).Value = 1 Then
Text1.Text = Check49(5).Caption
End If

If Check50(5).Value = 1 Then
Text1.Text = Check50(5).Caption
End If

If Check51(5).Value = 1 Then
Text1.Text = Check51(5).Caption
End If

If Check52(5).Value = 1 Then
Text1.Text = Check52(5).Caption
End If

If Check1.Value = 1 Then
Text1.Text = Check1.Caption
End If

'unloading
If Check49(6).Value = 1 Then
Text1.Text = Check49(6).Caption
End If

If Check50(6).Value = 1 Then
Text1.Text = Check50(6).Caption
End If

If Check51(6).Value = 1 Then
Text1.Text = Check51(6).Caption
End If

If Check52(6).Value = 1 Then
Text1.Text = Check52(6).Caption
End If

If Check53(6).Value = 1 Then
Text1.Text = Check53(6).Caption
End If

If Check54(6).Value = 1 Then
Text1.Text = Check54(6).Caption
End If

If Check2.Value = 1 Then
Text1.Text = Check2.Caption
End If

End Sub
