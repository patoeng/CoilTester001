VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "                                                                                                                          "
   ClientHeight    =   4080
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer7 
      Interval        =   500
      Left            =   1560
      Top             =   120
   End
   Begin VB.Timer Timer6 
      Interval        =   500
      Left            =   1080
      Top             =   120
   End
   Begin VB.CommandButton Command5 
      Caption         =   "1"
      Height          =   255
      Left            =   4320
      TabIndex        =   153
      Top             =   2040
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   12360
      TabIndex        =   152
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "1"
      Height          =   255
      Left            =   4320
      TabIndex        =   151
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "0"
      Height          =   255
      Left            =   3960
      TabIndex        =   150
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "0"
      Height          =   255
      Left            =   3960
      TabIndex        =   149
      Top             =   2040
      Width           =   255
   End
   Begin VB.Timer Timer5 
      Interval        =   500
      Left            =   14640
      Top             =   3960
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   14640
      Top             =   3000
   End
   Begin VB.Timer Timer4 
      Interval        =   500
      Left            =   14640
      Top             =   3480
   End
   Begin VB.TextBox txtRCL_OK 
      Height          =   285
      Left            =   5760
      TabIndex        =   148
      Text            =   "txtRCL_OK"
      Top             =   3720
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   14640
      Top             =   4440
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   1440
      TabIndex        =   136
      Text            =   "Text6"
      Top             =   9960
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   119
      Left            =   11280
      TabIndex        =   134
      Text            =   "0"
      Top             =   7680
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   118
      Left            =   11280
      TabIndex        =   133
      Text            =   "0"
      Top             =   7320
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   117
      Left            =   11280
      TabIndex        =   132
      Text            =   "0"
      Top             =   6960
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   116
      Left            =   11280
      TabIndex        =   131
      Text            =   "0"
      Top             =   6600
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   115
      Left            =   11280
      TabIndex        =   130
      Text            =   "0"
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   114
      Left            =   11280
      TabIndex        =   129
      Text            =   "0"
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   113
      Left            =   11280
      TabIndex        =   128
      Text            =   "0"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   112
      Left            =   11280
      TabIndex        =   127
      Text            =   "0"
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   111
      Left            =   11280
      TabIndex        =   126
      Text            =   "0"
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   110
      Left            =   11280
      TabIndex        =   125
      Text            =   "0"
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   109
      Left            =   11280
      TabIndex        =   124
      Text            =   "0"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   108
      Left            =   11280
      TabIndex        =   123
      Text            =   "0"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   107
      Left            =   11280
      TabIndex        =   122
      Text            =   "0"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   106
      Left            =   11280
      TabIndex        =   121
      Text            =   "0"
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   105
      Left            =   11280
      TabIndex        =   120
      Text            =   "0"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   104
      Left            =   11280
      TabIndex        =   119
      Text            =   "0"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   103
      Left            =   11280
      TabIndex        =   118
      Text            =   "0"
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   102
      Left            =   11280
      TabIndex        =   117
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   101
      Left            =   11280
      TabIndex        =   116
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   100
      Left            =   11280
      TabIndex        =   115
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   99
      Left            =   10800
      TabIndex        =   114
      Text            =   "0"
      Top             =   7680
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   98
      Left            =   10800
      TabIndex        =   113
      Text            =   "0"
      Top             =   7320
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   97
      Left            =   10800
      TabIndex        =   112
      Text            =   "0"
      Top             =   6960
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   96
      Left            =   10800
      TabIndex        =   111
      Text            =   "0"
      Top             =   6600
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   95
      Left            =   10800
      TabIndex        =   110
      Text            =   "0"
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   94
      Left            =   10800
      TabIndex        =   109
      Text            =   "0"
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   93
      Left            =   10800
      TabIndex        =   108
      Text            =   "0"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   92
      Left            =   10800
      TabIndex        =   107
      Text            =   "0"
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   91
      Left            =   10800
      TabIndex        =   106
      Text            =   "0"
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   90
      Left            =   10800
      TabIndex        =   105
      Text            =   "0"
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   89
      Left            =   10800
      TabIndex        =   104
      Text            =   "0"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   88
      Left            =   10800
      TabIndex        =   103
      Text            =   "0"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   87
      Left            =   10800
      TabIndex        =   102
      Text            =   "0"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   86
      Left            =   10800
      TabIndex        =   101
      Text            =   "0"
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   85
      Left            =   10800
      TabIndex        =   100
      Text            =   "0"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   84
      Left            =   10800
      TabIndex        =   99
      Text            =   "0"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   83
      Left            =   10800
      TabIndex        =   98
      Text            =   "0"
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   82
      Left            =   10800
      TabIndex        =   97
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   81
      Left            =   2160
      TabIndex        =   96
      Text            =   "0"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   80
      Left            =   2160
      TabIndex        =   95
      Text            =   "0"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   79
      Left            =   10320
      TabIndex        =   94
      Text            =   "0"
      Top             =   7680
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   78
      Left            =   10320
      TabIndex        =   93
      Text            =   "0"
      Top             =   7320
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   77
      Left            =   10320
      TabIndex        =   92
      Text            =   "0"
      Top             =   6960
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   76
      Left            =   10320
      TabIndex        =   91
      Text            =   "0"
      Top             =   6600
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   75
      Left            =   10320
      TabIndex        =   90
      Text            =   "0"
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   74
      Left            =   10320
      TabIndex        =   89
      Text            =   "0"
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   73
      Left            =   10320
      TabIndex        =   88
      Text            =   "0"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   72
      Left            =   10320
      TabIndex        =   87
      Text            =   "0"
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   71
      Left            =   2160
      TabIndex        =   86
      Text            =   "0"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   70
      Left            =   2160
      TabIndex        =   85
      Text            =   "0"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   69
      Left            =   10320
      TabIndex        =   84
      Text            =   "0"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   68
      Left            =   10320
      TabIndex        =   83
      Text            =   "0"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   67
      Left            =   10320
      TabIndex        =   82
      Text            =   "0"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   66
      Left            =   10320
      TabIndex        =   81
      Text            =   "0"
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   65
      Left            =   10320
      TabIndex        =   80
      Text            =   "0"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   64
      Left            =   10320
      TabIndex        =   79
      Text            =   "0"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   63
      Left            =   10320
      TabIndex        =   78
      Text            =   "0"
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   62
      Left            =   10320
      TabIndex        =   77
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   61
      Left            =   10320
      TabIndex        =   76
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   60
      Left            =   10320
      TabIndex        =   75
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   59
      Left            =   9840
      TabIndex        =   74
      Text            =   "0"
      Top             =   7680
      Width           =   375
   End
   Begin VB.TextBox txbDataLength 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   13080
      TabIndex        =   72
      Text            =   "0"
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton CmdConnect 
      Caption         =   "Connect"
      Height          =   495
      Left            =   13080
      TabIndex        =   67
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton CmdDissconect 
      Caption         =   "Disconnect"
      Height          =   495
      Left            =   13080
      TabIndex        =   66
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   13080
      TabIndex        =   65
      Text            =   "1000"
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   13080
      TabIndex        =   64
      Text            =   "120"
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton CmdRead 
      Caption         =   "Read registers"
      Height          =   495
      Left            =   13080
      TabIndex        =   63
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   62
      Text            =   "0"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   61
      Text            =   "0"
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   60
      Text            =   "0"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   59
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   58
      Text            =   "0"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   480
      TabIndex        =   57
      Text            =   "0"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   480
      TabIndex        =   56
      Text            =   "0"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   480
      TabIndex        =   55
      Text            =   "0"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   480
      TabIndex        =   54
      Text            =   "0"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   480
      TabIndex        =   53
      Text            =   "0"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   2160
      TabIndex        =   52
      Text            =   "0"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   2160
      TabIndex        =   51
      Text            =   "0"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   8880
      TabIndex        =   50
      Text            =   "0"
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   8880
      TabIndex        =   49
      Text            =   "0"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   8880
      TabIndex        =   48
      Text            =   "0"
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   8880
      TabIndex        =   47
      Text            =   "0"
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   8880
      TabIndex        =   46
      Text            =   "0"
      Top             =   6600
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   8880
      TabIndex        =   45
      Text            =   "0"
      Top             =   6960
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   8880
      TabIndex        =   44
      Tag             =   "T19"
      Text            =   "0"
      Top             =   7320
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   19
      Left            =   8880
      TabIndex        =   43
      Text            =   "0"
      Top             =   7680
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   20
      Left            =   2160
      TabIndex        =   42
      Text            =   "0"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   21
      Left            =   2160
      TabIndex        =   41
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   22
      Left            =   9360
      TabIndex        =   40
      Text            =   "0"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   23
      Left            =   9360
      TabIndex        =   39
      Text            =   "0"
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   24
      Left            =   9360
      TabIndex        =   38
      Text            =   "0"
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   25
      Left            =   9360
      TabIndex        =   37
      Text            =   "0"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   26
      Left            =   9360
      TabIndex        =   36
      Text            =   "0"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   27
      Left            =   9360
      TabIndex        =   35
      Text            =   "0"
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   28
      Left            =   9360
      TabIndex        =   34
      Text            =   "0"
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   29
      Left            =   9360
      TabIndex        =   33
      Text            =   "0"
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   30
      Left            =   2160
      TabIndex        =   32
      Text            =   "0"
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   31
      Left            =   9360
      TabIndex        =   31
      Text            =   "0"
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   32
      Left            =   9360
      TabIndex        =   30
      Text            =   "0"
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   33
      Left            =   9360
      TabIndex        =   29
      Text            =   "0"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   34
      Left            =   9360
      TabIndex        =   28
      Text            =   "0"
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   35
      Left            =   9360
      TabIndex        =   27
      Text            =   "0"
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   36
      Left            =   9360
      TabIndex        =   26
      Text            =   "0"
      Top             =   6600
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   37
      Left            =   9360
      TabIndex        =   25
      Text            =   "0"
      Top             =   6960
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   38
      Left            =   9360
      TabIndex        =   24
      Text            =   "0"
      Top             =   7320
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   39
      Left            =   9360
      TabIndex        =   23
      Text            =   "0"
      Top             =   7680
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   40
      Left            =   2160
      TabIndex        =   22
      Text            =   "0"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   41
      Left            =   9840
      TabIndex        =   21
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   42
      Left            =   9840
      TabIndex        =   20
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   43
      Left            =   9840
      TabIndex        =   19
      Text            =   "0"
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   44
      Left            =   9840
      TabIndex        =   18
      Text            =   "0"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   45
      Left            =   9840
      TabIndex        =   17
      Text            =   "0"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   46
      Left            =   9840
      TabIndex        =   16
      Text            =   "0"
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   47
      Left            =   9840
      TabIndex        =   15
      Text            =   "0"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   48
      Left            =   9840
      TabIndex        =   14
      Text            =   "0"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   49
      Left            =   9840
      TabIndex        =   13
      Text            =   "0"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   50
      Left            =   9840
      TabIndex        =   12
      Text            =   "0"
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   51
      Left            =   9840
      TabIndex        =   11
      Text            =   "0"
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   52
      Left            =   9840
      TabIndex        =   10
      Text            =   "0"
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   53
      Left            =   9840
      TabIndex        =   9
      Text            =   "0"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   54
      Left            =   9840
      TabIndex        =   8
      Text            =   "0"
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   55
      Left            =   9840
      TabIndex        =   7
      Text            =   "0"
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   56
      Left            =   9840
      TabIndex        =   6
      Text            =   "0"
      Top             =   6600
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   57
      Left            =   9840
      TabIndex        =   5
      Text            =   "0"
      Top             =   6960
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1045
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   58
      Left            =   9840
      TabIndex        =   4
      Text            =   "0"
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton CmdWrite 
      Caption         =   "Write Registers"
      Height          =   495
      Left            =   13080
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   13080
      TabIndex        =   2
      Text            =   "192.168.0.13"
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13080
      TabIndex        =   1
      Text            =   "Disconnected"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   14640
      Top             =   4920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   13080
      TabIndex        =   0
      Top             =   4920
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock1_Unused 
      Left            =   14640
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "85.16.23.23"
      RemotePort      =   502
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LASERDATA CHANGE"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2760
      TabIndex        =   156
      Top             =   240
      Width           =   1680
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID LASER"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1080
      TabIndex        =   155
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLEAR"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1080
      TabIndex        =   154
      Top             =   3120
      Width           =   525
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ECG FINISH"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2760
      TabIndex        =   147
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RCL FINISH"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2760
      TabIndex        =   146
      Top             =   3480
      Width           =   885
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RCL 2"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2760
      TabIndex        =   145
      Top             =   2760
      Width           =   450
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RCL 1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2760
      TabIndex        =   144
      Top             =   2400
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ECG ST 2"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2760
      TabIndex        =   143
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ECG ST 1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2760
      TabIndex        =   142
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ECG OK"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2760
      TabIndex        =   141
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RCL OK"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2760
      TabIndex        =   140
      Top             =   3120
      Width           =   585
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID UNLOAD"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1080
      TabIndex        =   139
      Top             =   2760
      Width           =   885
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID ECG"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1080
      TabIndex        =   138
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID RCL"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1080
      TabIndex        =   137
      Top             =   960
      Width           =   525
   End
   Begin VB.Label Label8 
      Caption         =   "SINKRON"
      Height          =   255
      Left            =   12240
      TabIndex        =   135
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "DataLength"
      Height          =   255
      Left            =   13080
      TabIndex        =   73
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Start register"
      Height          =   375
      Left            =   13080
      TabIndex        =   71
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Length"
      Height          =   255
      Left            =   13080
      TabIndex        =   70
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Adrress IP ETZ"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13080
      TabIndex        =   69
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Status"
      Height          =   255
      Left            =   13080
      TabIndex        =   68
      Top             =   6120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command2_Click()
    WriteModbus 1081, 0
End Sub

Private Sub Command3_Click()
    WriteModbus 1071, 0
End Sub

Private Sub Command4_Click()
    WriteModbus 1071, 1

End Sub

Private Sub Command5_Click()
    WriteModbus 1081, 1
End Sub


Private Sub Timer6_Timer() 'Capture Value Change
    Dim activeval As Integer
    If Text4(40).Text = 1 Then
    
        If Val(Text4(8).Text) >= 3 And Val(Text4(8).Text) <= 8 Then
            activeval = Val(Text4(8).Text) - 2
        ElseIf Val(Text4(8).Text) = 1 Then
            activeval = 7
        ElseIf Val(Text4(8).Text) = 2 Then
            activeval = 8
        End If
        
        frmMain.TxtLaserID.Text = Text4(5).Text
        
        frmMain.txtClearPosition.Text = activeval
        Timer6.Enabled = False
        Timer7.Enabled = True
    End If
End Sub

Private Sub Timer7_Timer()
    If Text4(40).Text = 0 Then
        Timer6.Enabled = True
        Timer7.Enabled = False
End If
End Sub

