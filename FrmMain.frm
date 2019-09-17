VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00996666&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7860
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChckOnTop 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      Caption         =   "Always On Top"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   130
      Top             =   7400
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7080
      Top             =   6960
   End
   Begin VB.TextBox txtType 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4920
      Width           =   6495
   End
   Begin VB.TextBox txtSubType 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5280
      Width           =   6495
   End
   Begin VB.TextBox txtFunctionKeys 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   5640
      Width           =   6495
   End
   Begin VB.TextBox txtLayoutID 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   6000
      Width           =   6495
   End
   Begin VB.TextBox txtLayoutName 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   6360
      Width           =   6495
   End
   Begin VB.TextBox txtEventLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1275
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "FrmMain.frx":1272
      Top             =   3240
      Width           =   8580
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "AVACO "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   132
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Keyboard Diagnostic 2002"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   480
      TabIndex        =   131
      Top             =   7190
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   3000
      X2              =   3000
      Y1              =   270
      Y2              =   -120
   End
   Begin VB.Label LblCheckAgain 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Check Again"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "FrmMain.frx":12A7
      MousePointer    =   99  'Custom
      TabIndex        =   129
      Top             =   30
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   60
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   7680
      Width           =   8565
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      X1              =   6960
      X2              =   6960
      Y1              =   270
      Y2              =   -120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00404040&
      X1              =   7800
      X2              =   7800
      Y1              =   270
      Y2              =   -120
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00404040&
      X1              =   1560
      X2              =   1560
      Y1              =   270
      Y2              =   -120
   End
   Begin VB.Label LblAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6960
      MouseIcon       =   "FrmMain.frx":15B1
      MousePointer    =   99  'Custom
      TabIndex        =   128
      Top             =   30
      Width           =   855
   End
   Begin VB.Label LblExit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7800
      MouseIcon       =   "FrmMain.frx":18BB
      MousePointer    =   99  'Custom
      TabIndex        =   127
      Top             =   30
      Width           =   855
   End
   Begin VB.Label LblRefKeys 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh All Keys"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      MouseIcon       =   "FrmMain.frx":1BC5
      MousePointer    =   99  'Custom
      TabIndex        =   126
      Top             =   30
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5400
      Picture         =   "FrmMain.frx":1ECF
      Top             =   7200
      Width           =   480
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   420
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   -120
      Width           =   8565
   End
   Begin VB.Label Label70 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Left            =   120
      TabIndex        =   125
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label69 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scroll"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   124
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label68 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Caps"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   123
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Num"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   122
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      FillColor       =   &H00404040&
      Height          =   2535
      Left            =   120
      Top             =   600
      Width           =   8535
   End
   Begin VB.Label Label119 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   121
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label118 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   120
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label55 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4200
      TabIndex        =   119
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   5400
      TabIndex        =   118
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label115 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   117
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label114 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   116
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label113 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   115
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label107 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ent"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   114
      Top             =   2520
      Width           =   390
   End
   Begin VB.Label Label106 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   113
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label100 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Space"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   112
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label99 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Alt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   111
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label98 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   840
      TabIndex        =   110
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label96 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   109
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label94 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   108
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label82 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   107
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label81 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   106
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label80 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   105
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label67 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   104
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label66 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   103
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label65 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lf"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   102
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label64 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   101
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label63 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   100
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label62 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   99
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label54 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mnu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   98
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Alt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   97
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   96
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   95
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   94
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   93
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   92
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   91
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   90
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   89
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label LblEnter 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5190
      TabIndex        =   88
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label117 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   87
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label116 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pd"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   86
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label87 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6250
      TabIndex        =   85
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label97 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Caps"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   84
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label88 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "'"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   83
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label86 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   82
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label61 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   81
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label60 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   80
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label59 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   79
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   78
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   77
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   76
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   75
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   74
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   73
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   72
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   71
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   70
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label104 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   69
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label112 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   68
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label110 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   67
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label105 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   66
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label92 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<--->"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   65
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label85 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   64
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label84 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "["
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   63
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label83 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   62
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label73 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   61
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label72 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "pup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   60
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label71 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ins"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   59
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   58
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label57 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   57
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   56
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   55
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   54
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   53
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   52
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   51
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   50
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label79 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "\"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   49
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label109 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   48
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label91 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   47
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label89 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   46
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label76 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scrl"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   45
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label75 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Brk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   44
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label74 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Prt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   43
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label52 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   42
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<---"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   41
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label50 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   40
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   39
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   38
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   37
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   36
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   35
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   34
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "`"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   32
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   30
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   29
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   28
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label111 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   27
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label95 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label93 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label90 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   24
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Sleep 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Slp"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   6240
      TabIndex        =   23
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label78 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "wup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   6610
      TabIndex        =   22
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label77 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pwr"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   5890
      TabIndex        =   21
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F11"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   19
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   17
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape Tabtb 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape num0 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   735
   End
   Begin VB.Shape del 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   375
   End
   Begin VB.Shape enter_num 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape num3 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape num2 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7440
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape num1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape num6 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape num5 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7440
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape num4 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape plus2 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape minus 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape num9 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape num8 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7440
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape num7 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Star 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape slash 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7440
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape numlock 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Delete 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   1935
      Width           =   375
   End
   Begin VB.Shape End1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   1935
      Width           =   375
   End
   Begin VB.Shape PgDwn 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   1935
      Width           =   375
   End
   Begin VB.Shape Up 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   2295
      Width           =   375
   End
   Begin VB.Shape Right 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   2655
      Width           =   375
   End
   Begin VB.Shape Down 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   2655
      Width           =   375
   End
   Begin VB.Shape Left1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   2655
      Width           =   375
   End
   Begin VB.Shape PgUp 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   1575
      Width           =   375
   End
   Begin VB.Shape Home 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   1575
      Width           =   375
   End
   Begin VB.Shape Insert 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   1575
      Width           =   375
   End
   Begin VB.Shape PBreak 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   1215
      Width           =   375
   End
   Begin VB.Shape ScrollLock 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   1215
      Width           =   375
   End
   Begin VB.Shape PrintScrn 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   1215
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5400
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape ctrl_r 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape popup 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4680
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape Alt_r 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape Start_r 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape space 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Shape start_l 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape alt_l 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shift_R 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4680
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   735
   End
   Begin VB.Shape GmKn 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape Titik 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape Comma 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3600
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape m 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape n 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape b 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape v 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape c 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape x 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape z 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape Enter2 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape enter 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape q 
      BackColor       =   &H00000000&
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape cmats 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape ttcm 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape l 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape k 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape j 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape h 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape g 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape f 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape d 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape s 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape a 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape Krng_kn 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape krng_kr 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape p 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape o 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape i 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape u 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape y 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2640
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape t 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape r 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape e 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape w 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Number1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Number2 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Number3 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Number4 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Number5 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Number6 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Number7 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Number8 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Number9 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Number0 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape min 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape plus 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape back 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5280
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   495
   End
   Begin VB.Shape gmkr 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape esc 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape f2 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape f3 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape f4 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape f5 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2640
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape f6 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape f7 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape f8 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape f9 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape f10 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4680
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape f11 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape f12 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5400
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape Power 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape Sleep1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape WakeUp 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape nekaj 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape capslock 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   735
   End
   Begin VB.Shape shift_L 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   855
   End
   Begin VB.Shape ctrl_l 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape f1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape ShpCapsLight 
      BorderColor     =   &H00404040&
      FillColor       =   &H00008080&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   720
      Width           =   195
   End
   Begin VB.Shape ShpScrollLight 
      BorderColor     =   &H00404040&
      FillColor       =   &H00008080&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   720
      Width           =   195
   End
   Begin VB.Shape ShpNumLight 
      BorderColor     =   &H00404040&
      FillColor       =   &H00008080&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   720
      Width           =   195
   End
   Begin VB.Label LblKeybInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Keyboard Information :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label lblLayoutName 
      BackStyle       =   0  'Transparent
      Caption         =   "Layout Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label lblFunctionKeys 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Function Keys"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label lblSubType 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label lblLayoutID 
      BackStyle       =   0  'Transparent
      Caption         =   "Layout ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label LblKeyDiag 
      BackStyle       =   0  'Transparent
      Caption         =   "Keyboard Diagnostic :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   330
      Width           =   2295
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--> Avaco Keyboard Diagnostic
'--> version 1.00
'--> Version Language : English
'--> By Agus Ramadhani
'--> avaco software
'--> http://avaco-software.tripod.com
'--> avaco@9cy.Com
'--> 2002-2003
'--> Don't forget to Vote :)



Private Sub Form_Load()
    ChckOnTop.Value = 1
    MeOnTop Me
    Timer1.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
With txtEventLog
Select Case Shift
Case 1
End Select
Select Case KeyCode

Case 27 'Esc
     esc.FillColor = &HFF&
    .Text = txtEventLog.Text & "Esc  Key Pressed" & " - KeyCode : " & KeyCode & Shift & vbCrLf
Case 112 'F1
     f1.FillColor = &HFF&
    .Text = txtEventLog.Text & "F1  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 113 'F2
     f2.FillColor = &HFF&
    .Text = txtEventLog.Text & "F2  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 114 'F3
     f3.FillColor = &HFF&
     .Text = txtEventLog.Text & "F3  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 115 'F4
     f4.FillColor = &HFF&
     .Text = txtEventLog.Text & "F4  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 116 'F5
     f5.FillColor = &HFF&
     .Text = txtEventLog.Text & "F5  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 117 'F5
     f6.FillColor = &HFF&
     .Text = txtEventLog.Text & "F6  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 118 'F7
     f7.FillColor = &HFF&
     .Text = txtEventLog.Text & "F7  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 119 'F8
     f8.FillColor = &HFF&
     .Text = txtEventLog.Text & "F8  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 120 'F9
     f9.FillColor = &HFF&
     .Text = txtEventLog.Text & "F9  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 121 'F10
     f10.FillColor = &HFF&
     .Text = txtEventLog.Text & "F10  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 122 'F11
     f11.FillColor = &HFF&
     .Text = txtEventLog.Text & "F11  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 123 'F12
     f12.FillColor = &HFF&
     .Text = txtEventLog.Text & "F12  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 192 '`
      nekaj.FillColor = &HFF&
     .Text = txtEventLog.Text & "`  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 49 '1
      Number1.FillColor = &HFF&
     .Text = txtEventLog.Text & "1  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 50 '2
     Number2.FillColor = &HFF&
     .Text = txtEventLog.Text & "2  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 51 '3
     Number3.FillColor = &HFF&
     .Text = txtEventLog.Text & "3  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 52 '4
     Number4.FillColor = &HFF&
     .Text = txtEventLog.Text & "4  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 53 '5
     Number5.FillColor = &HFF&
     .Text = txtEventLog.Text & "5  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 54 '6
     Number6.FillColor = &HFF&
     .Text = txtEventLog.Text & "6  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 55 '7
     Number7.FillColor = &HFF&
     .Text = txtEventLog.Text & "7  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 56 '8
     Number8.FillColor = &HFF&
     .Text = txtEventLog.Text & "8  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 57 '9
     Number9.FillColor = &HFF&
     .Text = txtEventLog.Text & "9  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 48 '0
     Number0.FillColor = &HFF&
     .Text = txtEventLog.Text & "0  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 65 'A
      a.FillColor = &HFF&
     .Text = txtEventLog.Text & "A  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 66 'B
      b.FillColor = &HFF&
     .Text = txtEventLog.Text & "B  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 67 'C
      c.FillColor = &HFF&
     .Text = txtEventLog.Text & "C  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 68 'D
      d.FillColor = &HFF&
     .Text = txtEventLog.Text & "D  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 69 'E
      e.FillColor = &HFF&
     .Text = txtEventLog.Text & "E  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 70 'F
     f.FillColor = &HFF&
     .Text = txtEventLog.Text & "F  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 71 'G
     g.FillColor = &HFF&
     .Text = txtEventLog.Text & "G  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 72 'H
      h.FillColor = &HFF&
     .Text = txtEventLog.Text & "H  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 73 'I
      i.FillColor = &HFF&
     .Text = txtEventLog.Text & "I  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 74 'J
      j.FillColor = &HFF&
     .Text = txtEventLog.Text & "J  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 75 'K
     k.FillColor = &HFF&
     .Text = txtEventLog.Text & "K  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 76 'L
      l.FillColor = &HFF&
     .Text = txtEventLog.Text & "L  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 77 'M
      m.FillColor = &HFF&
     .Text = txtEventLog.Text & "M  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 78 'N
      n.FillColor = &HFF&
     .Text = txtEventLog.Text & "N  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 79 'O
      o.FillColor = &HFF&
     .Text = txtEventLog.Text & "O  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 80 'P
      p.FillColor = &HFF&
     .Text = txtEventLog.Text & "P  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 81 'Q
     q.FillColor = &HFF&
     .Text = txtEventLog.Text & "Q  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 82 'R
      r.FillColor = &HFF&
     .Text = txtEventLog.Text & "R  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 83 'S
      s.FillColor = &HFF&
     .Text = txtEventLog.Text & "S  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 84 'T
     t.FillColor = &HFF&
     .Text = txtEventLog.Text & "T  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 85 'U
     u.FillColor = &HFF&
     .Text = txtEventLog.Text & "U  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 86 'V
     v.FillColor = &HFF&
     .Text = txtEventLog.Text & "V  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 87 'W
     w.FillColor = &HFF&
     .Text = txtEventLog.Text & "W  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 88 'X
     x.FillColor = &HFF&
     .Text = txtEventLog.Text & "X  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 89 'Y
     y.FillColor = &HFF&
     .Text = txtEventLog.Text & "Y  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 90 'Z
     z.FillColor = &HFF&
     .Text = txtEventLog.Text & "Z  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 145 'Scroll Lock
     ScrollLock.FillColor = &HFF&
     .Text = txtEventLog.Text & "Scroll Lock  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
     ShpScrollLight.FillColor = vbYellow
Case 19 'Pause
     PBreak.FillColor = &HFF&
     .Text = txtEventLog.Text & "Pause Break  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 9 'Tab
     Tabtb.FillColor = &HFF&
     .Text = txtEventLog.Text & "Tab  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
     
Case 20 'Caps Lock
     capslock.FillColor = &HFF&
     ShpCapsLight.FillColor = vbYellow
     .Text = txtEventLog.Text & "Caps Lock  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 16 'Shift
     shift_L.FillColor = &HFF&
     shift_R.FillColor = &HFF&
     .Text = txtEventLog.Text & "Shift Left  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
     .Text = txtEventLog.Text & "Shift Right  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 18 'Alt
     alt_l.FillColor = &HFF&
     Alt_r.FillColor = &HFF&
     .Text = txtEventLog.Text & "Alt Left  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
     .Text = txtEventLog.Text & "Alt Right  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 17 'Control
     ctrl_l.FillColor = &HFF&
     ctrl_r.FillColor = &HFF&
     .Text = txtEventLog.Text & "Ctrl Left  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
     .Text = txtEventLog.Text & "Ctrl Right  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 32 'Space
     space.FillColor = &HFF&
     .Text = txtEventLog.Text & "Space  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 13 'Enter
     enter.FillColor = &HFF&
     enter_num.FillColor = &HFF&
     Enter2.FillColor = &HFF&
     LblEnter.BackColor = &HFF&
     .Text = txtEventLog.Text & "Enter  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
     .Text = txtEventLog.Text & "Num Enter  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 8 'Back Space
     back.FillColor = &HFF&
     .Text = txtEventLog.Text & "Back Space  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 45 'Insert
     Insert.FillColor = &HFF&
     .Text = txtEventLog.Text & "Insert  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 36 'Home
     Home.FillColor = &HFF&
     .Text = txtEventLog.Text & "Home  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 33 'PgUp
     PgUp.FillColor = &HFF&
     .Text = txtEventLog.Text & "PgUp  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 46 'Delete
     Delete.FillColor = &HFF&
     .Text = txtEventLog.Text & "Delete  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 35 'End
      End1.FillColor = &HFF&
     .Text = txtEventLog.Text & "End  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 34 'PgDn
     PgDwn.FillColor = &HFF&
     .Text = txtEventLog.Text & "PgDn  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 38 'Up
     Up.FillColor = &HFF&
     .Text = txtEventLog.Text & "Up  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 37 'Left
     Left1.FillColor = &HFF&
     .Text = txtEventLog.Text & "Left  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 39 'Right
     Right.FillColor = &HFF&
     .Text = txtEventLog.Text & "Right  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 40 'Down
     Down.FillColor = &HFF&
     .Text = txtEventLog.Text & "Down  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 91 'Start
     start_l.FillColor = &HFF00&
     Start_r.FillColor = &HFF00&
     .Text = txtEventLog.Text & "Start Left  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
     .Text = txtEventLog.Text & "Start Right Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 189 '(-)
     min.FillColor = &HFF&
     .Text = txtEventLog.Text & "-  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 187 '(+)
     plus.FillColor = &HFF&
     .Text = txtEventLog.Text & "+  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 220 '(\)
     gmkr.FillColor = &HFF&
     .Text = txtEventLog.Text & "\  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 191 '(/)
     GmKn.FillColor = &HFF&
     .Text = txtEventLog.Text & "/  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 188 '(,)
     Comma.FillColor = &HFF&
     .Text = txtEventLog.Text & ",  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 190 '(.)
     Titik.FillColor = &HFF&
     .Text = txtEventLog.Text & ".  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 93 '(popup)
     popup.FillColor = &HFF&
     .Text = txtEventLog.Text & "PopUp  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 219 '([)
     krng_kr.FillColor = &HFF&
     .Text = txtEventLog.Text & "[  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 221 '(])
     Krng_kn.FillColor = &HFF&
     .Text = txtEventLog.Text & "]  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 186 '(;)
     ttcm.FillColor = &HFF&
     .Text = txtEventLog.Text & ";  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 222 '(')
     cmats.FillColor = &HFF&
     .Text = txtEventLog.Text & "'  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 44 'PrintScrn
     PrintScrn.FillColor = &HFF&
     .Text = txtEventLog.Text & "Print Screen  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 144 '(NumLock)
     Me.numlock.FillColor = &HFF&
     ShpNumLight.FillColor = vbYellow
     .Text = txtEventLog.Text & "Num Lock  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 111 '(/)
     Me.slash.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num /  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 106 '(*)
     Star.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num *  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 109 '(-)
     minus.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num -  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 107 '(+)
     plus2.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num +  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 110 'Num Del
     del.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num Del  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 96 'num 0
     num0.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num 0  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 97 'num 1
    num1.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num 1  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 98 'num 2
     num2.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num 2  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 99 'num 3
     num3.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num 3  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 100 'num 4
     num4.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num 4  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 101 'num 5
     num5.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num 5  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 102 'num 6
     num6.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num 6  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 103 'num 7
     num7.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num 7  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 104 'num 8
     num8.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num 8  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 105 'num 9
     num9.FillColor = &HFF&
     .Text = txtEventLog.Text & "Num 9  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
     
End Select
.SelStart = Len(.Text)
End With

End Sub

Private Sub Form_keyUp(KeyCode As Integer, Shift As Integer)
With txtEventLog
Select Case KeyCode
Case 27 'Esc
     esc.FillColor = &HFF00&
Case 112 'F1
     f1.FillColor = &HFF00&
Case 113 'F2
     f2.FillColor = &HFF00&
Case 114 'F3
     f3.FillColor = &HFF00&
Case 115 'F4
     f4.FillColor = &HFF00&
Case 116 'F5
     f5.FillColor = &HFF00&
Case 117 'F5
     f6.FillColor = &HFF00&
Case 118 'F7
     f7.FillColor = &HFF00&
Case 119 'F8
     f8.FillColor = &HFF00&
Case 120 'F9
     f9.FillColor = &HFF00&
Case 121 'F10
     f10.FillColor = &HFF00&
Case 122 'F11
     f11.FillColor = &HFF00&
Case 123 'F12
     f12.FillColor = &HFF00&
Case 192 '(`)
      nekaj.FillColor = &HFF00&
Case 49 '1
      Number1.FillColor = &HFF00&
Case 50 '2
     Number2.FillColor = &HFF00&
Case 51 '3
     Number3.FillColor = &HFF00&
Case 52 '4
     Number4.FillColor = &HFF00&
Case 53 '5
     Number5.FillColor = &HFF00&
Case 54 '6
     Number6.FillColor = &HFF00&
Case 55 '7
     Number7.FillColor = &HFF00&
Case 56 '8
     Number8.FillColor = &HFF00&
Case 57 '9
     Number9.FillColor = &HFF00&
Case 48 '0
     Number0.FillColor = &HFF00&
Case 65 'A
      a.FillColor = &HFF00&
Case 66 'B
      b.FillColor = &HFF00&
Case 67 'C
      c.FillColor = &HFF00&
Case 68, d
      d.FillColor = &HFF00&
Case 69 'E
      e.FillColor = &HFF00&
Case 70 'F
     f.FillColor = &HFF00&
Case 71 'G
     g.FillColor = &HFF00&
Case 72 'H
      h.FillColor = &HFF00&
Case 73 'I
      i.FillColor = &HFF00&
Case 74 'J
      j.FillColor = &HFF00&
Case 75 'K
     k.FillColor = &HFF00&
Case 76 'L
      l.FillColor = &HFF00&
Case 77 'M
      m.FillColor = &HFF00&
Case 78 'N
      n.FillColor = &HFF00&
Case 79 'O
      o.FillColor = &HFF00&
Case 80 'P
      p.FillColor = &HFF00&
Case 81 'Q
     q.FillColor = &HFF00&
Case 82 'R
      r.FillColor = &HFF00&
Case 83 'S
      s.FillColor = &HFF00&
Case 84 'T
     t.FillColor = &HFF00&
Case 85 'U
     u.FillColor = &HFF00&
Case 86 'V
     v.FillColor = &HFF00&
Case 87 'W
     w.FillColor = &HFF00&
Case 88 'X
     x.FillColor = &HFF00&
Case 89 'Y
     y.FillColor = &HFF00&
Case 90 'X
     z.FillColor = &HFF00&
Case 145 'Scroll Lock
     ScrollLock.FillColor = &HFF00&
     ShpScrollLight.FillColor = &H8080&
Case 19 'Pause
     PBreak.FillColor = &HFF00&
Case 9 'Tab
     Tabtb.FillColor = &HFF00&
Case 20 'Caps Lock
     capslock.FillColor = &HFF00&
     ShpCapsLight.FillColor = &H8080&
Case 16 'Shift
     Me.shift_L.FillColor = &HFF00&
     Me.shift_R.FillColor = &HFF00&
Case 18 'Alt
     alt_l.FillColor = &HFF00&
     Alt_r.FillColor = &HFF00&
Case 17 'Control
     ctrl_l.FillColor = &HFF00&
     ctrl_r.FillColor = &HFF00&
Case 32 'Space
     space.FillColor = &HFF00&
Case 13 'Enter
     enter.FillColor = &HFF00&
     enter_num.FillColor = &HFF00&
     Enter2.FillColor = &HFF00&
     LblEnter.BackColor = &HFF00&
Case 8 'Back Space
     Me.back.FillColor = &HFF00&
Case 45 'Insert
     Insert.FillColor = &HFF00&
Case 36 'Home
     Home.FillColor = &HFF00&
Case 33 'PgUp
     PgUp.FillColor = &HFF00&
Case 46 'Delete
     Delete.FillColor = &HFF00&
Case 35 'End
      End1.FillColor = &HFF00&
Case 34 'PgDn
     PgDwn.FillColor = &HFF00&
Case 38 'Up
     Up.FillColor = &HFF00&
Case 37 'Left
     Left1.FillColor = &HFF00&
Case 39 'Right
     Right.FillColor = &HFF00&
Case 40 'Down
     Down.FillColor = &HFF00&
Case 91 'Start
     start_l.FillColor = &HFF00&
     Start_r.FillColor = &HFF00&
Case 189 '(-)
     min.FillColor = &HFF00&
Case 187 '(+)
     plus.FillColor = &HFF00&
Case 220 '(\)
     gmkr.FillColor = &HFF00&
Case 191 '(/)
     GmKn.FillColor = &HFF00&
Case 188 '(,)
     Comma.FillColor = &HFF00&
Case 190 '(.)
     Titik.FillColor = &HFF00&
Case 93 '(popup)
     popup.FillColor = &HFF00&
Case 219 '([)
     krng_kr.FillColor = &HFF00&
Case 221 '(])
     Krng_kn.FillColor = &HFF00&
Case 186 '(])
     ttcm.FillColor = &HFF00&
Case 222 '(])
     cmats.FillColor = &HFF00&
Case 44 ' PrintScrn
     PrintScrn.FillColor = &HFF00&
     .Text = txtEventLog.Text & "Print Screen  Key Pressed" & " - KeyCode : " & KeyCode & vbCrLf
Case 144 '(NumLock)
     numlock.FillColor = &HFF00&
     ShpNumLight.FillColor = &H8080&
Case 111 '(/)
     slash.FillColor = &HFF00&
Case 106 '(*)
     Star.FillColor = &HFF00&
Case 109 '(-)
     minus.FillColor = &HFF00&
Case 107 '(+)
     plus2.FillColor = &HFF00&
Case 110 'Num Del
     del.FillColor = &HFF00&
Case 96 'num 0
     num0.FillColor = &HFF00&
Case 97 'num 1
    num1.FillColor = &HFF00&
Case 98 'num 2
     num2.FillColor = &HFF00&
Case 99 'num 3
     num3.FillColor = &HFF00&
Case 100 'num 4
     num4.FillColor = &HFF00&
Case 101 'num 5
     num5.FillColor = &HFF00&
Case 102 'num 6
     num6.FillColor = &HFF00&
Case 103 'num 7
     num7.FillColor = &HFF00&
Case 104 'num 8
     num8.FillColor = &HFF00&
Case 105 'num 9
     num9.FillColor = &HFF00&

End Select
.SelStart = Len(.Text)
End With
End Sub

Private Sub ChckOnTop_Click()
If ChckOnTop.Value = 1 Then
MeOnTop Me
Else
MeDown Me
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub

Private Sub Label120_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub





Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub

Private Sub LblAbout_Click()
FrmMain.Hide
FrmAbout.Show
End Sub


Private Sub LblExit_Click()
Unload Me
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub

Private Sub Timer1_Timer()
With txtEventLog
    .Text = txtEventLog.Text & ".... " & Get_KeyboardType & vbCrLf
    .Text = txtEventLog.Text & ".... " & GetKeyboardType(1) & vbCrLf
    .Text = txtEventLog.Text & ".... " & Get_KeyboardFuncKeys & vbCrLf
    .Text = txtEventLog.Text & ".... " & Get_KeyboardLayout & vbCrLf
    .Text = txtEventLog.Text & ".... " & LangIdent(Get_KeyboardLayout) & vbCrLf
    .Text = txtEventLog.Text & "Press Keys To Keyboard Testing " & vbCrLf
    .SelStart = Len(.Text)
End With
    txtFunctionKeys.Text = Get_KeyboardFuncKeys
    txtLayoutName.Text = LangIdent(Get_KeyboardLayout)
    txtLayoutID.Text = Get_KeyboardLayout
    txtSubType.Text = GetKeyboardType(1)
    txtType.Text = Get_KeyboardType
Timer1.Enabled = False
End Sub

Private Sub LblRefKeys_Click()
  'Esc
     esc.FillColor = &HE0E0E0
  'F1
     f1.FillColor = &HE0E0E0
  'F2
     f2.FillColor = &HE0E0E0
 'F3
     f3.FillColor = &HE0E0E0
  'F4
     f4.FillColor = &HE0E0E0
  'F5
     f5.FillColor = &HE0E0E0
  'F5
     f6.FillColor = &HE0E0E0
  'F7
     f7.FillColor = &HE0E0E0
  'F8
     f8.FillColor = &HE0E0E0
  'F9
     f9.FillColor = &HE0E0E0
  'F10
     f10.FillColor = &HE0E0E0
  'F11
     f11.FillColor = &HE0E0E0
  'F12
     f12.FillColor = &HE0E0E0
  '(`)
      nekaj.FillColor = &HE0E0E0
  '1
      Number1.FillColor = &HE0E0E0
 '2
     Number2.FillColor = &HE0E0E0
 '3
     Number3.FillColor = &HE0E0E0
  '4
     Number4.FillColor = &HE0E0E0
  '5
     Number5.FillColor = &HE0E0E0
  '6
     Number6.FillColor = &HE0E0E0
 '7
     Number7.FillColor = &HE0E0E0
  '8
     Number8.FillColor = &HE0E0E0
  '9
     Number9.FillColor = &HE0E0E0
  '0
     Number0.FillColor = &HE0E0E0
 'A
      a.FillColor = &HE0E0E0
  'B
      b.FillColor = &HE0E0E0
  'C
      c.FillColor = &HE0E0E0
' d
      d.FillColor = &HE0E0E0
  'E
      e.FillColor = &HE0E0E0
 'F
     f.FillColor = &HE0E0E0
 'G
     g.FillColor = &HE0E0E0
  'H
      h.FillColor = &HE0E0E0
  'I
      i.FillColor = &HE0E0E0
 'J
      j.FillColor = &HE0E0E0
 'K
     k.FillColor = &HE0E0E0
  'L
      l.FillColor = &HE0E0E0
  'M
      m.FillColor = &HE0E0E0
  'N
      n.FillColor = &HE0E0E0
  'O
      o.FillColor = &HE0E0E0
  'P
      p.FillColor = &HE0E0E0
  'Q
     q.FillColor = &HE0E0E0
  'R
      r.FillColor = &HE0E0E0
  'S
      s.FillColor = &HE0E0E0
  'T
     t.FillColor = &HE0E0E0
 'U
     u.FillColor = &HE0E0E0
  'V
     v.FillColor = &HE0E0E0
  'W
     w.FillColor = &HE0E0E0
  'X
     x.FillColor = &HE0E0E0
  'Y
     y.FillColor = &HE0E0E0
  'X
     z.FillColor = &HE0E0E0
 'Scroll Lock
     ScrollLock.FillColor = &HE0E0E0
     ShpScrollLight.FillColor = &H8080&
  'Pause
     PBreak.FillColor = &HE0E0E0
 'Tab
     Tabtb.FillColor = &HE0E0E0
  'Caps Lock
     capslock.FillColor = &HE0E0E0
     ShpCapsLight.FillColor = &H8080&
  'Shift
     Me.shift_L.FillColor = &HE0E0E0
     Me.shift_R.FillColor = &HE0E0E0
  'Alt
     alt_l.FillColor = &HE0E0E0
     Alt_r.FillColor = &HE0E0E0
  'Control
     ctrl_l.FillColor = &HE0E0E0
     ctrl_r.FillColor = &HE0E0E0
  'Space
     space.FillColor = &HE0E0E0
  'Enter
     enter.FillColor = &HE0E0E0
     enter_num.FillColor = &HE0E0E0
     Enter2.FillColor = &HE0E0E0
     LblEnter.BackColor = &HE0E0E0
  'Back Space
     Me.back.FillColor = &HE0E0E0
  'Insert
     Insert.FillColor = &HE0E0E0
  'Home
     Home.FillColor = &HE0E0E0
  'PgUp
     PgUp.FillColor = &HE0E0E0
  'Delete
     Delete.FillColor = &HE0E0E0
  'End
      End1.FillColor = &HE0E0E0
  'PgDn
     PgDwn.FillColor = &HE0E0E0
  'Up
     Up.FillColor = &HE0E0E0
  'Left
     Left1.FillColor = &HE0E0E0
  'Right
     Right.FillColor = &HE0E0E0
  'Down
     Down.FillColor = &HE0E0E0
  'Start
     start_l.FillColor = &HE0E0E0
     Start_r.FillColor = &HE0E0E0
  '(-)
     min.FillColor = &HE0E0E0
  '(+)
     plus.FillColor = &HE0E0E0
  '(\)
     gmkr.FillColor = &HE0E0E0
  '(/)
     GmKn.FillColor = &HE0E0E0
 '(,)
     Comma.FillColor = &HE0E0E0
  '(.)
     Titik.FillColor = &HE0E0E0
  '(popup)
     popup.FillColor = &HE0E0E0
  '([)
     krng_kr.FillColor = &HE0E0E0
  '(])
     Krng_kn.FillColor = &HE0E0E0
  '(])
     ttcm.FillColor = &HE0E0E0
 '(])
     cmats.FillColor = &HE0E0E0
  ' PrintScrn
     PrintScrn.FillColor = &HE0E0E0
    
 '(NumLock)
     numlock.FillColor = &HE0E0E0
     ShpNumLight.FillColor = &H8080&
  '(/)
     slash.FillColor = &HE0E0E0
  '(*)
     Star.FillColor = &HE0E0E0
  '(-)
     minus.FillColor = &HE0E0E0
  '(+)
     plus2.FillColor = &HE0E0E0
  'Num Del
     del.FillColor = &HE0E0E0
  'num 0
     num0.FillColor = &HE0E0E0
  'num 1
    num1.FillColor = &HE0E0E0
 'num 2
     num2.FillColor = &HE0E0E0
  'num 3
     num3.FillColor = &HE0E0E0
  'num 4
     num4.FillColor = &HE0E0E0
  'num 5
     num5.FillColor = &HE0E0E0
  'num 6
     num6.FillColor = &HE0E0E0
  'num 7
     num7.FillColor = &HE0E0E0
 'num 8
     num8.FillColor = &HE0E0E0
 'num 9
     num9.FillColor = &HE0E0E0
End Sub
Private Sub LblCheckAgain_Click()
'Esc
     esc.FillColor = &HE0E0E0
  'F1
     f1.FillColor = &HE0E0E0
  'F2
     f2.FillColor = &HE0E0E0
 'F3
     f3.FillColor = &HE0E0E0
  'F4
     f4.FillColor = &HE0E0E0
  'F5
     f5.FillColor = &HE0E0E0
  'F5
     f6.FillColor = &HE0E0E0
  'F7
     f7.FillColor = &HE0E0E0
  'F8
     f8.FillColor = &HE0E0E0
  'F9
     f9.FillColor = &HE0E0E0
  'F10
     f10.FillColor = &HE0E0E0
  'F11
     f11.FillColor = &HE0E0E0
  'F12
     f12.FillColor = &HE0E0E0
  '(`)
      nekaj.FillColor = &HE0E0E0
  '1
      Number1.FillColor = &HE0E0E0
 '2
     Number2.FillColor = &HE0E0E0
 '3
     Number3.FillColor = &HE0E0E0
  '4
     Number4.FillColor = &HE0E0E0
  '5
     Number5.FillColor = &HE0E0E0
  '6
     Number6.FillColor = &HE0E0E0
 '7
     Number7.FillColor = &HE0E0E0
  '8
     Number8.FillColor = &HE0E0E0
  '9
     Number9.FillColor = &HE0E0E0
  '0
     Number0.FillColor = &HE0E0E0
 'A
      a.FillColor = &HE0E0E0
  'B
      b.FillColor = &HE0E0E0
  'C
      c.FillColor = &HE0E0E0
' d
      d.FillColor = &HE0E0E0
  'E
      e.FillColor = &HE0E0E0
 'F
     f.FillColor = &HE0E0E0
 'G
     g.FillColor = &HE0E0E0
  'H
      h.FillColor = &HE0E0E0
  'I
      i.FillColor = &HE0E0E0
 'J
      j.FillColor = &HE0E0E0
 'K
     k.FillColor = &HE0E0E0
  'L
      l.FillColor = &HE0E0E0
  'M
      m.FillColor = &HE0E0E0
  'N
      n.FillColor = &HE0E0E0
  'O
      o.FillColor = &HE0E0E0
  'P
      p.FillColor = &HE0E0E0
  'Q
     q.FillColor = &HE0E0E0
  'R
      r.FillColor = &HE0E0E0
  'S
      s.FillColor = &HE0E0E0
  'T
     t.FillColor = &HE0E0E0
 'U
     u.FillColor = &HE0E0E0
  'V
     v.FillColor = &HE0E0E0
  'W
     w.FillColor = &HE0E0E0
  'X
     x.FillColor = &HE0E0E0
  'Y
     y.FillColor = &HE0E0E0
  'X
     z.FillColor = &HE0E0E0
 'Scroll Lock
     ScrollLock.FillColor = &HE0E0E0
     ShpScrollLight.FillColor = &H8080&
  'Pause
     PBreak.FillColor = &HE0E0E0
 'Tab
     Tabtb.FillColor = &HE0E0E0
  'Caps Lock
     capslock.FillColor = &HE0E0E0
     ShpCapsLight.FillColor = &H8080&
  'Shift
     Me.shift_L.FillColor = &HE0E0E0
     Me.shift_R.FillColor = &HE0E0E0
  'Alt
     alt_l.FillColor = &HE0E0E0
     Alt_r.FillColor = &HE0E0E0
  'Control
     ctrl_l.FillColor = &HE0E0E0
     ctrl_r.FillColor = &HE0E0E0
  'Space
     space.FillColor = &HE0E0E0
  'Enter
     enter.FillColor = &HE0E0E0
     enter_num.FillColor = &HE0E0E0
     Enter2.FillColor = &HE0E0E0
     LblEnter.BackColor = &HE0E0E0
  'Back Space
     Me.back.FillColor = &HE0E0E0
  'Insert
     Insert.FillColor = &HE0E0E0
  'Home
     Home.FillColor = &HE0E0E0
  'PgUp
     PgUp.FillColor = &HE0E0E0
  'Delete
     Delete.FillColor = &HE0E0E0
  'End
      End1.FillColor = &HE0E0E0
  'PgDn
     PgDwn.FillColor = &HE0E0E0
  'Up
     Up.FillColor = &HE0E0E0
  'Left
     Left1.FillColor = &HE0E0E0
  'Right
     Right.FillColor = &HE0E0E0
  'Down
     Down.FillColor = &HE0E0E0
  'Start
     start_l.FillColor = &HE0E0E0
     Start_r.FillColor = &HE0E0E0
  '(-)
     min.FillColor = &HE0E0E0
  '(+)
     plus.FillColor = &HE0E0E0
  '(\)
     gmkr.FillColor = &HE0E0E0
  '(/)
     GmKn.FillColor = &HE0E0E0
 '(,)
     Comma.FillColor = &HE0E0E0
  '(.)
     Titik.FillColor = &HE0E0E0
  '(popup)
     popup.FillColor = &HE0E0E0
  '([)
     krng_kr.FillColor = &HE0E0E0
  '(])
     Krng_kn.FillColor = &HE0E0E0
  '(])
     ttcm.FillColor = &HE0E0E0
 '(])
     cmats.FillColor = &HE0E0E0
  ' PrintScrn
     PrintScrn.FillColor = &HE0E0E0
    
 '(NumLock)
     numlock.FillColor = &HE0E0E0
     ShpNumLight.FillColor = &H8080&
  '(/)
     slash.FillColor = &HE0E0E0
  '(*)
     Star.FillColor = &HE0E0E0
  '(-)
     minus.FillColor = &HE0E0E0
  '(+)
     plus2.FillColor = &HE0E0E0
  'Num Del
     del.FillColor = &HE0E0E0
  'num 0
     num0.FillColor = &HE0E0E0
  'num 1
    num1.FillColor = &HE0E0E0
 'num 2
     num2.FillColor = &HE0E0E0
  'num 3
     num3.FillColor = &HE0E0E0
  'num 4
     num4.FillColor = &HE0E0E0
  'num 5
     num5.FillColor = &HE0E0E0
  'num 6
     num6.FillColor = &HE0E0E0
  'num 7
     num7.FillColor = &HE0E0E0
 'num 8
     num8.FillColor = &HE0E0E0
 'num 9
     num9.FillColor = &HE0E0E0
 
    txtFunctionKeys.Text = ""
    txtLayoutName.Text = ""
    txtLayoutID.Text = ""
    txtSubType.Text = ""
    txtType.Text = ""
    txtEventLog.Text = ""
    txtEventLog.Text = "Checking Keyboard Information Please wait ........"
    Timer1.Enabled = True
End Sub
