VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdchr 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Chr"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdoct 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Oct"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdbin 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bin"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Int"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   40
      Text            =   "0."
      Top             =   120
      Width           =   6975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   360
      TabIndex        =   37
      Top             =   600
      Width           =   2655
      Begin VB.OptionButton optrad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Radians"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optdeg 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Degrees"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdxtoy 
      BackColor       =   &H00C0C0C0&
      Caption         =   "x^y"
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdsquare 
      BackColor       =   &H00C0C0C0&
      Caption         =   "x^2"
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdcube 
      BackColor       =   &H00C0C0C0&
      Caption         =   "x^3"
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdHex 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hex"
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdfact 
      BackColor       =   &H00C0C0C0&
      Caption         =   "!n"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdln 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ln"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdpi 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pi"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdfix 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fix"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdsecant 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sec"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdcosec 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cosec"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdcotan 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cotan"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdexp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exp"
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdlog 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Log"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdrnd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rnd"
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdatan 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Atan"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdcos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cos"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdsin 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sin"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdtan 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tan"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdabs 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Abs"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1920
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   4320
      Top             =   1320
   End
   Begin VB.CommandButton cmdmplus 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&M+"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdms 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&MS"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdmr 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&MR"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   360
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1290
      Width           =   615
   End
   Begin VB.CommandButton cmdc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&C"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton cmdce 
      BackColor       =   &H00C0C0C0&
      Caption         =   "C&E"
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdbackspace 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Back"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmddecimal 
      BackColor       =   &H00C0C0C0&
      Caption         =   "."
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdplusminus 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+/-"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdequal 
      BackColor       =   &H00C0C0C0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdoneoverx 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1/x"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdpercent 
      BackColor       =   &H00C0C0C0&
      Caption         =   "%"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdsqrt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sqrt"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdplus 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdsubtract 
      BackColor       =   &H00C0C0C0&
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
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdmultiply 
      BackColor       =   &H00C0C0C0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmddivide 
      BackColor       =   &H00C0C0C0&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   615
   End
   Begin XPKalkulator.xpcmdbutton cmdmc 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "&MC"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPKalkulator.xpcmdbutton cmddigits 
      Height          =   495
      Index           =   7
      Left            =   1080
      TabIndex        =   45
      Top             =   1920
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3720
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XPKalkulator.xpcmdbutton cmddigits 
      Height          =   495
      Index           =   99
      Left            =   1680
      TabIndex        =   46
      Top             =   1920
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPKalkulator.xpcmdbutton cmddigits 
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   47
      Top             =   1920
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPKalkulator.xpcmdbutton cmddigits 
      Height          =   495
      Index           =   2
      Left            =   1080
      TabIndex        =   48
      Top             =   2400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPKalkulator.xpcmdbutton cmddigits 
      Height          =   495
      Index           =   3
      Left            =   1680
      TabIndex        =   49
      Top             =   2400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPKalkulator.xpcmdbutton cmddigits 
      Height          =   495
      Index           =   4
      Left            =   2280
      TabIndex        =   50
      Top             =   2400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPKalkulator.xpcmdbutton cmddigits 
      Height          =   495
      Index           =   5
      Left            =   1080
      TabIndex        =   51
      Top             =   2880
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPKalkulator.xpcmdbutton cmddigits 
      Height          =   495
      Index           =   6
      Left            =   1680
      TabIndex        =   52
      Top             =   2880
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPKalkulator.xpcmdbutton cmddigits 
      Height          =   495
      Index           =   8
      Left            =   2280
      TabIndex        =   53
      Top             =   2880
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPKalkulator.xpcmdbutton cmddigits 
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   54
      Top             =   3360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   0
      X2              =   8040
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
