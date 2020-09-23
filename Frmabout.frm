VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                     About Calculator"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "Frmabout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frmabout.frx":030A
   ScaleHeight     =   3795
   ScaleWidth      =   4485
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   3000
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   "www.realminds.cjb.net"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   5
      ToolTipText     =   "Click to Open my Web with a lot of more Kool Stuff."
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   "                                 I Am Waiting for your Comments"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "By"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Muhammad Junaid Raza"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      MouseIcon       =   "Frmabout.frx":E0DA
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Ver 1.0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "XP Advance Kalkulator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer
Private Sub Form_Activate()
Bubble frmabout
End Sub
Sub Bubble(frm As Form)
   Dim a As Integer
   Dim b As Integer
   Dim c As Integer
   Dim d As Integer
   Dim e As Integer
   Dim f As Integer
   Dim w As Integer
   Dim X As Integer
   Dim Y As Integer
   Dim z As Integer
   Dim current As Double
   Call frm.Move(0, 0)
   w = frm.Height: X = frm.Width: Y = frm.Top: z = frm.Left
   a = 0: b = 0: c = w: d = X: e = Y: f = z
   Do While a < frm.Height / 15 Or b < frm.Width / 15
      a = a + 25
      b = b + 25
      e = e + 40
      f = f + 148
      If a > frm.Height / 15 Then a = a - 24
      If b > frm.Width / 15 Then b = b - 24
      Call frm.Move(f, e, d, c)
      current = Timer
      Do While Timer - current < 0.01
         DoEvents
      Loop
      Call SetWindowRgn(frm.hWnd, CreateEllipticRgn(0, 0, b, a), True)
   Loop
   current = Timer
   Do While Timer - current < 1
      DoEvents
   Loop
  End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub


Private Sub Label3_Click()
RetVal = ShellExecute(Me.hWnd, vbNullString, "mailto:razajunaid@hotmail.com?subject=" & MainTitle, vbNullString, Left$(CurDir$, 3), SW_SHOWNORMAL)
End Sub

Private Sub Label6_Click()
ShellExecute Me.hWnd, "open", "http://www.realminds.cjb.net", vbNullString, Left$(CurDir$, 3), SW_SHOWNORMAL
End Sub

Private Sub Timer1_Timer()
Label5.Caption = Mid(Label5.Caption, 2) & Left(Label5.Caption, 1)
End Sub
