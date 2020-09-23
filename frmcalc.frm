VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "xp_canvas1"
   ClientHeight    =   3915
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmcalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   261
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   481
   StartUpPosition =   2  'CenterScreen
   Begin XPKalkulator.xp_canvas xp_canvas1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7011
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnucopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnupaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnuoptions 
         Caption         =   "&Options"
         Begin VB.Menu mnuphysical 
            Caption         =   "&Physical"
            Begin VB.Menu mnuavo 
               Caption         =   "&Avogadro constant [1/Kmol]"
            End
            Begin VB.Menu mnuboltz 
               Caption         =   "&Boltzman constant [J/K]"
            End
            Begin VB.Menu mnuelecradius 
               Caption         =   "&Electron classical radius [m]"
            End
            Begin VB.Menu mnuelecmass 
               Caption         =   "Electron &mass [Kg]"
            End
            Begin VB.Menu mnuelemcharg 
               Caption         =   "Elementary &charge [C]"
            End
            Begin VB.Menu mnugravity 
               Caption         =   "&Gravity constant [N*m^2/Kg^2]"
            End
            Begin VB.Menu mnumuon 
               Caption         =   "Mu&on mass [Kg]"
            End
            Begin VB.Menu mnuplanck 
               Caption         =   "&Planck constant [J*s]"
            End
            Begin VB.Menu mnuproton 
               Caption         =   "P&roton mass [Kg]"
            End
            Begin VB.Menu mnurydberg 
               Caption         =   "Ry&dberg constant [1/m]"
            End
            Begin VB.Menu mnulight 
               Caption         =   "Speed of &light [m/s]"
            End
            Begin VB.Menu mnuvacum 
               Caption         =   "&Vacuum permeability [N/A^2]"
            End
            Begin VB.Menu mnuvacuum1 
               Caption         =   "Vacuum permi&ttivity [F/m]"
            End
         End
         Begin VB.Menu mnusep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnubkcolor 
            Caption         =   "&Back Color"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnudisplay 
            Caption         =   "&Display Color"
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuforecolor 
            Caption         =   "&Fore Color"
            Shortcut        =   ^F
         End
         Begin VB.Menu mnusepfont 
            Caption         =   "-"
         End
         Begin VB.Menu mnusound 
            Caption         =   "&Sound"
            Shortcut        =   {F5}
         End
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetTextColor Lib "Gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "Gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Dim digit1 As Double
Dim operator As String
Dim memory As Double
Dim copy As Double
Dim timerflag As Integer
Dim flag As Integer
Dim SaveRes As String
Dim enab As Integer
Static Function log10(X As Long)

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_log10

log10 = Log(X) / Log(10#)
    

    Exit Function

err_log10:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: log10" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Function
Private Sub chkinv_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_chkinv_Click

chkinv.TabStop = False
    

    Exit Sub

err_chkinv_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: chkinv_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cm_Click()

End Sub

Private Sub cmdabs_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdabs_Click

Text2.Text = Abs(Val(Text2.Text))
    

    Exit Sub

err_cmdabs_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdabs_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdatan_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdatan_Click

If optdeg.Value = True Then
Text2.Text = Atn(Val(Text2.Text) * 0.01745)
End If
If optrad.Value = True Then
Text2.Text = Atn(Val(Text2.Text))
End If
    

    Exit Sub

err_cmdatan_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdatan_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdavg_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdavg_Click


    

    Exit Sub

err_cmdavg_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdavg_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdbackspace_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdbackspace_Click

If Text2.Text = "" Then
Exit Sub
End If
If Text2.Text = "0." Then Exit Sub

Text2.Text = Mid$(Text2.Text, 1, Len(Text2.Text) - 1)
    

    Exit Sub

err_cmdbackspace_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdbackspace_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdbin_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdbin_Click

Text2.Text = Bin(Val(Text2.Text))
    

    Exit Sub

err_cmdbin_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdbin_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdc_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdc_Click

Text2.Text = "0."
    

    Exit Sub

err_cmdc_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdc_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub
Private Sub cmdce_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdce_Click

Text2.Text = "0."
    

    Exit Sub

err_cmdce_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdce_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdchr_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdchr_Click

On Error GoTo handler
Text2.Text = Chr(Val(Text2.Text))
Exit Sub
handler:
MsgBox Err.Description
    

    Exit Sub

err_cmdchr_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdchr_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdcos_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdcos_Click

If optdeg.Value = True Then
Text2.Text = Cos(Val(Text2.Text) * 0.01745)
End If
If optrad.Value = True Then
Text2.Text = Cos(Val(Text2.Text))
End If
    

    Exit Sub

err_cmdcos_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdcos_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub
Private Sub cmdcosec_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdcosec_Click

If optdeg.Value = True Then
Text2.Text = 1 / Sin(Val(Text2.Text) * 0.01745)
End If
If optrad.Value = True Then
Text2.Text = 1 / Sin(Val(Text2.Text))
End If
    

    Exit Sub

err_cmdcosec_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdcosec_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdcotan_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdcotan_Click

If optdeg.Value = True Then
Text2.Text = 1 / Tan(Val(Text2.Text) * 0.01745)
End If
If optrad.Value = True Then
Text2.Text = 1 / Tan(Val(Text2.Text))
End If
    

    Exit Sub

err_cmdcotan_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdcotan_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdcube_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdcube_Click

Text2.Text = Val(Text2.Text) * Val(Text2.Text) * Val(Text2.Text)
    

    Exit Sub

err_cmdcube_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdcube_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmddecimal_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmddecimal_Click

If InStr(Text2.Text, ".") Then
Exit Sub
Else
Text2.Text = Text2.Text + "."
End If
    

    Exit Sub

err_cmddecimal_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmddecimal_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub
Private Sub cmddigits_Click(Index As Integer)

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmddigits_Click

If Text2.Text = "0." Or Text2.Text = "0" Then
Text2.Text = ""
End If
Text2.Text = Text2.Text + cmddigits(Index).Caption
'ts.speak cmddigits(Index).Caption
        Exit Sub
err_cmddigits_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmddigits_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmddivide_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmddivide_Click

digit1 = Val(Text2.Text)
Text2.Text = ""
operator = "/"
'ts.speak "operator" & "Divide"
    

    Exit Sub

err_cmddivide_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmddivide_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdenable_Click()

End Sub

Private Sub cmdequal_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdequal_Click

If operator = "/" Then
If Val(Text2.Text) = 0 Then
Text2.Text = ""
MsgBox "Can't divide by Zero"
Exit Sub
End If
Text2.Text = digit1 / Val(Text2.Text)
SaveRes = Text2.Text
ElseIf operator = "*" Then
Text2.Text = digit1 * Val(Text2.Text)
SaveRes = Text2.Text
ElseIf operator = "+" Then
Text2.Text = digit1 + Val(Text2.Text)
SaveRes = Text2.Text
ElseIf operator = "-" Then
Text2.Text = digit1 - Val(Text2.Text)
SaveRes = Text2.Text
ElseIf operator = "^" Then
Text2.Text = digit1 ^ Val(Text2.Text)
SaveRes = Text2.Text
'''''''''''''''''''


End If
'sp1.Speak text2.text
'ts.speak Text2.Text
    

    Exit Sub

err_cmdequal_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdequal_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdexp_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdexp_Click

If Val(Text2.Text) < 0 Then
MsgBox "Please Enter a Positive Value"
Else
Text2.Text = Exp(Val(Text2.Text))
End If
    

    Exit Sub

err_cmdexp_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdexp_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdfact_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdfact_Click

Dim temp
If Val(Text2.Text) < 0 Then
Text2.Text = "Invalid Input"
Exit Sub
End If
temp = 1
For i = 1 To Val(Text2.Text)
temp = temp * i
Next
Text2.Text = temp
    

    Exit Sub

err_cmdfact_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdfact_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdfix_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdfix_Click

Text2.Text = Fix(Val(Text2.Text))

    

    Exit Sub

err_cmdfix_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdfix_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdHex_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdHex_Click

Text2.Text = Hex(Val(Text2.Text))
    

    Exit Sub

err_cmdHex_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdHex_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdint_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdint_Click

Text2.Text = Int(Val(Text2.Text))
    

    Exit Sub

err_cmdint_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdint_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdln_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdln_Click

If Not Val(Text2.Text) < 0 Then
Text2.Text = Log(Val(Text2.Text))
Else
MsgBox "Please Enter a Positive Value"
End If
    

    Exit Sub

err_cmdln_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdln_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdlog_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdlog_Click

'If Not Val(Text2.Text) < 0 Then
'Text2.Text = Log10(Val(Text2.Text) / Log(10#))
'Else
'MsgBox "Please Enter a Positive Value"
'End If
Text2.Text = log10(Val(Text2.Text))
    

    Exit Sub

err_cmdlog_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdlog_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdmc_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdmc_Click

cmdmc.Enabled = False
cmdmplus.Enabled = False
cmdmr.Enabled = False
memory = 0
Text1.Text = ""
    

    Exit Sub

err_cmdmc_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdmc_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdmplus_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdmplus_Click

memory = memory + Val(Text2.Text)
    

    Exit Sub

err_cmdmplus_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdmplus_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdmr_Click()
Text2.Text = memory
End Sub
Private Sub cmdms_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdms_Click

If Text2.Text = "" Or Text2.Text = "0." Or Text2.Text = "0" Then
Exit Sub
End If
cmdmc.Enabled = True
cmdmplus.Enabled = True
cmdmr.Enabled = True
memory = Text2.Text
Text1.Text = "M"
    

    Exit Sub

err_cmdms_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdms_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdmultiply_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdmultiply_Click

digit1 = Val(Text2.Text)
Text2.Text = ""
operator = "*"
'ts.speak "operator" & "Multiply"
    

    Exit Sub

err_cmdmultiply_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdmultiply_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdoct_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdoct_Click

On Error GoTo ehand
Text2.Text = Oct(Val(Text2.Text))
Exit Sub
ehand:
MsgBox Err.Description
    

    Exit Sub

err_cmdoct_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdoct_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdoneoverx_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdoneoverx_Click

Text2.Text = 1 / Val(Text2.Text)
    

    Exit Sub

err_cmdoneoverx_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdoneoverx_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdpercent_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdpercent_Click

Text2.Text = (digit1 / Val(Text2.Text)) * 100
    

    Exit Sub

err_cmdpercent_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdpercent_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub
Private Sub cmdpi_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdpi_Click

Text2.Text = 3.14159265358979
    

    Exit Sub

err_cmdpi_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdpi_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub
Private Sub cmdplus_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdplus_Click

digit1 = Val(Text2.Text)
Text2.Text = ""
operator = "+"
'ts.speak "Operator" & "Plus"
    

    Exit Sub

err_cmdplus_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdplus_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdplusminus_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdplusminus_Click

Text2.Text = -Val(Text2.Text)
    

    Exit Sub

err_cmdplusminus_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdplusminus_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdrnd_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdrnd_Click

Text2.Text = Int(Rnd * 100) + 1
    

    Exit Sub

err_cmdrnd_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdrnd_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdsecant_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdsecant_Click

If optdeg.Value = True Then
Text2.Text = 1 / Cos(Val(Text2.Text) * 0.01745)
End If
If optrad.Value = True Then
Text2.Text = 1 / Cos(Val(Text2.Text))
End If
    

    Exit Sub

err_cmdsecant_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdsecant_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdsin_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdsin_Click

If optdeg.Value = True Then
Text2.Text = Sin(Val(Text2.Text) * 0.01745)
End If
If optrad.Value = True Then
Text2.Text = Sin(Val(Text2.Text))
End If
    

    Exit Sub

err_cmdsin_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdsin_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdsqrt_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdsqrt_Click

If Val(Text2.Text) < 0 Then
MsgBox "Can't take squareroot of negative value", vbOKOnly, App.Title
Exit Sub
End If
Text2.Text = Sqr(Val(Text2.Text))
    

    Exit Sub

err_cmdsqrt_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdsqrt_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub
Private Sub cmdsquare_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdsquare_Click

Text2.Text = Val(Text2.Text) * Val(Text2.Text)
    

    Exit Sub

err_cmdsquare_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdsquare_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub
Private Sub cmdsubtract_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdsubtract_Click

digit1 = Val(Text2.Text)
Text2.Text = ""
operator = "-"
'ts.speak "operator" & "subtraction"
    

    Exit Sub

err_cmdsubtract_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdsubtract_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub
Private Sub cmdtan_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdtan_Click

If optdeg.Value = True Then
Text2.Text = Tan(Val(Text2.Text) * 0.01745)
End If
If optrad.Value = True Then
Text2.Text = Tan(Val(Text2.Text))
End If
    

    Exit Sub

err_cmdtan_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdtan_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub cmdxtoy_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_cmdxtoy_Click

digit1 = Val(Text2.Text)
Text2.Text = ""
operator = "^"
    

    Exit Sub

err_cmdxtoy_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: cmdxtoy_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub



Private Sub Form_Activate()
   '//Dimension Variables to Allocate Memory

    On Error GoTo err_Form_Activate

Text2.SetFocus

    

    Exit Sub

err_Form_Activate:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: Form_Activate" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub Form_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_Form_Click

If Button = 2 Then
PopupMenu mnuview
End If
    

    Exit Sub

err_Form_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: Form_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_Form_KeyDown


'Dim AltDown
 '  AltDown = (Shift And vbAltMask) > 0

  'Alt + A = Shortcut for AddNew
  ' If AltDown And KeyCode = vbKeyB Then   ' A = Add
   '   Shell "notepad", vbMaximizedFocus
   'End If
    

    Exit Sub

err_Form_KeyDown:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: Form_KeyDown" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
   End Sub

Private Sub Form_Load()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_Form_Load

cmdmc.Enabled = False
cmdmplus.Enabled = False
cmdmr.Enabled = False
mnupaste.Enabled = False
Timer1.Enabled = False

    
mnuedit.Visible = False
mnuoptions.Visible = False
mnuabout.Visible = False
    Exit Sub

err_Form_Load:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: Form_Load" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_Form_MouseDown

If Button = 2 Then
PopupMenu mnuedit
End If
    

    Exit Sub

err_Form_MouseDown:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: Form_MouseDown" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_Form_Unload
      
If SaveRes <> "0." And savres <> "0" And SaveRes <> "" And Val(SaveRes) <> 0 Then
If MsgBox("Do you want to save your work", vbYesNo) = vbYes Then
Open App.Path & "/result.txt" For Append As #1
Print #1, "Result", Now, SaveRes
Close #1
End If
Else
Cancel = 0
End If
    

    Exit Sub

err_Form_Unload:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: Form_Unload" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnuabout_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuabout_Click

frmabout.Show
    

    Exit Sub

err_mnuabout_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuabout_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnuavo_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuavo_Click

Text2.Text = "6.022136736E+26"

    

    Exit Sub

err_mnuavo_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuavo_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnubkcolor_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnubkcolor_Click

On Error GoTo tt
cd.ShowColor
Me.BackColor = cd.Color
Frame2.BackColor = cd.Color
Frame1.BackColor = cd.Color
Exit Sub
tt:
MsgBox Err.Description
    

    Exit Sub

err_mnubkcolor_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnubkcolor_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnublink_Click()
End Sub

Private Sub mnuboltz_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuboltz_Click

Text2.Text = "1.38065812E-23"
    

    Exit Sub

err_mnuboltz_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuboltz_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnucopy_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnucopy_Click

'mnupaste.Enabled = True
'copy = Val(Text2.Text)
Clipboard.SetText (Text2.Text)
mnupaste.Enabled = True
    

    Exit Sub

err_mnucopy_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnucopy_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnudefault_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnudefault_Click

Text2.Height = 375
Text2.Font.Size = 10
    

    Exit Sub

err_mnudefault_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnudefault_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnudisplay_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnudisplay_Click

On Error GoTo tt
cd.ShowColor
Text2.BackColor = cd.Color
Exit Sub
tt:
MsgBox Err.Description
    

    Exit Sub

err_mnudisplay_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnudisplay_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnuelecmass_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuelecmass_Click

Text2.Text = "1.6021773349E-19"
    

    Exit Sub

err_mnuelecmass_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuelecmass_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnuelecradius_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuelecradius_Click

Text2.Text = "2.8179409238E-15"
    

    Exit Sub

err_mnuelecradius_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuelecradius_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnuelemcharg_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuelemcharg_Click

Text2.Text = "1.6021773349E-19"
    

    Exit Sub

err_mnuelemcharg_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuelemcharg_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnuforecolor_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuforecolor_Click

On Error GoTo tt
cd.ShowColor
Text2.ForeColor = cd.Color
Exit Sub
tt:
MsgBox Err.Description
    

    Exit Sub

err_mnuforecolor_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuforecolor_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnugravity_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnugravity_Click

Text2.Text = "6.6725985E-11"
    

    Exit Sub

err_mnugravity_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnugravity_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnulight_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnulight_Click

Text2.Text = "2.99792458E+08"
    

    Exit Sub

err_mnulight_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnulight_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnumuon_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnumuon_Click

Text2.Text = "1.883532711E-28"
    

    Exit Sub

err_mnumuon_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnumuon_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnunormal_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnunormal_Click

mnuscientific.Checked = False
mnunormal.Checked = True
    

    Exit Sub

err_mnunormal_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnunormal_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnupaste_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnupaste_Click

Text2.Text = Clipboard.GetText
    

    Exit Sub

err_mnupaste_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnupaste_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnuplanck_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuplanck_Click

Text2.Text = "6.62607554E-34"
    

    Exit Sub

err_mnuplanck_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuplanck_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub
Private Sub mnuplus8_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuplus8_Click

Text2.FontSize = Text2.FontSize + 10
Text2.Height = Text2.Height + 100
    

    Exit Sub

err_mnuplus8_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuplus8_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnuplusfour_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuplusfour_Click

If Not Text2.FontSize > 19 Then
Text2.FontSize = Text2.FontSize + 5
Else
Text2.FontSize = 14
End If
    

    Exit Sub

err_mnuplusfour_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuplusfour_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnuproton_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuproton_Click

Text2.Text = "1.67262311E-27"
    

    Exit Sub

err_mnuproton_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuproton_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnurydberg_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnurydberg_Click

Text2.Text = "1.097373153413E+07"
    

    Exit Sub

err_mnurydberg_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnurydberg_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnuscientific_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuscientific_Click

mnuscientific.Checked = True
mnunormal.Checked = False
    

    Exit Sub

err_mnuscientific_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuscientific_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnuvacum_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuvacum_Click

Text2.Text = "0.00000125663706144"
    

    Exit Sub

err_mnuvacum_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuvacum_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub mnuvacuum1_Click()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_mnuvacuum1_Click

Text2.Text = "8.854187817E-12"
    

    Exit Sub

err_mnuvacuum1_Click:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: mnuvacuum1_Click" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_Text2_KeyPress

If Text2.Text = "0." Or Text2.Text = "0" Then Text2.Text = ""
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
    

    Exit Sub

err_Text2_KeyPress:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: Text2_KeyPress" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Private Sub Timer1_Timer()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_Timer1_Timer

If flag = 0 Then
mnuedit.Caption = "&Edit"
mnuhelp.Caption = "&Help"
mnuview.Caption = "&View"
flag = 1
Else
mnuedit.Caption = ""
mnuhelp.Caption = ""
mnuview.Caption = ""
flag = 0
End If
    

    Exit Sub

err_Timer1_Timer:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: Timer1_Timer" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Public Sub disable()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_disable

For Each Control In Me.Controls
If TypeOf Control Is CommandButton Then
Control.Enabled = False
End If
Next
    

    Exit Sub

err_disable:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: disable" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub

Public Sub enable()

    '//Dimension Variables to Allocate Memory

    On Error GoTo err_enable

For Each Control In Me.Controls
If TypeOf Control Is CommandButton Then
Control.Enabled = True
End If
Next
    

    Exit Sub

err_enable:
    Screen.MousePointer = vbNormal
    MsgBox "An error has occured." & vbCrLf & vbTab & _
        "Procedure: enable" & vbCrLf & vbTab & _
        "Error Number: " & Err.Number & vbCrLf & vbTab & _
        "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName
End Sub
Private Function Bin(ByVal X As Long) As String

    '//Dimension Variables to Allocate Memory

    'On Error GoTo err_Bin

Dim temp As String

temp = ""
'start translation to binary
Do


' Check whether it is 1 bit or 0 bit
If X Mod 2 Then
      temp = "1" + temp
Else
      temp = "0" + temp
End If

X = X \ 2
'  Normal division     7/2 = 3.5
' Integer division     7\2 = 3
'

If X < 1 Then Exit Do

Loop '
Bin = temp

End Function
    

    'Exit Function

'err_Bin:
 '   Screen.MousePointer = vbNormal
  '  MsgBox "An error has occured." & vbCrLf & vbTab & _
   '     "Procedure: Bin" & vbCrLf & vbTab & _
    '    "Error Number: " & Err.Number & vbCrLf & vbTab & _
     '   "Error Description: " & Err.Description, vbCritical + vbOKOnly, App.EXEName

'End Sub
