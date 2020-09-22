VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculator by A s i F"
   ClientHeight    =   6240
   ClientLeft      =   6255
   ClientTop       =   3180
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command25 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   35
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command24 
      Caption         =   "MC"
      Height          =   615
      Left            =   4560
      TabIndex        =   34
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command23 
      Caption         =   "M"
      Height          =   615
      Left            =   4560
      TabIndex        =   33
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command22 
      Caption         =   "M+"
      Height          =   615
      Left            =   4560
      TabIndex        =   32
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Sqrt"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   31
      ToolTipText     =   "Square Root"
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox t3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3720
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command20 
      Caption         =   "n!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   29
      ToolTipText     =   "Factorial"
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Exp"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   28
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   27
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   26
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Tan"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   25
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Sin"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   24
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   23
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   22
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "BrushScript BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   21
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox t1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      MaxLength       =   35
      TabIndex        =   20
      Top             =   480
      Width           =   4935
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Clear  All"
      BeginProperty Font 
         Name            =   "MandarinD"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   19
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   18
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Commandj 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   23
      Left            =   3120
      TabIndex        =   17
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Commandh 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   2400
      TabIndex        =   16
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Cube"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   960
      TabIndex        =   15
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Sqr"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Commandk 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   3840
      TabIndex        =   13
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   1680
      TabIndex        =   12
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Commandq 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   3120
      TabIndex        =   11
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Commandn 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   2400
      TabIndex        =   10
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Commandb 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Commandv 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   960
      TabIndex        =   8
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Commandx 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "pi"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   960
      TabIndex        =   6
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   3120
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2400
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox t 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      MaxLength       =   35
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
   Begin VB.Menu f 
      Caption         =   "   &File"
      Begin VB.Menu d 
         Caption         =   "&Demonstrate Codes"
      End
      Begin VB.Menu e 
         Caption         =   "           &Exit"
      End
   End
   Begin VB.Menu v 
      Caption         =   "&View"
      Begin VB.Menu g 
         Caption         =   "&General"
      End
      Begin VB.Menu s 
         Caption         =   "&Scientific"
      End
   End
   Begin VB.Menu n 
      Caption         =   "&Sundry"
      Begin VB.Menu c 
         Caption         =   "&Convert"
      End
      Begin VB.Menu m 
         Caption         =   "&Date and Time"
      End
   End
   Begin VB.Menu h 
      Caption         =   "&Help"
      Begin VB.Menu k 
         Caption         =   "&Keyboard Shotcuts"
      End
      Begin VB.Menu a 
         Caption         =   "&About Calculator"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Byte
Dim p As Double
Private Sub X()
MsgBox "Scratch a Number First", vbInformation + vbOKOnly, "Erroneous Attempt"
End Sub
Private Sub a_Click()
form3.Show
End Sub
Private Sub c_Click()
Form6.Show
End Sub
Private Sub Command1_Click(Index As Integer)
If t.Text = t3.Text Then
t = ""
t = t + "0"
Else
t = t + "0"
End If
End Sub
Private Sub Command10_Click()
On Error Resume Next
Command13.Enabled = 1
Command14.Enabled = 1
If r = 4 Then
t = Val(t1) / Val(t)
ElseIf r = 2 Then
t = Val(t1) - Val(t)
ElseIf r = 3 Then
t = Val(t) * Val(t1)
ElseIf r = 1 Then
t = Val(t) + Val(t1)
ElseIf r = 5 Then
t = Val(t1) / Val(t) * 100
t = t + "%"
End If
t3.Text = t.Text
End Sub
Private Sub Command11_Click()
Command13.Enabled = 1
Command14.Enabled = 1
t = ""
t1 = ""
t.SetFocus
End Sub
Private Sub Command12_Click()
r = 5
Command13.Enabled = 1
Command14.Enabled = 1
t1.Enabled = 1
t1.Text = t.Text
t = ""
t.SetFocus
End Sub
Private Sub Command13_Click()
t = t + "."
Command13.Enabled = 0
End Sub
Private Sub Command14_Click()
t = "-" + t
Command14.Enabled = 0
End Sub
Private Sub Command15_Click()
If t.Text = t3.Text Then t = ""
Command13.Enabled = 1
Command14.Enabled = 1
If t = "" Then
Call X
Else
t = Sin(t)
End If
t.SetFocus
t3.Text = t.Text
End Sub
Private Sub Command16_Click()
If t.Text = t3.Text Then t = ""
Command13.Enabled = 1
Command14.Enabled = 1
If t = "" Then
Call X
Else
t = Tan(t)
End If
t.SetFocus
t3.Text = t.Text
End Sub
Private Sub Command17_Click()
If t.Text = t3.Text Then t = ""
Command13.Enabled = 1
Command14.Enabled = 1
If t = "" Then
Call X
Else
t = Cos(t)
End If
t.SetFocus
t3.Text = t.Text
End Sub
Private Sub Command18_Click()
If t.Text = t3.Text Then t = ""
On Error GoTo 55
Command13.Enabled = 1
Command14.Enabled = 1
If t = "" Then
Call X
Else
t = Log(t)
End If
t3.Text = t.Text
t.SetFocus: Exit Sub
55 MsgBox "Logarithm of a Negative Value is Impossible.", vbInformation + vbOKOnly, "Error"
t.SetFocus
t = ""
End Sub
Private Sub Command19_Click()
If t.Text = t3.Text Then t = ""
Command13.Enabled = 1
Command14.Enabled = 1
If t = "" Then
Call X
Else
t = Exp(t)
End If
t.SetFocus
t3.Text = t.Text
End Sub
Private Sub Command2_Click(Index As Integer)
If t.Text = t3.Text Then
t = ""
t = t + "1"
Else
t = t + "1"
End If
End Sub
Private Sub Command20_Click()
If t.Text = t3.Text Then t = ""
Command13.Enabled = 1
Command14.Enabled = 1
If t = "" Then Call X: GoTo 1
t = Round(t)
1 For i = 1 To Val(t) - 1
t = i * t
Next
t.SetFocus
t3.Text = t.Text
End Sub
Private Sub Command21_Click()
If t.Text = t3.Text Then t = ""
Command13.Enabled = 1
Command14.Enabled = 1
If t = "" Then
Call X
Else
t = Sqr(t)
End If
t.SetFocus
t3.Text = t.Text
End Sub
Private Sub Command22_Click()
p = Val(t)
t3.Text = t.Text
Command22.Enabled = False
End Sub
Private Sub Command23_Click()
Command22.Enabled = 1
t = p
t3.Text = t.Text
End Sub
Private Sub Command24_Click()
p = 0
Command22.Enabled = 1
End Sub
Private Sub Command25_Click()
t.SetFocus
SendKeys "{end}"
SendKeys "{backspace}"
End Sub
Private Sub Command3_Click(Index As Integer)
If t.Text = t3.Text Then
t = ""
t = t + "2"
Else
t = t + "2"
End If
End Sub
Private Sub Command4_Click(Index As Integer)
If t.Text = t3.Text Then
t = ""
t = t + "3"
Else
t = t + "3"
End If
End Sub
Private Sub Command5_Click(Index As Integer)
If t.Text = t3.Text Then
t = ""
t = t + "4"
Else
t = t + "4"
End If
End Sub
Private Sub Command6_Click(Index As Integer)
t = 3.14159265358979
End Sub
Private Sub Command7_Click(Index As Integer)
If t = "" Then
Call X
t.SetFocus
Else: GoTo 1
Exit Sub
1 t = t * t
End If
t3.Text = t.Text
End Sub
Private Sub Command8_Click(Index As Integer)
If t = "" Then
Call X
t.SetFocus
Else: GoTo 1
Exit Sub
1 t = t * t * t
End If
t3.Text = t.Text
End Sub
Private Sub Command9_Click(Index As Integer)
r = 1
Command13.Enabled = 1
Command14.Enabled = 1
t1.Text = t.Text
t = ""
t.SetFocus
End Sub
Private Sub Commandh_Click(Index As Integer)
r = 2
Command13.Enabled = 1
Command14.Enabled = 1
t1.Text = t.Text
t = ""
t.SetFocus
End Sub
Private Sub Commandj_Click(Index As Integer)
r = 3
Command13.Enabled = 1
Command14.Enabled = 1
t1.Text = t.Text
t = ""
t.SetFocus
End Sub
Private Sub Commandk_Click(Index As Integer)
r = 4
Command13.Enabled = 1
Command14.Enabled = 1
t1.Text = t.Text
t = ""
t.SetFocus
End Sub
Private Sub Commandx_Click(Index As Integer)
If t.Text = t3.Text Then
t = ""
t = t + "5"
Else
t = t + "5"
End If
End Sub
Private Sub Commandv_Click(Index As Integer)
If t.Text = t3.Text Then
t = ""
t = t + "6"
Else
t = t + "6"
End If
End Sub
Private Sub Commandb_Click(Index As Integer)
If t.Text = t3.Text Then
t = ""
t = t + "7"
Else
t = t + "7"
End If
End Sub
Private Sub Commandn_Click(Index As Integer)
If t.Text = t3.Text Then
t = ""
t = t + "8"
Else
t = t + "8"
End If
End Sub
Private Sub Commandq_Click(Index As Integer)
If t.Text = t3.Text Then
t = ""
t = t + "9"
Else
t = t + "9"
End If
End Sub
Private Sub d_Click()
Form4.Show
End Sub
Private Sub e_Click()
Form2.Hide
form3.Hide
Form4.Hide
Form5.Hide
End
End Sub
Private Sub Form_Load()
s.Checked = 1
t1.Visible = 0
Form1.Height = 6900
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form2.Hide
form3.Hide
Form4.Hide
Form5.Hide
Call e_Click
End Sub
Private Sub g_Click()
Form1.Height = 5865
g.Checked = 1
s.Checked = 0
End Sub
Private Sub k_Click()
Form2.Show
End Sub
Private Sub m_Click()
Form5.Show
End Sub
Private Sub s_Click()
Form1.Height = 6900
s.Checked = True
g.Checked = 0
End Sub
Private Sub t_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 126) Then
KeyAscii = 0
End If
End Sub
