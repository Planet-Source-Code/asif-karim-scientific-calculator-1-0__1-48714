VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Convert"
   ClientHeight    =   4530
   ClientLeft      =   615
   ClientTop       =   2715
   ClientWidth     =   3600
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox t1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      MaxLength       =   18
      TabIndex        =   4
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox t2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   2535
   End
   Begin VB.ComboBox c3 
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Text            =   "3"
      Top             =   3960
      Width           =   615
   End
   Begin VB.ComboBox c2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form6.frx":030A
      Left            =   840
      List            =   "Form6.frx":030C
      TabIndex        =   1
      Text            =   "Milimetre"
      Top             =   3120
      Width           =   2535
   End
   Begin VB.ComboBox c1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form6.frx":030E
      Left            =   840
      List            =   "Form6.frx":0310
      TabIndex        =   0
      Text            =   "Metre"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "     Equals"
      BeginProperty Font 
         Name            =   "WST_Swed"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Decimal Places"
      BeginProperty Font 
         Name            =   "Palette"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub c1_click()
c2.Clear
End Sub
Private Sub t1_click()
c2.Clear
End Sub
Private Sub c3_click()
Call t1_change
End Sub
Private Sub Form_Load()
With c1
.AddItem "Metre"
.AddItem "Litre"
.AddItem "Gram"
.AddItem "Newton"
.AddItem "Joule"
.AddItem "Celsius"
.AddItem "Decimal"
End With
With c3
.AddItem "1"
.AddItem "2"
.AddItem "3"
.AddItem "4"
.AddItem "5"
End With
End Sub
Private Sub c2_gotfocus()
If c1 = "Metre" Then
With c2
.AddItem "Milimetre"
.AddItem "Centimetre"
.AddItem "Kilometre"
.AddItem "Inch"
.AddItem "Foot"
.AddItem "Yard"
.AddItem "Mile"
End With
ElseIf c1 = "Litre" Then
c2.AddItem "Gallon"
ElseIf c1 = "Gram" Then
c2.AddItem "Kilogram"
ElseIf c1 = "Newton" Then
With c2
.AddItem "Dyne"
.AddItem "Poundal"
End With
ElseIf c1 = "Joule" Then
c2.AddItem "Calory"
ElseIf c1 = "Celsius" Then
With c2
.AddItem "Kelvin"
.AddItem "Fahrenheit"
End With
ElseIf c1 = "Decimal" Then
With c2
.AddItem "Hexadecimal"
.AddItem "Octal"
End With
End If
End Sub
Private Sub t1_change()
If c2 = "Milimetre" Then
t2 = Val(t1) * 1000
ElseIf c2 = "Centimetre" Then
t2 = Val(t1) * 100
ElseIf c2 = "Kilometre" Then
t2 = Val(t1) / 1000
ElseIf c2 = "Inch" Then
t2 = Val(t1) / 0.0254
ElseIf c2 = "Foot" Then
t2 = Val(t1) / 0.3
ElseIf c2 = "Yard" Then
t2 = Val(t1) * 1.094
ElseIf c2 = "Mile" Then
t2 = Val(t1) * 0.000621
ElseIf c2 = "Gallon" Then
t2 = Val(t1) * 0.22
ElseIf c2 = "Kilogram" Then
t2 = Val(t1) / 1000
ElseIf c2 = "Dyne" Then
t2 = Val(t1) * 100000
ElseIf c2 = "Poundal" Then
t2 = Val(t1) * 7.233
ElseIf c2 = "Calory" Then
t2 = Val(t1) * 0.238846
ElseIf c2 = "Kelvin" Then
t2 = Val(t1) + 274.15
ElseIf c2 = "Fahrenheit" Then
t2 = ((Val(t1) / 5) * 9) + 32
ElseIf c2 = "Hexadecimal" Then
On Error Resume Next
t2 = Hex$(t1.Text)
ElseIf c2 = "Octal" Then
On Error Resume Next
t2 = Oct$(t1.Text)
End If
If c3 = "1" Then
t2 = Format(t2, ".#")
ElseIf c3 = "2" Then
t2 = Format(t2, ".##")
ElseIf c3 = "3" Then
t2 = Format(t2, ".###")
ElseIf c3 = "4" Then
t2 = Format(t2, ".####")
ElseIf c3 = "5" Then
t2 = Format(t2, ".#####")
End If
End Sub
Private Sub c2_click()
Call t1_change
End Sub

Private Sub t2_click()
c2.Clear
End Sub
