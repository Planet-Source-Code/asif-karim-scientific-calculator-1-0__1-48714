VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Date & Time"
   ClientHeight    =   3120
   ClientLeft      =   10755
   ClientTop       =   2580
   ClientWidth     =   4395
   Icon            =   "Formdd5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   Begin VB.Label l3 
      BeginProperty Font 
         Name            =   "WST_Swed"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "ZappedChancellor"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "ZappedChancellor"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label l1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   1
      Top             =   2160
      Width           =   105
   End
   Begin VB.Label l2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4200
      TabIndex        =   0
      Top             =   600
      Width           =   105
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
l1 = Date
l2 = Time
l3 = Format(Now, "dddd")
End Sub

