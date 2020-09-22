VERSION 5.00
Begin VB.Form form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " About"
   ClientHeight    =   2055
   ClientLeft      =   11460
   ClientTop       =   6645
   ClientWidth     =   3165
   Icon            =   "formcv3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label3 
      Caption         =   "Version 1.00"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Shovon"
      BeginProperty Font 
         Name            =   "Americana BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Palette"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   1200
      Left            =   120
      Picture         =   "formcv3.frx":030A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1560
   End
End
Attribute VB_Name = "form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
