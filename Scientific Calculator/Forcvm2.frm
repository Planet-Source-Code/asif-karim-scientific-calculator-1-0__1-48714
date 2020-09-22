VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Shortcuts"
   ClientHeight    =   2865
   ClientLeft      =   1035
   ClientTop       =   7770
   ClientWidth     =   2925
   Icon            =   "Forcvm2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "  Paste                        Ctrl+V"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "  Copy                         Ctrl+C"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "   Cut                           Ctrl+X    "
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   " Delete                     Backspace      Cut"
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
