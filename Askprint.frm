VERSION 5.00
Begin VB.Form Askprint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   Icon            =   "Askprint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3600
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Copies"
      Height          =   675
      Left            =   120
      TabIndex        =   5
      Top             =   750
      Width           =   3375
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Print"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print What"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton Option2 
         Caption         =   "Selected Text"
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Entire Document"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Askprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Askprint.Visible = False
DoEvents
Printing.Show
DoEvents
End Sub

Private Sub Form_Load()
With Combo1
For i = 1 To 10
.AddItem (i)
Next i
.Text = 1
End With
End Sub
