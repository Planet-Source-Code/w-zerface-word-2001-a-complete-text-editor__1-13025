VERSION 5.00
Begin VB.Form FrmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ForeColor       =   &H8000000F&
   Icon            =   "FrmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Auto Save"
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4455
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   4215
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   4
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Auto save every"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Enable Auto Save"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Disable Auto Save"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
a = Mid(CurDir, 1, 3)
a = a & "windows\"
Open a & "Word2001.dat" For Output As #1
Write #1, Option1.Value, Text1.Text, Combo1.Text
Close #1
Unload Me
End Sub

Private Sub Form_Load()
Combo1.AddItem ("Second(s)")
Combo1.AddItem ("Minute(s)")
Combo1.Text = "Minute(s)"
Text1.Text = 5
On Error GoTo hhh:
a = Mid(CurDir, 1, 3)
a = a & "windows\"
Open a & "Word2001.dat" For Input As #1
Input #1, a, b, c
Close #1
If a = True Then
Option1.Value = True
Option2.Value = False
Else
Option2.Value = True
Option1.Value = False
Text1.Text = b
If c = "Second(s)" Or c = "Minute(s)" Then Combo1.Text = c
End If
Me.SetFocus
hhh:
 UpdateAll
End Sub

Private Sub Option1_Click()
 UpdateAll
End Sub

Private Sub Option2_Click()
 UpdateAll
End Sub
Private Sub UpdateAll()
If Option1.Value = True Then
Text1.Enabled = False
Combo1.Enabled = False
Text1.BackColor = &HC0C0C0
Else
Text1.Enabled = True
Combo1.Enabled = True
Text1.BackColor = vbWhite
End If
End Sub
