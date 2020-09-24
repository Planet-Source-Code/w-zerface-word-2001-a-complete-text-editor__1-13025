VERSION 5.00
Begin VB.Form frmDt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Date / Time"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   Icon            =   "InsertDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formats"
      Height          =   2775
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmDt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmMain.Rt.SelText = List1.Text
Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub list1_DblClick()
Command1_Click
End Sub
Private Sub Form_Load()
List1.AddItem Format(Now, "long time")
List1.AddItem Format(Now, "short time")
List1.AddItem Format(Now, "medium time")
List1.AddItem Format(Now, "general date")
List1.AddItem Format(Now, "long date")
List1.AddItem Format(Now, "medium date")
List1.AddItem Format(Now, "short date")
List1.AddItem (Date)
List1.AddItem Format(Date, "dd - mm - yyyy")
List1.AddItem Format(Date, "dd-mm-yy")
List1.AddItem Format(Date, "dd/mm/yy")
List1.AddItem Format(Date, "dd/mm/yyyy")
List1.AddItem Format(Date, "dd/mm")
List1.AddItem Format(Date, "dd")
List1.AddItem Format(Time, "hh-mm-ss")
List1.AddItem Format(Time, "hh.mm.ss")
List1.AddItem Format(Time, "hh-mm")
End Sub

