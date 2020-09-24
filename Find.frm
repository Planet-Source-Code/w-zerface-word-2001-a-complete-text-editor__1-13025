VERSION 5.00
Begin VB.Form FrmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "Find.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5235
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Find"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Match Case"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Replace with:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter string to find:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "FrmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Dim opt
Const conSwpShowWindow = &H40

Private Sub Check1_Click()
If Check1.Value = 1 Then
opt = 4
Else
opt = 0
End If
End Sub

Private Sub Command1_Click()
frmMain.Rt.SetFocus
If Command1.Caption = "&Find" Then
frmMain.Rt.SelStart = 0
nul = frmMain.Rt.Find(Text1.Text, , , opt)
Command1.Caption = "&Find Next"
Else
 frmMain.Rt.SelStart = frmMain.Rt.SelStart + Len(frmMain.Rt.SelText)
nul = frmMain.Rt.Find(Text1.Text, , Len(frmMain.Rt.Text), opt)
End If
If nul = -1 Then
nul = MsgBox("Word 2001 has finished searching the document!", vbInformation + vbOKOnly, "Finished!")
Unload Me
End If
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Not frmMain.Rt.SelText = "" Then
frmMain.Rt.SelText = Text2.Text
Command1_Click
Else
nul = MsgBox("You have not found anything to replace!", vbOKOnly + vbExclamation, "Error!")
End If
End Sub

Private Sub Form_Load()
Me.Hide
SetWindowPos hWnd, conHwndTopmost, 4000 / 15, 4000 / 15, FrmFind.Width / 15, FrmFind.Height / 15, conSwpNoActivate Or conSwpShowWindow
Me.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
End Sub
