VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Printing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing..."
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   Icon            =   "Printing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3030
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   500
   End
   Begin VB.Label Label1 
      Caption         =   "Please wait... Sending document to printer..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Printing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_activate()
On Error Resume Next
ProgressBar1.Max = Askprint.Combo1.Text
If Askprint.Option1.Value = True Then
For i = 1 To Askprint.Combo1.Text
ProgressBar1.Value = ProgressBar1.Value + 1
frmMain.Rt.SelStart = 0
frmMain.Rt.SelPrint (Printer.hDC)
Next i
Else
For i = 1 To Askprint.Combo1.Text
ProgressBar1.Value = ProgressBar1.Value + 1
frmMain.Rt.SelPrint (Printer.hDC)
Next i
End If
ProgressBar1.Value = ProgressBar1.Max
Unload Askprint
Unload Me
End Sub
