VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Word2001"
   ClientHeight    =   4740
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7155
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer AutoSave 
      Left            =   5280
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5760
      Top             =   480
   End
   Begin VB.PictureBox Marline 
      BackColor       =   &H00808080&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3255
      ScaleWidth      =   15
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   15
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Margins"
      Top             =   840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   393216
      Max             =   100
      TickStyle       =   3
      TickFrequency   =   2
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Combo2"
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   2655
   End
   Begin RichTextLib.RichTextBox Rt 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0442
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   9
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   13
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageIndex      =   13
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   14
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Color"
            Object.ToolTipText     =   "Change Color"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find Text"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   4470
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6959
            Text            =   "Status"
            TextSave        =   "Status"
            Object.ToolTipText     =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "11/22/2000"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "2:24 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   4680
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   3960
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":04FC
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":060E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0726
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0838
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":094A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A5C
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B6E
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C80
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D92
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EA4
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FB6
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10C8
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11DA
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12EC
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t        Ctrl+X"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy     Ctrl+C"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste    Ctrl+V"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu Selall 
         Caption         =   "&Select All"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu Find 
         Caption         =   "&Find"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mb 
         Caption         =   "&Margin Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu ml 
         Caption         =   "&Margin Line"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu refre 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu insert 
      Caption         =   "&Insert"
      Begin VB.Menu dt 
         Caption         =   "&Date / Time"
      End
   End
   Begin VB.Menu Format 
      Caption         =   "&Format"
      Begin VB.Menu style 
         Caption         =   "&Style"
         Begin VB.Menu bold 
            Caption         =   "&Bold"
         End
         Begin VB.Menu italic 
            Caption         =   "&Italic"
         End
         Begin VB.Menu Underline 
            Caption         =   "&Underline"
         End
         Begin VB.Menu StrikeThru 
            Caption         =   "&StrikeThru"
         End
         Begin VB.Menu bullets 
            Caption         =   "&Bullets"
         End
      End
      Begin VB.Menu Alignment 
         Caption         =   "&Alignment"
         Begin VB.Menu aLeft 
            Caption         =   "&Left"
         End
         Begin VB.Menu aCenter 
            Caption         =   "&Center"
         End
         Begin VB.Menu aRight 
            Caption         =   "&Right"
         End
      End
      Begin VB.Menu script 
         Caption         =   "&Script"
         Begin VB.Menu Nos 
            Caption         =   "&No Scripting"
         End
         Begin VB.Menu sups 
            Caption         =   "S&uper Script"
         End
         Begin VB.Menu subs 
            Caption         =   "&Sub Script"
         End
      End
      Begin VB.Menu clr 
         Caption         =   "&Color"
         Begin VB.Menu black 
            Caption         =   "&Black"
            Checked         =   -1  'True
         End
         Begin VB.Menu gray 
            Caption         =   "&Gray"
         End
         Begin VB.Menu red 
            Caption         =   "&Red"
         End
         Begin VB.Menu orange 
            Caption         =   "&Orange"
         End
         Begin VB.Menu yellow 
            Caption         =   "&Yellow"
         End
         Begin VB.Menu green 
            Caption         =   "&Green"
         End
         Begin VB.Menu blue 
            Caption         =   "&Blue"
         End
         Begin VB.Menu purple 
            Caption         =   "&Purple"
         End
         Begin VB.Menu chngecolor 
            Caption         =   "&Other..."
         End
      End
      Begin VB.Menu cc 
         Caption         =   "&Change Case"
         Begin VB.Menu lc 
            Caption         =   "&Lower Case"
         End
         Begin VB.Menu uc 
            Caption         =   "&Upper Case"
         End
      End
      Begin VB.Menu font 
         Caption         =   "Set &Font"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SFile
Dim Startupcomplete
Dim mdpath

Private Sub aCenter_Click()
If aCenter.Checked = False Then
Clearali
aCenter.Checked = True
tbToolBar.Buttons(16).Value = tbrPressed
Rt.SelAlignment = 2
sbStatusBar.Panels(1).Text = "Text Centered"
Else
Clearali
aCenter.Checked = False
tbToolBar.Buttons(16).Value = tbrUnpressed
sbStatusBar.Panels(1).Text = "Text NOT Centered"
End If

End Sub

Private Sub aLeft_Click()
If aLeft.Checked = False Then
Clearali
Rt.SelAlignment = 0
aLeft.Checked = True
tbToolBar.Buttons(15).Value = tbrPressed
sbStatusBar.Panels(1).Text = "Text Set Left"
Else
Clearali
aLeft.Checked = False
tbToolBar.Buttons(15).Value = tbrUnpressed
sbStatusBar.Panels(1).Text = "Text NOT Set Left"
End If

End Sub

Private Sub aRight_Click()
If aRight.Checked = False Then
Clearali
aRight.Checked = True
tbToolBar.Buttons(17).Value = tbrPressed
Rt.SelAlignment = 1
sbStatusBar.Panels(1).Text = "Text Set Right"
Else
Clearali
aRight.Checked = False
tbToolBar.Buttons(17).Value = tbrUnpressed
sbStatusBar.Panels(1).Text = "Text NOT Set Right"
End If
End Sub

Private Sub AutoSave_Timer()
If Not frmMain.Caption = "Word 2001 - Untitled" Then
Beep
mnuFileSave_Click
End If
sbStatusBar.Panels(1).Text = "File Auto Saved: " & Time
End Sub

Private Sub black_Click()
ClearCChecks
black.Checked = True
Rt.SelColor = vbBlack
sbStatusBar.Panels(1).Text = "Color set to black"
End Sub

Private Sub blue_Click()
ClearCChecks
black.Checked = True
Rt.SelColor = vbBlue
sbStatusBar.Panels(1).Text = "Color set to blue"
End Sub

Private Sub bold_Click()
If bold.Checked = False Then
bold.Checked = True
Rt.SelBold = True
tbToolBar.Buttons(11).Value = tbrPressed
sbStatusBar.Panels(1).Text = "Style is Bold"
Else
tbToolBar.Buttons(11).Value = tbrUnpressed
bold.Checked = False
Rt.SelBold = False
sbStatusBar.Panels(1).Text = "Style is NOT Bold"
End If
End Sub

Private Sub bullets_Click()
If bullets.Checked = False Then
bullets.Checked = True
Rt.SelBullet = True
sbStatusBar.Panels(1).Text = "Bullets Enabled"
Else
bullets.Checked = False
Rt.SelBullet = False
sbStatusBar.Panels(1).Text = "Bullets Disabled"
End If
End Sub

Private Sub chngecolor_Click()
sbStatusBar.Panels(1).Text = "Select a color"
On Error Resume Next
dlgCommonDialog.Color = Rt.SelColor
dlgCommonDialog.Flags = 1
dlgCommonDialog.ShowColor
Rt.SelColor = dlgCommonDialog.Color
End Sub

Private Sub Combo1_Click()
sbStatusBar.Panels(1).Text = "Select a font"
On Error Resume Next
Rt.SelFontName = Combo1.Text
If Startupcomplete = True Then
Rt.SetFocus
End If
End Sub

Private Sub Combo2_Click()
sbStatusBar.Panels(1).Text = "Select a font size"
On Error Resume Next
Rt.SelFontSize = Combo2.Text
If Startupcomplete = True Then
Rt.SetFocus
End If
End Sub


Private Sub Combo2_KeyPress(KeyAscii As Integer)
sbStatusBar.Panels(1).Text = "Enter a size number (1 - 100000)"
On Error Resume Next
If KeyAscii = 13 Then
Beep
Rt.SelFontSize = Combo2.Text
End If
End Sub

Private Sub di_Click()
Rt.SelIndent = Rt.SelIndent - 10
Slider1.Value = Rt.SelIndent / 50
End Sub

Private Sub dt_Click()
frmDt.Show
End Sub

Private Sub Find_Click()
sbStatusBar.Panels(1).Text = "Find Text"
FrmFind.Show
End Sub

Private Sub font_Click()
sbStatusBar.Panels(1).Text = "Select a font"
On Error Resume Next
dlgCommonDialog.Flags = 1
dlgCommonDialog.ShowFont
Rt.SelFontName = dlgCommonDialog.FontName
Rt.SelFontSize = dlgCommonDialog.FontSize
Rt.SelUnderline = dlgCommonDialog.FontUnderline
Rt.SelBold = dlgCommonDialog.FontBold
Rt.SelItalic = dlgCommonDialog.FontItalic
Rt.SelStrikeThru = dlgCommonDialog.FontStrikethru
End Sub

Private Sub Form_activate()
On Error GoTo h:
frmSplash.SetFocus
h:
If Startupcomplete = False Then
LoadFonts
End If
Startupcomplete = True
sbStatusBar.Panels(1).Text = "Status"

End Sub


Private Sub Form_Load()
frmSplash.Show
DoEvents
Slider1.Value = 10
curdrive = Mid(CurDir, 1, 3)
mdpath = curdrive & "MyDocu~1\*.*"
aLeft_Click
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
sbStatusBar.Panels(1).Text = "Status"
mnuFileNew_Click
GetOptions
End Sub



Private Sub Form_Resize()
On Error Resume Next
If tbToolBar.Visible = True Then
tbtb = tbToolBar.Height
Else
tbtb = 0
End If

If frmMain.Height < 3000 Then frmMain.Height = 3000
If frmMain.Width < 6900 Then frmMain.Width = 6900
Rt.Width = frmMain.Width - 100
Combo1.Top = tbtb + 10
Combo2.Top = Combo1.Top
If sbStatusBar.Visible = True Then
vvv = 1000
Else
vvv = 775
End If
If Slider1.Visible = True Then
Rt.Top = tbtb + Combo1.Height + Slider1.Height + 100
Rt.Height = frmMain.Height - (tbtb + vvv) - Combo1.Height - Slider1.Height
Else
Rt.Top = tbtb + Combo1.Height + 100
Rt.Height = frmMain.Height - (tbtb + vvv) - Combo1.Height
End If
Slider1.Top = Combo1.Height + Combo1.Top
Slider1.Width = Rt.Width
Slider1.Left = Rt.Left
Slider1.Max = Rt.Width / 50
Slider1_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload frmSplash
Dim i As Integer
For i = Forms.Count - 1 To 1 Step -1
Unload Forms(i)
Next i
If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
    SaveSetting App.Title, "Settings", "MainLeft", Me.Left
    SaveSetting App.Title, "Settings", "MainTop", Me.Top
    SaveSetting App.Title, "Settings", "MainWidth", Me.Width
    SaveSetting App.Title, "Settings", "MainHeight", Me.Height
End If
End
End Sub

Private Sub Format_Click()
sbStatusBar.Panels(1).Text = "Format"
clr_click
alignment_click
End Sub

Private Sub gray_Click()
sbStatusBar.Panels(1).Text = "Color set to gray"
ClearCChecks
black.Checked = True
Rt.SelColor = &H808080
End Sub

Private Sub green_Click()
sbStatusBar.Panels(1).Text = "Color set to green"
ClearCChecks
black.Checked = True
Rt.SelColor = vbGreen
End Sub

Private Sub ii_Click()
Rt.SelIndent = Rt.SelIndent + 10
Slider1.Value = Rt.SelIndent / 50
End Sub

Private Sub italic_Click()
If italic.Checked = False Then
italic.Checked = True
Rt.SelItalic = True
tbToolBar.Buttons(12).Value = tbrPressed
sbStatusBar.Panels(1).Text = "Style is Italic"
Else
tbToolBar.Buttons(12).Value = tbrUnpressed
italic.Checked = False
Rt.SelItalic = False
sbStatusBar.Panels(1).Text = "Style is NOT Italic"
End If
End Sub

Private Sub lc_Click()
sbStatusBar.Panels(1).Text = "Selected text is now all lower case"
Clipboard.SetText frmMain.Rt.SelText
frmMain.Rt.SelText = LCase(Clipboard.GetText)
End Sub

Private Sub mb_Click()
If mb.Checked = True Then
Slider1.Visible = False
sbStatusBar.Panels(1).Text = "Margin bar hidden"
Form_Resize
mb.Checked = False
Else
Slider1.Visible = True
sbStatusBar.Panels(1).Text = "Margin bar shown"
Form_Resize
mb.Checked = True
End If
End Sub

Private Sub ml_Click()
If ml.Checked = True Then
ml.Checked = False
Marline.BackColor = Rt.BackColor
sbStatusBar.Panels(1).Text = "Margin line hidden"
Else
ml.Checked = True
Marline.BackColor = &H8000000F
sbStatusBar.Panels(1).Text = "Margin line shown"
End If
End Sub

Private Sub mnuFileSave_Click()
If SFile = "" Then
mnuFileSaveAs_Click
GoTo ok:
End If
On Error GoTo jjj:
Rt.SaveFile (SFile)
GoTo ok:
jjj:
asdsd = MsgBox("There was an error in saving the file!", vbExclamation + vbOKOnly, "Error!")
ok:
frmMain.Caption = "Word 2001 - " & SFile
End Sub

Private Sub Nos_Click()
frmMain.Rt.SelCharOffset = 0
End Sub

Private Sub orange_Click()
ClearCChecks
black.Checked = True
Rt.SelColor = &H80FF&
sbStatusBar.Panels(1).Text = "Color set to orange"
End Sub

Private Sub purple_Click()
ClearCChecks
black.Checked = True
Rt.SelColor = &HFF00FF
sbStatusBar.Panels(1).Text = "Color set to purple"
End Sub

Private Sub red_Click()
ClearCChecks
black.Checked = True
Rt.SelColor = vbRed
sbStatusBar.Panels(1).Text = "Color set to red"
End Sub
Private Sub refre_Click()
sbStatusBar.Panels(1).Text = "Refreshing..."
RefreshRt
End Sub

Private Sub Rt_Click()
chngesub
End Sub

Private Sub Rt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
sbStatusBar.Panels(1).Text = "Refreshing..."
RefreshRt
End If
End Sub

Private Sub Rt_KeyUp(KeyCode As Integer, Shift As Integer)
chngesub
End Sub

Private Sub Selall_Click()
sbStatusBar.Panels(1).Text = "Selected all"
Rt.SetFocus
SendKeys ("^{a}")
End Sub

Private Sub Slider1_Change()
sbStatusBar.Panels(1).Text = "Margins have been changed"
On Error Resume Next
If Slider1.Value > Slider1.Max / 1.5 Then Slider1.Value = Slider1.Max / 1.5
Marline.Visible = True
Rt.SelStart = 0
Rt.SelLength = Len(Rt.Text)
Rt.SelIndent = (Slider1.Value * 48.5)
Marline.Height = Rt.Height - 100
Marline.Left = Rt.SelIndent + 100
Marline.Top = Rt.Top + 50
Rt.SelIndent = Rt.SelIndent + 100
Rt.SelLength = 0
If Startupcomplete = True Then
Rt.SetFocus
End If
End Sub

Private Sub StrikeThru_Click()

If StrikeThru.Checked = False Then
StrikeThru.Checked = True
Rt.SelStrikeThru = True
sbStatusBar.Panels(1).Text = "Style is Strikthru"
Else
StrikeThru.Checked = False
Rt.SelStrikeThru = False
sbStatusBar.Panels(1).Text = "Style is NOT Strikthru"
End If
End Sub

Private Sub subs_Click()
frmMain.Rt.SelCharOffset = -55
End Sub

Private Sub sups_Click()
sbStatusBar.Panels(1).Text = "Superscript"
frmMain.Rt.SelCharOffset = 55
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Find"
        Find_Click
        Case "Color"
        PopupMenu clr, , Button.Left, Button.Top + Button.Height
        Case "New"
        mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Bold"
bold_Click
        Case "Italic"
italic_Click
        Case "Underline"
Underline_Click
        Case "Align Left"
aLeft_Click
        Case "Center"
aCenter_Click
        Case "Align Right"
aRight_Click
    End Select
End Sub

Private Sub mnuViewOptions_Click()
sbStatusBar.Panels(1).Text = "Options"
FrmOptions.Show
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
Form_Resize
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
    Form_Resize
End Sub

Private Sub mnuEditPaste_Click()
sbStatusBar.Panels(1).Text = "Clipboard pasted"
Rt.SetFocus
SendKeys "^{v}"
End Sub

Private Sub mnuEditCopy_Click()
sbStatusBar.Panels(1).Text = "Selected text sent to clipboard"
Rt.SetFocus
SendKeys "^{c}"
End Sub

Private Sub mnuEditCut_Click()
sbStatusBar.Panels(1).Text = "Selected text sent to clipboard"
Rt.SetFocus
SendKeys "^{x}"
End Sub

Private Sub mnuEditUndo_Click()
Rt.SetFocus
SendKeys "^{z}"
End Sub

Private Sub mnuFileExit_Click()
sbStatusBar.Panels(1).Text = "goodbye!"
Unload Me
End
End Sub

Private Sub mnuFilePrint_Click()
sbStatusBar.Panels(1).Text = "Printing..."
Askprint.Show
End Sub

Private Sub mnuFilePageSetup_Click()
    sbStatusBar.Panels(1).Text = "Page setup"
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFilePrintPreview_Click()
sbStatusBar.Panels(1).Text = "Print Preview"
    'ToDo: Add 'mnuFilePrintPreview_Click' code.
    MsgBox "Add 'mnuFilePrintPreview_Click' code."
End Sub

Private Sub mnuFileSaveAs_Click()
sbStatusBar.Panels(1).Text = "Save As..."
On Error GoTo ghj:
dlgCommonDialog.Flags = 1
dlgCommonDialog.Filter = "Word for Windows 6.0|*.doc"
dlgCommonDialog.ShowSave
If dlgCommonDialog.FileName = "" Then GoTo h:
Rt.SaveFile (dlgCommonDialog.FileName)
SFile = dlgCommonDialog.FileName
h:
frmMain.Caption = "Word 2001 - " & SFile
GoTo h2:
ghj:
nul = MsgBox("Error in saving file!", vbExclamation + vbOKOnly, "Error!")
h2:
End Sub





Private Sub mnuFileOpen_Click()
sbStatusBar.Panels(1).Text = "Open File..."
Dim SFile As String
    With dlgCommonDialog
        .DialogTitle = "Pick a File to Open"
        .CancelError = False
        .Flags = 1
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
    End With
SFile = dlgCommonDialog.FileName
Rt.LoadFile (SFile)
frmMain.Caption = "Word 2001 - " & SFile
Form_Resize
End Sub

Private Sub mnuFileNew_Click()
frmMain.Caption = "Word 2001 - Untitled"
With Rt
.Text = ""
.SelAlignment = 0
.FileName = ""
.SelFontName = "Times New Romans"
.SelBold = False
.SelBullet = False
.SelColor = vbBlack
.SelFontSize = 10
.SelItalic = False
.SelRightIndent = False
.SelStrikeThru = False
.SelUnderline = False
.SelCharOffset = 0
End With
RefreshRt
End Sub

Private Sub Timer1_Timer()
Unload frmSplash
Timer1.Enabled = False
End Sub

Private Sub uc_Click()
sbStatusBar.Panels(1).Text = "Selected text now all upper case"
Clipboard.SetText frmMain.Rt.SelText
frmMain.Rt.SelText = UCase(Clipboard.GetText)
End Sub

Private Sub Underline_Click()
If Underline.Checked = False Then
Underline.Checked = True
Rt.SelUnderline = True
tbToolBar.Buttons(13).Value = tbrPressed
sbStatusBar.Panels(1).Text = "Style is underline"
Else
tbToolBar.Buttons(13).Value = tbrUnpressed
Underline.Checked = False
Rt.SelUnderline = False
sbStatusBar.Panels(1).Text = "Style is NOT underline"
End If
End Sub
Private Sub style_click()
If Rt.SelUnderline = True Then
Underline.Checked = True
Else
Underline.Checked = False
End If
If Rt.SelItalic = True Then
italic.Checked = True
Else
italic.Checked = False
End If
If Rt.SelBold = True Then
bold.Checked = True
Else
bold.Checked = False
End If
If Rt.SelStrikeThru = True Then
StrikeThru.Checked = True
Else
StrikeThru.Checked = False
End If
If Rt.SelBullet = True Then
bullets.Checked = True
Else
bullets.Checked = False
End If
End Sub
Private Sub LoadFonts()
On Error Resume Next
Combo1.Clear
Combo2.Clear
sbStatusBar.Panels(1).Text = "Loading Fonts..."
For i = 0 To Screen.FontCount - 1
Combo1.AddItem (Screen.Fonts(i))
DoEvents
Next i
DoEvents
sbStatusBar.Panels(1).Text = "Loading Loaded."
For i = 1 To 100
Combo2.AddItem (i)
Next i

Combo2.Text = 10
Combo1.Text = "Times New Roman"
End Sub
Private Sub ClearCChecks()
black.Checked = False
gray.Checked = False
red.Checked = False
orange.Checked = False
yellow.Checked = False
green.Checked = False
blue.Checked = False
purple.Checked = False
End Sub

Private Sub yellow_Click()
ClearCChecks
black.Checked = True
Rt.SelColor = vbYellow
sbStatusBar.Panels(1).Text = "Color set to yellow"
End Sub
Private Sub clr_click()
ClearCChecks
Select Case Rt.SelColor
Case vbBlack
black.Checked = True
Case &H808080
gray.Checked = True
Case vbRed
red.Checked = True
Case &H80FF&
orange.Checked = True
Case vbYellow
yellow.Checked = True
Case vbGreen
green.Checked = True
Case vbBlue
blue.Checked = True
Case &HFF00FF
purple.Checked = True
End Select
End Sub
Private Sub Clearali()
tbToolBar.Buttons(15).Value = tbrUnpressed
tbToolBar.Buttons(16).Value = tbrUnpressed
tbToolBar.Buttons(17).Value = tbrUnpressed
aLeft.Checked = False
aCenter.Checked = False
aRight.Checked = False
End Sub
Private Sub alignment_click()
Clearali
Select Case Rt.SelAlignment
Case 0
aLeft.Checked = True
tbToolBar.Buttons(15).Value = tbrPressed
Case 1
aRight.Checked = True
tbToolBar.Buttons(17).Value = tbrPressed
Case 2
aCenter.Checked = True
tbToolBar.Buttons(16).Value = tbrPressed
End Select
End Sub
Private Sub chngesub()
Static amt
Static lamt
amt = Rt.SelAlignment
If amt <> lamt Then
alignment_click
End If
lamt = Rt.SelAlignment
If Rt.SelBold = True Then
bold.Checked = True
tbToolBar.Buttons(11).Value = tbrPressed
Else
bold.Checked = False
tbToolBar.Buttons(11).Value = tbrUnpressed
End If

If Rt.SelItalic = True Then
italic.Checked = True
tbToolBar.Buttons(12).Value = tbrPressed
Else
italic.Checked = False
tbToolBar.Buttons(12).Value = tbrUnpressed
End If

If Rt.SelUnderline = True Then
Underline.Checked = True
tbToolBar.Buttons(13).Value = tbrPressed
Else
Underline.Checked = False
tbToolBar.Buttons(13).Value = tbrUnpressed
End If


End Sub
Private Sub RefreshRt()
sbStatusBar.Panels(1).Text = "Refreshing..."
DoEvents
Format_Click
DoEvents
Combo1.Refresh
DoEvents
Combo2.Refresh
DoEvents
tbToolBar.Refresh
DoEvents
frmMain.Refresh
DoEvents
frmMain.Cls
DoEvents
Rt.Refresh
DoEvents
LoadFonts
DoEvents
sbStatusBar.Panels(1).Text = "Status"
GetOptions
End Sub
Private Sub IND(amount)
Rt.SelIndent = amount
End Sub

Private Sub GetOptions()
On Error GoTo hhh:
a = Mid(CurDir, 1, 3)
a = a & "windows\"
Open a & "Word2001.dat" For Input As #1
Input #1, a, b, c
Close #1
If a = True Then
'No autosave
AutoSave.Enabled = False
Else
AutoSave.Enabled = True
Select Case c
Case "Second(s)"
AutoSave.Interval = b * 1000
Case "Minute(s)"
AutoSave.Interval = b * 60000
Case Else
'File corrupt or something
AutoSave.Enabled = False
End Select
End If
Me.SetFocus
hhh:
End Sub
