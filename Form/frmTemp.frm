VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTemp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Picture"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmTemp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1035
      TabIndex        =   35
      Text            =   "Combo1"
      Top             =   45
      Visible         =   0   'False
      Width           =   3720
   End
   Begin VB.TextBox txtCopies 
      Height          =   285
      Left            =   4110
      TabIndex        =   34
      Text            =   "1"
      Top             =   2490
      Width           =   915
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   315
      Left            =   4110
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2130
      Width           =   2325
   End
   Begin VB.CheckBox chkStretch 
      Caption         =   "Stretch"
      Height          =   195
      Left            =   2340
      TabIndex        =   4
      Top             =   810
      Width           =   825
   End
   Begin VB.Frame fraMargins 
      Caption         =   " Printer Margins (Pixel) "
      Height          =   1005
      Left            =   90
      TabIndex        =   23
      Top             =   2130
      Width           =   3345
      Begin VB.TextBox txtTop 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2250
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtBottom 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2250
         TabIndex        =   30
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtRight 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   630
         TabIndex        =   29
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtLeft 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   630
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblTop 
         Caption         =   "Top:"
         Height          =   195
         Left            =   1890
         TabIndex        =   27
         Top             =   270
         Width           =   315
      End
      Begin VB.Label lblRight 
         Caption         =   "Right:"
         Height          =   195
         Left            =   210
         TabIndex        =   26
         Top             =   630
         Width           =   405
      End
      Begin VB.Label lblBottom 
         Caption         =   "Bottom:"
         Height          =   165
         Left            =   1680
         TabIndex        =   25
         Top             =   630
         Width           =   525
      End
      Begin VB.Label lblLeft 
         Caption         =   "Left:"
         Height          =   165
         Left            =   300
         TabIndex        =   24
         Top             =   270
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   5550
      TabIndex        =   12
      Top             =   2820
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   345
      Left            =   4560
      TabIndex        =   11
      Top             =   2820
      Width           =   915
   End
   Begin VB.CheckBox chkShowBorder 
      Caption         =   "Show Border"
      Height          =   195
      Left            =   1020
      TabIndex        =   3
      Top             =   810
      Width           =   1215
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   4230
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1140
      Width           =   795
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   4230
      MaxLength       =   3
      TabIndex        =   6
      Top             =   810
      Width           =   795
   End
   Begin VB.ComboBox cmbPaperSize 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   2475
   End
   Begin VB.Frame fraColors 
      Caption         =   " Colors "
      Height          =   615
      Left            =   90
      TabIndex        =   17
      Top             =   1470
      Width           =   4935
      Begin VB.PictureBox AASelectColor1 
         Height          =   315
         Left            =   1110
         ScaleHeight     =   255
         ScaleWidth      =   915
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   210
         Width           =   975
      End
      Begin VB.PictureBox AASelectColor2 
         Height          =   315
         Left            =   3870
         ScaleHeight     =   255
         ScaleWidth      =   915
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   210
         Width           =   975
      End
      Begin VB.Label lblBorderColor 
         Caption         =   "Border:"
         Height          =   165
         Left            =   3270
         TabIndex        =   19
         Top             =   270
         Width           =   495
      End
      Begin VB.Label lblBGColor 
         Caption         =   "BackGround:"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbAlign 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   420
      Width           =   4005
   End
   Begin MSComDlg.CommonDialog CDlgPicture 
      Left            =   5970
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpen 
      Height          =   285
      Left            =   4770
      Picture         =   "frmTemp.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   705
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CheckBox chkPreview 
      Caption         =   "Preview"
      Height          =   195
      Left            =   5280
      TabIndex        =   13
      Top             =   1830
      Value           =   1  'Checked
      Width           =   885
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   5070
      Picture         =   "frmTemp.frx":0DFC
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   14
      Top             =   30
      Width           =   1395
      Begin Project1.AAPrintPicture AAPrintPicture1 
         Left            =   930
         Top             =   30
         _ExtentX        =   688
         _ExtentY        =   661
         Picture         =   "frmTemp.frx":19FE0
      End
   End
   Begin VB.Label lblCopies 
      Caption         =   "Copies:"
      Height          =   195
      Left            =   3540
      TabIndex        =   33
      Top             =   2520
      Width           =   525
   End
   Begin VB.Label lblPrinter 
      Caption         =   "Printer:"
      Height          =   195
      Left            =   3540
      TabIndex        =   32
      Top             =   2190
      Width           =   525
   End
   Begin VB.Label lblPaperHeight 
      Caption         =   "Height:"
      Height          =   195
      Left            =   3660
      TabIndex        =   22
      Top             =   1170
      Width           =   525
   End
   Begin VB.Label lblPaperWidth 
      Caption         =   "Width:"
      Height          =   195
      Left            =   3690
      TabIndex        =   21
      Top             =   840
      Width           =   525
   End
   Begin VB.Label lblPaperSize 
      Caption         =   "Paper Size:"
      Height          =   195
      Left            =   150
      TabIndex        =   20
      Top             =   1110
      Width           =   825
   End
   Begin VB.Label lblAlign 
      Caption         =   "Picture Align:"
      Height          =   195
      Left            =   30
      TabIndex        =   16
      Top             =   450
      Width           =   945
   End
   Begin VB.Label lblFileName 
      Caption         =   "File Name:"
      Height          =   195
      Left            =   210
      TabIndex        =   15
      Top             =   90
      Width           =   765
   End
End
Attribute VB_Name = "frmTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub GeneratePreview()
  If (chkPreview.Value = vbUnchecked) Then Exit Sub
  picPreview.Cls
  Call picPreview.PaintPicture(AAPrintPicture1.GetPreview, 0, 0, picPreview.Width, picPreview.Height)
End Sub

Private Sub AAPrintPicture1_MarginReaded(ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngRight As Long, ByVal lngBottom As Long)
  txtTop.Text = CStr(lngTop)
  txtLeft.Text = CStr(lngLeft)
  txtRight.Text = CStr(lngRight)
  txtBottom.Text = CStr(lngBottom)
End Sub

Private Sub AASelectColor1_Changed()
  'AAPrintPicture1.BackColor = AASelectColor1.Color
  Call GeneratePreview
End Sub

Private Sub AASelectColor2_Changed()
  'AAPrintPicture1.BorderColor = AASelectColor2.Color
  Call GeneratePreview
End Sub

Private Sub chkPreview_Click()
  Select Case chkPreview.Value
    Case vbChecked: Call GeneratePreview
    Case vbUnchecked:  picPreview.Cls
  End Select
End Sub

Private Sub chkShowBorder_Click()
  Select Case chkShowBorder.Value
    Case vbChecked: AAPrintPicture1.ShowBorder = True
    Case vbUnchecked: AAPrintPicture1.ShowBorder = False
  End Select
  Call GeneratePreview
End Sub

Private Sub chkStretch_Click()
  If (chkStretch.Value = vbChecked) Then
    AAPrintPicture1.Stretch = True
    cmbAlign.Enabled = False
  Else
    AAPrintPicture1.Stretch = False
    cmbAlign.Enabled = True
  End If
  Call GeneratePreview
End Sub

Private Sub cmbAlign_Click()
  AAPrintPicture1.Align = cmbAlign.ListIndex
  Call GeneratePreview
End Sub

Private Sub cmbPaperSize_Click()
  Dim intWidth As Integer, intHeight As Integer
  On Error Resume Next
  AAPrintPicture1.PaperSize = cmbPaperSize.ListIndex
  If (cmbPaperSize.ListIndex = 0) Then
    txtWidth.Enabled = True
    txtHeight.Enabled = True
    txtWidth.Text = AAPrintPicture1.PaperWidth
    txtHeight.Text = AAPrintPicture1.PaperHeight
  Else
    Call AAPrintPicture1.GetPaperWH(intWidth, intHeight)
    txtWidth.Enabled = False
    txtHeight.Enabled = False
    txtWidth.Text = CStr(intWidth)
    txtHeight.Text = CStr(intHeight)
  End If
  Call GeneratePreview
End Sub

Private Sub cmbPrinter_Click()
  AAPrintPicture1.DefultPrinter = cmbPrinter.Text
  Call GeneratePreview
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdOpen_Click()
  With CDlgPicture
    .FileName = ""
    .Filter = "All Supported Images|*.gif;*.jpg;*.bmp"
    .DialogTitle = "Open Image File For Print"
    .InitDir = App.Path
    
    .ShowOpen
    
    If (.FileName = "") Then Exit Sub
    txtFileName.Text = .FileName
    txtFileName.SelStart = Len(txtFileName.Text)
    Set AAPrintPicture1.Picture = LoadPicture(.FileName)
    Call GeneratePreview
  End With
End Sub

Private Sub cmdPrint_Click()
Dim i As Integer
  If (cmbPrinter.Text = "") Then
    MsgBox "Please select Printer", vbCritical
    Exit Sub
  End If
  DoEvents
For i = 0 To 15
    AAPrintPicture1.ShowBorder = False
    Set AAPrintPicture1.Picture = Poster.Picture1(i).Image
    Call GeneratePreview
  AAPrintPicture1.PrintPicture
Next
End Sub

Private Sub Form_Activate()
chkStretch.Value = 1
chkShowBorder.Value = 0
picPreview.Picture = Poster.Picture1(0).Image
End Sub

Private Sub Form_Load()
  Dim intCounter As Integer
  SetTopMostWindow Me.hWnd, True
  cmbAlign.AddItem "Top - Left"
  cmbAlign.AddItem "Top - Center"
  cmbAlign.AddItem "Top - Right"
  cmbAlign.AddItem "Middle - Left"
  cmbAlign.AddItem "Middle - Center"
  cmbAlign.AddItem "Middle - Right"
  cmbAlign.AddItem "Bottom - Left"
  cmbAlign.AddItem "Bottom - Center"
  cmbAlign.AddItem "Bottom - Right"
  cmbAlign.ListIndex = 4
  
  cmbPaperSize.AddItem "Manual ->"
  cmbPaperSize.AddItem "A5"
  cmbPaperSize.AddItem "A4"
  cmbPaperSize.AddItem "A3"
  cmbPaperSize.ListIndex = 2
  
  For intCounter = 0 To Printers.Count - 1
    cmbPrinter.AddItem Printers(intCounter).DeviceName
    If (Printers(intCounter).DeviceName = AAPrintPicture1.DefultPrinter) Then cmbPrinter.ListIndex = intCounter
  Next intCounter
  If (cmbPrinter.ListCount = 0) Then cmdPrint.Enabled = False
  Call GeneratePreview
End Sub

Private Sub txtCopies_Change()
  If (Val(txtCopies.Text) > 0) Then AAPrintPicture1.Copies = Val(txtCopies.Text)
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    AAPrintPicture1.PaperHeight = Val(txtHeight.Text)
    Call GeneratePreview
  End If
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    AAPrintPicture1.PaperWidth = Val(txtWidth.Text)
    Call GeneratePreview
  End If
End Sub
