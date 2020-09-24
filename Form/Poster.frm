VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Poster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Poster Child"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9705
   Icon            =   "Poster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7980
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   14076
      _Version        =   393216
      Tabs            =   18
      TabsPerRow      =   9
      TabHeight       =   529
      TabCaption(0)   =   "Original"
      TabPicture(0)   =   "Poster.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Picture2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmDlg"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "1"
      TabPicture(1)   =   "Poster.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "2"
      TabPicture(2)   =   "Poster.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1(1)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "3"
      TabPicture(3)   =   "Poster.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture1(2)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "4"
      TabPicture(4)   =   "Poster.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture1(3)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "5"
      TabPicture(5)   =   "Poster.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Picture1(4)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "6"
      TabPicture(6)   =   "Poster.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Picture1(5)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "7"
      TabPicture(7)   =   "Poster.frx":098E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Picture1(6)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "8"
      TabPicture(8)   =   "Poster.frx":09AA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Picture1(7)"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "9"
      TabPicture(9)   =   "Poster.frx":09C6
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Picture1(8)"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "10"
      TabPicture(10)  =   "Poster.frx":09E2
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Picture1(9)"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "11"
      TabPicture(11)  =   "Poster.frx":09FE
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Picture1(10)"
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "12"
      TabPicture(12)  =   "Poster.frx":0A1A
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Picture1(11)"
      Tab(12).ControlCount=   1
      TabCaption(13)  =   "13"
      TabPicture(13)  =   "Poster.frx":0A36
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "Picture1(12)"
      Tab(13).ControlCount=   1
      TabCaption(14)  =   "14"
      TabPicture(14)  =   "Poster.frx":0A52
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "Picture1(13)"
      Tab(14).ControlCount=   1
      TabCaption(15)  =   "15"
      TabPicture(15)  =   "Poster.frx":0A6E
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "Picture1(14)"
      Tab(15).ControlCount=   1
      TabCaption(16)  =   "17"
      TabPicture(16)  =   "Poster.frx":0A8A
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "Picture1(15)"
      Tab(16).ControlCount=   1
      TabCaption(17)  =   "Final"
      TabPicture(17)  =   "Poster.frx":0AA6
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "Picture3"
      Tab(17).ControlCount=   1
      Begin Project1.cmdopen CmDlg 
         Left            =   6885
         Top             =   7455
         _ExtentX        =   661
         _ExtentY        =   635
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000E&
         Caption         =   "Print Poster"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1080
         Width           =   2115
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000012&
         Height          =   345
         Left            =   2130
         TabIndex        =   19
         Top             =   705
         Width           =   1155
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   ". . ."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   45
            TabIndex        =   20
            Top             =   -195
            Width           =   1065
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   28800
         Left            =   -75015
         ScaleHeight     =   1918
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   2558
         TabIndex        =   18
         Top             =   645
         Width           =   38400
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5655
         Left            =   2025
         Picture         =   "Poster.frx":0AC2
         ScaleHeight     =   375
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   379
         TabIndex        =   17
         Top             =   1725
         Width           =   5715
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   15
         Left            =   -75000
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   16
         Top             =   615
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   14
         Left            =   -74985
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   15
         Top             =   630
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   13
         Left            =   -74985
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   14
         Top             =   600
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   12
         Left            =   -75105
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   13
         Top             =   585
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   11
         Left            =   -75075
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   12
         Top             =   615
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   10
         Left            =   -75060
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   11
         Top             =   600
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   9
         Left            =   -75030
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   10
         Top             =   615
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   8
         Left            =   -75105
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   9
         Top             =   630
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   7
         Left            =   -75015
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   8
         Top             =   615
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   6
         Left            =   -75015
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   7
         Top             =   600
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   5
         Left            =   -75030
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   6
         Top             =   600
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   4
         Left            =   -75000
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   5
         Top             =   600
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   3
         Left            =   -75060
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   4
         Top             =   600
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   2
         Left            =   -75015
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   3
         Top             =   630
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   1
         Left            =   -75135
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   2
         Top             =   615
         Width           =   9600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   0
         Left            =   -75030
         ScaleHeight     =   478
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   638
         TabIndex        =   1
         Top             =   570
         Width           =   9600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Picture               and Process Poster"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   75
         TabIndex        =   21
         Top             =   660
         Width           =   6015
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   7335
         Left            =   0
         Top             =   600
         Width           =   9555
      End
   End
End
Attribute VB_Name = "Poster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load frmTemp
frmTemp.Show
End Sub

Private Sub Form_Load()
ProcessPicture
End Sub
Sub ProcessPicture()
    Picture3.PaintPicture Picture2.Image, _
                0, _
                0, _
                2560, _
                1920, _
                0, _
                0, _
                Picture2.ScaleWidth, _
                Picture2.ScaleHeight
    Picture1(0).Refresh
'***********************************
WWidth = Int(Picture2.ScaleWidth / 4)
HHeight = Int(Picture2.ScaleHeight / 4)
    Picture1(0).PaintPicture Picture3.Image, _
                0, _
                0, _
                640, _
                480, _
                0, _
                0, _
                640, _
                480
    Picture1(0).Refresh
    Picture1(1).PaintPicture Picture3.Image, _
                0, _
                0, _
                640, _
                480, _
                640, _
                0, _
                640, _
                480
    Picture1(1).Refresh
    Picture1(2).PaintPicture Picture3.Image, _
            0, _
            0, _
            640, _
            480, _
            1280, _
            0, _
            640, _
            480
    Picture1(2).Refresh
    Picture1(3).PaintPicture Picture3.Image, _
            0, _
            0, _
            640, _
            480, _
            1920, _
            0, _
            640, _
            480
    Picture1(3).Refresh
'*************************
    Picture1(4).PaintPicture Picture3.Image, _
                0, _
                0, _
                640, _
                480, _
                0, _
                480, _
                640, _
                480
    Picture1(4).Refresh
    Picture1(5).PaintPicture Picture3.Image, _
                0, _
                0, _
                640, _
                480, _
                640, _
                480, _
                640, _
                480
    Picture1(5).Refresh
    Picture1(6).PaintPicture Picture3.Image, _
            0, _
            0, _
            640, _
            480, _
            1280, _
            480, _
            640, _
            480
    Picture1(6).Refresh
    Picture1(7).PaintPicture Picture3.Image, _
            0, _
            0, _
            640, _
            480, _
            1920, _
            480, _
            640, _
            480
    Picture1(7).Refresh

'*************************
    Picture1(8).PaintPicture Picture3.Image, _
                0, _
                0, _
                640, _
                480, _
                0, _
                960, _
                640, _
                480
    Picture1(8).Refresh
    Picture1(9).PaintPicture Picture3.Image, _
                0, _
                0, _
                640, _
                480, _
                640, _
                960, _
                640, _
                480
    Picture1(9).Refresh
    Picture1(10).PaintPicture Picture3.Image, _
            0, _
            0, _
            640, _
            480, _
            1280, _
            960, _
            640, _
            480
    Picture1(10).Refresh
    Picture1(11).PaintPicture Picture3.Image, _
            0, _
            0, _
            640, _
            480, _
            1920, _
            960, _
            640, _
            480
    Picture1(11).Refresh

'*************************
    Picture1(12).PaintPicture Picture3.Image, _
                0, _
                0, _
                640, _
                480, _
                0, _
                1440, _
                640, _
                480
    Picture1(12).Refresh
    Picture1(13).PaintPicture Picture3.Image, _
                0, _
                0, _
                640, _
                480, _
                640, _
                1440, _
                640, _
                480
    Picture1(13).Refresh
    Picture1(14).PaintPicture Picture3.Image, _
            0, _
            0, _
            640, _
            480, _
            1280, _
            1440, _
            640, _
            480
    Picture1(14).Refresh
    Picture1(15).PaintPicture Picture3.Image, _
            0, _
            0, _
            640, _
            480, _
            1920, _
            1440, _
            640, _
            480
    Picture1(15).Refresh


End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Poster = Nothing
Unload frmTemp
Set frmTemp = Nothing
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    With Poster
        .CmDlg.InitialDir = File1.Path
        .CmDlg.CancelError = True 'Set cancel error to true
        .CmDlg.MultiSelect = True   'True 'Allow multi select
        .CmDlg.DialogTitle = "Open All" 'Set dialog title
        .CmDlg.Filter = "All Graphic Files" & Chr$(0) & _
        "*.gif;*.jpg;*.bmp;*.wmf" _
        & Chr$(0) & _
        "Gif Files (*.gif)" & Chr$(0) & "*.gif" & Chr$(0) & "Jpeg Files (*.jpg)" & Chr$(0) & "*.jpg" _
        & Chr$(0) & "BMP Files (*.bmp)" & Chr$(0) _
        & "*.bmp" & Chr$(0) & "Meta Files (*.wmf)" & Chr$(0) & "*.wmf"
        
        .CmDlg.FilterIndex = 1 'Set filter index
        .CmDlg.ShowOpen
    End With
Picture2.Picture = LoadPicture(CmDlg.cFileName(1))
ProcessPicture
End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Picture1(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Picture2.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Picture3.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
        frmTemp.picPreview.Picture = Picture1(SSTab1.Tab - 1).Image
End Sub

