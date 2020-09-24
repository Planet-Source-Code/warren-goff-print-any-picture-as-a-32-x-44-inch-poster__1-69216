VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl AASelectColor 
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1740
   ScaleHeight     =   885
   ScaleWidth      =   1740
   Begin MSComDlg.CommonDialog CDlgColor 
      Left            =   1230
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "AASelectColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Changed()

Private Sub picColor_DblClick()
  With CDlgColor
    .DialogTitle = "Select Color"
    .Color = picColor.BackColor
    
    .ShowColor
    
    picColor.BackColor = .Color
    RaiseEvent Changed
  End With
End Sub

Private Sub UserControl_Initialize()
  picColor.BackColor = vbWhite
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  picColor.BackColor = PropBag.ReadProperty("Color", vbWhite)
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = picColor.Width
  UserControl.Height = picColor.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Color", picColor.BackColor, vbWhite)
End Sub

Public Property Get Color() As OLE_COLOR
  Color = picColor.BackColor
End Property

Public Property Let Color(ByVal colNewColor As OLE_COLOR)
  picColor.BackColor = colNewColor
  PropertyChanged "Color"
End Property
