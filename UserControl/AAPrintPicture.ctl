VERSION 5.00
Begin VB.UserControl AAPrintPicture 
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   99
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   ToolboxBitmap   =   "AAPrintPicture.ctx":0000
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   720
      Picture         =   "AAPrintPicture.ctx":0532
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   2
      Top             =   750
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picPrint 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   1740
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   1
      Top             =   120
      Width           =   1245
   End
   Begin VB.PictureBox picPrinter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "AAPrintPicture.ctx":0B9C
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   0
      Top             =   0
      Width           =   390
   End
End
Attribute VB_Name = "AAPrintPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'             Public Events
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Event MarginReaded(ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngRight As Long, ByVal lngBottom As Long)


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'             Public Enums
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Enum PrintingAlign
  Top_Left = 0
  Top_Center = 1
  Top_Right = 2
  Middle_Left = 3
  Middle_Center = 4
  Middle_Right = 5
  Bottom_Left = 6
  Bottom_Center = 7
  Bottom_Right = 8
End Enum
Public Enum PaperSizeType
  Manual_Paper = 0
  A5_Paper = 1
  A4_Paper = 2
  A3_Paper = 3
End Enum

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'          Private Variables
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private alnPicture As PrintingAlign
Private lngBackColor As OLE_COLOR
Private lngBorderColor As OLE_COLOR
Private blnShowBorder As Boolean
Private strDefultPrinter As String
Private prpPrint As PaperSizeType
Private intPaperWidth As Integer
Private intPaperHeight As Integer
Private blnStretch As Boolean
Private intCopies As Integer

Private lngMarginTop As Long
Private lngMarginLeft As Long
Private lngMarginRight As Long
Private lngMarginBottom As Long

Private Const PHYSICALWIDTH = 110
Private Const PHYSICALHEIGHT = 111
Private Const PHYSICALOFFSETX = 112
Private Const PHYSICALOFFSETY = 113
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'     UserControl Private Functions
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub GetPrinterMargin()
  Dim lngWidth As Long, lngHeight As Long
  Dim prnTemp As Printer
  
  For Each prnTemp In Printers
   If prnTemp.DeviceName = strDefultPrinter Then
      Set Printer = prnTemp
      Exit For
    End If
  Next
  Printer.ScaleMode = vbPixels
  Select Case prpPrint
    Case Manual_Paper:
      Printer.PaperSize = vbPRPSUser
      Printer.Width = intPaperWidth
      Printer.Height = intPaperHeight
    Case A5_Paper: Printer.PaperSize = vbPRPSA5
    Case A4_Paper: Printer.PaperSize = vbPRPSA4
    Case A3_Paper: Printer.PaperSize = vbPRPSA3
  End Select
  lngWidth = GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
  lngHeight = GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
  lngMarginTop = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY)
  lngMarginBottom = lngHeight - (lngMarginTop + Printer.ScaleHeight)
  lngMarginLeft = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX)
  lngMarginRight = lngWidth - (lngMarginLeft + Printer.ScaleWidth)
  RaiseEvent MarginReaded(lngMarginLeft, lngMarginTop, lngMarginRight, lngMarginBottom)
End Sub

Private Sub GeneratePrintPicture()
  Dim intWidth As Integer, intHeight As Integer
  Dim intTop As Integer, intLeft As Integer
  
  Call GetPrinterMargin
  Call GetPaperWH(intWidth, intHeight)
  picPrint.Width = UserControl.ScaleX(intWidth, vbMillimeters, vbPixels) - ((lngMarginRight + lngMarginLeft) / 4)
  picPrint.Height = UserControl.ScaleX(intHeight, vbMillimeters, vbPixels) - ((lngMarginBottom + lngMarginTop) / 4)
  Select Case alnPicture
    Case Top_Left:
      intTop = 0
      intLeft = 0
    
    Case Top_Center:
      intTop = 0
      intLeft = Int((picPrint.Width - picPicture.Width) / 2)
    
    Case Top_Right
      intTop = 0
      intLeft = picPrint.Width - picPicture.Width
    
    Case Middle_Left:
      intTop = Int((picPrint.Height - picPicture.Height) / 2)
      intLeft = 0
    
    Case Middle_Center:
      intTop = Int((picPrint.Height - picPicture.Height) / 2)
      intLeft = Int((picPrint.Width - picPicture.Width) / 2)
    
    Case Middle_Right
      intTop = Int((picPrint.Height - picPicture.Height) / 2)
      intLeft = picPrint.Width - picPicture.Width
    
    Case Bottom_Left:
      intTop = picPrint.Height - picPicture.Height
      intLeft = 0
    
    Case Bottom_Center:
      intTop = picPrint.Height - picPicture.Height
      intLeft = Int((picPrint.Width - picPicture.Width) / 2)
    
    Case Bottom_Right
      intTop = picPrint.Height - picPicture.Height
      intLeft = picPrint.Width - picPicture.Width
  End Select
  picPrint.BackColor = lngBackColor
  picPrint.Cls
  If (blnStretch) Then
    Call picPrint.PaintPicture(picPicture.Picture, 0, 0, picPrint.Width, picPrint.Height)
  Else
    Call picPrint.PaintPicture(picPicture.Picture, intLeft, intTop, picPicture.Width, picPicture.Height)
  End If
  If (blnShowBorder) Then picPrint.Line (intLeft, intTop)-(intLeft + picPicture.Width, intTop + picPicture.Height), lngBorderColor, B
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'     UserControl Public Functions
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub GetPaperWH(ByRef intWidth As Integer, ByRef intHeight As Integer)
  Select Case prpPrint
    Case Manual_Paper:
      intWidth = intPaperWidth
      intHeight = intPaperHeight
    Case A3_Paper:
      intWidth = 297
      intHeight = 420
    Case A4_Paper:
      intWidth = 210
      intHeight = 297
    Case A5_Paper:
      intWidth = 148
      intHeight = 210
  End Select
End Sub

Public Function GetPreview() As IPictureDisp
  Call GeneratePrintPicture
  Set GetPreview = picPrint.Image
End Function

Public Sub PrintPicture()
  Dim lngX As Long, lngY As Long
  Dim intCounter As Integer
  Printer.ScaleMode = vbMillimeters
  Call GeneratePrintPicture
  lngX = lngMarginLeft / GetDeviceCaps(picPrint.hDC, LOGPIXELSX)
  lngY = lngMarginTop / GetDeviceCaps(picPrint.hDC, LOGPIXELSY)
  For intCounter = 1 To intCopies
    Call Printer.PaintPicture(picPrint.Image, lngMarginLeft, lngMarginTop)
    Printer.EndDoc
  Next intCounter
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'          UserControl Subs
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub UserControl_Initialize()
  alnPicture = Middle_Center
  lngBackColor = vbWhite
  lngBorderColor = vbBlack
  blnShowBorder = True
  strDefultPrinter = ""
  prpPrint = A4_Paper
  intPaperWidth = 210
  intPaperHeight = 297
  blnStretch = False
  intCopies = 1
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = UserControl.ScaleX(picPrinter.Width, vbPixels, vbTwips)
  UserControl.Height = UserControl.ScaleY(picPrinter.Height, vbPixels, vbTwips)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  alnPicture = PropBag.ReadProperty("Align", Middle_Center)
  Set picPicture.Picture = PropBag.ReadProperty("Picture", picPrinter.Picture)
  lngBackColor = PropBag.ReadProperty("BackColor", vbWhite)
  lngBorderColor = PropBag.ReadProperty("BorderColor", vbBlack)
  blnShowBorder = PropBag.ReadProperty("ShowBorder", True)
  strDefultPrinter = PropBag.ReadProperty("Printer", "")
  prpPrint = PropBag.ReadProperty("PaperSize", A4_Paper)
  intPaperWidth = PropBag.ReadProperty("PaperWidth", 210)
  intPaperHeight = PropBag.ReadProperty("PaperHeight", 297)
  blnStretch = PropBag.ReadProperty("Stretch", False)
  intCopies = PropBag.ReadProperty("Copies", 1)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Align", alnPicture, Middle_Center)
  Call PropBag.WriteProperty("Picture", picPicture.Picture, picPrinter.Picture)
  Call PropBag.WriteProperty("BackColor", lngBackColor, vbWhite)
  Call PropBag.WriteProperty("BorderColor", lngBorderColor, vbBlack)
  Call PropBag.WriteProperty("ShowBorder", blnShowBorder, True)
  Call PropBag.WriteProperty("Printer", strDefultPrinter, "")
  Call PropBag.WriteProperty("PaperSize", prpPrint, A4_Paper)
  Call PropBag.WriteProperty("PaperWidth", intPaperWidth, 210)
  Call PropBag.WriteProperty("PaperHeight", intPaperHeight, 297)
  Call PropBag.WriteProperty("Stretch", blnStretch, False)
  Call PropBag.WriteProperty("Copies", intCopies, 1)
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'        UserControl Properties
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Public Property Get Align() As PrintingAlign
  Align = alnPicture
End Property

Public Property Let Align(ByVal alnNewAlign As PrintingAlign)
  alnPicture = alnNewAlign
  PropertyChanged "Align"
End Property

Public Property Get Picture() As Picture
  Set Picture = picPicture.Picture
End Property

Public Property Set Picture(ByVal picNewPicture As Picture)
  Set picPicture.Picture = picNewPicture
  PropertyChanged "Picture"
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = lngBackColor
End Property

Public Property Let BackColor(ByVal colNewBackColor As OLE_COLOR)
  lngBackColor = colNewBackColor
  PropertyChanged "BackColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
  BorderColor = lngBorderColor
End Property

Public Property Let BorderColor(ByVal colNewBorderColor As OLE_COLOR)
  lngBorderColor = colNewBorderColor
  PropertyChanged "BorderColor"
End Property

Public Property Get ShowBorder() As Boolean
  ShowBorder = blnShowBorder
End Property

Public Property Let ShowBorder(ByVal blnNewShowBorder As Boolean)
  blnShowBorder = blnNewShowBorder
  PropertyChanged "ShowBorder"
End Property

Public Property Get DefultPrinter() As String
  DefultPrinter = strDefultPrinter
End Property

Public Property Let DefultPrinter(ByVal strNewDefultPrinter As String)
  strDefultPrinter = strNewDefultPrinter
  PropertyChanged "DefultPrinter"
End Property

Public Property Get PaperSize() As PaperSizeType
  PaperSize = prpPrint
End Property

Public Property Let PaperSize(ByVal prpNewPaperSize As PaperSizeType)
  prpPrint = prpNewPaperSize
  PropertyChanged "PaperSize"
End Property

Public Property Get PaperWidth() As Integer
  PaperWidth = intPaperWidth
End Property

Public Property Let PaperWidth(ByVal intNewPaperWidth As Integer)
  intPaperWidth = intNewPaperWidth
  PropertyChanged "PaperWidth"
End Property

Public Property Get PaperHeight() As Integer
  PaperHeight = intPaperHeight
End Property

Public Property Let PaperHeight(ByVal intNewPaperHeight As Integer)
  intPaperHeight = intNewPaperHeight
  PropertyChanged "PaperHeight"
End Property

Public Property Get Stretch() As Boolean
  Stretch = blnStretch
End Property

Public Property Let Stretch(ByVal blnNewStretch As Boolean)
  blnStretch = blnNewStretch
  PropertyChanged "Stretch"
End Property

Public Property Get Copies() As Integer
  Copies = intCopies
End Property

Public Property Let Copies(ByVal intNewCopies As Integer)
  intCopies = intNewCopies
  PropertyChanged "Copies"
End Property
