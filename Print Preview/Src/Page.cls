VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'////////////////////////////////////////////////////////
'///                 Output Page Class
'///                    (Page.cls)
'///_____________________________________________________
'/// This class mimics a printer page, it let us call
'/// 'Draw' methods that adds page elements as data structure.
'/// Printer information is stored here (Paper size, Port,
'/// Paper Try, etc.). The resulting pages could be exported
'/// to printer or any serial format as Text, HTML, RTF, etc.
'///_____________________________________________________
'/// Last modification  : Ago/10/2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Created
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit
#Const KEEP_VALUES = False

Private m_iSectionType As Integer
Private m_iStoredSection As Integer

Public Type PageElement
    Text As String
    Left As Long
    Top As Long
    Width As Long
    Height As Long
    #If KEEP_VALUES Then
    OrigLeft As Single
    OrigTop As Single
    OrigWidth As Single
    OrigHeight As Single
    #End If
    BackColor As Long
    Pen As Long
    ForeColor As Long
    Aligment As Long
    Type As Integer
    Size As Single
    Picture As StdPicture
    FontIndex As Integer
    BandIndex As Integer
    SectionType As Integer
    DisplayType As Integer
    Sunken As Boolean
    Checked As Boolean
End Type

Private m_oElements As Collection
Private m_lParentPtr As Long
Private m_lRangePtr As Long
Private m_iIndex As Long

Private m_iScaleMode As ScaleModeConstants
Private m_iMFactor As Single
Private m_PaperSize As Integer
Private m_nDisplayWidth As Single
Private m_nDisplayHeight As Single
Private m_nScaledWidth As Single
Private m_nScaledHeight As Single
Private m_bLandScape As Boolean
' Ratio conversion
Private m_sS2PRatioX As Single
Private m_sS2PRatioY As Single
Private m_lPrnGapX As Integer
Private m_lPrnGapY As Integer

Public PrinterPort As String
Public PrinterTray As Integer

Private Const TYPE_TEXT = 1
Private Const TYPE_LINE = 2
Private Const TYPE_BOX = 3
Private Const TYPE_PICT = 4
Private Const TYPE_CHKBOX = 5

Private Const TWIPS_PER_INCHE = 1440
Private Const TWIPS_PER_CENTIMETER = 567
Private Const TWIPS_PER_POINT = 72
Private Const TWIPS_PER_PIXEL = 15
Private Const TWIPS_PER_TWIP = 1

'Drawing functions
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function DrawTextAPI Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function Arc Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const BKMODE_TRANSPARENT = 1
' System related
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Const TECHNOLOGY = 2 ' Device type returns DT_PLOTTER, DT_RASDISPLAY, DT_RASPRINTER, DT_RASCAMERA, DT_CHARSTREAM, DT_METAFILE, or DT_DISPFILE
Private Const DT_PLOTTER = 0 ' Vector plotter
Private Const DT_RASDISPLAY = 1 ' Raster display
Private Const DT_RASPRINTER = 2 ' Raster printer
Private Const DT_RASCAMERA = 3 ' Raster camera
Private Const DT_CHARSTREAM = 4 ' Character stream
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const VERTSIZE = 6 ' Width, in millimeters, of the physical screen.
Private Const HORZSIZE = 4 ' Height, in millimeters, of the physical screen.
Private Const HORZRES = 8  ' Width, in pixels, of the screen.
Private Const VERTRES = 10 ' Height, in raster lines, of the screen.
Private Const LOGPIXELSX = 88 ' (&H58) Number of pixels per logical inch along the screen width.
Private Const LOGPIXELSY = 90 ' (&H5A) Number of pixels per logical inch along the screen height.' For printing devices:
Private Const PHYSICALWIDTH = 110 ' (&H6E) The physical width, in device units.
Private Const PHYSICALHEIGHT = 111 ' (&H6F) The physical height, in device units.
Private Const PHYSICALOFFSETX = 112 ' (&H70) The physical printable area horizontal margin.
Private Const PHYSICALOFFSETY = 113 '(&H71) The physical printable area vertical margin.
Private Const SCALINGFACTORX = 114 ' (&H72)  The scaling factor along the horizontal axis.
Private Const SCALINGFACTORY = 115 ' (&H73  The scaling factor along the vertical axis.

Public Sub DrawShape(iType As Integer, Left As Single, Top As Single, _
    Width As Single, Height As Single, _
    Optional lForeColor As OLE_COLOR, Optional lBackColor As OLE_COLOR = -1, _
    Optional iLineWidth As Integer = 1, Optional lPen As Long = 0)
    Dim LrElement As PageElement
    
    With LrElement
        .Type = TYPE_BOX
        #If KEEP_VALUES Then
        .OrigLeft = Left
        .OrigTop = Top
        .OrigWidth = Width
        .OrigHeight = Height
        #End If
        .Left = (Left * m_iMFactor)
        .Top = (Top * m_iMFactor)
        .Width = (Width * m_iMFactor)
        .Height = (Height * m_iMFactor)
        .BackColor = lBackColor
        .ForeColor = lForeColor
        .Size = iLineWidth
        .Pen = lPen
        .SectionType = m_iSectionType
        .DisplayType = iType
    End With
    m_oElements.Add LrElement
End Sub

Public Sub DrawCheckBox(ByVal iType As Integer, ByVal blnChecked As Boolean, _
    ByVal lLeft As Single, ByVal lTop As Single, Width As Single, Height As Single, Optional lfColor As OLE_COLOR, _
    Optional lbColor As OLE_COLOR = -1, Optional iLineWidth As Integer = 1, Optional blnSunken As Boolean)
    
    Dim LrElement As PageElement

    With LrElement
        .Type = TYPE_CHKBOX
        #If KEEP_VALUES Then
        .OrigLeft = lLeft
        .OrigTop = lTop
        #End If
        .Left = (lLeft * m_iMFactor)
        .Top = (lTop * m_iMFactor)
        .Width = (Width * m_iMFactor)
        .Height = (Height * m_iMFactor)
        .ForeColor = lfColor
        .BackColor = lbColor
        .SectionType = m_iSectionType
        .DisplayType = iType
        .Size = iLineWidth
        .Sunken = blnSunken
        .Checked = blnChecked
    End With
    m_oElements.Add LrElement

End Sub

Public Sub DrawLine(X1 As Single, Y1 As Single, _
    X2 As Single, Y2 As Single, _
    Optional lColor As OLE_COLOR, Optional iWidth As Integer = 1, Optional lPen As Long = 0)
    Dim LrElement As PageElement
    
    With LrElement
        .Type = TYPE_LINE
        #If KEEP_VALUES Then
        .OrigLeft = X1
        .OrigTop = Y1
        .OrigWidth = X2
        .OrigHeight = Y2
        #End If
        .Left = (X1 * m_iMFactor)
        .Top = (Y1 * m_iMFactor)
        .Width = (X2 * m_iMFactor)
        .Height = (Y2 * m_iMFactor)
        .Size = iWidth
        .Pen = lPen
        .ForeColor = lColor
        .SectionType = m_iSectionType
    End With
    m_oElements.Add LrElement
End Sub

Public Sub DrawPicture(Left As Single, Top As Single, _
    Width As Single, Height As Single, _
    oPicture As StdPicture, Optional bKeepRatio As Boolean = True, _
    Optional sFileName As String)
    Dim LrElement As PageElement
    
    With LrElement
        .Type = TYPE_PICT
        #If KEEP_VALUES Then
        .OrigLeft = Left
        .OrigTop = Top
        .OrigWidth = Width
        .OrigHeight = Height
        #End If
        .Left = (Left * m_iMFactor)
        .Top = (Top * m_iMFactor)
        .Width = (Width * m_iMFactor)
        .Height = (Height * m_iMFactor)
        Set .Picture = oPicture
        .Text = sFileName
        .SectionType = m_iSectionType
    End With
    m_oElements.Add LrElement
End Sub

Friend Property Get Elements() As Collection
    Set Elements = m_oElements
End Property

Public Sub NewSection()
    m_iSectionType = (m_iSectionType + 1)
End Sub

Friend Sub PrintIt()
    If (Printers.Count = 0) Then
        Exit Sub
    End If
    Dim LnIdx As Long
    Dim LoRange As clsRangeSelector
    Dim LoFMap As FontMap
    Dim LhOldFont As Long
    Dim LhOldFntIdx As Integer
    
    On Error GoTo ERR_H
    Set LoRange = Range
    Set LoFMap = Parent.FontMap
    ' Verify if is a printable page....
    If (LoRange.ElementInRange(m_iIndex)) Then
        If (m_iIndex = LoRange.RangeMin) Then
            ' This is the first printed page
            Printer.KillDoc
        Else
            Printer.NewPage
        End If
        If (PrinterPort <> Printer.Port) Then
        '////////////////////////////////////////////////
        '/// TODO: Code for Printer Change on the Fly ///
        '////////////////////////////////////////////////
        End If
        If (PrinterTray <> Printer.PaperBin) Then
            On Error Resume Next
            Printer.PaperBin = PrinterTray
            On Error GoTo 0
        End If
        ' Adjust conversion ratio
        prvGetPrn2ScrRatio
        ' Prepare Printer page
        prvSetPrinterPage
        ' Inits page
        On Error GoTo ERR_H
        Printer.Print ""
        ' Prints page elements
        ' Stores original Font object
        LhOldFntIdx = 1
        #If USE_LOG_FONT Then
        LhOldFont = SelectObject(Printer.hDC, LoFMap.Item(LhOldFntIdx).Handle)
        #Else
        With LoFMap.Item(1)
            Printer.FontName = .FaceName
            Printer.FontSize = .Size
            Printer.FontBold = .Bold
            Printer.FontItalic = .Italic
            Printer.FontUnderline = .Underline
            Printer.FontStrikethru = .Strikethrough
            Printer.FontTransparent = True
            Printer.Print ""
        End With
        #End If
        If (m_oElements.Count > 0) Then
            Printer.Print ""
        End If
        DoEvents
        For LnIdx = 1 To m_oElements.Count
            With m_oElements.Item(LnIdx)
                Printer.CurrentX = 0
                Printer.CurrentY = 0
                Select Case .Type
                    Case TYPE_TEXT
                        If (LhOldFntIdx <> .FontIndex) Then
                            Dim LnFntIdx As Integer
                            
                            LnFntIdx = .FontIndex
                            #If USE_LOG_FONT Then
                            Call SelectObject(Printer.hDC, LoFMap.Item(LnFntIdx).Handle)
                            Printer.Print ""
                            #Else
                            
                            With LoFMap.Item(LnFntIdx)
                                Printer.FontName = .FaceName
                                Printer.FontSize = .Size
                                Printer.FontBold = .Bold
                                Printer.FontItalic = .Italic
                                Printer.FontUnderline = .Underline
                                Printer.FontStrikethru = .Strikethrough
                                Printer.FontTransparent = True
                                Printer.Print ""
                            End With
                            #End If
                            LhOldFntIdx = LnFntIdx
                        End If
                        Call prvPrintText(.Text, (.Left * m_sS2PRatioX), (.Top * m_sS2PRatioY), _
                            (.Width * m_sS2PRatioX), (.Height * m_sS2PRatioY), .ForeColor, .BackColor, .Aligment)
                    Case TYPE_LINE
                        Call prvPrintLine((.Left * m_sS2PRatioX), (.Top * m_sS2PRatioY), _
                            (.Width * m_sS2PRatioX), (.Height * m_sS2PRatioY), _
                            .ForeColor, CInt(.Size * m_sS2PRatioY), .Pen)
                    Case TYPE_BOX
                        Call prvPrintShape(.DisplayType, (.Left * m_sS2PRatioX), (.Top * m_sS2PRatioY), _
                            (.Width * m_sS2PRatioX), (.Height * m_sS2PRatioY), _
                            CInt(.Size * m_sS2PRatioY), .ForeColor, .BackColor, .Pen)
                    Case TYPE_CHKBOX
                        Call prvPrintCheckBox(.DisplayType, .Checked, (.Left * m_sS2PRatioX), (.Top * m_sS2PRatioY), _
                        (.Width * m_sS2PRatioX), (.Height * m_sS2PRatioY), CInt(.Size * m_sS2PRatioY), .ForeColor, .BackColor, .Sunken)
                    Case TYPE_PICT
                        Call prvPrintPicture(.Picture, (.Left * m_sS2PRatioX), (.Top * m_sS2PRatioY), _
                            (.Width * m_sS2PRatioX), (.Height * m_sS2PRatioY))
                End Select
            End With
            DoEvents
        Next
        ' Restores Original Font
        #If USE_LOG_FONT Then
        LhOldFont = SelectObject(Printer.hDC, LhOldFont)
        #Else
        #End If
    End If
    Exit Sub
ERR_H:
    Me.RaiseErr Err.Number, Err.Source, Err.Description
End Sub

Private Sub prvPrintText(ByVal sText As String, ByVal lLeft As Long, ByVal lTop As Long, _
    ByVal lWidth As Long, ByVal lHeight As Long, ByVal lfColor As OLE_COLOR, _
    ByVal lbColor As OLE_COLOR, ByVal lFlags As Long)
    On Error GoTo ERR_H
    Dim LhDC  As Long
    Dim lBrush As Long
    Dim LnOldBMode As Long
    Dim LrBox As RECT
    
    LhDC = Printer.hDC
'////////////////////////////////////////
'///      Coordenates Calculation
'////////////////////////////////////////
    With LrBox
        .Left = (lLeft - m_lPrnGapX)
        .Top = (lTop - m_lPrnGapY)
        .Right = (.Left + lWidth)
        .Bottom = (.Top + lHeight)
    End With
'////////////////////////////////////////
'///       Defaults Resolution
'////////////////////////////////////////
    If (lFlags = 0) Then
        lFlags = (DT_WORDBREAK Or DT_LEFT)
    ElseIf (lFlags = 1) Then
        lFlags = (DT_WORDBREAK Or DT_RIGHT)
    ElseIf (lFlags = 2) Then
        lFlags = (DT_WORDBREAK Or DT_CENTER)
    End If
    If (lfColor = -1) Then
        lfColor = vbBlack
    End If
'////////////////////////////////////////
'///         Background Drawing
'///_____________________________________
'/// If background is not transparent
'/// (<> -1) then draws it...
'////////////////////////////////////////
    If (lbColor <> -1) Then
        lBrush = CreateSolidBrush(lbColor)
        Call FillRect(LhDC, LrBox, lBrush)
        DeleteObject (lBrush)
    End If
'////////////////////////////////////////
'///           Text Drawing
'////////////////////////////////////////
    LnOldBMode = SetBkMode(LhDC, BKMODE_TRANSPARENT)
'    #If USE_LOG_FONT Then
    Call SetTextColor(LhDC, lfColor)
    DrawTextAPI LhDC, sText, Len(sText), LrBox, lFlags Or DT_NOPREFIX
'    #Else
'    Printer.FillColor = lFColor
'    printer.CurrentX=
'    #End If
    Call SetBkMode(LhDC, LnOldBMode)
    Exit Sub
ERR_H:
    Me.RaiseErr Err.Number, "DrawText"
End Sub

Private Sub prvPrintLine(lX1 As Long, lY1 As Long, lX2 As Long, lY2 As Long, _
    lColor As OLE_COLOR, Optional lWidth As Long = 1, Optional lPen As Long)
    
    On Error GoTo ERR_H
    Dim LhDC As Long
    Dim LrPos As POINTAPI
    Dim lhPen As Long
    Dim LhOldPen As Long
    Dim LrBox As RECT
    
    ' Gets printer DC
    LhDC = Printer.hDC
    ' Converts coordinates
    With LrBox
        .Left = (lX1 - m_lPrnGapX)
        .Top = (lY1 - m_lPrnGapY)
        .Right = (lX2 - m_lPrnGapX)
        .Bottom = (lY2 - m_lPrnGapY)
    End With
    ' Creates apropiate pen and selects it in the DC
    If (lWidth = 0) Then
        lhPen = CreatePen(PS_DOT, 1, lColor)
    Else
        lhPen = CreatePen(lPen, lWidth, lColor)
    End If
    If lhPen Then
        LhOldPen = SelectObject(LhDC, lhPen)
    End If
    ' Draws box lines
    With LrBox
        MoveToEx LhDC, .Left, .Top, LrPos
        LineTo LhDC, .Right, .Bottom
    End With
    If lhPen Then
        Call SelectObject(LhDC, LhOldPen)
        Call DeleteObject(lhPen)
    End If
    Exit Sub
ERR_H:
    Me.RaiseErr Err.Number, "DrawLine"
End Sub

Private Sub prvPrintPicture(oPicture As StdPicture, lLeft As Long, lTop As Long, _
    lWidth As Long, lHeight As Long, Optional bKeepRatio As Boolean = True, _
    Optional bVCenter As Boolean = True, Optional bHCenter As Boolean = True, Optional lPen As Long)
    On Error GoTo ERR_H
#Const USE_PALETTE = True
#Const USE_DISPLAY = False
#Const USE_BITBLT = True
#Const USE_BKMODE = False
#Const USE_HALFTONE = False
    Dim LhDC As Long
    Dim LhDCTemp As Long
    Dim LhOldPalette As Long
    Dim LhOldBitmap As Long
    Dim LrPoint As POINTAPI
    Dim LrReal As RECT
    Dim LrBitmap As BITMAP
    #If USE_HALFTONE Then
    Dim LnOldStrMode As Long
    #End If
    
    ' Gets printer DC
    LhDC = Printer.hDC
    'Validate that a bitmap was passed in
    If (oPicture.Type = vbPicTypeBitmap) Then
        With oPicture
            ' get bitmap info
            GetObjectAPI .Handle, Len(LrBitmap), LrBitmap   'dest
            #If USE_DISPLAY Then
            Dim LhDC As Long
            LhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
            LhDCTemp = CreateCompatibleDC(LhDC)
            DeleteDC LhDC
            #Else
            LhDCTemp = CreateCompatibleDC(LhDC)
            #End If
            #If USE_PALETTE Then
            LhOldPalette = SelectPalette(LhDCTemp, .hPal, True)
            RealizePalette LhDCTemp
            #End If
            'Select bitmap into DC
            LhOldBitmap = SelectObject(LhDCTemp, .Handle)
        End With
        #If USE_BITBLT Then
        If ((lWidth = LrBitmap.bmWidth) And (lHeight = LrBitmap.bmHeight)) Then
            #If USE_HALFTONE Then
            LnOldStrMode = SetStretchBltMode(LhDC, HALFTONE)
            #End If
            BitBlt LhDC, lLeft, lTop, lWidth, lHeight, LhDCTemp, 0, 0, SRCCOPY
            #If USE_HALFTONE Then
            LnOldStrMode = SetStretchBltMode(LhDC, LnOldStrMode)
            #End If
        Else
        #Else
        If True Then
        #End If
            With LrReal
                If bKeepRatio Then
                    Dim LnRatio As Single
                    Dim LnDrwRatio As Single
                    Dim LnImgRatio As Single
                    
                    With LrBitmap
                        LnDrwRatio = (lHeight / lWidth)
                        LnImgRatio = (.bmHeight / .bmWidth)
                        If (LnDrwRatio > LnImgRatio) Then
                            LnRatio = (lWidth / .bmWidth)
                        Else
                            LnRatio = (lHeight / .bmHeight)
                        End If
                    End With
                    ' Used for width & height
                    .Right = (LrBitmap.bmWidth * LnRatio)
                    .Bottom = (LrBitmap.bmHeight * LnRatio)
                    ' Left & Top positions
                    If bHCenter Then
                        .Left = (lLeft + ((lWidth - .Right) / 2))
                    Else
                        .Left = lLeft
                    End If
                    If bVCenter Then
                        .Top = (lTop + ((lHeight - .Bottom) / 2))
                    Else
                        .Top = lTop
                    End If
                    ' Now converts coordinates to printer ratio
                    .Left = (.Left - m_lPrnGapX)
                    .Top = (.Top - m_lPrnGapY)
                    .Right = (.Right - m_lPrnGapX)
                    .Bottom = (.Bottom - m_lPrnGapY)
                Else
                    .Left = (lLeft - m_lPrnGapX)
                    .Top = (lTop - m_lPrnGapY)
                    ' Used for width & height
                    .Right = (lWidth - m_lPrnGapX)
                    .Bottom = (lHeight - m_lPrnGapY)
                End If
                'Copy to destination DC
                #If USE_BKMODE Then
                Dim LnOldBkColor As Long
                LnOldBkColor = SetBkColor(LhDC, 0&)
                #End If
                #If USE_HALFTONE Then
                LnOldStrMode = SetStretchBltMode(LhDC, HALFTONE)
                #End If
                Call SetBrushOrgEx(LhDC, 0, 0, LrPoint)
                Call StretchBlt(LhDC, .Left, .Top, .Right, .Bottom, LhDCTemp, 0, 0, LrBitmap.bmWidth, LrBitmap.bmHeight, SRCCOPY)
                #If USE_BKMODE Then
                LnOldBkColor = SetBkColor(LhDC, LnOldBkColor)
                #End If
                #If USE_HALFTONE Then
                LnOldStrMode = SetStretchBltMode(LhDC, LnOldStrMode)
                #End If
            End With
        End If
        'Cleanup
        SelectObject LhDCTemp, LhOldBitmap
        #If USE_PALETTE Then
        SelectPalette LhDCTemp, LhOldPalette, True
        RealizePalette LhDCTemp
        #End If
        DeleteDC LhDCTemp
    Else
        Me.RaiseErr ecInvalidObjType, "DrawPicture", "Invalid picture format."
    End If
    Exit Sub
ERR_H:
    Me.RaiseErr Err.Number, "DrawPicture"
End Sub

Private Sub prvGetPrn2ScrRatio()
    Dim LhDC As Long
    Dim LnScrResX As Long
    Dim LnScrResY As Long
    Dim LnPrnResX As Long
    Dim LnPrnResY As Long
    ' Gets Screen logical inches resolution
    LhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    LnScrResX = GetDeviceCaps(LhDC, LOGPIXELSX)
    LnScrResY = GetDeviceCaps(LhDC, LOGPIXELSY)
    DeleteDC LhDC
    ' Gets Printer logical inches resolution
    With Printer
        LnPrnResX = GetDeviceCaps(.hDC, LOGPIXELSX)
        LnPrnResY = GetDeviceCaps(.hDC, LOGPIXELSY)
        ' Gets Screen logical inches resolution
        m_lPrnGapX = GetDeviceCaps(.hDC, PHYSICALOFFSETX)
        m_lPrnGapY = GetDeviceCaps(.hDC, PHYSICALOFFSETY)
    End With
    m_sS2PRatioX = (LnPrnResX / LnScrResX)
    m_sS2PRatioY = (LnPrnResY / LnScrResY)
End Sub


Private Sub prvValidatePaperSize()
    If (m_nDisplayWidth = 816) Then ' Letter, A4 or Legal (Portrait)
        Select Case m_nDisplayHeight
            Case 1056 ' Letter
                m_PaperSize = vbPRPSLetter
            Case 1152 ' A4
                m_PaperSize = vbPRPSA4
            Case 1344 ' Legal
                m_PaperSize = vbPRPSLegal
            Case Else ' User
                m_PaperSize = vbPRPSUser
        End Select
    ElseIf (m_nDisplayWidth = 1056) Then ' Letter
        If (m_nDisplayHeight = 816) Then ' Letter
            m_PaperSize = vbPRPSLetter
        Else ' User
            m_PaperSize = vbPRPSUser
        End If
    ElseIf (m_nDisplayWidth = 1152) Then ' A4
        If (m_nDisplayHeight = 816) Then ' A4
            m_PaperSize = vbPRPSA4
        Else ' User
            m_PaperSize = vbPRPSUser
        End If
    ElseIf (m_nDisplayWidth = 1344) Then ' Legal
        If (m_nDisplayHeight = 816) Then ' Legal
            m_PaperSize = vbPRPSLegal
        Else ' User
            m_PaperSize = vbPRPSUser
        End If
    Else
        m_PaperSize = vbPRPSUser
    End If
    m_bLandScape = (m_nDisplayWidth > m_nDisplayHeight)
End Sub

Public Property Get SectionType() As Integer
    SectionType = m_iSectionType
End Property

Public Property Let SectionType(iType As Integer)
    m_iSectionType = iType
End Property


Friend Sub SetPaperProps(ByVal iPaperSize As Integer, _
    iScaleMode As ScaleModeConstants, iMFactor As Single, _
    lWidth As Single, lHeight As Single, bLandscape As Boolean)
    m_PaperSize = iPaperSize
    m_iScaleMode = iScaleMode
    m_iMFactor = iMFactor
    m_nDisplayWidth = lWidth
    m_nDisplayHeight = lHeight
    m_bLandScape = bLandscape
End Sub

'Public Sub AddElement(iType As Integer, lLeft As Long, lTop As Long, _
'    lWidth As Long, lHeight As Long, Optional nSize As Single = 1, _
'    Optional lBackColor As Long = -1, Optional lForeColor As Long = -1, _
'    Optional sText As String, Optional lFlags As Long, _
'    Optional oPicture As StdPicture, Optional iFontIndex As Integer)
'    m_iElementsCount = (m_iElementsCount + 1)
'    ReDim Preserve m_oElements(1 To m_iElementsCount)
'    With m_oElements(m_iElementsCount)
'        .Type = iType
'        .Left = lLeft
'        .Top = lTop
'        .Width = lWidth
'        .Height = lHeight
'        .Text = sText
'        .Size = nSize
'        .BackColor = lBackColor
'        .ForeColor = lForeColor
'        .Aligment = lFlags
'        If (iFontIndex = 0) Then
'            .FontIndex = Parent.FontMap.ActiveFont.Index
'        Else
'            .FontIndex = iFontIndex
'        End If
'        Set .Picture = oPicture
'    End With
'End Sub

Public Property Get PaperSize() As Integer
    PaperSize = m_PaperSize
End Property

Public Function SetFont(FaceName As String, Size As Single, _
    Optional Bold As Boolean, Optional Italic As Boolean, _
    Optional Underline As Boolean, Optional Strikethrough As Boolean, _
    Optional Rotation As Integer) As Integer
    SetFont = Parent.FontMap.Add(FaceName, Size, Bold, _
        Italic, Underline, Strikethrough, Rotation)
End Function

Public Property Get Count() As Integer
    Count = m_oElements.Count
End Property

Public Property Get DisplayWidth() As Long
    DisplayWidth = m_nDisplayWidth
End Property

Public Property Get DisplayHeight() As Long
    DisplayHeight = m_nDisplayHeight
End Property

Public Sub DrawText(sText As String, lLeft As Single, lTop As Single, _
    lWidth As Single, lHeight As Single, _
    Optional lfColor As OLE_COLOR = vbBlack, Optional lbColor As OLE_COLOR = -1, _
    Optional lAligment As Long, Optional iFontIndex As Integer)
    Dim LrElement As PageElement
    
    With LrElement
        .Type = TYPE_TEXT
        #If KEEP_VALUES Then
        .OrigLeft = lLeft
        .OrigTop = lTop
        .OrigWidth = lWidth
        .OrigHeight = lHeight
        #End If
        .Left = (lLeft * m_iMFactor)
        .Top = (lTop * m_iMFactor)
        .Width = (lWidth * m_iMFactor)
        .Height = (lHeight * m_iMFactor)
        .Text = sText
        .BackColor = lbColor
        .ForeColor = lfColor
        .Aligment = lAligment
        If (iFontIndex = 0) Then
            .FontIndex = Parent.FontMap.ActiveFont.Index
        Else
            .FontIndex = iFontIndex
        End If
        .SectionType = m_iSectionType
    End With
    m_oElements.Add LrElement
End Sub

Private Sub prvSetPrinterPage()
    On Error GoTo ERR_TRAP
    Printer.ScaleMode = m_iScaleMode
'    Printer.ScaleMode = vbPixels
    Select Case PaperSize
        Case vbPRPSLetter, vbPRPSLegal, vbPRPSA4
            Printer.PaperSize = m_PaperSize
        Case Else
            With Printer
                .PaperSize = vbPRPSUser
                .Width = m_nDisplayWidth
                .Height = m_nDisplayHeight
                .ScaleWidth = m_nDisplayWidth
                .ScaleHeight = m_nDisplayHeight
            End With
    End Select
    Printer.ScaleMode = m_iScaleMode
    Printer.Orientation = IIf(m_bLandScape, vbPRORLandscape, vbPRORPortrait)
    Exit Sub
ERR_TRAP:
    If (Err.Number = 380) Then Resume Next
    'Me.RaiseErr Err.Number, "prvSetPrinterPage"
End Sub

Public Property Get Element(Index As Integer) As PageElement
    Element = m_oElements.Item(Index)
End Property

Public Property Get Enabled() As Boolean
    Enabled = Range.ElementInRange(m_iIndex)
End Property

Public Property Get Index() As Integer
    Index = m_iIndex
End Property

Friend Property Let Index(iIndex As Integer)
    m_iIndex = iIndex
End Property

Public Property Get Parent() As Pages
    Dim lObjTmp As Object
        ' Turn the pointer into an illegal, uncounted interface
    CopyMemory lObjTmp, m_lParentPtr, 4
    ' Do NOT hit the End button here! You will crash!
    ' Assign to legal reference
    Set Parent = lObjTmp
    ' Still do NOT hit the End button here! You will still crash!
    ' Destroy the illegal reference
    CopyMemory lObjTmp, 0&, 4
End Property

Friend Property Get Range() As clsRangeSelector
    Dim lObjTmp As Object
        ' Turn the pointer into an illegal, uncounted interface
    CopyMemory lObjTmp, m_lRangePtr, 4
    ' Do NOT hit the End button here! You will crash!
    ' Assign to legal reference
    Set Range = lObjTmp
    ' Still do NOT hit the End button here! You will still crash!
    ' Destroy the illegal reference
    CopyMemory lObjTmp, 0&, 4
End Property

Friend Property Set Parent(oParent As Pages)
    m_lParentPtr = ObjPtr(oParent)
End Property

Friend Property Set Range(oRange As clsRangeSelector)
    m_lRangePtr = ObjPtr(oRange)
End Property

Public Sub SetActive()
    Parent.SelectPage m_iIndex
End Sub

Public Sub Clear()
    With m_oElements
        While .Count
            .Remove .Count
        Wend
    End With
    m_iSectionType = 0
End Sub

Public Property Let Landscape(ByVal vData As Boolean)
    Dim LnTemp As Long
    
    m_bLandScape = vData
    If m_bLandScape Then
        If (m_nDisplayWidth < m_nDisplayHeight) Then
            LnTemp = m_nDisplayWidth
            m_nDisplayWidth = m_nDisplayHeight
            m_nDisplayHeight = LnTemp
        End If
    Else
        If (m_nDisplayWidth > m_nDisplayHeight) Then
            LnTemp = m_nDisplayWidth
            m_nDisplayWidth = m_nDisplayHeight
            m_nDisplayHeight = LnTemp
        End If
    End If
End Property

Public Property Get Landscape() As Boolean
    Landscape = m_bLandScape
End Property

Public Property Let Height(ByVal vData As Single)
    m_nScaledHeight = vData
    m_nDisplayHeight = (m_nScaledHeight * m_iMFactor)
    prvValidatePaperSize
End Property

Public Sub StoreSection()
    m_iStoredSection = m_iSectionType
End Sub

Public Sub RestoreSection()
    m_iSectionType = m_iStoredSection
End Sub

Public Property Get Width() As Single
    Width = m_nScaledWidth
End Property

Public Property Get Height() As Single
    Height = m_nScaledHeight
End Property

Public Property Let Width(ByVal vData As Single)
    m_nScaledWidth = vData
    m_nDisplayWidth = (m_nScaledWidth * m_iMFactor)
    prvValidatePaperSize
End Property

Friend Sub RaiseErr(ByVal lErrNum As RSErrorCode, Optional sRoutineName As String, _
    Optional sDescription As String)
    RaiseError lErrNum, TypeName(Me), sRoutineName, sDescription, Erl
End Sub

Private Sub prvPrintShape(ByVal iType As Integer, ByVal lLeft As Long, ByVal lTop As Long, _
    ByVal lWidth As Long, ByVal lHeight As Long, Optional lLineWidth As Integer = 1, _
    Optional lfColor As OLE_COLOR, Optional lbColor As OLE_COLOR = -1, Optional lPen As Long = 0)
    On Error GoTo ERR_H
    
'* Note from RG (04/08/2006)
'       Modified prvPrintBox sub to draw Rectangle, Square, Oval, Circle, Rounded Rectangle and Rounded Square
'       Renamed sub prvPrintShape

    Dim LhDC As Long
    Dim LnTop As Long
    Dim LnLeft As Long
    Dim LnWidth As Long
    Dim LnHeight As Long
    Dim LrBox As RECT
    Dim LrPos As POINTAPI
    Dim lhPen As Long
    Dim LhOldPen As Long
    Dim LhBrush As Long
    
    LhDC = Printer.hDC
    If iType = 0 Or iType = 2 Or iType = 4 Then     'if rectangle, ellipse or rounded rectangle
        With LrBox
            .Left = lLeft
            .Top = lTop
            .Right = lLeft + lWidth
            .Bottom = lTop + lHeight
            ' Now converts coordinates to printer ratio
            .Left = (.Left - m_lPrnGapX)
            .Top = (.Top - m_lPrnGapY)
            .Right = (.Right - m_lPrnGapX)
            .Bottom = (.Bottom - m_lPrnGapY)
        End With
    Else                                            'otherwise will be square, circle or rounded square
        With LrBox
            If lWidth < lHeight Then
                .Top = lTop + (lHeight / 2) - (lWidth / 2)
                lHeight = lWidth
                .Left = lLeft
            Else
                .Left = lLeft + (lWidth / 2) - (lHeight / 2)
                lWidth = lHeight
                .Top = lTop
            End If
            .Right = .Left + lWidth
            .Bottom = .Top + lHeight
        
            ' Now converts coordinates to printer ratio
            .Left = (.Left - m_lPrnGapX)
            .Top = (.Top - m_lPrnGapY)
            .Right = (.Right - m_lPrnGapX)
            .Bottom = (.Bottom - m_lPrnGapY)
        End With
    End If
    
    If iType < 2 Then      'rectangle or square
        If (lbColor <> -1) Then
            LhBrush = CreateSolidBrush(lbColor)
            Call FillRect(LhDC, LrBox, LhBrush)
            Call DeleteObject(LhBrush)
        End If
        
        If (lfColor <> -1) Then
            If (lLineWidth = 0) Then
                lhPen = CreatePen(PS_DOT, 1, lfColor)
            Else
                lhPen = CreatePen(lPen, lLineWidth, lfColor)
            End If
            If (lhPen <> 0) Then
                LhOldPen = SelectObject(LhDC, lhPen)
            End If
            With LrBox
                MoveToEx LhDC, .Left, .Top, LrPos
                LineTo LhDC, .Right, .Top
                LineTo LhDC, .Right, .Bottom
                LineTo LhDC, .Left, .Bottom
                LineTo LhDC, .Left, .Top
            End With
            If lhPen Then
                Call SelectObject(LhDC, LhOldPen)
                Call DeleteObject(lhPen)
            End If
        End If
    
    ElseIf iType < 4 Then           'ellipse or circle
        Dim EllipReg As Long
        EllipReg = CreateEllipticRgn(LrBox.Left, LrBox.Top, LrBox.Right, LrBox.Bottom)
        If (lbColor <> -1) Then
            LhBrush = CreateSolidBrush(lbColor)
            Call FillRgn(LhDC, EllipReg, LhBrush)
            Call DeleteObject(LhBrush)
        End If
    
        If (lfColor <> -1) Then
            If (lLineWidth = 0) Then
                lhPen = CreatePen(PS_DOT, 1, lfColor)
            Else
                lhPen = CreatePen(lPen, lLineWidth, lfColor)
            End If
            If (lhPen <> 0) Then
                LhOldPen = SelectObject(LhDC, lhPen)
            End If
        
            Call Arc(LhDC, LrBox.Left, LrBox.Top, LrBox.Right, LrBox.Bottom, _
            LrBox.Right, LrBox.Top + (LrBox.Bottom - LrBox.Top) / 2, LrBox.Right, LrBox.Top + (LrBox.Bottom - LrBox.Top) / 2)
        
            If lhPen Then
                Call SelectObject(LhDC, LhOldPen)
                Call DeleteObject(lhPen)
            End If
        End If
    
    ElseIf iType < 6 Then          'rounded rectangle or rounded square
        Dim RnRcReg As Long
        Dim ArcDiam As Long
        Dim LineOffset As Long
        If (LrBox.Right - LrBox.Left) < (LrBox.Bottom - LrBox.Top) Then
            ArcDiam = (LrBox.Right - LrBox.Left) / 4
        Else
            ArcDiam = (LrBox.Bottom - LrBox.Top) / 4
        End If
        LineOffset = ArcDiam / 2
        RnRcReg = CreateRoundRectRgn(LrBox.Left, LrBox.Top, LrBox.Right, LrBox.Bottom, ArcDiam, ArcDiam)
        If (lbColor <> -1) Then
            LhBrush = CreateSolidBrush(lbColor)
            Call FillRgn(LhDC, RnRcReg, LhBrush)
            Call DeleteObject(LhBrush)
        End If
        
        If (lfColor <> -1) Then
            If (lLineWidth = 0) Then
                lhPen = CreatePen(PS_DOT, 1, lfColor)
            Else
                lhPen = CreatePen(lPen, lLineWidth, lfColor)
            End If
            If (lhPen <> 0) Then
                LhOldPen = SelectObject(LhDC, lhPen)
            End If
        
            MoveToEx LhDC, LrBox.Left + LineOffset, LrBox.Top, LrPos
            LineTo LhDC, LrBox.Right - LineOffset, LrBox.Top
            MoveToEx LhDC, LrBox.Right, LrBox.Top + LineOffset, LrPos
            LineTo LhDC, LrBox.Right, LrBox.Bottom - LineOffset
            MoveToEx LhDC, LrBox.Right - LineOffset, LrBox.Bottom, LrPos
            LineTo LhDC, LrBox.Left + LineOffset, LrBox.Bottom
            MoveToEx LhDC, LrBox.Left, LrBox.Bottom - LineOffset, LrPos
            LineTo LhDC, LrBox.Left, LrBox.Top + LineOffset
        
            Arc LhDC, LrBox.Right - ArcDiam, LrBox.Top, LrBox.Right, LrBox.Top + ArcDiam, _
            LrBox.Right, LrBox.Top + LineOffset, LrBox.Right - LineOffset, LrBox.Top
            
            Arc LhDC, LrBox.Right - ArcDiam, LrBox.Bottom - ArcDiam, LrBox.Right, LrBox.Bottom, _
            LrBox.Right - LineOffset, LrBox.Bottom, LrBox.Right, LrBox.Bottom - LineOffset
            
            Arc LhDC, LrBox.Left, LrBox.Bottom - ArcDiam, LrBox.Left + ArcDiam, LrBox.Bottom, _
            LrBox.Left, LrBox.Bottom - LineOffset, LrBox.Left + LineOffset, LrBox.Bottom
            
            Arc LhDC, LrBox.Left, LrBox.Top, LrBox.Left + ArcDiam, LrBox.Top + ArcDiam, _
            LrBox.Left + LineOffset, LrBox.Top, LrBox.Left, LrBox.Top + LineOffset
        
            If lhPen Then
                Call SelectObject(LhDC, LhOldPen)
                Call DeleteObject(lhPen)
            End If
        End If
    
    End If
    
    Exit Sub

ERR_H:
    Me.RaiseErr Err.Number, "DrawShape"
End Sub

Public Sub prvPrintCheckBox(ByVal iType As Integer, ByVal blnChecked As Boolean, ByVal lLeft As Long, _
ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long, Optional lLineWidth As Integer = 1, _
Optional lfColor As OLE_COLOR = -1, Optional lbColor As OLE_COLOR = -1, Optional blnSunken As Boolean)
    
    On Error GoTo ERR_H
    Dim LhDC As Long
    Dim LnTop As Long
    Dim LnLeft As Long
    Dim LrBox As RECT
    Dim LrPos As POINTAPI
    Dim lhPen As Long
    Dim LhOldPen As Long
    Dim LhBrush As Long
    
    LhDC = Printer.hDC
    
    If iType = 0 Then       'box with check mark
        With LrBox
            .Left = lLeft
            .Top = lTop + (lHeight * 0.1875)
            .Right = lLeft + lWidth
            .Bottom = lTop + lHeight
        
            ' Now converts coordinates to printer ratio
            .Left = (.Left - m_lPrnGapX)
            .Top = (.Top - m_lPrnGapY)
            .Right = (.Right - m_lPrnGapX)
            .Bottom = (.Bottom - m_lPrnGapY)
        End With
        
        If (lbColor <> -1) Then
            LhBrush = CreateSolidBrush(lbColor)
            Call FillRect(LhDC, LrBox, LhBrush)
            Call DeleteObject(LhBrush)
        End If
    
        lhPen = CreatePen(PS_SOLID, lLineWidth, lfColor)
        If (lhPen <> 0) Then
            LhOldPen = SelectObject(LhDC, lhPen)
        End If
        If blnSunken Then
            MoveToEx LhDC, LrBox.Left + (lWidth * 0.0833), LrBox.Top + (lHeight * 0.9375), LrPos
            LineTo LhDC, LrBox.Left + (lWidth * 0.0833), LrBox.Top + (lHeight * 0.25)
            LineTo LhDC, LrBox.Left + (lWidth * 0.9167), LrBox.Top + (lHeight * 0.25)
        End If
        With LrBox
            MoveToEx LhDC, .Left, .Top, LrPos
            LineTo LhDC, .Right, .Top
            LineTo LhDC, .Right, .Bottom
            LineTo LhDC, .Left, .Bottom
            LineTo LhDC, .Left, .Top
        End With
        
        If blnChecked Then
            lhPen = CreatePen(PS_SOLID, lLineWidth * 3, lfColor)
            If (lhPen <> 0) Then
                LhOldPen = SelectObject(LhDC, lhPen)
            End If
            MoveToEx LhDC, LrBox.Left + (lWidth * 0.25), LrBox.Top + (lHeight * 0.4375), LrPos
            LineTo LhDC, LrBox.Left + (lWidth * 0.5), LrBox.Top + (lHeight * 0.8125)
            LineTo LhDC, LrBox.Left + (lWidth * 0.833), LrBox.Top
        End If
        
    ElseIf iType = 1 Then       'round radio button with center dot
        With LrBox
            .Left = lLeft
            .Top = lTop
            .Right = lLeft + (lWidth * 1.0833)
            .Bottom = lTop + (lHeight * 0.8125)
            
            ' Now converts coordinates to printer ratio
            .Left = (.Left - m_lPrnGapX)
            .Top = (.Top - m_lPrnGapY)
            .Right = (.Right - m_lPrnGapX)
            .Bottom = (.Bottom - m_lPrnGapY)
        End With
        
        Dim lEllReg As Long
        If (lbColor <> -1) Then
            lEllReg = CreateEllipticRgn(LrBox.Left, LrBox.Top, LrBox.Right, LrBox.Bottom)
            LhBrush = CreateSolidBrush(lbColor)
            Call FillRgn(LhDC, lEllReg, LhBrush)
            Call DeleteObject(LhBrush)
        End If
        
        lhPen = CreatePen(PS_SOLID, lLineWidth * 2, lfColor)
        If (lhPen <> 0) Then
            LhOldPen = SelectObject(LhDC, lhPen)
        End If

        Call Arc(LhDC, LrBox.Left, LrBox.Top, LrBox.Right, LrBox.Bottom, _
        LrBox.Left + (lWidth * 0.5), LrBox.Top, LrBox.Left + (lWidth * 0.5), LrBox.Top)
        
        If blnSunken Then
            Call Arc(LhDC, LrBox.Left + (lWidth * 0.0833), LrBox.Top + (lHeight * 0.0625), LrBox.Right, LrBox.Bottom, _
            LrBox.Left + (lWidth * 0.5), LrBox.Top, LrBox.Left + (lWidth * 0.5), LrBox.Top)
        End If

        If blnChecked Then
            With LrBox
                .Left = lLeft + (lWidth * 0.25)
                .Top = lTop + (lWidth * 0.25)
                .Right = lLeft + (lWidth * 0.9167)
                .Bottom = lTop + (lHeight * 0.6875)
                
                ' Now converts coordinates to printer ratio
                .Left = (.Left - m_lPrnGapX)
                .Top = (.Top - m_lPrnGapY)
                .Right = (.Right - m_lPrnGapX)
                .Bottom = (.Bottom - m_lPrnGapY)
                
                lEllReg = CreateEllipticRgn(.Left, .Top, .Right, .Bottom)
            End With
            
            LhBrush = CreateSolidBrush(lfColor)
            Call FillRgn(LhDC, lEllReg, LhBrush)
            Call DeleteObject(LhBrush)
        End If
        
    ElseIf iType = 2 Then       'box with X in it
        With LrBox
            .Left = lLeft
            .Top = lTop
            .Right = lLeft + lWidth
            .Bottom = lTop + lHeight * 0.75
        
            ' Now converts coordinates to printer ratio
            .Left = (.Left - m_lPrnGapX)
            .Top = (.Top - m_lPrnGapY)
            .Right = (.Right - m_lPrnGapX)
            .Bottom = (.Bottom - m_lPrnGapY)
        End With
        
        If (lbColor <> -1) Then
            LhBrush = CreateSolidBrush(lbColor)
            Call FillRect(LhDC, LrBox, LhBrush)
            Call DeleteObject(LhBrush)
        End If
    
        lhPen = CreatePen(PS_SOLID, lLineWidth, lfColor)
        If (lhPen <> 0) Then
            LhOldPen = SelectObject(LhDC, lhPen)
        End If
        
        If blnSunken Then
            MoveToEx LhDC, LrBox.Left + (lWidth * 0.0833), LrBox.Top + (lHeight * 0.75), LrPos
            LineTo LhDC, LrBox.Left + (lWidth * 0.0833), LrBox.Top + (lHeight * 0.0625)
            LineTo LhDC, LrBox.Left + (lWidth * 0.9167), LrBox.Top + (lHeight * 0.0625)
        End If
        
        With LrBox
            MoveToEx LhDC, .Left, .Top, LrPos
            LineTo LhDC, .Right, .Top
            LineTo LhDC, .Right, .Bottom
            LineTo LhDC, .Left, .Bottom
            LineTo LhDC, .Left, .Top
        End With
        
        If blnChecked Then
            lhPen = CreatePen(PS_SOLID, lLineWidth * 2, lfColor)
            If (lhPen <> 0) Then
                LhOldPen = SelectObject(LhDC, lhPen)
            End If
            MoveToEx LhDC, LrBox.Left + (lWidth * 0.25), LrBox.Top + (lWidth * 0.25), LrPos
            LineTo LhDC, LrBox.Left + (lWidth * 0.833), LrBox.Top + (lHeight * 0.625)
            MoveToEx LhDC, LrBox.Left + (lWidth * 0.25), LrBox.Top + (lHeight * 0.625), LrPos
            LineTo LhDC, LrBox.Left + (lWidth * 0.833), LrBox.Top + (lHeight * 0.1875)
        End If
        
    End If
    
    Exit Sub

ERR_H:
    Me.RaiseErr Err.Number, "DrawCheckBox"

End Sub

Public Property Let PaperSize(iSize As Integer)
    m_PaperSize = iSize
    Select Case m_PaperSize
        Case vbPRPSLetter ' Letter
            If m_bLandScape Then
                m_nDisplayWidth = 1056
                m_nDisplayHeight = 816
            Else
                m_nDisplayWidth = 816
                m_nDisplayHeight = 1056
            End If
        Case vbPRPSA4 ' A4
            If m_bLandScape Then
                m_nDisplayWidth = 1152
                m_nDisplayHeight = 816
            Else
                m_nDisplayWidth = 816
                m_nDisplayHeight = 1152
            End If
        Case vbPRPSLegal ' Legal
            If m_bLandScape Then
                m_nDisplayWidth = 1344
                m_nDisplayHeight = 816
            Else
                m_nDisplayWidth = 816
                m_nDisplayHeight = 1344
            End If
        Case Else ' User
            m_PaperSize = vbPRPSUser
    End Select
End Property

Private Sub Class_Initialize()
    Set m_oElements = New Collection
    m_nDisplayWidth = 816 ' 8.5 inches to pixels
    m_nDisplayHeight = 1056 '11 inches to pixels
    m_iScaleMode = vbInches
    m_PaperSize = vbPRPSLetter
    m_bLandScape = False
    PrinterTray = vbPRBNAuto
    If (Printers.Count = 0) Then
        PrinterPort = "[Unknown]"
    Else
        PrinterPort = Printer.Port
    End If
End Sub

