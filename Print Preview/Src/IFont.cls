VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'////////////////////////////////////////////////////////
'///                Font Definition Class
'///                     (IFont.cls)
'///_____________________________________________________
'/// Font Definition Class. Handles Font definition and
'/// creates a LOG_FONT object as HTML tags for the apropiate
'/// font size.
'///_____________________________________________________
'/// Last modification  : Ago/07/2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Created
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit

Public Bold As Boolean
Public Italic As Boolean
Public Underline As Boolean
Public Strikethrough As Boolean

Private m_sFaceName As String
Private m_nSize As Single
Private m_iRotation As Integer
Private m_hFont As Long
Private m_sOpenTag As String
Private m_sCloseTag As String
Private m_Index As Integer
Private m_iType As Integer

Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_CHARACTER_PRECIS = 1
Private Const CLIP_STROKE_PRECIS = 2
Private Const CLIP_MASK = &HF
Private Const CLIP_LH_ANGLES = 16
Private Const CLIP_TT_ALWAYS = 32
Private Const CLIP_EMBEDDED = 128

Private Const DEFAULT_QUALITY = 0
Private Const DRAFT_QUALITY = 1
Private Const PROOF_QUALITY = 2

Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2

Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const HANGEUL_CHARSET = 129
Private Const CHINESEBIG5_CHARSET = 136
Private Const OEM_CHARSET = 255

' Font Families
'
Private Const FF_DONTCARE = 0    '  Don't care or don't know.
Private Const FF_ROMAN = 16      '  Variable stroke width, serifed.

' Times Roman, Century Schoolbook, etc.
Private Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.

' Helvetica, Swiss, etc.
Private Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.

' Pica, Elite, Courier, etc.
Private Const FF_SCRIPT = 64     '  Cursive, etc.
Private Const FF_DECORATIVE = 80 '  Old English, etc.

' Font Weights
Private Const FW_DONTCARE = 0
Private Const FW_THIN = 100
Private Const FW_EXTRALIGHT = 200
Private Const FW_LIGHT = 300
Private Const FW_NORMAL = 400
Private Const FW_MEDIUM = 500
Private Const FW_SEMIBOLD = 600
Private Const FW_BOLD = 700
Private Const FW_EXTRABOLD = 800
Private Const FW_HEAVY = 900

Private Const FW_ULTRALIGHT = FW_EXTRALIGHT
Private Const FW_REGULAR = FW_NORMAL
Private Const FW_DEMIBOLD = FW_SEMIBOLD
Private Const FW_ULTRABOLD = FW_EXTRABOLD
Private Const FW_BLACK = FW_HEAVY

Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_STRING_PRECIS = 1
Private Const OUT_CHARACTER_PRECIS = 2
Private Const OUT_STROKE_PRECIS = 3
Private Const OUT_TT_PRECIS = 4
Private Const OUT_DEVICE_PRECIS = 5
Private Const OUT_RASTER_PRECIS = 6
Private Const OUT_TT_ONLY_PRECIS = 7
Private Const OUT_OUTLINE_PRECIS = 8

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Type OLECOLOR
    RedOrSys As Byte
    Green As Byte
    Blue As Byte
    Type As Byte
End Type


Public Function HTMLColor(lColor As Long) As String
    Dim LsColor As String
    Dim LsHexVal As String
    ' Red value
    LsHexVal = Hex(GetR(lColor))
    If (Len(LsHexVal) = 1) Then LsHexVal = "0" & LsHexVal
    HTMLColor = """#" & LsHexVal
    ' Green value
    LsHexVal = Hex(GetG(lColor))
    If (Len(LsHexVal) = 1) Then LsHexVal = "0" & LsHexVal
    HTMLColor = HTMLColor & LsHexVal
    ' Blue value
    LsHexVal = Hex(GetB(lColor))
    If (Len(LsHexVal) = 1) Then LsHexVal = "0" & LsHexVal
    HTMLColor = HTMLColor & LsHexVal & """"
End Function


Friend Property Let Index(iIndex As Integer)
    m_Index = iIndex
End Property

Public Property Get Index() As Integer
    Index = m_Index
End Property

Private Sub prvDestroyFont()
    If (m_hFont <> 0) Then
        Call DeleteObject(m_hFont)
        m_hFont = 0
    End If
End Sub

Public Property Get Rotation() As Integer
    Rotation = m_iRotation
End Property

Public Property Let Rotation(iRot As Integer)
    m_iRotation = (iRot Mod 360)
End Property

Public Property Let Size(nSize As Single)
    m_nSize = Abs(nSize)
    If (m_nSize = 0) Then
        m_nSize = 8.5
    End If
End Property

Public Property Get Size() As Single
    Size = m_nSize
End Property

Public Property Let FaceName(sName As String)
    If Len(sName) = 0 Then
        m_sFaceName = "Arial"
    Else
        m_sFaceName = sName
    End If
End Property

Public Property Get FaceName() As String
    FaceName = m_sFaceName
End Property

Public Property Get Handle() As Long
    Handle = m_hFont
End Property

Private Function prvCreateLogFont(hDC As Long) As Long
    Dim LrFont As LOGFONT
    Dim LaTempArray() As Byte ' byte array to hold the fontname
    Dim LnIdx As Integer
    prvDestroyFont
    With LrFont
        ' All but two properties are very straight-forward,
        ' even with rotation, and map directly.
        .lfHeight = -(Size * GetDeviceCaps(hDC, LOGPIXELSY)) / 72
        .lfWidth = 0
        .lfEscapement = (m_iRotation * 10)
        .lfOrientation = .lfEscapement
        
'Note from RG (04/08/2006) - changed the bold settings below to lighter values

        If Bold Then ' if true set the weight to the appropriate value
            .lfWeight = FW_BOLD ' FW_EXTRABOLD ' FW_BOLD
        Else
            .lfWeight = FW_NORMAL ' FW_MEDIUM ' FW_NORMAL
        End If
        .lfItalic = Abs(Italic)
        .lfUnderline = Abs(Underline)
        .lfStrikeOut = Abs(Strikethrough)
        .lfClipPrecision = CLIP_DEFAULT_PRECIS
        .lfQuality = PROOF_QUALITY
        .lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
        ' OEM fonts can't rotate, and we must force
        ' substitution with something ANSI.
        .lfCharSet = DEFAULT_CHARSET
        If .lfCharSet = OEM_CHARSET Then
            If (m_iRotation <> 0) Then
                .lfCharSet = ANSI_CHARSET
            End If
        End If
        ' Only TrueType fonts can rotate, so we must
        ' specify TT-only if angle is not zero.
        If (m_iRotation <> 0) Then
            .lfOutPrecision = OUT_TT_ONLY_PRECIS
        Else
            .lfOutPrecision = OUT_DEFAULT_PRECIS
        End If
        ' converts the font name from unicode
        LaTempArray = StrConv(m_sFaceName & vbNullChar, vbFromUnicode)
    '   ***** puts the font name into the byte array for face name
        For LnIdx = 0 To UBound(LaTempArray)
            .lfFaceName(LnIdx) = LaTempArray(LnIdx)
        Next LnIdx
    End With
    prvCreateLogFont = CreateFontIndirect(LrFont)
End Function

Public Property Get OpenTag(Optional lColor As Long) As String
    OpenTag = m_sOpenTag & HTMLColor(lColor) & ">"
End Property

Private Function GetR(Color As Long) As Byte
    CopyMemory GetR, WinColor(Color), 1
End Function

Private Function GetG(Color As Long) As Byte
    CopyMemory GetG, ByVal VarPtr(WinColor(Color)) + 1, 1
End Function

Private Function GetB(Color As Long) As Byte
    CopyMemory GetB, ByVal VarPtr(WinColor(Color)) + 2, 1
End Function

Private Function WinColor(VBColor As Long) As Long
    Dim SysClr As OLECOLOR

    CopyMemory SysClr, VBColor, Len(SysClr)

    If SysClr.Type = &H80 Then 'Es ist eine Systemfarbe
        'SysClr.RedOrSys ist die Nummer der Systemfarbe
        WinColor = GetSysColor(SysClr.RedOrSys)
    Else 'Es ist keine Systemfarbe
        'SysClr.RedOrSys ist lediglich der Rotanteil
        WinColor = VBColor
    End If
End Function

Public Property Get CloseTag() As String
    CloseTag = m_sCloseTag
End Property

Friend Sub Create(Optional ByVal hDC As Long)
    Dim LnRelSize As Integer
    Dim LhWnd As Long
    Dim LbDestroy As Boolean
    
    ' Obtains a LOG_FONT handler
    If (hDC = 0) Then
        LhWnd = GetDesktopWindow
        hDC = GetDC(LhWnd)
        LbDestroy = True
    End If
    m_hFont = prvCreateLogFont(hDC)
    If LbDestroy Then
        Call ReleaseDC(LhWnd, hDC)
    End If
    ' Calculates HTML font size
    If (m_nSize <= 9) Then
        LnRelSize = 1
    ElseIf (m_nSize <= 11) Then
        LnRelSize = 2
    ElseIf (m_nSize <= 13) Then
        LnRelSize = 3
    ElseIf (m_nSize <= 16) Then
        LnRelSize = 4
    ElseIf (m_nSize <= 21) Then
        LnRelSize = 5
    ElseIf (m_nSize <= 30) Then
        LnRelSize = 6
    Else
        LnRelSize = 7
    End If
    ' Creates html tags...
    m_sOpenTag = ""
    m_sCloseTag = ""
    If Bold Then
        m_sOpenTag = "<B>"
        m_sCloseTag = "</B>"
    End If
    If Italic Then
        m_sOpenTag = m_sOpenTag & "<I>"
        m_sCloseTag = "</I>" & m_sCloseTag
    End If
    m_sOpenTag = m_sOpenTag & "<FONT FACE=""" & m_sFaceName & _
        """ SIZE=""" & LnRelSize & """ COLOR="
    m_sCloseTag = "</FONT>" & m_sCloseTag
'///////////// PRESERVED CODE /////////////
'    m_iType = iType
'    Select Case m_iType
'        Case 1
'            m_hFont = prvCreateLogFont
'        Case 3
'            Dim LnRelSize As Integer
'
'            If (Size <= 9) Then
'                LnRelSize = 1
'            ElseIf (Size <= 11) Then
'                LnRelSize = 2
'            ElseIf (Size <= 13) Then
'                LnRelSize = 3
'            ElseIf (Size <= 16) Then
'                LnRelSize = 4
'            ElseIf (Size <= 21) Then
'                LnRelSize = 5
'            ElseIf (Size <= 30) Then
'                LnRelSize = 6
'            Else
'                LnRelSize = 7
'            End If
'            m_sOpenTag = "<FONT face=" & FaceName & _
'                " Size = " & LnRelSize & " color=#"
'            m_sCloseTag = "</FONT>"
'            If Bold Then
'                m_sOpenTag = m_sOpenTag & "<B>"
'                m_sCloseTag = m_sCloseTag & "</B>"
'            End If
'            If Italic Then
'                m_sOpenTag = m_sOpenTag & "<I>"
'                m_sCloseTag = m_sCloseTag & "</I>"
'            End If
'        Case 4
'            m_sOpenTag = "\f0" & "\fs" & "\b"
'    End Select
End Sub

Private Sub Class_Terminate()
    prvDestroyFont
End Sub


