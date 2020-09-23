VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'////////////////////////////////////////////////////////
'///                  Image List Class
'///                 (clsImageList.cls)
'///_____________________________________________________
'/// Class to handle a CommonCtl Image List.
'///_____________________________________________________
'/// Last modification  : Ago/07/2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Created
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Modified by Leontti A. Ramos M. (leontti@leontti.net)
'/// based in code from Steve McMahon (vbaccelerator.com)
'////////////////////////////////////////////////////////
Option Explicit

Private m_ILDMonoHDC As Long
Private m_ILDMonoHBMP As Long
Private m_ILDMonoHBMPOld As Long

Private m_ILDColorHDC As Long
Private m_ILDColorHBMP As Long
Private m_ILDColorHBMPOld As Long

Public Enum eilIconSize
  Size16 = 16
  Size32 = 32
End Enum

Public Enum eilIconState
  Normal = 0
  Disabled = 1
End Enum
'
Private m_hIml As Long

Private Const ILC_MASK = &H1
Private Const ILC_COLOR = &H0
Private Const ILC_COLORDDB = &H0
Private Const ILC_COLOR4 = &H4
Private Const ILC_COLOR8 = &H8
Private Const ILC_COLOR16 = &H10
Private Const ILC_COLOR24 = &H18
Private Const ILC_COLOR32 = &H20
 
Private Const CLR_NONE = -1
Private Const CLR_DEFAULT = -16777216
Private Const CLR_HILIGHT = -16777216

Public Enum ImageTypes
  IMAGE_BITMAP = 0
  IMAGE_ICON = 1
  IMAGE_CURSOR = 2
  'IMAGE_ENHMETAFILE = 3
End Enum
 
Private Type IMAGEINFO
    hBitmapImage As Long
    hBitmapMask As Long
    cPlanes As Long
    cBitsPerPixel As Long
    rcImage As RECT
End Type

Private Declare Function ImageList_SetBkColor Lib "COMCTL32" (ByVal hImageList As Long, ByVal clrBk As Long) As Long
Private Declare Function ImageList_GetBkColor Lib "COMCTL32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_ReplaceIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_Draw Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hDCDest As Long, ByVal X As Long, ByVal Y As Long, ByVal Flags As Long) As Long
Private Declare Function ImageList_DrawEx Lib "COMCTL32" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_Convert Lib "COMCTL32" Alias "ImageList_Draw" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hDCDest As Long, ByVal X As Long, ByVal Y As Long, ByVal Flags As Long) As Long
Private Declare Function ImageList_Create Lib "COMCTL32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_AddMasked Lib "COMCTL32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_Replace Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
Private Declare Function ImageList_Add Lib "COMCTL32" (ByVal hImageList As Long, ByVal hbmImage As Long, hbmMask As Long) As Long
Private Declare Function ImageList_Remove Lib "COMCTL32" (ByVal hImageList As Long, ImgIndex As Long) As Long
Private Declare Function ImageList_GetImageInfo Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, pImageInfo As IMAGEINFO) As Long
Private Declare Function ImageList_AddIcon Lib "COMCTL32" (ByVal hIml As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_GetIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, hbmMask As Long) As Long
Private Declare Function ImageList_SetImageCount Lib "COMCTL32" (ByVal hImageList As Long, uNewCount As Long)
Private Declare Function ImageList_GetImageCount Lib "COMCTL32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_Destroy Lib "COMCTL32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Private Declare Function ImageList_SetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Private Declare Function ImageList_LoadImage Lib "COMCTL32" Alias "ImageList_LoadImageA" (ByVal hInst As Long, ByVal lpbmp As String, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long)

Friend Sub RaiseErr(ByVal lErrNum As RSErrorCode, Optional sRoutineName As String, _
    Optional sDescription As String)
    RaiseError lErrNum, TypeName(Me), sRoutineName, sDescription, Erl
End Sub


Public Function Create(ByVal hDC As Long, _
    ByVal ImgSize As eilIconSize) As Boolean
    ' Do we already have an image list?
    Destroy
    'Create the Imagelist
    m_hIml = ImageList_Create(ImgSize, ImgSize, ILC_MASK, 4, 4)
    If ((m_hIml <> 0) And (m_hIml <> -1)) Then
        prvMakeWorkDCS hDC, ImgSize
    Else
        m_hIml = 0
    End If
End Function

Public Sub Destroy()
    If (hIml <> 0) Then
        ImageList_Destroy hIml
        prvClearUpWorkDCS
        m_hIml = 0
    End If
End Sub

Public Sub DrawImage(ByVal iImgIndex As Long, ByVal hDC As Long, _
        ByVal xPixels As Integer, ByVal yPixels As Integer, _
        Optional ByVal bSelected = False, Optional ByVal bDisabled = False)
    Dim hIcon As Long
    Dim lFlags As Long
    Dim lhIml As Long
   
    lhIml = hIml
    lFlags = ILD_TRANSPARENT
    If (bSelected) Or (bDisabled) Then
        lFlags = lFlags Or ILD_SELECTED
    End If
    If (bDisabled) Then
        lFlags = lFlags Or ILD_SELECTED
        ImageList_DrawEx _
              lhIml, _
              iImgIndex, _
              hDC, _
              xPixels, yPixels, 0, 0, _
              CLR_NONE, GetSysColor(COLOR_WINDOW), _
              lFlags
    Else
        ImageList_Draw _
            lhIml, _
            iImgIndex, _
            hDC, _
            xPixels, _
            yPixels, _
            lFlags
    End If
End Sub

Public Property Get IconSize() As Integer
Dim ImgHeight As Long, ImgWidth As Long
    ImageList_GetIconSize hIml, ImgHeight, ImgWidth
    IconSize = ImgHeight
End Property

Public Property Get ImageCount() As Integer
    ImageCount = ImageList_GetImageCount(hIml)
End Property

Public Sub RemoveImage(ByVal Index As Integer)
    ImageList_Remove hIml, ByVal Index
End Sub

Public Sub Clear()
    ImageList_Remove hIml, -1
End Sub

Public Function AddFromPictureBox(ByVal hDC As Long, pic As Object, _
    Optional ByVal LeftPixels As Long = 0, Optional ByVal TopPixels As Long = 0, _
    Optional ByVal lBackColor As OLE_COLOR = -1) As Long
    Dim LhDC As Long
    Dim lhBmp As Long, lhBmpOld As Long
    Dim tBm As BITMAP
    Dim lAColor As Long
    Dim lW As Long, lH As Long
    Dim hBrush As Long
    Dim tR As RECT
    Dim lR As Long
    Dim lIconSize As Long
    Dim lBPixel As Long
    
    lIconSize = IconSize
    ' Create a DC to hold the bitmap to transfer into the image list:
    LhDC = CreateCompatibleDC(hDC)
    If (LhDC <> 0) Then
        ' Create a bitmap compatible with the current device
        ' to copy the picture into:
        'GetObjectAPI pic.Picture.Handle, LenB(tBm), tBm
        'tBm.bmBits = 0
        'tBm.bmWidth = lIconSize
        'tBm.bmHeight = lIconSize
        'lHbmp = CreateBitmapIndirect(tBm)
        lhBmp = CreateCompatibleBitmap(hDC, lIconSize, lIconSize)
        If (lhBmp <> 0) Then
            ' Get the backcolor to use:
            If (lBackColor = -1) Then
                ' None specified, use the colour at 0,0:
                lBackColor = GetPixel(pic.hDC, 0, 0)
            Else
                ' Try to get the specified backcolor:
                If OleTranslateColor(lBackColor, 0, lAColor) Then
                    ' Failed- use default of silver
                    lBackColor = &HC0C0C0
                Else
                    ' Set to GDI version of OLE Color
                    lBackColor = lAColor
                End If
            End If
            ' Select the bitmap into the DC
            lhBmpOld = SelectObject(LhDC, lhBmp)
            ' Clear the background:
            hBrush = CreateSolidBrush(lBackColor)
            tR.Right = lIconSize: tR.Bottom = lIconSize
            FillRect LhDC, tR, hBrush
            DeleteObject hBrush
            ' Get the source picture's dimension:
            GetObjectAPI pic.Picture.Handle, LenB(tBm), tBm
            lW = 16
            lH = 16
            If (lW + LeftPixels > tBm.bmWidth) Then
                lW = tBm.bmWidth - LeftPixels
            End If
            If (lH + TopPixels > tBm.bmHeight) Then
                lH = tBm.bmHeight - TopPixels
            End If
            If (lW > 0) And (lH > 0) Then
                ' Blt from the picture into the bitmap:
                lR = BitBlt(LhDC, 0, 0, lW, lH, hDC, LeftPixels, TopPixels, SRCCOPY)
                Debug.Assert (lR <> 0)
            End If
            ' We now have the image in the bitmap, so select it out of the DC:
            SelectObject LhDC, lhBmpOld
            ' And add it to the image list:
            lR = ImageList_AddMasked(hIml, lhBmp, lBackColor)
            Debug.Assert (lR <> -1)
            AddFromPictureBox = lR
            DeleteObject lhBmp
        End If
        ' Clear up the DC:
        DeleteObject LhDC
    End If
End Function

Public Property Get hIml() As Long
    hIml = m_hIml
End Property

Private Sub prvMakeWorkDCS(ByVal lHDCBasis As Long, ByVal lIconSize As Long)
    m_ILDMonoHDC = CreateCompatibleDC(0)
    If (m_ILDMonoHDC <> 0) Then
        m_ILDMonoHBMP = CreateCompatibleBitmap(m_ILDMonoHDC, lIconSize, lIconSize * 3)
        If (m_ILDMonoHBMP <> 0) Then
            m_ILDMonoHBMPOld = SelectObject(m_ILDMonoHDC, m_ILDMonoHBMP)
        End If
    End If
    m_ILDColorHDC = CreateCompatibleDC(lHDCBasis)
    If (m_ILDColorHDC <> 0) Then
        m_ILDColorHBMP = CreateCompatibleBitmap(lHDCBasis, lIconSize, lIconSize * 2)
        If (m_ILDColorHBMP <> 0) Then
            m_ILDColorHBMPOld = SelectObject(m_ILDColorHDC, m_ILDColorHBMP)
        End If
    End If
End Sub

Private Sub prvClearUpWorkDCS()
    If (m_ILDMonoHDC <> 0) Then
        If (m_ILDMonoHBMP <> 0) Then
            SelectObject m_ILDMonoHDC, m_ILDMonoHBMPOld
            DeleteObject m_ILDMonoHBMP
        End If
        DeleteObject m_ILDMonoHDC
    End If
    If (m_ILDColorHDC <> 0) Then
        If (m_ILDColorHBMP <> 0) Then
            SelectObject m_ILDColorHDC, m_ILDColorHBMPOld
            DeleteObject m_ILDColorHBMP
        End If
        DeleteObject m_ILDColorHDC
    End If
End Sub

Private Sub pImageListDrawIconDisabled(ByVal LhDC As Long, _
        ByVal hIml As Long, ByVal iIconIndex As Long, _
        ByVal lX As Long, ByVal lY As Long, ByVal lSize As Long)
    Dim tR As RECT
    Dim hBrush As Long
    Dim lStyle As Long

    ' Firstly, create the mask & image:
    ' Draw the image into the top square of the mono DC:
    BitBlt m_ILDMonoHDC, 0, 0, lSize, lSize * 3, m_ILDMonoHDC, 0, 0, WHITENESS
    lStyle = ILD_IMAGE
    ImageList_Draw hIml, iIconIndex, m_ILDMonoHDC, 0, 0, lStyle
    ' Draw the Mask into the second square:
    lStyle = ILD_MASK
    ImageList_Draw hIml, iIconIndex, m_ILDMonoHDC, 0, lSize, lStyle
    ' Or the mask & mono image together:
    BitBlt m_ILDMonoHDC, 0, 0, lSize, lSize, m_ILDMonoHDC, 0, lSize, SRCPAINT
    ' Invert the thing:
   'BitBlt m_ILDMonoHDC, 0, lSize * 2, lSize, lSize, m_ILDMonoHDC, 0, 0, WHITENESS
    BitBlt m_ILDMonoHDC, 0, lSize * 2, lSize, lSize, m_ILDMonoHDC, 0, lSize, SRCINVERT

    ' Now create white & button shadow copies of it:
    BitBlt m_ILDColorHDC, 0, 0, lSize, lSize, m_ILDMonoHDC, 0, lSize * 2, SRCCOPY
    hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNSHADOW))
    tR.Left = 0
    tR.Right = lSize
    tR.Top = lSize
    tR.Bottom = lSize * 2
    FillRect m_ILDColorHDC, tR, hBrush
    DeleteObject hBrush
    BitBlt m_ILDColorHDC, 0, lSize, lSize, lSize, m_ILDMonoHDC, 0, lSize * 2, SRCAND
    BitBlt m_ILDColorHDC, 0, lSize, lSize, lSize, m_ILDMonoHDC, 0, lSize, SRCPAINT
    ' Finally, we blit the disabled verson to the DC:
    ' Draw white version, offset by 1 pixel in x & y:
    BitBlt LhDC, lX + 1, lY + 1, lSize - 1, lSize - 1, m_ILDColorHDC, 0, 0, SRCPAINT
    ' Draw mask for dark version:
    BitBlt LhDC, lX, lY, lSize, lSize, m_ILDColorHDC, 0, 0, SRCPAINT
    ' Finally draw the button shadow version:
    BitBlt LhDC, lX, lY, lSize, lSize, m_ILDColorHDC, 0, lSize, SRCAND
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub
