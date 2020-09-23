Attribute VB_Name = "modMain"
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, _
ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long


Public GridSpace As Long

Public Sub ShowGrid()
Dim i As Integer, j As Integer, ptX As Single, ptY As Single
Dim SectionWidth As Long, SectionHeight As Long
Dim a As Integer

    For a = 0 To 2
        SectionWidth = frmMain.picSection(a).Width * 1440 / Screen.TwipsPerPixelX
        SectionHeight = frmMain.picSection(a).Height * 1440 / Screen.TwipsPerPixelY
        frmMain.picSection(a).Cls
        ptX = 0
        For i = 0 To Int(SectionWidth / GridSpace)
            ptX = ptX + GridSpace
            ptY = 0
            For j = 0 To Int(SectionHeight / GridSpace)
                ptY = ptY + GridSpace
                SetPixelV frmMain.picSection(a).hdc, ptX, ptY, vbBlack
            Next j
        Next i
    Next a

End Sub

