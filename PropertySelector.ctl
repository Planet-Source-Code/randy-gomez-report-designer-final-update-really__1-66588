VERSION 5.00
Begin VB.UserControl PropertySelector 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   ScaleHeight     =   172
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   202
End
Attribute VB_Name = "PropertySelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private FocusRECT As Rect
Private MouseYPos As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As Rect) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private ColorRect As Rect
Private ColorBrush As Long

Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2

Private SelMode As Integer
Private ShowTransparentOption As Boolean
Private Colors() As Long                  'array of colors for color picker
Private PrevSelColor As Long
Private NewSelColor As Long               'selected color
Private oldX As Long, oldY As Long      'stores previous mouse location on color picker grid
Private ucShowDialog As Boolean

Private SelLineStyle As Integer
Private SelLineWidth As Integer

Public Event Clicked()

Private Sub UserControl_Click()
    
    RaiseEvent Clicked

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim NewX As Long, NewY As Long
Dim GetX As Long, GetY As Long

    UserControl.DrawStyle = 0
    If SelMode <= 3 Or SelMode = 6 Then
        GetX = CLng(x)
        GetY = CLng(y)
        
        UserControl.ForeColor = UserControl.BackColor
        UserControl.Line (oldX - 1, oldY - 1)-(oldX + 16, oldY + 16), , B
                
        If GetX > 16 And GetX < 186 And GetY > 16 And GetY < 143 Then
            NewX = SnapToProperty(GetX, 22, 16)
            NewY = SnapToProperty(GetY, 22, 16)
            NewSelColor = Colors(Int(NewX / 22), Int(NewY / 22))
            oldX = NewX
            oldY = NewY
        ElseIf GetX > 16 And GetX < 32 And GetY > 148 And GetY < 164 Then
            NewSelColor = -1
        ElseIf GetX > 140 And GetX < 186 And GetY > 148 And GetY < 164 Then
            ucShowDialog = True
            ShowMoreButton True
        Else
            NewSelColor = PrevSelColor
            oldX = -16
            oldY = -16
            ucShowDialog = False
            ShowMoreButton False
        End If
        
        If NewX > 0 And NewY > 0 Then
            UserControl.ForeColor = vbBlack
            UserControl.Line (NewX - 1, NewY - 1)-(NewX + 16, NewY + 16), , B
        End If
    ElseIf SelMode = 4 Then
        GetY = CLng(y)
        FocusRECT.Top = SnapToProperty(GetY, 20, 2)
        FocusRECT.Bottom = FocusRECT.Top + 15
        SelLineStyle = (FocusRECT.Top - 2) / 20
        ShowLineStyles
        DrawFocusRect UserControl.hdc, FocusRECT
        DrawBorder
    ElseIf SelMode = 5 Then
        GetY = CLng(y)
        FocusRECT.Top = SnapToProperty(GetY, 20, 5)
        FocusRECT.Bottom = FocusRECT.Top + 15
        SelLineWidth = (FocusRECT.Top - 5) / 20 + 1
        ShowLineWidths
        DrawFocusRect UserControl.hdc, FocusRECT
        DrawBorder
    End If

End Sub

Private Function SnapToProperty(Coord As Long, Grid As Long, Offset As Long) As Long
Dim Remainder As Long

    Remainder = Int((Coord - Offset) / Grid)
    Remainder = Remainder * Grid
    Remainder = Coord - Remainder

    SnapToProperty = Coord - Remainder + Offset

End Function

Private Sub ShowMoreButton(Hilite As Boolean)

    UserControl.ForeColor = RGB(180, 180, 180)
    UserControl.Line (140, 148)-(186, 164), , BF
    
    UserControl.ForeColor = vbWhite
    UserControl.Line (142, 150)-(184, 162), , BF
    
    UserControl.ForeColor = RGB(100, 100, 100)
    UserControl.Line (140, 148)-(186, 164), , B
    
    ColorRect.Left = 147
    ColorRect.Right = 184
    ColorRect.Top = 150
    ColorRect.Bottom = 164
    
    If Hilite Then
        UserControl.ForeColor = &H8000000D
    Else
        UserControl.ForeColor = vbBlack
    End If
    DrawText UserControl.hdc, "More...", 7, ColorRect, vbNormal

End Sub

Private Sub ShowColors()
Dim i As Integer, j As Integer

    ReDim Colors(7, 5)
    Colors(0, 0) = RGB(255, 255, 255)   'white to black range
    Colors(0, 1) = RGB(215, 215, 215)
    Colors(0, 2) = RGB(185, 185, 185)
    Colors(0, 3) = RGB(145, 145, 145)
    Colors(0, 4) = RGB(95, 95, 95)
    Colors(0, 5) = RGB(0, 0, 0)
    Colors(1, 0) = RGB(255, 195, 195)   'red range
    Colors(1, 1) = RGB(255, 125, 125)
    Colors(1, 2) = RGB(255, 45, 45)
    Colors(1, 3) = RGB(205, 0, 0)
    Colors(1, 4) = RGB(165, 0, 0)
    Colors(1, 5) = RGB(105, 0, 0)
    Colors(2, 0) = RGB(255, 205, 155)   'orange range
    Colors(2, 1) = RGB(255, 155, 115)
    Colors(2, 2) = RGB(255, 125, 75)
    Colors(2, 3) = RGB(195, 95, 45)
    Colors(2, 4) = RGB(125, 65, 0)
    Colors(2, 5) = RGB(95, 45, 0)
    Colors(3, 0) = RGB(255, 255, 205)   'yellow range
    Colors(3, 1) = RGB(255, 255, 155)
    Colors(3, 2) = RGB(255, 255, 0)
    Colors(3, 3) = RGB(195, 195, 0)
    Colors(3, 4) = RGB(155, 155, 0)
    Colors(3, 5) = RGB(105, 105, 0)
    Colors(4, 0) = RGB(205, 255, 205)   'green range
    Colors(4, 1) = RGB(155, 255, 155)
    Colors(4, 2) = RGB(0, 255, 0)
    Colors(4, 3) = RGB(0, 195, 0)
    Colors(4, 4) = RGB(0, 155, 0)
    Colors(4, 5) = RGB(0, 95, 0)
    Colors(5, 0) = RGB(205, 255, 255)   'cyan range
    Colors(5, 1) = RGB(155, 255, 255)
    Colors(5, 2) = RGB(0, 255, 255)
    Colors(5, 3) = RGB(0, 195, 195)
    Colors(5, 4) = RGB(0, 155, 155)
    Colors(5, 5) = RGB(0, 95, 95)
    Colors(6, 0) = RGB(205, 205, 255)   'blue range
    Colors(6, 1) = RGB(155, 155, 255)
    Colors(6, 2) = RGB(0, 0, 255)
    Colors(6, 3) = RGB(0, 0, 195)
    Colors(6, 4) = RGB(0, 0, 155)
    Colors(6, 5) = RGB(0, 0, 95)
    Colors(7, 0) = RGB(255, 205, 255)   'magenta
    Colors(7, 1) = RGB(255, 195, 255)
    Colors(7, 2) = RGB(255, 0, 255)
    Colors(7, 3) = RGB(195, 0, 195)
    Colors(7, 4) = RGB(155, 0, 155)
    Colors(7, 5) = RGB(95, 0, 95)
    
    UserControl.width = 3030
    UserControl.Height = 2600
    UserControl.Cls
    'fill in colors on properties window
    For i = 0 To 7
        For j = 0 To 5
            ColorRect.Left = i * 22 + 16
            ColorRect.Right = ColorRect.Left + 16
            ColorRect.Top = j * 22 + 16
            ColorRect.Bottom = ColorRect.Top + 16
            ColorBrush = CreateSolidBrush(Colors(i, j))
            FillRect UserControl.hdc, ColorRect, ColorBrush
        Next j
    Next i
    
End Sub

Private Sub DrawTransparentOption(Showing As Boolean)

    If Showing Then
        UserControl.Height = 2600
    End If
    
    UserControl.DrawStyle = 0
    UserControl.ForeColor = UserControl.BackColor
    UserControl.Line (2, 147)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2), , BF
    
    If Showing Then
        UserControl.ForeColor = vbBlack
        UserControl.Line (16, 148)-(32, 164), , B
        UserControl.Line (16, 148)-(32, 164)
        UserControl.Line (16, 164)-(32, 148)
        
        ColorRect.Left = 35
        ColorRect.Right = 100
        ColorRect.Top = 149
        ColorRect.Bottom = 180
        
        DrawText UserControl.hdc, "Transparent", 11, ColorRect, vbNormal
    End If
    
End Sub

Private Sub DrawBorder()
Dim BorderRec As Rect

    BorderRec.Right = UserControl.ScaleWidth
    BorderRec.Bottom = UserControl.ScaleHeight

    DrawEdge UserControl.hdc, BorderRec, (BDR_RAISEDINNER Or BDR_RAISEDOUTER), BF_RECT

End Sub

Private Sub ShowLineStyles()
Dim i As Integer
    
    UserControl.width = 1200
    UserControl.Height = 1600
    UserControl.Cls
    UserControl.DrawWidth = 1
    UserControl.ForeColor = vbBlack
    For i = 0 To 4
        UserControl.DrawStyle = i
        UserControl.Line (10, i * 20 + 10)-(UserControl.ScaleWidth - 10, i * 20 + 10)
    Next i

End Sub

Private Sub ShowLineWidths()
Dim i As Integer, j As Integer

    UserControl.width = 700
    UserControl.Height = 2200
    UserControl.Cls
    UserControl.DrawStyle = 0
    UserControl.ForeColor = vbBlack
    ColorRect.Left = 5
    ColorRect.Right = 20
    For i = 0 To 6
        For j = 0 To i
            UserControl.Line (20, i * 20 + 10 + j)-(UserControl.ScaleWidth - 10, i * 20 + 10 + j)
        Next j
        UserControl.DrawWidth = 1
        ColorRect.Top = i * 20 + 5
        ColorRect.Bottom = i * 20 + 50
        DrawText UserControl.hdc, Str(i + 1), 2, ColorRect, vbNormal
    Next i

End Sub

Public Property Get SelectedColor() As Long
    SelectedColor = NewSelColor
End Property

Public Property Let SelectedColor(ByVal ClrValue As Long)
    PrevSelColor = ClrValue
End Property
Public Property Get TransparentOption() As Boolean
    TransparentOption = ShowTransparentOption
End Property
Public Property Let TransparentOption(ByVal TransValue As Boolean)
    ShowTransparentOption = TransValue
End Property

Public Property Get SelectedLineStyle() As Integer
    SelectedLineStyle = SelLineStyle
End Property

Public Property Let SelectedLineStyle(ByVal NewValue As Integer)
    SelLineStyle = NewValue
End Property

Public Property Get SelectedLineWidth() As Integer
    SelectedLineWidth = SelLineWidth
End Property

Public Property Let SelectedLineWidth(ByVal NewValue As Integer)
    SelLineWidth = NewValue
End Property

Public Property Get SelectMode() As Integer
    SelectMode = SelMode
End Property

Public Property Let SelectMode(ByVal NewValue As Integer)
    SelMode = NewValue
    UserControl_Show
End Property

Private Sub UserControl_Show()

    ucShowDialog = False
    If SelMode <= 3 Then
        ShowColors
        DrawTransparentOption ShowTransparentOption
        ShowMoreButton False
    ElseIf SelMode = 4 Then
        ShowLineStyles
        FocusRECT.Left = 5
        FocusRECT.Right = UserControl.ScaleWidth - 5
    ElseIf SelMode = 5 Then
        FocusRECT.Left = 5
        FocusRECT.Right = UserControl.ScaleWidth - 5
        ShowLineWidths
    ElseIf SelMode = 6 Then
        PrevSelColor = -1
        ShowColors
        DrawTransparentOption ShowTransparentOption
        ShowMoreButton False
    End If
    DrawBorder
        
End Sub

Public Property Get ShowDialog() As Boolean
    ShowDialog = ucShowDialog
End Property

Public Property Let ShowDialog(ByVal NewValue As Boolean)
    ucShowDialog = NewValue
End Property
