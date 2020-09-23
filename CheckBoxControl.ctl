VERSION 5.00
Begin VB.UserControl CheckBoxControl 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   195
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   ScaleHeight     =   16
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   13
End
Attribute VB_Name = "CheckBoxControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event DblClick()

Private ucBorderColor As Long
Private ucBackColor As Long
Private ucSunken As Boolean
Private ucDisplay As Integer
Private ucDataField As String
Private ucToolTip As String


Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()

    DrawControl

End Sub

Private Sub DrawControl()

    UserControl.Cls
        
    If ucDisplay = 0 Then
        DrawCheck
    ElseIf ucDisplay = 1 Then
        DrawRadio
    ElseIf ucDisplay = 2 Then
        DrawXBox
    End If
        
    UserControl.MaskPicture = UserControl.Image

End Sub

Private Sub DrawCheck()
Dim StartPt As POINTAPI

    If UserControl.Height < 240 Then UserControl.Height = 240
    
    UserControl.DrawWidth = 1
    UserControl.FillColor = ucBackColor
    UserControl.ForeColor = ucBorderColor
    Rectangle UserControl.hdc, 0, 3, 12, 16
    
    If ucSunken Then
        MoveToEx UserControl.hdc, 1, 15, StartPt
        LineTo UserControl.hdc, 1, 4
        LineTo UserControl.hdc, 11, 4
    End If
    
    UserControl.DrawWidth = 3
    MoveToEx UserControl.hdc, 3, 7, StartPt
    LineTo UserControl.hdc, 6, 13
    LineTo UserControl.hdc, 10, 0

End Sub

Private Sub DrawRadio()
Dim StartPt As POINTAPI

    UserControl.Height = 195
    
    UserControl.DrawWidth = 2
    UserControl.FillColor = ucBackColor
    
    UserControl.ForeColor = ucBorderColor
    Ellipse UserControl.hdc, 0, 0, 13, 13

    If ucSunken Then
        Ellipse UserControl.hdc, 1, 1, 13, 13
    End If

    UserControl.FillColor = ucBorderColor
    Ellipse UserControl.hdc, 4, 4, 9, 9

End Sub

Private Sub DrawXBox()
Dim StartPt As POINTAPI

    UserControl.Height = 195
    
    UserControl.DrawWidth = 1
    UserControl.FillColor = ucBackColor
    Rectangle UserControl.hdc, 0, 0, 12, 12
    
    If ucSunken Then
        MoveToEx UserControl.hdc, 1, 12, StartPt
        LineTo UserControl.hdc, 1, 1
        LineTo UserControl.hdc, 11, 1
    End If

    UserControl.ForeColor = ucBorderColor
    UserControl.DrawWidth = 2
    MoveToEx UserControl.hdc, 2, 2, StartPt
    LineTo UserControl.hdc, 9, 9
    MoveToEx UserControl.hdc, 2, 9, StartPt
    LineTo UserControl.hdc, 9, 2

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Property Get BorderColor() As Long
    BorderColor = ucBorderColor
End Property

Public Property Let BorderColor(ByVal NewValue As Long)
    ucBorderColor = NewValue
    DrawControl
End Property

Public Property Get Sunken() As Boolean
    Sunken = ucSunken
End Property

Public Property Let Sunken(ByVal NewValue As Boolean)
    ucSunken = NewValue
    DrawControl
End Property

Private Sub UserControl_Resize()

    If UserControl.width > 195 Then UserControl.width = 195
    If UserControl.Height > 240 Then UserControl.Height = 240

End Sub

Public Property Get DisplayType() As Integer
    DisplayType = ucDisplay
End Property

Public Property Let DisplayType(ByVal NewValue As Integer)
    ucDisplay = NewValue
    DrawControl
End Property

Public Property Get DataField() As String
    DataField = ucDataField
End Property

Public Property Let DataField(ByVal NewValue As String)
    ucDataField = NewValue
End Property

Public Property Get ToolTipText() As String
    ToolTipText = ucToolTip
End Property

Public Property Let ToolTipText(ByVal NewValue As String)
    ucToolTip = NewValue
End Property

Public Property Get BackColor() As Long
    BackColor = ucBackColor
End Property

Public Property Let BackColor(ByVal NewValue As Long)
    ucBackColor = NewValue
    DrawControl
End Property
