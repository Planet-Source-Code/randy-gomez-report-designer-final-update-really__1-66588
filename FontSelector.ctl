VERSION 5.00
Begin VB.UserControl FontSelector 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2970
   ScaleHeight     =   330
   ScaleWidth      =   2970
   Begin VB.TextBox txtDisplayFont 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   2570
   End
   Begin VB.ComboBox cboFontList 
      Height          =   315
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   2955
   End
End
Attribute VB_Name = "FontSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim selFont As String

Private OldFontName As String
Private ClickedOnce As Boolean
Public Event FontSelected()

Private Sub cmdCancel_Click()

    If ClickedOnce Then
        txtDisplayFont.FontName = OldFontName
        txtDisplayFont.FontBold = False
        txtDisplayFont.FontItalic = False
        txtDisplayFont.FontSize = 10
        txtDisplayFont = " " & OldFontName
    End If
    
End Sub

Private Sub cmdDropList_Click()

    cboFontList = txtDisplayFont.FontName
    ClickedOnce = False

End Sub

Private Sub lstFontList_Click()

    If Not ClickedOnce Then
        OldFontName = txtDisplayFont.FontName
        ClickedOnce = True
    End If

    txtDisplayFont.FontName = cboFontList
    txtDisplayFont.FontBold = False
    txtDisplayFont.FontItalic = False
    txtDisplayFont.FontSize = 10
    txtDisplayFont = " " & cboFontList

End Sub

Private Sub cboFontList_Click()

    selFont = cboFontList
    RaiseEvent FontSelected
    txtDisplayFont.FontName = cboFontList
    txtDisplayFont.FontBold = False
    txtDisplayFont.FontItalic = False
    txtDisplayFont.FontSize = 10
    txtDisplayFont = " " & cboFontList

End Sub

Private Sub cboFontList_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        selFont = cboFontList.List(cboFontList.ListIndex)
        RaiseEvent FontSelected
    ElseIf KeyCode = vbKeyEscape Then
        CancelSelect
    End If

End Sub

Private Sub CancelSelect()

    If ClickedOnce Then
        txtDisplayFont.FontName = OldFontName
        txtDisplayFont.FontBold = False
        txtDisplayFont.FontItalic = False
        txtDisplayFont.FontSize = 10
        txtDisplayFont = " " & OldFontName
    End If
    
End Sub
Private Sub UserControl_ExitFocus()

    ClickedOnce = False

End Sub

Public Property Get FontFace() As String

    FontFace = selFont

End Property

Public Property Let FontFace(ByVal FaceValue As String)

    selFont = FaceValue
    txtDisplayFont.FontName = FaceValue
    txtDisplayFont = " " & FaceValue

End Property

Private Sub UserControl_Initialize()
Dim i As Integer
    
    For i = 0 To Screen.FontCount - 1
        cboFontList.AddItem Screen.Fonts(i)
    Next i
    txtDisplayFont.FontName = "Arial"
    txtDisplayFont.text = " Arial"
    
End Sub

Private Sub UserControl_Resize()

    cboFontList.width = UserControl.width
    txtDisplayFont.width = cboFontList.width - 385

End Sub
