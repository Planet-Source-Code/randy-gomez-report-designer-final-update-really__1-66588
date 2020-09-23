VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FontMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"IFont"
Attribute VB_Ext_KEY = "Member0" ,"IFont"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'////////////////////////////////////////////////////////
'///               Fonts Collection Class
'///                   (FontMap.cls)
'///_____________________________________________________
'/// Font Definition Collection Class. Handles Stores IFont
'/// elements, let user to append font definition, if a
'/// definition already exists returns only the proper index
'/// for the already defined font. The Create method let the
'/// user create HTML tags as well as Log Font handles.
'///_____________________________________________________
'/// Last modification  : Ago/13/2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Created
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit

'local variable to hold collection
Private m_oCol As Collection
Private m_iFontIdx As Integer

Public Property Get ActiveFont() As IFont
    On Error Resume Next
    Set ActiveFont = m_oCol.Item(m_iFontIdx)
End Property

Public Sub Clear()
    With m_oCol
        While (.Count > 0)
            .Remove .Count
        Wend
    End With
End Sub

Public Sub Create(Optional ByVal hDC As Long)
    Dim LoFont As IFont
    Dim LhWnd As Long
    Dim LbDestroy As Boolean
    
    ' Obtains a LOG_FONT handler
    If (hDC = 0) Then
        LhWnd = GetDesktopWindow
        hDC = GetDC(LhWnd)
        LbDestroy = True
    End If
    For Each LoFont In m_oCol
        LoFont.Create hDC
    Next
    If LbDestroy Then
        Call ReleaseDC(LhWnd, hDC)
    End If
End Sub


Public Function Add(FaceName As String, Size As Single, _
    Optional Bold As Boolean, Optional Italic As Boolean, _
    Optional Underline As Boolean, Optional Strikethrough As Boolean, _
    Optional Rotation As Integer) As Integer
    Dim LoNewFont As IFont
    Dim LsKey As String
    ' Builds the key to identify the font
    LsKey = FaceName & Size & (Rotation Mod 360) & _
        Abs(Bold) & Abs(Italic) & Abs(Underline) & Abs(Strikethrough)
    On Error Resume Next
    ' Checks if the requested font already exists
    Set LoNewFont = m_oCol.Item(LsKey)
    If (Err.Number = 0) Then
        ' If already exists, no error is raised, therefore
        ' just obtains the font index
        m_iFontIdx = LoNewFont.Index
        Add = m_iFontIdx
    Else
        ' If error was raised, the font does not exists
        ' Therefore we must add it to the collection
        Set LoNewFont = New IFont
        'set the properties passed into the method
        With LoNewFont
            .FaceName = FaceName
            .Size = Size
            .Rotation = Rotation
            .Bold = Bold
            .Italic = Italic
            .Underline = Underline
            .Strikethrough = Strikethrough
            ' Obtains the current font index
            .Index = (m_oCol.Count + 1)
        End With
        ' Adds the new font to the collection
        m_oCol.Add LoNewFont, LsKey
        ' Returns the font index
        If (Err.Number <> 0) Then
            m_iFontIdx = LoNewFont.Index
            Add = m_iFontIdx
        End If
    End If
    Set LoNewFont = Nothing
End Function

Public Property Get Item(Key As Variant) As IFont
Attribute Item.VB_UserMemId = 0
    Set Item = m_oCol(Key)
End Property



Public Property Get Count() As Long
    Count = m_oCol.Count
End Property


Public Sub Remove(Key As Variant)
    m_oCol.Remove Key
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_oCol.[_NewEnum]
End Property


Friend Sub SelectFont(iIndex As Integer)
    m_iFontIdx = iIndex
End Sub

Private Sub Class_Initialize()
    ' Creates the collection when this class is created
    Set m_oCol = New Collection
    ' Adds almost one valid font
    Me.Add "Arial", 8.5
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set m_oCol = Nothing
End Sub

