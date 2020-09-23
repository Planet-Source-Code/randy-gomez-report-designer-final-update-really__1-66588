VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Page"
Attribute VB_Ext_KEY = "Member0" ,"Page"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'////////////////////////////////////////////////////////
'///               Pages Collection Class
'///                    (Pages.cls)
'///_____________________________________________________
'/// Report Pages Collection Class. Handles pages manipulation
'///_____________________________________________________
'/// Last modification  : Ago/07/2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Created
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit

Private m_oRange As clsRangeSelector
Private m_oFontMap As FontMap
Private m_oCol As Collection
Private m_lSelected As Long
Private m_lParentPtr As Long
'Private m_lPixelsPerUnit As Long
'Private m_iLineRatio As Integer

Private m_iScaleMode As ScaleModeConstants
Private m_iMFactor As Single
' Paper size properties
Private m_PaperSize As Integer
Private m_nDisplayWidth As Single
Private m_nDisplayHeight As Single
Private m_nScaledWidth As Single
Private m_nScaledHeight As Single
Private m_bLandScape As Boolean

Public Event PageAdded(iCount As Long)
Public Event Cleared()
Public Event PageSelected(ByVal lPage As Long, lCount As Long)

Private Const TWIPS_PER_INCHE = 1440
Private Const TWIPS_PER_CENTIMETER = 567
Private Const TWIPS_PER_POINT = 72
Private Const TWIPS_PER_PIXEL = 15
Private Const TWIPS_PER_TWIP = 1

Public Enum ARDestination
    arScreen = 0
    arPrinter
    arText
    arCSV
    arHTML
    arXML
    arRTF
    arPDF
    arBMP = 101
    arJPG = 102
    arCustom = 255
End Enum

Friend Function GetTempFile(Optional sPrefix As String) As String
    Dim LsTmpName As String
    Dim LsTempPath As String
    Dim LnRet As Long

    LsTempPath = GetTmpDir
    LsTmpName = String(512, 0)
    LnRet = GetTempFileName(LsTempPath, sPrefix, 0, LsTmpName)
    If (LnRet = 0) Then
        Me.RaiseErr ecErrorImportingData, "GetTempFile", "Error getting temporary name."
    Else
        GetTempFile = Left(LsTmpName, InStr(LsTmpName, vbNullChar) - 1)
    End If
End Function

Friend Function GetTmpDir() As String
    Dim LsBuffer As String
    Dim LlResult As Long

    LsBuffer = String(255, 0)
    LlResult = GetTempPath(Len(LsBuffer), LsBuffer)
    GetTmpDir = Left(LsBuffer, LlResult)
    If Right(GetTmpDir, 1) <> "\" Then GetTmpDir = GetTmpDir & "\"
End Function






Friend Property Set Parent(oParent As Preview)
    m_lParentPtr = ObjPtr(oParent)
End Property

Friend Property Get Parent() As Preview
    Set Parent = ObjFromPtr(m_lParentPtr)
End Property


Public Property Get Width() As Single
    Width = m_nScaledWidth
End Property

Public Property Get Height() As Single
    Height = m_nScaledHeight
End Property


Public Property Get PaperSize() As Integer
    PaperSize = m_PaperSize
End Property

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

Private Sub prvSetPrinterPage(oPage As Page)
    On Error GoTo ERR_TRAP
    Printer.ScaleMode = vbPixels
    Select Case oPage.PaperSize
        Case vbPRPSLetter, vbPRPSLegal, vbPRPSA4
            Printer.PaperSize = m_PaperSize
        Case Else
            With Printer
                .PaperSize = vbPRPSUser
                .Width = oPage.DisplayWidth
                .Height = oPage.DisplayHeight
                .ScaleWidth = oPage.DisplayWidth
                .ScaleHeight = oPage.DisplayHeight
            End With
    End Select
    Printer.ScaleMode = ScaleMode
    Printer.Orientation = IIf(m_bLandScape, vbPRORLandscape, vbPRORPortrait)
    Exit Sub
ERR_TRAP:
    If (Err.Number = 380) Then Resume Next
    Me.RaiseErr Err.Number, "prvSetPrinterPage"
End Sub

Public Property Get FontMap() As FontMap
    Set FontMap = m_oFontMap
End Property

Public Function SetFont(FaceName As String, Size As Single, _
    Optional Bold As Boolean, Optional Italic As Boolean, _
    Optional Underline As Boolean, Optional Strikethrough As Boolean, _
    Optional Rotation As Integer) As Integer
    SetFont = m_oFontMap.Add(FaceName, Size, Bold, Italic, _
        Underline, Strikethrough, Rotation)
End Function

Public Property Get ActivePage() As Page
    On Error Resume Next
    Set ActivePage = m_oCol(m_lSelected)
End Property

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

Public Property Let ScaleMode(iMode As ScaleModeConstants)
    If (m_oCol.Count > 0) Then
        If (m_iScaleMode <> iMode) Then
            Me.RaiseErr ecInvalidPropertyUsage, "ScaleMode[Let]", "Could not change the ScaleMode property when pages already exists."
        End If
    Else
        Select Case iMode
            Case vbInches, vbCentimeters, vbPixels, vbTwips, vbPoints
                m_iScaleMode = iMode
                Select Case m_iScaleMode
                    Case vbCentimeters
                        m_iMFactor = (TWIPS_PER_CENTIMETER / TWIPS_PER_PIXEL)
                    Case vbPoints
                        m_iMFactor = (TWIPS_PER_POINT / TWIPS_PER_PIXEL)
                    Case vbTwips
                        m_iMFactor = (TWIPS_PER_TWIP / TWIPS_PER_PIXEL)
                    Case vbPixels
                        m_iMFactor = 1 '(TWIPS_PER_PIXEL / TWIPS_PER_PIXEL)
                    Case Else 'vbInches
                        m_iMFactor = (TWIPS_PER_INCHE / TWIPS_PER_PIXEL)
                End Select
            Case Else
                ' No accepted ScaleMode
        End Select
    End If
End Property

Public Property Get ScaleMode() As ScaleModeConstants
    ScaleMode = m_iScaleMode
End Property


Public Function Add() As Page
    'create a new object
    Dim LoNewItem As Page
    
    Set LoNewItem = New Page
    'set the properties passed into the method
    With LoNewItem
        .Index = (m_oCol.Count + 1)
        .SetPaperProps m_PaperSize, m_iScaleMode, m_iMFactor, _
            m_nDisplayWidth, m_nDisplayHeight, m_bLandScape
        .Height = m_nScaledHeight
        .Width = m_nScaledWidth
        Set .Parent = Me
        Set .Range = m_oRange
    End With
    m_oCol.Add LoNewItem
    'return the object created
    Set Add = LoNewItem
    Set LoNewItem = Nothing
    m_lSelected = m_oCol.Count
    RaiseEvent PageAdded(m_lSelected)
End Function

Public Sub Clear()
    With m_oCol
        While (.Count > 0)
            .Remove .Count
        Wend
    End With
    RaiseEvent Cleared
End Sub


Public Property Get Item(Key As Variant) As Page
Attribute Item.VB_UserMemId = 0
    Set Item = m_oCol(Key)
End Property



Public Property Get Count() As Long
    Count = m_oCol.Count
End Property


Friend Sub NavigateTo(ByVal Index As Integer)
    Dim LbRaise As Boolean
    If (m_oCol.Count > 0) Then
        Select Case Index
            Case 1 ' First page
                m_lSelected = 1
                LbRaise = True
            Case 2 ' Previous page
                If (m_lSelected > 1) Then
                    m_lSelected = m_lSelected - 1
                    LbRaise = True
                End If
            Case 4 'Next page
                If (m_lSelected < m_oCol.Count) Then
                    m_lSelected = m_lSelected + 1
                    LbRaise = True
                End If
            Case 5 ' Last page
                m_lSelected = m_oCol.Count
                LbRaise = True
        End Select
        If LbRaise Then
            RaiseEvent PageSelected(m_lSelected, m_oCol.Count)
        End If
    End If
End Sub


Public Sub Remove(Key As Variant)
    m_oCol.Remove Key
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_oCol.[_NewEnum]
End Property


Public Property Get Selection() As String
    Selection = m_oRange.SelectionString
End Property

Public Property Let Selection(sSelect As String)
    m_oRange.SelectionString = sSelect
    Parent.ShowPageSelection m_oRange.SelectionString
End Property


Public Function SelectPage(Index As Long) As Long
    If ((Index > 0) And (Index <= m_oCol.Count)) Then
        m_lSelected = Index
        SelectPage = m_lSelected
        RaiseEvent PageSelected(m_lSelected, m_oCol.Count)
    Else
'        Me.RaiseErr 5, "SelectItem"
    End If
End Function

Friend Sub RaiseErr(ByVal lErrNum As RSErrorCode, Optional sRoutineName As String, _
    Optional sDescription As String)
    RaiseError lErrNum, TypeName(Me), sRoutineName, sDescription, Erl
End Sub


Friend Function TestEnable(ByVal Index As Integer) As Boolean
    If (m_oCol.Count = 0) Then
        TestEnable = False
    Else
        Select Case Index
            Case 1, 2 ' First page, Previous page
                TestEnable = (m_lSelected > 1)
            Case 3 ' Goto page
                TestEnable = (m_oCol.Count > 1)
            Case 4, 5 'Next page, Last page
                TestEnable = (m_lSelected < m_oCol.Count)
        End Select
    End If
End Function

Private Sub Class_Initialize()
    m_iScaleMode = vbInches
    m_iMFactor = (TWIPS_PER_INCHE / TWIPS_PER_PIXEL)
    m_nDisplayWidth = 816 ' 8.5 inches to pixels
    m_nDisplayHeight = 1056 '11 inches to pixels
    m_PaperSize = vbPRPSLetter
    m_bLandScape = False
    Set m_oCol = New Collection
    Set m_oRange = New clsRangeSelector
    Set m_oFontMap = New FontMap
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


Friend Property Get DisplayHeight() As Single
    DisplayHeight = m_nDisplayHeight
End Property

Public Property Let Width(ByVal vData As Single)
    m_nScaledWidth = vData
    m_nDisplayWidth = (m_nScaledWidth * m_iMFactor)
    prvValidatePaperSize
End Property


Friend Property Get DisplayWidth() As Single
    DisplayWidth = m_nDisplayWidth
End Property

Private Sub Class_Terminate()
    Set m_oCol = Nothing
End Sub

