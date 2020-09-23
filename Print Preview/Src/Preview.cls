VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Pages"
'////////////////////////////////////////////////////////
'///                 Print Preview Class
'///                   (clsPreview.cls)
'///_____________________________________________________
'/// Print Preview class with Navigation Bar.
'///_____________________________________________________
'/// Last modification  : Ago/09/2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Created
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const WM_SETREDRAW = &HB

Private m_oMemDC As clsMemDC
Private m_oImgLst As clsImageList

Private WithEvents m_oForm As frmPreview
Attribute m_oForm.VB_VarHelpID = -1
Private WithEvents m_oViewPort As PictureBox
Attribute m_oViewPort.VB_VarHelpID = -1
Private WithEvents m_oPreview As PictureBox
Attribute m_oPreview.VB_VarHelpID = -1
Private WithEvents m_oPage As PictureBox
Attribute m_oPage.VB_VarHelpID = -1
Private WithEvents m_oPages As Pages
Attribute m_oPages.VB_VarHelpID = -1
Private WithEvents m_oPageBack As PictureBox
Attribute m_oPageBack.VB_VarHelpID = -1

Private LnTPPX As Integer
Private m_iLastButton As ButtonIndex
Private m_iButtonCount As Integer
Private m_lOldParent As Long
Private m_lPageWidth As Long
Private m_lPageHeight As Long
Private m_iPageSize As Integer
Private m_lContainer As Long
Private m_iZoomLevel As ZoomRatio
Private m_bDragging As Boolean
Private m_rDragPos As POINTAPI
Private m_bLockRedraw As Boolean
Private m_bLockWndRedraw As Boolean
Private m_lPage As Long

Private Const PN_BUTTON_SIDE = 24

Private Type ButtonInfo
    ImageIdx As Integer
    Left As Long
    Rigth As Long
    Enabled As Boolean
    Caption As String
    ToolTipText As String
    Key As String
    IsSeparator As Boolean
End Type

Private Buttons() As ButtonInfo

Private Enum ButtonIndex
    [No Button] = 0
    PageFirst
    PagePrevious
    PageGoto
    PageNext
    PageLast
    [Sep 1]
    [Print Pages]
    [Sep 2]
    ZoomPreview
    [Sep 3]
    PreviewTools
    CustomCommand
    [Sep 4]
    ClosePreview
End Enum

Private Enum ZoomRatio
    [100%] = 0
    [75%]
    [50%]
    [25%]
    [Page Width]
    [Full Page]
    [Custom Ratio]
    [Hide Page]
End Enum

Private Enum CursorResource
    HandOpen = 101
    HandClosed = 102
    HandClosedArrows = 103
    HandClosedUpDown = 104
    HandClosedLeftRight = 105
End Enum

Public Event UserCommand()
Public Event Printing()
Public Event Done(iFormat As ARDestination)

Public OutputFile As String
Public LaunchFile As Boolean

' Temp filke related
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 260
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Public Sub AboutBox()
    frmAbout.Start
End Sub

Public Property Get LockRedraw() As Boolean
    LockRedraw = m_bLockRedraw
End Property

Public Property Let LockRedraw(bLock As Boolean)
    If (m_bLockRedraw <> bLock) Then
        If m_bLockRedraw Then
            Redraw
        Else
            m_bLockRedraw = bLock
        End If
    End If
End Property


Public Sub Redraw()
    m_bLockRedraw = False
    If m_bLockWndRedraw Then
        LockWindowUpdate 0
        m_bLockWndRedraw = False
    Else
        prvBuildPageImage
    End If
End Sub


Friend Sub ShowPageSelection(sSelection As String)
    m_oForm.lblRange = sSelection
End Sub

Public Function SyncShell(sCmdLine As String) As Long
    Dim lR As Long
    On Error Resume Next
    lR = ShellExecute(0&, "Open", sCmdLine, "", "", vbNormalFocus)
End Function

Private Function GetTempFile(Optional Prefix As String) As String
'    Dim PathName As String
'    Dim sRet As String
'
'    If Prefix = "" Then Prefix = ""
'    PathName = GetTempDir
'    sRet = String(MAX_PATH, 0)
'    GetTempFileName PathName, Prefix, 0, sRet
'    GetTempFile = StrZToStr(sRet)

    Dim PathName As String

    If Prefix = "" Then Prefix = "ARExp"
    PathName = GetTempDir
    If (Right(PathName, 1) <> "\") Then PathName = PathName & "\"
    GetTempFile = PathName & Prefix & Hex(Timer) & ".tmp"
End Function

Private Function GetTempDir() As String
    Dim sRet As String
    Dim c As Long
    sRet = String(MAX_PATH, 0)
    c = GetTempPath(MAX_PATH, sRet)
    If c = 0 Then
        GetTempDir = App.Path
    Else
        GetTempDir = Left$(sRet, c)
    End If
End Function


Public Sub Cls()
    m_oMemDC.Cls
End Sub

Public Property Get Pages() As Pages
    Set Pages = m_oPages
End Property

Private Sub prvDisableButtons()
    Dim LnIdx As Integer
    For LnIdx = PageFirst To ZoomPreview
        Buttons(LnIdx).Enabled = False
    Next
    If (Not m_bLockRedraw) Then prvDrawToolbar
End Sub

Private Sub prvEnableButtons()
    Dim LnIdx As Integer
    
    With m_oPages
        Buttons(PageFirst).Enabled = .TestEnable(PageFirst)
        Buttons(PagePrevious).Enabled = .TestEnable(PagePrevious)
        Buttons(PageGoto).Enabled = .TestEnable(PageGoto)
        Buttons(PageNext).Enabled = .TestEnable(PageNext)
        Buttons(PageLast).Enabled = .TestEnable(PageLast)
    End With
     Buttons([Print Pages]).Enabled = (Printers.Count > 0)
'    For LnIdx = PrintPages To ZoomPreview
'        Buttons(LnIdx).Enabled = True
'    Next
     Buttons(ZoomPreview).Enabled = True
    prvDrawToolbar
End Sub

Private Sub prvLoadForm()
    Load frmPreview
    Set m_oForm = frmPreview
    m_lContainer = m_oForm.hWnd
    Set m_oViewPort = m_oForm.picViewPort
    Set m_oPreview = m_oForm.picPreview
    Set m_oPage = m_oForm.picPage
    Set m_oPageBack = m_oForm.picBack
    
End Sub

Private Sub prvBuildPageImage()
    If m_bLockRedraw Then Exit Sub
    Dim LnIdx As Integer
    Dim LbRebuildDC As Boolean
    
    With Pages.ActivePage
        m_lPageWidth = .DisplayWidth
        m_lPageHeight = .DisplayHeight
    End With
    With m_oMemDC
        LbRebuildDC = ((.Width <> m_lPageWidth) Or (.Height <> m_lPageHeight))
        If LbRebuildDC Then
            .Create m_lPageWidth, m_lPageHeight
        Else
            .Cls
        End If
    End With
    With Pages.ActivePage
        For LnIdx = 1 To .Count
            With .Element(LnIdx)
                Select Case .Type
                    Case 1 ' Text
                        Dim LhOldFont As Long
                        
                        LhOldFont = SelectObject(m_oMemDC.hDC, Pages.FontMap.Item(.FontIndex).Handle)
                        m_oMemDC.DrawText .Text, .Left, .Top, .Width, .Height, .ForeColor, .BackColor, .Aligment
                        LhOldFont = SelectObject(m_oMemDC.hDC, LhOldFont)
                    Case 2 ' Line
                        m_oMemDC.DrawLine .Left, .Top, .Width, .Height, .ForeColor, CLng(.Size), .Pen
                    Case 3 ' Box
                        m_oMemDC.DrawShape .DisplayType, .Left, .Top, .Width, .Height, CInt(.Size), .ForeColor, .BackColor, .Pen
                    Case 4 ' Picture
                        On Error Resume Next
                        m_oMemDC.DrawPicture .Picture, .Left, .Top, .Width, .Height
                    Case 5 ' Checkbox
                        m_oMemDC.DrawCheckBox .DisplayType, .Checked, .Left, .Top, .Width, .Height, .ForeColor, .BackColor, CLng(.Size), .Sunken
                End Select
            End With
        Next
    End With
    If LbRebuildDC Then
        prvSetZoomRatio m_iZoomLevel
    Else
        With m_oPage
            .Visible = True
            m_oMemDC.BlitImage .hDC, 0, 0, .ScaleWidth, .ScaleHeight, Pages.ActivePage.Enabled
        End With
    End If
End Sub

Private Sub prvUnloadForm()
    ' Returns to original container
    Container = m_oForm.hWnd
    m_lContainer = 0
    ' Remove controls references
    Set m_oViewPort = Nothing
    Set m_oPreview = Nothing
    ' Unload support form
    Set m_oForm = Nothing
    Unload frmPreview
End Sub

Friend Sub RaiseErr(ByVal lErrNum As RSErrorCode, Optional sRoutineName As String, _
    Optional sDescription As String)
    RaiseError lErrNum, TypeName(Me), sRoutineName, sDescription, Erl
End Sub

Public Property Let Container(LhWnd As Long)
    If (LhWnd = m_lContainer) Then
        AdjustClient
    Else
        If (LhWnd = 0) Then
            Call SetParent(m_oPreview.hWnd, m_oForm.hWnd)
            Buttons(ClosePreview).Enabled = True
        Else
            If IsWindowLocal(LhWnd) Then
                m_lOldParent = SetParent(m_oPreview.hWnd, LhWnd)
            Else
                Me.RaiseErr ecNoExternalWindow, "Container[Let]"
            End If
        End If
        m_lContainer = GetParent(m_oPreview.hWnd)
        Buttons(ClosePreview).Enabled = (m_lContainer = m_oForm.hWnd)
        AdjustClient
    End If
End Property

Public Property Get Container() As Long
    Container = m_lContainer
End Property

Private Sub prvButtonClick(iButton As ButtonIndex)
    Select Case iButton
        Case [Print Pages]
            PrintPages
        Case PageFirst, PagePrevious, PageNext, PageLast, PageGoto
            prvNavigateTo iButton
        Case ZoomPreview, PreviewTools
            Dim LrPoint As POINTAPI
        
            With LrPoint
                .X = Buttons(iButton).Left
                .Y = PN_BUTTON_SIDE
            End With
            ClientToScreen m_lContainer, LrPoint
            ScreenToClient m_oForm.hWnd, LrPoint
            With m_oForm
                If (iButton = ZoomPreview) Then
                    .PopupMenu .mnuZoom, , LrPoint.X, LrPoint.Y
                ElseIf (iButton = PreviewTools) Then
                    .PopupMenu .mnuTools, , LrPoint.X, LrPoint.Y
                Else
                    .PopupMenu .mnuExport, , LrPoint.X, LrPoint.Y
                End If
            End With
        Case CustomCommand
            RaiseEvent UserCommand
        Case ClosePreview
            m_oForm.Hide
            'Unload m_oForm
    End Select
End Sub

Private Sub prvLoadImages()
    Dim LnIdx As Integer
    
    With m_oImgLst
        .Create m_oForm.picImages.hDC, Size16
        .Clear
        For LnIdx = 0 To 12
            .AddFromPictureBox m_oForm.picImages.hDC, m_oForm.picImages, (LnIdx * 16)
        Next
        'Debug.Print .ImageCount & " Images."
    End With
End Sub

Private Sub prvNavigateTo(iAction As ButtonIndex)
    Select Case iAction
        Case PageFirst, PagePrevious, PageNext, PageLast
            m_oPages.NavigateTo iAction
        Case PageGoto
            Dim LsPrompt As String
            
            If (m_oPages.Count = 0) Then
                Buttons(PageGoto).Enabled = False
                prvDrawToolbar PageGoto
                LsPrompt = "There are not created pages yet"
                MsgBox LsPrompt, vbOKOnly Or vbInformation
            Else
                Dim LnPage As Long
                
                LsPrompt = "Please enter the page number you want to go to" & vbCrLf & _
                    "(a valid number between 1 to " & m_oPages.Count & ")"
                LnPage = Val(InputBox(LsPrompt, "ActiveReporter"))
                m_oPages.SelectPage LnPage
            End If
    End Select
End Sub

Private Sub prvSetZoomRatio(ByVal iRatio As ZoomRatio)
    If (m_oPages.Count = 0) Then Exit Sub
    Dim LnZoomRatio As Single
    Dim LrPage As RECT
    Dim LrClient As RECT
    Dim LnCurFlag As Byte
    
    GetClientRect m_oViewPort.hWnd, LrClient
    InflateRect LrClient, -4, -4
    OffsetRect LrClient, -(LrClient.Left), -(LrClient.Top)
    m_iZoomLevel = iRatio
    Select Case m_iZoomLevel
        Case [100%]
            LnZoomRatio = 1
        Case [75%]
            LnZoomRatio = 0.75
        Case [50%]
            LnZoomRatio = 0.5
        Case [25%]
            LnZoomRatio = 0.25
        Case [Page Width]
            LnZoomRatio = (LrClient.Right / m_lPageWidth)
        Case [Full Page]
            Dim LnCRatio As Single
            Dim LnPRatio As Single
            
            With LrClient
                LnCRatio = (.Bottom / .Right)
                'In case preview is being destroyed
                If ((m_lPageHeight = 0) Or (m_lPageWidth = 0)) Then
                    prvBuildPageImage
                    Exit Sub
                End If
                LnPRatio = (m_lPageHeight / m_lPageWidth)
                If (LnCRatio > LnPRatio) Then
                    LnZoomRatio = (LrClient.Right / m_lPageWidth)
                Else
                    LnZoomRatio = (LrClient.Bottom / m_lPageHeight)
                End If
            End With
        Case [Custom Ratio]
            Dim LnUsrRatio As Single
            
            LnUsrRatio = Val(InputBox("Enter a zoom ratio between 10 and 100"))
            If ((LnUsrRatio < 10) Or (LnUsrRatio > 100)) Then
                Exit Sub
            Else
                LnZoomRatio = (LnUsrRatio / 100)
                m_iZoomLevel = LnUsrRatio
            End If
        Case [Hide Page]
            m_oPage.Visible = False
            m_oForm.VScroll1.Enabled = False
            m_oForm.HScroll1.Enabled = False
            Buttons(ZoomPreview).Caption = "Zoom Level"
            prvDrawToolbar ZoomPreview
            Exit Sub
        Case Else
            LnZoomRatio = (m_iZoomLevel / 100)
    End Select
    If (LnZoomRatio < 0.1) Then
        LnZoomRatio = 0.1
    End If
    Buttons(ZoomPreview).Caption = Format(LnZoomRatio, "##0.0#%")
    prvDrawToolbar ZoomPreview
    m_oPage.MousePointer = vbDefault
    With LrPage
        .Right = (m_lPageWidth * LnZoomRatio)
        .Bottom = (m_lPageHeight * LnZoomRatio)
        If (.Right > LrClient.Right) Then
            .Left = 4
            With m_oForm.HScroll1
                .Enabled = True
                .Min = 4
                .Max = (LrClient.Right - LrPage.Right)
                .LargeChange = LrPage.Bottom
                .SmallChange = (.LargeChange / 5)
                m_oPage.MousePointer = vbCustom
                m_bLockWndRedraw = True
                .Value = .Min
                LnCurFlag = 1
            End With
        Else
            .Left = ((LrClient.Right - .Right) / 2) + 4
            m_oForm.HScroll1.Enabled = False
        End If
        If (.Bottom > LrClient.Bottom) Then
            .Top = 4
            With m_oForm.VScroll1
                .Enabled = True
                .Min = 4
                .Max = (LrClient.Bottom - LrPage.Bottom)
                .LargeChange = LrPage.Right
                .SmallChange = (.LargeChange / 5)
                m_oPage.MousePointer = vbCustom
                m_bLockWndRedraw = True
                .Value = .Min
                LnCurFlag = LnCurFlag Or 2
            End With
        Else
            .Top = ((LrClient.Bottom - .Bottom) / 2) + 4
            m_oForm.VScroll1.Enabled = False
        End If
    End With
    Select Case LnCurFlag
        Case 1 ' Horizontal only
            m_oPage.DragIcon = LoadResPicture(HandClosedLeftRight, vbResCursor)
        Case 2 ' Vertical only
            m_oPage.DragIcon = LoadResPicture(HandClosedUpDown, vbResCursor)
        Case 3 ' Both
            m_oPage.DragIcon = LoadResPicture(HandClosedArrows, vbResCursor)
    End Select
    LockWindowUpdate m_oViewPort.hWnd
    With LrPage
        MoveWindow m_oPageBack.hWnd, .Left + 5, .Top + 5, .Right, .Bottom, True
        MoveWindow m_oPage.hWnd, .Left, .Top, .Right, .Bottom, True
    End With
    With m_oPage
        .Visible = True
        m_oMemDC.BlitImage .hDC, 0, 0, .ScaleWidth, .ScaleHeight, .Enabled
    End With
    m_oPageBack.Visible = True
    If (Not m_bLockRedraw) Then
        LockWindowUpdate 0
        m_bLockWndRedraw = False
    End If
End Sub

Friend Sub AdjustClient()
    Dim LrRect As RECT
    
    GetClientRect m_lContainer, LrRect
    With LrRect
        OffsetRect LrRect, (-.Left), (-.Top)
        MoveWindow m_oPreview.hWnd, 0, 0, .Right, .Bottom, True
    End With
    prvSetZoomRatio m_iZoomLevel
End Sub

Public Sub SetCustomCommand(sCaption As String)
    With Buttons(CustomCommand)
        .Enabled = (Len(sCaption) > 0)
        If .Enabled Then
            .ToolTipText = sCaption
        Else
            .ToolTipText = "Custom Command"
        End If
    End With
    prvDrawToolbar CustomCommand
End Sub

Public Sub Show()
    With m_oPages
        .FontMap.Create
        If (.Count > 0) Then
            .SelectPage 1
            m_bLockRedraw = False
            prvBuildPageImage
            prvSetZoomRatio [Full Page]
        End If
    End With
    If (m_lContainer = m_oForm.hWnd) Then
        Buttons(ZoomPreview).Caption = "Zoom Level"
        m_oForm.Start
    End If
End Sub

Public Sub PrintPages()
'////////////////////////////////////////////
'/// Send output to printer...
'////////////////////////////////////////////
    If (Printers.Count = 0) Then
        Me.RaiseErr ecNoPrintersDefined, "PrintPages"
    Else
        Dim LoPage As Page
        RaiseEvent Printing
        #If USE_LOG_FONT Then
        ' Recreates LOG_FONT objects to match with printer resolution
        Printer.Print ""
        m_oPages.FontMap.Create Printer.hDC
        Printer.KillDoc
        #End If
        For Each LoPage In m_oPages
            LoPage.PrintIt
        Next LoPage
        Printer.EndDoc
        #If USE_LOG_FONT Then
        ' Restores the LOG_FONT objects for screen usage
        m_oPages.FontMap.Create
        #End If
        RaiseEvent Done(arPrinter)
    End If
End Sub

Private Sub Class_Initialize()
    m_iPageSize = 1
    Set m_oPages = New Pages
    Set m_oPages.Parent = Me
    m_iZoomLevel = [Full Page]
    Set m_oMemDC = New clsMemDC
    Set m_oImgLst = New clsImageList
    prvLoadForm
    prvCreateButtons
    prvLoadImages
    
End Sub

Private Sub Class_Terminate()
    prvUnloadForm
End Sub

Private Sub m_oForm_ChangeZoomRatio(iRatioIndex As Integer)
    prvSetZoomRatio iRatioIndex
End Sub

Private Sub m_oForm_DoScroll(lValue As Long, bVertical As Boolean)
    Dim LrPage As RECT
    Dim LrPoint As POINTAPI
    
    GetWindowRect m_oPage.hWnd, LrPage
    ScreenToClient m_oViewPort.hWnd, LrPoint
    OffsetRect LrPage, LrPoint.X, LrPoint.Y
    LockWindowUpdate m_oViewPort.hWnd
'    Call SendMessage(m_oViewPort.hwnd, WM_SETREDRAW, False, 0)
    If bVertical Then
        With LrPage
            MoveWindow m_oPage.hWnd, .Left, lValue, (.Right - .Left), (.Bottom - .Top), True
        End With
    Else
        With LrPage
            MoveWindow m_oPage.hWnd, lValue, .Top, (.Right - .Left), (.Bottom - .Top), True
        End With
    End If
'    Call SendMessage(m_oViewPort.hwnd, WM_SETREDRAW, True, 0)
    If Not m_bLockWndRedraw Then LockWindowUpdate 0
End Sub

Private Sub m_oForm_GotoPage()
    prvButtonClick PageGoto
End Sub

Private Sub m_oForm_ToolSelected(iToolIndex As Integer)
    Select Case iToolIndex
        Case 0 'Choose printer
            Dim LoPrnSel As clsPrintDialog
            
            Set LoPrnSel = New clsPrintDialog
            With LoPrnSel
                .Min = 1
                .Max = 1000
                .Flags = cdlPDNoPageNums Or cdlPDNoSelection
                '.FromPage = m_oRange.RangeMin
                '.ToPage = m_oRange.RangeMax
                .hWndOwner = m_lContainer
                .ShowPrinter
            End With
        Case 1 'Select pages
            With Pages
                .Selection = InputBox("Enter pages range separated with commas." & vbCrLf & _
                    "(i.e. '1,3-5,7-*'), '*' for all pages.", , .Selection)
                m_oPage_Paint
            End With
    End Select
End Sub

Private Sub m_oPage_DragDrop(Source As Control, X As Single, Y As Single)
    Dim LnDifference As Long
    Dim LnNewValue As Long
    Dim LrPos As POINTAPI
    
    LockWindowUpdate m_oViewPort.hWnd
    m_bLockWndRedraw = True
    GetCursorPos LrPos
    With m_oForm.VScroll1
        If .Enabled Then
            LnDifference = (LrPos.Y - m_rDragPos.Y)
            If (.Max < 0) Then
                LnNewValue = .Value + LnDifference
                If (LnNewValue > .Min) Then
                    LnNewValue = .Min
                ElseIf (LnNewValue < .Max) Then
                    LnNewValue = .Max
                End If
            Else
            End If
            .Value = LnNewValue
        End If
    End With
    With m_oForm.HScroll1
        If .Enabled Then
            LnDifference = (LrPos.X - m_rDragPos.X)
            If (.Max < 0) Then
                LnNewValue = .Value + LnDifference
                If (LnNewValue > .Min) Then
                    LnNewValue = .Min
                ElseIf (LnNewValue < .Max) Then
                    LnNewValue = .Max
                End If
            Else
            End If
            .Value = LnNewValue
        End If
    End With
    m_oPage.Drag vbEndDrag
    m_bDragging = False
    m_bLockWndRedraw = False
    LockWindowUpdate 0
End Sub

Private Sub m_oPage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        If m_oPage.MousePointer = vbCustom Then
            GetCursorPos m_rDragPos
            m_oPage.Drag vbBeginDrag
            m_bDragging = True
        End If
    Else
        m_bDragging = False
    End If
End Sub

Private Sub m_oPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_bDragging Then
        m_oPage_DragDrop m_oPage, X, Y
    End If
End Sub

Private Sub m_oPage_Paint()
    With m_oPage
        If (.Visible And Not (Pages.ActivePage Is Nothing)) Then
            m_oMemDC.BlitImage .hDC, 0, 0, .ScaleWidth, _
                .ScaleHeight, Pages.ActivePage.Enabled
        End If
    End With
End Sub

Private Sub m_oPages_Cleared()
    prvDisableButtons
    m_lPage = 0
    m_oForm.lblPageIndex = "No pages"
    m_oPage.Visible = False
End Sub

Private Sub m_oPages_PageAdded(iCount As Long)
    prvEnableButtons
    m_oForm.lblPageIndex = "Page " & m_lPage & "/" & iCount
End Sub

Private Sub m_oPages_PageSelected(ByVal lPage As Long, lCount As Long)
    Dim LnIdx As Integer
    
    m_lPage = lPage
    m_oPages_PageAdded lCount
    prvBuildPageImage
End Sub

Private Sub m_oPreview_Click()
    If (m_iLastButton > 0) Then
        If Buttons(m_iLastButton).Enabled Then
            prvButtonClick m_iLastButton
        End If
    End If
End Sub

Private Sub m_oPreview_KeyDown(KeyCode As Integer, Shift As Integer)
    'Debug.Print "KeyCode:" & KeyCode
    Select Case KeyCode
        Case 27 ' Escape
            prvButtonClick ClosePreview
        Case 33 ' Page Up
            prvButtonClick PagePrevious
        Case 34 ' Page Down
            prvButtonClick PageNext
        Case 35 ' End
            prvButtonClick PageLast
        Case 36 ' Home
            prvButtonClick PageFirst
    End Select
End Sub

Private Sub m_oPreview_KeyPress(KeyAscii As Integer)
    'Debug.Print "KeyAscii:" & KeyAscii
    Select Case KeyAscii
        Case 43 ' + key
'            If (m_iZoomLevel < [25%]) Then
'                prvSetZoomRatio (m_iZoomLevel - 1)
'            ElseIf (m_iZoomLevel > [25%]) Then
'                prvSetZoomRatio [Full Page]
'            End If
        Case 45 ' - key
'            If (m_iZoomLevel > [100%]) Then
'                prvSetZoomRatio (m_iZoomLevel + 1)
'            ElseIf (m_iZoomLevel > [25%]) Then
'                prvSetZoomRatio [Full Page]
'            End If
    End Select
End Sub

Private Sub m_oPreview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_iLastButton = prvButtonFromPoint(X)
    If (m_iLastButton > 0) Then
        If Buttons(m_iLastButton).Enabled Then
            prvDrawToolbar m_iLastButton, True
        End If
    End If
End Sub

Private Sub m_oPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim LnButton As Integer
    
    LnButton = prvButtonFromPoint(X)
    If (m_iLastButton <> LnButton) Then
        m_iLastButton = LnButton
    End If
    If (m_iLastButton > 0) Then
        m_oPreview.ToolTipText = Buttons(m_iLastButton).ToolTipText
    Else
        m_oPreview.ToolTipText = ""
    End If
End Sub

Private Function prvButtonFromPoint(X As Single) As ButtonIndex
    Dim LnIdx As Integer
    
    For LnIdx = 1 To m_iButtonCount
        If ((X >= Buttons(LnIdx).Left) And (X <= Buttons(LnIdx).Rigth)) Then
            prvButtonFromPoint = LnIdx
            Exit For
        End If
    Next
End Function

Private Sub prvCreateButtons()
    LnTPPX = Screen.TwipsPerPixelX
    prvAddButton "first", "", 0, "First Page", False
    prvAddButton "previous", "", 1, "Previous Page", False
    prvAddButton "goto", "", 2, "Go To Page", False
    prvAddButton "next", "", 3, "Next Page", False
    prvAddButton "last", "", 4, "Last Page", False
    prvAddButton "", "", , "", False, True
    prvAddButton "print", "Print", 7, "Send To Printer", False
    prvAddButton "", "", , "", False, True
    prvAddButton "zoom", "Zoom Level", 6, "Zoom Level", False
    prvAddButton "", "", , "", False, True
    prvAddButton "printer", "", 8, "Select Default Printer"
    prvAddButton "custom", "", 5, "Custom Command", False
    prvAddButton "", "", , "", False, True
    prvAddButton "close", "Close", , "Close Preview"
End Sub

Private Sub prvAddButton(Optional sKey As String, Optional sText As String, _
    Optional iImage As Integer = -1, Optional sTipText As String, _
    Optional bEnabled As Boolean = True, Optional bSeparator As Boolean)
    Static SnLastLeft As Long
    
    m_iButtonCount = m_iButtonCount + 1
    ReDim Preserve Buttons(1 To m_iButtonCount)
    With Buttons(m_iButtonCount)
        .Caption = sText
        .ImageIdx = iImage
        .IsSeparator = bSeparator
        .ToolTipText = sTipText
        .Enabled = bEnabled
        .Key = sKey
        .Left = SnLastLeft
        If .IsSeparator Then
            .Rigth = (.Left + 4)
        ElseIf (.ImageIdx = -1) Then
            .Rigth = (.Left + m_oPreview.TextWidth(sText) + 8)
        Else
            If (Len(sText) = 0) Then
                .Rigth = (.Left + PN_BUTTON_SIDE - 1)
            Else
                .Rigth = (.Left + PN_BUTTON_SIDE + m_oPreview.TextWidth(sText) + 4)
            End If
        End If
        SnLastLeft = (.Rigth + 1)
    End With
End Sub

Private Sub m_oPreview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    prvDrawToolbar m_iLastButton
End Sub

Private Sub m_oPreview_Paint()
    prvDrawToolbar
End Sub

Private Sub prvDrawToolbar(Optional ByVal iButton As ButtonIndex, _
    Optional bDepressed As Boolean)
    If iButton Then
        Dim LnLTColor As Long
        Dim LnRBColor As Long
        Dim LnOffsetX As Long
        Dim LnOffsetY As Long
        Dim LnBottom As Long
        Dim LnTextW As Single
        
        If bDepressed Then
            LnLTColor = vbButtonShadow
            LnRBColor = vbWindowBackground
            LnOffsetX = 4
            LnOffsetY = 5
        Else
            LnLTColor = vbWindowBackground
            LnRBColor = vbButtonShadow
            LnOffsetX = 3
            LnOffsetY = 4
        End If
        LnBottom = (PN_BUTTON_SIDE - 2) '(LnTPPX * 2))
        With Buttons(iButton)
            If Not .IsSeparator Then
                ' Back
                m_oPreview.Line (.Left, 1)-(.Rigth, LnBottom), vbButtonFace, BF
                ' Top line.
                m_oPreview.Line (.Left, 1)-(.Rigth, 1), LnLTColor
                ' Left line.
                m_oPreview.Line (.Left, 1)-Step(0, LnBottom), LnLTColor
                ' Right line.
                m_oPreview.Line (.Rigth, 1)-Step(0, LnBottom), LnRBColor
                ' Bottom line
                m_oPreview.Line (.Left, LnBottom)-(.Rigth, LnBottom), LnRBColor
                ' Draw image
                m_oImgLst.DrawImage .ImageIdx, m_oPreview.hDC, .Left + LnOffsetX, LnOffsetY, , (Not .Enabled)
                ' Draw caption
                If .Enabled Then
                    m_oPreview.ForeColor = vbBlack
                Else
                    m_oPreview.ForeColor = vbButtonShadow
                End If
                LnTextW = m_oPreview.TextWidth(.Caption)
                m_oPreview.CurrentX = .Left + IIf(.ImageIdx = -1, 4, PN_BUTTON_SIDE) + 1
                If bDepressed Then
                    m_oPreview.CurrentY = 5
                Else
                    m_oPreview.CurrentX = .Left + IIf(.ImageIdx = -1, 4, PN_BUTTON_SIDE)
                    m_oPreview.CurrentY = 4
                End If
                m_oPreview.Print .Caption
            End If
        End With
    Else
        Dim LnIdx As Integer
        
        For LnIdx = 1 To m_iButtonCount
            prvDrawToolbar LnIdx
        Next
    End If
End Sub

Private Sub m_oPreview_Resize()
    prvSetZoomRatio m_iZoomLevel
End Sub

Private Sub m_oViewPort_DragDrop(Source As Control, X As Single, Y As Single)
    If (Source Is m_oPage) Then
        m_oPage_DragDrop m_oPage, 0, 0
    End If
End Sub


