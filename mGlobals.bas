Attribute VB_Name = "mGlobals"
'============ Control Placement and Selection ================
Public ActiveGrip As Integer               'which object grip box has been selected
Public LineIndex As Integer                'index counter for line objects
Public ShapeIndex As Integer                 'index counter for box objects
Public LabelIndex As Integer               'index counter for label object
Public FieldIndex As Integer               'index counter for field objects
Public ImageIndex As Integer               'index counter for static image objects
Public BoundImageIndex As Integer          'index counter for data-bound image objects
Public CheckShapeIndex As Integer            'index counter for check box objects

Public ctlTest As Control                  'pointer used for looping through controls
Public ctlActive As Control                'pointer to currently selected control
Public blnControlFound As Boolean          'a control has been found at mouse location
Public blnControlSelected As Boolean       'control has been selected with the mouse
Public blnGroupSelected As Boolean         'a group of controls has been selected by dragging mouse
Public blnGripFound As Boolean

Public blnSelectArrayInit As Boolean
Public ButtonTop As Single                 'saves mouse position on section dividers for dragging
Public FromSection As Integer
Public CurrSection As Integer              'saves the current page section for right-click actions
Public blnDragStarted As Boolean           'indicates if user has started dragging mouse across page
Public NumInGrp As Integer                 'number of objects selected
Public CutActive As Boolean                'if Cut mode is active
Public CopyActive As Boolean               'if Copy mode is active
Public strSpecialFieldContent As String    'for date, page and total fields
Public strCalcDataFieldContents As String      'for calculated fields contents
Public blnEditExisting As Boolean          'indicates if existing field is being editted
Public intDateFormatType As Integer
Public FontsLoaded As Boolean
Public PropertySelectMode As Integer

Public intGroupRestraint As Integer         'for select group moves - tracks direction that move is
Public Const resNone = 0                    'restrained by controls in group reaching limits of section
Public Const resTop = 1
Public Const resBottom = 2

'=========== Page Section Variables ============
Public FirstSectionVis As Integer
Public LastSectionVis As Integer
Public MinSectionHt(10) As Single           'tracks minimum section heights that can be set
Public MinPageWidth As Single              'tracks minimum page width that can be set
Public StartScaleLeft As Single            'tracks starting position of horizontal scale
Public GroupHVis(2) As Boolean
Public GroupFVis(2) As Boolean
Public blnPageChanged As Boolean
Public blnOnLastSection As Boolean

'=========== Program State Variables ===========
Public lngState As Long                    'tracks current program state or mode
Public lngPrevState As Long                'saves previous program state for restoration if needed
Public lngCmdState As Long                 'saves the current side toolbutton command state or mode

Public Const Default = 0               'program state constants
Public Const PlaceNewControl = 1
Public Const OverControl = 2
Public Const SelectControl = 3
Public Const DeleteControl = 4
Public Const MoveGrip = 5
Public Const EditLabel = 6
Public Const EditField = 7
Public Const EditPicture = 8
Public Const ResizeSection = 9
Public Const ResizePageWidth = 10
Public Const MoveControl = 11
Public Const SelectGroup = 13
Public Const MoveGroup = 14

'=========== Group Selection =============
Public Type SelectGroup
    blnOnClipBrd As Boolean
    ActiveGrip As Integer
    ctl As Control
    dX1 As Single
    dY1 As Single
    dX2 As Single
    dY2 As Single
End Type

Public SelectedCtl() As SelectGroup

'========== Report Design Parameters ===========
Public blnReportSaved As Boolean       'is currently open report saved
Public blnReportDataBound As Boolean   'does current report require data connection
Public blnGridOn As Boolean
Public GridSpace As Single
Public blnSnapOn As Boolean
Public blnCustomGrid As Boolean
Public PageSizeName As String
Public PageWd As Single
Public PageHt As Single
Public blnPageWidChanged As Boolean
Public PageOrient As Integer
Public Const cPortrait = 1
Public Const cLandscape = 2
Public LeftMarg As Single
Public RightMarg As Single
Public TopMarg As Single
Public BottomMarg As Single
Public blnHasPics As Boolean
Public strImgPathTable As String
Public strImgPathField As String
Public strImageFolder As String

'============ Active Placement Settings ============
Public Type BorderSettings
    Color As Long
    width As Integer
    Style As Integer
End Type

Public Type BackGroundSettings
    Style As Integer
    Color As Long
End Type

Public Type TextSettings
    Color As Long
    FontName As String
    FontSize As Integer
    IsBold As Boolean
    IsItalic As Boolean
    IsUnderline As Boolean
    Align As Integer
    BorderOn As Boolean
End Type

Public Type CheckBoxSettings
    DisplayType As Integer
    Sunken As Boolean
End Type

Public ActiveBorder As BorderSettings
Public ActiveBack As BackGroundSettings
Public ActiveText As TextSettings
Public ActiveChkBox As CheckBoxSettings
Public ActiveShape As Integer

'=========== Windows API SetPixelV function for drawing grid points ================
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
ByVal y As Long) As Long

Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, _
ByVal hBrush As Long) As Long

Public Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, _
ByVal HPALETTE As Long, pccolorref As Long) As Long

'============= Windows API DrawText function for horiz and vert scale numbers =============
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, _
ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long

Public Const DT_CENTER = &H1

Public Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Sub Main()

    Load frmDesign
    DoEvents
    
    frmDesign.Show

End Sub

Public Sub ShowGrid(Optional SectNo As Integer = -1)
Dim i As Integer, j As Integer, ptX As Single, ptY As Single
Dim SectionWidth As Long, SectionHeight As Long
Dim a As Integer

    If SectNo = -1 Then
        For a = FirstSectionVis To LastSectionVis
            frmDesign.picSection(a).Cls
        Next a
        If Not blnGridOn Then Exit Sub
        
        If PageScaleUnits = scEnglish Then
            For a = FirstSectionVis To LastSectionVis
                SectionWidth = frmDesign.picSection(a).width * 1440 / Screen.TwipsPerPixelX
                SectionHeight = frmDesign.picSection(a).Height * 1440 / Screen.TwipsPerPixelY
                frmDesign.picSection(a).DrawStyle = 0
                frmDesign.picSection(a).ForeColor = &HFF8080
                ptX = 0
                For i = 0 To Int(SectionWidth / 12)
                    ptX = ptX + 12
                    ptY = 0
                    For j = 0 To Int(SectionHeight / 12)
                        ptY = ptY + 12
                        SetPixelV frmDesign.picSection(a).hdc, ptX, ptY, &HFF8080
                    Next j
                Next i
            Next a
        ElseIf PageScaleUnits = scMetric Then
            For a = FirstSectionVis To LastSectionVis
                SectionWidth = frmDesign.picSection(a).width * 1440 / Screen.TwipsPerPixelX
                SectionHeight = frmDesign.picSection(a).Height * 1440 / Screen.TwipsPerPixelY
                frmDesign.picSection(a).DrawStyle = 0
                frmDesign.picSection(a).ForeColor = &HFF8080
                ptX = 0
                For i = 0 To Int(SectionWidth / 7.56)
                    ptX = ptX + 7.56
                    ptY = 0
                    For j = 0 To Int(SectionHeight / 7.56)
                        ptY = ptY + 7.56
                        SetPixelV frmDesign.picSection(a).hdc, ptX, ptY, &HFF8080
                    Next j
                Next i
            Next a
        End If
    Else
        SectionWidth = frmDesign.picSection(SectNo).width * 1440 / Screen.TwipsPerPixelX
        SectionHeight = frmDesign.picSection(SectNo).Height * 1440 / Screen.TwipsPerPixelY
        frmDesign.picSection(SectNo).DrawStyle = 0
        frmDesign.picSection(SectNo).ForeColor = &HFF8080
        frmDesign.picSection(SectNo).Cls
        ptX = 0
        For i = 0 To Int(SectionWidth / 12)
            ptX = ptX + 12
            ptY = 0
            For j = 0 To Int(SectionHeight / 12)
                ptY = ptY + 12
                SetPixelV frmDesign.picSection(SectNo).hdc, ptX, ptY, &HFF8080
            Next j
        Next i
    End If

    If PageScaleUnits = scEnglish Then
        For j = FirstSectionVis To LastSectionVis
            i = 1
            Do While i < frmDesign.picSection(j).width
                frmDesign.picSection(j).Line (i, 0)-(i, frmDesign.picSection(j).Height)
                i = i + 1
            Loop
            i = 1
            Do While i < frmDesign.picSection(j).Height
                frmDesign.picSection(j).Line (0, i)-(frmDesign.picSection(j).width, i)
                i = i + 1
            Loop
        Next j
    ElseIf PageScaleUnits = scMetric Then
        Dim metline As Single
        For j = FirstSectionVis To LastSectionVis
            metline = 0.394
            Do While metline < frmDesign.picSection(j).width
                frmDesign.picSection(j).Line (metline, 0)-(metline, frmDesign.picSection(j).Height)
                metline = metline + 0.394
            Loop
            metline = 0.394
            Do While metline < frmDesign.picSection(j).Height
                frmDesign.picSection(j).Line (0, metline)-(frmDesign.picSection(j).width, metline)
                metline = metline + 0.394
            Loop
        Next j
    End If
            
End Sub

Public Sub LoadFieldNames()
Dim i As Integer
Dim AddedItem As ListItem
Dim strFieldType As String

    frmSelField.lstFields.ListItems.Clear
    
    For i = 0 To UBound(DataField) - 1
        Set AddedItem = frmSelField.lstFields.ListItems.Add(, , DataField(i, 0))
        Select Case DataField(i, 1)
            Case 3: strFieldType = "Long Integer"
            Case 4: strFieldType = "Decimal"
            Case 6: strFieldType = "Currency"
            Case 7: strFieldType = "Date/Time"
            Case 11: strFieldType = "True/False"
            Case 202: strFieldType = "Text"
            Case 203: strFieldType = "Hyperlink/Memo"
            Case 205: strFieldType = "OLE Object"
        End Select
        AddedItem.ListSubItems.Add , , strFieldType
    Next i
    
   
End Sub

Public Sub ResetMinSectionHt(GetSection As Integer)
Dim i As Integer, ctlTestHt As Control

        MinSectionHt(GetSection) = 0
        For i = 0 To frmDesign.Controls.count - 1
            Set ctlTestHt = frmDesign.Controls(i)
            If ctlTestHt.Tag > "" Then
                If ctlTestHt.Visible = True Then
                    If TypeOf ctlTestHt Is Line Then
                        If ctlTestHt.Tag = GetSection Then
                            If ctlTestHt.Y1 >= ctlTestHt.Y2 Then
                                If ctlTestHt.Y1 > MinSectionHt(GetSection) Then
                                    MinSectionHt(GetSection) = ctlTestHt.Y1
                                End If
                            ElseIf ctlTestHt.Y2 > ctlTestHt.Y1 Then
                                If ctlTestHt.Y2 > MinSectionHt(GetSection) Then
                                    MinSectionHt(GetSection) = ctlTestHt.Y2
                                End If
                            End If
                        End If
                    ElseIf TypeOf ctlTestHt Is Shape Then
                        If ctlTestHt.Tag = GetSection And ctlTestHt.Name <> "Grip" Then
                            If ctlTestHt.Top + ctlTestHt.Height > MinSectionHt(GetSection) Then
                                MinSectionHt(GetSection) = ctlTestHt.Top + ctlTestHt.Height
                            End If
                        End If
                    ElseIf TypeOf ctlTestHt Is Label Or TypeOf ctlTestHt Is MSForms.Image Then
                        If ctlTestHt.Tag = GetSection Then
                            If ctlTestHt.Top + ctlTestHt.Height > MinSectionHt(GetSection) Then
                                MinSectionHt(GetSection) = ctlTestHt.Top + ctlTestHt.Height
                            End If
                        End If
                    End If
                End If
            End If
        Next i

End Sub

Public Sub ResetMinPageWidth()
Dim i As Integer, ctlTestWd As Control

        MinPageWidth = 0.2
        For i = 0 To frmDesign.Controls.count - 1
            Set ctlTestWd = frmDesign.Controls(i)
            If ctlTestWd.Tag > "" Then
                If TypeOf ctlTestWd Is Line Then
                    If ctlTestWd.X1 >= ctlTestWd.X2 Then
                        If ctlTestWd.X1 > MinPageWidth Then
                            MinPageWidth = ctlTestWd.X1
                        End If
                    ElseIf ctlTestWd.X2 > ctlTestWd.X1 Then
                        If ctlTestWd.X2 > MinPageWidth Then
                            MinPageWidth = ctlTestWd.X2
                        End If
                    End If
                ElseIf TypeOf ctlTestWd Is Shape Then
                    If ctlTestWd.Name <> "Grip" Then
                        If ctlTestWd.Left + ctlTestWd.width > MinPageWidth Then
                            MinPageWidth = ctlTestWd.Left + ctlTestWd.width
                        End If
                    End If
                ElseIf TypeOf ctlTestWd Is Label Or TypeOf ctlTestWd Is MSForms.Image Then
                    If ctlTestWd.Left + ctlTestWd.width > MinPageWidth Then
                        MinPageWidth = ctlTestWd.Left + ctlTestWd.width
                    End If
                End If
            End If
        Next i

End Sub
