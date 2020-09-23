Attribute VB_Name = "mPreviewData"
'mPreviewData
'was written to be self-contained so that it can be used simply for previewing
'or printing existing reports without opening them in the Report Designer
'include this module in your application and execute with:
'        PreviewReport <True/False>, FilePath, <Host Window Handle>
'                           |           |                |
'        Do Preview First --O           |                |
'                         Report File --O                |
'             form handle or full screen if left blank --O
'
'
'======= Global variables ==========
Public OpenFileName As String              'name/location of current report file
Public PageScaleUnits As Integer
Public PageNo As Integer
Public NumPages As Integer
Public PgHeadTop As Single
Public PgFootTop As Single
Public PageFreeHt As Single
Public PageFreeWd As Single
Public FirstPgFreeHt As Single
Public LastPgFreeHt As Single
Public DetTop As Single
Public NumRecs As Integer
Public NumRecsFirstPage As Integer
Public NumRecsBodyPage As Integer
Public NumRecsLastPage As Integer
Public SectionTop As Single
Private blnHasCalcDataFields As Boolean
Private CalcNum As Integer
Private TotalPageControlNum As Integer          'which control shows total number of pages
Private TotalPageControlSection As Integer      'which section is the above control in

'Control Type Descriptor
'==================================
Public Type ControlInfo
    Name As String
    Index As Long
    Type As Integer
    SecNo As Integer
    X1 As Single
    Y1 As Single
    X2 As Single
    Y2 As Single
    Left As Single
    Top As Single
    width As Single
    Height As Single
    BdrWd As Integer
    BdrStl As Long
    BdrClr As Long
    BckStl As Integer
    BckClr As Long
    ForClr As Long
    FntNam As String
    FntSiz As Single
    FntBld As Boolean
    FntItl As Boolean
    FntUnd As Boolean
    Align As Long
    strText As String
    ImgData As String
    DisplayType As Integer
    Sunken As Boolean
    Fieldname As String
End Type

'Report File Save Format
'=======================================
Public Type ReportStructure
    DataBound As Boolean
    DBName As String
    DBSource As String
    HasPics As Boolean
    ImgPathTable As String
    ImgPathField As String
    ImageFolder As String
    SortField(2) As String
    SortDescending(2) As Boolean
    PageSclUnit As Integer
    PageSzNam As String
    PageWd As Single
    PageHt As Single
    Orient As Integer
    LMarg As Single
    RMarg As Single
    TMarg As Single
    BMarg As Single
    DesWd As Single
    HeaderVis(4) As Boolean
    FooterVis(4) As Boolean
    HeaderHt(4) As Single
    FooterHt(4) As Single
    NewPageOnHeader(2) As Boolean
    DetHt As Single
    SectColor(10) As Long
    RpControl() As ControlInfo
End Type

'============ Control Type Constants ==========
Public intControlType As Integer       'type of control about to be placed
Public Const cNone = 0
Public Const cDataField = 1                'control type constants
Public Const cLabel = 2
Public Const cCheckBox = 3
Public Const cLine = 4
Public Const cBox = 5
Public Const cImage = 6
Public Const cBoundImage = 7
Public Const cDatePageField = 8
Public Const cCalcField = 9
Public Const cSumField = 10

'============= Page Scale Unit Constants ============
Public Const scEnglish = 0
Public Const scMetric = 1

Public ReportFile As ReportStructure       'for saving report file to disk

'=========== Calculated Field Descriptor ===========
Private Type CalcInfo
    strFieldName As String
    strValue As String
End Type

Private SaveCalc() As CalcInfo

'=========== Preview ==============
Public PP As Preview

'============= ADO Connection ============
Public dbConn As ADODB.Connection
Public rstTables As ADODB.Recordset
Public rstData As ADODB.Recordset
Public strConnErrMsg As String
Public strDataFileName As String
Public strSortField(2) As String
Public blnNewPageOnHeader(2) As Boolean
Public blnSortDescending(2) As Boolean
Public strTableName As String
Public DataField() As String
Public FieldNo(2) As Integer

Public Sub UnloadPreview()

    Set PP = Nothing

End Sub

Public Function ConnectToDataFile() As Boolean
'attempts to establish the ADO connection - currently set up for MS Access only
On Error GoTo ConnectError
Dim strCnn As String

    strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDataFileName & ";Persist Security Info=False"
    Set dbConn = New ADODB.Connection
    dbConn.Open strCnn        'open connection to database file (MS Access)
    ConnectToDataFile = True
    Exit Function

ConnectError:
    MsgBox Err.Description, vbOKOnly, "Connection Error"
    ConnectToDataFile = False
    strDataFileName = ""
    strConnErrMsg = "Error in ConnectToDataFile : " & Err.Description

End Function

Public Function GetTables() As Boolean
'extracts list of tables/queries from the opened database
On Error GoTo NoTables

    Set rstTables = dbConn.OpenSchema(adSchemaTables)
    GetTables = True
    Exit Function

NoTables:
    GetTables = False
    strConnErrMsg = Err.Description
    MsgBox Err.Description, vbOKOnly, "Connection Error"

End Function

Public Sub OpenData(strTableName As String, Optional GetSort As String)
'opens a table/query and extracts a recordset, also populates the DataField array
'used for quick access to the field list of the table/query
On Error GoTo NoOpenData
Dim i As Integer

    Set rstData = New ADODB.Recordset
    rstData.Open "Select * from [" & strTableName & "] " & GetSort, dbConn, adOpenStatic, adLockOptimistic, adCmdText

    rstData.MoveLast
    rstData.MoveFirst
    
    ReDim DataField(rstData.Fields.count, 1)
    For i = 0 To UBound(DataField) - 1
        DataField(i, 0) = rstData.Fields(i).Name
        DataField(i, 1) = rstData.Fields(i).Type
    Next i
    
    Exit Sub

NoOpenData:
    MsgBox "Error in OpenData : " & Err.Description

End Sub

Public Sub PreviewReport(GoPreview As Boolean, Optional strFilePath As String = "%Current%", Optional HostWnd As Long)
'this the master sub for previewing/printing a report
'makes calls to GetReportFile, PreviewWithNoData and GenerateReport
'accepts arguments for preview first, external filepath to preview and host window handle
On Error GoTo NoPreview
Dim i As Integer
Dim Counter As Integer

'get an external report file to preview if specified
    If strFilePath <> "%Current%" Then
        If GetReportFile(strFilePath) = False Then
            MsgBox "Could not open '" & strFilePath & "'", vbOKOnly + vbInformation
            Exit Sub
        End If
    End If
        
'if no data connection found just preview what the user has put on the page
    If ReportFile.DataBound = False Then
        PreviewWithNoData GoPreview
        Exit Sub
    End If

'check to see if there are any calculated fields or total pages fields in the report
    With ReportFile
        For i = 0 To UBound(.RpControl)
            If .RpControl(i).Type = cCalcField Then
                Counter = Counter + 1
            ElseIf .RpControl(i).Type = cDatePageField Then
                If InStr(1, .RpControl(i).strText, "[NumPages]") > 0 Then
                    TotalPageControlNum = i
                    TotalPageControlSection = .RpControl(i).SecNo
                End If
            End If
        Next i
    End With
    If Counter > 0 Then
        blnHasCalcDataFields = True
        ReDim SaveCalc((Counter) * rstData.RecordCount)
    End If
    CalcNum = 0

'initialize the Print Preview object
    Set PP = New Preview
'run it in a form window if specified, otherwise full screen
    If Not IsNull(HostWnd) Then PP.Container = HostWnd
'generate the pages and items of the report
    GenerateReport

'go to the preview screen if specified, otherwise just print it
    If GoPreview Then
        PP.Show
    Else
        PP.PrintPages
    End If

'destroy the preview object to free memory
    Set PP = Nothing
    Exit Sub

NoPreview:

    MsgBox "Error in PreviewReport : " & Err.Description

End Sub

Private Function GetReportFile(strGetFilePath As String) As Boolean
'attempt to open and load an external report file into the ReportFile object
On Error GoTo NoOpen

Dim FileNum As Long
Dim GetNum As Integer

    If Dir(strGetFilePath) = "" Then
        MsgBox "Could not find specified file : '" & strGetFilePath & "'", vbOKOnly + vbInformation
        Exit Function
    End If

    FileNum = FreeFile()
    
    Open strGetFilePath For Binary Access Read Lock Write As FileNum
    Get FileNum, , GetNum
    ReDim ReportFile.RpControl(GetNum)
    Get FileNum, , ReportFile
    Close FileNum
    
    strDataFileName = ReportFile.DBName
    strTableName = ReportFile.DBSource
    
    If strDataFileName = "" Then
        MsgBox "No data file loaded.  Report cannot be previewed.", vbOKOnly + vbInformation
        GetReportFile = False
        Exit Function
    End If
    
'set up the data connection, open the underlying table/query
    If ConnectToDataFile Then
        GetTables
        OpenData strTableName
    Else
        GetReportFile = False
        Exit Function
    End If
    
    GetReportFile = True
    Exit Function

NoOpen:

    GetReportFile = False
    MsgBox "Error in GetReportFile : " & Err.Description

End Function

Private Sub GenerateReport()
'this sub does the work of generating the pages and their contents
On Error GoTo NoGenerate

Dim i As Integer, j As Integer, k As Integer, l As Integer
Dim PrevValue(2) As Variant
Dim blnFirstHeader(3) As Boolean
Dim TopOfFreeSpace As Single
Dim CurrPagePos As Single
Dim EndOfPage As Boolean
Dim EndOfReport As Boolean

    NumRecs = 0
    NumPages = 0

    For j = 0 To 2
        FieldNo(j) = -1
        blnFirstHeader(j) = True
    Next j
'find and store which fields (by number) - if any - are used for sorting
    For j = 0 To 2
        If strSortField(j) > "" Then
            For i = 0 To rstData.Fields.count - 1
                If rstData.Fields(i).Name = ReportFile.SortField(j) Then
                    FieldNo(j) = i
                End If
            Next i
        End If
    Next j
    
'set up the report page scalemode, size, orientation
    rstData.MoveFirst
    PageNo = 0
    PP.Cls
    With PP.Pages
        .ScaleMode = vbInches
        If ReportFile.Orient = cPortrait Then
            .Landscape = False
        Else
            .Landscape = True
        End If
        .width = ReportFile.PageWd
        .Height = ReportFile.PageHt
        .Add
    End With
    PageNo = PageNo + 1
        
    With ReportFile
        PageFreeWd = .DesWd
'main loop for report
        Do While Not EndOfReport
'set up free space for first page
            If PageNo = 1 Then
                PageFreeHt = .PageHt - .TMarg - .BMarg - .HeaderHt(0) - .HeaderHt(1) - .FooterHt(1)
                If .SectColor(0) <> vbWhite Then
                    PP.Pages.ActivePage.DrawShape 0, .LMarg, .TMarg, PageFreeWd, .HeaderHt(0), -1, .SectColor(0)
                End If
                If .SectColor(1) <> vbWhite Then
                    PP.Pages.ActivePage.DrawShape 0, .LMarg, .TMarg + .HeaderHt(0), PageFreeWd, .HeaderHt(1), -1, .SectColor(1)
                End If
                If .SectColor(9) <> vbWhite Then
                    PP.Pages.ActivePage.DrawShape 0, .LMarg, .PageHt - .BMarg - .FooterHt(1), PageFreeWd, .FooterHt(1), -1, .SectColor(9)
                End If
                For i = 1 To UBound(.RpControl)
                    If .RpControl(i).SecNo = 0 Then
                        DrawObject i, .TMarg
                    ElseIf .RpControl(i).SecNo = 1 Then
                        If i <> TotalPageControlNum Then
                            DrawObject i, .TMarg + .HeaderHt(0)
                        End If
                    ElseIf .RpControl(i).SecNo = 9 Then
                        If i <> TotalPageControlNum Then
                            DrawObject i, .PageHt - .BMarg - .FooterHt(1)
                        End If
                    End If
                Next i
                TopOfFreeSpace = .TMarg + .HeaderHt(0) + .HeaderHt(1)
'set up free space for all other pages
            Else
                DoPageHeaderFooter
                TopOfFreeSpace = .TMarg + .HeaderHt(1)
            End If
            CurrPagePos = TopOfFreeSpace
            EndOfPage = False
            
'loop for each page
            Do While Not EndOfPage
' process group headers
                If Not EndOfPage And Not rstData.EOF Then
                    For j = 0 To 2
                        If .HeaderVis(j + 2) Then
                            If rstData.Fields(FieldNo(j)) <> PrevValue(j) Then
                                If blnNewPageOnHeader(j) Then
                                    If Not blnFirstHeader(j) Then
                                        PP.Pages.Add
                                        PageNo = PageNo + 1
                                        DoPageHeaderFooter
                                        TopOfFreeSpace = .TMarg + .HeaderHt(1)
                                        CurrPagePos = .TMarg + .HeaderHt(1)
                                        If j = 0 Then
                                            blnFirstHeader(1) = True
                                            blnFirstHeader(2) = True
                                        ElseIf j = 1 Then
                                            blnFirstHeader(2) = True
                                        End If
                                    Else
                                        blnFirstHeader(j) = False
                                    End If
                                End If
                                If CurrPagePos + .HeaderHt(j + 2) < PageFreeHt + TopOfFreeSpace Then
                                    If .SectColor(j + 2) <> vbWhite Then
                                        PP.Pages.ActivePage.DrawShape 0, .LMarg, CurrPagePos, PageFreeWd, .HeaderHt(j + 2), -1, .SectColor(j + 2)
                                    End If
                                    For i = 1 To UBound(.RpControl)
                                        If .RpControl(i).SecNo = j + 2 Then
                                            DrawObject i, CurrPagePos
                                        End If
                                    Next i
                                    CurrPagePos = CurrPagePos + .HeaderHt(j + 2)
                                Else
                                    EndOfPage = True
                                End If
                            End If
                        End If
                    Next j
                Else
                    EndOfPage = True
                End If
'process detail section
                If Not EndOfPage And Not rstData.EOF Then
                    If CurrPagePos + .DetHt < PageFreeHt + TopOfFreeSpace Then
                        For i = 1 To UBound(.RpControl)
                            If .RpControl(i).SecNo = 5 Then
                                DrawObject i, CurrPagePos
                            End If
                        Next i
                        CurrPagePos = CurrPagePos + .DetHt
                        For i = 0 To 2
                            If FieldNo(i) > -1 Then
                                PrevValue(i) = rstData.Fields(FieldNo(i)).value
                            End If
                        Next i
                        rstData.MoveNext
                    Else
                        EndOfPage = True
                    End If
                Else
                    EndOfPage = True
                End If
'process group footers
                If Not EndOfPage And Not rstData.EOF Then
                    For j = 2 To 0 Step -1
                        If .FooterVis(j + 2) Then
                            If rstData.Fields(FieldNo(j)) <> PrevValue(j) Then
                                If CurrPagePos + .FooterHt(j + 2) < PageFreeHt + TopOfFreeSpace Then
                                    If .SectColor(8 - j) <> vbWhite Then
                                        PP.Pages.ActivePage.DrawShape 0, .LMarg, CurrPagePos, PageFreeWd, .FooterHt(j + 2), -1, .SectColor(8 - j)
                                    End If
                                    If Not rstData.BOF Then rstData.MovePrevious
                                    For i = 1 To UBound(.RpControl)
                                        If .RpControl(i).SecNo = 8 - j Then
                                            DrawObject i, CurrPagePos
                                        End If
                                    Next i
                                    CurrPagePos = CurrPagePos + .FooterHt(j + 2)
                                    rstData.MoveNext
                                Else
                                    EndOfPage = True
                                End If
                            End If
                        End If
                    Next j
'if at the end of data, do last set of group footers
                Else
                    rstData.MovePrevious
                    For j = 2 To 0 Step -1
                        If .FooterVis(j + 2) Then
                            If CurrPagePos + .FooterHt(j + 2) < PageFreeHt + TopOfFreeSpace Then
                                If .SectColor(8 - j) <> vbWhite Then
                                    PP.Pages.ActivePage.DrawShape 0, .LMarg, CurrPagePos, PageFreeWd, .FooterHt(j + 2), -1, .SectColor(8 - j)
                                End If
                                For i = 1 To UBound(.RpControl)
                                    If .RpControl(i).SecNo = 8 - j Then
                                        DrawObject i, CurrPagePos
                                    End If
                                Next i
                                CurrPagePos = CurrPagePos + .FooterHt(j + 2)
                            Else
                                EndOfPage = True
                            End If
                        End If
                    Next j
                    EndOfPage = True
                    rstData.MoveNext
                End If
                
            Loop
            
'if not at the end of data, add another page
            If Not rstData.EOF Then
                PP.Pages.Add
                PageNo = PageNo + 1
'if at the end of data do report footer
            Else
'if it fits put it in
                If CurrPagePos + .FooterHt(0) < PageFreeHt + TopOfFreeSpace Then
                    If .SectColor(10) <> vbWhite Then
                        PP.Pages.ActivePage.DrawShape 0, .LMarg, CurrPagePos, PageFreeWd, .FooterHt(0), -1, .SectColor(10)
                    End If
                    For i = 1 To UBound(.RpControl)
                        If ReportFile.RpControl(i).SecNo = 10 Then
                            DrawObject i, CurrPagePos
                        End If
                    Next i
'otherwise add another page and put it in
                Else
                    PP.Pages.Add
                    PageNo = PageNo + 1
                    DoPageHeaderFooter
                    CurrPagePos = ReportFile.TMarg + .HeaderHt(1)
                    If .SectColor(10) <> vbWhite Then
                        PP.Pages.ActivePage.DrawShape 0, .LMarg, CurrPagePos, PageFreeWd, .FooterHt(0), -1, .SectColor(10)
                    End If
                    For i = 1 To UBound(.RpControl)
                        If ReportFile.RpControl(i).SecNo = 10 Then
                            DrawObject i, CurrPagePos
                        End If
                    Next i
                End If
                EndOfReport = True
            End If
            
        Loop
        
        NumPages = PageNo
        If TotalPageControlNum > -1 Then
            For i = 1 To NumPages
                PageNo = i
                If TotalPageControlSection = 1 Then
                    If i = 1 Then
                        If .HeaderVis(0) Then
                            CurrPagePos = .TMarg + .HeaderHt(0)
                        Else
                            CurrPagePos = .TMarg
                        End If
                    Else
                        CurrPagePos = .TMarg
                    End If
                ElseIf TotalPageControlSection = 9 Then
                    CurrPagePos = PageHt - .BMarg - .FooterHt(1)
                End If
                PP.Pages.SelectPage CLng(i)
                DrawObject TotalPageControlNum, CurrPagePos
            Next i
        End If
    
    End With
    
    Exit Sub


NoGenerate:

    Set PP = Nothing
    MsgBox "Error in GenerateReport : " & Err.Description

End Sub

Private Sub DoPageHeaderFooter()
'draws items in the page header and footer for a given page when called
Dim i As Integer

    With ReportFile
        PageFreeHt = .PageHt - .TMarg - .BMarg - .HeaderHt(1) - .FooterHt(1)
        If .SectColor(1) <> vbWhite Then
            PP.Pages.ActivePage.DrawShape 0, .LMarg, .TMarg, PageFreeWd, .HeaderHt(1), -1, .SectColor(1)
        End If
        If .SectColor(9) <> vbWhite Then
            PP.Pages.ActivePage.DrawShape 0, .LMarg, .PageHt - .BMarg - .FooterHt(1), PageFreeWd, .FooterHt(1), -1, .SectColor(9)
        End If
        
        For i = 1 To UBound(.RpControl)
            If ReportFile.RpControl(i).SecNo = 1 Then
                If i <> TotalPageControlNum Then
                    DrawObject i, .TMarg
                End If
            ElseIf ReportFile.RpControl(i).SecNo = 9 Then
                If i <> TotalPageControlNum Then
                    DrawObject i, .PageHt - .BMarg - .FooterHt(1)
                End If
            End If
        Next i
    End With

End Sub

Private Sub DrawObject(Index As Integer, TopOffset As Single)
'this sub handles drawing of each individual control item in the report
'accepts arguments for control index number and offset from the top of the page
On Error GoTo NoDraw
Dim i As Integer, j As Integer
Dim strFieldName As String
Dim strFValue As String
Dim bkcolor As Long
Dim bdrcolor As Long
Dim GetPic As StdPicture
Dim strGetFormat As String
Dim AggFunc As String
Dim ctlwidth As Single
Dim ctlheight As Single

With ReportFile.RpControl(Index)

    If .BdrStl = 0 Then
        bdrcolor = -1
    Else
        bdrcolor = .BdrClr
    End If
    
    If .BckStl = 0 Then
        bkcolor = -1
    Else
        bkcolor = .BckClr
    End If
    
    ctlwidth = .width
    ctlheight = .Height
    
    If .Type = cLine Then
        PP.Pages.ActivePage.DrawLine ReportFile.LMarg + .X1, TopOffset + .Y1, _
        ReportFile.LMarg + .X2, TopOffset + .Y2, .BdrClr, .BdrWd, .BdrStl - 1
    ElseIf .Type = cBox Then
        PP.Pages.ActivePage.DrawShape .DisplayType, ReportFile.LMarg + .Left, TopOffset + .Top, _
        ctlwidth, ctlheight, bdrcolor, bkcolor, .BdrWd, .BdrStl - 1
    ElseIf .Type = cLabel Then       'label control used for labels and fields on the report
        PP.Pages.ActivePage.DrawShape 0, ReportFile.LMarg + .Left, TopOffset + .Top, _
        ctlwidth, ctlheight, bdrcolor, bkcolor, 1, .BdrStl - 1
        
        PP.Pages.ActivePage.SetFont .FntNam, .FntSiz, .FntBld, .FntItl, .FntUnd, False, 0
        PP.Pages.ActivePage.DrawText .strText, ReportFile.LMarg + .Left + 0.01, TopOffset + .Top + 0.01, _
        ctlwidth - 0.02, ctlheight - 0.02, .ForClr, bkcolor, .Align
        
'check for either a database field or a special field with date, page no., etc.
    ElseIf .Type = cDataField Or .Type = cDatePageField Or .Type = cCalcField Or .Type = cSumField Then
        strFValue = "Error!"
        If .Type = cDatePageField Then
            If InStr(1, .strText, "=[Date") > 0 Then
                strFValue = Trim(.strText)
                strFValue = Mid(strFValue, 8, Len(strFValue) - 8)
                If Left(strFValue, 4) = "wwww" Then
                    strFValue = FormatDateTime(Now(), vbLongDate)
                Else
                    strFValue = Format(Now(), strFValue)
                End If
            ElseIf InStr(1, .strText, "[PageNo]") > 0 Then
                strFValue = "Page " & PageNo
                If InStr(1, .strText, "[NumPages]") > 0 Then
                    strFValue = strFValue & " of " & NumPages
                End If
            End If
        ElseIf .Type = cCalcField Then
            If ReportFile.DataBound Then
                Dim strOper As String
                Dim strGetField As String
                Dim strRemainder As String
                Dim strModified As String
                Dim StartNextField As Integer
                Dim EndLastField As Integer
                Dim blnFieldFound As Boolean
                StartNextField = 1
                EndLastField = 1
                strRemainder = Mid(.strText, 3, InStr(1, .strText, "}") - 3) 'parse out stuff in curly brackets
                blnFieldFound = True
                Do While blnFieldFound
                    StartNextField = InStr(1, strRemainder, "[")
                    If StartNextField > 0 Then
                        blnFieldFound = True
                        strModified = strModified & Mid(strRemainder, EndLastField, StartNextField - EndLastField)
                        strRemainder = Right(strRemainder, Len(strRemainder) - StartNextField)
                        strGetField = Left(strRemainder, InStr(1, strRemainder, "]") - 1)
                        EndLastField = InStr(1, strRemainder, "]") + 1
                        For i = 0 To rstData.Fields.count - 1
                            If rstData.Fields(i).Name = strGetField Then
                                strModified = strModified & Str(rstData.Fields(i).value)
                                Exit For
                            End If
                        Next i
                    Else
                        blnFieldFound = False
                    End If
                Loop
            End If
            If Len(strRemainder) > EndLastField - 1 Then
                strModified = strModified & Right(strRemainder, Len(strRemainder) - (EndLastField - 1))
            End If
            strFValue = Eval(strModified)
            SaveCalc(CalcNum).strFieldName = .Fieldname
            SaveCalc(CalcNum).strValue = strFValue
            CalcNum = CalcNum + 1
        ElseIf .Type = cSumField Then
            If ReportFile.DataBound Then
                Dim rst As ADODB.Recordset
                Dim TotalNum As Double
                Dim AvgNum As Double
                Dim MinNum As Double
                Dim MaxNum As Double
                Dim count As Integer
                Dim strFilter As String
                AggFunc = Mid(.strText, 3, 5)
                Set rst = rstData.Clone
                If .SecNo > 5 And .SecNo < 9 Then
                    For i = 0 To (8 - .SecNo)
                        If ReportFile.SortField(i) > "" Then
                            If strFilter > "" Then
                                strFilter = strFilter & " AND " & ReportFile.SortField(i) & " = '" & CStr(rstData.Fields(FieldNo(i)).value) & "'"
                            Else
                                strFilter = ReportFile.SortField(i) & " = '" & CStr(rstData.Fields(FieldNo(i)).value) & "'"
                            End If
                        End If
                    Next i
                    rst.filter = strFilter
                End If
                rst.MoveFirst
                For i = 0 To rst.Fields.count - 1
                    If rst.Fields(i).Name = .Fieldname Then
                        MinNum = rst.Fields(i).value
                        For j = 0 To rst.RecordCount - 1
                            TotalNum = TotalNum + Abs(rst.Fields(i).value)
                            If rst.Fields(i).value < MinNum Then MinNum = rst.Fields(i).value
                            If rst.Fields(i).value > MaxNum Then MaxNum = rst.Fields(i).value
                            count = count + 1
                            If Not rst.EOF Then rst.MoveNext
                        Next j
                        AvgNum = TotalNum / count
                        Select Case AggFunc
                            Case "SumOf": strFValue = TotalNum
                            Case "AvgOf": strFValue = AvgNum
                            Case "MinOf": strFValue = MinNum
                            Case "MaxOf": strFValue = MaxNum
                            Case "CntOf": strFValue = count
                        End Select
                        Exit For
                    End If
                Next i
                Set rst = Nothing
                If blnHasCalcDataFields Then
                    Dim varValue As Variant
                    MinNum = Val(SaveCalc(0).strValue)
                    For i = 0 To UBound(SaveCalc)
                        If SaveCalc(i).strFieldName = .Fieldname Then
                            varValue = Val(SaveCalc(i).strValue)
                            TotalNum = TotalNum + Abs(varValue)
                            If varValue < MinNum Then MinNum = varValue
                            If varValue > MaxNum Then MaxNum = varValue
                            count = count + 1
                        End If
                    Next i
                    AvgNum = TotalNum / count
                    Select Case AggFunc
                        Case "SumOf": strFValue = TotalNum
                        Case "AvgOf": strFValue = AvgNum
                        Case "MinOf": strFValue = MinNum
                        Case "MaxOf": strFValue = MaxNum
                        Case "CntOf": strFValue = count
                    End Select
                End If
            End If
        Else
            If ReportFile.DataBound Then
                If Not rstData.EOF Then
                    For i = 0 To UBound(DataField) - 1
                        If DataField(i, 0) = .Fieldname Then
                            strFValue = rstData.Fields(i).value
                            Exit For
                        End If
                    Next i
                End If
            End If
        End If
        If InStr(1, .strText, "|") > 0 Then
            strGetFormat = Right(.strText, Len(.strText) - (InStr(1, .strText, "|")))
            If strGetFormat = "(Whole)" Then
                strFValue = FormatNumber(strFValue, 0, vbFalse, vbFalse)
            ElseIf Left(strGetFormat, 8) = "(Decimal" Then
                strFValue = FormatNumber(strFValue, Val(Mid(strGetFormat, 10, 2)))
            ElseIf Left(strGetFormat, 8) = "(Percent" Then
                strFValue = FormatPercent(strFValue, Val(Mid(strGetFormat, 10, 2)))
            ElseIf Left(strGetFormat, 9) = "(Currency" Then
                strFValue = FormatCurrency(strFValue)
            ElseIf Left(strGetFormat, 5) = "(Date" Then
                If Mid(strGetFormat, 7, 4) = "wwww" Then
                    strFValue = FormatDateTime(strFValue, vbLongDate)
                Else
                    strFValue = Format(strFValue, Mid(strGetFormat, 7, Len(strGetFormat) - 1 - InStr(1, strGetFormat, ":")))
                End If
            ElseIf Left(strGetFormat, 5) = "(Time" Then
                If Mid(strGetFormat, 7, 5) = "hh:ss" Then
                    strFValue = FormatDateTime(strFValue, vbShortTime)
                ElseIf Mid(strGetFormat, 7, 10) = "h:mm:ss AM" Then
                    strFValue = FormatDateTime(strFValue, vbLongTime)
                End If
            End If
        End If
        If .BckStl = 0 Then
            bkcolor = -1
        Else
            bkcolor = .BckClr
        End If
        
        PP.Pages.ActivePage.DrawShape 0, ReportFile.LMarg + .Left, TopOffset + .Top, _
        ctlwidth, ctlheight, bdrcolor, bkcolor, 1, .BdrStl - 1
        
        PP.Pages.ActivePage.SetFont .FntNam, .FntSiz, .FntBld, .FntItl, .FntUnd, False, 0
        PP.Pages.ActivePage.DrawText strFValue, ReportFile.LMarg + .Left + .BdrWd * 0.01, TopOffset + .Top + .BdrWd * 0.01, _
        ctlwidth - .BdrWd * 0.02, ctlheight - .BdrWd * 0.02, .ForClr, bkcolor, .Align
    
    ElseIf .Type = cBoundImage Then
        If ReportFile.DataBound Then                                    'to dynamically load the picture into the report
            If Not rstData.EOF Then                                     'OLE picture field in Access, which bloats the
                For i = 0 To UBound(DataField) - 1                        'size of the database terribly
                    If DataField(i, 0) = .Fieldname Then
                        strFValue = rstData.Fields(i).value
                        Exit For
                    End If
                Next i
            End If
        End If
        If ReportFile.ImageFolder > "" Then
            PP.Pages.ActivePage.DrawPicture .Left + ReportFile.LMarg, .Top + TopOffset, ctlwidth, ctlheight, LoadPicture(ReportFile.ImageFolder & "\" & strFValue), True
        End If
    ElseIf .Type = cImage Then
        Open App.Path & "\tmpfile" For Binary As #5
        Put 5, , .ImgData
        Close 5
        PP.Pages.ActivePage.DrawPicture ReportFile.LMarg + .Left, .Top + TopOffset, ctlwidth, ctlheight, LoadPicture(App.Path & "\tmpfile"), True
        Kill App.Path & "\tmpfile"
        DoEvents
    ElseIf .Type = cCheckBox Then
        Dim blnValue As Boolean
        If ReportFile.DataBound Then
            If Not rstData.EOF Then                                     'OLE picture field in Access, which bloats the
                For i = 0 To UBound(DataField) - 1                        'size of the database terribly
                    If DataField(i, 0) = .Fieldname Then
                        blnValue = rstData.Fields(i).value
                        Exit For
                    End If
                Next i
            End If
        End If
        PP.Pages.ActivePage.DrawCheckBox .DisplayType, blnValue, .Left + ReportFile.LMarg, _
        .Top + TopOffset, 0.125, 0.166, .ForClr, .BckClr, 1, .Sunken
    End If
    
End With
    
Exit Sub

NoDraw:
    MsgBox "Error in DrawObject : " & Err.Description
    
    
End Sub

Public Sub PreviewWithNoData(GoPreview As Boolean)
'On Error GoTo NoPreview

Dim PgHeadTop As Single
Dim PgFootTop As Single
Dim PageFreeHt As Single
Dim PageFreeWd As Single
Dim DetTop As Single
Dim PageNo As Integer
Dim i As Integer
Dim j As Integer

'set variables for free space heights to be used below
    With ReportFile
        PageFreeHt = .PageHt - .TMarg - .BMarg - .HeaderHt(1) - .FooterHt(1) - .HeaderHt(0) - .FooterHt(0)
        PageFreeWd = .PageWd - .LMarg - .RMarg
    End With

'generate the report page
    Set PP = New Preview
    PageNo = 0
    With PP
        .Cls
        With .Pages
            .ScaleMode = vbInches
            If ReportFile.Orient = cPortrait Then
                .Landscape = False
            Else
                .Landscape = True
            End If
            .width = ReportFile.PageWd
            .Height = ReportFile.PageHt
            .Add
            PageNo = PageNo + 1
            If ReportFile.SectColor(0) <> vbWhite Then
                PP.Pages.ActivePage.DrawShape 0, ReportFile.LMarg, ReportFile.TMarg, _
                PageFreeWd, ReportFile.HeaderHt(0), -1, ReportFile.SectColor(0)
            End If
            PgHeadTop = ReportFile.HeaderHt(0) + ReportFile.TMarg
            If ReportFile.SectColor(1) <> vbWhite Then
                PP.Pages.ActivePage.DrawShape 0, ReportFile.LMarg, PgHeadTop, _
                PageFreeWd, ReportFile.HeaderHt(1), -1, ReportFile.SectColor(1)
            End If
            DetTop = PgHeadTop + ReportFile.HeaderHt(1)
            If ReportFile.SectColor(5) <> vbWhite Then
                PP.Pages.ActivePage.DrawShape 0, ReportFile.LMarg, DetTop, _
                PageFreeWd, ReportFile.DetHt, -1, ReportFile.SectColor(5)
            End If
            If ReportFile.SectColor(9) <> vbWhite Then
                PP.Pages.ActivePage.DrawShape 0, ReportFile.LMarg, DetTop + PageFreeHt, _
                PageFreeWd, ReportFile.FooterHt(9), -1, ReportFile.SectColor(9)
            End If
            If ReportFile.SectColor(10) <> vbWhite Then
                PP.Pages.ActivePage.DrawShape 0, ReportFile.LMarg, DetTop + ReportFile.DetHt, _
                PageFreeWd, ReportFile.FooterHt(0), -1, ReportFile.SectColor(10)
            End If
            For i = 1 To UBound(ReportFile.RpControl)
                If ReportFile.RpControl(i).SecNo = 0 Then
                    DrawObject i, ReportFile.TMarg
                ElseIf ReportFile.RpControl(i).SecNo = 1 Then
                    DrawObject i, PgHeadTop
                ElseIf ReportFile.RpControl(i).SecNo = 5 Then
                    DrawObject i, DetTop
                ElseIf ReportFile.RpControl(i).SecNo = 9 Then
                    DrawObject i, DetTop + PageFreeHt
                ElseIf ReportFile.RpControl(i).SecNo = 10 Then
                    DrawObject i, DetTop + ReportFile.DetHt
                End If
            Next i
        End With
    If GoPreview Then
        .Show
    Else
        .PrintPages
    End If
    End With

'destroy the preview object to free memory
    Set PP = Nothing
    Exit Sub

NoPreview:

    Set PP = Nothing
    MsgBox "Error in PreviewWithNoData : " & Err.Description

End Sub
