VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'////////////////////////////////////////////////////////
'///            HTML Exportation Component
'///                 (ExportHTML.cls)
'///_____________________________________________________
'/// Component responsible of HTML exportation. Implements
'/// the IExport interface.
'///_____________________________________________________
'/// Last modification  : Ago/15/2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Created
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit

Implements IExport

Private m_oBuffer As clsAppendString

Private m_bBusy As Boolean
Private m_sFileName As String
Private m_sSubject As String

Private Enum FilePathPart
    fpDrive = 1
    fpPath = 2
    fpFileName = 4
    fpExtension = 8
End Enum

Private Enum AlignConstans
    acTLeft = &H0
    acTCenter = &H1
    acTRight = &H2
    acTVCenter = &H4
    acTBottom = &H8
    acTWordBreak = &H10
    acTSingleLine = &H20
    acTNoClip = &H100
End Enum

Private Function prvTokenize(sInput As String) As String
    prvTokenize = sInput
    prvTokenize = Replace(prvTokenize, "&", "&amp;")
    prvTokenize = Replace(prvTokenize, "<", "&lt;")
    prvTokenize = Replace(prvTokenize, ">", "&gt;")
    prvTokenize = Replace(prvTokenize, " ", "&nbsp;")
    prvTokenize = Replace(prvTokenize, vbCrLf, "<BR>")
    prvTokenize = Replace(prvTokenize, vbCr, "<BR>")
    prvTokenize = Replace(prvTokenize, vbLf, "<BR>")
End Function

Private Sub Class_Initialize()
    Set m_oBuffer = New clsAppendString
End Sub

Private Function IExport_DoAction(ByVal iAction As Integer, ByVal vValue As Variant) As Integer
    Select Case iAction
        Case 1 ' Subject
            m_sSubject = vValue
    End Select
End Function

Private Function IExport_Export(oPages As Pages) As Integer
    Dim LoPage As Page
    Dim LoFonts As FontMap
    Dim LoFont As IFont
    Dim LnIdx As Integer
    Dim LsMainBuffer As String
    Dim LsBufferTmp As String
    Dim LsHdrBuffer As String
    Dim LnLastSection As Integer
    Dim LnLastBand As Integer
    Dim LbOpenTable As Boolean
    Dim LrElement As PageElement
    Dim LbIDE As Boolean
    
    '////////////////////////////////////////
    '/// Inits values
    '////////////////////////////////////////
    LbIDE = IsIDE
    Set LoFonts = oPages.FontMap
    Set LoFont = LoFonts.Item(1)
    LoFonts.Create
    LsMainBuffer = String(65536, 0)
    LsMainBuffer = ""
    '////////////////////////////////////////
    '/// Build HTML page header.
    '////////////////////////////////////////
    LsHdrBuffer = String(1024, 0)
    LsHdrBuffer = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2 Final//EN"">" & vbCrLf
    LsHdrBuffer = LsHdrBuffer & "<HTML>" & vbCrLf & "<HEAD>" & vbCrLf & "<TITLE>" & m_sSubject & "</TITLE>" & vbCrLf
    LsHdrBuffer = LsHdrBuffer & "<META NAME=""Generator"" CONTENT=""" & App.ProductName & " v" & App.Major & "." & App.Minor & """>" & vbCrLf
    LsHdrBuffer = LsHdrBuffer & "<META NAME=""Author"" CONTENT=""?"">" & vbCrLf & "<META NAME=""Keywords"" CONTENT=""?"">" & vbCrLf
    LsHdrBuffer = LsHdrBuffer & "<META NAME=""Description"" CONTENT=""?"">" & vbCrLf & "</HEAD>" & vbCrLf & "<BODY>" & vbCrLf
    LsHdrBuffer = LsHdrBuffer & "<A NAME=""TOP""><FONT SIZE=4>Index</FONT></A><BR>" & vbCrLf
    '////////////////////////////////////////
    '/// Create pages section tables
    '////////////////////////////////////////
    For Each LoPage In oPages
        '////////////////////////////////////////
        '/// Adds a hyperlink to jump to the report
        '/// page section in the complete HTML page
        '////////////////////////////////////////
        LsHdrBuffer = LsHdrBuffer & "<A HREF=""#PAGE" & LoPage.Index & """>Page " & LoPage.Index & "</A><BR>" & vbCrLf
        LsBufferTmp = String(1024, 0)
        LsBufferTmp = "<HR SIZE=1 COLOR=""Red"" WIDTH=""100%"">" & vbCrLf
        LsBufferTmp = LsBufferTmp & "<A NAME=""PAGE" & LoPage.Index & """><FONT SIZE=4 COLOR=""Black"">Page " & LoPage.Index & "</FONT></A><BR>" & vbCrLf
        '////////////////////////////////////////
        '/// Now encodes page elements.
        '////////////////////////////////////////
        LnLastSection = -1
        LnLastBand = -1
        With LoPage.Elements
            If (.Count > 0) Then
                For LnIdx = 1 To .Count
                    LrElement = .Item(LnIdx)
                    With LrElement
                        '////////////////////////////////////////
                        '/// Open and close tables as needed.
                        '////////////////////////////////////////
                        If (LnLastSection = .SectionType) Then
                        Else
                            LnLastSection = .SectionType
                            LnLastBand = .BandIndex
                            If LbOpenTable Then
                                LsBufferTmp = LsBufferTmp & "</TR>" & vbCrLf & "</TABLE>" & vbCrLf
                            End If
                            If LbIDE Then
                                LsBufferTmp = LsBufferTmp & "<TABLE BORDER=1 CELLSPACING=0 CELLPADDING=0>" & vbCrLf & "<TR>" & vbCrLf
                            Else
                                LsBufferTmp = LsBufferTmp & "<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0>" & vbCrLf & "<TR>" & vbCrLf
                            End If
                            LbOpenTable = True
                        End If
                        '////////////////////////////////////////
                        '/// verify if is needed to close the row tag
                        '////////////////////////////////////////
                        If (LnLastBand <> .BandIndex) Then
                            LsBufferTmp = LsBufferTmp & "</TR>" & vbCrLf & "<TR>" & vbCrLf
                            LnLastBand = .BandIndex
                        End If
                        '////////////////////////////////////////
                        '/// Encodes item
                        '////////////////////////////////////////
                        Select Case .Type
                            Case 1 ' Text
                                Dim LsTagContent As String
                                
                                LsTagContent = " BAND=""" & .BandIndex & """"
                                If (.BackColor <> -1) Then
                                    LsTagContent = LsTagContent & " BGCOLOR=" & LoFont.HTMLColor(.BackColor)
                                End If
                                LsTagContent = LsTagContent & " WIDTH=""" & Fix(.Width * 1.2) & "px"""
                                ' Horizontal aligment
                                If (.Aligment And acTCenter) Then
                                    LsTagContent = LsTagContent & " ALIGN=""Center"""
                                ElseIf (.Aligment And acTRight) Then
                                    LsTagContent = LsTagContent & " ALIGN=""Right"""
                                Else
                                    LsTagContent = LsTagContent & " ALIGN=""Left"""
                                End If
                                ' Vertical aligment
                                If (.Aligment And acTVCenter) Then
                                    LsTagContent = LsTagContent & " VALIGN=""Middle"""
                                ElseIf (.Aligment And acTBottom) Then
                                    LsTagContent = LsTagContent & " VALIGN=""Bottom"""
                                Else
                                    LsTagContent = LsTagContent & " VALIGN=""Top"""
                                End If
                                LsBufferTmp = LsBufferTmp & "<TD" & LsTagContent & ">"
                                LsBufferTmp = LsBufferTmp & LoFonts.Item(.FontIndex).OpenTag(.ForeColor)
                                If (Len(.Text) = 0) Then
                                    LsBufferTmp = LsBufferTmp & "&nbsp;"
                                Else
                                    LsBufferTmp = LsBufferTmp & prvTokenize(.Text)
                                End If
                                LsBufferTmp = LsBufferTmp & LoFonts.Item(.FontIndex).CloseTag
                                LsBufferTmp = LsBufferTmp & "</TD>" & vbCrLf '"</TD><TD>&nbsp;&nbsp;</TD>" & vbCrLf
                            Case 2 ' Line
                                If (.Height < .Width) Then
                                    ' Horizontal Rule
                                    LsBufferTmp = LsBufferTmp & "<TD><HR COLOR=" & LoFont.HTMLColor(.ForeColor) & _
                                        " SIZE=""" & .Size & """ WIDTH=""" & Fix((.Width - .Left) * 1.2) & "px""></TD>"
                                Else
                                End If
                            Case 3 ' Box
                            Case 4 ' Picture
                        End Select
                        '////////////////////////////////////////
                        '///
                        '////////////////////////////////////////
                    End With
                Next LnIdx
            End If
        End With
        '////////////////////////////////////////
        '/// Closes any open table
        '////////////////////////////////////////
        If LbOpenTable Then
            LsBufferTmp = LsBufferTmp & "</TR>" & vbCrLf & "</TABLE>" & vbCrLf
            LbOpenTable = False
        End If
        LsBufferTmp = LsBufferTmp & "<A HREF=""#TOP"">Top</A><BR>" & vbCrLf
        '////////////////////////////////////////
        '/// Apends page code to the main buffer
        '////////////////////////////////////////
        LsMainBuffer = LsMainBuffer & LsBufferTmp
        'm_oBuffer.AppendString LsBufferTmp
        DoEvents
    Next LoPage
    '////////////////////////////////////////
    '/// Finalize adding closing tags to page
    '////////////////////////////////////////
    LsMainBuffer = LsMainBuffer & "<HR SIZE=1 COLOR=""RED"" WIDTH=""100%"">" & vbCrLf
    LsMainBuffer = LsMainBuffer & "</BODY>" & vbCrLf & "</HTML>"
'    With m_oBuffer
'        .AppendString "<HR SIZE=1 COLOR=""RED"" WIDTH=""100%"">" & vbCrLf
'        .AppendString "</BODY>" & vbCrLf & "</HTML>"
'    End With
    '////////////////////////////////////////
    '/// Saves page buffer to disk
    '////////////////////////////////////////
    Dim LnFileHandler As Integer
    
    LnFileHandler = FreeFile
'    If (Len(m_sFileName) = 0) Then
'        ' Get temporal file name....
'        ' Next line for testing only
'        m_sFileName = "C:\ARExportTest.html"
'    End If
    On Error Resume Next
    Kill m_sFileName
    Open m_sFileName For Append As #LnFileHandler
    Print #LnFileHandler, LsHdrBuffer
    Print #LnFileHandler, LsMainBuffer 'm_oBuffer.ReturnString
    Close #LnFileHandler
End Function

Private Function prvParseFileName(ByVal sTempPath As String, _
    iReturnType As FilePathPart) As String
    Dim LsDrive As String
    Dim LsPath As String
    Dim LsFileName As String
    Dim LsExtension As String
    Dim LnPathLength As Integer
    Dim LnThisLength As Integer
    Dim LnOffset As Integer
    Dim LbFileNameFound As Boolean

    LsDrive = ""
    LsPath = ""
    LsFileName = ""
    LsExtension = ""

    If Mid(sTempPath, 2, 1) = ":" Then ' Find the drive letter.
        LsDrive = UCase(Left(sTempPath, 2))
        sTempPath = Mid(sTempPath, 3)
    ElseIf (Left(sTempPath, 2) = "\\") Then
        Dim LnPos As Integer
        LnPos = InStr(3, sTempPath, "\")
        LsDrive = Left(sTempPath, LnPos - 1)
        sTempPath = Mid(sTempPath, LnPos)
    End If

    LnPathLength = Len(sTempPath)

    For LnOffset = LnPathLength To 1 Step -1 ' Find the next delimiter.
        Select Case Mid(sTempPath, LnOffset, 1)
            Case ".": ' This indicates either an LsExtension or a . or a ..
                LnThisLength = Len(sTempPath) - LnOffset
                If LnThisLength >= 1 Then ' LsExtension
                    LsExtension = Mid(sTempPath, LnOffset, LnThisLength + 1)
                End If
                sTempPath = Left(sTempPath, LnOffset - 1)
            Case "\": ' This indicates a path delimiter.
                LnThisLength = Len(sTempPath) - LnOffset
                If LnThisLength >= 1 Then ' Filename
                    LsFileName = Mid(sTempPath, LnOffset + 1, LnThisLength)
                    sTempPath = Left(sTempPath, LnOffset)
                    LbFileNameFound = True
                    Exit For
                End If
            Case Else
        End Select
    Next LnOffset

    If LbFileNameFound Then
        LsPath = sTempPath
    Else
        LsFileName = sTempPath
    End If

    prvParseFileName = ""
    If iReturnType And fpDrive Then prvParseFileName = LsDrive
    If iReturnType And fpPath Then prvParseFileName = prvParseFileName & LsPath
    If iReturnType And fpFileName Then prvParseFileName = prvParseFileName & LsFileName
    If iReturnType And fpExtension Then prvParseFileName = prvParseFileName & LCase(LsExtension)
'    Debug.Print "ParseFilename-> " & ParseFilename
End Function

Private Property Let IExport_FileName(ByVal RHS As String)
    Dim LsFile As String
    
    If (Len(RHS) > 0) Then
        LsFile = prvParseFileName(RHS, fpDrive Or fpPath Or fpFileName)
        m_sFileName = LsFile & ".htm"
    End If
End Property

Private Property Get IExport_FileName() As String
    IExport_FileName = m_sFileName
End Property

Private Function IExport_QueryInfo(iType As Integer) As Variant
    Select Case iType
        Case 0 ' Status
            IExport_QueryInfo = m_bBusy
        Case 1 ' File name
            IExport_QueryInfo = m_sFileName
        Case 2 ' ???
    End Select
End Function

Private Property Get IExport_StillWorking() As Boolean
    IExport_StillWorking = m_bBusy
End Property


