VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ExportTXT"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'////////////////////////////////////////////////////////
'///             CVS Exportation Component
'///                 (ExportCSV.cls)
'///_____________________________________________________
'/// Component responsible of CSV (Comma Separated Values)
'/// exportation. Implements the IExport interface.
'///_____________________________________________________
'/// Last modification  : Ago/25/2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Created
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit

Implements IExport

Private m_bBusy As Boolean
Private m_sFileName As String
Private m_sSubject As String
Private m_sSeparator As String

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
    prvTokenize = Replace(prvTokenize, vbCr, "[Cr]")
    prvTokenize = Replace(prvTokenize, vbLf, "[Lf]")
    prvTokenize = Replace(prvTokenize, vbTab, "[Tab]")
End Function

Private Sub Class_Initialize()
    m_sSeparator = ","
End Sub

Private Function IExport_DoAction(ByVal iAction As Integer, ByVal vValue As Variant) As Integer
    Select Case iAction
        Case 1 ' Subject
            m_sSubject = vValue
    End Select
End Function


Private Function IExport_Export(oPages As Pages) As Integer
    Dim LoPage As Page
    Dim LnIdx As Integer
    Dim LsMainBuffer As String
    Dim LsBufferTmp As String
    Dim LnLastSection As Integer
    Dim LnLastBand As Integer
    Dim LrElement As PageElement
    '////////////////////////////////////////
    '/// Inits values
    '////////////////////////////////////////
    LsMainBuffer = String(65536, 0)
    LsMainBuffer = ""
    '////////////////////////////////////////
    '/// Build HTML page header.
    '////////////////////////////////////////
    '////////////////////////////////////////
    '/// Create pages section tables
    '////////////////////////////////////////
    For Each LoPage In oPages
        LsBufferTmp = String(1024, 0)
        LsBufferTmp = "Page " & LoPage.Index & vbCrLf
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
                        If (LnLastSection <> .SectionType) Then
                            LnLastSection = .SectionType
                            LsBufferTmp = LsBufferTmp & vbCrLf '& String(100, "-") & vbCrLf
                            LnLastBand = .BandIndex
                        End If
                        '////////////////////////////////////////
                        '/// verify if is needed to close the row tag
                        '////////////////////////////////////////
                        If (LnLastBand <> .BandIndex) Then
                            If (Right$(LsBufferTmp, 1) = vbTab) Then
                                LsBufferTmp = Left$(LsBufferTmp, Len(LsBufferTmp) - 1)
                            End If
                            LsBufferTmp = LsBufferTmp & vbCrLf
                            LnLastBand = .BandIndex
                        End If
                        '////////////////////////////////////////
                        '/// Encodes item
                        '////////////////////////////////////////
                        Select Case .Type
                            Case 1 ' Text
                                LsBufferTmp = LsBufferTmp & """" & prvTokenize(.Text) & """" & m_sSeparator
                            Case 2 ' Line
'                                If (.Height < .Width) Then
'                                    ' Horizontal Rule
'                                    LsBufferTmp = LsBufferTmp & String(CLng(.Width \ 8), "_")
'                                Else
'                                End If
                            Case 3 ' Box
                            Case 4 ' Picture
                        End Select
                    End With
                Next LnIdx
            End If
        End With
        '////////////////////////////////////////
        '/// Apends page code to the main buffer
        '////////////////////////////////////////
        LsMainBuffer = LsMainBuffer & LsBufferTmp
        'm_oBuffer.AppendString LsBufferTmp
        DoEvents
    Next LoPage
    '////////////////////////////////////////
    '/// Saves page buffer to disk
    '////////////////////////////////////////
    Dim LnFileHandler As Integer
    
    LnFileHandler = FreeFile
    On Error Resume Next
    Kill m_sFileName
    Open m_sFileName For Append As #LnFileHandler
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
        m_sFileName = LsFile & ".csv"
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




