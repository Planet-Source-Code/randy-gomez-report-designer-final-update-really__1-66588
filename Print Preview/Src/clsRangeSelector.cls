VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRangeSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'////////////////////////////////////////////////////////
'///                 Range Selector Class
'///                (clsRangeSelector.cls)
'///_____________________________________________________
'/// This class lets you specify a range descriptor string
'/// to indicate a pages range to be printed.
'/// (i.e. "1-2,5,10-*"), wildcard use is allowed and the
'/// keyword "All" to specify complete range.
'///_____________________________________________________
'/// Last modification  : Ago/10/2000
'/// Last modified by   : Leontti R.
'/// modification reason: Created
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit
Option Compare Text
Option Base 0

Private m_sSelection As String
Private m_bAll As Boolean
Private m_asRanges() As RangeLimit
Private m_lCount As Long

Private Const PS_ALL = "All"

Private Type RangeLimit
    First As Long
    Last As Long
    Descr As String
    IsRange As Boolean
    Enabled As Boolean
End Type

Private Const PN_HIGH_LIMIT = 2147483646
Private Const PN_LOW_LIMIT = 1

Public RaiseEvents As Boolean
Public Event Changed(sSelectionStr As String)
Public Sub AddRange(lFrom As Long, Optional lTo As Long)
    m_lCount = m_lCount + 1
    ReDim Preserve m_asRanges(m_lCount)
    With m_asRanges(m_lCount)
        .Enabled = True
        .First = lFrom
        If (lTo = 0) Then
            .Last = PN_HIGH_LIMIT
        Else
            .Last = lTo
        End If
    End With
    prvRangesOptimize
End Sub

Public Sub Clear()
    m_bAll = True
    m_sSelection = PS_ALL
    m_lCount = 0
    ReDim m_asRanges(m_lCount)
End Sub

Public Function ElementInRange(ByVal lIndex As Long) As Boolean
    If m_bAll Then
        ElementInRange = True
    Else
        ElementInRange = prvRangesSeek(lIndex)
    End If
End Function

Public Property Get RangeMin() As Long
    If m_bAll Then
        RangeMin = PN_LOW_LIMIT
    Else
        RangeMin = m_asRanges(1).First
    End If
End Property

Public Property Get RangeMax() As Long
    If m_bAll Then
        RangeMax = PN_HIGH_LIMIT
    Else
        RangeMax = m_asRanges(m_lCount).Last
    End If
End Property


Private Sub prvRangesAnalizeDescr(ByRef rRangeLimit As RangeLimit)
    With rRangeLimit
        .Descr = Trim(.Descr)
        ' Valid by default
        .Enabled = True
        ' If only has a dash char or is empty, is an invalid entry
        If ((.Descr = "-") Or (Len(.Descr) = 0)) Then
            .Enabled = False
            .Descr = ""
        Else
            Dim LnDivPos As Integer
            Dim LnNxtPos As Integer
            ' Checks if descriptor string has a range mark (-)
            LnDivPos = InStr(1, .Descr, "-")
            If (LnDivPos > 0) Then
                .IsRange = True
                LnNxtPos = (LnDivPos + 1)
            Else
                LnDivPos = InStr(1, .Descr, " to ", vbTextCompare)
                If (LnDivPos > 0) Then
                    .IsRange = True
                    LnNxtPos = (LnDivPos + 4)
                End If
            End If
            If .IsRange Then
            ' If is a range, parse first and last limits
                Dim LsFirst As String
                Dim LsLast As String
                
                LsFirst = RTrim(Left$(.Descr, LnDivPos - 1))
                LsLast = LTrim(Mid$(.Descr, LnNxtPos))
                ' Evaluates limits
                If (Len(LsFirst) = 0) Or (LsFirst = "*") Then
                    .First = PN_LOW_LIMIT
                Else
                    .First = Val(LsFirst)
                End If
                If (Len(LsLast) = 0) Or (LsLast = "*") Then
                    .Last = PN_HIGH_LIMIT
                Else
                    .Last = Val(LsLast)
                End If
                ' Check any special condition
                ' Invalid order
                If (.First > .Last) Then
                    ' Swaps values
                    Dim LnTemp As Long
                    
                    LnTemp = .First
                    .First = .Last
                    .Last = LnTemp
                ElseIf (.First = .Last) Then
                    ' Single value
                    .IsRange = False
                ElseIf ((.First = 0) Or (.Last = 0)) Then
                    ' Any other invalid chars
                    .Enabled = False
                End If
            Else
            ' If doesn't have a range mark (-), is a single number or a
            ' special indication (*=All, empty=Invalid mark)
                If (.Descr = "*") Then
                    ' All numbers mark, overrides all other ranges
                    .First = PN_LOW_LIMIT
                    .Last = PN_HIGH_LIMIT
                    .IsRange = True
                ElseIf (Val(.Descr) = 0) Then
                    ' Alpha char is invalid
                    .Enabled = False = False
                Else
                    ' Otherwise is a number
                    .First = Val(.Descr)
                    .Last = .First
                End If
            End If
        End If
    End With
End Sub

Private Function prvRangesCleanUp() As Boolean
    Dim LnIdx As Integer

    For LnIdx = m_lCount To 1 Step -1
        If Not m_asRanges(LnIdx).Enabled Then
            Call prvRemoveItem(LnIdx)
        End If
    Next
    For LnIdx = 1 To m_lCount
        With m_asRanges(LnIdx)
            If .Enabled Then
                .IsRange = (.First < .Last)
                If .IsRange Then
                    If (.Last = PN_HIGH_LIMIT) Then
                        If (.First = PN_LOW_LIMIT) Then
                            .Descr = "*"
                            prvRangesCleanUp = True
                        Else
                            .Descr = CStr(.First) & "-*"
                        End If
                    Else
                        .Descr = CStr(.First) & "-" & CStr(.Last)
                    End If
                Else
                    .Descr = CStr(.First)
                End If
            End If
        End With
    Next
End Function

Private Function prvGetSelectionString() As String
    Dim LnIdx As Integer
    
    For LnIdx = 1 To m_lCount
        With m_asRanges(LnIdx)
            If .Enabled Then
                prvGetSelectionString = prvGetSelectionString & .Descr & ", "
            End If
        End With
    Next
    prvGetSelectionString = Left(prvGetSelectionString, Len(prvGetSelectionString) - 2)
End Function

Private Function prvRangesMerge(rMain As RangeLimit, rSub As RangeLimit) As Boolean
    ' By default we disable sub range
    prvRangesMerge = True
    If (rSub.Last = (rMain.First - 1)) Then
    '       |5|6|7|8|9 -- Main
    '1|2|3|4|          -- Sub
    ' Just merge ranges into main
        rMain.First = rSub.First
    ElseIf (rSub.First = (rMain.Last + 1)) Then
    '1|2|3|4|          -- Main
    '       |5|6|7|8|9 -- Sub
    ' Just merge ranges into main
        rMain.Last = rSub.Last
    ElseIf ((rSub.First >= rMain.First) And (rSub.Last <= rMain.Last)) Then
    '1|2|3|4|5|6|7|8|9 -- Main
    '    3|4|5|6|7     -- Sub
    ' Just keep sub disabled
    ElseIf ((rMain.First >= rSub.First) And (rMain.Last <= rSub.Last)) Then
    '    3|4|5|6|7     -- Main
    '1|2|3|4|5|6|7|8|9 -- Sub
    ' Make main the bigger one
        rMain.First = rSub.First
        rMain.Last = rSub.Last
    ElseIf ((rMain.First >= rSub.First) And (rMain.First <= rSub.Last) _
         And (rMain.Last >= rSub.Last)) Then
    '     |4|5|6|7|8|9 -- Main
    '1|2|3|4|5|6|      -- Sub
    ' Merge ranges into main
        rMain.First = rSub.First
    ElseIf ((rSub.First >= rMain.First) And (rSub.First <= rMain.Last) _
         And (rSub.Last >= rMain.Last)) Then
    '1|2|3|4|5|6|      -- Main
    '     |4|5|6|7|8|9 -- Sub
    ' Merge ranges into main
        rMain.Last = rSub.Last
    Else
        prvRangesMerge = False
    End If
    'Disable sub
    rSub.Enabled = (Not prvRangesMerge)
End Function

Private Sub prvRangesOptimize()
    Dim LbDirty  As Boolean
    
    Do
        LbDirty = prvRangesReduce
    Loop While LbDirty
    m_bAll = prvRangesCleanUp
    prvRangesSort
    If m_bAll Then
        m_sSelection = PS_ALL
    Else
        m_sSelection = prvGetSelectionString
    End If
    If RaiseEvents Then RaiseEvent Changed(m_sSelection)
End Sub

Private Sub prvAnalizeSelectionString(sSelection As String)
    On Error Resume Next
    Dim LsAddPages As String
    
    If ((sSelection = PS_ALL) Or (sSelection = "*")) Then
        Clear
    Else
        Dim LasDescrs() As String
        Dim LnIdx As Long
        
        LasDescrs = Split(sSelection, ",")
        m_lCount = UBound(LasDescrs)
        m_lCount = m_lCount + 1
        ReDim m_asRanges(m_lCount)
        For LnIdx = 1 To m_lCount
            With m_asRanges(LnIdx)
                .Descr = LasDescrs(LnIdx - 1)
                prvRangesAnalizeDescr m_asRanges(LnIdx)
            End With
        Next
        prvRangesOptimize
    End If
End Sub


Private Function prvRangesReduce() As Boolean
    Dim LnIndex As Integer
    Dim LnIdx As Integer
    
    If (m_lCount > 1) Then
        For LnIndex = 1 To (m_lCount - 1)
            If m_asRanges(LnIndex).Enabled Then
                For LnIdx = (LnIndex + 1) To m_lCount
                    If m_asRanges(LnIdx).Enabled Then
                        If prvRangesMerge(m_asRanges(LnIndex), _
                            m_asRanges(LnIdx)) Then
                            prvRangesReduce = True
                        End If
                    End If
                Next LnIdx
            End If
        Next LnIndex
    End If
End Function

Private Function prvRemoveItem(iIndex As Integer) As Integer
    Dim LnIdx As Long
    
    For LnIdx = iIndex To (m_lCount - 1)
        m_asRanges(LnIdx) = m_asRanges(LnIdx + 1)
    Next
    m_lCount = UBound(m_asRanges)
    m_lCount = (m_lCount - 1)
    ReDim Preserve m_asRanges(m_lCount)
    prvRemoveItem = m_lCount
End Function

Private Function prvRangesSeek(lIndex As Long) As Boolean
    Dim LnIdx As Integer

    For LnIdx = 1 To m_lCount
        With m_asRanges(LnIdx)
            prvRangesSeek = ((lIndex >= .First) And (lIndex <= .Last))
            If prvRangesSeek Then
                Exit For
            End If
        End With
    Next
End Function

Private Sub prvRangesSort()
    Dim LrTemp As RangeLimit
    Dim LbSorted As Boolean
    Dim LnIdx As Long

    LbSorted = False
    Do While Not LbSorted
        LbSorted = True
        For LnIdx = (m_lCount - 1) To 1 Step -1
            If m_asRanges(LnIdx + 1).First < m_asRanges(LnIdx).First Then
                DoEvents
                LbSorted = False
                LrTemp = m_asRanges(LnIdx)
                m_asRanges(LnIdx) = m_asRanges(LnIdx + 1)
                m_asRanges(LnIdx + 1) = LrTemp
            End If
      Next LnIdx
    Loop
End Sub

Public Property Get SelectionString() As String
    SelectionString = m_sSelection
    Debug.Print "Output--> " & SelectionString
End Property

Public Property Let SelectionString(sValue As String)
    Debug.Print "Input--> " & sValue
    sValue = Trim(sValue)
    If (Len(sValue) = 0) Then
        Exit Property
    Else
        prvAnalizeSelectionString sValue
    End If
End Property

Private Sub Class_Initialize()
    m_bAll = True
    m_sSelection = PS_ALL
End Sub


