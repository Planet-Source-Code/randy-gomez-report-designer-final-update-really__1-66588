VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSortPageElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'////////////////////////////////////////////////////////
'///            Page Elements Sorting Class
'///             (clsSortPageElements.cls)
'///_____________________________________________________
'/// This class sorts runtime page elements to be used in
'///      Y  Z     a serial format (as HTML) to ensure the
'///      | /      right position in the X and Y axis based
'///      |/       in the left and top values overriding
'/// X----+----    the Z order assigned in design time (at
'///      |        design time objects are sorted over Z axis)
'///_____________________________________________________
'/// Last modification  : Sep/01/2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Minor bug fixes.
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit
Public Align2Top As Boolean
Private m_nToleranceRatio As Single

Public Function GetBands(oItems As Collection) As Collection
    On Error GoTo ERR_H
    Dim LoBand As Collection
    ' First, the collection is sorted over the Y axis.
    prvSortAxis oItems, False
    ' Then extracts the collection bands.
    Set GetBands = New Collection
    Do
        Set LoBand = prvGetBand(oItems, (GetBands.Count + 1))
        ' If a band was found, adds it to the bands collection
        If (Not (LoBand Is Nothing)) Then
            ' The band is resorted over X axis.
            Call prvSortAxis(LoBand, True)
            GetBands.Add LoBand
        End If
    Loop Until ((oItems.Count = 0) Or (GetBands.Count > 100))
    ' No more than 100 bands...(temporary)
    Exit Function
ERR_H:
    Me.RaiseErr Err.Number, "GetBands", Err.Description
    'Resume
End Function

Private Sub prvSortAxis(oColl As Collection, bHorizAxis As Boolean)
    On Error GoTo ERR_H
    Dim LnIdx As Long
    Dim LnIdx2 As Long
    Dim LnGap As Long
    Dim LbResult As Boolean
    Dim LrTmpElm As PageElement
    Dim LrTmpElm2 As PageElement
    
    LnGap = (oColl.Count / 2)
    Do While LnGap > 0
        For LnIdx = LnGap To (oColl.Count - 1)
            LrTmpElm = oColl(LnIdx + 1)
            LnIdx2 = LnIdx
            If bHorizAxis Then
                LbResult = (LrTmpElm.Left < oColl(LnIdx2 - LnGap + 1).Left)
            Else
                LbResult = (LrTmpElm.Top < oColl(LnIdx2 - LnGap + 1).Top)
            End If
            Do While ((LnIdx2 >= LnGap) And LbResult)
                LrTmpElm2 = oColl(LnIdx2 - LnGap + 1)
                prvSortSwap oColl, LrTmpElm2, LnIdx2
                LnIdx2 = (LnIdx2 - LnGap)
                If (LnIdx2 >= LnGap) Then
                    If bHorizAxis Then
                        LbResult = (LrTmpElm.Left < oColl(LnIdx2 - LnGap + 1).Left)
                    Else
                        LbResult = (LrTmpElm.Top < oColl(LnIdx2 - LnGap + 1).Top)
                    End If
                End If
            Loop
            prvSortSwap oColl, LrTmpElm, LnIdx2
        Next
        LnGap = (LnGap / 2)
    Loop
    Exit Sub
ERR_H:
    Me.RaiseErr Err.Number, "prvSortAxis", Err.Description
End Sub

Private Function prvGetBand(oItems As Collection, ByVal iBandIdx As Long) As Collection
    On Error GoTo ERR_H
    Dim LrElem As PageElement
    Dim LnLimit As Long
    Dim LnTop As Long
    Dim LnIdx As Long
    
    ' Gets the band limits, if there are no elements to the band, it
    ' will return false
    If (oItems.Count > 0) Then
        LnIdx = 1
        ' Gets first element reference
        LrElem = oItems.Item(LnIdx)
        With LrElem
            ' Gets the first element top position
            LnTop = .Top
            ' Calculates the bottom limit to be part of the band
            If (.Type = 2) Then ' Type 2 is line
                ' When is a line, Height means Y2
                LnLimit = (LnTop + CLng(m_nToleranceRatio * (.Height - .Top)))
            Else
                LnLimit = (LnTop + CLng(m_nToleranceRatio * .Height))
            End If
        End With
        ' Creates new band element
        Set prvGetBand = New Collection
        ' Loops to get band elements
        Do
            LrElem = oItems.Item(LnIdx)
            ' If the top position is located inside the limits
            ' that means is part of the band.
            If (LrElem.Top >= LnTop) And _
                (LrElem.Top <= LnLimit) Then
                'Debug.Print "Top:" & LrElem.Top & " Band:" & iBandIdx & ":" & LrElem.Text
                ' If must be realigned, does it.
                If Align2Top Then
                    LrElem.Top = LnTop
                End If
                ' Marks item with the band index
                LrElem.BandIndex = iBandIdx
                ' Adds the item to the band
                prvGetBand.Add LrElem
                ' Removes the otem from the original collection
                oItems.Remove LnIdx
                DoEvents
            Else
                ' Because the items are already sorted over y axis,
                ' we can break now to save time
                'LnIdx = (LnIdx + 1)
                Exit Do
            End If
        Loop Until (LnIdx > oItems.Count)
        LnTop = LnLimit
    End If
EXIT_WHILE:
    Exit Function
ERR_H:
    Me.RaiseErr Err.Number, "prvGetBand", Err.Description
End Function

Private Sub prvSortSwap(oColl As Collection, _
    LrElem As PageElement, iIdx As Long)
    oColl.Remove (iIdx + 1)
    If ((iIdx + 1) > oColl.Count) Then
        oColl.Add LrElem
    Else
        oColl.Add LrElem, , (iIdx + 1)
    End If
End Sub

Public Sub SortElements(ByRef oColl As Collection)
    On Error GoTo ERR_H
    Dim LoBands As Collection
    Dim LoBand As Collection
    Dim LrElem As PageElement
    Dim LnIdx As Long
    
    Set LoBands = GetBands(oColl)
    For Each LoBand In LoBands
        If (LoBand.Count > 0) Then
            For LnIdx = 1 To LoBand.Count
                oColl.Add LoBand.Item(LnIdx)
            Next LnIdx
        End If
    Next LoBand
    Set LoBand = Nothing
    Set LoBands = Nothing
    Exit Sub
ERR_H:
    Me.RaiseErr Err.Number, "SortElements", Err.Description
    'Resume
End Sub

Public Property Get ToleranceRatio() As Single
    ToleranceRatio = m_nToleranceRatio
End Property

Public Property Let ToleranceRatio(nRatio As Single)
    m_nToleranceRatio = nRatio
    If (m_nToleranceRatio > 1) Then
        m_nToleranceRatio = 1
    End If
    If (m_nToleranceRatio < 0) Then
        m_nToleranceRatio = 0
    End If
End Property

Private Sub Class_Initialize()
    m_nToleranceRatio = 0.4
End Sub

Friend Sub RaiseErr(ByVal lErrNum As RSErrorCode, Optional sRoutineName As String, _
    Optional sDescription As String)
    RaiseError lErrNum, TypeName(Me), sRoutineName, sDescription, Erl
End Sub


