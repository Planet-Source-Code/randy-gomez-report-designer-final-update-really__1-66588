Attribute VB_Name = "mUndo"
'type descriptor for undo information
Public Type UndoStep
    UndoCategory As Integer
    Type As Integer
End Type

'array for storing undo information
Public UndoList() As UndoStep
'stores current position in Undo list
Public CurrUndoPos As Integer
Public blnFirstUndo As Boolean
Public LastCat As Integer
Public LastType As Integer

'undo category constants
Public Const unControl = 1
Public Const unSection = 2
Public Const unPage = 3

'undo step type constants
Public Const unPlace = 1
Public Const unMove = 2
Public Const unResize = 3
Public Const unFormat = 4
Public Const unEdit = 5
Public Const unCut = 6
Public Const unPaste = 7
Public Const unSendBack = 8
Public Const unBringFront = 9
Public Const unDelete = 10
Public Const unSectWidth = 11
Public Const unSectHeight = 12
Public Const unSectColor = 13
Public Const unPageSize = 14
Public Const unPageOrient = 15
Public Const unPageMargin = 16

'type for undone control
Public Type ControlUndoInfo
    UndoIDNo As Long
    ctl As ControlInfo
End Type

'array to store undone control information
Public UndoCtl() As ControlUndoInfo

'type for undo page section
Public Type SectUndoInfo
    UndoIDNo As Long
    SectNo As Integer
    SectWidth As Single
    SectHeight As Single
    SectColor As Long
End Type

'array to store undo page section information
Public UndoSect() As SectUndoInfo

'type for undo page information
Public Type PageUndoInfo
    UndoIDNo As Long
    PageWidth As Single
    PageHeight As Single
    PageOrient As Integer
    PageLMarg As Single
    PageTMarg As Single
    PageRMarg As Single
    PageBMarg As Single
End Type

'array to store undo page information
Public UndoPage() As PageUndoInfo

Public Sub InitUndoArrays()

    ReDim UndoList(0)
    ReDim UndoCtl(0)
    ReDim UndoSect(0)
    ReDim UndoPage(0)
    CurrUndoPos = 0

End Sub

Public Sub WriteToUndoList(UndoCat As Integer, UndoType As Integer)
Dim i As Integer
Dim NewPos As Integer
Dim NewNum As Integer

    NewPos = UBound(UndoList) + 1
    CurrUndoPos = NewPos
    
    ReDim Preserve UndoList(NewPos)

    UndoList(NewPos).UndoCategory = UndoCat
    UndoList(NewPos).Type = UndoType
        
    If UndoCat = unControl Then
        For i = 0 To UBound(UndoCtl)
            If UndoCtl(i).UndoIDNo > CurrUndoPos Then
                ReDim Preserve UndoCtl(i)
                Exit For
            End If
        Next i
        If blnControlSelected Then
            NewNum = UBound(UndoCtl) + 1
            ReDim Preserve UndoCtl(NewNum)
            UndoCtl(NewNum).UndoIDNo = NewPos
            StoreUndoControlInfo ctlActive, NewNum
        ElseIf blnGroupSelected Then
            For i = 0 To UBound(SelectedCtl)
                NewNum = UBound(UndoCtl) + 1
                ReDim Preserve UndoCtl(NewNum)
                UndoCtl(NewNum).UndoIDNo = NewPos
                StoreUndoControlInfo SelectedCtl(i).ctl, NewNum
            Next i
        End If
    
    ElseIf UndoCat = unSection Then
        NewNum = UBound(UndoSect) + 1
        ReDim Preserve UndoSect(NewNum)
        UndoSect(NewNum).UndoIDNo = NewPos
        If UndoList(NewPos).Type = unSectHeight Then
            If Not blnOnLastSection Then
                If CurrSection > 0 Then
                    For i = (CurrSection - 1) To FirstSectionVis Step -1
                        If frmDesign.picSection(i).Visible Then
                            Exit For
                        End If
                    Next i
                Else
                    i = 0
                End If
            Else
                i = 10
            End If
            UndoSect(NewNum).SectNo = i
            UndoSect(NewNum).SectHeight = frmDesign.picSection(i).Height
        ElseIf UndoList(NewPos).Type = unSectColor Then
            UndoSect(NewNum).SectHeight = frmDesign.picSection(CurrSection).Height
            UndoSect(NewNum).SectNo = CurrSection
            UndoSect(NewNum).SectColor = frmDesign.picSection(CurrSection).BackColor
        ElseIf UndoList(NewPos).Type = unSectWidth Then
            UndoSect(NewNum).SectWidth = frmDesign.picSection(CurrSection).width
        End If
        
    ElseIf UndoCat = unPage Then
        NewNum = UBound(UndoPage) + 1
        ReDim Preserve UndoPage(NewNum)
        UndoPage(NewNum).UndoIDNo = NewPos
        UndoPage(NewNum).PageWidth = PageWd
        UndoPage(NewNum).PageHeight = PageHt
        UndoPage(NewNum).PageOrient = PageOrient
        UndoPage(NewNum).PageLMarg = LeftMarg
        UndoPage(NewNum).PageRMarg = RightMarg
        UndoPage(NewNum).PageTMarg = TopMarg
        UndoPage(NewNum).PageBMarg = BottomMarg
    End If
    
    blnFirstUndo = True
    LastType = UndoType
    LastCat = UndoCat
    
End Sub

Private Sub StoreUndoControlInfo(GetCtl As Control, GetUndoNo As Integer)
Dim FLen As Long

    
    With UndoCtl(GetUndoNo).ctl
        .Name = GetCtl.Name
        .Index = GetCtl.Index
        .SecNo = GetCtl.Tag
        If TypeOf GetCtl Is Line Then
            .Type = cLine
            .X1 = GetCtl.X1
            .Y1 = GetCtl.Y1
            .X2 = GetCtl.X2
            .Y2 = GetCtl.Y2
            .BdrClr = GetCtl.BorderColor
            .BdrStl = GetCtl.BorderStyle
            .BdrWd = GetCtl.BorderWidth
        ElseIf TypeOf GetCtl Is Shape Then
            .Type = cBox
            .Left = GetCtl.Left
            .Top = GetCtl.Top
            .width = GetCtl.width
            .Height = GetCtl.Height
            .BckClr = GetCtl.BackColor
            .BckStl = GetCtl.BackStyle
            .BdrClr = GetCtl.BorderColor
            .BdrStl = GetCtl.BorderStyle
            .BdrWd = GetCtl.BorderWidth
            .DisplayType = GetCtl.Shape
        ElseIf TypeOf GetCtl Is Label Then
            .Type = GetCtl.LinkTimeout
            .Left = GetCtl.Left
            .Top = GetCtl.Top
            .width = GetCtl.width
            .Height = GetCtl.Height
            .strText = GetCtl.Caption
            .Align = GetCtl.Alignment
            .BckStl = GetCtl.BackStyle
            .BckClr = GetCtl.BackColor
            .BdrStl = GetCtl.BorderStyle
            .FntNam = GetCtl.FontName
            .FntBld = GetCtl.FontBold
            .FntItl = GetCtl.FontItalic
            .FntUnd = GetCtl.FontUnderline
            .FntSiz = GetCtl.FontSize
            .ForClr = GetCtl.ForeColor
            .Fieldname = GetCtl.DataField
        ElseIf TypeOf GetCtl Is MSForms.Image Then
            .Type = cImage
            .Left = GetCtl.Left
            .Top = GetCtl.Top
            .width = GetCtl.width
            .Height = GetCtl.Height
            .BckClr = GetCtl.BackColor
            .BdrStl = GetCtl.BorderStyle
            SavePicture GetCtl.Picture, App.Path & "\tmpfile"
            Open App.Path & "\tmpfile" For Binary As #1
            FLen = LOF(1)
            .ImgData = String$(FLen, " ")
            Get #1, 1, .ImgData
            Close #1
            Kill App.Path & "\tmpfile"
        ElseIf TypeOf GetCtl Is CheckBoxControl Then
            .Type = cCheckBox
            .Left = GetCtl.Left
            .Top = GetCtl.Top
            .width = GetCtl.width
            .Height = GetCtl.Height
            .BckClr = GetCtl.BackColor
            .BdrClr = GetCtl.BorderColor
            .DisplayType = GetCtl.DisplayType
            .Sunken = GetCtl.Sunken
            .Fieldname = GetCtl.DataField
        End If
    End With

End Sub

Public Sub RestoreFromUndoList()
Dim i As Integer, j As Integer
    
    ReDim SelectedCtl(0)
    blnGroupSelected = False
    
    If CurrUndoPos > 0 Then
        CurrUndoPos = CurrUndoPos - 1
    End If
    
    If UndoList(CurrUndoPos).UndoCategory = unControl Then
        For i = 1 To UBound(UndoCtl)
            If UndoCtl(i).UndoIDNo = CurrUndoPos Then
                If UndoList(CurrUndoPos).Type = unDelete Then
                    ReCreateControl UndoCtl(i).ctl
                Else
                    For j = 0 To frmDesign.Controls.count - 1
                        If frmDesign.Controls(j).Tag = UndoCtl(i).ctl.SecNo Then
                            If frmDesign.Controls(j).Name = UndoCtl(i).ctl.Name And _
                                frmDesign.Controls(j).Index = UndoCtl(i).ctl.Index Then
                                If UndoList(CurrUndoPos).Type = unPlace Then
                                    Unload frmDesign.Controls(j)
                                ElseIf UndoList(CurrUndoPos).Type = unMove Then
                                    If TypeOf frmDesign.Controls(j) Is Line Then
                                        frmDesign.Controls(j).X1 = UndoCtl(i).ctl.X1
                                        frmDesign.Controls(j).Y1 = UndoCtl(i).ctl.Y1
                                        frmDesign.Controls(j).X2 = UndoCtl(i).ctl.X2
                                        frmDesign.Controls(j).Y2 = UndoCtl(i).ctl.Y2
                                    Else
                                        frmDesign.Controls(j).Left = UndoCtl(i).ctl.Left
                                        frmDesign.Controls(j).Top = UndoCtl(i).ctl.Top
                                    End If
                                ElseIf UndoList(CurrUndoPos).Type = unResize Then
                                    If TypeOf frmDesign.Controls(j) Is Line Then
                                        frmDesign.Controls(j).X1 = UndoCtl(i).ctl.X1
                                        frmDesign.Controls(j).Y1 = UndoCtl(i).ctl.Y1
                                        frmDesign.Controls(j).X2 = UndoCtl(i).ctl.X2
                                        frmDesign.Controls(j).Y2 = UndoCtl(i).ctl.Y2
                                    Else
                                        frmDesign.Controls(j).Left = UndoCtl(i).ctl.Left
                                        frmDesign.Controls(j).Top = UndoCtl(i).ctl.Top
                                        frmDesign.Controls(j).width = UndoCtl(i).ctl.width
                                        frmDesign.Controls(j).Height = UndoCtl(i).ctl.Height
                                    End If
                                ElseIf UndoList(CurrUndoPos).Type = unFormat Then
                                    If TypeOf frmDesign.Controls(j) Is Line Then
                                        frmDesign.Controls(j).BorderColor = UndoCtl(i).ctl.BdrClr
                                        frmDesign.Controls(j).BorderStyle = UndoCtl(i).ctl.BdrStl
                                        frmDesign.Controls(j).BorderWidth = UndoCtl(i).ctl.BdrWd
                                    ElseIf TypeOf frmDesign.Controls(j) Is Shape Then
                                        frmDesign.Controls(j).BorderColor = UndoCtl(i).ctl.BdrClr
                                        frmDesign.Controls(j).BorderStyle = UndoCtl(i).ctl.BdrStl
                                        frmDesign.Controls(j).BorderWidth = UndoCtl(i).ctl.BdrWd
                                        frmDesign.Controls(j).BackColor = UndoCtl(i).ctl.BckClr
                                        frmDesign.Controls(j).BackStyle = UndoCtl(i).ctl.BckStl
                                        frmDesign.Controls(j).Shape = UndoCtl(i).ctl.DisplayType
                                    ElseIf TypeOf frmDesign.Controls(j) Is Label Then
                                        frmDesign.Controls(j).BorderStyle = UndoCtl(i).ctl.BdrClr
                                        frmDesign.Controls(j).BackColor = UndoCtl(i).ctl.BckClr
                                        frmDesign.Controls(j).BackStyle = UndoCtl(i).ctl.BckStl
                                        frmDesign.Controls(j).ForeColor = UndoCtl(i).ctl.ForClr
                                        frmDesign.Controls(j).FontName = UndoCtl(i).ctl.FntNam
                                        frmDesign.Controls(j).FontSize = UndoCtl(i).ctl.FntSiz
                                        frmDesign.Controls(j).FontBold = UndoCtl(i).ctl.FntBld
                                        frmDesign.Controls(j).FontItalic = UndoCtl(i).ctl.FntItl
                                        frmDesign.Controls(j).FontUnderline = UndoCtl(i).ctl.FntUnd
                                        frmDesign.Controls(j).Caption = UndoCtl(i).ctl.strText
                                        frmDesign.Controls(j).Alignment = UndoCtl(i).ctl.Align
                                    ElseIf TypeOf frmDesign.Controls(j) Is CheckBoxControl Then
                                        frmDesign.Controls(j).BackColor = UndoCtl(i).ctl.BckClr
                                        frmDesign.Controls(j).BorderColor = UndoCtl(i).ctl.BdrClr
                                        frmDesign.Controls(j).DisplayType = UndoCtl(i).ctl.DisplayType
                                        frmDesign.Controls(j).Sunken = UndoCtl(i).ctl.Sunken
                                    End If
                                End If
                            End If
                        End If
                    Next j
                End If
            End If
        Next i
    ElseIf UndoList(CurrUndoPos).UndoCategory = unSection Then
        If UndoList(CurrUndoPos).Type = unSectWidth Then
            For i = 0 To UBound(UndoSect)
                If UndoSect(i).UndoIDNo = CurrUndoPos Then
                    For j = 0 To 10
                        frmDesign.picSection(j).width = UndoSect(i).SectWidth
                    Next j
                End If
            Next i
        ElseIf UndoList(CurrUndoPos).Type = unSectHeight Then
            For i = 0 To UBound(UndoSect)
                If UndoSect(i).UndoIDNo = CurrUndoPos Then
                    frmDesign.picSection(UndoSect(i).SectNo).Height = UndoSect(i).SectHeight
                End If
            Next i
        ElseIf UndoList(CurrUndoPos).Type = unSectColor Then
            For i = 0 To UBound(UndoSect)
                If UndoSect(i).UndoIDNo = CurrUndoPos Then
                    frmDesign.picSection(UndoSect(i).SectNo).BackColor = UndoSect(i).SectColor
                End If
            Next i
        ElseIf UndoList(CurrUndoPos).Type = unSectWidth Then
            For i = 0 To 10
                frmDesign.picSection(i).width = UndoSect(i).SectWidth
            Next i
        End If
    End If
    
    blnFirstUndo = False
    
End Sub

Private Sub ReCreateControl(ctl As ControlInfo)
Dim i As Integer

    For i = 0 To frmDesign.Controls.count - 1
        If frmDesign.Controls(i).Name = ctl.Name Then
            If frmDesign.Controls(i).Index = ctl.Index Then
                Exit Sub
            End If
        End If
    Next i

    If ctl.Type = cLine Then
        With ctl
            Load frmDesign.Lin(.Index)
            Set frmDesign.Lin(.Index).Container = frmDesign.picSection(.SecNo)
            frmDesign.Lin(.Index).Tag = .SecNo
            frmDesign.Lin(.Index).X1 = .X1
            frmDesign.Lin(.Index).Y1 = .Y1
            frmDesign.Lin(.Index).X2 = .X2
            frmDesign.Lin(.Index).Y2 = .Y2
            frmDesign.Lin(.Index).BorderColor = .BdrClr
            frmDesign.Lin(.Index).BorderStyle = .BdrStl
            frmDesign.Lin(.Index).BorderWidth = .BdrWd
            frmDesign.Lin(.Index).Visible = True
        End With
    ElseIf ctl.Type = cBox Then
        With ctl
            Load frmDesign.Shape(.Index)
            Set frmDesign.Shape(.Index).Container = frmDesign.picSection(.SecNo)
            frmDesign.Shape(.Index).Tag = .SecNo
            frmDesign.Shape(.Index).Left = .Left
            frmDesign.Shape(.Index).Top = .Top
            frmDesign.Shape(.Index).width = .width
            frmDesign.Shape(.Index).Height = .Height
            frmDesign.Shape(.Index).BorderColor = .BdrClr
            frmDesign.Shape(.Index).BorderStyle = .BdrStl
            frmDesign.Shape(.Index).BorderWidth = .BdrWd
            frmDesign.Shape(.Index).BackColor = .BckClr
            frmDesign.Shape(.Index).BackStyle = .BckStl
            frmDesign.Shape(.Index).Shape = .DisplayType
            frmDesign.Shape(.Index).Visible = True
        End With
    ElseIf ctl.Type = cLabel Or ctl.Type = cDataField Or ctl.Type = cDatePageField Or _
        ctl.Type = cSumField Or ctl.Type = cCalcField Then
        Dim ctlLoaded As Control
        If ctl.Type = cLabel Then
            Load frmDesign.Label(ctl.Index)
            Set ctlLoaded = frmDesign.Label(ctl.Index)
        Else
            Load frmDesign.Field(ctl.Index)
            Set ctlLoaded = frmDesign.Field(ctl.Index)
        End If
        With ctl
            Set ctlLoaded.Container = frmDesign.picSection(.SecNo)
            ctlLoaded.Tag = .SecNo
            ctlLoaded.Left = .Left
            ctlLoaded.Top = .Top
            ctlLoaded.width = .width
            ctlLoaded.Height = .Height
            ctlLoaded.BorderStyle = .BdrStl
            ctlLoaded.BackColor = .BckClr
            ctlLoaded.BackStyle = .BckStl
            ctlLoaded.FontName = .FntNam
            ctlLoaded.FontSize = .FntSiz
            ctlLoaded.FontBold = .FntBld
            ctlLoaded.FontItalic = .FntItl
            ctlLoaded.FontUnderline = .FntUnd
            ctlLoaded.Alignment = .Align
            ctlLoaded.Caption = .strText
            ctlLoaded.LinkTimeout = .Type
            ctlLoaded.DataField = .Fieldname
            ctlLoaded.Visible = True
        End With
    ElseIf ctl.Type = cCheckBox Then
        With ctl
            Load frmDesign.Chkbox(.Index)
            Set frmDesign.Chkbox(.Index).Container = frmDesign.picSection(.SecNo)
            frmDesign.Chkbox(.Index).Tag = .SecNo
            frmDesign.Chkbox(.Index).Left = .Left
            frmDesign.Chkbox(.Index).Top = .Top
            frmDesign.Chkbox(.Index).width = .width
            frmDesign.Chkbox(.Index).Height = .Height
            frmDesign.Chkbox(.Index).BackColor = .BckClr
            frmDesign.Chkbox(.Index).DisplayType = .DisplayType
            frmDesign.Chkbox(.Index).DataField = .Fieldname
            frmDesign.Chkbox(.Index).Sunken = .Sunken
            frmDesign.Chkbox(.Index).BorderColor = .BdrClr
            frmDesign.Chkbox(.Index).Visible = True
        End With
    End If

End Sub
