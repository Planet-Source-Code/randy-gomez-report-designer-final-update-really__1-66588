VERSION 5.00
Begin VB.Form frmAggField 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Place Summary Field"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6465
   Icon            =   "frmAggField.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboPlaceIn 
      Height          =   315
      Left            =   1260
      TabIndex        =   13
      Top             =   3540
      Width           =   2715
   End
   Begin VB.ListBox lstCalcFields 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   4125
      TabIndex        =   2
      Top             =   615
      Width           =   2175
   End
   Begin VB.TextBox txtRepFieldName 
      Height          =   375
      Left            =   3810
      TabIndex        =   3
      Top             =   2775
      Width           =   2490
   End
   Begin VB.ListBox lstSelectCalc 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   225
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.ListBox lstFields 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   1815
      TabIndex        =   1
      Top             =   615
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4170
      TabIndex        =   4
      Top             =   3510
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5280
      TabIndex        =   6
      Top             =   3510
      Width           =   1020
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      Caption         =   "Place Field in :"
      Height          =   255
      Left            =   1260
      TabIndex        =   14
      Top             =   3300
      Width           =   1155
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      Caption         =   "Select Calculated Field :"
      Height          =   195
      Index           =   1
      Left            =   4125
      TabIndex        =   12
      Top             =   375
      Width           =   1965
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000B&
      Caption         =   "Report Field Name :"
      Height          =   255
      Left            =   3825
      TabIndex        =   11
      Top             =   2535
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Field Contents :"
      Height          =   255
      Left            =   210
      TabIndex        =   10
      Top             =   2550
      Width           =   1215
   End
   Begin VB.Label lblContents 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   9
      Top             =   2790
      Width           =   3420
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Double-click lists to add functions or fields."
      Height          =   210
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   3420
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      Caption         =   "Select Data Field :"
      Height          =   195
      Index           =   0
      Left            =   1830
      TabIndex        =   7
      Top             =   375
      Width           =   1605
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "Select Function :"
      Height          =   255
      Left            =   255
      TabIndex        =   5
      Top             =   375
      Width           =   1635
   End
End
Attribute VB_Name = "frmAggField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strAggFunc As String
Dim strFieldName As String
Dim strFieldContents As String
Dim strAggFieldName As String
Dim strRepFieldName1 As String
Dim strRepFieldName2 As String
Dim strSaveContents As String

Private Sub UpdateContents()
'called at form loading and whenever selections have been made

    strFieldContents = "={" & strAggFunc & "([" & strFieldName & "])}"
    lblContents = strFieldContents
    strAggFieldName = strRepFieldName1 & strRepFieldName2
    txtRepFieldName.text = strAggFieldName
    
End Sub

Private Sub cmdCancel_Click()

    blnEditExisting = False
    frmDesign.optAction(0).value = True
    Unload Me

End Sub

Private Sub cmdOK_Click()

'if the aggregate function has been changed, update the active control
'replace '~' character with new function and field - preserves any formatting info
    If blnEditExisting Then
        strFieldContents = Replace(strSaveContents, "~", strFieldContents)
        ctlActive.Caption = strFieldContents
        ctlActive.DataField = strFieldName
        ctlActive.ToolTipText = strFieldContents
        blnEditExisting = False
    Else                    'otherwise create a new summary field control
        If cboPlaceIn > "" Then
            frmDesign.CreateField cboPlaceIn.ItemData(cboPlaceIn.ListIndex), "Agg", strFieldContents, strFieldName, True
        Else
            MsgBox "You must select a section to place the summary field in.", vbInformation
            cboPlaceIn.SetFocus
            Exit Sub
        End If
    End If
    
    Unload Me

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim strGetFunc As String

'get data field from database recordset
    If strDataFileName > "" Then
        If strTableName > "" Then
            For i = 0 To rstData.Fields.count - 1
                    lstFields.AddItem rstData.Fields(i).Name
            Next i
        End If
    End If

'get any calculated fields from report
    For i = 0 To frmDesign.Controls.count - 1
        If frmDesign.Controls(i).Tag = 5 Then
            Set ctlTest = frmDesign.Controls(i)
            If TypeOf ctlTest Is Label Then
                If ctlTest.LinkTimeout = cCalcField Then
                    lstCalcFields.AddItem ctlTest.DataField
                End If
            End If
        End If
    Next i
    
'fill the summary function listbox
    With Me.lstSelectCalc
        .Enabled = True
        .Clear
        .AddItem "Total"
        .AddItem "Average"
        .AddItem "Minimum"
        .AddItem "Maximum"
        .AddItem "Count"
    End With

'if editting an existing summary field, load the current contents
    If blnEditExisting Then
        strSaveContents = ctlActive.Caption
        'parse out the equation portion of the contents, temporarily replace with '~'
        If InStr(1, strSaveContents, "|") > 0 Then
            lblContents = Left(strSaveContents, InStr(1, strSaveContents, "}"))
            strGetFunc = Left(strSaveContents, InStr(1, strSaveContents, "}"))
            strSaveContents = Replace(strSaveContents, strGetFunc, "~")
        Else
            lblContents = strSaveContents
        End If
        'save the various portions of the contents to the appropriate variables
        strAggFunc = Mid(ctlActive.Caption, 3, 5)   'get function name
        strRepFieldName1 = strAggFunc               'save to variable
        strFieldName = ctlActive.DataField          'get field name
        strRepFieldName2 = strFieldName             'save to variable
        
'set function listbox selection from variable
        Select Case strRepFieldName1
            Case "SumOf"
                lstSelectCalc = "Total"
            Case "AvgOf"
                lstSelectCalc = "Average"
            Case "MinOf"
                lstSelectCalc = "Minimum"
            Case "MaxOf"
                lstSelectCalc = "Maximum"
            Case "CntOf"
                lstSelectCalc = "Count"
        End Select
        
'set fieldname listbox selection from variable
        For i = 0 To lstFields.ListCount - 1
            If lstFields.List(i) = strRepFieldName2 Then
                lstFields = lstFields.List(i)
                Exit For
            End If
        Next i
        
'set calculated field listbox selection from variable
        For i = 0 To lstCalcFields.ListCount - 1
            If lstCalcFields.List(i) = strRepFieldName2 Then
                lstCalcFields = lstCalcFields.List(i)
                Exit For
            End If
        Next i
        
        UpdateContents
        
        lstCalcFields.Enabled = True
        lstFields.Enabled = True
    End If
    
    For i = 6 To 8
        If frmDesign.picSection(i).Visible Then
            cboPlaceIn.AddItem frmDesign.cmdDivider(i).Caption
            cboPlaceIn.ItemData(cboPlaceIn.NewIndex) = i
        End If
    Next i
    cboPlaceIn.AddItem frmDesign.cmdDivider(10).Caption
    cboPlaceIn.ItemData(cboPlaceIn.NewIndex) = 10

End Sub

Private Sub lstCalcFields_Click()

    strFieldName = lstCalcFields
    strRepFieldName2 = lstCalcFields
    UpdateContents

End Sub

Private Sub lstCalcFields_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    lstFields.ListIndex = -1

End Sub

Private Sub lstFields_Click()

    strFieldName = lstFields
    strRepFieldName2 = lstFields
    UpdateContents

End Sub

Private Sub lstFields_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    lstCalcFields.ListIndex = -1

End Sub

Private Sub lstSelectCalc_Click()

    Select Case lstSelectCalc
        Case "Total"
            strAggFunc = "SumOf"
        Case "Average"
            strAggFunc = "AvgOf"
        Case "Minimum"
            strAggFunc = "MinOf"
        Case "Maximum"
            strAggFunc = "MaxOf"
        Case "Count"
            strAggFunc = "CntOf"
    End Select
    
    strRepFieldName1 = strAggFunc
    UpdateContents
    lstFields.Enabled = True
    lstCalcFields.Enabled = True

End Sub
