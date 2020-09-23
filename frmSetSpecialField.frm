VERSION 5.00
Begin VB.Form frmSetSpecialField 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Date Field"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboAggFuncList 
      Height          =   315
      ItemData        =   "frmSetSpecialField.frx":0000
      Left            =   510
      List            =   "frmSetSpecialField.frx":0013
      TabIndex        =   8
      Top             =   1125
      Visible         =   0   'False
      Width           =   2850
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
      Height          =   420
      Left            =   4185
      TabIndex        =   6
      Top             =   810
      Width           =   1230
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
      Height          =   420
      Left            =   4170
      TabIndex        =   5
      Top             =   360
      Width           =   1230
   End
   Begin VB.Frame fraPageFormat 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   330
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   3390
      Begin VB.OptionButton optPageNumAndTotal 
         Caption         =   "Page # of (Total Pages)"
         Height          =   285
         Left            =   1170
         TabIndex        =   4
         Top             =   0
         Width           =   2040
      End
      Begin VB.OptionButton OptPageNum 
         Caption         =   "Page #"
         Height          =   285
         Left            =   135
         TabIndex        =   3
         Top             =   0
         Width           =   1365
      End
   End
   Begin VB.ComboBox cboFieldList 
      Height          =   315
      Left            =   510
      TabIndex        =   1
      Top             =   450
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.ComboBox cboDateFormat 
      Height          =   315
      ItemData        =   "frmSetSpecialField.frx":0040
      Left            =   510
      List            =   "frmSetSpecialField.frx":0042
      TabIndex        =   0
      Top             =   450
      Width           =   2850
   End
   Begin VB.Label Label1 
      Caption         =   "Calculation :"
      Height          =   270
      Left            =   510
      TabIndex        =   9
      Top             =   870
      Width           =   960
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Date Format :"
      Height          =   285
      Left            =   510
      TabIndex        =   7
      Top             =   180
      Width           =   3555
   End
End
Attribute VB_Name = "frmSetSpecialField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AggFunc As String

Private Sub cboAggFuncList_Click()
    
    Select Case cboAggFuncList
        Case "Total": AggFunc = "={SumOf["
        Case "Average": AggFunc = "={AvgOf["
        Case "Minimum": AggFunc = "={MinOf["
        Case "Maximum": AggFunc = "={MaxOf["
        Case "Count": AggFunc = "={CntOf["
    End Select
        
End Sub

Private Sub cboDateFormat_Click()

    intDateFormatType = cboDateFormat.ListIndex

End Sub

Private Sub cmdCancel_Click()

    Me.Hide

End Sub

Private Sub cmdOK_Click()

    lngState = PlaceNewControl
    lngPrevState = PlaceNewControl
    If intControlType = cDateField Then
        Select Case cboDateFormat.ListIndex
            Case 0: strSpecialFieldContent = "=[Date:mm/dd/yy]"
            Case 1: strSpecialFieldContent = "=[Date:mm/dd/yyyy]"
            Case 2: strSpecialFieldContent = "=[Date:mmm. dd, yyyy]"
            Case 3: strSpecialFieldContent = "=[Date:mmmm dd, yyyy]"
            Case 4: strSpecialFieldContent = "=[Date:wwww, mmmm dd, yyyy]"
        End Select
    ElseIf intControlType = cPageField Then
        If Me.OptPageNum = True Then
            strSpecialFieldContent = "Page [PageNo]"
        ElseIf Me.optPageNumAndTotal = True Then
            strSpecialFieldContent = "Page [PageNo] of [NumPages]"
        End If
    ElseIf intControlType = cAggField Then
        strSpecialFieldContent = AggFunc & Me.cboFieldList & "]}"
    End If
    
    Me.Hide

End Sub

Private Sub Form_Load()
Dim i As Integer

    If intControlType = cDateField Then
        Me.Caption = "Set Date Field Format"
        lblPrompt = "Date Format :"
        Me.cboDateFormat.AddItem Format(Now(), "mm/dd/yy")
        Me.cboDateFormat.AddItem Format(Now(), "mm/dd/yyyy")
        Me.cboDateFormat.AddItem Format(Now(), "mmm. dd, yyyy")
        Me.cboDateFormat.AddItem Format(Now(), "mmmm dd, yyyy")
        Me.cboDateFormat.AddItem FormatDateTime(Now(), vbLongDate)
        Me.cboDateFormat.Visible = True
        Me.cboFieldList.Visible = False
        Me.cboAggFuncList.Visible = False
        Me.Label1.Visible = False
        Me.fraPageFormat.Visible = False
    ElseIf intControlType = cPageField Then
        Me.Caption = "Set Page Number Field Format"
        lblPrompt = "Page Format :"
        Me.fraPageFormat.Visible = True
        Me.cboDateFormat.Visible = False
        Me.cboFieldList.Visible = False
        Me.cboAggFuncList.Visible = False
        Me.Label1.Visible = False
    ElseIf intControlType = cAggField Then
        Me.Caption = "Select Calculated Field"
        lblPrompt = "Select Field (Place in Report Footer only)"
        For i = 0 To rstData.Fields.Count - 1
            If IsNumeric(rstData.Fields(i)) Then
                Me.cboFieldList.AddItem DataField(i)
            End If
        Next i
        Me.fraPageFormat.Visible = False
        Me.cboDateFormat.Visible = False
        Me.cboFieldList.Visible = True
        Me.cboAggFuncList.Visible = True
        Me.Label1.Visible = True
    End If

End Sub
