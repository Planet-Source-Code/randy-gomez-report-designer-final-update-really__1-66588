VERSION 5.00
Begin VB.Form frmDatePageField 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Place Date/Page Field in Report"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6765
   Icon            =   "frmDatePageField.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton OptPlacePage 
      Caption         =   "Page Number Field"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   180
      Width           =   1935
   End
   Begin VB.OptionButton OptPlaceDate 
      Caption         =   "Date Field"
      Height          =   315
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   180
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.Frame fraPageFormat 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   4020
      TabIndex        =   3
      Top             =   1320
      Width           =   2550
      Begin VB.OptionButton optPageNumAndTotal 
         BackColor       =   &H8000000B&
         Caption         =   "Page # of (Total Pages)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   10
         Top             =   420
         Width           =   2520
      End
      Begin VB.OptionButton OptPageNum 
         BackColor       =   &H8000000B&
         Caption         =   "Page #"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.ListBox lstDateFormat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   240
      TabIndex        =   2
      Top             =   900
      Width           =   2955
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
      Left            =   5340
      TabIndex        =   6
      Top             =   3360
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
      Left            =   5340
      TabIndex        =   4
      Top             =   2880
      Width           =   1230
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   6780
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   3540
      X2              =   3540
      Y1              =   0
      Y2              =   2640
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "Select Format :"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   900
      Width           =   1155
   End
   Begin VB.Label lblPrompt 
      BackColor       =   &H8000000B&
      Caption         =   "Select Format :"
      Height          =   225
      Left            =   240
      TabIndex        =   8
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      Caption         =   "Field Contents :"
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   2940
      Width           =   1155
   End
   Begin VB.Label lblFieldString 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   300
      TabIndex        =   5
      Top             =   3180
      Width           =   4575
   End
End
Attribute VB_Name = "frmDatePageField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intFieldType As Integer

Private Sub cboDateFormat_Click()

    intDateFormatType = cboDateFormat.ListIndex

End Sub

Private Sub cmdCancel_Click()

    frmDesign.optAction(0).value = True
    Unload Me

End Sub

Private Sub cmdOK_Click()

    If intFieldType = 1 Then
        Select Case lstDateFormat.ListIndex
            Case 0: strSpecialFieldContent = "=[Date:mm/dd/yy]"
            Case 1: strSpecialFieldContent = "=[Date:mm/dd/yyyy]"
            Case 2: strSpecialFieldContent = "=[Date:mmm. dd, yyyy]"
            Case 3: strSpecialFieldContent = "=[Date:mmmm dd, yyyy]"
            Case 4: strSpecialFieldContent = "=[Date:wwww, mmmm dd, yyyy]"
        End Select
    ElseIf intFieldType = 2 Then
        If Me.OptPageNum = True Then
            strSpecialFieldContent = "=Page [PageNo]"
        ElseIf Me.optPageNumAndTotal = True Then
            strSpecialFieldContent = "=Page [PageNo] of [NumPages]"
        End If
    End If
    
    If blnEditExisting Then
        ctlActive.Caption = strSpecialFieldContent
        ctlActive.ToolTipText = strSpecialFieldContent
    Else
        lngState = PlaceNewControl
        lngPrevState = PlaceNewControl
    End If
    
    Unload Me

End Sub

Private Sub Form_Load()

    Me.lstDateFormat.AddItem Format(Now(), "mm/dd/yy")
    Me.lstDateFormat.AddItem Format(Now(), "mm/dd/yyyy")
    Me.lstDateFormat.AddItem Format(Now(), "mmm. dd, yyyy")
    Me.lstDateFormat.AddItem Format(Now(), "mmmm dd, yyyy")
    Me.lstDateFormat.AddItem FormatDateTime(Now(), vbLongDate)
    intFieldType = 1

End Sub

Private Sub lstDateFormat_Click()

    Select Case lstDateFormat.ListIndex
        Case 0: strSpecialFieldContent = "=[Date:mm/dd/yy]"
        Case 1: strSpecialFieldContent = "=[Date:mm/dd/yyyy]"
        Case 2: strSpecialFieldContent = "=[Date:mmm. dd, yyyy]"
        Case 3: strSpecialFieldContent = "=[Date:mmmm dd, yyyy]"
        Case 4: strSpecialFieldContent = "=[Date:wwww, mmmm dd, yyyy]"
    End Select

    Me.lblFieldString = strSpecialFieldContent

End Sub

Private Sub OptPageNum_Click()

    strSpecialFieldContent = "=Page [PageNo]"
    Me.lblFieldString = strSpecialFieldContent


End Sub

Private Sub optPageNumAndTotal_Click()

    strSpecialFieldContent = "=Page [PageNo] of [NumPages]"
    Me.lblFieldString = strSpecialFieldContent

End Sub

Private Sub OptPlaceDate_Click()

    
    Me.OptPageNum.Enabled = False
    Me.optPageNumAndTotal.Enabled = False
    Me.lstDateFormat.Enabled = True
    Me.lstDateFormat.ListIndex = 0
    intFieldType = 1
    lstDateFormat_Click

End Sub

Private Sub OptPlacePage_Click()

    Me.OptPageNum.Enabled = True
    Me.optPageNumAndTotal.Enabled = True
    Me.OptPageNum = True
    strSpecialFieldContent = "=Page [PageNo]"
    Me.lblFieldString = strSpecialFieldContent
    intFieldType = 2
    Me.lstDateFormat.Enabled = False

End Sub
