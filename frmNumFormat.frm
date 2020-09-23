VERSION 5.00
Begin VB.Form frmNumFormat 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Set Field Format"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3690
   Icon            =   "frmNumFormat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   1890
      TabIndex        =   12
      Top             =   3540
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
      Height          =   345
      Left            =   450
      TabIndex        =   11
      Top             =   3540
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Caption         =   "Number Format "
      Height          =   3255
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   3555
      Begin VB.OptionButton OptFormat 
         BackColor       =   &H8000000B&
         Caption         =   "None"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   3000
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton OptFormat 
         BackColor       =   &H8000000B&
         Caption         =   "Percent  (%125.00)"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cboTimeFormat 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   10
         Top             =   2460
         Width           =   2295
      End
      Begin VB.OptionButton OptFormat 
         BackColor       =   &H8000000B&
         Caption         =   "Time"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox cboDateFormat 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   9
         Top             =   1980
         Width           =   2295
      End
      Begin VB.OptionButton OptFormat 
         BackColor       =   &H8000000B&
         Caption         =   "Date"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   675
      End
      Begin VB.OptionButton OptFormat 
         BackColor       =   &H8000000B&
         Caption         =   "Whole Number   (6,125)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   2055
      End
      Begin VB.OptionButton OptFormat 
         BackColor       =   &H8000000B&
         Caption         =   "Currency   (-$6,125.00)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1995
      End
      Begin VB.OptionButton OptFormat 
         BackColor       =   &H8000000B&
         Caption         =   "Decimal  (6,125.00)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1875
      End
      Begin VB.TextBox txtDecPlaces 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2820
         TabIndex        =   8
         Text            =   "2"
         Top             =   840
         Width           =   375
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   2460
         X2              =   2700
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   2460
         X2              =   2460
         Y1              =   720
         Y2              =   1200
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   2220
         X2              =   2460
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   2220
         X2              =   2460
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "Decimal Places"
         Height          =   435
         Left            =   2700
         TabIndex        =   13
         Top             =   360
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmNumFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFormat As String
Dim strDateFormat As String
Dim strTimeFormat As String

Private Sub cboDateFormat_Click()

    Select Case cboDateFormat.ListIndex
        Case 0: strDateFormat = "mm/dd/yy"
        Case 1: strDateFormat = "mm/dd/yyyy"
        Case 2: strDateFormat = "mmm. dd, yyyy"
        Case 3: strDateFormat = "mmmm dd, yyyy"
        Case 4: strDateFormat = "wwww mmmm dd, yyyy"
    End Select
    
    strFormat = "|(Date:" & strDateFormat & ")"

End Sub

Private Sub cboTimeFormat_Click()

    Select Case cboTimeFormat.ListIndex
        Case 0: strTimeFormat = "hh:mm"
        Case 1: strTimeFormat = "h:mm:ss AM"
    End Select

    strFormat = "|(Time:" & strTimeFormat & ")"

End Sub

Private Sub cmdCancel_Click()

    blnEditExisting = False
    Unload Me

End Sub

Private Sub cmdOK_Click()

    If InStr(1, ctlActive.ToolTipText, "|") > 0 Then
        ctlActive.Caption = Left(ctlActive.ToolTipText, InStr(1, ctlActive.ToolTipText, "|") - 1) & strFormat
        ctlActive.ToolTipText = ctlActive.Caption
    Else
        ctlActive.Caption = ctlActive.ToolTipText & strFormat
        ctlActive.ToolTipText = ctlActive.Caption
    End If
    blnEditExisting = False
    
    Unload Me

End Sub

Private Sub Form_Load()

    Me.cboDateFormat.AddItem Format(Now(), "mm/dd/yy")
    Me.cboDateFormat.AddItem Format(Now(), "mm/dd/yyyy")
    Me.cboDateFormat.AddItem Format(Now(), "mmm. dd, yyyy")
    Me.cboDateFormat.AddItem Format(Now(), "mmmm dd, yyyy")
    Me.cboDateFormat.AddItem FormatDateTime(Now(), vbLongDate)

    Me.cboTimeFormat.AddItem FormatDateTime(Now(), vbShortTime)
    Me.cboTimeFormat.AddItem FormatDateTime(Now(), vbLongTime)
    
    If blnEditExisting Then
        Dim strGetFormat As String
        strGetFormat = Right(ctlActive.Caption, Len(ctlActive.Caption) - InStr(1, ctlActive.Caption, "|"))
        If InStr(1, strGetFormat, "Whole") > 0 Then
            OptFormat(0).value = True
        ElseIf InStr(1, strGetFormat, "Decimal") > 0 Then
            OptFormat(1).value = True
            txtDecPlaces = Mid(strGetFormat, 10, 2)
        ElseIf InStr(1, strGetFormat, "Currency") > 0 Then
            OptFormat(2).value = True
        ElseIf InStr(1, strGetFormat, "Date") > 0 Then
            OptFormat(3).value = True
            If Mid(strGetFormat, 7, Len(strGetFormat) - 7) = "wwww mmmm dd, yyyy" Then
                Me.cboDateFormat = FormatDateTime(Now(), vbLongDate)
            Else
                cboDateFormat = Format(Now(), Mid(strGetFormat, 7, Len(strGetFormat) - 7))
            End If
        ElseIf InStr(1, strGetFormat, "Time") > 0 Then
            OptFormat(4).value = True
            cboTimeFormat = Format(Now(), Mid(strGetFormat, 7, Len(strGetFormat) - 7))
        ElseIf InStr(1, strGetFormat, "Percent") > 0 Then
            OptFormat(5).value = True
            txtDecPlaces = Mid(strGetFormat, 10, 2)
        End If
    End If
    
End Sub

Private Sub OptFormat_Click(Index As Integer)

    Select Case Index
        Case 0: strFormat = "|(Whole)"
        Case 1: strFormat = "|(Decimal:" & Format(txtDecPlaces, "00") & ")"
        Case 2: strFormat = "|(Currency)"
        Case 3: strFormat = "|(Date:" & strDateFormat & ")"
        Case 4: strFormat = "|(Time:" & strTimeFormat & ")"
        Case 5: strFormat = "|(Percent:" & Format(txtDecPlaces, "00") & ")"
        Case 6: strFormat = ""
    End Select

    txtDecPlaces.Enabled = (Index = 1) Or (Index = 5)
    cboDateFormat.Enabled = (Index = 3)
    cboTimeFormat.Enabled = (Index = 4)

End Sub

Private Sub txtDecPlaces_Change()

    If OptFormat(1).value = True Then
        strFormat = "|(Decimal:" & Format(txtDecPlaces, "00") & ")"
    ElseIf OptFormat(5).value = True Then
        strFormat = "|(Percent:" & Format(txtDecPlaces, "00") & ")"
    End If

End Sub
