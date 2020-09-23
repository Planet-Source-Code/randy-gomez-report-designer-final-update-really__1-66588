VERSION 5.00
Begin VB.Form frmGridSpace 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Design Grid Settings"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4035
   Icon            =   "frmGridSpace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSnap 
      BackColor       =   &H8000000B&
      Caption         =   "Snap Enabled"
      Height          =   195
      Left            =   2295
      TabIndex        =   3
      Top             =   1095
      Width           =   1455
   End
   Begin VB.CheckBox chkShowGrid 
      BackColor       =   &H8000000B&
      Caption         =   "Visible"
      Height          =   255
      Left            =   2295
      TabIndex        =   4
      Top             =   1410
      Width           =   795
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
      Height          =   315
      Left            =   2070
      TabIndex        =   6
      Top             =   2055
      Width           =   1005
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
      Height          =   315
      Left            =   990
      TabIndex        =   5
      Top             =   2055
      Width           =   1005
   End
   Begin VB.TextBox txtGrSpc 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2295
      TabIndex        =   2
      Top             =   495
      Width           =   945
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Select Grid Spacing"
      ForeColor       =   &H00000000&
      Height          =   1665
      Left            =   195
      TabIndex        =   1
      Top             =   180
      Width           =   1695
      Begin VB.OptionButton OptGS 
         BackColor       =   &H8000000B&
         Caption         =   "1/4"""
         Height          =   285
         Index           =   3
         Left            =   465
         TabIndex        =   11
         Tag             =   "24"
         Top             =   1260
         Width           =   1005
      End
      Begin VB.OptionButton OptGS 
         BackColor       =   &H8000000B&
         Caption         =   "1/8"""
         Height          =   285
         Index           =   2
         Left            =   465
         TabIndex        =   8
         Tag             =   "12"
         Top             =   960
         Width           =   1005
      End
      Begin VB.OptionButton OptGS 
         BackColor       =   &H8000000B&
         Caption         =   "1/16"""
         Height          =   285
         Index           =   1
         Left            =   465
         TabIndex        =   7
         Tag             =   "6"
         Top             =   615
         Width           =   1065
      End
      Begin VB.OptionButton OptGS 
         BackColor       =   &H8000000B&
         Caption         =   "1/32"""
         Height          =   285
         Index           =   0
         Left            =   465
         TabIndex        =   0
         Tag             =   "3"
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.Label lblUnits 
      BackColor       =   &H8000000B&
      Caption         =   "inches"
      Height          =   225
      Left            =   3315
      TabIndex        =   10
      Top             =   555
      Width           =   525
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Or Enter in decimal format:"
      Height          =   255
      Left            =   1995
      TabIndex        =   9
      Top             =   195
      Width           =   1935
   End
End
Attribute VB_Name = "frmGridSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnSelectedCtl As Boolean

Private Sub chkShowGrid_Click()

    blnGridOn = -(chkShowGrid.value)

End Sub

Private Sub chkSnap_Click()

    blnSnapOn = -(chkSnap.value)

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()
Dim a As Integer

    If blnCustomGrid Then
        If txtGrSpc.text > "" Then
            If Val(txtGrSpc) * 96 <= 2 Then
                MsgBox "The value you have entered is invalid", vbOKOnly, "Grid Space Value Invalid"
                txtGrSpc.SetFocus
                Exit Sub
            Else
                GridSpace = Val(txtGrSpc) * 96
            End If
        End If
    Else
        For a = 0 To 3
            If OptGS(a).value Then
                GridSpace = Val(OptGS(a).Tag)
            End If
        Next a
    End If

    Unload Me
    
    ShowGrid
    
    If blnGroupSelected Then
        frmDesign.HiliteObjects
    End If

End Sub

Private Sub Form_Load()
Dim i As Integer

    If blnCustomGrid Then
        OptGS(0).value = False
        OptGS(1).value = False
        OptGS(2).value = False
        OptGS(3).value = False
    Else
        If PageScaleUnits = scEnglish Then
            SetupEngUnitOptions
        ElseIf PageScaleUnits = scMetric Then
            SetupMetricUnitOptions
        End If
    End If
    
    For i = 0 To 3
        If OptGS(i).Tag = GridSpace Then
            OptGS(i).value = True
            Exit For
        End If
    Next i
    
    txtGrSpc.text = Format(Round(GridSpace / 96, 4), "0.000")
    If blnGridOn Then chkShowGrid.value = 1
    If blnSnapOn Then chkSnap.value = 1

End Sub

Private Sub SetupEngUnitOptions()

    OptGS(0).Tag = 3
    OptGS(1).Tag = 6
    OptGS(2).Tag = 12
    OptGS(3).Tag = 24
    OptGS(0).Caption = "1/32"""
    OptGS(1).Caption = "1/16"""
    OptGS(2).Caption = "1/8"""
    OptGS(3).Caption = "1/4"""
    lblUnits.Caption = "inches"

End Sub

Private Sub SetupMetricUnitOptions()

    OptGS(0).Tag = 3.78
    OptGS(1).Tag = 7.56
    OptGS(2).Tag = 11.34
    OptGS(3).Tag = 18.9
    OptGS(0).Caption = "0.1 cm"
    OptGS(1).Caption = "0.2 cm"
    OptGS(2).Caption = "0.3 cm"
    OptGS(3).Caption = "0.5 cm"
    lblUnits.Caption = "cm"

End Sub

Private Sub OptGS_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    txtGrSpc.text = Format(Round(Val(OptGS(Index).Tag) / 96, 4), "0.000")
    blnCustomGrid = False


End Sub

Private Sub txtGrSpc_KeyUp(KeyCode As Integer, Shift As Integer)

    blnCustomGrid = True
    
End Sub
