VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDataConn 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect to Data Source"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   Icon            =   "frmDataConn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNewPage 
      Caption         =   "Check1"
      Height          =   195
      Index           =   2
      Left            =   5940
      TabIndex        =   50
      Top             =   4500
      Width           =   195
   End
   Begin VB.CheckBox chkNewPage 
      Caption         =   "Check1"
      Height          =   195
      Index           =   1
      Left            =   5940
      TabIndex        =   49
      Top             =   4000
      Width           =   195
   End
   Begin VB.CheckBox chkNewPage 
      Caption         =   "Check1"
      Height          =   195
      Index           =   0
      Left            =   5940
      TabIndex        =   47
      Top             =   3480
      Width           =   195
   End
   Begin VB.ComboBox cboFolderPath 
      Height          =   315
      Left            =   210
      TabIndex        =   45
      Top             =   6720
      Width           =   6075
   End
   Begin VB.CheckBox chkGrpFVis 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   2
      Left            =   5220
      TabIndex        =   43
      Top             =   4500
      Width           =   195
   End
   Begin VB.CheckBox chkGrpFVis 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   5220
      TabIndex        =   42
      Top             =   4000
      Width           =   195
   End
   Begin VB.CheckBox chkGrpFVis 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   5220
      TabIndex        =   41
      Top             =   3480
      Width           =   195
   End
   Begin VB.CheckBox chkGrpHVis 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   2
      Left            =   4620
      TabIndex        =   39
      Top             =   4500
      Width           =   195
   End
   Begin VB.CheckBox chkGrpHVis 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   4620
      TabIndex        =   38
      Top             =   4000
      Width           =   195
   End
   Begin VB.CheckBox chkGrpHVis 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   4620
      TabIndex        =   37
      Top             =   3480
      Width           =   195
   End
   Begin VB.CommandButton cmdViewData 
      Caption         =   "View Data..."
      Enabled         =   0   'False
      Height          =   360
      Left            =   1740
      TabIndex        =   35
      Top             =   2070
      Width           =   1200
   End
   Begin VB.Frame FraOrder 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   10
      Top             =   4515
      Width           =   795
      Begin VB.OptionButton OptAsc 
         BackColor       =   &H8000000B&
         Height          =   195
         Index           =   2
         Left            =   45
         TabIndex        =   32
         Top             =   15
         Value           =   -1  'True
         Width           =   195
      End
      Begin VB.OptionButton OptDesc 
         BackColor       =   &H8000000B&
         Height          =   210
         Index           =   2
         Left            =   480
         TabIndex        =   31
         Top             =   15
         Width           =   210
      End
   End
   Begin VB.Frame FraOrder 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   3480
      TabIndex        =   9
      Top             =   4005
      Width           =   795
      Begin VB.OptionButton OptAsc 
         BackColor       =   &H8000000B&
         Height          =   195
         Index           =   1
         Left            =   45
         TabIndex        =   30
         Top             =   15
         Value           =   -1  'True
         Width           =   195
      End
      Begin VB.OptionButton OptDesc 
         BackColor       =   &H8000000B&
         Height          =   210
         Index           =   1
         Left            =   480
         TabIndex        =   29
         Top             =   15
         Width           =   210
      End
   End
   Begin VB.Frame FraOrder 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   3480
      TabIndex        =   8
      Top             =   3495
      Width           =   795
      Begin VB.OptionButton OptDesc 
         BackColor       =   &H8000000B&
         Height          =   210
         Index           =   0
         Left            =   480
         TabIndex        =   28
         Top             =   0
         Width           =   210
      End
      Begin VB.OptionButton OptAsc 
         BackColor       =   &H8000000B&
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   27
         Top             =   0
         Value           =   -1  'True
         Width           =   195
      End
   End
   Begin VB.ComboBox cboFields 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   2850
      TabIndex        =   13
      Top             =   6135
      Width           =   2445
   End
   Begin VB.ComboBox cboTables 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   210
      TabIndex        =   12
      Top             =   6135
      Width           =   2445
   End
   Begin VB.CheckBox chkHasPics 
      BackColor       =   &H8000000B&
      Caption         =   "Images associated with records."
      Enabled         =   0   'False
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
      Left            =   240
      TabIndex        =   11
      Top             =   5220
      Width           =   3165
   End
   Begin VB.TextBox txtPicFolder 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   7920
      Width           =   6105
   End
   Begin VB.CommandButton cmdFindPicFolder 
      Caption         =   "Select Folder..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   14
      Top             =   7140
      Width           =   1575
   End
   Begin VB.ComboBox cboFields 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   810
      TabIndex        =   7
      Top             =   4455
      Width           =   2265
   End
   Begin VB.ComboBox cboFields 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   810
      TabIndex        =   6
      Top             =   3945
      Width           =   2265
   End
   Begin VB.ComboBox cboFields 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   810
      TabIndex        =   5
      Top             =   3435
      Width           =   2265
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   960
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5820
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "mdb"
      DialogTitle     =   "Open Data File"
      Filter          =   "MS Access Files (*.mdb)|*.mdb"
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
      Height          =   375
      Left            =   5220
      TabIndex        =   16
      Top             =   8460
      Width           =   1095
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
      Height          =   375
      Left            =   4020
      TabIndex        =   15
      Top             =   8460
      Width           =   1095
   End
   Begin VB.ComboBox cboTables 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   1740
      TabIndex        =   4
      Top             =   1695
      Width           =   2865
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   315
      Left            =   5340
      TabIndex        =   0
      Top             =   240
      Width           =   1035
   End
   Begin VB.TextBox txtDataFile 
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   570
      Width           =   6195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start New Page"
      Height          =   615
      Left            =   5760
      TabIndex        =   48
      Top             =   2760
      Width           =   555
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   1620
      X2              =   780
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   780
      X2              =   780
      Y1              =   1440
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   180
      X2              =   6350
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   180
      X2              =   6350
      Y1              =   2655
      Y2              =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   180
      X2              =   6350
      Y1              =   5115
      Y2              =   5115
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   180
      X2              =   6350
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Label lblSelPath 
      BackColor       =   &H8000000B&
      Caption         =   "Select folder path"
      Height          =   255
      Left            =   180
      TabIndex        =   46
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label lblFooter 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Footer"
      Enabled         =   0   'False
      Height          =   195
      Left            =   5040
      TabIndex        =   44
      Top             =   3135
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Header"
      Enabled         =   0   'False
      Height          =   195
      Left            =   4400
      TabIndex        =   40
      Top             =   3135
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblOR 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   36
      Top             =   7200
      Width           =   375
   End
   Begin VB.Label lblImagePathHead 
      BackColor       =   &H8000000B&
      Caption         =   "Image Folder Path stored in Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   34
      Top             =   5700
      Width           =   3315
   End
   Begin VB.Label lblSortHead 
      BackColor       =   &H8000000B&
      Caption         =   "Select sorting and grouping (optional)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   33
      Top             =   2820
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   3960
      Picture         =   "frmDataConn.frx":058A
      Top             =   3135
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   3510
      Picture         =   "frmDataConn.frx":0B14
      Top             =   3135
      Width           =   240
   End
   Begin VB.Label lblImgField 
      BackColor       =   &H8000000B&
      Caption         =   "Field:"
      Height          =   240
      Left            =   2850
      TabIndex        =   26
      Top             =   5940
      Width           =   1065
   End
   Begin VB.Label lblImgTable 
      BackColor       =   &H8000000B&
      Caption         =   "Table:"
      Height          =   285
      Left            =   210
      TabIndex        =   25
      Top             =   5940
      Width           =   1170
   End
   Begin VB.Label lblImgPath 
      BackColor       =   &H8000000B&
      Caption         =   "Image Folder path:"
      Height          =   285
      Left            =   180
      TabIndex        =   24
      Top             =   7680
      Width           =   1410
   End
   Begin VB.Label lbl3rd 
      BackStyle       =   0  'Transparent
      Caption         =   "3rd"
      Enabled         =   0   'False
      Height          =   225
      Left            =   450
      TabIndex        =   22
      Top             =   4485
      Width           =   315
   End
   Begin VB.Label lbl2nd 
      BackStyle       =   0  'Transparent
      Caption         =   "2nd"
      Enabled         =   0   'False
      Height          =   285
      Left            =   420
      TabIndex        =   21
      Top             =   3975
      Width           =   375
   End
   Begin VB.Label lbl1st 
      BackStyle       =   0  'Transparent
      Caption         =   "1st"
      Enabled         =   0   'False
      Height          =   225
      Left            =   450
      TabIndex        =   20
      Top             =   3495
      Width           =   315
   End
   Begin VB.Label lblSortField 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Select sort field(s)"
      Enabled         =   0   'False
      Height          =   240
      Left            =   840
      TabIndex        =   19
      Top             =   3195
      Width           =   2055
   End
   Begin VB.Label lblConnStatus 
      BackColor       =   &H8000000B&
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   945
      Width           =   4845
   End
   Begin VB.Label lblTable 
      BackColor       =   &H8000000B&
      Caption         =   "Select an existing Table / Query :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   1740
      TabIndex        =   17
      Top             =   1440
      Width           =   2925
   End
   Begin VB.Label lblDbFile 
      BackColor       =   &H8000000B&
      Caption         =   "Database file to connect to (currently MS Access only)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   300
      Width           =   4725
   End
End
Attribute VB_Name = "frmDataConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TableIndex As Integer
Dim TablesPerSection(2) As Integer
Dim strSort As String
Dim rstImage As ADODB.Recordset
Dim strOldDataFileName As String
Dim strOldTableName As String
Dim blnConnected As Boolean
Dim blnViewingData As Boolean

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

'// BrowseForFolder constants
Private Const BIF_RETURNONLYFSDIRS = &H1

'// BrowseForFolders APIs
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST


'// Frees memory allocated by the OLE (shell's) task allocator.
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Function BrowseFolder(hwnd As Long, szDialogTitle As String) As String
'displays the default Windows folder browser dialog
    Dim bi As BROWSEINFO
    Dim dwIList As Long
    Dim szPath As String * 260
    
    bi.hOwner = hwnd
    bi.lpszTitle = szDialogTitle
    bi.ulFlags = BIF_RETURNONLYFSDIRS
    
    dwIList = SHBrowseForFolder(bi)
    If dwIList Then
        If SHGetPathFromIDList(dwIList, szPath) Then
            BrowseFolder = Left$(szPath, InStr(szPath, vbNullChar) - 1)
        End If
        ' Free the memory the shell allocated for the pidl.
        Call CoTaskMemFree(dwIList)
    End If

End Function

Private Sub cboFields_Click(Index As Integer)
'fired when user clicks on the combo box to select a field name from the list
Dim i As Integer, j As Integer

    If Index = 3 Then                       'this is the image folder field name combo box
        For i = 0 To rstImage.Fields.count - 1
            If rstImage.Fields(i).Name = cboFields(3) Then
                If InStr(1, rstImage.Fields(i).value, "\") = 0 Then
                    Me.txtPicFolder = "Field does not contain a valid path!"
                Else
                    Me.cboFolderPath.Clear
                    If rstImage.RecordCount > 0 Then
                        For j = 0 To rstImage.RecordCount - 1
                            If Not IsNull(rstImage.Fields(i).value) Then
                                Me.cboFolderPath.AddItem rstImage.Fields(i).value
                            End If
                            rstImage.MoveNext
                        Next j
                    End If
                End If
                Exit For
            End If
        Next i
    Else
        If cboFields(Index) > "" Then
            Me.FraOrder(Index).Enabled = True
            Me.chkGrpHVis(Index).Enabled = True
            Me.chkGrpFVis(Index).Enabled = True
        End If
    End If

End Sub

Private Sub cboFolderPath_Click()

    If cboFolderPath > "" Then
        txtPicFolder = cboFolderPath
        strImageFolder = cboFolderPath
    End If

End Sub

Private Sub cboTables_Click(Index As Integer)
'fired when user clicks on the database tables combo box to select one from the list
Dim strImageTableName As String

    If cboTables(Index).ListIndex = -1 Then Exit Sub
    
    If Index = 0 Then               'this is the connection table combo box
        strTableName = Right(cboTables(Index), Len(cboTables(Index)) - 9)
        OpenData strTableName
        LoadSortFieldCombos
        chkHasPics.Enabled = True
        cmdViewData.Enabled = True
        frmDesign.Toolbar1.Buttons(16).Enabled = True
        lblSortField.Enabled = True
        lblHeader.Enabled = True
        lblFooter.Enabled = True
        lbl1st.Enabled = True
        lbl2nd.Enabled = True
        lbl3rd.Enabled = True
    ElseIf Index = 1 Then                   'this is the image path table combo box
        strImageTableName = cboTables(Index)
        Set rstImage = New ADODB.Recordset
        rstImage.Open "Select * from [" & strImageTableName & "]" & GetSort, dbConn, adOpenStatic, adLockOptimistic, adCmdText
        rstImage.MoveLast
        rstImage.MoveFirst
        Me.cboFields(3).Clear
        For i = 0 To rstImage.Fields.count - 1
            cboFields(3).AddItem rstImage.Fields(i).Name
        Next i
        cboFields(3).Enabled = True
    End If

End Sub

Private Sub chkGrpFVis_Click(Index As Integer)

    If chkGrpFVis(Index).value = 0 Then
        If SectionHasControls(10 - Index) Then
            MsgBox "This section has existing controls!" & Chr$(13) & _
            "Please delete all controls first before hiding this section", vbCritical
            chkGrpFVis(Index).value = 1
            Exit Sub
        End If
    End If
    
    GroupFVis(Index) = -(chkGrpFVis(Index).value)

End Sub

Private Sub chkGrpHVis_Click(Index As Integer)

    If chkGrpHVis(Index).value = 0 Then
        If SectionHasControls(Index + 2) Then
            MsgBox "This section has existing controls!" & Chr$(13) & _
            "Please delete all controls first before hiding this section", vbCritical
            chkGrpHVis(Index).value = 1
            Exit Sub
        End If
    End If
    
    GroupHVis(Index) = -(chkGrpHVis(Index).value)

End Sub

Private Function SectionHasControls(SecNo As Integer) As Boolean
Dim a As Integer
    
    For a = 0 To frmDesign.Controls.count - 1
        Set ctlTest = frmDesign.Controls(a)
        If ctlTest.Tag = SecNo Then
            SectionHasControls = True
            Exit For
        End If
    Next a
        
End Function

Private Sub chkHasPics_Click()
'fired when the user clicks on the check box to select/deselect if there are images associated with records
Dim blnOnOff As Boolean

    blnOnOff = -chkHasPics.value

    Me.txtPicFolder.Enabled = blnOnOff
    Me.cboTables(1).Enabled = blnOnOff
    Me.cboFields(3).Enabled = blnOnOff
    Me.cmdFindPicFolder.Enabled = blnOnOff
    Me.cboFolderPath.Enabled = blnOnOff
    
    Me.lblImagePathHead.Enabled = blnOnOff
    Me.lblImgTable.Enabled = blnOnOff
    Me.lblImgField.Enabled = blnOnOff
    Me.lblSelPath.Enabled = blnOnOff
    Me.lblImgPath.Enabled = blnOnOff
    Me.lblOR.Enabled = blnOnOff
    Me.txtPicFolder.Enabled = blnOnOff
    Me.cmdFindPicFolder.Enabled = blnOnOff
    
    Me.cboTables(1) = ""
    Me.cboFields(3) = ""
    If txtPicFolder.Enabled Then    'user has selected the check box
        Me.cboTables(1).BackColor = vbWhite
        Me.cboFields(3).BackColor = vbWhite
    Else                                        'user has deselected the check box
        Me.cboTables(1).BackColor = &H8000000F
        Me.cboFields(3).BackColor = &H8000000F
        strPhotoFolder = ""
    End If
    
    blnHasPics = blnOnOff

End Sub

Private Sub cmdBrowse_Click()
'fired on button click - displays the open file dialog box using the onboard CommonDialog control
'and sets variables to the selected file name
On Error GoTo NoFile

    CommonDialog1.ShowOpen
    txtDataFile.text = CommonDialog1.FileName
    strDataFileName = CommonDialog1.FileName
    Me.lblConnStatus.Caption = ""
    Me.cmdConnect.Enabled = True
    Exit Sub

NoFile:

End Sub

Private Sub cmdCancel_Click()
'fired on button click - exits the form and resets data connection to the way it was
    If strOldDataFileName > "" Then
        strDataFileName = strOldDataFileName
        ConnectToDataFile
        If strOldTableName > "" Then
            If GetTables Then
                strTableName = strOldTableName
                OpenData strTableName
                LoadFieldNames
            End If
        Set rstTables = Nothing
        Set rstImage = Nothing
        End If
    End If
    blnPageChanged = False
    Me.Hide

End Sub

Private Sub cmdConnect_Click()
'fired on button click - calls ConnectToDataFile to set up ADO data connection - currently rigged for MS Access
'but I think it could be easily modified
    
    Me.lblConnStatus.Caption = "Connecting to database..."
    DoEvents
    
    If ConnectToDataFile Then
        blnConnected = True
        Me.lblConnStatus.Caption = "Connection successful."
        If strConnErrMsg > "" Then              'this indicates that stored connection
            OpenData strTableName, strSort      'attempted on opening existing file was not
            GetTables                           'successful - if so then re-establish the
            LoadTableCombos                     'stored data parameters (assumes that you
            LoadFieldNames                      're-connected to the right database)
            cboTables(0) = strTableName
            blnReportDataBound = True
            LoadDataSelections
            strConnErrMsg = ""
            blnReportSaved = False
            Exit Sub
        End If
    Else
        Me.lblConnStatus.Caption = "Error : " & strConnErrMsg
        Exit Sub
    End If
    
    If GetTables Then           'if we were successful getting table names from the database then
        LoadTableCombos
    Else
        lblConnStatus.Caption = "Error : " & strConnErrMsg
    End If

End Sub

Private Sub LoadTableCombos()

    cboTables(0).Clear      'populate the appropriate combo boxes with them
    cboTables(1).Clear
    While Not rstTables.EOF     'this is the recordset opened in call to GetTables
        If rstTables!TABLE_TYPE <> "ACCESS TABLE" And rstTables!TABLE_TYPE <> "SYSTEM TABLE" And rstTables!TABLE_TYPE <> "LINK" Then
            If rstTables!TABLE_TYPE = "TABLE" Then
                Me.cboTables(0).AddItem "(Table)  " & rstTables!TABLE_NAME
                Me.cboTables(1).AddItem rstTables!TABLE_NAME
            ElseIf rstTables!TABLE_TYPE = "VIEW" Then
                Me.cboTables(0).AddItem "(Query)  " & rstTables!TABLE_NAME
            End If
        End If
        rstTables.MoveNext
    Wend
    rstTables.Close
    cboTables(0).BackColor = vbWhite
    cboTables(0).Enabled = True

End Sub

Private Sub cmdFindPicFolder_Click()

'calls Windows API function to browse for a folder
'returns result to the textbox if successful
    Dim InDir As String
        
    InDir = BrowseFolder(Me.hwnd, "Select folder containing images")
    If InDir = "" Then
        Exit Sub
    End If
    
    strImageFolder = InDir
    Me.txtPicFolder = InDir

End Sub

Private Sub cmdOK_Click()

 Dim i As Integer

'build sort portion of SQL statement
    strSort = ""
    For i = 0 To 2
        If cboFields(i).text > "" Then
            strSortField(i) = cboFields(i).text
            blnSortDescending(i) = Me.OptDesc(i).value
            If strSort > "" Then strSort = strSort & ", "
            strSort = strSort & cboFields(i).text
            If OptDesc(i) = True Then strSort = strSort & " DESC"
        End If
    Next i
    If strSort <> "" Then strSort = "ORDER BY " & strSort

'if table or query selected then call sub to open it and load the field names
'into the frmSelField listbox
    If strTableName > "" Then
        strDataFileName = txtDataFile.text
        OpenData strTableName, strSort
        LoadFieldNames
        blnReportDataBound = True
        frmDesign.StatusBar1.Panels(2).text = "Data Source = " & _
        cboTables(0).List(cboTables(0).ListIndex) & " - from " & Right(strDataFileName, Len(strDataFileName) - InStrRev(strDataFileName, "\"))
        For i = 2 To 4
            frmDesign.cmdDivider(i).Visible = GroupHVis(i - 2)
            frmDesign.cmdDivider(i).Caption = "'" & cboFields(i - 2) & "' Group Header"
            frmDesign.picSection(i).Visible = GroupHVis(i - 2)
            frmDesign.cmdDivider(10 - i).Visible = GroupFVis(i - 2)
            frmDesign.cmdDivider(10 - i).Caption = "'" & cboFields(i - 2) & "' Group Footer"
            frmDesign.picSection(10 - i).Visible = GroupFVis(i - 2)
        Next i
        blnPageChanged = True
        strConnErrMsg = ""
        For i = 0 To 2
            blnNewPageOnHeader(i) = -(chkNewPage(i).value)
        Next i
        strImgPathTable = Me.cboTables(1)
        strImgPathField = Me.cboFields(3)
        strImageFolder = Me.txtPicFolder

        Me.Hide
    Else
        If MsgBox("You have not selected a table or query as a data source for your report" & Chr$(13) _
        & "Do you wish to abort the data connection?", vbYesNo + vbQuestion) = vbYes Then
            If blnConnected Then
                dbConn.Close
                Set dbConn = Nothing
            End If
            strDataFileName = ""
            Unload Me
            Set frmDataConn = Nothing
        End If
    End If
    
    blnReportSaved = False

End Sub

Private Sub cmdViewData_Click()

    blnViewingData = True
    frmViewData.Show vbModal

End Sub

Private Sub Form_Activate()

'sets up listboxes and textboxes if there is already a data connection established
'checks to see if there is already a data connection established

    If blnViewingData Then
        blnViewingData = False
        Exit Sub
    End If

    If strConnErrMsg > "" Then
        txtDataFile.text = "<Enter database file name or press 'Browse'>"
        Exit Sub
    End If
    
    If strDataFileName > "" Then strOldDataFileName = strDataFileName
    If strTableName > "" Then strOldTableName = strTableName

    LoadDataSelections

End Sub

Private Sub LoadDataSelections()
Dim i As Integer

    If blnReportDataBound Then
        cmdConnect.Enabled = True
        If strDataFileName > "" Then
            txtDataFile.text = strDataFileName
            If GetTables And cboTables(0).ListCount = 0 Then
                LoadTableCombos
            End If
            If strTableName > "" Then
                cboTables(0) = strTableName
                cboTables(0).text = strTableName
                cboTables(0).Enabled = True
                cboTables(0).BackColor = vbWhite
                Me.chkHasPics.Enabled = True
                LoadSortFieldCombos
                For i = 0 To 2
                    cboFields(i) = strSortField(i)
                    If cboFields(i) > "" Then FraOrder(i).Enabled = True
                    chkGrpHVis(i).value = -(GroupHVis(i))
                    chkGrpHVis(i).Enabled = GroupHVis(i)
                    chkGrpFVis(i).value = -(GroupFVis(i))
                    chkGrpFVis(i).Enabled = GroupFVis(i)
                    chkNewPage(i).value = -(blnNewPageOnHeader(i))
                    chkNewPage(i).Enabled = GroupHVis(i)
                Next i
                cmdViewData.Enabled = True
                lblSortField.Enabled = True
                lblHeader.Enabled = True
                lblFooter.Enabled = True
                lbl1st.Enabled = True
                lbl2nd.Enabled = True
                lbl3rd.Enabled = True
                Me.chkHasPics.value = -(strImageFolder > "")
                chkHasPics_Click
                If strImgPathTable > "" Then cboTables(1) = strImgPathTable
                If strImgPathField > "" Then cboFields(3) = strImgPathField
                If strImageFolder > "" Then
                    txtPicFolder.text = strImageFolder
                    Me.cboFolderPath = strImageFolder
                End If
            End If
        End If
    Else
        cmdConnect.Enabled = False
        txtDataFile.text = ""
        cboTables(0) = ""
        cboTables(0).Clear
        cboTables(0).Enabled = False
        cboTables(1) = ""
        cboTables(1).Clear
        cboTables(1).Enabled = False
        For i = 0 To 2
            FraOrder(i).Enabled = False
            cboFields(i) = ""
            cboFields(i).Clear
            cboFields(i).Enabled = False
            chkGrpHVis(i).value = 0
            chkGrpHVis(i).Enabled = False
            chkGrpFVis(i).value = 0
            chkGrpFVis(i).Enabled = False
        Next i
        chkHasPics.value = -blnHasPics
        chkHasPics_Click
        chkHasPics.Enabled = False
        txtPicFolder.text = ""
    End If

End Sub

Private Sub LoadSortFieldCombos()
Dim i As Integer

    For i = 0 To 2
        cboFields(i).Clear
        cboFields(i).Enabled = True
        cboFields(i).BackColor = vbWhite
    Next i
    
    For i = 0 To UBound(DataField) - 1
        cboFields(0).AddItem rstData.Fields(i).Name
        cboFields(1).AddItem rstData.Fields(i).Name
        cboFields(2).AddItem rstData.Fields(i).Name
    Next i
    
End Sub
