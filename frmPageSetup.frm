VERSION 5.00
Begin VB.Form frmPageSetup 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Page Setup"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   Icon            =   "frmPageSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2.458
   ScaleMode       =   5  'Inch
   ScaleWidth      =   4.427
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmUnits 
      BackColor       =   &H8000000B&
      Caption         =   " Measurement Units "
      Height          =   1155
      Left            =   180
      TabIndex        =   21
      Top             =   120
      Width           =   1755
      Begin VB.OptionButton optEng 
         BackColor       =   &H8000000B&
         Caption         =   "English"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optMetric 
         BackColor       =   &H8000000B&
         Caption         =   "Metric (SI)"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   1035
      End
   End
   Begin VB.TextBox txtPageHt 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   3060
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1770
      Width           =   525
   End
   Begin VB.TextBox txtPageWid 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2100
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1770
      Width           =   525
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
      Left            =   5220
      TabIndex        =   8
      Top             =   3060
      Width           =   975
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
      Left            =   4140
      TabIndex        =   7
      Top             =   3060
      Width           =   975
   End
   Begin VB.ComboBox cboPageSize 
      Height          =   315
      ItemData        =   "frmPageSetup.frx":0442
      Left            =   150
      List            =   "frmPageSetup.frx":044F
      TabIndex        =   2
      Text            =   "Letter"
      Top             =   1770
      Width           =   1605
   End
   Begin VB.TextBox txtBM 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1980
      TabIndex        =   6
      Top             =   3060
      Width           =   495
   End
   Begin VB.TextBox txtTM 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1980
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtRM 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2700
      TabIndex        =   4
      Top             =   2670
      Width           =   495
   End
   Begin VB.TextBox txtLM 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   2700
      Width           =   495
   End
   Begin VB.Frame FraOrientation 
      BackColor       =   &H8000000B&
      Caption         =   "Paper Orientation"
      Height          =   1155
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1545
      Begin VB.OptionButton OptLandscape 
         BackColor       =   &H8000000B&
         Caption         =   "Landscape"
         Height          =   255
         Left            =   270
         TabIndex        =   9
         Top             =   720
         Width           =   1125
      End
      Begin VB.OptionButton OptPortrait 
         BackColor       =   &H8000000B&
         Caption         =   "Portrait"
         Height          =   255
         Left            =   270
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   825
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "X"
      Height          =   195
      Left            =   2700
      TabIndex        =   20
      Top             =   1830
      Width           =   285
   End
   Begin VB.Label lblHeight 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Height (in.)"
      Height          =   285
      Left            =   2940
      TabIndex        =   19
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label lblWidth 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Width (in.)"
      Height          =   225
      Left            =   1920
      TabIndex        =   17
      Top             =   1500
      Width           =   945
   End
   Begin VB.Label lblPageMarg 
      BackColor       =   &H8000000B&
      Caption         =   "Page Margins  (Inches)"
      Height          =   465
      Left            =   150
      TabIndex        =   15
      Top             =   2220
      Width           =   1035
   End
   Begin VB.Shape shpMargBord 
      BorderStyle     =   3  'Dot
      Height          =   1440
      Left            =   4536
      Top             =   360
      Width           =   1080
   End
   Begin VB.Shape shpPage 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   1584
      Index           =   1
      Left            =   4464
      Top             =   288
      Width           =   1224
   End
   Begin VB.Shape shpPage 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1584
      Index           =   0
      Left            =   4507
      Top             =   331
      Width           =   1224
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000B&
      Caption         =   "Paper size"
      Height          =   255
      Left            =   150
      TabIndex        =   14
      Top             =   1530
      Width           =   765
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      Caption         =   "Bottom"
      Height          =   225
      Left            =   1980
      TabIndex        =   13
      Top             =   2820
      Width           =   555
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Top"
      Height          =   225
      Left            =   1980
      TabIndex        =   12
      Top             =   2160
      Width           =   525
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Right"
      Height          =   225
      Left            =   2700
      TabIndex        =   11
      Top             =   2460
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "Left"
      Height          =   225
      Left            =   1260
      TabIndex        =   10
      Top             =   2460
      Width           =   495
   End
End
Attribute VB_Name = "frmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ScaleFactor As Single
Private Type PageParameters
    Name As String
    width As Single
    Height As Single
    PgScale As Single
End Type

Private Type MargRect
    Left As Single
    Right As Single
    Top As Single
    Bottom As Single
End Type

Dim Margin As MargRect

Dim PageSize(3) As PageParameters
Dim CurrPageWd As Single
Dim CurrPageHt As Single
Dim UnitScaleMod As Single
Dim CurrUnits As Integer
Dim blnLoading As Boolean

Private Sub LoadPageSizes()

    PageSize(0).Name = "Letter"
    PageSize(0).width = 8.5
    PageSize(0).Height = 11
    PageSize(0).PgScale = 1

    PageSize(1).Name = "Legal"
    PageSize(1).width = 8.5
    PageSize(1).Height = 14
    PageSize(1).PgScale = 0.9

    PageSize(2).Name = "B-Size"
    PageSize(2).width = 11
    PageSize(2).Height = 17
    PageSize(2).PgScale = 0.7

End Sub

Private Sub SetPageDimensions()
Dim i As Integer
    
    For i = 0 To 3
        If cboPageSize = PageSize(i).Name Then
            CurrPageWd = PageSize(i).width * UnitScaleMod
            CurrPageHt = PageSize(i).Height * UnitScaleMod
            If OptPortrait Then
                txtPageWid = CurrPageWd
                txtPageHt = CurrPageHt
            ElseIf OptLandscape Then
                txtPageWid = CurrPageHt
                txtPageHt = CurrPageWd
            End If
            ScaleFactor = PageSize(i).PgScale
            Exit Sub
        End If
    Next i

End Sub

Private Sub UpdatePreview()

    If OptLandscape Then
        shpPage(0).Left = 3.03
        shpPage(0).Top = 0.33
        shpPage(1).Left = 3
        shpPage(1).Top = 0.3
    Else
        shpPage(0).Left = 3.13
        shpPage(0).Top = 0.23
        shpPage(1).Left = 3.1
        shpPage(1).Top = 0.2
    End If

    shpPage(0).width = txtPageWid / UnitScaleMod * 0.1 * ScaleFactor
    shpPage(0).Height = txtPageHt / UnitScaleMod * 0.1 * ScaleFactor
    shpPage(1).width = txtPageWid / UnitScaleMod * 0.1 * ScaleFactor
    shpPage(1).Height = txtPageHt / UnitScaleMod * 0.1 * ScaleFactor
    shpMargBord.Left = shpPage(1).Left + ((Val(txtLM) / UnitScaleMod) / 10) * ScaleFactor
    shpMargBord.Top = shpPage(1).Top + ((Val(txtTM) / UnitScaleMod) / 10) * ScaleFactor
    shpMargBord.width = shpPage(1).width - ((Val(txtLM) / UnitScaleMod) / 10) * ScaleFactor - ((Val(txtRM) / UnitScaleMod) / 10) * ScaleFactor
    shpMargBord.Height = shpPage(1).Height - ((Val(txtTM) / UnitScaleMod) / 10) * ScaleFactor - ((Val(txtBM) / UnitScaleMod) / 10) * ScaleFactor

End Sub

Private Sub cboPageSize_Click()

    SetPageDimensions
    UpdatePreview

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    blnPageWidChanged = True
    PageSizeName = cboPageSize
    If OptPortrait Then
        PageWd = CurrPageWd / UnitScaleMod
        PageHt = CurrPageHt / UnitScaleMod
    Else
        PageWd = CurrPageHt / UnitScaleMod
        PageHt = CurrPageWd / UnitScaleMod
    End If
    LeftMarg = Val(txtLM) / UnitScaleMod
    RightMarg = Val(txtRM) / UnitScaleMod
    TopMarg = Val(txtTM) / UnitScaleMod
    BottomMarg = Val(txtBM) / UnitScaleMod
    If OptPortrait Then
        PageOrient = cPortrait
    Else
        PageOrient = cLandscape
    End If
    frmDesign.ShowHorizScale
    frmDesign.ShowVertScale
    ShowGrid
    If PageScaleUnits <> CurrUnits And blnSnapOn Then
        frmGridSpace.Show
    End If
    Unload Me

End Sub

Private Sub Form_Load()

    blnLoading = True
    CurrUnits = PageScaleUnits
    cboPageSize = PageSizeName
    If PageScaleUnits = scEnglish Then
        UnitScaleMod = 1
        optEng.value = True
        lblWidth = "Width (in.)"
        lblHeight = "Height (in.)"
        lblPageMarg.Caption = "Page Margins  (Inches)"
    ElseIf PageScaleUnits = scMetric Then
        UnitScaleMod = 2.54
        optMetric.value = True
        lblWidth = "Width (cm)"
        lblHeight = "Height (cm)"
        lblPageMarg.Caption = "Page Margins  (cm)"
    End If
    
    If PageOrient = cPortrait Then
        OptPortrait = True
    Else
        OptLandscape = True
    End If
    
    cboPageSize.text = PageSizeName
    LoadPageSizes
    Margin.Left = LeftMarg
    Margin.Right = RightMarg
    Margin.Top = TopMarg
    Margin.Bottom = BottomMarg
    
    txtLM.text = Margin.Left * UnitScaleMod
    txtTM.text = Margin.Top * UnitScaleMod
    txtRM.text = Margin.Right * UnitScaleMod
    txtBM.text = Margin.Bottom * UnitScaleMod
    SetPageDimensions
    UpdatePreview
    blnLoading = False

End Sub

Private Sub optEng_Click()

    PageScaleUnits = scEnglish
    UnitScaleMod = 1
    lblWidth = "Width (in.)"
    lblHeight = "Height (in.)"
    lblPageMarg.Caption = "Page Margins  (Inches)"
    txtLM.text = Margin.Left * UnitScaleMod
    txtTM.text = Margin.Top * UnitScaleMod
    txtRM.text = Margin.Right * UnitScaleMod
    txtBM.text = Margin.Bottom * UnitScaleMod
    SetPageDimensions

End Sub

Private Sub OptLandscape_Click()

    SetPageDimensions
    UpdatePreview

End Sub

Private Sub optMetric_Click()

    PageScaleUnits = scMetric
    UnitScaleMod = 2.54
    lblWidth = "Width (cm)"
    lblHeight = "Height (cm)"
    lblPageMarg.Caption = "Page Margins  (cm)"
    txtLM.text = Margin.Left * UnitScaleMod
    txtTM.text = Margin.Top * UnitScaleMod
    txtRM.text = Margin.Right * UnitScaleMod
    txtBM.text = Margin.Bottom * UnitScaleMod
    SetPageDimensions

End Sub

Private Sub OptPortrait_Click()

    SetPageDimensions
    UpdatePreview

End Sub

Private Sub txtBM_GotFocus()

    txtBM.SelStart = 0
    txtBM.SelLength = Len(txtBM)

End Sub

Private Sub txtBM_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Margin.Bottom = Val(txtBM)
        UpdatePreview
    End If

End Sub

Private Sub txtLM_GotFocus()

    txtLM.SelStart = 0
    txtLM.SelLength = Len(txtLM)

End Sub

Private Sub txtLM_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Margin.Left = Val(txtLM)
        UpdatePreview
    End If

End Sub

Private Sub txtRM_GotFocus()
    
    txtRM.SelStart = 0
    txtRM.SelLength = Len(txtRM)

End Sub

Private Sub txtRM_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Margin.Right = Val(txtRM)
        UpdatePreview
    End If

End Sub

Private Sub txtTM_GotFocus()

    txtTM.SelStart = 0
    txtTM.SelLength = Len(txtTM)

End Sub

Private Sub txtTM_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Margin.Top = Val(txtTM)
        UpdatePreview
    End If

End Sub
