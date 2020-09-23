VERSION 5.00
Begin VB.Form frmPreview 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3150
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "ActiveReporter"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   174
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPreview 
      BorderStyle     =   0  'None
      Height          =   2595
      Left            =   0
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   569
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.PictureBox picViewPort 
         BackColor       =   &H8000000C&
         Height          =   1995
         Left            =   0
         ScaleHeight     =   1935
         ScaleWidth      =   2835
         TabIndex        =   4
         Top             =   360
         Width           =   2895
         Begin VB.PictureBox picPage 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DragIcon        =   "frmPreview.frx":014A
            ForeColor       =   &H80000008&
            Height          =   1755
            Left            =   840
            MouseIcon       =   "frmPreview.frx":0454
            ScaleHeight     =   117
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   85
            TabIndex        =   5
            Top             =   60
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.PictureBox picBack 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            DragIcon        =   "frmPreview.frx":075E
            ForeColor       =   &H80000008&
            Height          =   1755
            Left            =   900
            MouseIcon       =   "frmPreview.frx":0A68
            ScaleHeight     =   115
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   83
            TabIndex        =   8
            Top             =   120
            Visible         =   0   'False
            Width           =   1275
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1995
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   240
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         Left            =   0
         TabIndex        =   2
         Top             =   2340
         Width           =   2895
      End
      Begin VB.PictureBox picImages 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   0
         Picture         =   "frmPreview.frx":0D72
         ScaleHeight     =   240
         ScaleWidth      =   3120
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.Label lblRange 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "All"
         Height          =   300
         Left            =   7260
         TabIndex        =   7
         Top             =   15
         Width           =   1425
      End
      Begin VB.Label lblPageIndex 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No pages"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5820
         TabIndex        =   6
         Top             =   15
         Width           =   1425
      End
   End
   Begin VB.Image imgDither 
      Height          =   120
      Left            =   3240
      Picture         =   "frmPreview.frx":1474
      Top             =   300
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Menu mnuZoom 
      Caption         =   "Zoom"
      Visible         =   0   'False
      Begin VB.Menu mnuRatio 
         Caption         =   "100 %"
         Index           =   0
      End
      Begin VB.Menu mnuRatio 
         Caption         =   "75 %"
         Index           =   1
      End
      Begin VB.Menu mnuRatio 
         Caption         =   "50 %"
         Index           =   2
      End
      Begin VB.Menu mnuRatio 
         Caption         =   "25 %"
         Index           =   3
      End
      Begin VB.Menu mnuRatio 
         Caption         =   "Page Width"
         Index           =   4
      End
      Begin VB.Menu mnuRatio 
         Caption         =   "Full Page"
         Index           =   5
      End
      Begin VB.Menu mnuRatio 
         Caption         =   "Custom"
         Index           =   6
      End
   End
   Begin VB.Menu mnuExport 
      Caption         =   "Export"
      Visible         =   0   'False
      Begin VB.Menu mnuFormat 
         Caption         =   "Text"
         Index           =   0
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "CSV"
         Index           =   1
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "HTML"
         Index           =   2
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "RTF"
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "JPG"
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuSet 
         Caption         =   "Choose Printer"
         Index           =   0
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Select pages"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////
'///                Preview Support Form
'///                  (frmPreview.frm)
'///_____________________________________________________
'/// Form tha lets us show  the preview viewport when no
'/// external container is specified, preview design is
'/// allowed too.
'///_____________________________________________________
'/// Last modification  : Ago/10/2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Created
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Event ChangeZoomRatio(iRatioIndex As Integer)
Public Event DoScroll(lValue As Long, bVertical As Boolean)
Public Event ToolSelected(iToolIndex As Integer)
Public Event ExportAction(iFormatIndex As Integer)
Public Event GotoPage()

Public Sub Start()
    WindowState = vbMaximized
    picPreview_Resize
    Show vbModal
End Sub

Private Sub Form_Resize()
    With picPreview
        .Move 0, 0, ScaleWidth, ScaleHeight
    End With
End Sub

Private Sub HScroll1_Change()
    RaiseEvent DoScroll(HScroll1.Value, False)
End Sub

Private Sub lblPageIndex_DblClick()
    RaiseEvent GotoPage
End Sub

Private Sub lblRange_DblClick()
    mnuSet_Click (1)
End Sub

Private Sub mnuFormat_Click(Index As Integer)
    RaiseEvent ExportAction(Index)
End Sub

Private Sub mnuRatio_Click(Index As Integer)
    RaiseEvent ChangeZoomRatio(Index)
End Sub

Private Sub mnuSet_Click(Index As Integer)
    RaiseEvent ToolSelected(Index)
End Sub

Private Sub picPreview_Resize()
    LockWindowUpdate picPreview.hWnd
    With picViewPort
        .Move 0, .Top, (picPreview.ScaleWidth - 16), (picPreview.ScaleHeight - .Top - 16)
    End With
    With picViewPort
        VScroll1.Move .Width, .Top, 16, .Height
        HScroll1.Move 0, (.Height + .Top), .Width, 16
    End With
    LockWindowUpdate 0
End Sub

Private Sub VScroll1_Change()
    RaiseEvent DoScroll(VScroll1.Value, True)
End Sub

