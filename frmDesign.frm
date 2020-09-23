VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmDesign 
   AutoRedraw      =   -1  'True
   Caption         =   "RG Report Designer"
   ClientHeight    =   9015
   ClientLeft      =   -495
   ClientTop       =   975
   ClientWidth     =   12375
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000015&
   Icon            =   "frmDesign.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmDesign.frx":0442
   ScaleHeight     =   6.26
   ScaleMode       =   5  'Inch
   ScaleWidth      =   8.594
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   390
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   688
      BandCount       =   2
      EmbossPicture   =   -1  'True
      _CBWidth        =   12375
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   0.229
      Width1          =   3.396
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "Toolbar2"
      MinHeight2      =   0.229
      Width2          =   1.958
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   165
         TabIndex        =   57
         Top             =   30
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "MenuImages"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "New File"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open File"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save File"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   0.15
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Print Report"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Preview"
               Object.ToolTipText     =   "Preview Report"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   0.15
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Paste"
               Object.ToolTipText     =   "Paste"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   0.15
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Undo"
               Object.ToolTipText     =   "Undo"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "DesignGrid"
               Object.ToolTipText     =   "Design Grid Settings"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "DataConnect"
               Object.ToolTipText     =   "Connect to Data Source"
               ImageIndex      =   19
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "ViewData"
               Object.ToolTipText     =   "View Data"
               ImageIndex      =   20
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   5085
         TabIndex        =   51
         Top             =   30
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "MenuImages"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   17
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "Bold"
               ImageIndex      =   11
               Style           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "Italic"
               ImageIndex      =   12
               Style           =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Object.ToolTipText     =   "Underline"
               ImageIndex      =   13
               Style           =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "LeftAlign"
               Object.ToolTipText     =   "Left Align"
               Object.Tag             =   "0"
               ImageIndex      =   14
               Style           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "CenterAlign"
               Object.ToolTipText     =   "Center Align"
               Object.Tag             =   "2"
               ImageIndex      =   15
               Style           =   1
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "RightAlign"
               Object.ToolTipText     =   "Right Align"
               Object.Tag             =   "1"
               ImageIndex      =   16
               Style           =   1
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "TextColor"
               ImageIndex      =   23
               Style           =   5
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   0.05
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "LineColor"
               ImageIndex      =   21
               Style           =   5
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   0.05
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FillColor"
               ImageIndex      =   22
               Style           =   5
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   0.05
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "LineStyle"
               ImageIndex      =   24
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   0.05
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "LineWidth"
               ImageIndex      =   25
            EndProperty
         EndProperty
         Begin VB.ComboBox cboFontSize 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmDesign.frx":074C
            Left            =   8280
            List            =   "frmDesign.frx":077D
            TabIndex        =   56
            Text            =   "10"
            Top             =   20
            Width           =   645
         End
         Begin Project1.FontSelector FontSelector1 
            Height          =   315
            Left            =   5340
            TabIndex        =   55
            Top             =   20
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
         End
         Begin VB.PictureBox picTxtCol 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   60
            Left            =   2355
            ScaleHeight     =   4
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   16
            TabIndex        =   54
            Top             =   240
            Width           =   240
         End
         Begin VB.PictureBox picLinCol 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   60
            Left            =   3015
            ScaleHeight     =   60
            ScaleWidth      =   240
            TabIndex        =   53
            Top             =   240
            Width           =   240
         End
         Begin VB.PictureBox picFillCol 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   60
            Left            =   3660
            ScaleHeight     =   30
            ScaleWidth      =   210
            TabIndex        =   52
            Top             =   240
            Width           =   240
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   33
      Top             =   8700
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   556
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3810
            MinWidth        =   3810
            Key             =   "Coords"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12700
            MinWidth        =   12700
            Key             =   "Contents"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12700
            MinWidth        =   12700
            Key             =   "DataConn"
         EndProperty
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.CheckBoxControl Chkbox 
      Height          =   240
      Index           =   0
      Left            =   11760
      TabIndex        =   32
      Top             =   1845
      Visible         =   0   'False
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   423
   End
   Begin Project1.PropertySelector PropertySelector1 
      Height          =   2595
      Left            =   8460
      TabIndex        =   30
      Top             =   6000
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   4577
   End
   Begin MSComctlLib.ImageList MenuImages 
      Left            =   10425
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":07BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":0D55
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":12EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1889
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":1E23
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":23BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":2957
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":2EF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":348B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":3A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":3FBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":4559
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":4AF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":508D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":5627
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":5BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":615B
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":66F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":6C8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":7229
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":77C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":7D5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":82F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":8891
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesign.frx":8E2B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   10470
      TabIndex        =   26
      ToolTipText     =   "Press 'Enter' to accept changes"
      Top             =   2475
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdPgSetup 
      Height          =   490
      Left            =   540
      Picture         =   "frmDesign.frx":93C5
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   825
      Width           =   274
   End
   Begin VB.PictureBox picVertScale 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      ForeColor       =   &H00FFFFFF&
      Height          =   7275
      Left            =   540
      MousePointer    =   7  'Size N S
      ScaleHeight     =   5.01
      ScaleMode       =   5  'Inch
      ScaleWidth      =   0.156
      TabIndex        =   17
      Top             =   1350
      Width           =   285
      Begin VB.Line lnPos 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   -0.042
         X2              =   0.167
         Y1              =   0.167
         Y2              =   0.167
      End
   End
   Begin VB.PictureBox picHorizScale 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      MousePointer    =   9  'Size W E
      ScaleHeight     =   0.156
      ScaleMode       =   5  'Inch
      ScaleWidth      =   6.427
      TabIndex        =   16
      Top             =   825
      Width           =   9315
      Begin VB.PictureBox picWidMark 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1920
         Picture         =   "frmDesign.frx":994F
         ScaleHeight     =   0.135
         ScaleMode       =   5  'Inch
         ScaleWidth      =   0.156
         TabIndex        =   23
         Top             =   0
         Width           =   230
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10455
      Top             =   3915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Report Files (*.drp)|*.drp"
      FontName        =   "Arial"
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6600
      Left            =   60
      TabIndex        =   9
      Top             =   795
      Width           =   435
      Begin VB.OptionButton optAction 
         DownPicture     =   "frmDesign.frx":9ED9
         Height          =   520
         Index           =   9
         Left            =   0
         Picture         =   "frmDesign.frx":A463
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Place Calculated Field"
         Top             =   2820
         Width           =   400
      End
      Begin VB.OptionButton optAction 
         DownPicture     =   "frmDesign.frx":A9ED
         Height          =   520
         Index           =   10
         Left            =   0
         Picture         =   "frmDesign.frx":AF77
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Place Summary Field"
         Top             =   3345
         Width           =   400
      End
      Begin VB.OptionButton optAction 
         DownPicture     =   "frmDesign.frx":B501
         Height          =   520
         Index           =   8
         Left            =   0
         Picture         =   "frmDesign.frx":BA8B
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Place Date/Page No. Field"
         Top             =   2295
         Width           =   400
      End
      Begin VB.OptionButton optAction 
         DownPicture     =   "frmDesign.frx":C015
         Height          =   520
         Index           =   3
         Left            =   0
         Picture         =   "frmDesign.frx":C59F
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Place Check Box"
         Top             =   1245
         Width           =   400
      End
      Begin VB.OptionButton optAction 
         DownPicture     =   "frmDesign.frx":CB29
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Index           =   7
         Left            =   0
         Picture         =   "frmDesign.frx":D0B3
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Place Data Bound Image"
         Top             =   1770
         Width           =   400
      End
      Begin VB.OptionButton optAction 
         DownPicture     =   "frmDesign.frx":D63D
         Height          =   520
         Index           =   0
         Left            =   0
         Picture         =   "frmDesign.frx":DBC7
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Select object"
         Top             =   30
         Width           =   400
      End
      Begin VB.OptionButton optAction 
         DownPicture     =   "frmDesign.frx":E151
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Index           =   6
         Left            =   0
         Picture         =   "frmDesign.frx":E6DB
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Place Image"
         Top             =   4620
         Width           =   400
      End
      Begin VB.OptionButton optAction 
         DownPicture     =   "frmDesign.frx":EC65
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Index           =   1
         Left            =   0
         Picture         =   "frmDesign.frx":F1EF
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Place Data Field"
         Top             =   720
         Width           =   400
      End
      Begin VB.OptionButton optAction 
         Appearance      =   0  'Flat
         DownPicture     =   "frmDesign.frx":F779
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   520
         Index           =   2
         Left            =   0
         Picture         =   "frmDesign.frx":FD03
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Place Label"
         Top             =   4080
         Width           =   400
      End
      Begin VB.OptionButton optAction 
         DownPicture     =   "frmDesign.frx":1028D
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Index           =   5
         Left            =   0
         Picture         =   "frmDesign.frx":10817
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Place Shape"
         Top             =   5700
         Width           =   400
      End
      Begin VB.OptionButton optAction 
         DownPicture     =   "frmDesign.frx":10DA1
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Index           =   4
         Left            =   0
         Picture         =   "frmDesign.frx":1132B
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Place Line"
         Top             =   5160
         Width           =   400
      End
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      LargeChange     =   10
      Left            =   840
      Max             =   10
      TabIndex        =   5
      Top             =   8280
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.VScrollBar VScroll 
      Height          =   2355
      LargeChange     =   10
      Left            =   9960
      Max             =   5
      TabIndex        =   4
      Top             =   1140
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picContainer 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   840
      ScaleHeight     =   4.969
      ScaleMode       =   5  'Inch
      ScaleWidth      =   6.156
      TabIndex        =   0
      Top             =   1106
      Width           =   8925
      Begin VB.Frame lnSelect 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   1
         Left            =   6540
         TabIndex        =   49
         Top             =   2880
         Visible         =   0   'False
         Width           =   58
      End
      Begin VB.Frame lnSelect 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1935
         Index           =   0
         Left            =   6000
         TabIndex        =   48
         Top             =   2880
         Visible         =   0   'False
         Width           =   58
      End
      Begin VB.PictureBox picSection 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FF8080&
         Height          =   225
         Index           =   8
         Left            =   3420
         ScaleHeight     =   0.156
         ScaleMode       =   5  'Inch
         ScaleWidth      =   1.208
         TabIndex        =   47
         Top             =   4380
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.CommandButton cmdDivider 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Group Footer 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   4140
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.PictureBox picSection 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FF8080&
         Height          =   225
         Index           =   7
         Left            =   3420
         ScaleHeight     =   0.156
         ScaleMode       =   5  'Inch
         ScaleWidth      =   1.208
         TabIndex        =   45
         Top             =   4980
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.CommandButton cmdDivider 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Group Footer 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   4740
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.PictureBox picSection 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FF8080&
         Height          =   225
         Index           =   6
         Left            =   3420
         ScaleHeight     =   0.156
         ScaleMode       =   5  'Inch
         ScaleWidth      =   1.208
         TabIndex        =   43
         Top             =   5700
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.CommandButton cmdDivider 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Group Footer 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   5460
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.CommandButton cmdDivider 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Group Header 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   5460
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.PictureBox picSection 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FF8080&
         Height          =   225
         Index           =   4
         Left            =   660
         ScaleHeight     =   0.156
         ScaleMode       =   5  'Inch
         ScaleWidth      =   1.208
         TabIndex        =   40
         Top             =   5700
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.CommandButton cmdDivider 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Group Header 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4740
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.PictureBox picSection 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FF8080&
         Height          =   225
         Index           =   3
         Left            =   660
         ScaleHeight     =   0.156
         ScaleMode       =   5  'Inch
         ScaleWidth      =   1.208
         TabIndex        =   38
         Top             =   4980
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.CommandButton cmdDivider 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Group Header 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4140
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.PictureBox picSection 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FF8080&
         Height          =   225
         Index           =   2
         Left            =   660
         ScaleHeight     =   0.156
         ScaleMode       =   5  'Inch
         ScaleWidth      =   1.208
         TabIndex        =   36
         Top             =   4380
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.PictureBox picSection 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FF8080&
         Height          =   540
         Index           =   10
         Left            =   245
         ScaleHeight     =   0.375
         ScaleMode       =   5  'Inch
         ScaleWidth      =   5.708
         TabIndex        =   22
         Top             =   3360
         Width           =   8220
      End
      Begin VB.CommandButton cmdDivider 
         Caption         =   "Report Footer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3139
         Width           =   8445
      End
      Begin VB.PictureBox picSection 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FF8080&
         Height          =   540
         Index           =   9
         Left            =   240
         ScaleHeight     =   0.375
         ScaleMode       =   5  'Inch
         ScaleWidth      =   5.708
         TabIndex        =   20
         Top             =   2580
         Width           =   8220
      End
      Begin VB.CommandButton cmdDivider 
         Caption         =   "Page Footer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2354
         Width           =   8445
      End
      Begin VB.PictureBox picSection 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FF8080&
         Height          =   540
         Index           =   1
         Left            =   245
         ScaleHeight     =   0.375
         ScaleMode       =   5  'Inch
         ScaleWidth      =   5.708
         TabIndex        =   8
         Top             =   1030
         Width           =   8220
      End
      Begin VB.PictureBox picSection 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FF8080&
         Height          =   540
         Index           =   5
         Left            =   245
         ScaleHeight     =   0.375
         ScaleMode       =   5  'Inch
         ScaleWidth      =   5.708
         TabIndex        =   7
         Top             =   1814
         Width           =   8220
      End
      Begin VB.CommandButton cmdDivider 
         Caption         =   "Report Header"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   8475
      End
      Begin VB.CommandButton cmdDivider 
         Caption         =   "Detail Section"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1570
         Width           =   8445
      End
      Begin VB.CommandButton cmdDivider 
         Caption         =   "Page Header"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   785
         Width           =   8445
      End
      Begin VB.PictureBox picSection 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   540
         Index           =   0
         Left            =   240
         ScaleHeight     =   0.375
         ScaleMode       =   5  'Inch
         ScaleWidth      =   5.708
         TabIndex        =   3
         Top             =   240
         Width           =   8220
      End
   End
   Begin VB.Label Field 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " New Label 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   75
      Index           =   0
      Left            =   10800
      TabIndex        =   35
      Top             =   2040
      Width           =   75
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " New Label 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   75
      Index           =   0
      Left            =   10500
      TabIndex        =   34
      Top             =   2040
      Width           =   75
   End
   Begin MSForms.Image Img 
      Height          =   105
      Index           =   0
      Left            =   11160
      Top             =   2040
      Visible         =   0   'False
      Width           =   120
      SizeMode        =   3
      Size            =   "212;185"
   End
   Begin VB.Shape Shape 
      Height          =   75
      Index           =   0
      Left            =   10650
      Top             =   2055
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label BoundImg 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   75
      Index           =   0
      Left            =   10980
      TabIndex        =   31
      Top             =   2040
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape SelectBox 
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   11520
      Top             =   1260
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Shape Grip 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   90
      Index           =   7
      Left            =   11100
      Top             =   1470
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape Grip 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   90
      Index           =   6
      Left            =   10920
      Top             =   1470
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape Grip 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   90
      Index           =   4
      Left            =   10560
      Top             =   1470
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape Grip 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   90
      Index           =   5
      Left            =   10740
      Top             =   1470
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape Grip 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   90
      Index           =   3
      Left            =   11100
      Top             =   1230
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape Grip 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   90
      Index           =   2
      Left            =   10920
      Top             =   1230
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape Grip 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   90
      Index           =   0
      Left            =   10560
      Top             =   1230
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape Grip 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   90
      Index           =   1
      Left            =   10740
      Top             =   1230
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Line Lin 
      Index           =   0
      Visible         =   0   'False
      X1              =   7.333
      X2              =   8.061
      Y1              =   1.604
      Y2              =   1.604
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu NewFile 
         Caption         =   "New"
      End
      Begin VB.Menu OpenFile 
         Caption         =   "Open..."
      End
      Begin VB.Menu SaveFile 
         Caption         =   "Save"
      End
      Begin VB.Menu SaveFileAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu PageSet 
         Caption         =   "Page Setup..."
      End
      Begin VB.Menu prnt 
         Caption         =   "Print..."
      End
      Begin VB.Menu PrntPrev 
         Caption         =   "Print Preview..."
      End
      Begin VB.Menu ExitProg 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCutObj 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopyObj 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPasteObj 
         Caption         =   "Paste"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDelObj 
         Caption         =   "Delete..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu ViewMenu 
      Caption         =   "View"
      Begin VB.Menu PageHdFt 
         Caption         =   "Page Header/Footer"
         Checked         =   -1  'True
      End
      Begin VB.Menu RepHdFt 
         Caption         =   "Report Header/Footer"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu ToolsMenu 
      Caption         =   "Tools"
      Begin VB.Menu GridSpacing 
         Caption         =   "Design Grid Settings..."
      End
      Begin VB.Menu DataConn 
         Caption         =   "Connect to Data Source..."
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu ReportObjectPopUp 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu EditText 
         Caption         =   "Edit Text"
         Visible         =   0   'False
      End
      Begin VB.Menu EditCalc 
         Caption         =   "Edit Calculation..."
         Visible         =   0   'False
      End
      Begin VB.Menu EditSummary 
         Caption         =   "Edit Summary Field..."
         Visible         =   0   'False
      End
      Begin VB.Menu EditDatePage 
         Caption         =   "Edit Date/Page Field..."
         Visible         =   0   'False
      End
      Begin VB.Menu AssignField 
         Caption         =   "Assign Data Field..."
         Visible         =   0   'False
      End
      Begin VB.Menu FormatText 
         Caption         =   "Set Data Format..."
         Visible         =   0   'False
      End
      Begin VB.Menu CutObj 
         Caption         =   "Cut"
         Enabled         =   0   'False
      End
      Begin VB.Menu CopyObj 
         Caption         =   "Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu PasteObj 
         Caption         =   "Paste"
         Enabled         =   0   'False
      End
      Begin VB.Menu DelObj 
         Caption         =   "Delete..."
         Enabled         =   0   'False
      End
      Begin VB.Menu BringFront 
         Caption         =   "Bring to front"
      End
      Begin VB.Menu SendBack 
         Caption         =   "Send to back"
      End
      Begin VB.Menu ToggleSnap 
         Caption         =   "Grid Snap"
      End
      Begin VB.Menu SetBackColor 
         Caption         =   "Set Background Color..."
      End
      Begin VB.Menu chk3d 
         Caption         =   "Sunken"
         Visible         =   0   'False
      End
      Begin VB.Menu ChkOptions 
         Caption         =   "Display as..."
         Begin VB.Menu chkCheck 
            Caption         =   "Check Box"
         End
         Begin VB.Menu chkRadio 
            Caption         =   "Radio Button"
         End
         Begin VB.Menu chkXBox 
            Caption         =   "X Box"
         End
      End
      Begin VB.Menu ChkOptions2 
         Caption         =   "Display as..."
         Begin VB.Menu Rect 
            Caption         =   "Rectangle"
         End
         Begin VB.Menu Sqre 
            Caption         =   "Square"
         End
         Begin VB.Menu Ellipse 
            Caption         =   "Oval"
         End
         Begin VB.Menu Circ 
            Caption         =   "Circle"
         End
         Begin VB.Menu RRect 
            Caption         =   "Rounded Rectangle"
         End
         Begin VB.Menu RSqr 
            Caption         =   "Rounded Square"
         End
      End
      Begin VB.Menu TxtBordToggle 
         Caption         =   "Border"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TopOfDesForm As Single
Dim ToolX As Single

'=========== Mouse Tracking ==============
Dim blnMouseIsDown As Boolean           'indicates if mouse button is pressed
Dim StartX As Single                    'mouse X-coord when button pressed
Dim StartY As Single                    'mouse Y-coord when button pressed
Dim dX1 As Single                       'saves distance from mouse to left of object (or line X1)
Dim dY1 As Single                       'saves distance from mouse to top of object (or line Y1)
Dim dX2 As Single                       'saves distance from mouse to line X2
Dim dY2 As Single                       'saves distance from mouse to line Y2
Dim blnAtLimit As Boolean

'============ Page Measurement Scale Arrays =============
Dim recVSclText(60) As Rect             'holds values for vertical scale text rectangles
Dim recHSclText(60) As Rect             'holds values for horizontal scale text rectangles

'===== variables for drag selection in horizontal/vertical scale bars ======
Dim ScaleStartX As Single           'horizontal scale start X coord
Dim ScaleStartY As Single           'vertical scale start Y coord
Dim ScaleEndX As Single             'horizontal scale end X coord
Dim ScaleEndY As Single             'vertical scale end Y coord
Dim ScaleLastX As Long              'previous X coord while dragging
Dim ScaleLastY As Long              'previous Y coord while dragging

Private Type DrawOrder  'type for set/get draw order for objects on the page
    ctlName As String
    ctlIndex As Integer
End Type

Dim DrawList() As DrawOrder     'array to keep track of back-to-front order of controls
Dim dLstCount As Integer        'sets/gets the current number of control objects

Dim blnFormLoad As Boolean          'whether form is being loaded
Dim blnCreatingControl As Boolean   'control is being created
Dim NewBound As Integer             'array dimension for number of controls created
Dim PgFreeWid As Single             'current page width showing on form
Dim blnDialogOpen As Boolean        'open file dialog was opened


Private Sub GetTextFormat()
'gets the format of the currently selected text control
Dim i As Integer

    If blnControlSelected Then
        If TypeOf ctlActive Is Label Then
            cboFontSize = Round(ctlActive.FontSize, 0)
            Toolbar2.Buttons(1).value = -ctlActive.FontBold
            Toolbar2.Buttons(2).value = -ctlActive.FontItalic
            Toolbar2.Buttons(3).value = -ctlActive.FontUnderline
            For i = 6 To 8
                If Val(Toolbar2.Buttons(i).Tag) = ctlActive.Alignment Then
                    Toolbar2.Buttons(i).value = tbrPressed
                Else
                    Toolbar2.Buttons(i).value = tbrUnpressed
                End If
            Next i
            FontSelector1.FontFace = ctlActive.FontName
        End If
    End If

End Sub

Private Sub AssignField_Click()
'If a regular data field has been clicked call SetDataField
    
    If ctlActive.LinkTimeout = cDataField Then
        SetDataField
    End If

End Sub

Private Sub SetDataField()
'opens frmSelField for assigning a data source to a field - if a valid data connection exists

    If frmSelField.lstFields.ListItems.count > 0 Then   'lstFields is only filled when data source has been established
        frmSelField.Left = (ctlActive.Left + picSection(ctlActive.Tag).Left + picContainer.Left) * 1440 _
            + Me.Left
        frmSelField.Top = (ctlActive.Top + ctlActive.Height + picSection(ctlActive.Tag).Top + picContainer.Top) * 1440 _
             + Me.Top + 800
        frmSelField.lstFields.Visible = True
        frmSelField.Caption = "Select Field from [" & strTableName & "]"
        frmSelField.Show vbModal
    Else
        If lngState <> MoveGrip Then        'don't want to trigger when field is just being resized
            MsgBox "You need to connect to a database and select a" & Chr$(13) & _
            "table or query to retrieve a list of fields to use." & Chr$(13) & Chr$(13) _
            & "Click on Connect to Database in the Tools menu.", _
            vbInformation, "No Fields Available"
        End If
    End If

End Sub

Private Sub BringFront_Click()
'brings the currently selected control to the front (on top of all other controls - drawn last)

Dim i As Integer
Dim SaveName As String, SaveIndex As Integer, SaveLoc As Integer
Dim Last As Long

    ctlActive.ZOrder (0)    'set the selected controls ZOrder to 0 (puts it on top)
    SaveName = ctlActive.Name       'save the selected controls name and control array index no.
    SaveIndex = ctlActive.Index

    Last = UBound(DrawList)

    For i = 0 To Last       'loop through draw order list and find the selected control
        If DrawList(i).ctlName = SaveName And DrawList(i).ctlIndex = SaveIndex Then
            SaveLoc = i     'if found, save the location and exit the loop
            Exit For
        End If
    Next i

    For i = 0 To Last           'loop through the draw order list and move all the controls after the selected control
        If i > SaveLoc Then     'up one so that selected control can be moved to the end
            DrawList(i - 1).ctlName = DrawList(i).ctlName
            DrawList(i - 1).ctlIndex = DrawList(i).ctlIndex
        End If
    Next i

    DrawList(Last).ctlName = SaveName   'add the selected control to the end of the draw order list
    DrawList(Last).ctlIndex = SaveIndex
    
End Sub

Private Sub cboFontSize_Click()
'when a font size is selected from the combo box, update the text control if selected

    ActiveText.FontSize = cboFontSize
    If blnControlSelected Then
        If TypeOf ctlActive Is Label Then
            ctlActive.FontSize = cboFontSize
        End If
    ElseIf blnGroupSelected Then
        Dim i As Integer
        For i = 0 To UBound(SelectedCtl)
            If TypeOf SelectedCtl(i).ctl Is Label Then
                SelectedCtl(i).ctl.FontSize = ActiveText.FontSize
            End If
        Next i
    End If

End Sub

Private Sub BoldClicked()
'when Bold is selected from the check button, update any text control selected
'and make it active for control placement

    WriteToUndoList unControl, unFormat
    
    ActiveText.IsBold = -(Toolbar2.Buttons(1).value)
    If blnControlSelected = True Then
        If TypeOf ctlActive Is Label Then
            ctlActive.FontBold = -(Toolbar2.Buttons(1).value)
        End If
    End If
        
    If blnGroupSelected Then
        For i = 0 To UBound(SelectedCtl)
            If TypeOf SelectedCtl(i).ctl Is Label Then
                SelectedCtl(i).ctl.FontBold = -(Toolbar2.Buttons(1).value)
            End If
        Next i
    End If

End Sub

Private Sub chk3d_Click()
'toggle whether check box displays sunken or not

    chk3d.Checked = Not chk3d.Checked
    ActiveChkBox.Sunken = chk3d.Checked
    ctlActive.Sunken = ActiveChkBox.Sunken

End Sub

Private Sub ChkBox_DblClick(Index As Integer)
'call SetDataField when check box is double-clicked

    SetDataField

End Sub

Private Sub ChkBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'handles left and right-click on a checkbox control in the report
'left-click selects it and prepares it for being moved
'right-click brings up a pop-up menu and enables the appropriate options

    CurrSection = Chkbox(Index).Tag
    blnMouseIsDown = True
    blnDragStarted = False
    
    If Button = 2 Then
        Set ctlActive = Chkbox(Index)
        ShowPopUp cCheckBox
        Exit Sub
    End If

    SetupControlSelection Chkbox(Index), Chkbox(Index).Tag, 96, x, y

End Sub

Private Sub ChkBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'handles dragging of a checkbox control in the report - can only be dragged within a section

    MoveInControl x, y, Chkbox(Index).ToolTipText

End Sub

Private Sub ChkBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'when mouse button is released on a checkbox control in the report
'sets new position for control and maintains current selection

    ConfirmControlSelection Chkbox(Index), Chkbox(Index).Tag
    
End Sub

Private Sub chkCheck_Click()
'change checkbox display to check mark and make active

    ActiveChkBox.DisplayType = 0
    ctlActive.DisplayType = 0

End Sub

Private Sub ItalicClicked()
'when Italic is selected from the check button, update any text control selected
'and make it active for control placement

    WriteToUndoList unControl, unFormat
    
    ActiveText.IsItalic = -(Toolbar2.Buttons(2).value)
    If blnControlSelected = True Then
        If TypeOf ctlActive Is Label Then
            ctlActive.FontItalic = -(Toolbar2.Buttons(2).value)
        End If
    End If
        
    If blnGroupSelected Then
        For i = 0 To UBound(SelectedCtl)
            If TypeOf SelectedCtl(i).ctl Is Label Then
                SelectedCtl(i).ctl.FontItalic = -(Toolbar2.Buttons(2).value)
            End If
        Next i
    End If


End Sub

Private Sub chkRadio_Click()
'change checkbox display to radio button and make active

    ActiveChkBox.DisplayType = 1
    ctlActive.DisplayType = 1

End Sub

Private Sub UnderlineClicked()
'when Underline is selected from the check button, update any text control selected
'and make it active for control placement

    WriteToUndoList unControl, unFormat
    
    ActiveText.IsUnderline = -(Toolbar2.Buttons(3).value)
    If blnControlSelected = True Then
        If TypeOf ctlActive Is Label Then
            ctlActive.FontUnderline = -(Toolbar2.Buttons(3).value)
        End If
    End If
        
    If blnGroupSelected Then
        For i = 0 To UBound(SelectedCtl)
            If TypeOf SelectedCtl(i).ctl Is Label Then
                SelectedCtl(i).ctl.FontUnderline = -(Toolbar2.Buttons(3).value)
            End If
        Next i
    End If

End Sub

Private Sub chkXBox_Click()
'change checkbox display to x-box and make active

    ActiveChkBox.DisplayType = 2
    ctlActive.DisplayType = 2

End Sub

Private Sub Circ_Click()
'change shape to display circle
    
    WriteToUndoList unControl, unFormat
    ActiveShape = 3
    UpdateShape

End Sub

Private Sub UpdateShape()
'when user has changed shape display type, update the display
Dim i As Integer

    If blnControlSelected Then
        ctlActive.Shape = ActiveShape
    ElseIf blnGroupSelected Then
        For i = 0 To UBound(SelectedCtl)
            If TypeOf SelectedCtl(i).ctl Is Shape Then
                SelectedCtl(i).ctl.Shape = ActiveShape
            End If
        Next i
    End If

End Sub

Private Sub cmdDivider_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'sets up page section divider for being moved

    ClearSelection
    PropertySelector1.Visible = False
    Me.lnPos.Y1 = cmdDivider(Index).Top - 0.17
    Me.lnPos.Y2 = Me.lnPos.Y1
    Me.lnPos.Visible = True
    SetDividerColours Index
    ShowVertScale
    CurrSection = Index
    
    blnMouseIsDown = True
    ButtonTop = y

End Sub

Private Sub SetDividerColours(GetIndex As Integer)
'update the section divider highlight color when user clicks on them or sections
Dim i As Integer

    For i = FirstSectionVis To LastSectionVis
        If i = GetIndex Then
            cmdDivider(i).BackColor = &H80000013
        Else
            If (i > 1 And i < 5) Or (i > 5 And i < 9) Then
                cmdDivider(i).BackColor = &HFFC0C0
            Else
                cmdDivider(i).BackColor = &H8000000F
            End If
        End If
    Next i
    
End Sub

Private Sub cmdDivider_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'moves the page section divider with the mouse
Dim SectionY As Single
Dim i As Integer

    If Index > FirstSectionVis Then
        Me.MousePointer = 7
        If blnMouseIsDown Then
            lngState = ResizeSection
            For i = (Index - 1) To FirstSectionVis Step -1
                If picSection(i).Visible Then
                    Exit For
                End If
            Next i
            SectionY = picSection(i).Height + (y - ButtonTop) / 1440
            ExecuteState i, lngState, x, SectionY
        End If
    End If

End Sub

Private Sub cmdDivider_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    blnMouseIsDown = False
    blnDragStarted = False
    Me.lnPos.Visible = False
    If Index > FirstSectionVis Then
        If lngState = ResizeSection Then
            ShowGrid
            SetPageComponents
            ShowVertScale
        End If
    End If

End Sub

Private Sub GetFirstLastSections()
'set variables for which are the first and last sections currently displayed
Dim i As Integer

    For i = 0 To 5
        If cmdDivider(i).Visible Then
            FirstSectionVis = i
            Exit For
        End If
    Next i
    For i = 10 To 5 Step -1
        If cmdDivider(i).Visible = True Then
            LastSectionVis = i
            Exit For
        End If
    Next i

End Sub

Public Sub SetPageComponents()
'sets the dimensions and display of the page sections, dividers and scroll bars
Dim GetHeight As Single
Dim i As Integer

    If TotalPageHeight() < picContainer.Height Then
        cmdDivider(FirstSectionVis).Top = 0
        StartScaleTop = 0
        VScroll.value = 0
    Else
        cmdDivider(FirstSectionVis).Top = -VScroll / 8
    End If
    GetHeight = cmdDivider(FirstSectionVis).Top + cmdDivider(FirstSectionVis).Height
    picSection(FirstSectionVis).Top = GetHeight
    
    If picSection(FirstSectionVis).width < picContainer.width Then
        For i = 0 To 10
            picSection(i).Left = 0.17
        Next i
        HScroll.value = 0
    End If
    
    For i = (FirstSectionVis + 1) To LastSectionVis
        If blnFormLoad Then
            GetHeight = GetHeight + picSection(i - 1).Height
            cmdDivider(i).Top = GetHeight
            GetHeight = GetHeight + cmdDivider(i).Height
            picSection(i).Top = GetHeight
        Else
            If picSection(i - 1).Visible = True Then
                GetHeight = GetHeight + picSection(i - 1).Height
                cmdDivider(i).Top = GetHeight
                GetHeight = GetHeight + cmdDivider(i).Height
                picSection(i).Top = GetHeight
            Else
                cmdDivider(i).Top = GetHeight - cmdDivider(i).Height
                picSection(i).Top = GetHeight
            End If
        End If
    Next i

End Sub

Private Sub OpenNewFile()
'clears off current display and loads blank new file
Dim i As Integer

'prompt user if currently open file need saving
    If Not blnReportSaved Then
        If MsgBox("Do you want to save changes to " & OpenFileName & "?", vbYesNo, "Confirm Save Changes...") = vbYes Then
            SaveReportDesign False
        End If
    End If
    
'unload any currently loaded controls
    For i = Me.Controls.count - 1 To 0 Step -1
        Set ctlTest = Me.Controls(i)
        If ctlTest.Tag > "" Then
            Unload Me.Controls(i)
        End If
    Next i
        
    ClearSelection
    
'load default page values
    PageSizeName = "Letter"
    PageScaleUnits = scEnglish
    PageWd = 8.5
    PageHt = 11
    LeftMarg = 0.5
    RightMarg = 0.5
    TopMarg = 0.5
    BottomMarg = 0.5
    PageOrient = cPortrait

'reset page section sizes
    For i = 0 To 10
        picSection(i).Height = 0.5
        picSection(i).width = PageWd - LeftMarg - RightMarg
        picSection(i).BackColor = vbWhite
        ResetMinSectionHt i
    Next i
    
    blnControlSelected = False
    blnGroupSelected = False
    
    'hide group headers and footers if they are visible
    For i = 2 To 4
        cmdDivider(i).Visible = False
        picSection(i).Visible = False
        picSection(i).Height = 0
        cmdDivider(10 - i).Visible = False
        picSection(10 - i).Visible = False
        picSection(10 - i).Height = 0
    Next i
    
'run various setup subs also used at program startup
    ResetMinPageWidth
    GetFirstLastSections
    SetPageComponents
    InitUndoArrays
    SetupDefaultSettings
    
    blnGridOn = True
    ShowGrid
    ShowHorizScale
    ShowVertScale
    
'reset scroll bars
    HScroll.Visible = (picSection(0).width > picContainer.width)
    VScroll.Visible = (TotalPageHeight() > picContainer.Height)
    
'clear off global variables
    blnReportDataBound = False
    blnHasPics = False
    strDataFileName = ""
    strTableName = ""
    strConnErrMsg = ""
    strImgPathTable = ""
    strImgPathField = ""
    strImageFolder = ""

'zero off the global control counters
    dLstCount = 0
    ReDim DrawList(0)
    LineIndex = 0
    ShapeIndex = 0
    LabelIndex = 0
    FieldIndex = 0
    ImageIndex = 0
    BoundImageIndex = 0
    
    For i = 0 To 2
        GroupHVis(i) = False
        GroupFVis(i) = False
        blnNewPageOnHeader(i) = False
        strSortField(i) = ""
        blnSortDescending(i) = False
    Next i
    
'prompt user if new report will use a data connection
    If MsgBox("Will this new report require a data connection?", vbYesNo + vbQuestion, "Data Connection Required") _
    = vbYes Then
        frmDataConn.Show
    End If
    

    Me.Caption = "RG Report Designer - New File"
    blnReportSaved = True

End Sub

Private Sub OpenSavedFile()
'clears off current display and prompts for, then opens a saved report file
Dim i As Integer, j As Integer
Dim GetNum As Integer
Dim FLen As Long
Dim strLocalFonts() As String
Dim strUseFont As String

'poll system for a list of locally available fonts - will need this to confirm
'that fonts specified in saved report file are on this computer
    ReDim strLocalFonts(Screen.FontCount)
    For i = 0 To Screen.FontCount - 1
        strLocalFonts(i) = Screen.Fonts(i)
    Next i

'prompt user if current report file needs saving
    If Not blnReportSaved Then
        If MsgBox("Do you wish to save this report?", vbYesNo + vbQuestion, "Save Report") = vbYes Then
            SaveReportDesign False
        End If
    End If
    
'show open file dialog
    On Error GoTo ExitSub
    blnDialogOpen = True
    CommonDialog1.filter = "Report Files (*.drp)|*.drp"
    CommonDialog1.ShowOpen
    OpenFileName = CommonDialog1.FileName
    
'unload any currently loaded controls
    For i = Me.Controls.count - 1 To 0 Step -1
        Set ctlTest = Me.Controls(i)
        If ctlTest.Tag > "" Then
            Unload Me.Controls(i)
        End If
    Next i
    ClearSelection
    
'load the saved report file into the ReportFile object type
    On Error GoTo NoOpen
    Open OpenFileName For Binary Access Read Lock Write As #3
    Get 3, , GetNum
    ReDim ReportFile.RpControl(GetNum)
    Get 3, , ReportFile
    Close 3
    
'zero off the global control counters
    dLstCount = 0
    ReDim DrawList(0)
    LineIndex = 0
    ShapeIndex = 0
    LabelIndex = 0
    FieldIndex = 0
    ImageIndex = 0
    BoundImageIndex = 0
        
'extract the report and control information and load it into the display
    With ReportFile
        strDataFileName = .DBName
        strTableName = .DBSource
        strImgPathTable = .ImgPathTable
        strImgPathField = .ImgPathField
        strImageFolder = .ImageFolder
        For i = 0 To 2
            strSortField(i) = .SortField(i)
            blnSortDescending(i) = .SortDescending(i)
            blnNewPageOnHeader(i) = .NewPageOnHeader(i)
        Next i
        PageSizeName = .PageSzNam
        PageScaleUnits = .PageSclUnit
        PageWd = .PageWd
        PageHt = .PageHt
        PageOrient = .Orient
        LeftMarg = .LMarg
        RightMarg = .RMarg
        TopMarg = .TMarg
        BottomMarg = .BMarg
       
        For i = 0 To 10
            picSection(i).width = .DesWd
            picSection(i).BackColor = .SectColor(i)
        Next i
        
        For i = 0 To 4
            cmdDivider(i).Visible = .HeaderVis(i)
            picSection(i).Visible = .HeaderVis(i)
            cmdDivider(10 - i).Visible = .FooterVis(i)
            picSection(10 - i).Visible = .FooterVis(i)
            picSection(i).Height = .HeaderHt(i)
            picSection(10 - i).Height = .FooterHt(i)
        Next i
        picSection(5).Height = .DetHt
        RepHdFt.Checked = -(picSection(0).Visible)
        Me.PageHdFt.Checked = -(picSection(1).Visible)
        
        For i = 0 To 2
            GroupHVis(i) = .HeaderVis(i + 2)
            GroupFVis(i) = .FooterVis(i + 2)
        Next i
        
        For i = 2 To 4
            cmdDivider(i).Caption = IIf(.HeaderVis(i), "'" & .SortField(i - 2) & "' Group Header", "")
            cmdDivider(10 - i).Caption = IIf(.FooterVis(i), "'" & .SortField(i - 2) & "' Group Footer", "")
        Next i
        
        For i = 0 To 10
            picSection(i).BackColor = .SectColor(i)
        Next i
       
        If UBound(.RpControl) > 0 Then
            For i = 1 To UBound(.RpControl)
                If .RpControl(i).Type = cLine Then
                    LineIndex = LineIndex + 1
                    Load Lin(LineIndex)
                    Set Lin(LineIndex).Container = picSection(.RpControl(i).SecNo)
                    Lin(LineIndex).ZOrder (0)
                    Lin(LineIndex).Tag = .RpControl(i).SecNo
                    Lin(LineIndex).X1 = .RpControl(i).X1
                    Lin(LineIndex).Y1 = .RpControl(i).Y1
                    Lin(LineIndex).X2 = .RpControl(i).X2
                    Lin(LineIndex).Y2 = .RpControl(i).Y2
                    Lin(LineIndex).BorderColor = .RpControl(i).BdrClr
                    Lin(LineIndex).BorderStyle = .RpControl(i).BdrStl
                    Lin(LineIndex).BorderWidth = .RpControl(i).BdrWd
                    Lin(LineIndex).Visible = True
                    Set ctlActive = Lin(LineIndex)
                ElseIf .RpControl(i).Type = cBox Then
                    ShapeIndex = ShapeIndex + 1
                    Load Shape(ShapeIndex)
                    Set Shape(ShapeIndex).Container = picSection(.RpControl(i).SecNo)
                    Shape(ShapeIndex).ZOrder (0)
                    Shape(ShapeIndex).Tag = .RpControl(i).SecNo
                    Shape(ShapeIndex).Left = .RpControl(i).Left
                    Shape(ShapeIndex).Top = .RpControl(i).Top
                    Shape(ShapeIndex).width = .RpControl(i).width
                    Shape(ShapeIndex).Height = .RpControl(i).Height
                    Shape(ShapeIndex).BorderColor = .RpControl(i).BdrClr
                    Shape(ShapeIndex).BorderStyle = .RpControl(i).BdrStl
                    Shape(ShapeIndex).BorderWidth = .RpControl(i).BdrWd
                    Shape(ShapeIndex).BackColor = .RpControl(i).BckClr
                    Shape(ShapeIndex).BackStyle = .RpControl(i).BckStl
                    Shape(ShapeIndex).Shape = .RpControl(i).DisplayType
                    Shape(ShapeIndex).Visible = True
                    Set ctlActive = Shape(ShapeIndex)
                ElseIf .RpControl(i).Type = cLabel Then
                    LabelIndex = LabelIndex + 1
                    Load Label(LabelIndex)
                    Set Label(LabelIndex).Container = picSection(.RpControl(i).SecNo)
                    Label(LabelIndex).ZOrder (0)
                    Label(LabelIndex).Tag = .RpControl(i).SecNo
                    Label(LabelIndex).Left = .RpControl(i).Left
                    Label(LabelIndex).Top = .RpControl(i).Top
                    Label(LabelIndex).width = .RpControl(i).width
                    Label(LabelIndex).Height = .RpControl(i).Height
                    Label(LabelIndex).BackStyle = .RpControl(i).BckStl
                    Label(LabelIndex).BackColor = .RpControl(i).BckClr
                    Label(LabelIndex).ForeColor = .RpControl(i).ForClr
                    Label(LabelIndex).BorderStyle = .RpControl(i).BdrStl
                    strUseFont = "Arial"  'use as default font if font name not found on this computer
                    For j = 0 To UBound(strLocalFonts)  'search for font name
                        If .RpControl(i).FntNam = strLocalFonts(j) Then
                            strUseFont = .RpControl(i).FntNam
                            Exit For
                        End If
                    Next j
                    Label(LabelIndex).FontName = strUseFont
                    Label(LabelIndex).FontSize = .RpControl(i).FntSiz
                    Label(LabelIndex).FontBold = .RpControl(i).FntBld
                    Label(LabelIndex).FontItalic = .RpControl(i).FntItl
                    Label(LabelIndex).FontUnderline = .RpControl(i).FntUnd
                    Label(LabelIndex).Alignment = .RpControl(i).Align
                    Label(LabelIndex).Caption = .RpControl(i).strText
                    Label(LabelIndex).Visible = True
                    Set ctlActive = Label(LabelIndex)
                ElseIf .RpControl(i).Type = cDataField Or .RpControl(i).Type = cDatePageField Or _
                .RpControl(i).Type = cCalcField Or .RpControl(i).Type = cSumField Then
                    FieldIndex = FieldIndex + 1
                    Load Field(FieldIndex)
                    Set Field(FieldIndex).Container = picSection(.RpControl(i).SecNo)
                    Field(FieldIndex).ZOrder (0)
                    Field(FieldIndex).Tag = .RpControl(i).SecNo
                    Field(FieldIndex).Left = .RpControl(i).Left
                    Field(FieldIndex).Top = .RpControl(i).Top
                    Field(FieldIndex).width = .RpControl(i).width
                    Field(FieldIndex).Height = .RpControl(i).Height
                    Field(FieldIndex).LinkTimeout = .RpControl(i).Type
                    Field(FieldIndex).BackColor = .RpControl(i).BckClr
                    Field(FieldIndex).ForeColor = .RpControl(i).ForClr
                    Field(FieldIndex).BorderStyle = .RpControl(i).BdrStl
                    strUseFont = "Arial"
                    For j = 0 To UBound(strLocalFonts)
                        If .RpControl(i).FntNam = strLocalFonts(j) Then
                            strUseFont = .RpControl(i).FntNam
                            Exit For
                        End If
                    Next j
                    Field(FieldIndex).FontName = strUseFont
                    Field(FieldIndex).FontSize = .RpControl(i).FntSiz
                    Field(FieldIndex).FontBold = .RpControl(i).FntBld
                    Field(FieldIndex).FontItalic = .RpControl(i).FntItl
                    Field(FieldIndex).FontUnderline = .RpControl(i).FntUnd
                    Field(FieldIndex).Alignment = .RpControl(i).Align
                    Field(FieldIndex).Caption = .RpControl(i).strText
                    Field(FieldIndex).ToolTipText = .RpControl(i).strText
                    Field(FieldIndex).DataField = .RpControl(i).Fieldname
                    Field(FieldIndex).Visible = True
                    Set ctlActive = Field(FieldIndex)
                ElseIf .RpControl(i).Type = cImage Then
                    ImageIndex = ImageIndex + 1
                    Load Img(ImageIndex)
                    Set Img(ImageIndex).Container = picSection(.RpControl(i).SecNo)
                    Img(ImageIndex).Tag = .RpControl(i).SecNo
                    Img(ImageIndex).Left = .RpControl(i).Left
                    Img(ImageIndex).Top = .RpControl(i).Top
                    Img(ImageIndex).width = .RpControl(i).width
                    Img(ImageIndex).Height = .RpControl(i).Height
                    Img(ImageIndex).BackColor = .RpControl(i).BckClr
                    Img(ImageIndex).BackStyle = .RpControl(i).BckStl
                    Img(ImageIndex).BorderColor = .RpControl(i).BdrClr
                    Img(ImageIndex).BorderStyle = .RpControl(i).BdrStl
                    Open App.Path & "\tmpfile" For Binary As #4
                    Put 4, , .RpControl(i).ImgData
                    Close 4
                    Img(ImageIndex).Picture = LoadPicture(App.Path & "\tmpfile")
                    Kill App.Path & "\tmpfile"
                    DoEvents
                    Img(ImageIndex).Visible = True
                    Set ctlActive = Img(ImageIndex)
                ElseIf .RpControl(i).Type = cBoundImage Then
                    BoundImageIndex = BoundImageIndex + 1
                    Load BoundImg(BoundImageIndex)
                    Set BoundImg(BoundImageIndex).Container = picSection(.RpControl(i).SecNo)
                    BoundImg(BoundImageIndex).Tag = .RpControl(i).SecNo
                    BoundImg(BoundImageIndex).Left = .RpControl(i).Left
                    BoundImg(BoundImageIndex).Top = .RpControl(i).Top
                    BoundImg(BoundImageIndex).width = .RpControl(i).width
                    BoundImg(BoundImageIndex).Height = .RpControl(i).Height
                    BoundImg(BoundImageIndex).Caption = .RpControl(i).strText
                    BoundImg(BoundImageIndex).LinkTimeout = .RpControl(i).Type
                    BoundImg(BoundImageIndex).ToolTipText = .RpControl(i).strText
                    BoundImg(BoundImageIndex).DataField = .RpControl(i).Fieldname
                    BoundImg(BoundImageIndex).Visible = True
                    Set ctlActive = BoundImg(BoundImageIndex)
                ElseIf .RpControl(i).Type = cCheckBox Then
                    CheckShapeIndex = CheckShapeIndex + 1
                    Load Chkbox(CheckShapeIndex)
                    Set Chkbox(CheckShapeIndex).Container = picSection(.RpControl(i).SecNo)
                    Chkbox(CheckShapeIndex).Tag = .RpControl(i).SecNo
                    Chkbox(CheckShapeIndex).Left = .RpControl(i).Left
                    Chkbox(CheckShapeIndex).Top = .RpControl(i).Top
                    Chkbox(CheckShapeIndex).width = .RpControl(i).width
                    Chkbox(CheckShapeIndex).Height = .RpControl(i).Height
                    Chkbox(CheckShapeIndex).BorderColor = .RpControl(i).BdrClr
                    Chkbox(CheckShapeIndex).BackColor = .RpControl(i).BckClr
                    Chkbox(CheckShapeIndex).DataField = .RpControl(i).Fieldname
                    Chkbox(CheckShapeIndex).ToolTipText = "=![" & .RpControl(i).Fieldname & "]"
                    Chkbox(CheckShapeIndex).DisplayType = .RpControl(i).DisplayType
                    Chkbox(CheckShapeIndex).Sunken = .RpControl(i).Sunken
                    Chkbox(CheckShapeIndex).Visible = True
                    Set ctlActive = Chkbox(CheckShapeIndex)
                End If
                DoEvents
                AddToDrawList
            Next i
        End If
    End With
    
    GetFirstLastSections
    
    ResetMinPageWidth
    For i = FirstSectionVis To LastSectionVis
        ResetMinSectionHt i
    Next i
        
    SetPageComponents
    VScroll.Max = TotalPageHeight() * 8
    VScroll.Visible = (TotalPageHeight() > picContainer.Height)
    HScroll.Visible = (picSection(0).width + 0.2 >= picContainer.width)
    ShowVertScale
    ShowHorizScale
    InitUndoArrays
    
    If blnGridOn Then ShowGrid
    
'if a data connection has been detected then load all the connection and data parameter info
    blnReportDataBound = False
    If ReportFile.DataBound Then
        If strDataFileName > "" Then
            If Dir(strDataFileName, vbNormal) > "" Then
                If ConnectToDataFile Then
                    Dim strSort As String
                'build sort portion of SQL statement
                    strSort = ""
                    For i = 0 To 2
                        If strSortField(i) > "" Then
                            If strSort > "" Then strSort = strSort & ", "
                            strSort = strSort & strSortField(i)
                            If blnSortDescending(i) = True Then strSort = strSort & " DESC"
                        End If
                    Next i
                    If strSort <> "" Then strSort = "ORDER BY " & strSort
                    
                    GetTables
                    OpenData strTableName, strSort
                    LoadFieldNames
                    blnReportDataBound = True
                    StatusBar1.Panels(2).text = "Data Source : [" & strTableName & "] from " & strDataFileName
                    Toolbar1.Buttons(16).Enabled = True
                Else
                    strConnErrMsg = "Could not connect to " & strDataFileName
                    MsgBox strConnErrMsg, vbOKOnly
                    frmDataConn.Show vbModal
                End If
            Else
                strConnErrMsg = "Could not find " & strDataFileName
                MsgBox strConnErrMsg
                frmDataConn.Show vbModal
                Exit Sub
            End If
        Else
            strConnErrMsg = "No database file specified"
            MsgBox strConnErrMsg
        End If
    End If
    
    Me.Caption = "RG Report Designer - " & OpenFileName
    blnReportSaved = True
    
    Exit Sub

NoOpen:
    MsgBox "Error opening file : " & Err.Description

ExitSub:

End Sub

Private Sub cmdPgSetup_Click()

    frmPageSetup.Show

End Sub

Private Sub PreviewTheReport()
'runs the Print Preview of the report
'the mPreviewData module is fairly self-contained and I tried to write it so that it
'can be used independently of this project for previewing report files
'there is an optional file name parameter for PreviewReport to specify an external
'report file to be loaded and previewed
'(would still require a reference to PrintPreview.dll)

On Error GoTo NoPreview

    If ReportCanPrint Then
        LoadCurrentReport
        PreviewReport True
    End If
    
    Exit Sub

NoPreview:
    MsgBox Err.Description

End Sub

Private Function ReportCanPrint() As Boolean
Dim i As Integer
Dim ObjCount As Integer

    ReportCanPrint = False

'get number of report controls and test if there are any to preview or print
    ObjCount = 0
    For i = 0 To Me.Controls.count - 1
        If Me.Controls(i).Tag > "" Then
            ObjCount = ObjCount + 1
        End If
    Next i
    If ObjCount = 0 Then
        MsgBox "No report has been created yet. There is nothing to preview.", vbOKOnly + vbInformation
        Exit Function
    End If

'check first to see if user has connected to data file
    If blnReportDataBound Then
        If strDataFileName = "" Then
            If MsgBox("You  need to connect to a database in order to preview this report" & Chr$(13) & _
            "Do you wish to connect now?", vbYesNo, "Confirm Connect to Data") = vbNo Then
                Exit Function
            Else
                frmDataConn.Show
                Exit Function
            End If
        End If
    End If
    
    ReportCanPrint = True

End Function

Private Sub PrintTheReport()
'sends report directly to printer with no print preview
'NOTE - I have not tested this very much as paper and ink cartridges are expensive

    If ReportCanPrint Then
        LoadCurrentReport
        PreviewReport False
        Exit Sub
    End If

End Sub

Private Sub SaveReportDesign(SaveNewName As Boolean)
'save the current report file to the specified location

Dim i As Integer
Dim j As Integer
Dim FileNum As Long
Dim SaveFileName As String

    If Not SaveNewName Then
        If blnReportSaved Then Exit Sub
    End If
    
'open the Save File dialog
    SaveFileName = ""
    On Error GoTo ExitSub
    CommonDialog1.filter = "Report Files (*.drp)|*.drp"
'if it's a new file being saved
    If OpenFileName = "New File" Then
        CommonDialog1.ShowSave
        SaveFileName = CommonDialog1.FileName
        Me.Caption = "RG Report Designer - " & SaveFileName
'if it's an existing file being saved
    ElseIf SaveNewName Then
        CommonDialog1.FileName = Left$(OpenFileName, Len(OpenFileName) - 4)
        CommonDialog1.ShowSave
        SaveFileName = CommonDialog1.FileName
        Me.Caption = "RG Report Designer - " & OpenFileName
    Else
        SaveFileName = OpenFileName
    End If
    
'load report file data into ReportFile object
    LoadCurrentReport
'check if file already exists, prompt to overwrite, delete existing file if needed
    If SaveFileName > "" Then
        If SaveFileName <> OpenFileName Then
            If Dir(SaveFileName) > "" Then
                If MsgBox(SaveFileName & " already exists.  Do you wish to overwrite it?", vbYesNo) = vbYes Then
                    Kill SaveFileName
                Else
                    Exit Sub
                End If
            End If
        Else
            Kill SaveFileName
        End If
    Else
        Exit Sub
    End If
    DoEvents
'open a file, write to it and close it
    FileNum = FreeFile
    On Error GoTo NoSave
    Open SaveFileName For Binary Access Write Lock Write As FileNum
    
    Put FileNum, , NewBound
    Put FileNum, , ReportFile
    Close FileNum
    
    blnReportSaved = True
    
    Exit Sub
    
NoSave:
    MsgBox "Error opening '" & OpenFileName & " : " & Err.Description

ExitSub:

End Sub

Private Sub ShowFillClrClicked()
'user has clicked the active fill color drop-down button, show the color selector

    With PropertySelector1
        If picFillCol.FillStyle = 1 Then
            .SelectedColor = -1
        Else
            .SelectedColor = shpFill.FillColor
        End If
        .TransparentOption = True
        SetPropertyPanelPos
        If PropertySelectMode <> 3 Then
            .SelectMode = 3
            .Visible = True
        Else
            .Visible = Not .Visible
        End If
    End With

    PropertySelectMode = 3

End Sub

Private Sub ShowLineClrClicked()
'user has clicked the active line color drop-down button, show the color selector

    With PropertySelector1
        .TransparentOption = True
        SetPropertyPanelPos
        If PropertySelectMode <> 2 Then
            .SelectMode = 2
            .Visible = True
        Else
            .Visible = Not .Visible
        End If
    End With

    PropertySelectMode = 2

End Sub

Private Sub ShowTextClrClicked()
'user has clicked the active text color drop-down button, show the color selector

    With PropertySelector1
        .SelectedColor = picTxtCol.BackColor
        .TransparentOption = False
        SetPropertyPanelPos
        If PropertySelectMode <> 1 Then
            .SelectMode = 1
            .Visible = True
        Else
            .Visible = Not .Visible
        End If
    End With

    PropertySelectMode = 1

End Sub

Private Sub CopyClicked()
'user has clicked Copy in the popup context menu for later pasting
'control is really just flagged in the selection list - sorry, it's not actually on
'the Windows clipboard!
Dim i As Integer

    CopyActive = True
    CutActive = False
    Toolbar1.Buttons(10).Enabled = True
    PasteObj.Enabled = True
    blnMouseIsDown = False
    If blnGroupSelected Then
        For i = 0 To UBound(SelectedCtl)
            SelectedCtl(i).blnOnClipBrd = True
        Next i
        FromSection = CurrSection
    ElseIf blnControlSelected Then
        ReDim SelectedCtl(0)
        Set SelectedCtl(0).ctl = ctlActive
        SelectedCtl(0).blnOnClipBrd = True
        FromSection = ctlActive.Tag
    End If
    
End Sub

Private Sub CutClicked()
'user has clicked Cut in the popup context menu for later pasting
'control is really just hidden and flagged in the selection list - sorry, it's not actually on
'the Windows clipboard!
Dim i As Integer
    
    CopyActive = True
    Toolbar1.Buttons(10).Enabled = True
    PasteObj.Enabled = True
    blnMouseIsDown = False
    ClearSelection
    If blnGroupSelected Then
        For i = 0 To UBound(SelectedCtl)
            SelectedCtl(i).ctl.Visible = False
            SelectedCtl(i).blnOnClipBrd = True
        Next i
        FromSection = CurrSection
    ElseIf blnControlSelected Then
        ctlActive.Visible = False
        ReDim SelectedCtl(0)
        Set SelectedCtl(0).ctl = ctlActive
        SelectedCtl(0).blnOnClipBrd = True
        FromSection = ctlActive.Tag
    End If
    
    ResetMinSectionHt FromSection
    ResetMinPageWidth

End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
'user has dragged one of the toolbars vertically, changing the height of the coolbar
'adjust the top of form's controls accordingly

    TopOfDesForm = NewHeight
    SetupDesignScreen

End Sub

Private Sub CopyObj_Click()

    CopyClicked

End Sub

Private Sub CutObj_Click()

    CutClicked

End Sub

Private Sub DataConn_Click()

    frmDataConn.Show vbModal

End Sub

Private Sub DelObj_Click()

    DeleteSelectedCtlObject

End Sub

Private Sub EditCalc_Click()
'user clicked Edit Calculation in the context menu, show the dialog or prompt
'if no data connection

    If blnReportDataBound Then
        If strDataFileName <> "" Then
            strCalcDataFieldContents = ctlActive.ToolTipText
            blnEditExisting = True
            frmCalcDataField.Show vbModal
        End If
    Else
        MsgBox "There is no open database connection, no fields available to calculate", vbOKOnly, "No Connection Found"
        lngState = Default
        lngPrevState = Default
    End If

End Sub

Private Sub EditDatePage_Click()
'user clicked Edit Date/Page field in the context menu, show the dialog

    blnEditExisting = True
    frmDatePageField.Show

End Sub

Private Sub EditSummary_Click()
'user clicked Edit Summary field in the context menu, show the dialog

    blnEditExisting = True
    frmAggField.Show

End Sub

Private Sub EditText_Click()
'user has double-clicked on label or selected Edit Text from the context menu
'Uses a text box overlayed on the label for entering the new text, transferred to
'the label when the user hits Enter

    ctlActive.Visible = False
    With Me.txtInput
        Set .Container = picSection(CurrSection)
        .Left = ctlActive.Left
        .Top = ctlActive.Top
        .width = ctlActive.width
        .Height = ctlActive.Height
        .FontName = ctlActive.FontName
        .FontSize = ctlActive.FontSize
        .Alignment = ctlActive.Alignment
        .BackColor = ctlActive.BackColor
        .ForeColor = ctlActive.ForeColor
        .FontBold = ctlActive.FontBold
        .FontItalic = ctlActive.FontItalic
        .FontUnderline = ctlActive.FontUnderline
        .text = ctlActive.Caption
        .ZOrder (0)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.text)
        .SetFocus
    End With
    Me.StatusBar1.Panels(3).text = "Press 'Enter' to accept"

    Me.MousePointer = 3

End Sub

Private Sub Ellipse_Click()
'user has clicked on Ellipse display option in context menu, update shape

    WriteToUndoList unControl, unFormat
    ActiveShape = 2
    UpdateShape

End Sub

Private Sub ExitProg_Click()
    
    Unload Me

End Sub

Private Sub FontSelector1_FontSelected()
'user has clicked on Font Selector, set active font and update any selected controls

    ActiveText.FontName = FontSelector1.FontFace
    If blnControlSelected Then
        If TypeOf ctlActive Is Label Then
            ctlActive.FontName = ActiveText.FontName
        End If
    ElseIf blnGroupSelected Then
        Dim i As Integer
        For i = 0 To UBound(SelectedCtl)
            If TypeOf SelectedCtl(i).ctl Is Label Then
                SelectedCtl(i).ctl.FontName = ActiveText.FontName
            End If
        Next i
    End If

End Sub

Private Sub Form_Activate()

    If blnPageChanged Then
        Form_Resize
        blnPageChanged = False
    End If

End Sub

Private Sub Form_Click()

    PropertySelector1.Visible = False

End Sub

Private Sub Form_GotFocus()

    If blnGridOn Then ShowGrid

End Sub

Private Sub SetupDefaultSettings()
'sets up default active settings for controls when form or new file loads

    With ActiveBorder
        .Color = vbBlack
        .Style = 0
        .width = 1
    End With
    
    With ActiveBack
        .Style = 0
        .Color = vbWhite
    End With
    
    With ActiveText
        .Align = vbLeftJustify
        .FontName = "Arial"
        .FontSize = 10
        .IsBold = False
        .IsItalic = False
        .IsUnderline = False
        .Color = vbBlack
        .BorderOn = False
    End With
    
    ActiveChkBox.DisplayType = 0
    ActiveChkBox.Sunken = False
    
    SelectedColor = vbBlack

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'reset program state to control selection mode (default) when user hits escape key

    If KeyCode = vbKeyEscape Then
        lngState = Default
        lngCmdState = Default
        optAction(0).value = True
        ShowStateMouse
    ElseIf KeyCode = vbKeyDelete Then
        DeleteSelectedCtlObject
    End If

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim AddItem As ComboItem
    
    
    TopOfDesForm = CoolBar1.Height
    blnFormLoad = True
    blnReportSaved = True
    UnitScaleMod = 1
    dLstCount = 0
    RepHdrFtrVis = True
    PagHdrFtrVis = True
    optAction(0).value = True
    NumInGrp = -1
    ReDim DrawList(0)
    blnGridOn = True
    GridSpace = 3
    lngState = Default
    lngPrevState = lngState

    SetupDefaultSettings
    PropertySelector1.SelectedColor = 0
    
    StatusBar1.Panels(2).text = "Data Source : (None)"
    
    OpenFileName = "New File"
    Me.Caption = "RG Report Designer - New File"

    'set page section sizes
    For i = 0 To 10
        cmdDivider(i).Left = 0
        picSection(i).Left = 0.17
        picSection(i).width = 7.5
        picSection(i).Height = 0.375
        picSection(i).BackColor = vbWhite
    Next i
    
    InitUndoArrays
    
    'load default page values
    PageSizeName = "Letter"
    PageWd = 8.5
    PageHt = 11
    LeftMarg = 0.5
    RightMarg = 0.5
    TopMarg = 0.5
    BottomMarg = 0.5
    PageOrient = cPortrait
    MinPageWidth = 0.2

    Load frmSelField
    blnFormLoad = False
    
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Me.MousePointer = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim intResponse As Integer

'prompt to save the current report if needed
    If Not blnReportSaved Then
        intResponse = MsgBox("Do you wish to save " & OpenFileName & " ?", vbYesNo + vbQuestion, "Save Report")
        If intResponse = vbYes Then
            SaveReportDesign False
        End If
    End If
    
'unload any forms that have been loaded
    If VB.Forms.count > 0 Then
        For i = VB.Forms.count - 1 To 0 Step -1
            Unload Forms(i)
        Next i
    End If

End Sub

Private Sub Form_Resize()
'resize all the design screen and page components to suit

    SetupDesignScreen
    GetFirstLastSections
    SetPageComponents
    ShowVertScale
    ShowHorizScale
    If blnGridOn Then ShowGrid
    
End Sub

Private Sub SetupDesignScreen()
'sets up the design screen components based on the size of the form
Dim i As Integer

    Me.CoolBar1.width = Me.width
    If Me.ScaleWidth <= 5 Or Me.ScaleHeight <= 4 Then Exit Sub
    picContainer.width = Me.ScaleWidth - 0.8
    picContainer.Height = Me.ScaleHeight - (TopOfDesForm - 0.2708333) - 0.75
    
    picHorizScale.Top = TopOfDesForm + 0.05
    picContainer.Top = TopOfDesForm + 0.247
    picVertScale.Top = picContainer.Top + 0.167
    Frame1.Top = TopOfDesForm + 0.03
    cmdPgSetup.Top = TopOfDesForm + 0.05
    
    
    For i = 0 To 10
        cmdDivider(i).width = picContainer.ScaleWidth
    Next i
    
    VScroll.Top = picContainer.Top + 0.02
    VScroll.Left = picContainer.Left + picContainer.width - 0.02 - VScroll.width
    VScroll.Height = picContainer.Height - 0.05
    VScroll.Max = TotalPageHeight() * 8
    VScroll.Visible = (TotalPageHeight() > picContainer.Height)
    
    HScroll.Left = picContainer.Left + 0.02
    HScroll.Top = picContainer.Top + picContainer.Height - 0.02 - HScroll.Height
    If VScroll.Visible Then
        HScroll.width = picContainer.width - 0.04 - VScroll.width
    Else
        HScroll.width = picContainer.width - 0.04
    End If
    HScroll.Max = picSection(0).width * 8
    HScroll.Visible = (picSection(0).width + 0.2 >= picContainer.width)
    
    picHorizScale.width = picContainer.width
    picVertScale.Height = picContainer.Height - 0.167

End Sub

Private Sub FormatText_Click()
'user has clicked Format Text in the context menu, check if text already formatted
'and open text format dialog

    If InStr(1, ctlActive.Caption, "|") > 0 Then blnEditExisting = True
    frmNumFormat.Show

End Sub

Private Sub GridSpacing_Click()
'user has clicked on Design Grid menu item, show dialog

    frmGridSpace.Show vbModal
    ShowGrid

End Sub

Private Sub HScroll_Change()
'user has clicked on the horizontal scroll bar, update the design screen display
Dim i As Integer

    For i = 0 To 10
        picSection(i).Left = (0.167 - HScroll / 8)
    Next i

    StartScaleLeft = HScroll / 8
    ShowHorizScale
    ShowGrid

End Sub

Private Sub HScroll_GotFocus()

    Me.MousePointer = 0

End Sub

Private Sub HScroll_Scroll()
'user has scrolled the horizontal scroll bar, update the design screen display
Dim i As Integer

    For i = 0 To 10
        picSection(i).Left = (0.167 - HScroll / 8)
    Next i

    StartScaleLeft = HScroll / 8
    ShowHorizScale
    ShowGrid

End Sub

Private Sub SelectImage()
'user has placed an image control, open dialog to select image and load it into control
On Error GoTo NoFile

'set dialog to look for JPEG or BMP files (JPEG default)
    CommonDialog1.filter = "JPEG Files (*.jpg)|*.jpg|Windows Bitmaps|*.bmp|All Files|*.*"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName > "" Then
'check if file is Windows MetaFile - MSForms.Image control can't handle it (ironic, huh?)
        If Right(CommonDialog1.FileName, 4) = ".wmf" Or Right(CommonDialog1.FileName, 4) = ".WMF" Then
            MsgBox "Sorry, Windows Metafile format is not supported", vbOKOnly, "Format Not Supported"
            SelectImage
            Exit Sub
        End If
        If Dir(CommonDialog1.FileName, vbNormal) > "" Then
            ctlActive.Picture = LoadPicture(CommonDialog1.FileName)
'   adjust the size of the control to fit within the report section using the picture dimensions
            If ctlActive.Top + ctlActive.Height > picSection(ctlActive.Tag).Height Then
                ctlActive.Height = ctlActive.Picture.Height * 0.8 * (picSection(ctlActive.Tag).Height / ctlActive.Picture.Height)
                ctlActive.width = (ctlActive.Picture.width / ctlActive.Picture.Height) * ctlActive.Height
            End If
            ShowGrips (ctlActive.Tag)
            SetupSelectBox
            DoEvents
            picSection(ctlActive.Tag).Refresh
            CommonDialog1.FileName = ""
            WriteToUndoList unControl, unPlace
        End If
        
    Else
        DeleteSelectedCtlObject
    End If
    Exit Sub

NoFile:

End Sub

Private Sub Img_DblClick(Index As Integer, Cancel As MSForms.ReturnBoolean)

    SelectImage

End Sub

Private Sub Img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    CurrSection = Img(Index).Tag
    blnMouseIsDown = True
    blnDragStarted = False
   
    If Button = 2 Then
        Set ctlActive = Img(Index)
        ShowPopUp cImage
        Exit Sub
    End If
    
    SetupControlSelection Img(Index), Img(Index).Tag, 72, x, y

End Sub

Private Sub Img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    MoveInControl x, y

End Sub

Private Sub Img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    ConfirmControlSelection Img(Index), Img(Index).Tag
    
End Sub

Private Sub SetupControlSelection(ctl As Control, SecNo As Integer, ctlUnits As Long, GetX As Single, GetY As Single)
'generic sub to process selection of controls (other than line and shape)
'called on MouseDown event of control
    
    picSection(SecNo).Cls
    ShowGrid
    
    StartX = ctl.Left
    StartY = ctl.Top
    
    If Shift = 1 Then
        If NumInGrp = -1 Then
            NumInGrp = NumInGrp + 1
            ReDim Preserve SelectedCtl(NumInGrp)
            Set SelectedCtl(NumInGrp).ctl = ctlActive
            ClearGrips
        End If
        NumInGrp = NumInGrp + 1
        ReDim Preserve SelectedCtl(NumInGrp)
        Set SelectedCtl(NumInGrp).ctl = ctl
        blnGroupSelected = True
    End If
    If txtInput.Visible Then
        ctlActive.Caption = txtInput.text
        txtInput.Visible = False
        ctlActive.Visible = True
    End If
    Set ctlActive = ctl
    
    If blnSnapOn Then
        dX1 = SnapToGrid(GetX / ctlUnits)
        dY1 = SnapToGrid(GetY / ctlUnits)
    Else
        dX1 = GetX / ctlUnits
        dY1 = GetY / ctlUnits
    End If
    
    SetupSelectBox
    If blnGroupSelected Then
        blnControlSelected = False
        ShowGrid
        HiliteObjects
        intGroupRestraint = resNone
        GetRelativeDistances StartX + dX1, StartY + dY1
    Else
        blnControlSelected = True
        CurrSection = SecNo
        ShowGrips SecNo
        SetDividerColours SecNo
    End If

End Sub

Private Sub ConfirmControlSelection(ctl As Control, SecNo As Integer)
'generic sub to process selection of controls (other than line and shape)
'called on MouseUp event of control

    If blnDialogOpen Then
        blnDialogOpen = False
        Exit Sub
    End If
    
    blnMouseIsDown = False
    
    ctl.Visible = True
    If lngState <> EditLabel Then
        ctl.Left = SelectBox.Left
        ctl.Top = SelectBox.Top
        ResetMinSectionHt ctl.Tag
        ResetMinPageWidth
        blnDragStarted = False
    End If

    Set ctlActive = ctl
    ShowStateMouse
    GetTextFormat
    
    If blnGroupSelected Then
        If lngState = MoveControl Or lngState = MoveGrip Then HiliteObjects
    ElseIf blnControlSelected Then
        ShowGrips ctlActive.Tag
    End If
    
    blnReportSaved = False
    lngState = OverControl

End Sub

Private Sub Label_DblClick(Index As Integer)
    
    EditText_Click

End Sub

Private Sub Label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    CurrSection = Label(Index).Tag
    blnMouseIsDown = True
    blnDragStarted = False
        
    If Button = 2 Then
        Set ctlActive = Label(Index)
        ShowPopUp cLabel
        Exit Sub
    End If
    
    SetupControlSelection Label(Index), Label(Index).Tag, 1440, x, y

End Sub

Private Sub Label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    MoveInControl x, y
    
End Sub

Private Sub Label_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    ConfirmControlSelection Label(Index), Label(Index).Tag

End Sub

Private Sub SetupSelectBox()
'displays the dashed-line selection box at location of selected control

    SelectBox.Left = ctlActive.Left
    SelectBox.Top = ctlActive.Top
    SelectBox.width = ctlActive.width
    SelectBox.Height = ctlActive.Height
    
    Set SelectBox.Container = picSection(CurrSection)
    SelectBox.Visible = True
    SelectBox.ZOrder (0)
    
End Sub

Private Sub ShowGrips(SectIndex As Integer)
'displays the appropriate selection grips (boxes) for the selected control
'shape controls used for single selection, multiple selection uses owner drawn boxes
Dim i As Integer

    For i = 0 To 7
        Set Grip(i).Container = picSection(SectIndex)
    Next i
    
        If TypeOf ctlActive Is Line Then
            With ctlActive
                Grip(0).Left = .X1 - 0.03
                Grip(0).Top = .Y1 - 0.03
                Grip(1).Left = .X2 - 0.03
                Grip(1).Top = .Y2 - 0.03
                Grip(0).Visible = True: Grip(0).ZOrder (0)
                Grip(1).Visible = True: Grip(1).ZOrder (0)
            End With
        Else
            With ctlActive
                Grip(0).Left = .Left - 0.05
                Grip(0).Top = .Top - 0.05
                Grip(1).Left = .Left + .width - 0.01
                Grip(1).Top = .Top - 0.06
                Grip(2).Left = .Left + .width - 0.01
                Grip(2).Top = .Top + .Height - 0.02
                Grip(3).Left = .Left - 0.06
                Grip(3).Top = .Top + .Height - 0.02
                Grip(4).Left = .Left + .width / 2 - 0.03
                Grip(4).Top = .Top - 0.06
                Grip(5).Left = .Left + .width - 0.01
                Grip(5).Top = .Top + .Height / 2 - 0.03
                Grip(6).Left = .Left + .width / 2 - 0.03
                Grip(6).Top = .Top + .Height - 0.01
                Grip(7).Left = .Left - 0.05
                Grip(7).Top = .Top + .Height / 2 - 0.03
            End With
            For i = 0 To 7
                Grip(i).Visible = True
                Grip(i).ZOrder (0)
            Next i
        End If

End Sub

Private Sub BoundImg_DblClick(Index As Integer)

    SetDataField
    
End Sub

Private Sub BoundImg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    CurrSection = BoundImg(Index).Tag
    blnMouseIsDown = True
    blnDragStarted = False
        
    If Button = 2 Then
        Set ctlActive = BoundImg(Index)
        ShowPopUp cBoundImage
        Exit Sub
    End If
    
    SetupControlSelection BoundImg(Index), BoundImg(Index).Tag, 1440, x, y

End Sub

Private Sub BoundImg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    MoveInControl x, y, BoundImg(Index).ToolTipText
    
End Sub

Private Sub BoundImg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    ConfirmControlSelection BoundImg(Index), BoundImg(Index).Tag

End Sub

Private Sub mnuAbout_Click()

    frmAbout.Show

End Sub

Private Sub mnuCopyObj_Click()

    CopyObj_Click

End Sub

Private Sub mnuCutObj_Click()

    CutObj_Click
    
End Sub

Private Sub mnuDelObj_Click()

    DelObj_Click

End Sub

Private Sub mnuPasteObj_Click()

    PasteObj_Click

End Sub

Private Sub NewFile_Click()

    OpenNewFile

End Sub

Private Sub OpenFile_Click()

    OpenSavedFile

End Sub

Private Sub optAction_Click(Index As Integer)
'this handles actions for the command toolbar menu on the left side of the design screen,
'which is a frame with option buttons

    PropertySelector1.Visible = False
    lngPrevState = lngState
    If Index = 0 Then
        lngState = Default
        lngPrevState = Default
        lngCmdState = Default
        Set ctlActive = Nothing
        blnControlSelected = False
        ClearSelection
    Else
        intControlType = Index
        lngState = PlaceNewControl
        lngPrevState = PlaceNewControl
        lngCmdState = PlaceNewControl
        If Index = 8 Then
            frmDatePageField.Show vbModal
            optAction(Index).value = False
        ElseIf Index = 9 Then
            If blnReportDataBound Then
                If strDataFileName <> "" Then
                    frmCalcField.Show vbModal
                End If
            Else
                MsgBox "There is no open database connection, no fields available to calculate", vbOKOnly, "No Connection Found"
                lngState = Default
                lngCmdState = Default
                intControlType = cNone
            End If
            optAction(Index).value = False
        ElseIf Index = 10 Then
            If blnReportDataBound Then
                If strDataFileName <> "" Then
                    frmAggField.Show vbModal
                End If
            Else
                MsgBox "There is no open database connection, no fields available to summarize", vbOKOnly, "No Connection Found"
                intControlType = cNone
            End If
            lngState = Default
            lngCmdState = Default
            optAction(Index).value = False
        End If
    End If

    ShowStateMouse

End Sub

Private Sub ShowStateMouse()
'sets the appearance of the mouse pointer based on the current state of the program
    
    Select Case lngState
        Case Default
            Me.MousePointer = 0
        Case PlaceNewControl
            Me.MousePointer = 2
        Case MoveGrip
            Select Case ActiveGrip
                Case 0: Me.MousePointer = 8
                Case 1: Me.MousePointer = 6
                Case 2: Me.MousePointer = 8
                Case 3: Me.MousePointer = 6
                Case 4: Me.MousePointer = 7
                Case 5: Me.MousePointer = 9
                Case 6: Me.MousePointer = 7
                Case 7: Me.MousePointer = 9
            End Select
        Case ResizeSection
            Me.MousePointer = 7
        Case ResizePageWidth
            Me.MousePointer = 9
        Case SelectControl
            Me.MousePointer = 15
        Case OverControl
            Me.MousePointer = 5
    End Select
        
End Sub

Private Sub ClearSelection()
'clears off display of any current control selections
Dim i As Integer

    For i = 0 To 7
        Grip(i).Visible = False
        Grip(i).Visible = False
        Grip(i).Visible = False
    Next i

    For i = 0 To 4
        picSection(i).Cls
    Next i
    SelectBox.Visible = False
    
    For i = 1 To 9
        Toolbar2.Buttons(i).value = tbrUnpressed
    Next i

    If blnGridOn Then ShowGrid

End Sub

Private Sub ClearGrips()
'clears display of control grips (boxes)
Dim i As Integer

    For i = 0 To 7
        Grip(i).Visible = False
        Grip(i).Visible = False
        Grip(i).Visible = False
    Next i

End Sub

Private Function IsInSelectGroup(GetName As String, GetIndex As Long) As Boolean
'checks whether control clicked on is already in the selected group
Dim i As Integer

    For i = 0 To UBound(SelectedCtl)
        If SelectedCtl(i).ctl.Name = GetName And SelectedCtl(i).ctl.Index = GetIndex Then
            IsInSelectGroup = True
            Exit For
        End If
    Next i

End Function

Private Sub ExecuteState(SectionIndex As Integer, GetState As Long, GetX As Single, GetY As Single)
'this sub controls the state of the program while the mouse is moving
Dim i As Integer, j As Integer

    If blnSnapOn Then
        GetX = SnapToGrid(GetX)
        GetY = SnapToGrid(GetY)
    End If

    Select Case GetState
        Case Default
            Me.MousePointer = 0
        Case PlaceNewControl
            blnReportSaved = False
            ClearSelection
            If intControlType = cLine Then
                ActiveGrip = 0
                CreateLine SectionIndex, GetX, GetY
            ElseIf intControlType = cBox Then
                ActiveGrip = 2
                CreateShape SectionIndex
            ElseIf intControlType = cLabel Then
                ActiveGrip = 2
                CreateLabel SectionIndex
            ElseIf intControlType = cDataField Then
                ActiveGrip = 2
                CreateField SectionIndex, "Data"
            ElseIf intControlType = cCheckBox Then
                ActiveGrip = 0
            ElseIf intControlType = cImage Then
                ActiveGrip = 2
                CreateImageBox SectionIndex
            ElseIf intControlType = cBoundImage Then
                ActiveGrip = 2
                CreateBoundImageBox SectionIndex
            ElseIf intControlType >= cDateField Then
                ActiveGrip = 2
                CreateField SectionIndex, "Date", strSpecialFieldContent
                Me.optAction(8).value = False
            End If
        Case MoveGrip
            blnReportSaved = False
            If blnControlSelected Then
                If TypeOf ctlActive Is Line Then
                    ResizeLine SectionIndex, GetX, GetY
                Else
                    If ctlActive.Name <> "ChkBox" Then
                        ResizeShape SectionIndex, GetX, GetY
                    End If
                End If
            ElseIf blnGroupSelected Then
                SelectBox.Visible = False
                ResizeSelectedObjects GetX, GetY
            End If
        Case MoveControl
            blnReportSaved = False
            If blnGroupSelected Then
                If IsInSelectGroup(ctlActive.Name, ctlActive.Index) Then
                    MoveSelectedObjects SectionIndex, GetX, GetY
                Else
                    blnGroupSelected = False
                    ClearSelection
                    blnControlSelected = True
                End If
            Else
                MoveTheControl SectionIndex, GetX, GetY
            End If
        Case ResizeSection
            blnReportSaved = False
            For i = FirstSectionVis To LastSectionVis
                If i = SectionIndex Then
                    If GetY > MinSectionHt(i) Then
                        picSection(i).Height = GetY
                    Else
                        picSection(i).Height = MinSectionHt(i)
                    End If
                    Me.lnPos.Y1 = picSection(i).Top + picSection(i).Height - 0.17
                    Me.lnPos.Y2 = Me.lnPos.Y1
                    If i < LastSectionVis Then
                        For j = i + 1 To LastSectionVis
                            If picSection(j).Visible Then
                                cmdDivider(j).Top = picSection(i).Height + picSection(i).Top
                                Exit For
                            End If
                        Next j
                    End If
                ElseIf i > SectionIndex Then
                    If picSection(i).Visible Then
                        picSection(i).Top = cmdDivider(i).Top + cmdDivider(i).Height
                        For j = i + 1 To LastSectionVis
                            If picSection(j).Visible Then
                                cmdDivider(j).Top = picSection(i).Top + picSection(i).Height
                                Exit For
                            End If
                        Next j
                    End If
                End If
            Next i
            If Not blnDragStarted Then
                WriteToUndoList unSection, unSectHeight
                blnDragStarted = True
            End If
            VScroll.Max = TotalPageHeight() * 8
            VScroll.Visible = TotalPageHeight() > picContainer.Height
            If VScroll.Visible = False Then
                HScroll.width = picContainer.width - 0.04
            Else
                HScroll.width = picContainer.width - 0.04 - VScroll.width
            End If
        Case ResizePageWidth
            blnReportSaved = False
            If GetX > MinPageWidth Then
                For i = 0 To 10
                    picSection(i).width = GetX
                Next i
            Else
                For i = 0 To 10
                    picSection(i).width = MinPageWidth
                Next i
            End If
            If Not blnDragStarted Then
                WriteToUndoList unSection, unSectWidth
                blnDragStarted = True
            End If
            HScroll.Max = picSection(0).width * 8
            HScroll.Visible = (picSection(0).width + 0.2 >= picContainer.width)
        Case SelectGroup
            ReSizeSelectBox SectionIndex, GetX, GetY
        Case ReadyForAggControl
            Me.MousePointer = 2
    End Select

End Sub

Private Sub ReSizeSelectBox(GetSec As Integer, MousX As Single, MousY As Single)
'sizes the selection box based on the mouse coordinates
Dim i As Integer, k As Integer

    If blnGridOn Then
        ShowGrid
    Else
        picSection(GetSec).Cls
    End If
    If Not blnDragStarted Then
        SelectBox.ZOrder (0)
        blnDragStarted = True
    End If
    
    blnGroupSelected = False
    NumInGrp = -1
    
    With SelectBox
        If MousX >= StartX Then
            .Left = StartX
            .width = MousX - StartX
        Else
            .Left = MousX
            .width = StartX - MousX
        End If
        
        If MousY >= StartY Then
            .Top = StartY
            .Height = MousY - StartY
        Else
            .Top = MousY
            .Height = StartY - MousY
        End If
        .Visible = True
        
    End With
        
End Sub

Private Sub SelectObjects(GetSec As Integer)
'finds and selects object inside the selection rectangle
Dim i As Integer, k As Integer

    NumInGrp = -1
    With SelectBox
        For i = 0 To Me.Controls.count - 1
            Set ctlTest = Me.Controls(i)
            If ctlTest.Tag = GetSec Then
                If TypeOf ctlTest Is Line Then
                    If .Left < ctlTest.X1 And .Left < ctlTest.X2 And _
                    .Left + .width > ctlTest.X1 And _
                    .Left + .width > ctlTest.X2 Then
                        If .Top < ctlTest.Y1 And .Top < ctlTest.Y2 And _
                        .Top + .Height > ctlTest.Y1 And _
                        .Top + .Height > ctlTest.Y2 Then
                            blnGroupSelected = True
                            NumInGrp = NumInGrp + 1
                            ReDim Preserve SelectedCtl(NumInGrp)
                            Set SelectedCtl(NumInGrp).ctl = ctlTest
                        End If
                    End If
                Else
                    If .Left < ctlTest.Left And _
                    .Left + .width > ctlTest.Left + ctlTest.width Then
                        If .Top < ctlTest.Top And _
                        .Top + .Height > ctlTest.Top + ctlTest.Height Then
                            blnGroupSelected = True
                            NumInGrp = NumInGrp + 1
                            ReDim Preserve SelectedCtl(NumInGrp)
                            Set SelectedCtl(NumInGrp).ctl = ctlTest
                        End If
                    End If
                End If
            End If
            blnSelectArrayInit = True
        Next i
    End With
        
End Sub

Public Sub HiliteObjects()
'shows control selection grips for all selected controls
Dim k As Integer

    If blnGroupSelected Then
        Dim ctlSec As Integer
        For k = FirstSectionVis To LastSectionVis
            picSection(k).Cls
            ShowGrid
        Next k
        For k = 0 To UBound(SelectedCtl)
            ctlSec = SelectedCtl(k).ctl.Tag
            picSection(ctlSec).FillStyle = 1
            picSection(ctlSec).DrawStyle = 0
            picSection(ctlSec).ForeColor = &H8000000D
            If TypeOf SelectedCtl(k).ctl Is Line Then
                With SelectedCtl(k).ctl
                    picSection(ctlSec).Line (.X1 - 0.03, .Y1 - 0.03)-(.X1 + 0.03, .Y1 + 0.03), , BF
                    picSection(ctlSec).Line (.X2 - 0.03, .Y2 - 0.03)-(.X2 + 0.03, .Y2 + 0.03), , BF
                End With
            Else
                With SelectedCtl(k).ctl
                    picSection(ctlSec).Line (.Left - 0.06, .Top - 0.06)-(.Left, .Top), , BF
                    picSection(ctlSec).Line (.Left + .width - 0.01, .Top - 0.06)-(.Left + .width + 0.05, .Top), , BF
                    picSection(ctlSec).Line (.Left + .width - 0.01, .Top + .Height - 0.01)-(.Left + .width + 0.05, .Top + .Height + 0.05), , BF
                    picSection(ctlSec).Line (.Left - 0.06, .Top + .Height - 0.01)-(.Left, .Top + .Height + 0.05), , BF
                    picSection(ctlSec).Line (.Left + .width / 2 - 0.03, .Top - 0.06)-(.Left + .width / 2 + 0.03, .Top), , BF
                    picSection(ctlSec).Line (.Left + .width - 0.01, .Top + .Height / 2 - 0.03)-(.Left + .width + 0.05, .Top + .Height / 2 + 0.03), , BF
                    picSection(ctlSec).Line (.Left + .width / 2 - 0.03, .Top + .Height - 0.01)-(.Left + .width / 2 + 0.03, .Top + .Height + 0.05), , BF
                    picSection(ctlSec).Line (.Left - 0.06, .Top + .Height / 2 - 0.03)-(.Left, .Top + .Height / 2 + 0.03), , BF
                End With
            End If
        Next k
    End If
        
End Sub

Private Sub MoveSelectedObjects(SecNo As Integer, MousX As Single, MousY As Single)
'moves all selected controls relative to mouse coordinates
Dim i As Integer

    If Not blnDragStarted Then
        ClearSelection
        WriteToUndoList unControl, unMove
        blnDragStarted = True
    End If
    
    blnAtLimit = False
    For i = 0 To UBound(SelectedCtl)
        With SelectedCtl(i)
                If TypeOf .ctl Is Line Then
                    If (MousY - .dY1 <= 0 Or MousY - .dY1 >= picSection(SecNo).Height) Or _
                     (MousY - .dY2 <= 0 Or MousY - .dY2 >= picSection(SecNo).Height) Then
                        blnAtLimit = True
                    End If
                ElseIf TypeOf .ctl Is Shape Then
                    If MousY - .dY1 <= 0 Or (MousY - .dY1 + .ctl.Height) >= picSection(SecNo).Height Then
                        blnAtLimit = True
                    End If
                Else
                    If .ctl.Visible Then
                        If .ctl.Name <> ctlActive.Name Or .ctl.Index <> ctlActive.Index Then
                            If MousY - .dY1 <= 0 Or (MousY - .dY1 + .ctl.Height) >= picSection(SecNo).Height Then
                                blnAtLimit = True
                            End If
                        End If
                    End If
                End If
        End With
    Next i
    
    For i = 0 To UBound(SelectedCtl)
        With SelectedCtl(i)
            If TypeOf .ctl Is Line Then
                .ctl.X1 = MousX - .dX1
                .ctl.X2 = MousX - .dX2
                If Not blnAtLimit Then
                    .ctl.Y1 = MousY - .dY1
                    .ctl.Y2 = MousY - .dY2
               End If
            ElseIf TypeOf .ctl Is Shape Then
                .ctl.Left = MousX - .dX1
                If Not blnAtLimit Then .ctl.Top = MousY - .dY1
            Else
                If .ctl.Visible Then
                    If .ctl.Name <> ctlActive.Name Or .ctl.Index <> ctlActive.Index Then
                        .ctl.Left = MousX - .dX1
                        If Not blnAtLimit Then .ctl.Top = MousY - .dY1
                    End If
                End If
            End If
        End With
    Next i
 
End Sub
 
Private Function TotalPageHeight() As Single
'calculates the current total height of all displayed sections
'used in display of scroll bars
Dim GetHt As Single
Dim i As Integer

    For i = 0 To LastSectionVis
        If picSection(i).Visible Then
            GetHt = GetHt + picSection(i).Height + 0.17
        End If
    Next i

    TotalPageHeight = GetHt + 0.08

End Function

Private Sub GetRelativeDistances(MouseX As Single, MouseY As Single)
'calculates the relative distance of the mouse to the selected control
'control and then to all the other selected controls
    If blnSnapOn Then
        MouseX = SnapToGrid(MouseX)
        MouseY = SnapToGrid(MouseY)
    End If

    If TypeOf ctlActive Is Line Then
        dX1 = MouseX - ctlActive.X1
        dY1 = MouseY - ctlActive.Y1
        dX2 = MouseX - ctlActive.X2
        dY2 = MouseY - ctlActive.Y2
    ElseIf TypeOf ctlActive Is Shape Then
        dX1 = MouseX - ctlActive.Left
        dY1 = MouseY - ctlActive.Top
    End If
    
    If blnGroupSelected Then
        Dim i As Integer
        If NumInGrp >= 0 Then
            For i = 0 To UBound(SelectedCtl)
                If TypeOf SelectedCtl(i).ctl Is Line Then
                    SelectedCtl(i).dX1 = MouseX - SelectedCtl(i).ctl.X1
                    SelectedCtl(i).dY1 = MouseY - SelectedCtl(i).ctl.Y1
                    SelectedCtl(i).dX2 = MouseX - SelectedCtl(i).ctl.X2
                    SelectedCtl(i).dY2 = MouseY - SelectedCtl(i).ctl.Y2
                Else
                    SelectedCtl(i).dX1 = MouseX - SelectedCtl(i).ctl.Left
                    SelectedCtl(i).dY1 = MouseY - SelectedCtl(i).ctl.Top
                End If
                SelectedCtl(i).ActiveGrip = ActiveGrip
            Next i
        End If
    End If

End Sub

Private Sub MoveTheControl(GetSection As Integer, MouseX As Single, MouseY As Single)
'sub for moving lines and shapes which have no Windows handle
    
    If Not blnDragStarted Then
        WriteToUndoList unControl, unMove
        blnDragStarted = True
        ClearSelection
    End If
    If TypeOf ctlActive Is Line Then
        ctlActive.X1 = MouseX - dX1
        ctlActive.X2 = MouseX - dX2
        If (MouseY - dY1 >= 0 And MouseY - dY1 <= picSection(GetSection).Height) And _
         (MouseY - dY2 >= 0 And MouseY - dY2 <= picSection(GetSection).Height) Then
            ctlActive.Y1 = MouseY - dY1
            ctlActive.Y2 = MouseY - dY2
        End If
    ElseIf TypeOf ctlActive Is Shape Then
        ctlActive.Left = MouseX - dX1
        If MouseY - dY1 >= 0 And (MouseY - dY1 + ctlActive.Height) <= picSection(GetSection).Height Then
            ctlActive.Top = MouseY - dY1
        End If
    End If

End Sub

Private Sub AddToDrawList()
'adds a newly created control to the draw order list

        ReDim Preserve DrawList(dLstCount)
        
        DrawList(dLstCount).ctlName = ctlActive.Name
        DrawList(dLstCount).ctlIndex = ctlActive.Index
        dLstCount = dLstCount + 1
        
End Sub

Private Sub CreateLine(GetIndex As Integer, GetX As Single, GetY As Single)
'create a new line control

    LineIndex = LineIndex + 1
    Load Lin(LineIndex)
    Set ctlActive = Lin(LineIndex)
    Set ctlActive.Container = picSection(GetIndex)
    With ctlActive
        .ZOrder (0)
        .Tag = GetIndex
        If blnSnapOn Then
            .X1 = SnapToGrid(StartX)
            .Y1 = SnapToGrid(StartY)
        Else
            .X1 = StartX
            .Y1 = StartY
        End If
        .X2 = GetX
        .Y2 = YInsideSection(GetY, picSection(GetIndex).Height)
        .BorderWidth = ActiveBorder.width
        .BorderColor = ActiveBorder.Color
        .BorderStyle = ActiveBorder.Style + 1
    End With
    ctlActive.Visible = True
    
    Set Grip(0).Container = picSection(GetIndex)
    Set Grip(1).Container = picSection(GetIndex)
    Grip(0).Left = StartX - 0.04
    Grip(0).Top = StartY - 0.04
    Grip(1).Left = GetX - 0.04
    Grip(1).Top = GetY - 0.04
    Grip(0).Visible = True
    Grip(1).Visible = True
    
    blnControlSelected = True
    blnCreatingControl = True
    lngState = MoveGrip
    
    AddToDrawList
   
End Sub

Private Sub ResizeLine(GetIndex As Integer, GetX As Single, GetY As Single)
'resize a line from the current mouse coordinates
Dim LimitedY As Single

    LimitedY = YInsideSection(GetY, picSection(GetIndex).Height)

    If ActiveGrip > 0 Then
        ctlActive.X2 = GetX
        ctlActive.Y2 = LimitedY
    Else
        ctlActive.X1 = GetX
        ctlActive.Y1 = LimitedY
    End If
    
    Grip(ActiveGrip).Left = GetX - 0.04
    Grip(ActiveGrip).Top = LimitedY - 0.04
        
End Sub

Private Sub CreateShape(GetIndex As Integer)
'creates a new shape control
Dim i As Integer
    
    ShapeIndex = ShapeIndex + 1
    Load Shape(ShapeIndex)
    Set ctlActive = Shape(ShapeIndex)
    Set ctlActive.Container = picSection(GetIndex)
    With ctlActive
        .Shape = ActiveShape
        .ZOrder (0)
        .Tag = GetIndex
        If blnSnapOn Then
            .Left = SnapToGrid(StartX)
            .Top = SnapToGrid(StartY)
        Else
            .Left = StartX
            .Top = StartY
        End If
        .BorderStyle = ActiveBorder.Style + 1
        .BorderWidth = ActiveBorder.width
        .BorderColor = ActiveBorder.Color
        .BackStyle = ActiveBack.Style
        .BackColor = ActiveBack.Color
        .Visible = True
    End With
    For i = 0 To 7
        Grip(i).Left = StartX - 0.04
        Grip(i).Top = StartY - 0.04
        Grip(i).Visible = True
    Next i

    blnControlSelected = True
    blnCreatingControl = True
    lngState = MoveGrip

    AddToDrawList

End Sub

Private Sub ResizeShape(GetIndex As Integer, GetX As Single, GetY As Single)
'resize shape control based on current mouse coordinates
'and which corner is being moved
Dim oldLeft As Single, oldTop As Single, OldRight As Single, OldBottom As Single
Dim LimitedY As Single

    LimitedY = YInsideSection(GetY, picSection(GetIndex).Height)
'
'  Grip Positions
'          0----4----1
'          |         |
'          7         5
'          |         |
'          3----6----2
'
    With ctlActive
        Select Case ActiveGrip
            Case 0
                If .Left + .width > GetX Then
                    OldRight = .Left + .width
                    .Left = GetX
                    .width = OldRight - GetX
                Else
                    ActiveGrip = 1
                End If
                If .Top + .Height > GetY Then
                    OldBottom = .Top + .Height
                    .Top = LimitedY
                    .Height = OldBottom - LimitedY
                Else
                    ActiveGrip = 3
                End If
            Case 1
                If .Left < GetX Then
                    .width = GetX - .Left
                Else
                    ActiveGrip = 0
                End If
                If .Top + .Height > LimitedY Then
                    OldBottom = .Top + .Height
                    .Top = LimitedY
                    .Height = OldBottom - LimitedY
                Else
                    ActiveGrip = 2
                End If
            Case 2
                If .Left < GetX Then
                    .width = GetX - .Left
                Else
                    ActiveGrip = 3
                End If
                If .Top < LimitedY Then
                    .Height = LimitedY - .Top
                Else
                    ActiveGrip = 1
                End If
            Case 3
                If .Left + .width > GetX Then
                    OldRight = .Left + .width
                    .Left = GetX
                    .width = OldRight - GetX
                Else
                    ActiveGrip = 2
                End If
                If .Top < LimitedY Then
                    .Height = LimitedY - .Top
                Else
                    ActiveGrip = 1
                End If
            Case 4
                If .Top + .Height > GetY Then
                    OldBottom = .Top + .Height
                    .Top = LimitedY
                    .Height = OldBottom - LimitedY
                Else
                    ActiveGrip = 6
                End If
            Case 5
                If .Left < GetX Then
                    .width = GetX - .Left
                Else
                    ActiveGrip = 7
                End If
            Case 6
                If .Top < LimitedY Then
                    .Height = LimitedY - .Top
                Else
                    ActiveGrip = 4
                End If
            Case 7
                If .Left + .width > GetX Then
                    OldRight = .Left + .width
                    .Left = GetX
                    .width = OldRight - GetX
                Else
                    ActiveGrip = 5
                End If
        End Select
    End With
    
    SetupSelectBox
    blnControlSelected = True
    ShowGrips GetIndex
    
End Sub

Private Sub ResizeSelectedObjects(GetX As Single, GetY As Single)
'resize all selected controls based on current mouse coordinates and
'which corner of selected control is being moved
Dim i As Integer

'
'  Grip Positions
'          0----4----1
'          |         |
'          7         5
'          |         |
'          3----6----2
'
    dX2 = GetX - dX1
    dY2 = GetY - dY1

    For i = 0 To UBound(SelectedCtl)
        With SelectedCtl(i)
            If TypeOf .ctl Is Line Then
                If .ActiveGrip = 0 Then
                    .ctl.X1 = .ctl.X1 + dX2
                    .ctl.Y1 = .ctl.Y1 + dY2
                ElseIf .ActiveGrip < 4 Then
                    .ctl.X2 = .ctl.X2 + dX2
                    .ctl.Y2 = .ctl.Y2 + dY2
                ElseIf .ActiveGrip = 4 Then
                    .ctl.Y1 = .ctl.Y1 + dY2
                ElseIf .ActiveGrip = 5 Then
                    .ctl.X1 = .ctl.X1 + dX2
                ElseIf .ActiveGrip = 6 Then
                    .ctl.Y2 = .ctl.Y2 + dY2
                ElseIf .ActiveGrip = 7 Then
                    .ctl.X2 = .ctl.X2 + dX2
                End If
            Else
                Select Case .ActiveGrip
                    Case 0
                        If .ctl.width - dX2 > 0 Then
                            .ctl.Left = .ctl.Left + dX2
                            .ctl.width = .ctl.width - dX2
                        Else
                            .ActiveGrip = 1
                        End If
                        If .ctl.Height - dY2 > 0 Then
                            .ctl.Top = .ctl.Top + dY2
                            .ctl.Height = .ctl.Height - dY2
                        Else
                            .ActiveGrip = 3
                        End If
                    Case 1
                        If .ctl.width + dX2 > 0 Then
                            .ctl.width = .ctl.width + dX2
                        Else
                            .ActiveGrip = 0
                        End If
                        If .ctl.Height - dY2 > 0 Then
                            .ctl.Top = .ctl.Top + dY2
                            .ctl.Height = .ctl.Height - dY2
                        Else
                            .ActiveGrip = 2
                        End If
                    Case 2
                        If .ctl.width + dX2 > 0 Then
                            .ctl.width = .ctl.width + dX2
                        Else
                            .ActiveGrip = 3
                        End If
                        If .ctl.Height + dY2 > 0 Then
                            .ctl.Height = .ctl.Height + dY2
                        Else
                            .ActiveGrip = 1
                        End If
                    Case 3
                        If .ctl.width - dX2 > 0 Then
                            .ctl.width = .ctl.width - dX2
                        Else
                            .ActiveGrip = 2
                        End If
                        If .ctl.Height + dY2 > 0 Then
                            .ctl.Left = .ctl.Left + dX2
                            .ctl.Height = .ctl.Height + dY2
                        Else
                            .ActiveGrip = 0
                        End If
                    Case 4
                        If .ctl.Height - dY2 > 0 Then
                            .ctl.Top = .ctl.Top + dY2
                            .ctl.Height = .ctl.Height - dY2
                        Else
                            .ActiveGrip = 6
                        End If
                    Case 5
                        If .ctl.width + dX2 > 0 Then
                            .ctl.width = .ctl.width + dX2
                        Else
                            .ActiveGrip = 7
                        End If
                    Case 6
                        If .ctl.Height + dY2 > 0 Then
                            .ctl.Height = .ctl.Height + dY2
                        Else
                            .ActiveGrip = 4
                        End If
                    Case 7
                        If .ctl.width - dX2 > 0 Then
                            .ctl.Left = .ctl.Left + dX2
                            .ctl.width = .ctl.width - dX2
                        Else
                            .ActiveGrip = 5
                        End If
                End Select
            End If
        End With
    Next i
    
    dX1 = GetX
    dY1 = GetY
    
    ShowGrid ctlActive.Tag

End Sub

Private Sub CreateLabel(GetIndex As Integer)

    LabelIndex = LabelIndex + 1
    Load Label(LabelIndex)
    Set ctlActive = Label(LabelIndex)
    Set ctlActive.Container = picSection(GetIndex)
    With ctlActive
        .ZOrder (0)
        .FontName = ActiveText.FontName
        .FontSize = ActiveText.FontSize
        .FontBold = ActiveText.IsBold
        .FontUnderline = ActiveText.IsUnderline
        .FontItalic = ActiveText.IsItalic
        .Alignment = ActiveText.Align
        .BackStyle = ActiveBack.Style
        .BackColor = ActiveBack.Color
        .ForeColor = ActiveText.Color
        .BorderStyle = -(ActiveText.BorderOn)
        .LinkTimeout = cLabel
        .Tag = GetIndex
        If blnSnapOn Then
            .Left = SnapToGrid(StartX)
            .Top = SnapToGrid(StartY)
        Else
            .Left = StartX
            .Top = StartY
        End If
        .Caption = "New Label" & LabelIndex
        .Visible = True
    End With
    ShowGrips GetIndex
    GetTextFormat
    
    blnControlSelected = True
    blnCreatingControl = True
    lngState = MoveGrip
    
    AddToDrawList

End Sub

Private Sub CreateBoundImageBox(GetIndex As Integer)

    BoundImageIndex = BoundImageIndex + 1
    Load BoundImg(BoundImageIndex)
    Set ctlActive = BoundImg(BoundImageIndex)
    Set ctlActive.Container = picSection(GetIndex)
    ctlActive.LinkTimeout = cBoundImage
    With ctlActive
        .ZOrder (0)
        .Alignment = 0
        .FontName = "Arial"
        .FontSize = 10
        .Tag = GetIndex
        .Caption = "Unbound"
        If blnSnapOn Then
            .Left = SnapToGrid(StartX)
            .Top = SnapToGrid(StartY)
        Else
            .Left = StartX
            .Top = StartY
        End If
        .Visible = True
    End With
    ShowGrips GetIndex

    blnCreatingControl = True
    blnControlSelected = True
    lngState = MoveGrip

    AddToDrawList

End Sub

Public Sub CreateField(GetIndex As Integer, fldType As String, _
Optional strContent As String = "Unbound", Optional strFieldName As String = "", _
Optional blnPlaceNew As Boolean = False)

    FieldIndex = FieldIndex + 1
    Load Field(FieldIndex)
    Set ctlActive = Field(FieldIndex)
    Set ctlActive.Container = picSection(GetIndex)
    With ctlActive
        .ZOrder (0)
        .FontName = ActiveText.FontName
        .FontSize = ActiveText.FontSize
        .FontBold = ActiveText.IsBold
        .FontUnderline = ActiveText.IsUnderline
        .FontItalic = ActiveText.IsItalic
        .Alignment = ActiveText.Align
        .BorderStyle = -(ActiveText.BorderOn)
        .BackStyle = ActiveBack.Style
        .BackColor = ActiveBack.Color
        .ForeColor = ActiveText.Color
        .BorderStyle = -(ActiveText.BorderOn)
        .LinkTimeout = intControlType
        .Tag = GetIndex
        If blnPlaceNew Then
            .Left = 0.05
            .Top = 0.05
            .width = 0.5
            .Height = 0.2
        ElseIf blnSnapOn Then
            .Left = SnapToGrid(StartX)
            .Top = SnapToGrid(StartY)
        Else
            .Left = StartX
            .Top = StartY
        End If
        If .LinkTimeout = cDataField And strContent <> "Unbound" Then
            .Caption = strFieldName
            .ToolTipText = strContent
        Else
            .Caption = strContent
            .ToolTipText = strContent
        End If
        If strFieldName > "" Then .DataField = strFieldName
        .Visible = True
    End With
    ShowGrips GetIndex
    
    blnControlSelected = True
    blnCreatingControl = True
    lngState = MoveGrip

    AddToDrawList

End Sub

Private Sub CreateImageBox(GetIndex As Integer)

    ImageIndex = ImageIndex + 1
    Load Img(ImageIndex)
    Set ctlActive = Img(ImageIndex)
    Set ctlActive.Container = picSection(GetIndex)
    With ctlActive
        .ZOrder (0)
        .BorderStyle = 0
        .BackStyle = 1
        .Tag = GetIndex
        If blnSnapOn Then
            .Left = SnapToGrid(StartX)
            .Top = SnapToGrid(StartY)
        Else
            .Left = StartX
            .Top = StartY
        End If
        .Visible = True
    End With
    ShowGrips GetIndex
    
    blnControlSelected = True
    blnCreatingControl = True
    lngState = MoveGrip

    AddToDrawList

End Sub

Private Sub CreateCheckBox(GetIndex As Integer)

    CheckShapeIndex = CheckShapeIndex + 1
    Load Chkbox(CheckShapeIndex)
    Set ctlActive = Chkbox(CheckShapeIndex)
    Set ctlActive.Container = picSection(GetIndex)
    With ctlActive
        .ZOrder (0)
        .BackColor = ActiveBack.Color
        .BorderColor = ActiveBorder.Color
        .Tag = GetIndex
        If blnSnapOn Then
            .Left = SnapToGrid(StartX)
            .Top = SnapToGrid(StartY)
        Else
            .Left = StartX
            .Top = StartY
        End If
        .DisplayType = ActiveChkBox.DisplayType
        .Sunken = ActiveChkBox.Sunken
        .ToolTipText = "Double-click to link to a data field"
        .Visible = True
    End With
    ShowGrips GetIndex
    intControlType = cCheckBox
    
    blnControlSelected = True
    blnCreatingControl = True
    lngState = PlaceNewControl
   
    AddToDrawList
    WriteToUndoList unControl, unPlace

End Sub

Private Sub OptJustifyClicked(intJustify As Integer)
    
    WriteToUndoList unControl, unFormat
    
    ActiveText.Align = intJustify
    If blnControlSelected = True Then
        If TypeOf ctlActive Is Label Then
            ctlActive.Alignment = intJustify
        End If
    End If

    GetTextFormat

End Sub

Private Function YInsideSection(GetY As Single, GetMax As Single) As Single
'attempts to keep user from creating control beyond page section boundaries vertically
'doesn't always work!

    If GetY < 0 Then
        YInsideSection = 0
    ElseIf GetY > GetMax Then
        YInsideSection = GetMax
    Else
        YInsideSection = GetY
    End If

End Function

Public Sub ShowVertScale()
'draw the vertical scale markers and numbers for all page sections
Dim a As Integer
        
    picVertScale.Cls
    
    For a = FirstSectionVis To LastSectionVis
        If picSection(a).Visible Then
            DoSectionVertScale a
            DoScaleDividers a
        End If
    Next a
        
End Sub

Private Sub DoSectionVertScale(SecNo As Integer)
'draw the vertical scale markers and numbers for a given page section
'uses the Picturebox.Line method and the Windows DrawText API

Dim TickY As Single, sglValue As Single, i As Integer
Dim j As Integer, SecStartY As Single
Dim LowerLimit As Single

    If SecNo = LastSectionVis Then
        LowerLimit = picVertScale.ScaleHeight
    Else
        LowerLimit = picSection(SecNo).Top + picSection(SecNo).Height
    End If

    TickY = picSection(SecNo).Top
    SecStartY = TickY
    If PageScaleUnits = scEnglish Then
        Do While TickY <= LowerLimit
            TickY = TickY + 0.125
            If Round(TickY - SecStartY, 3) = 1 Then
                sglValue = sglValue + 1
                recVSclText(i).Top = (TickY - 0.257) * 1440 / Screen.TwipsPerPixelY
                recVSclText(i).Bottom = (TickY - 0.107) * 1440 / Screen.TwipsPerPixelY
                If sglValue < 10 Then
                    recVSclText(i).Left = 0.04 * 1440 / Screen.TwipsPerPixelX
                    recVSclText(i).Right = 0.2 * 1440 / Screen.TwipsPerPixelX
                    DrawText picVertScale.hdc, Format(sglValue, "0"), 1, recVSclText(i), vbLeftJustify
                Else
                    recVSclText(i).Left = 0.015 * 1440 / Screen.TwipsPerPixelX
                    recVSclText(i).Right = 0.2 * 1440 / Screen.TwipsPerPixelX
                    DrawText picVertScale.hdc, Format(sglValue, "0"), 2, recVSclText(i), vbLeftJustify
                End If
                SecStartY = TickY
                i = i + 1
            ElseIf Round(TickY - SecStartY, 3) = 0.5 Then
                picVertScale.Line (0.03, TickY - 0.177)-(0.1, TickY - 0.177)
            Else
                picVertScale.Line (0.05, TickY - 0.177)-(0.08, TickY - 0.177)
            End If
        Loop
    ElseIf PageScaleUnits = scMetric Then
        Do While TickY <= LowerLimit
            TickY = TickY + 0.0788
            If Round(TickY - SecStartY, 4) = 0.394 Then
                sglValue = sglValue + 1
                recVSclText(i).Top = (TickY - 0.257) * 1440 / Screen.TwipsPerPixelY
                recVSclText(i).Bottom = (TickY - 0.107) * 1440 / Screen.TwipsPerPixelY
                If sglValue < 10 Then
                    recVSclText(i).Left = 0.04 * 1440 / Screen.TwipsPerPixelX
                    recVSclText(i).Right = 0.2 * 1440 / Screen.TwipsPerPixelX
                    DrawText picVertScale.hdc, Format(sglValue, "0"), 1, recVSclText(i), vbLeftJustify
                Else
                    recVSclText(i).Left = 0.015 * 1440 / Screen.TwipsPerPixelX
                    recVSclText(i).Right = 0.2 * 1440 / Screen.TwipsPerPixelX
                    DrawText picVertScale.hdc, Format(sglValue, "0"), 2, recVSclText(i), vbLeftJustify
                End If
                SecStartY = TickY
                i = i + 1
            Else
                picVertScale.Line (0.05, TickY - 0.177)-(0.08, TickY - 0.177)
            End If
        Loop
    End If
    
End Sub

Private Sub DoScaleDividers(SecNo As Integer)
'draws the page section dividers in the vertical scale bar
'all done with the Picturebox.Line method - seems to work well enough
Dim LowerLimit As Single

    If SecNo = LastSectionVis Then
        LowerLimit = picVertScale.ScaleHeight
    Else
        LowerLimit = picSection(SecNo).Top + picSection(SecNo).Height
    End If

    If SecNo < LastSectionVis Then
        For i = SecNo + 1 To LastSectionVis
            If cmdDivider(i).Visible Then
                picVertScale.ForeColor = cmdDivider(i).BackColor
                Exit For
            End If
        Next i
        picVertScale.Line (0, LowerLimit - 0.17)-(0.15, LowerLimit), , BF
        picVertScale.DrawWidth = 2
        picVertScale.ForeColor = &H80000014
        picVertScale.Line (-0.01, LowerLimit - 0.16)-(0.15, LowerLimit - 0.16)
        picVertScale.ForeColor = &H80000015
        picVertScale.Line (-0.01, LowerLimit)-(0.15, LowerLimit)
        picVertScale.DrawWidth = 1
        picVertScale.ForeColor = &HFFFFFF
    End If

End Sub

Public Sub ShowHorizScale()
'draw the horizontal scale markers and numbers
'uses the Picturebox.Line method and the Windows DrawText API
Dim TickX As Single, sglValue As Single
Dim i As Integer, j As Integer

    picHorizScale.Cls
    TickX = 0.167 - StartScaleLeft
    If PageScaleUnits = scEnglish Then
        Do While TickX <= picHorizScale.width
                For j = 1 To 7
                    If TickX + 0.125 * j > 0 Then
                        picHorizScale.Line (TickX + 0.125 * j, 0.06)-(TickX + 0.125 * j, 0.09)
                    End If
                Next j
                If TickX + 0.5 > 0.167 Then picHorizScale.Line (TickX + 0.5, 0.04)-(TickX + 0.5, 0.1)
            recHSclText(i).Top = 0.005 * 1440 / Screen.TwipsPerPixelY
            recHSclText(i).Bottom = 0.15 * 1440 / Screen.TwipsPerPixelY
            If sglValue < 10 Then
                recHSclText(i).Left = (TickX - 0.03) * 1440 / Screen.TwipsPerPixelX
                recHSclText(i).Right = recHSclText(i).Left + 0.2 * 1440 / Screen.TwipsPerPixelX
                If recHSclText(i).Left >= 0.167 Then
                    DrawText picHorizScale.hdc, Format(sglValue, "0"), 1, recHSclText(i), vbLeftJustify
                End If
            Else
                recHSclText(i).Left = (TickX - 0.06) * 1440 / Screen.TwipsPerPixelX
                recHSclText(i).Right = recHSclText(i).Left + 0.25 * 1440 / Screen.TwipsPerPixelX
                If recHSclText(i).Left >= 0.167 Then
                    DrawText picHorizScale.hdc, Format(sglValue, "0"), 2, recHSclText(i), vbLeftJustify
                End If
            End If
            TickX = TickX + 1
            i = i + 1
            sglValue = sglValue + 1
        Loop
    ElseIf PageScaleUnits = scMetric Then
        Do While TickX <= picHorizScale.width
                For j = 1 To 4
                    If TickX + 0.0788 * j > 0 And 0.0788 * j < 0.394 Then
                        picHorizScale.Line (TickX + 0.0788 * j, 0.06)-(TickX + 0.0788 * j, 0.09)
                    End If
                Next j
            recHSclText(i).Top = 0.005 * 1440 / Screen.TwipsPerPixelY
            recHSclText(i).Bottom = 0.15 * 1440 / Screen.TwipsPerPixelY
            If sglValue < 10 Then
                recHSclText(i).Left = (TickX - 0.03) * 1440 / Screen.TwipsPerPixelX
                recHSclText(i).Right = recHSclText(i).Left + 0.2 * 1440 / Screen.TwipsPerPixelX
                If recHSclText(i).Left >= 0.167 Then
                    DrawText picHorizScale.hdc, Format(sglValue, "0"), 1, recHSclText(i), vbLeftJustify
                End If
            Else
                recHSclText(i).Left = (TickX - 0.06) * 1440 / Screen.TwipsPerPixelX
                recHSclText(i).Right = recHSclText(i).Left + 0.25 * 1440 / Screen.TwipsPerPixelX
                If recHSclText(i).Left >= 0.167 Then
                    DrawText picHorizScale.hdc, Format(sglValue, "0"), 2, recHSclText(i), vbLeftJustify
                End If
            End If
            TickX = TickX + 0.394
            i = i + 1
            sglValue = sglValue + 1
        Loop
    End If
    
    PgFreeWid = PageWd - LeftMarg - RightMarg
    Me.picWidMark.Left = (PgFreeWid + 0.167 - StartScaleLeft) - 0.08

End Sub

Private Sub PageHdFt_Click()
'toggles the View Page Header/Footer option
Dim blnValue As Boolean

    blnValue = Not PageHdFt.Checked
    TogglePageHeaderFooter blnValue
    
End Sub

Public Sub TogglePageHeaderFooter(IsOn As Boolean)
'sets the display of the page header/footer sections
Dim i As Integer
Dim blnValue As Boolean

    blnValue = IsOn
    
    cmdDivider(1).Visible = blnValue
    picSection(1).Visible = blnValue
    cmdDivider(9).Visible = blnValue
    picSection(9).Visible = blnValue
    
    PageHdFt.Checked = blnValue
    
    For i = 0 To LastSectionVis
        If picSection(i).Visible Then
            FirstSectionVis = i
            Exit For
        End If
    Next i
    
    If blnValue = True Then
        If LastSectionVis < 5 Then LastSectionVis = 5
    End If
    For i = 10 To 5 Step -1
        If picSection(i).Visible Then
            LastSectionVis = i
            Exit For
        End If
    Next i
    
    GetFirstLastSections
    SetPageComponents
    ShowVertScale
    VScroll.Max = TotalPageHeight() * 8
    VScroll.Visible = TotalPageHeight() > picContainer.Height
    ShowGrid

End Sub

Private Sub PageSet_Click()
'opens the Page Setup dialog

    frmPageSetup.Show vbModal

End Sub

Private Sub PasteClicked()
'user has selected Paste from the context or pull-down menus
'searches through selected control array for items flagged as 'on the clipboard'
'(either visible or not if Cut used) and makes a copy of them in the current section
Dim i As Integer

    ClearSelection
    
    For i = 0 To UBound(SelectedCtl)
        If SelectedCtl(i).blnOnClipBrd Then
            If TypeOf SelectedCtl(i).ctl Is Line Then
                LineIndex = LineIndex + 1
                Load Lin(LineIndex)
                Set Lin(LineIndex).Container = picSection(CurrSection)
                Lin(LineIndex).ZOrder (0)
                Set ctlActive = Lin(LineIndex)
                With Lin(LineIndex)
                    If CurrSection <> SelectedCtl(i).ctl.Tag Then
                        .X1 = SelectedCtl(i).ctl.X1
                        .Y1 = SelectedCtl(i).ctl.Y1
                        .X2 = SelectedCtl(i).ctl.X2
                        .Y2 = SelectedCtl(i).ctl.Y2
                    Else
                        If blnSnapOn Then
                            .X1 = SnapToGrid(SelectedCtl(i).ctl.X1 + 0.1)
                            .Y1 = SnapToGrid(SelectedCtl(i).ctl.Y1 + 0.1)
                            .X2 = SnapToGrid(SelectedCtl(i).ctl.X2 + 0.1)
                            .Y2 = SnapToGrid(SelectedCtl(i).ctl.Y2 + 0.1)
                        Else
                            .X1 = SelectedCtl(i).ctl.X1 + 0.1
                            .Y1 = SelectedCtl(i).ctl.Y1 + 0.1
                            .X2 = SelectedCtl(i).ctl.X2 + 0.1
                            .Y2 = SelectedCtl(i).ctl.Y2 + 0.1
                        End If
                    End If
                    .BorderColor = SelectedCtl(i).ctl.BorderColor
                    .BorderWidth = SelectedCtl(i).ctl.BorderWidth
                    .BorderStyle = SelectedCtl(i).ctl.BorderStyle
                    .Tag = CurrSection
                    .Visible = True
                End With
                If picSection(CurrSection).Height < Lin(LineIndex).Y1 Then
                    picSection(CurrSection).Height = Lin(LineIndex).Y1
                    SetPageComponents
                ElseIf picSection(CurrSection).Height < Lin(LineIndex).Y2 Then
                    picSection(CurrSection).Height = Lin(LineIndex).Y2
                    SetPageComponents
                End If
            Else
                If TypeOf SelectedCtl(i).ctl Is Shape Then
                    ShapeIndex = ShapeIndex + 1
                    Load Shape(ShapeIndex)
                    Set Shape(ShapeIndex).Container = picSection(CurrSection)
                    Shape(ShapeIndex).ZOrder (0)
                    Set ctlActive = Shape(ShapeIndex)
                    With Shape(ShapeIndex)
                        .width = SelectedCtl(i).ctl.width
                        .Height = SelectedCtl(i).ctl.Height
                        .BorderColor = SelectedCtl(i).ctl.BorderColor
                        .BorderWidth = SelectedCtl(i).ctl.BorderWidth
                        .BackStyle = SelectedCtl(i).ctl.BackStyle
                        .BackColor = SelectedCtl(i).ctl.BackColor
                        .Shape = SelectedCtl(i).ctl.Shape
                        .Tag = CurrSection
                        .Visible = True
                    End With
                ElseIf TypeOf SelectedCtl(i).ctl Is Label Then
                    If SelectedCtl(i).ctl.Name = "Label" Then
                        LabelIndex = LabelIndex + 1
                        Load Label(LabelIndex)
                        Set Label(LabelIndex).Container = picSection(CurrSection)
                        Label(LabelIndex).ZOrder (0)
                        Set ctlActive = Label(LabelIndex)
                    ElseIf SelectedCtl(i).ctl.Name = "Field" Then
                        FieldIndex = FieldIndex + 1
                        Load Field(FieldIndex)
                        Set Field(FieldIndex).Container = picSection(CurrSection)
                        Field(FieldIndex).ZOrder (0)
                        Set ctlActive = Field(FieldIndex)
                    ElseIf SelectedCtl(i).ctl.Name = "BoundImg" Then
                        BoundImageIndex = BoundImageIndex + 1
                        Load BoundImg(BoundImageIndex)
                        Set BoundImg(BoundImageIndex).Container = picSection(CurrSection)
                        BoundImg(BoundImageIndex).ZOrder (0)
                        Set ctlActive = BoundImg(BoundImageIndex)
                    End If
                    With ctlActive
                        .FontName = SelectedCtl(i).ctl.FontName
                        .FontSize = SelectedCtl(i).ctl.FontSize
                        .FontBold = SelectedCtl(i).ctl.FontBold
                        .FontItalic = SelectedCtl(i).ctl.FontItalic
                        .FontUnderline = SelectedCtl(i).ctl.FontUnderline
                        .Alignment = SelectedCtl(i).ctl.Alignment
                        .BorderStyle = SelectedCtl(i).ctl.BorderStyle
                        .BackColor = SelectedCtl(i).ctl.BackColor
                        .ForeColor = SelectedCtl(i).ctl.ForeColor
                        .Tag = CurrSection
                        .DataField = SelectedCtl(i).ctl.DataField
                        .LinkTimeout = SelectedCtl(i).ctl.LinkTimeout
                        .width = SelectedCtl(i).ctl.width
                        .Height = SelectedCtl(i).ctl.Height
                        .Caption = SelectedCtl(i).ctl.Caption
                        .Visible = True
                    End With
                ElseIf TypeOf SelectedCtl(i).ctl Is MSForms.Image Then
                    ImageIndex = ImageIndex + 1
                    Load Img(ImageIndex)
                    Set Img(ImageIndex).Container = picSection(CurrSection)
                    Img(ImageIndex).ZOrder (0)
                    Set ctlActive = Img(ImageIndex)
                    With Img(ImageIndex)
                        .width = SelectedCtl(i).ctl.width
                        .Height = SelectedCtl(i).ctl.Height
                        .BorderStyle = SelectedCtl(i).ctl.BorderStyle
                        .BackColor = SelectedCtl(i).ctl.BackColor
                        .Picture = SelectedCtl(i).ctl.Picture
                        .Tag = CurrSection
                        .Visible = True
                    End With
                ElseIf TypeOf SelectedCtl(i).ctl Is CheckBoxControl Then
                    CheckShapeIndex = CheckShapeIndex + 1
                    Load Chkbox(CheckShapeIndex)
                    Set Chkbox(CheckShapeIndex).Container = picSection(CurrSection)
                    Chkbox(CheckShapeIndex).ZOrder (0)
                    Set ctlActive = Chkbox(CheckShapeIndex)
                    With Chkbox(CheckShapeIndex)
                        .BorderColor = SelectedCtl(i).ctl.BorderColor
                        .BackColor = SelectedCtl(i).ctl.BackColor
                        .Tag = CurrSection
                        .DisplayType = SelectedCtl(i).ctl.DisplayType
                        .Visible = True
                    End With
                End If
                With ctlActive
                    If blnControlSelected Then
                        .Left = StartX
                        .Top = StartY
                    Else
                        If CurrSection <> SelectedCtl(i).ctl.Tag Then
                            .Left = SelectedCtl(i).ctl.Left
                            .Top = SelectedCtl(i).ctl.Top
                        Else
                            If blnSnapOn Then
                                .Left = SnapToGrid(SelectedCtl(i).ctl.Left + 0.1)
                                .Top = SnapToGrid(SelectedCtl(i).ctl.Top + 0.1)
                            Else
                                .Left = SelectedCtl(i).ctl.Left + 0.1
                                .Top = SelectedCtl(i).ctl.Top + 0.1
                            End If
                        End If
                    End If
                End With
            End If
        End If
        AddToDrawList
    Next i
    ShowGrid
    ResetMinSectionHt CurrSection
    ResetMinPageWidth
    
    CleanOutOldControls     'get rid of any leftover controls from cut operation
        
        
End Sub
Private Sub CleanOutOldControls()
'clean out any controls made invisible by cut operations
Dim i As Integer, j As Integer
Dim blnPreserve As Boolean

    For i = Me.Controls.count - 1 To 0 Step -1
        If Me.Controls(i).Tag <> "" Then
            If Me.Controls(i).Visible = False Then
                If Me.Controls(i).Name <> "Grip" And Me.Controls(i).Name <> "SelectBox" Then
                    blnPreserve = False
                    For j = 0 To UBound(SelectedCtl)
                        If SelectedCtl(j).ctl.Name = Me.Controls(i).Name _
                            And SelectedCtl(j).ctl.Index = Me.Controls(i).Index Then
                            blnPreserve = True
                        End If
                    Next j
                    If Not blnPreserve Then Unload Me.Controls(i)
                End If
            End If
        End If
    Next i

End Sub

Private Sub PasteObj_Click()

    PasteClicked

End Sub

Private Sub picContainer_Click()
'de-select everything if user clicks on workspace area

    PropertySelector1.Visible = False
    blnGroupSelected = False
    blnControlSelected = False
    ClearSelection

End Sub

Private Sub SetFillClr()
'set fill colour of any selected controls to current active fill colour
Dim i As Integer

    If blnControlSelected Then
        If TypeOf ctlActive Is Shape Then
            ctlActive.BackColor = ActiveBack.Color
            ctlActive.BackStyle = ActiveBack.Style
        ElseIf TypeOf ctlActive Is Label Then
            ctlActive.BackColor = ActiveBack.Color
        End If
    End If

    If blnGroupSelected Then
        For i = 0 To UBound(SelectedCtl)
            If TypeOf SelectedCtl(i).ctl Is Shape Then
                SelectedCtl(i).ctl.BackColor = ActiveBack.Color
                SelectedCtl(i).ctl.BackStyle = ActiveBack.Style
            ElseIf TypeOf SelectedCtl(i).ctl Is Label Then
                SelectedCtl(i).ctl.BackColor = ActiveBack.Color
            End If
        Next i
    End If

End Sub

Private Sub picHorizScale_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        lngState = Default
        lngCmdState = Default
        optAction(0).value = True
        ShowStateMouse
    ElseIf KeyCode = vbKeyDelete Then
        DeleteSelectedCtlObject
    End If

End Sub

Private Sub picHorizScale_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'clear any current control selections and setup for drag select in horizontal scale bar
    
    PropertySelector1.Visible = False
    ClearSelection
    lngState = Default
    lngCmdState = Default
    optAction(0).value = True
    ShowStateMouse
    ScaleStartX = x
    picHorizScale.AutoRedraw = False

End Sub

Private Sub picHorizScale_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'execute drag select in the horizontal scale bar
'hilite drag area in the scale bar (using SetPixelV) and display vertical
'selection border bars at start and end of area (bars are frame controls)
'there are some graphic artifacts when reversing drag direction - haven't figured out how
'to avoid these
Dim i As Long, j As Long
Dim GetColor As Long, LongX As Long
Dim StepDir As Integer
Dim BackCol As Long

    LongX = CLng(x * 1440 / Screen.TwipsPerPixelX)
    If LongX >= ScaleLastX Then
        StepDir = 1
    Else
        StepDir = -1
    End If
    
    OleTranslateColor picHorizScale.BackColor, 0, BackCol
    
    If Button = 1 Then
        For i = 0 To 19
            For j = ScaleLastX To LongX Step StepDir
                GetColor = GetPixel(picHorizScale.hdc, j, i)
                If GetColor = BackCol Then
                    SetPixelV picHorizScale.hdc, j, i, vbBlack
                ElseIf GetColor = vbBlack Then
                    SetPixelV picHorizScale.hdc, j, i, BackCol
                End If
            Next j
        Next i
        lnSelect(0).Left = x - 0.02
        lnSelect(0).Top = 0
        lnSelect(0).width = 0.04
        lnSelect(0).Height = picContainer.Height
        lnSelect(0).Visible = True
        
        lnSelect(1).Left = ScaleStartX - 0.02
        lnSelect(1).Top = 0
        lnSelect(1).width = 0.04
        lnSelect(1).Height = picContainer.Height
        lnSelect(1).Visible = True
    End If
            
    ScaleLastX = LongX

End Sub

Private Sub picHorizScale_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'upon release of mouse button, refresh scale bar, hide selection border bars
'and highlight selected controls if there are any
    picHorizScale.AutoRedraw = True
    ShowHorizScale
    lnSelect(0).Visible = False
    lnSelect(1).Visible = False
    ScaleEndX = x
    HScaleSelectObjects
    If NumInGrp > -1 Then
        HiliteObjects
    End If

End Sub

Private Sub SetLineClr()
'set line colour of any selected controls to current active line colour
Dim i As Integer

    If blnControlSelected Then
        If TypeOf ctlActive Is Line Or TypeOf ctlActive Is Shape Then
            ctlActive.BorderColor = ActiveBorder.Color
        End If
    End If

    If blnGroupSelected Then
        For i = 0 To UBound(SelectedCtl)
            If TypeOf SelectedCtl(i).ctl Is Line Or TypeOf SelectedCtl(i).ctl Is Shape Then
                SelectedCtl(i).ctl.BorderColor = ActiveBorder.Color
            End If
        Next i
    End If

End Sub

Private Sub picContainer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Me.MousePointer = 0

End Sub

Public Sub DeleteSelectedCtlObject()
'delete selected controls
Dim i As Integer, j As Integer, k As Integer
Dim blnPreserve As Boolean

'if a group of controls selected, loop through array and delete
'unless control is flagged as 'on clipboard'
    If blnGroupSelected Then
        For i = 0 To UBound(SelectedCtl)
            For j = 0 To UBound(DrawList)
                If DrawList(j).ctlName = SelectedCtl(i).ctl.Name And DrawList(j).ctlIndex = SelectedCtl(i).ctl.Index Then
                    For k = j + 1 To UBound(DrawList)
                        DrawList(k - 1).ctlName = DrawList(k).ctlName
                        DrawList(k - 1).ctlIndex = DrawList(k).ctlIndex
                    Next k
                End If
            Next j
            dLstCount = dLstCount - 1
            SelectedCtl(i).ctl.Visible = False
            If Not SelectedCtl(i).blnOnClipBrd Then
                Unload SelectedCtl(i).ctl
            End If
        Next i
        If dLstCount >= 0 Then
            ReDim Preserve DrawList(dLstCount)
        End If
        blnGroupSelected = False
'if single control selected, delete unless flagged as 'on clipboard'
    ElseIf blnControlSelected Then
        If blnSelectArrayInit Then
            For i = 0 To UBound(SelectedCtl)
                If SelectedCtl(i).blnOnClipBrd Then
                    If SelectedCtl(i).ctl.Name = ctlActive.Name And SelectedCtl(i).ctl.Index = ctlActive.Index Then
                        blnPreserve = True
                        Exit For
                    End If
                End If
            Next i
        End If
        For i = 0 To Me.Controls.count - 1
            If Me.Controls(i).Name = ctlActive.Name Then
                If Me.Controls(i).Index = ctlActive.Index Then
                    For j = 0 To UBound(DrawList)
                        If DrawList(j).ctlName = ctlActive.Name And DrawList(j).ctlIndex = ctlActive.Index Then
                            For k = j + 1 To UBound(DrawList)
                                DrawList(k - 1).ctlName = DrawList(k).ctlName
                                DrawList(k - 1).ctlIndex = DrawList(k).ctlIndex
                            Next k
                        End If
                    Next j
                    dLstCount = dLstCount - 1
                    ReDim Preserve DrawList(dLstCount)
                    If Not blnPreserve Then
                        WriteToUndoList unControl, unDelete
                        Unload ctlActive
                    End If
                    blnControlSelected = False
                    Exit For
                End If
            End If
        Next i
    End If

'reset display and command state
    blnReportSaved = False
    ClearSelection
    ShowGrid
    For i = FirstSectionVis To LastSectionVis
        ResetMinSectionHt i
    Next i
    ResetMinPageWidth
    For i = 0 To 10
        optAction(i).value = False
    Next i
    blnMouseIsDown = False
    lngState = Default

End Sub

Private Sub picSection_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'handle applicable key down strokes while page section has focus
    
    If KeyCode = vbKeyEscape Then
        lngState = Default
        lngCmdState = Default
        optAction(0).value = True
        ShowStateMouse
    End If

End Sub

Private Sub picSection_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'handle applicable key up strokes while page section has focus
Dim CtrlDown

    If KeyCode = vbKeyDelete Then
        DeleteSelectedCtlObject
    End If

    CtrlDown = (Shift And vbCtrlMask) > 0
    If CtrlDown Then
        If KeyCode = vbKeyC Then
            CopyClicked
        ElseIf KeyCode = vbKeyX Then
            CutClicked
        ElseIf KeyCode = vbKeyV Then
            PasteClicked
        End If
    End If

End Sub

Private Sub picSection_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'handle mouse down event on page section which has focus
Dim i As Integer

'set colour of selected divider over section and in vertical scale bar
    SetDividerColours Index
    For i = FirstSectionVis To LastSectionVis
        If picSection(i).Visible Then
            DoScaleDividers i
        End If
    Next i

    If CurrSection <> Index Then
        CurrSection = Index
    End If
    
'store the current mouse coordinates for use elsewhere
    StartX = x
    StartY = y
    
'if right mouse button pressed, popup menu will be displayed at mouse up event
    If Button = 2 Then Exit Sub
    
    frmSelField.Visible = False
    blnMouseIsDown = True
'depending on current command state, place a check box, set up for moving a control,
'set up for resizing a control
    If lngState = PlaceNewControl Then
        If intControlType = cCheckBox Then CreateCheckBox Index
        ResetMinPageWidth
        ResetMinSectionHt Index
    ElseIf lngState = MoveGrip Then
        WriteToUndoList unControl, unResize
        If blnGripFound Then
            Set ctlActive = ctlTest
            GetRelativeDistances x, y
            dX1 = StartX
            dY1 = StartY
        End If
        ResetMinPageWidth
        ResetMinSectionHt Index
        Exit Sub
    ElseIf lngState = OverControl Then
        blnDragStarted = False
        lngState = MoveControl
    Else
        lngState = SelectGroup
        Set SelectBox.Container = picSection(Index)
    End If

    ClearSelection
    blnControlSelected = False
    
'if a control is found at mouse coordinates, add to selection if shift key is down
'and get relative coordinates to mouse to prepare for moving
    If blnControlFound Then
        If Shift = 1 Then
            If NumInGrp = -1 Then
                NumInGrp = NumInGrp + 1
                ReDim Preserve SelectedCtl(NumInGrp)
                Set SelectedCtl(NumInGrp).ctl = ctlActive
            End If
            NumInGrp = NumInGrp + 1
            ReDim Preserve SelectedCtl(NumInGrp)
            Set SelectedCtl(NumInGrp).ctl = ctlTest
            blnGroupSelected = True
        End If
        Set ctlActive = ctlTest
        GetRelativeDistances x, y
        If Not blnGroupSelected Then
            blnControlSelected = True
            ShowGrips Index
        End If
    Else
        NumInGrp = -1
        blnGroupSelected = False
    End If
    
    If NumInGrp > -1 Then
        HiliteObjects
    End If
    
'check whether mouse is on last visible page section and whether is at bottom edge or
'right edge in preparation for resizing section height or page width
    blnOnLastSection = False
    If y > picSection(LastSectionVis).Height - 0.08 And Index = LastSectionVis Then
        If Not lngState = MoveGrip Then
            blnDragStarted = False
            blnOnLastSection = True
            lngState = ResizeSection
            Me.lnPos.Y1 = picSection(LastSectionVis).Top + picSection(LastSectionVis).Height - 0.17
            Me.lnPos.Y2 = Me.lnPos.Y1
            Me.lnPos.Visible = True
        End If
    ElseIf x > picSection(Index).width - 0.08 Then
        If Not lngState = MoveGrip Then
            lngState = ResizePageWidth
        End If
    End If
    
    ShowStateMouse

End Sub

Private Sub picSection_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'handles mouse moving in page section - calls ExecuteState if left mouse button is down
'otherwise updates display, detects shapes and lines, sets program command state
Dim i As Integer
Dim ScaleFactor As Single, strUnits As String

    If blnMouseIsDown Then
        If Button = 1 Then
            ExecuteState Index, lngState, x, y
        ElseIf Button = 2 Then
            Exit Sub
        End If
    Else
        For i = 0 To 7
            If Grip(i).Visible And CurrSection = Index Then
                If (x > Grip(i).Left And x < Grip(i).Left + Grip(i).width) And _
                    (y > Grip(i).Top And y < Grip(i).Top + Grip(i).Height) Then
                    ActiveGrip = i
                    lngState = MoveGrip
                    ShowStateMouse
                    Exit Sub
                End If
            End If
        Next i
        If y > picSection(LastSectionVis).Height - 0.08 And Index = LastSectionVis Then
            If Not lngState = MoveGrip Then
                lngState = ResizeSection
            End If
        ElseIf x > picSection(Index).width - 0.08 Then
            If Not lngState = MoveGrip Then
                lngState = ResizePageWidth
            End If
        Else
            lngState = lngCmdState
        End If
        
        If lngState <> PlaceNewControl And lngState <> MoveGrip Then
            DetectControls Index, x, y
        End If
    End If

    If PageScaleUnits = scEnglish Then
        ScaleFactor = 1
        strUnits = "   (IN)"
    ElseIf PageScaleUnits = scMetric Then
        ScaleFactor = 2.54
        strUnits = "   (CM)"
    End If
    
    StatusBar1.Panels(1).text = Format(x * ScaleFactor, "0.000") & " , " & Format(y * ScaleFactor, "0.000") & strUnits
    StatusBar1.Panels(3).text = ""
    ShowStateMouse
    
End Sub

Private Sub picSection_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'handles mouse up event on page section which has focus
Dim i As Integer

'if the Open File dialog was open, just exit (to avoid problems if user double-clicks
'file in the dialog)
    If blnDialogOpen Then
        blnDialogOpen = False
        Exit Sub
    End If
    
    blnMouseIsDown = False

'if right mouse button pressed, show the appropriate popup menu
'whether a control is selected or not
    If Button = 2 Then
        If blnControlSelected Then
            If TypeOf ctlActive Is Shape Then
                ShowPopUp cBox
            ElseIf TypeOf ctlActive Is Line Then
                ShowPopUp cLine
            ElseIf TypeOf ctlActive Is Label Then
                If ctlActive.Name = "Label" Then
                    ShowPopUp cLabel
                ElseIf ctlActive.Name = "Field" Then
                    ShowPopUp cDataField
                ElseIf ctlActive.Name = "BoundImg" Then
                    ShowPopUp cBoundImage
                End If
            ElseIf TypeOf ctlActive Is MSForms.Image Then
                ShowPopUp cImage
            ElseIf TypeOf ctlActive Is CheckBoxControl Then
                ShowPopUp cCheckBox
            End If
        Else
            ShowPopUp cNone
        End If
        Exit Sub
    End If
    
'if left mouse button pressed
'if a control has been selected show the selection grips
    If blnControlSelected Then
        If CurrSection = Index Then
            ShowGrips Index
        End If
    End If

'and take the appropriate action depending on the current command state
    If lngState = PlaceNewControl Then
        If intControlType = cCheckBox Then
            SetDataField
        End If
    ElseIf lngState = MoveGrip Then
        ResetMinPageWidth
        ResetMinSectionHt Index
        HiliteObjects
        GetTextFormat
        
        If blnCreatingControl Then
            blnCreatingControl = False
            If intControlType = cImage Then
                SelectImage
            Else
                WriteToUndoList unControl, unPlace
            End If
            If intControlType = cDataField Then
                If ctlActive.LinkTimeout = cDataField Then
                    SetDataField
                End If
            ElseIf intControlType = cBoundImage Then
                SetDataField
            End If
        End If
        
    ElseIf lngState = ResizeSection Then
        If TotalPageHeight() < picContainer.Height Then
            Me.lnPos.Visible = False
            SetPageComponents
            StartScaleTop = 0
            ShowVertScale
        End If
        ShowGrid Index
    ElseIf lngState = ResizePageWidth Then
        ShowGrid
        If picSection(Index).width > PgFreeWid Then
            MsgBox "You have exceeded the available space between the left and right margin." & Chr$(13) _
            & "This may cause some report items to be truncated.", vbOKOnly + vbExclamation, "Page Width Exceeded"
        End If
        If picSection(Index).width + 0.156 < picContainer.width Then
            For i = 0 To 10
                picSection(i).Left = 0.156
            Next i
            StartScaleLeft = 0
            ShowHorizScale
        End If
    ElseIf lngState = SelectGroup Then
        If SelectBox.Visible Then
            SelectBox.Visible = False
            SelectObjects Index
            If NumInGrp > -1 Then
                HiliteObjects
            End If
        End If
        lngState = Default
    ElseIf lngState = MoveControl Then
        ResetMinSectionHt Index
        ResetMinPageWidth
        lngState = Default
        GetTextFormat
        If blnGroupSelected Then
            ShowGrid
            HiliteObjects
        End If
    End If
    
    If PropertySelector1.Visible Then PropertySelector1.Visible = False
    blnDragStarted = False
    ShowStateMouse
    
End Sub

Private Sub DisplayBoxIfNoBorder(picSectionNo As Integer)
'if label or shape has transparent border, display the selection box when they are selected

    picSection(picSectionNo).Cls
    If blnGridOn Then ShowGrid
    If blnControlSelected Then
        If TypeOf ctlActive Is Label Or TypeOf ctlActive Is Shape Then
            If ctlActive.BorderStyle = 0 Then
                picSection(picSectionNo).DrawStyle = 2
                picSection(picSectionNo).DrawWidth = 1
                picSection(picSectionNo).Line (ctlActive.Left - 0.02, ctlActive.Top - 0.02)-(ctlActive.Left + ctlActive.width + 0.02, ctlActive.Top + ctlActive.Height + 0.02), , B
            End If
        End If
    End If

End Sub

Private Sub DetectControls(GetSectionNo As Integer, GetX As Single, GetY As Single)
'detects line and shape controls which have no windows handle and thus will not be
'automatically detected by mouse

Dim i As Integer, j As Integer
Dim ctlLast As Control

    blnControlFound = False
    blnGripFound = False
    If dLstCount > 0 Then
    If lngState <> MoveGrip Then
        For i = 0 To UBound(DrawList)
            For j = 0 To Me.Controls.count - 1
                Set ctlTest = Me.Controls(j)
                If ctlTest.Tag = GetSectionNo Then
                    If ctlTest.Name = DrawList(i).ctlName And ctlTest.Index = DrawList(i).ctlIndex Then
                        If TypeOf ctlTest Is Line Then
                            If Abs(ctlTest.X1 - ctlTest.X2) < 0.05 Then
                                If Abs(GetX - ctlTest.X1) < 0.02 Then
                                    If (GetY < ctlTest.Y2 And GetY > ctlTest.Y1) Or (GetY > ctlTest.Y2 And GetY < ctlTest.Y1) Then
                                        Set ctlLast = ctlTest
                                        blnControlFound = True
                                    End If
                                End If
                            ElseIf Abs(ctlTest.Y1 - ctlTest.Y2) < 0.05 Then
                                If Abs(GetY - ctlTest.Y1) < 0.02 Then
                                    If (GetX < ctlTest.X2 And GetX > ctlTest.X1) Or (GetX > ctlTest.X2 And GetX < ctlTest.X1) Then
                                        Set ctlLast = ctlTest
                                        blnControlFound = True
                                    End If
                                End If
                            Else
                                If (GetX > ctlTest.X1 And GetX < ctlTest.X2) Or (GetX > ctlTest.X2 And GetX < ctlTest.X1) And _
                                    (GetY < ctlTest.Y2 And GetY > ctlTest.Y1) Or (GetY > ctlTest.Y2 And GetY < ctlTest.Y1) Then
                                    Result1 = (GetX - ctlTest.X1) / IIf((GetY - ctlTest.Y1) = 0, 0.001, (GetY - ctlTest.Y1))
                                    Result2 = (ctlTest.X2 - ctlTest.X1) / (ctlTest.Y2 - ctlTest.Y1)
                                    Result2 = Result1 / Result2
                                    If Result2 > 0.9 And Result2 < 1.1 Then
                                        Set ctlLast = ctlTest
                                        blnControlFound = True
                                    End If
                                End If
                            End If
                            If blnGroupSelected Then
                                If (GetX > ctlTest.X1 - 0.03 And GetX < ctlTest.X1 + 0.03) And _
                                (GetY > ctlTest.Y1 - 0.03 And GetY < ctlTest.Y1 + 0.03) Then
                                    ActiveGrip = 0
                                    blnGripFound = True
                                    Set ctlLast = ctlTest
                                    Exit For
                                ElseIf (GetX > ctlTest.X2 - 0.03 And GetX < ctlTest.X2 + 0.03) And _
                                (GetY > ctlTest.Y2 - 0.03 And GetY < ctlTest.Y2 + 0.03) Then
                                    ActiveGrip = 1
                                    blnGripFound = True
                                    Set ctlLast = ctlTest
                                    Exit For
                                End If
                            End If
                        Else
                            If TypeOf ctlTest Is Shape Then
                                If ctlTest.Name <> "Grip" Then
                                    If ctlTest.BackStyle = 0 Then
                                        If Abs(GetX - ctlTest.Left) < 0.03 Or Abs(GetX - (ctlTest.Left + ctlTest.width)) < 0.03 Then
                                            If GetY > ctlTest.Top And GetY < (ctlTest.Top + ctlTest.Height) Then
                                                Set ctlLast = ctlTest
                                                blnControlFound = True
                                            End If
                                        ElseIf Abs(GetY - ctlTest.Top) < 0.03 Or Abs(GetY - (ctlTest.Top + ctlTest.Height)) < 0.03 Then
                                            If GetX > ctlTest.Left And GetX < (ctlTest.Left + ctlTest.width) Then
                                                Set ctlLast = ctlTest
                                                blnControlFound = True
                                            End If
                                        End If
                                    Else
                                        If (GetX > ctlTest.Left And GetX < ctlTest.Left + ctlTest.width) And _
                                            (GetY > ctlTest.Top And GetY < ctlTest.Top + ctlTest.Height) Then
                                            Set ctlLast = ctlTest
                                            blnControlFound = True
                                        End If
                                    End If
                                End If
                            End If
                            If blnGroupSelected Then
                                With ctlTest
                                    If GetX > .Left - 0.05 And GetX < .Left - 0.01 Then
                                        If GetY > .Top - 0.05 And GetY < .Top - 0.01 Then
                                            ActiveGrip = 0
                                            blnGripFound = True
                                            Set ctlLast = ctlTest
                                            Exit For
                                        ElseIf GetY < .Top + .Height + 0.05 And GetY > .Top + .Height + 0.01 Then
                                            ActiveGrip = 3
                                            blnGripFound = True
                                            Set ctlLast = ctlTest
                                            Exit For
                                        ElseIf GetY < .Top + .Height / 2 + 0.03 And GetY > .Top + .Height / 2 - 0.03 Then
                                            ActiveGrip = 7
                                            blnGripFound = True
                                            Set ctlLast = ctlTest
                                            Exit For
                                        End If
                                    ElseIf GetX < .Left + .width + 0.05 And GetX > .Left + .width + 0.01 Then
                                        If GetY > .Top - 0.05 And GetY < .Top - 0.01 Then
                                            ActiveGrip = 1
                                            blnGripFound = True
                                            Set ctlLast = ctlTest
                                            Exit For
                                        ElseIf GetY < .Top + .Height + 0.05 And GetY > .Top + .Height + 0.01 Then
                                            ActiveGrip = 2
                                            blnGripFound = True
                                            Set ctlLast = ctlTest
                                            Exit For
                                        ElseIf GetY < .Top + .Height / 2 + 0.03 And GetY > .Top + .Height / 2 - 0.03 Then
                                            ActiveGrip = 5
                                            blnGripFound = True
                                            Set ctlLast = ctlTest
                                            Exit For
                                        End If
                                    ElseIf GetX < .Left + .width / 2 + 0.03 And GetX > .Left + .width / 2 - 0.03 Then
                                        If GetY > .Top - 0.05 And GetY < .Top - 0.01 Then
                                            ActiveGrip = 4
                                            blnGripFound = True
                                            Set ctlLast = ctlTest
                                            Exit For
                                        ElseIf GetY < .Top + .Height + 0.05 And GetY > .Top + .Height + 0.01 Then
                                            ActiveGrip = 6
                                            blnGripFound = True
                                            Set ctlLast = ctlTest
                                            Exit For
                                        End If
                                    End If
                                End With
                            End If
                        End If
                    End If
                End If
            Next j
        Next i
    End If
    End If

    If blnControlFound Then
        Set ctlTest = ctlLast
        lngPrevState = lngState
        lngState = OverControl
    ElseIf blnGripFound Then
        Set ctlTest = ctlLast
        lngPrevState = lngState
        lngState = MoveGrip
    End If

End Sub

Private Function SnapToGrid(Coord As Single) As Single
'modifies the passed coordinate to the nearest available grid coordinate
'based on the current grid snap setting
Dim InchGrid As Single, Remainder As Single

    InchGrid = GridSpace / 96
    Remainder = Int(Coord / InchGrid)

    Remainder = Remainder * InchGrid
    Remainder = Coord - Remainder

    If Remainder <= InchGrid / 2 Then
        SnapToGrid = Coord - Remainder
    Else
        SnapToGrid = Coord + InchGrid - Remainder
    End If

End Function

Private Sub Field_DblClick(Index As Integer)

    If ctlActive.LinkTimeout = cDataField Then
        SetDataField
    ElseIf ctlActive.LinkTimeout = cCalcField Then
        EditCalc_Click
    ElseIf ctlActive.LinkTimeout = cSumField Then
        EditSummary_Click
    End If

End Sub

Private Sub Field_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    CurrSection = Field(Index).Tag
    blnMouseIsDown = True
    blnDragStarted = False
        
    If Button = vbRightButton Then
        Set ctlActive = Field(Index)
        ShowPopUp ctlActive.LinkTimeout
        Exit Sub
    End If
    
    SetupControlSelection Field(Index), Field(Index).Tag, 1440, x, y
    

End Sub

Private Sub Field_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    MoveInControl x, y, Field(Index).ToolTipText

End Sub

Private Sub Field_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    ConfirmControlSelection Field(Index), Field(Index).Tag

End Sub

Private Sub SetTextClr()
Dim i As Integer

    If blnControlSelected Then
        If TypeOf ctlActive Is Label Then
            ctlActive.ForeColor = ActiveText.Color
        End If
    ElseIf blnGroupSelected Then
        For i = 0 To UBound(SelectedCtl)
            If TypeOf SelectedCtl(i).ctl Is Label Then
                SelectedCtl(i).ctl.ForeColor = ActiveText.Color
            End If
        Next i
    End If


End Sub

Private Sub picVertScale_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        lngState = Default
        lngCmdState = Default
        optAction(0).value = True
        ShowStateMouse
    ElseIf KeyCode = vbKeyDelete Then
        DeleteSelectedCtlObject
    End If

End Sub

Private Sub picVertScale_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    PropertySelector1.Visible = False
    ClearSelection
    lngState = Default
    lngCmdState = Default
    optAction(0).value = True
    ShowStateMouse
    ScaleStartY = y
    picVertScale.AutoRedraw = False
    
End Sub

Private Sub picVertScale_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Integer
Dim GetColor As Long, LongY As Long
Dim StepDir As Integer
Dim BackCol As Long

    LongY = CLng(y * 1440 / Screen.TwipsPerPixelX)
    If LongY >= ScaleLastY Then
        StepDir = 1
    Else
        StepDir = -1
    End If
    
    OleTranslateColor picVertScale.BackColor, 0, BackCol
    
    If Button = 1 Then
        For i = 0 To 19
            For j = ScaleLastY To LongY Step StepDir
                GetColor = GetPixel(picVertScale.hdc, i, j)
                If GetColor = BackCol Then
                    SetPixelV picVertScale.hdc, i, j, vbBlack
                ElseIf GetColor = vbBlack Then
                    SetPixelV picVertScale.hdc, i, j, BackCol
                End If
            Next j
        Next i
        lnSelect(0).Left = 0
        lnSelect(0).Top = y + 0.15
        lnSelect(0).width = picContainer.width
        lnSelect(0).Height = 0.04
        lnSelect(0).Visible = True
        
        lnSelect(1).Left = 0
        lnSelect(1).Top = ScaleStartY + 0.15
        lnSelect(1).width = picContainer.width
        lnSelect(1).Height = 0.04
        lnSelect(1).Visible = True
    End If

    ScaleLastY = LongY

End Sub

Private Sub picVertScale_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    picVertScale.AutoRedraw = True
    ShowVertScale
    lnSelect(0).Visible = False
    lnSelect(1).Visible = False
    ScaleEndY = y
    VScaleSelectObjects
    If NumInGrp > -1 Then
        HiliteObjects
    End If

End Sub

Private Sub prnt_Click()

    If ReportCanPrint Then
        LoadCurrentReport
        PreviewReport False
        Exit Sub
    End If

End Sub

Private Sub PrntPrev_Click()
On Error GoTo NoPreview

    If ReportCanPrint Then
        LoadCurrentReport
        PreviewReport True
    End If
    
    Exit Sub

NoPreview:
    MsgBox Err.Description

End Sub

Private Sub PropertySelector1_Clicked()
Dim i As Integer
Dim GetColor As Long
    
    If PropertySelector1.ShowDialog Then
        PropertySelector1.Visible = False
        PropertySelector1.ShowDialog = False
        CommonDialog1.Flags = cdlCCFullOpen
        On Error Resume Next
        CommonDialog1.ShowColor
        If Err.Number = cdlCancel Then Exit Sub
        GetColor = CommonDialog1.Color
    Else
        GetColor = PropertySelector1.SelectedColor
        blnReportSaved = False
    End If

    
    With PropertySelector1
        If PropertySelectMode = 1 Then
            WriteToUndoList unControl, unFormat
            picTxtCol.BackColor = GetColor
            ActiveText.Color = GetColor
            If blnControlSelected Then
                If TypeOf ctlActive Is Label Then
                    ctlActive.ForeColor = ActiveText.Color
                End If
            ElseIf blnGroupSelected Then
                For i = 0 To UBound(SelectedCtl)
                    If TypeOf SelectedCtl(i).ctl Is Label Then
                        SelectedCtl(i).ctl.ForeColor = ActiveText.Color
                    End If
                Next i
            End If
        ElseIf PropertySelectMode = 2 Then
            WriteToUndoList unControl, unFormat
            If GetColor = -1 Then
                picLinCol.BorderStyle = 1
                picLinCol.BackColor = &H8000000F
                ActiveBorder.Style = 0
                ActiveBorder.Color = vbWhite
            Else
                picLinCol.BorderStyle = 0
                picLinCol.BackColor = GetColor
                ActiveBorder.Style = 1
                ActiveBorder.Color = GetColor
            End If
            If GetColor <> -1 Then
                If blnControlSelected Then
                    If TypeOf ctlActive Is Line Or TypeOf ctlActive Is CheckBoxControl Then
                        ctlActive.BorderColor = ActiveBorder.Color
                    ElseIf TypeOf ctlActive Is Shape Then
                        ctlActive.BorderColor = ActiveBorder.Color
                        If ctlActive.BorderStyle = 0 Then
                            ctlActive.BorderStyle = 1
                        End If
                    End If
                ElseIf blnGroupSelected Then
                    For i = 0 To UBound(SelectedCtl)
                        If TypeOf SelectedCtl(i).ctl Is Line Or TypeOf SelectedCtl(i).ctl Is CheckBoxControl Then
                            SelectedCtl(i).ctl.BorderColor = ActiveBorder.Color
                        ElseIf TypeOf SelectedCtl(i).ctl Is Shape Then
                            SelectedCtl(i).ctl.BorderColor = ActiveBorder.Color
                            If SelectedCtl(i).ctl.BorderStyle = 0 Then
                                SelectedCtl(i).ctl.BorderStyle = 1
                            End If
                        End If
                    Next i
                End If
            Else
                If blnControlSelected Then
                    If TypeOf ctlActive Is Shape Then
                        ctlActive.BorderStyle = 0
                    End If
                ElseIf blnGroupSelected Then
                    For i = 0 To UBound(SelectedCtl)
                        If TypeOf SelectedCtl(i).ctl Is Shape Then
                            SelectedCtl(i).ctl.BorderStyle = 0
                        End If
                    Next i
                End If
            End If
        ElseIf PropertySelectMode = 3 Then
            WriteToUndoList unControl, unFormat
            If GetColor = -1 Then
                picFillCol.BackColor = &H8000000F
                picFillCol.BorderStyle = 1
                ActiveBack.Style = 0
                ActiveBack.Color = vbWhite
            Else
                picFillCol.BorderStyle = 0
                picFillCol.BackColor = GetColor
                ActiveBack.Style = 1
                ActiveBack.Color = GetColor
            End If
            If blnControlSelected Then
                If TypeOf ctlActive Is Shape Or TypeOf ctlActive Is Label Then
                    ctlActive.BackStyle = ActiveBack.Style
                    ctlActive.BackColor = ActiveBack.Color
                ElseIf TypeOf ctlActive Is CheckBoxControl Then
                    ctlActive.BackColor = ActiveBack.Color
                End If
            ElseIf blnGroupSelected Then
                For i = 0 To UBound(SelectedCtl)
                    If TypeOf SelectedCtl(i).ctl Is Shape Or TypeOf SelectedCtl(i).ctl Is Label Then
                        SelectedCtl(i).ctl.BackColor = ActiveBack.Color
                        SelectedCtl(i).ctl.BackStyle = ActiveBack.Style
                    ElseIf TypeOf SelectedCtl(i).ctl Is CheckBoxControl Then
                        SelectedCtl(i).ctl.BackColor = ActiveBack.Color
                    End If
                Next i
            End If
        ElseIf PropertySelectMode = 4 Then
            WriteToUndoList unControl, unFormat
            ActiveBorder.Style = .SelectedLineStyle
            If blnControlSelected Then
                If TypeOf ctlActive Is Line Or TypeOf ctlActive Is Shape Then
                    ctlActive.BorderStyle = ActiveBorder.Style + 1
                End If
            ElseIf blnGroupSelected Then
                For i = 0 To UBound(SelectedCtl)
                    If TypeOf SelectedCtl(i).ctl Is Shape Or TypeOf SelectedCtl(i).ctl Is Line Then
                        SelectedCtl(i).ctl.BorderStyle = ActiveBorder.Style + 1
                    End If
                Next i
            End If
        ElseIf PropertySelectMode = 5 Then
            WriteToUndoList unControl, unFormat
            Dim TextREC As Rect
            ActiveBorder.width = .SelectedLineWidth
            TextREC.Left = 0: TextREC.Right = 20: TextREC.Top = 3: TextREC.Bottom = 20
            If blnControlSelected Then
                If TypeOf ctlActive Is Line Or TypeOf ctlActive Is Shape Then
                    ctlActive.BorderWidth = ActiveBorder.width
                End If
            ElseIf blnGroupSelected Then
                For i = 0 To UBound(SelectedCtl)
                    If TypeOf SelectedCtl(i).ctl Is Line Or TypeOf SelectedCtl(i).ctl Is Shape Then
                        SelectedCtl(i).ctl.BorderWidth = ActiveBorder.width
                    End If
                Next i
            End If
        ElseIf PropertySelectMode = 6 Then
            WriteToUndoList unSection, unSectColor
            If GetColor > -1 Then
                picSection(CurrSection).BackColor = GetColor
                picSection(CurrSection).Refresh
                ShowGrid CurrSection
            End If
        End If
        .Visible = False
    End With


ExitSub:

End Sub

Private Sub DrawPicBorder(ctlPic As PictureBox)

    ctlPic.Cls
    ctlPic.DrawStyle = 0
    ctlPic.ForeColor = RGB(100, 100, 100)
'    ctlPic.ForeColor = vbWhite
    ctlPic.Line (ctlPic.ScaleWidth - 1, 0)-(ctlPic.ScaleWidth - 1, ctlPic.ScaleHeight - 1)
    ctlPic.Line (ctlPic.ScaleWidth - 1, ctlPic.ScaleHeight - 1)-(0, ctlPic.ScaleHeight - 1)
    ctlPic.Line (0, ctlPic.ScaleHeight - 1)-(0, 0)
    ctlPic.Line (0, 0)-(ctlPic.ScaleWidth - 1, 0)
    ctlPic.ForeColor = vbBlack

End Sub

Private Sub PropertySelector1_LostFocus()

    PropertySelector1.Visible = False

End Sub

Private Sub Rect_Click()

    ActiveShape = 0
    UpdateShape

End Sub

Private Sub RepHdFt_Click()
Dim blnValue As Boolean

    blnValue = Not RepHdFt.Checked
    ToggleRepHeaderFooter blnValue
    
End Sub

Public Sub ToggleRepHeaderFooter(IsOn As Boolean)
Dim blnValue As Boolean
Dim i As Integer

    blnValue = IsOn
    
    cmdDivider(0).Visible = blnValue
    picSection(0).Visible = blnValue
    cmdDivider(10).Visible = blnValue
    picSection(10).Visible = blnValue
    RepHdFt.Checked = blnValue
    
    If blnValue = True Then
        FirstSectionVis = 0
        LastSectionVis = 10
    Else
        For i = 0 To LastSectionVis
            If picSection(i).Visible Then
                FirstSectionVis = i
                Exit For
            End If
        Next i
        
        For i = 10 To 5 Step -1
            If picSection(i).Visible Then
                LastSectionVis = i
                Exit For
            End If
        Next i
    End If
    
    GetFirstLastSections
    SetPageComponents
    ShowVertScale
    VScroll.Max = TotalPageHeight() * 8
    VScroll.Visible = TotalPageHeight() > picContainer.Height
    ShowGrid

End Sub

Private Sub RRect_Click()

    WriteToUndoList unControl, unFormat
    ActiveShape = 4
    UpdateShape

End Sub

Private Sub RSqr_Click()

    WriteToUndoList unControl, unFormat
    ActiveShape = 5
    UpdateShape

End Sub

Private Sub SaveFile_Click()

    SaveReportDesign False

End Sub

Private Sub SaveFileAs_Click()

    SaveReportDesign True

End Sub

Private Sub SendBack_Click()
Dim i As Integer
Dim SaveName As String, SaveIndex As Integer, SaveLoc As Integer

    ctlActive.ZOrder (1)
    SaveName = ctlActive.Name
    SaveIndex = ctlActive.Index

    For i = 0 To UBound(DrawList)
        If DrawList(i).ctlName = SaveName And DrawList(i).ctlIndex = SaveIndex Then
            SaveLoc = i
            Exit For
        End If
    Next i

    For i = SaveLoc To 1 Step -1
        DrawList(i).ctlName = DrawList(i - 1).ctlName
        DrawList(i).ctlIndex = DrawList(i - 1).ctlIndex
    Next i

    DrawList(0).ctlName = SaveName
    DrawList(0).ctlIndex = SaveIndex
    
End Sub

Private Sub SetBackColor_Click()

    
    With PropertySelector1
        .TransparentOption = False
        .SelectMode = 6
        If StartX + .width < Me.picContainer.width Then
            .Left = StartX + picSection(CurrSection).Left + 0.75
        Else
            .Left = StartX + picSection(CurrSection).Left - 1.3
        End If
        If StartY + picSection(CurrSection).Top + .Height < Me.picContainer.Height Then
            .Top = StartY + picSection(CurrSection).Top + 0.75
        Else
            .Top = StartY + picSection(CurrSection).Top - 0.9
        End If
        .Visible = True
    End With

    PropertySelectMode = 6

End Sub

Private Sub Sqre_Click()

    WriteToUndoList unControl, unFormat
    ActiveShape = 1
    UpdateShape

End Sub

Private Sub ToggleSnap_Click()

    blnSnapOn = Not blnSnapOn
    ToggleSnap.Checked = blnSnapOn

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer

    Select Case Button.Key
        Case "New": OpenNewFile
        Case "Open": OpenSavedFile
        Case "Save": SaveReportDesign False
        Case "Print": PrintTheReport
        Case "Preview": PreviewTheReport
        Case "Copy": CopyClicked
        Case "Cut": CutClicked
        Case "Paste": PasteClicked
        Case "Undo"
            If blnFirstUndo Then
                WriteToUndoList unControl, LastType
            End If
            RestoreFromUndoList
            For i = 0 To 10
                ResetMinSectionHt i
            Next i
            ResetMinPageWidth
            If UndoList(CurrUndoPos).Type = unSectHeight Then
                SetPageComponents
                ShowVertScale
            End If
            ClearSelection
        Case "DesignGrid": GridSpacing_Click
        Case "DataConnect": DataConn_Click
        Case "ViewData": frmViewData.Show vbModal
    End Select

End Sub

Private Sub SetPropertyPanelPos()

    With PropertySelector1
        If CoolBar1.Height < 0.3 Then
            .Left = CoolBar1.Bands(1).width + ToolX
        Else
            .Left = ToolX
        End If
        .Top = TopOfDesForm
    End With

End Sub

Private Sub ShowLineStyleList()

    With PropertySelector1
        .TransparentOption = False
        SetPropertyPanelPos
        If PropertySelectMode <> 4 Then
            .SelectMode = 4
            .Visible = True
        Else
            .Visible = Not .Visible
        End If
    End With

    PropertySelectMode = 4

End Sub

Private Sub ShowLineWidthList()

    With PropertySelector1
        .TransparentOption = False
        SetPropertyPanelPos
        If PropertySelectMode <> 5 Then
            .SelectMode = 5
            .Visible = True
        Else
            .Visible = Not .Visible
        End If
    End With

    PropertySelectMode = 5

End Sub

Private Sub Toolbar1_Click()

    PropertySelector1.Visible = False

End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Me.MousePointer = 0

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "Bold": BoldClicked
        Case "Italic": ItalicClicked
        Case "Underline": UnderlineClicked
        Case "LeftAlign": OptJustifyClicked 0
        Case "CenterAlign": OptJustifyClicked 2
        Case "RightAlign": OptJustifyClicked 1
        Case "TextColor": SetTextClr
        Case "LineColor": SetLineClr
        Case "FillColor": SetFillClr
        Case "LineStyle": ShowLineStyleList
        Case "LineWidth": ShowLineWidthList
    End Select

End Sub

Private Sub Toolbar2_ButtonDropDown(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "TextColor": ShowTextClrClicked
        Case "LineColor": ShowLineClrClicked
        Case "FillColor": ShowFillClrClicked
    End Select

End Sub

Private Sub Toolbar2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ToolX = x / 1440

End Sub

Private Sub Toolbar2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Me.MousePointer = 0

End Sub

Private Sub TxtBordToggle_Click()

    If ctlActive.BorderStyle = 0 Then
        ctlActive.BorderStyle = 1
    Else
        ctlActive.BorderStyle = 0
    End If
    ActiveText.BorderOn = ctlActive.BorderStyle
    TxtBordToggle.Checked = (ctlActive.BorderStyle = 1)

End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        ctlActive.Caption = txtInput.text
        txtInput.Visible = False
        ctlActive.Visible = True
        Me.StatusBar1.Panels(3).text = ""
    End If

End Sub

Private Sub txtInput_LostFocus()

    If Not blnControlSelected Then
        ctlActive.Caption = txtInput.text
    End If
    ctlActive.Visible = True
    txtInput.Visible = False

End Sub

Private Sub VScroll_Change()

    SetPageComponents
    StartScaleTop = VScroll / 8
    ShowVertScale
    ShowGrid

End Sub

Private Sub VScroll_GotFocus()
    
    Me.MousePointer = 0

End Sub

Private Sub VScroll_Scroll()

    SetPageComponents
    StartScaleTop = VScroll / 8
    ShowVertScale

End Sub

Private Sub LoadCurrentReport()
Dim i As Integer
Dim j As Integer
Dim OldBound As Integer
Dim FLen As Long

    ReDim ReportFile.RpControl(0)
    NewBound = 0

    With ReportFile
        .DataBound = blnReportDataBound
        .DBName = strDataFileName
        .DBSource = strTableName
        .ImageFolder = strImageFolder
        .ImgPathTable = strImgPathTable
        .ImgPathField = strImgPathField
        For i = 0 To 2
            .SortField(i) = strSortField(i)
            .SortDescending(i) = blnSortDescending(i)
            .NewPageOnHeader(i) = blnNewPageOnHeader(i)
        Next i
        .PageSzNam = PageSizeName
        .PageSclUnit = PageScaleUnits
        .PageWd = PageWd
        .PageHt = PageHt
        .Orient = PageOrient
        .LMarg = LeftMarg
        .RMarg = RightMarg
        .TMarg = TopMarg
        .BMarg = BottomMarg
        .DesWd = picSection(0).width
        For i = 0 To 4
            .HeaderVis(i) = picSection(i).Visible
            .FooterVis(i) = picSection(10 - i).Visible
            .HeaderHt(i) = IIf(picSection(i).Visible, picSection(i).Height, 0)
            .FooterHt(i) = IIf(picSection(10 - i).Visible, picSection(10 - i).Height, 0)
        Next i
        .DetHt = picSection(5).Height
        For i = 0 To 10
            .SectColor(i) = picSection(i).BackColor
        Next i
        For j = 0 To UBound(DrawList)
            For i = 0 To Me.Controls.count - 1
                Set ctlTest = Me.Controls(i)
                If ctlTest.Tag > "" Then
                    If picSection(ctlTest.Tag).Visible And ctlTest.Name = DrawList(j).ctlName _
                    And ctlTest.Index = DrawList(j).ctlIndex And ctlTest.Visible = True Then
                        OldBound = UBound(.RpControl)
                        NewBound = OldBound + 1
                        ReDim Preserve .RpControl(NewBound)
                        .RpControl(NewBound).SecNo = ctlTest.Tag
                        If TypeOf ctlTest Is Line Then
                            .RpControl(NewBound).Type = cLine
                            .RpControl(NewBound).X1 = ctlTest.X1
                            .RpControl(NewBound).Y1 = ctlTest.Y1
                            .RpControl(NewBound).X2 = ctlTest.X2
                            .RpControl(NewBound).Y2 = ctlTest.Y2
                            .RpControl(NewBound).BdrClr = ctlTest.BorderColor
                            .RpControl(NewBound).BdrStl = ctlTest.BorderStyle
                            .RpControl(NewBound).BdrWd = ctlTest.BorderWidth
                        ElseIf TypeOf ctlTest Is Shape Then
                            .RpControl(NewBound).Type = cBox
                            .RpControl(NewBound).DisplayType = ctlTest.Shape
                            .RpControl(NewBound).Left = ctlTest.Left
                            .RpControl(NewBound).Top = ctlTest.Top
                            .RpControl(NewBound).width = ctlTest.width
                            .RpControl(NewBound).Height = ctlTest.Height
                            .RpControl(NewBound).BdrClr = ctlTest.BorderColor
                            .RpControl(NewBound).BdrStl = ctlTest.BorderStyle
                            .RpControl(NewBound).BdrWd = ctlTest.BorderWidth
                            .RpControl(NewBound).BckClr = ctlTest.BackColor
                            .RpControl(NewBound).BckStl = ctlTest.BackStyle
                        ElseIf TypeOf ctlTest Is Label Then
                            If ctlTest.Name = "Label" Then
                                .RpControl(NewBound).Type = cLabel
                            ElseIf ctlTest.Name = "Field" Or ctlTest.Name = "BoundImg" Then
                                .RpControl(NewBound).Type = ctlTest.LinkTimeout
                            End If
                            .RpControl(NewBound).Fieldname = ctlTest.DataField
                            .RpControl(NewBound).strText = ctlTest.Caption
                            .RpControl(NewBound).FntNam = ctlTest.FontName
                            .RpControl(NewBound).FntSiz = ctlTest.FontSize
                            .RpControl(NewBound).FntBld = ctlTest.FontBold
                            .RpControl(NewBound).FntItl = ctlTest.FontItalic
                            .RpControl(NewBound).FntUnd = ctlTest.FontUnderline
                            .RpControl(NewBound).BdrStl = ctlTest.BorderStyle
                            .RpControl(NewBound).Left = ctlTest.Left
                            .RpControl(NewBound).Top = ctlTest.Top
                            .RpControl(NewBound).width = ctlTest.width
                            .RpControl(NewBound).Height = ctlTest.Height
                            .RpControl(NewBound).BckStl = ctlTest.BackStyle
                            .RpControl(NewBound).BckClr = ctlTest.BackColor
                            .RpControl(NewBound).BdrStl = ctlTest.BorderStyle
                            .RpControl(NewBound).ForClr = ctlTest.ForeColor
                            .RpControl(NewBound).Align = ctlTest.Alignment
                        ElseIf TypeOf ctlTest Is MSForms.Image Then
                            .RpControl(NewBound).Type = cImage
                            .RpControl(NewBound).Left = ctlTest.Left
                            .RpControl(NewBound).Top = ctlTest.Top
                            .RpControl(NewBound).width = ctlTest.width
                            .RpControl(NewBound).Height = ctlTest.Height
                            SavePicture ctlTest.Picture, App.Path & "\tmpfile"
                            Open App.Path & "\tmpfile" For Binary As #1
                            FLen = LOF(1)
                            .RpControl(NewBound).ImgData = String$(FLen, " ")
                            Get #1, 1, .RpControl(NewBound).ImgData
                            Close #1
                            Kill App.Path & "\tmpfile"
                        ElseIf TypeOf ctlTest Is CheckBoxControl Then
                            .RpControl(NewBound).Type = cCheckBox
                            .RpControl(NewBound).Left = ctlTest.Left
                            .RpControl(NewBound).Top = ctlTest.Top
                            .RpControl(NewBound).width = ctlTest.width
                            .RpControl(NewBound).Height = ctlTest.Height
                            .RpControl(NewBound).ForClr = ctlTest.BorderColor
                            .RpControl(NewBound).BckClr = ctlTest.BackColor
                            .RpControl(NewBound).DisplayType = ctlTest.DisplayType
                            .RpControl(NewBound).Sunken = ctlTest.Sunken
                            .RpControl(NewBound).Fieldname = ctlTest.DataField
                        End If
                        Exit For
                    End If
                End If
            Next i
        Next j
    End With

End Sub

Private Sub ShowPopUp(WhatsClicked As Integer)

    CutObj.Enabled = blnControlSelected + blnGroupSelected
    CopyObj.Enabled = blnControlSelected + blnGroupSelected
    PasteObj.Enabled = CutActive Or CopyActive
    DelObj.Enabled = blnControlSelected + blnGroupSelected
    
    SetBackColor.Visible = (WhatsClicked = cNone)
    If (WhatsClicked = cLine) Or (WhatsClicked = cBox) Or (WhatsClicked = cNone) Then
        BringFront.Visible = True
        SendBack.Visible = True
        BringFront.Enabled = blnControlFound
        SendBack.Enabled = blnControlFound
    ElseIf (WhatsClicked <> cCheckBox) Then
        BringFront.Visible = True
        SendBack.Visible = True
        BringFront.Enabled = True
        SendBack.Enabled = True
    Else
        BringFront.Visible = False
        SendBack.Visible = False
    End If
    
    ChkOptions.Visible = (WhatsClicked = cCheckBox)
    chk3d.Visible = (WhatsClicked = cCheckBox)
    ChkOptions2.Visible = (WhatsClicked = cBox)
    
    If (WhatsClicked = cDataField) Or (WhatsClicked = cCalcField) Or (WhatsClicked = cSumField) Then
        FormatText.Visible = True
    Else
        FormatText.Visible = False
    End If
    
    AssignField.Visible = (WhatsClicked = cDataField)
    EditCalc.Visible = (WhatsClicked = cCalcField)
    EditSummary.Visible = (WhatsClicked = cSumField)
    EditDatePage.Visible = (WhatsClicked = cDatePageField)
    
    EditText.Visible = (WhatsClicked = cLabel)
    If (WhatsClicked = cLabel) Or (WhatsClicked = cDataField) Or (WhatsClicked = cDatePageField) Then
        TxtBordToggle.Visible = True
        TxtBordToggle.Checked = (ctlActive.BorderStyle = 1)
    Else
        TxtBordToggle.Visible = False
    End If
    
    If blnGroupSelected Then
        blnControlSelected = False
        ShowGrid
        HiliteObjects
        intGroupRestraint = resNone
    ElseIf blnControlSelected Then
        ShowGrips ctlActive.Tag
    End If
    
    PopupMenu ReportObjectPopUp
    blnMouseIsDown = False

End Sub

Private Sub MoveInControl(GetX As Single, GetY As Single, Optional GetContents As String = "")
Dim ctlUnits As Long

    lngPrevState = lngState
    lngState = OverControl
    Me.MousePointer = 5
    
    
    StatusBar1.Panels(3).text = GetContents
    
    If blnMouseIsDown Then
        If TypeOf ctlActive Is MSForms.Image Then        'unbound picture control
            ctlUnits = 72
        ElseIf TypeOf ctlActive Is CheckBoxControl Then          'checkbox control
            ctlActive.Visible = False
            ctlUnits = 96
        Else
            ctlActive.Visible = False
            ctlUnits = 1440
        End If
        
        If blnSnapOn Then
            GetX = SnapToGrid(GetX / ctlUnits)
            GetY = SnapToGrid(GetY / ctlUnits)
        Else
            GetX = GetX / ctlUnits
            GetY = GetY / ctlUnits
        End If
    
        lngState = MoveControl
        If blnControlSelected Then
            If Not blnDragStarted Then
                WriteToUndoList unControl, unMove
                ClearSelection
                blnDragStarted = True
            End If
            SelectBox.Left = GetX - dX1 + StartX
            SelectBox.Top = GetY - dY1 + StartY
            SelectBox.Visible = True
        ElseIf blnGroupSelected Then
            If IsInSelectGroup(ctlActive.Name, ctlActive.Index) Then
                SelectBox.Left = GetX - dX1 + StartX
                If Not blnAtLimit Then
                    SelectBox.Top = GetY - dY1 + StartY
                End If
                SelectBox.Visible = True
                MoveSelectedObjects ctlActive.Tag, SelectBox.Left + dX1, SelectBox.Top + dY1
            Else
                blnGroupSelected = False
                ClearSelection
                blnControlSelected = True
                SelectBox.Left = GetX - dX1 + StartX
                SelectBox.Top = GetY - dY1 + StartY
                SelectBox.Visible = True
            End If
        End If
    End If
    

End Sub

Private Sub HScaleSelectObjects()
Dim i As Integer
Dim LeftBorder As Single, RightBorder As Single

    NumInGrp = -1
    
    If ScaleStartX >= ScaleEndX Then
        LeftBorder = ScaleEndX - 0.17
        RightBorder = ScaleStartX - 0.17
    Else
        LeftBorder = ScaleStartX - 0.17
        RightBorder = ScaleEndX - 0.17
    End If

    For i = 0 To Me.Controls.count - 1
        Set ctlTest = Me.Controls(i)
        If ctlTest.Tag > "" Then
            If TypeOf ctlTest Is Line Then
                If LeftBorder < ctlTest.X1 And RightBorder > ctlTest.X2 Then
                    blnGroupSelected = True
                    NumInGrp = NumInGrp + 1
                    ReDim Preserve SelectedCtl(NumInGrp)
                    Set SelectedCtl(NumInGrp).ctl = ctlTest
                End If
            Else
                If LeftBorder < ctlTest.Left And RightBorder > ctlTest.Left + ctlTest.width Then
                    blnGroupSelected = True
                    NumInGrp = NumInGrp + 1
                    ReDim Preserve SelectedCtl(NumInGrp)
                    Set SelectedCtl(NumInGrp).ctl = ctlTest
                End If
            End If
        End If
        blnSelectArrayInit = True
    Next i

End Sub

Private Sub VScaleSelectObjects()
Dim i As Integer
Dim TopBorder As Single, BottomBorder As Single
Dim SecTop As Single

    NumInGrp = -1
    
    If ScaleStartY > ScaleEndY Then
        TopBorder = ScaleEndY
        BottomBorder = ScaleStartY
    Else
        TopBorder = ScaleStartY
        BottomBorder = ScaleEndY
    End If

    For i = 0 To Me.Controls.count - 1
        Set ctlTest = Me.Controls(i)
        If ctlTest.Tag > "" Then
            SecTop = picSection(ctlTest.Tag).Top
            If TypeOf ctlTest Is Line Then
                If TopBorder - SecTop < ctlTest.Y1 And BottomBorder + 0.17 - SecTop > ctlTest.Y2 Then
                    blnGroupSelected = True
                    NumInGrp = NumInGrp + 1
                    ReDim Preserve SelectedCtl(NumInGrp)
                    Set SelectedCtl(NumInGrp).ctl = ctlTest
                End If
            Else
                If TopBorder - SecTop < ctlTest.Top And BottomBorder + 0.17 - SecTop > ctlTest.Top + ctlTest.Height Then
                    blnGroupSelected = True
                    NumInGrp = NumInGrp + 1
                    ReDim Preserve SelectedCtl(NumInGrp)
                    Set SelectedCtl(NumInGrp).ctl = ctlTest
                End If
            End If
        End If
        blnSelectArrayInit = True
    Next i

End Sub
