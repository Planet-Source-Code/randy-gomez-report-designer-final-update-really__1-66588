VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About RG Report Designer"
   ClientHeight    =   4455
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6240
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3074.92
   ScaleMode       =   0  'User
   ScaleWidth      =   5859.682
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2460
      TabIndex        =   0
      Top             =   4020
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   2310
      Left            =   240
      Picture         =   "frmAbout.frx":000C
      Top             =   120
      Width           =   5745
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5803.339
      Y1              =   2225.952
      Y2              =   2225.952
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   $"frmAbout.frx":25BF
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   300
      TabIndex        =   1
      Top             =   2460
      Width           =   5745
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5746.996
      Y1              =   2236.305
      Y2              =   2236.305
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H8000000B&
      Caption         =   "Version 0.6"
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
      Left            =   180
      TabIndex        =   3
      Top             =   2940
      Width           =   1065
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H8000000B&
      Caption         =   $"frmAbout.frx":264A
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   60
      TabIndex        =   2
      Top             =   3360
      Width           =   6090
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

