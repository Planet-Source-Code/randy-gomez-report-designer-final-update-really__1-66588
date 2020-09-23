VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBackgroundBuffer 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   120
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   3855
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4320
      TabIndex        =   1
      Top             =   4020
      Width           =   1095
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Center
      Caption         =   "www.ramosoft.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2640
      MouseIcon       =   "frmAbout.frx":3A57
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1620
      Width           =   2775
   End
   Begin VB.Label lblRelease 
      Alignment       =   2  'Center
      Caption         =   "Release 5.0.0, Build 001006"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   6
      Top             =   4140
      Width           =   2415
   End
   Begin VB.Label lblMail 
      Alignment       =   2  'Center
      Caption         =   "Powered by RamoSoft de Mexico"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1380
      Width           =   2775
   End
   Begin VB.Label lblWarning 
      Caption         =   $"frmAbout.frx":3BA9
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2640
      TabIndex        =   4
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lblRights 
      Alignment       =   2  'Center
      Caption         =   "               Copyright © 2001                      RamoSoft de Mexico S.A. de C.V.    All rights reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "RamoSoft PrintPreview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub Start()
    lblRelease = "Release v" & App.Major & "." & _
        App.Minor & " (Build " & App.Revision & ")"
    Show vbModal
End Sub

Private Sub Form_Load()
    lblRights = "Copyright © 2001" & vbCrLf & "RamoSoft de Mexico S.A. de C.V." & vbCrLf & "All rights reserved."
End Sub

Private Sub lblURL_Click()
    Call ShellExecute(Me.hWnd, vbNullString, "http://www.ramosoft.com/", vbNullString, "C:\", 0)
End Sub

Private Sub cmdOk_Click()
   Unload Me
End Sub

