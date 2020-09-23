VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Properties"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboValues 
      Height          =   315
      Left            =   2610
      TabIndex        =   2
      Top             =   330
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Index           =   0
      Left            =   2610
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label lblParam 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
