VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelField 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Select Data Field"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   3480
      TabIndex        =   1
      Top             =   2760
      Width           =   795
   End
   Begin MSComctlLib.ListView lstFields 
      Height          =   2625
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   4630
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Field Name"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Type"
         Text            =   "Data Type"
         Object.Width           =   2822
      EndProperty
   End
End
Attribute VB_Name = "frmSelField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

    Me.Hide

End Sub

Private Sub lstFields_Click()
Dim i As Integer
    
    If TypeOf ctlActive Is Label Then
        ctlActive.Caption = "=![" & lstFields.SelectedItem & "]"
    ElseIf TypeOf ctlActive Is PictureBox Then
        For i = 0 To UBound(DataField)
            If DataField(i) = lstFields Then
                If rstData.Fields(i).Type <> adBoolean Then
                    MsgBox "A Checkbox must be linked to a True/False field!", vbOKOnly + vbExclamation
                    Exit Sub
                End If
            End If
        Next i
    End If
    ctlActive.DataField = lstFields.SelectedItem
    ctlActive.ToolTipText = "=![" & lstFields.SelectedItem & "]"
    frmSelField.Visible = False

End Sub


