VERSION 5.00
Begin VB.Form frmCalcField 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Place Calculated Field (in Detail section)"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   Icon            =   "frmCalcField.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFieldName 
      Height          =   330
      Left            =   2595
      TabIndex        =   10
      Top             =   2220
      Width           =   2490
   End
   Begin VB.CommandButton cmdOper 
      Height          =   360
      Index           =   4
      Left            =   4020
      Picture         =   "frmCalcField.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1155
      Width           =   345
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
      Left            =   3885
      TabIndex        =   12
      Top             =   2925
      Width           =   1020
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
      Left            =   2715
      TabIndex        =   11
      Top             =   2925
      Width           =   1020
   End
   Begin VB.CommandButton cmdOper 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   4740
      Picture         =   "frmCalcField.frx":0596
      TabIndex        =   9
      Top             =   1155
      Width           =   345
   End
   Begin VB.CommandButton cmdOper 
      Caption         =   "("
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   4380
      Picture         =   "frmCalcField.frx":0B20
      TabIndex        =   7
      Top             =   1155
      Width           =   345
   End
   Begin VB.ListBox lstFields 
      Height          =   2010
      Left            =   180
      TabIndex        =   1
      Top             =   1485
      Width           =   2160
   End
   Begin VB.CommandButton cmdOper 
      Height          =   360
      Index           =   3
      Left            =   3660
      Picture         =   "frmCalcField.frx":10AA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1155
      Width           =   345
   End
   Begin VB.CommandButton cmdOper 
      Height          =   360
      Index           =   2
      Left            =   3300
      Picture         =   "frmCalcField.frx":1634
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1155
      Width           =   345
   End
   Begin VB.CommandButton cmdOper 
      Height          =   360
      Index           =   1
      Left            =   2940
      Picture         =   "frmCalcField.frx":1BBE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1155
      Width           =   345
   End
   Begin VB.CommandButton cmdOper 
      Height          =   360
      Index           =   0
      Left            =   2565
      Picture         =   "frmCalcField.frx":2148
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1155
      Width           =   345
   End
   Begin VB.TextBox txtEquation 
      Height          =   765
      Left            =   165
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   330
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "Calculation Field Contents :"
      Height          =   240
      Left            =   180
      TabIndex        =   14
      Top             =   105
      Width           =   1995
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "Field Name :"
      Height          =   255
      Left            =   2610
      TabIndex        =   13
      Top             =   1980
      Width           =   960
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Available Fields : (click to add)"
      Height          =   270
      Left            =   180
      TabIndex        =   8
      Top             =   1230
      Width           =   2160
   End
End
Attribute VB_Name = "frmCalcField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strAdd As String
Dim strSaveContents As String
Dim strSaveFieldName As String

Private Sub cmdCancel_Click()

    blnEditExisting = False
    frmDesign.optAction(0).value = True
    Unload Me

End Sub

Private Sub cmdOK_Click()

    If txtFieldName = "" Then
        MsgBox "You must give this calculated field a name", vbOKOnly + vbExclamation
        txtFieldName.SetFocus
        Exit Sub
    End If
    
'if the calculation has been changed, update the active control
'replace '~' character with new calculation - preserves any formatting info
    If blnEditExisting Then
        strCalcDataFieldContents = Replace(strSaveContents, "~", txtEquation.text)
        strCalcDataFieldContents = Replace(strCalcDataFieldContents, strSaveFieldName, txtFieldName)
        ctlActive.Caption = strCalcDataFieldContents
        ctlActive.ToolTipText = strCalcDataFieldContents
        ctlActive.DataField = txtFieldName
    Else
        frmDesign.CreateField 5, "Calc", "={" & txtEquation & "} as " & txtFieldName, txtFieldName, True
    End If
    blnEditExisting = False
    Unload Me

End Sub

Private Sub cmdOper_Click(Index As Integer)

'inserts operator into the calculation based on selection
    Select Case Index
        Case 0: strAdd = " + "
        Case 1: strAdd = " - "
        Case 2: strAdd = " * "
        Case 3: strAdd = " / "
        Case 4: strAdd = "^"
        Case 5: strAdd = "("
        Case 6: strAdd = ")"
    End Select
    
    UpdateEquation

End Sub

Private Sub Form_Load()
Dim i As Integer

'fill in the field listbox
    If strDataFileName > "" Then
        If strTableName > "" Then
            For i = 0 To rstData.Fields.count - 1
                Me.lstFields.AddItem rstData.Fields(i).Name
            Next i
        End If
    Else
    End If

'save the current field contents to variables, temporarily replacing the calculation
'portion with the '~' character
    If blnEditExisting Then
        txtEquation = Mid(strCalcDataFieldContents, 3, InStr(1, strCalcDataFieldContents, "}") - 3)                     'parse out stuff in curly brackets
        strSaveContents = Replace(strCalcDataFieldContents, txtEquation, "~")
        txtFieldName = ctlActive.DataField
        strSaveFieldName = ctlActive.DataField
    End If

End Sub

Private Sub lstFields_Click()

    strAdd = "[" & lstFields & "]"
    UpdateEquation

End Sub

Private Sub UpdateEquation()
Dim StartChar As Integer

'called at form loading and whenever a selection is made
    StartChar = txtEquation.SelStart
    txtEquation = Left(txtEquation, txtEquation.SelStart) & _
    strAdd & _
    Right(txtEquation, Len(txtEquation) - (txtEquation.SelStart + txtEquation.SelLength))
    
    txtEquation.SelStart = StartChar + Len(strAdd)
    txtEquation.SetFocus

End Sub
