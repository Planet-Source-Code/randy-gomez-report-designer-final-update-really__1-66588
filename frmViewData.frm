VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewData 
   BackColor       =   &H8000000B&
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3690
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   6509
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmViewData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Integer
Dim ColIndex As Integer
Dim AddedItem As ListItem

    For i = 0 To rstData.Fields.count - 1
        ListView1.ColumnHeaders.Add , , rstData.Fields(i).Name
    Next i

    rstData.MoveFirst
    For i = 0 To rstData.RecordCount - 1
        If Not IsNull(rstData.Fields(0)) Then
            Set AddedItem = ListView1.ListItems.Add(, , rstData.Fields(0))
        Else
            Set AddedItem = ListView1.ListItems.Add(, , "")
        End If
        For ColIndex = 1 To rstData.Fields.count - 1
            If Not IsNull(rstData.Fields(ColIndex)) Then
                AddedItem.ListSubItems.Add , , rstData.Fields(ColIndex)
            End If
        Next ColIndex
        rstData.MoveNext
    Next i
    
    Me.Caption = frmDataConn.cboTables(ListIndex) & " : " & rstData.RecordCount & " records"

End Sub

Private Sub Form_Resize()

    ListView1.width = Me.ScaleWidth - 100
    ListView1.Height = Me.ScaleHeight - 100

End Sub
