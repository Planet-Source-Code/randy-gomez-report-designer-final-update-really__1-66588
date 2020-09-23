VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'////////////////////////////////////////////////////////
'///              Common Export Interface
'///                    (IExport.cls)
'///_____________________________________________________
'/// Abstract class that brings the common interface for
'/// the exportation process, each export component must
'/// implement this class.
'///_____________________________________________________
'/// Last modification  : Ago//2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Created
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit

Public FileName As String


Public Function DoAction(ByVal iAction As Integer, _
    ByVal vValue As Variant) As Integer
End Function

Public Function Export(oPages As Pages) As Integer
Attribute Export.VB_UserMemId = 0
End Function

Public Function QueryInfo(iType As Integer) As Variant
End Function

Public Property Get StillWorking() As Boolean
End Property

