VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAppendString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_lBufferPos As Long
Private m_sBuffer As String
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Sub AppendString(strAdd As String)
On Error GoTo ERR_H
    ' If the current position + length of string to add is greater than the
    ' length of the BigString, then increase the size of the BigString
    If (m_lBufferPos + LenB(strAdd)) > LenB(m_sBuffer) Then
        m_sBuffer = m_sBuffer & Space$(8192)
    End If
    'Add strAdd to the BigString
    CopyMemory ByVal StrPtr(m_sBuffer) + m_lBufferPos, ByVal StrPtr(strAdd), LenB(strAdd)
    'Move the pointer to show where the end of the string is
    m_lBufferPos = m_lBufferPos + LenB(strAdd)
Exit Sub
ERR_H:
    RaiseErr Err.Number, "AppendString", Err.Description
End Sub

Public Property Get ReturnString() As Variant
    'Trim off the blank characters and return the string
    m_sBuffer = Left$(m_sBuffer, m_lBufferPos \ 2) 'trim back to size
    ReturnString = m_sBuffer
End Property

Private Sub Class_Initialize()
'Set the starting position of the BigString to 0
    m_lBufferPos = 0
End Sub

Friend Sub RaiseErr(ByVal lErrNum As RSErrorCode, Optional sRoutineName As String, _
    Optional sDescription As String)
    RaiseError lErrNum, TypeName(Me), sRoutineName, sDescription, Erl
End Sub
