Attribute VB_Name = "modGlobals"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation _
      As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Function SetPrnPort(sPortName As String) As Boolean
    Dim oPrn As Printer

    sPortName = LCase(sPortName)
    If sPortName = LCase(Printer.Port) Then Exit Function
    For Each oPrn In Printers
        If LCase(oPrn.Port) = sPortName Then
            Set Printer = oPrn
            SetPrnPort = True
            Exit For
        End If
    Next oPrn
    Printer.Print ""
End Function


Public Function IsIDE() As Boolean
    On Error GoTo ERR_IDE
    Debug.Print (1 / 0)
    IsIDE = False
    Exit Function
ERR_IDE:
    IsIDE = True
    Err.Clear
End Function

Public Function ObjFromPtr(lObjPtr As Long) As Object
    If lObjPtr <> 0 Then
        Dim LoTmp As Object
        ' Turn the pointer into an illegal, uncounted interface
        CopyMemory LoTmp, lObjPtr, 4
        ' Do NOT hit the End button here! You will crash!
        ' Assign to legal reference
        Set ObjFromPtr = LoTmp
        ' Still do NOT hit the End button here! You will still crash!
        ' Destroy the illegal reference
        CopyMemory LoTmp, 0&, 4
        ' OK, hit the End button if you must--you'll probably still crash,
        ' but it will be because of the subclass, not the uncounted reference
    End If
End Function

Function IsWindowLocal(ByVal hWnd As Long) As Boolean
    Dim LnWinID As Long
    
    Call GetWindowThreadProcessId(hWnd, LnWinID)
    IsWindowLocal = (LnWinID = GetCurrentProcessId())
End Function



