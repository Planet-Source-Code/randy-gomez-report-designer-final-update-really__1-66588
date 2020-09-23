VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'////////////////////////////////////////////////////////
'///                  Image List Class
'///                (clsPrintDialog.cls)
'///_____________________________________________________
'/// Show a dialog to select a printer and to set printer
'/// properties. Selected printer will be set to Printer
'/// object WITHOUT CHANGING DEFAULT PRINTER.
'/// Printer object could print to selected printer then.
'/// This class work like "MS Common Dialogs" ShowPrinter
'/// method, but it set Printer object without changing
'/// default printer.
'///_____________________________________________________
'/// Last modification  : Ago/10/2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Created
'/// Project: RamoSoft Component Suite ' I borrowed this code from a another project from myself
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// based in code found in the internet
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit
' --- API CONSTANTS
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&
' --- API TYPES DEFINITION
Public hWndOwner As Long
Private Type PRINTDLG_TYPE
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Private Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
'
' --- API DECLARATIONS
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'
' --- PUBLIC ENUM
Public Enum PrinterConstants
    cdlPDAllPages = &H0
    cdlPDCollate = &H10
    cdlPDDisablePrintToFile = &H80000
    cdlPDHelpButton = &H800
    cdlPDHidePrintToFile = &H100000
    cdlPDNoPageNums = &H8
    cdlPDNoSelection = &H4
    cdlPDNoWarning = &H80
    cdlPDPageNums = &H2
    cdlPDPrintSetup = &H40
    cdlPDPrintToFile = &H20
    cdlPDReturnDC = &H100
    cdlPDReturnDefault = &H400
    cdlPDReturnIC = &H200
    cdlPDSelection = &H1
    cdlPDUseDevModeCopies = &H40000
End Enum

Public Enum ErrorConstants
    cdlCancel = 32755
End Enum
'
' --- PRIVATE VARIABLES
Private m_iMinPage As Integer  ' Local copy of Min
Private m_iMaxPage As Integer  ' Local copy of Max
Private m_iFromPage As Integer ' Local copy of FromPage
Private m_iToPage As Integer   ' Local copy of ToPage
' N.B. 0 >= Min >= FromPage >= ToPage >= Max
'      If Max=0 then no limits.
'
' --- PUBLIC VARIABLES
Public Flags As PrinterConstants
Public CancelError As Boolean
Private Sub Class_Initialize()
'Debug.Print "Creating " & TypeName(Me)
    Flags = cdlPDPageNums
    m_iMinPage = 0
    m_iMaxPage = 10000
    m_iFromPage = 0
    m_iToPage = 0
    CancelError = False
End Sub
Property Get Min() As Integer
    Min = m_iMinPage
End Property
Property Let Min(ByVal intNewValue As Integer)
    intNewValue = IIf(intNewValue > 0, intNewValue, 0)
    m_iMinPage = intNewValue
    If intNewValue > m_iFromPage Then m_iFromPage = intNewValue
    If intNewValue > m_iToPage Then m_iToPage = intNewValue
    If intNewValue > m_iMaxPage Then m_iMaxPage = intNewValue
End Property
Property Get FromPage() As Integer
  FromPage = m_iFromPage
End Property
Property Let FromPage(ByVal intNewValue As Integer)
    intNewValue = IIf(intNewValue > 0, intNewValue, 0)
    m_iFromPage = intNewValue
    If intNewValue > m_iToPage Then m_iToPage = intNewValue
    If intNewValue > m_iMaxPage Then m_iMaxPage = intNewValue
    If intNewValue < m_iMinPage Then m_iMinPage = intNewValue
    Flags = Flags Or cdlPDPageNums
End Property
Property Get ToPage() As Integer
    ToPage = m_iToPage
End Property
Property Let ToPage(ByVal intNewValue As Integer)
    intNewValue = IIf(intNewValue > 0, intNewValue, 0)
    m_iToPage = intNewValue
    If intNewValue > m_iMaxPage Then m_iMaxPage = intNewValue
    If intNewValue < m_iFromPage Then m_iFromPage = intNewValue
    If intNewValue < m_iMinPage Then m_iMinPage = intNewValue
    Flags = Flags Or cdlPDPageNums
End Property
Property Get Max() As Integer
    Max = m_iMaxPage
End Property
Property Let Max(ByVal intNewValue As Integer)
    intNewValue = IIf(intNewValue > 0, intNewValue, 0)
    m_iMaxPage = intNewValue
    If intNewValue < m_iToPage Then m_iToPage = intNewValue
    If intNewValue < m_iFromPage Then m_iFromPage = intNewValue
    If intNewValue < m_iMinPage Then m_iMinPage = intNewValue
    
    If m_iMaxPage = 0 Then Flags = (Flags And (&HFFFF Xor cdlPDPageNums))
End Property
Public Function ShowPrinter() As Boolean
    Dim LrPrintDlg As PRINTDLG_TYPE
    Dim LrDevMode As DEVMODE_TYPE
    Dim LrDevName As DEVNAMES_TYPE
    Dim LnDevMode As Long
    Dim LnDevName As Long
    Dim LnReturn As Integer
    Dim LoPrinter As Printer
    Dim LsNewPrtName As String
    Dim LbCancel   As Boolean
    
    LbCancel = False
    ' Use PrintDialog to get the handle to a memory block
    ' with a LrDevMode and LrDevName structures
    With LrPrintDlg
        .lStructSize = Len(LrPrintDlg)
        .hWndOwner = hWndOwner
        .Flags = Flags
        .nMinPage = m_iMinPage
        .nFromPage = m_iFromPage
        .nToPage = m_iToPage
        .nMaxPage = m_iMaxPage
    End With
    'Set the current orientation and duplex setting
    LrDevMode.dmDeviceName = Printer.DeviceName
    LrDevMode.dmSize = Len(LrDevMode)
    LrDevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
    LrDevMode.dmOrientation = Printer.Orientation
    On Error Resume Next
    LrDevMode.dmDuplex = Printer.Duplex
    On Error GoTo 0
    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    LrPrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(LrDevMode))
    LnDevMode = GlobalLock(LrPrintDlg.hDevMode)
    If LnDevMode > 0 Then
        CopyMemory ByVal LnDevMode, LrDevMode, Len(LrDevMode)
        LnReturn = GlobalUnlock(LnDevMode)
    End If
    'Set the current driver, device, and port name strings
    With LrDevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With
    With Printer
        LrDevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With
    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
    LrPrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(LrDevName))
    LnDevName = GlobalLock(LrPrintDlg.hDevNames)
    If LnDevName > 0 Then
        CopyMemory ByVal LnDevName, LrDevName, Len(LrDevName)
        LnReturn = GlobalUnlock(LnDevName)
    End If
    'Call the print dialog up and let the user make changes
    If PrintDialog(LrPrintDlg) Then
        'First get the LrDevName structure.
        LnDevName = GlobalLock(LrPrintDlg.hDevNames)
        CopyMemory LrDevName, ByVal LnDevName, 45
        LnReturn = GlobalUnlock(LnDevName)
        With LrPrintDlg
            Flags = .Flags
            m_iFromPage = .nFromPage
            m_iToPage = .nToPage
        End With
        GlobalFree LrPrintDlg.hDevNames
        'Next get the LrDevMode structure and set the printer
        'properties appropriately
        LnDevMode = GlobalLock(LrPrintDlg.hDevMode)
        CopyMemory LrDevMode, ByVal LnDevMode, Len(LrDevMode)
        LnReturn = GlobalUnlock(LrPrintDlg.hDevMode)
        GlobalFree LrPrintDlg.hDevMode
        LsNewPrtName = UCase(Left(LrDevMode.dmDeviceName, InStr(LrDevMode.dmDeviceName, Chr(0)) - 1))
        If Printer.DeviceName <> LsNewPrtName Then
            For Each LoPrinter In Printers
                If UCase(LoPrinter.DeviceName) = LsNewPrtName Then
                    Set Printer = LoPrinter
                End If
            Next
        End If
        On Error Resume Next
        'Set printer object properties according to selections made by user
        With Printer
            .Copies = LrDevMode.dmCopies
            .Duplex = LrDevMode.dmDuplex
            .Orientation = LrDevMode.dmOrientation
        End With
        On Error GoTo 0
    Else
        GlobalFree LrPrintDlg.hDevMode
        GlobalFree LrPrintDlg.hDevNames
        LbCancel = True
        If CancelError Then Err.Raise cdlCancel, "clsPrintDialog", "Selection Canceled"
    End If
    ShowPrinter = Not LbCancel
End Function

