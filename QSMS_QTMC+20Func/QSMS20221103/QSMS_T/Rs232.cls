VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rs232"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''**RQ09102710  Denver      2009.10.27    Add ²âÊÔLCR ÐÍºÅÎª4300 µÄÒÇÆ÷  £¨0063£©



'Application
Const AppTitle = "4300 Serial Control"
'About
Const ABOUTFORMCAPTION = "Wayne Kerr Electronics"

'General
Public OpenForms
Const EMPTYSTR = ""
Public MouseButton As Integer           '1 = left, 2 = Right, 4 = Centre
Public InstInitialised As Boolean

 Const MAXPREVIOUSCOMMANDS = 25
Public EnteredCommands As New Collection
Public ReturnedData As String

 Const PORTCOLUMN = 0
 Const BAUDCOLUMN = 1
 Const PARITYCOLUMN = 2
 Const DATABITSCOLUMN = 3
 Const STOPBITSCOLUMN = 4

'Serial Port
Public COMPortError As Boolean
Public COMPortEvent As Variant
Public setComport As Integer

Public BaudRateCol As New Collection
Public ParityCol As New Collection
Public DataBitsCol As New Collection
Public StopBitsCol As New Collection
Public ComPortCol As New Collection

'Default Serial port values COM Port 1, 9600 Baud, No Parity, 8 Data Bits, 1 stop bit
Const DEFCOMPORT = 1
Const DEFBAUDRATE = "9600"
Const DEFPARITY = "N"
Const DEFDATABITS = "8"
Const DEFSTOPBITS = "1"
Public strResult As String



Function SendCommand(MSComm1 As MSComm, ByVal intReadResultBack As Integer, Command) As Boolean


Dim OutString$
OutString$ = Command

If Right(OutString$, 1) <> vbLf Then OutString$ = OutString$ & vbLf

If OpenSerialPort(MSComm1) = True Then
    MSComm1.SThreshold = 1
    COMPortEvent = 0
    MSComm1.Output = OutString$
    
    SendCommand = True
    If intReadResultBack = 1 Then
        SendCommand = False
        Do Until (COMPortEvent = 2)
            DoEvents
            Sleep (10)
            DoEvents
            
        Loop
'        OutString$ = "*LOC"
'        If Right(OutString$, 1) <> vbLf Then OutString$ = OutString$ & vbLf
'        MSComm1.Output = OutString$
    End If
    Call Sleep(80)
    SendCommand = True
    
Else
    MsgBox "COM Port Error - No Data Sent", vbOKOnly + vbCritical, AppTitle & " SendCommand"
End If

End Function


Sub SetupSerialPort(MSComm1 As MSComm, setComport As Integer)
     
Dim Settings As String
    If CloseSerialPort(MSComm1) = False Then
        
            If setComport = 0 Then setComport = DEFCOMPORT
            MSComm1.CommPort = setComport
            Settings = DEFBAUDRATE & "," & DEFPARITY & "," & DEFDATABITS & "," & DEFSTOPBITS
            MSComm1.Settings = Settings
            MSComm1.InputLen = 512
            MSComm1.Handshaking = comRTSXOnXOff
            MSComm1.InBufferSize = 512
            MSComm1.OutBufferSize = 512
            MSComm1.SThreshold = 5
            MSComm1.OutBufferCount = 0
            MSComm1.RThreshold = 5
            MSComm1.InputMode = comInputModeText
    End If
End Sub
Function OpenSerialPort(MSComm1 As MSComm) As Boolean
On Error GoTo OpenSerialPortErr
    OpenSerialPort = MSComm1.PortOpen
    If OpenSerialPort = False Then
        MSComm1.PortOpen = True
        OpenForms = DoEvents
        OpenForms = DoEvents
        OpenSerialPort = MSComm1.PortOpen
        If OpenSerialPort = False Then MsgBox "COM Port Error - Cannot Open Port", vbOKOnly + vbCritical, "CloseSerialPort"
    End If
    Exit Function
OpenSerialPortErr:
    OpenSerialPort = False
    MsgBox "Error Number:  " & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbExclamation, AppTitle & " OpenSerialPort"
End Function

Function CloseSerialPort(MSComm1 As MSComm) As Boolean

    CloseSerialPort = MSComm1.PortOpen
    If CloseSerialPort = True Then
        MSComm1.PortOpen = False
        OpenForms = DoEvents
        CloseSerialPort = MSComm1.PortOpen
        If CloseSerialPort = True Then MsgBox "COM Port Error - Cannot Close Port", vbOKOnly + vbCritical, "CloseSerialPort"
    End If

End Function



