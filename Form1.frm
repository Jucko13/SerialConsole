VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   17835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm comm 
      Left            =   10065
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_SerialPort")

For Each objItem In colItems
    Debug.Print "Binary: " & objItem.Binary
    Debug.Print "Description: " & objItem.Description
    Debug.Print "Device ID: " & objItem.DeviceID
    Debug.Print "Maximum Baud Rate: " & objItem.MaxBaudRate
    Debug.Print "Maximum Input Buffer Size: " & objItem.MaximumInputBufferSize
    Debug.Print "Maximum Output Buffer Size: " & _
        objItem.MaximumOutputBufferSize
    Debug.Print "Name: " & objItem.Name
    Debug.Print "OS Auto Discovered: " & objItem.OSAutoDiscovered
    Debug.Print "PNP Device ID: " & objItem.PNPDeviceID
    Debug.Print "Provider Type: " & objItem.ProviderType
    Debug.Print "Settable Baud Rate: " & objItem.SettableBaudRate
    Debug.Print "Settable Data Bits: " & objItem.SettableDataBits
    Debug.Print "Settable Flow Control: " & objItem.SettableFlowControl
    Debug.Print "Settable Parity: " & objItem.SettableParity
    Debug.Print "Settable Parity Check: " & objItem.SettableParityCheck
    Debug.Print "Settable RLSD: " & objItem.SettableRLSD
    Debug.Print "Settable Stop Bits: " & objItem.SettableStopBits
    Debug.Print "Supports 16-Bit Mode: " & objItem.Supports16BitMode
    Debug.Print "Supports DTRDSR: " & objItem.SupportsDTRDSR
    Debug.Print "Supports Elapsed Timeouts: " & _
        objItem.SupportsElapsedTimeouts
    Debug.Print "Supports Int Timeouts: " & objItem.SupportsIntTimeouts
    Debug.Print "Supports Parity Check: " & objItem.SupportsParityCheck
    Debug.Print "Supports RLSD: " & objItem.SupportsRLSD
    Debug.Print "Supports RTSCTS: " & objItem.SupportsRTSCTS
    Debug.Print "Supports Special Characters: " & _
        objItem.SupportsSpecialCharacters
    Debug.Print "Supports XOn XOff: " & objItem.SupportsXOnXOff
    Debug.Print "Supports XOn XOff Setting: " & objItem.SupportsXOnXOffSet
Next


End Sub
