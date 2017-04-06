VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1189
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar scrollTest 
      Height          =   3015
      LargeChange     =   20
      Left            =   330
      Max             =   10
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin Project1.uTextBox txtStatus 
      Height          =   300
      Left            =   180
      TabIndex        =   4
      Top             =   4035
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   529
      BorderColor     =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uCheckBox chkDTR 
      Height          =   465
      Left            =   12285
      TabIndex        =   3
      Top             =   180
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   820
      BackgroundColor =   14737632
      BorderColor     =   8421504
      Caption         =   "DTR"
      CheckBorderColor=   8421504
      CheckSelectionColor=   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdConnect 
      Height          =   465
      Left            =   9945
      TabIndex        =   2
      Top             =   180
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   820
      BackgroundColor =   14737632
      BorderColor     =   8421504
      ForeColor       =   0
      MouseOverBackgroundColor=   12632256
      FocusColor      =   12632256
      BackgroundColorDisabled=   14737632
      BorderColorDisabled=   8421504
      ForeColorDisabled=   0
      MouseOverBackgroundColorDisabled=   12632256
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   12632256
      Caption         =   "Connect"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uDropDown drpCommports 
      Height          =   465
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   820
      BackgroundColor =   14737632
      BorderColor     =   8421504
      SelectionBackgroundColor=   14737632
      SelectionBorderColor=   14737632
      BackgroundColorDisabled=   14737632
      BorderColorDisabled=   8421504
      SelectionBackgroundColorDisabled=   14737632
      SelectionBorderColorDisabled=   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "drpCommports met een erg lange zin er achter aan"
      ItemHeight      =   20
      ScrollBarWidth  =   30
   End
   Begin MSCommLib.MSComm comm 
      Left            =   10995
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin Project1.uDropDown drpBaud 
      Height          =   465
      Left            =   6930
      TabIndex        =   1
      Top             =   180
      Width           =   2820
      _ExtentX        =   4974
      _ExtentY        =   820
      BackgroundColor =   14737632
      BorderColor     =   8421504
      SelectionBackgroundColor=   14737632
      SelectionBorderColor=   14737632
      BackgroundColorDisabled=   14737632
      BorderColorDisabled=   8421504
      SelectionBackgroundColorDisabled=   14737632
      SelectionBorderColorDisabled=   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "drpCommports met een erg lange zin er achter aan"
      ItemHeight      =   20
      ScrollBarWidth  =   30
   End
   Begin Project1.uTextBox txtReceived 
      Height          =   7980
      Left            =   180
      TabIndex        =   6
      Top             =   855
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   14076
      BackgroundColor =   0
      BorderColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65280
      RowLineColor    =   14737632
      RowNumberOnEveryLine=   -1  'True
      WordWrap        =   -1  'True
      MultiLine       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type SerialDevice
    name As String
    commport As String
    baudMax As String
End Type


Dim wmi As WbemScripting.SWbemServices

Dim serialDevices() As SerialDevice

Private Sub chkDTR_ActivateNextState(u_Cancel As Boolean, u_NewState As uCheckboxConstants)
    comm.DTREnable = (u_NewState = u_Checked)
    
End Sub

Private Sub cmdConnect_Click(Button As Integer, x As Single, y As Single)
    comm.PortOpen = True
End Sub

Private Sub comm_OnComm()
    Select Case comm.CommEvent
    
        Case 2   ' comEvReceive event occured
        
            Do While comm.InBufferCount > 0
                txtReceived.SelStart = txtReceived.TextLength
                txtReceived.SelLength = 0
                
                txtReceived.AddCharAtCursor comm.Input
            Loop
        
        Case Is > 1000
            
            ' The CommEvent property always returns a numerical value.
            ' Whenever the CommEvent property returns a number
            ' above 1000 then you know that an error occurred.
            txtStatus.Text = "Some ComPort Error occurred"
        Case Else
            ' What happened? It wasn't the arrival of data - and it wasn't
            ' an error. See the ' CommEvent property for a full listing
            ' of all the events and errors.
   End Select
   
   
End Sub

Private Sub drpBaud_ItemChange(ItemIndex As Long)
    comm.Settings = drpBaud.List(ItemIndex) & ",n,8,1"
End Sub

Private Sub drpCommports_ItemChange(ItemIndex As Long)
    comm.commport = Replace(serialDevices(ItemIndex).commport, "COM", "")
End Sub

Private Sub Form_Load()
    
    fillCommportList
    
    fillBaudList
    
    
    'txtReceived.ReCalculateMarkup
    'txtReceived.ReCalculateWords
    'txtReceived.Redraw
    
End Sub



Private Sub Form_Resize()
    txtReceived.Left = 0
    txtReceived.Width = Me.ScaleWidth
    txtReceived.Height = Me.ScaleHeight - txtReceived.Top - txtStatus.Height - 32
    
    txtStatus.Top = Me.ScaleHeight - txtStatus.Height - 16
    txtStatus.Left = 0
    txtStatus.Width = txtReceived.Width
    
End Sub


Sub fillBaudList()
drpBaud.Clear

Const bauds As String = "300,600,1200,2400,4800,9600,14400,19200,28800,38400,56000,57600,115200,128000,250000"
Dim tmpSplit() As String
Dim i As Long

tmpSplit = Split(bauds, ",")

For i = 0 To UBound(tmpSplit)
    drpBaud.AddItem tmpSplit(i)
   
    If tmpSplit(i) = "9600" Then
        drpBaud.ListIndex = i
    End If

Next i

drpBaud.ItemsVisible = UBound(tmpSplit) + 1
 
End Sub




Sub fillCommportList()
On Error Resume Next
    
    Dim colItems, objItem
    
    'ReDim serialDevices(0 to colItems
    
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    Set colItems = wmi.ExecQuery("Select * from Win32_SerialPort")
    
    
    'Set objWMIService =
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_SerialPort")
    
    ReDim serialDevices(0)
    Dim itemCount As Long
    
    
    
    For Each objItem In colItems
        ReDim Preserve serialDevices(0 To itemCount)
        
        With serialDevices(itemCount)
            .commport = objItem.DeviceID
            .name = objItem.Description
            .baudMax = objItem.MaxBaudRate
            
        End With
    '
    '    Debug.Print "Binary: " & objItem.Binary
    '    Debug.Print "Description: " & objItem.Description
    '    Debug.Print "Device ID: " & objItem.DeviceID
    '    Debug.Print "Maximum Baud Rate: " & objItem.MaxBaudRate
    '    Debug.Print "Maximum Input Buffer Size: " & objItem.MaximumInputBufferSize
    '    Debug.Print "Maximum Output Buffer Size: " & objItem.MaximumOutputBufferSize
    '    Debug.Print "Name: " & objItem.name
    '    Debug.Print "OS Auto Discovered: " & objItem.OSAutoDiscovered
    '    Debug.Print "PNP Device ID: " & objItem.PNPDeviceID
    '    Debug.Print "Provider Type: " & objItem.ProviderType
    '    Debug.Print "Settable Baud Rate: " & objItem.SettableBaudRate
    '    Debug.Print "Settable Data Bits: " & objItem.SettableDataBits
    '    Debug.Print "Settable Flow Control: " & objItem.SettableFlowControl
    '    Debug.Print "Settable Parity: " & objItem.SettableParity
    '    Debug.Print "Settable Parity Check: " & objItem.SettableParityCheck
    '    Debug.Print "Settable RLSD: " & objItem.SettableRLSD
    '    Debug.Print "Settable Stop Bits: " & objItem.SettableStopBits
    '    Debug.Print "Supports 16-Bit Mode: " & objItem.Supports16BitMode
    '    Debug.Print "Supports DTRDSR: " & objItem.SupportsDTRDSR
    '    Debug.Print "Supports Elapsed Timeouts: " & objItem.SupportsElapsedTimeouts
    '    Debug.Print "Supports Int Timeouts: " & objItem.SupportsIntTimeouts
    '    Debug.Print "Supports Parity Check: " & objItem.SupportsParityCheck
    '    Debug.Print "Supports RLSD: " & objItem.SupportsRLSD
    '    Debug.Print "Supports RTSCTS: " & objItem.SupportsRTSCTS
    '    Debug.Print "Supports Special Characters: " & objItem.SupportsSpecialCharacters
    '    Debug.Print "Supports XOn XOff: " & objItem.SupportsXOnXOff
    '    Debug.Print "Supports XOn XOff Setting: " & objItem.SupportsXOnXOffSet
    '
        itemCount = itemCount + 1
    Next
    
    Dim i As Long
    
    drpCommports.Clear
    
    For i = 0 To itemCount - 1
        drpCommports.AddItem serialDevices(i).name & " (" & serialDevices(i).commport & ", " & serialDevices(i).baudMax & ")", i, , -1
    Next i
    
    drpCommports.ItemsVisible = itemCount

End Sub


Private Sub scrollTest_Scroll()
    txtReceived.m_lScrollTop = scrollTest.Value
    txtReceived.Redraw
    
End Sub

Private Sub txtReceived_Changed()
    Dim i As Long
    Dim s As String
    Dim t As String


    s = txtReceived.Text
    txtReceived.RedrawPause

    For i = 1 To Len(s) - 1
        t = Mid$(s, i, 1)
        txtReceived.setCharBackColor i - 1, -1

        Select Case t


            Case "S", "s"
                txtReceived.setCharBackColor i - 1, vbRed

            Case "a", "A"
                txtReceived.setCharBackColor i - 1, vbBlue

        End Select

    Next i
    
    
    txtReceived.RedrawResume
End Sub

Private Sub txtReceived_KeyDown(KeyCode As Integer, Shift As Integer)
    'KeyCode = 0
    'Shift = 0
End Sub

Private Sub txtReceived_KeyPress(KeyAscii As Integer)
    'KeyAscii = 0
End Sub
