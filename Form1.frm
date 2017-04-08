VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Zebromote - V1.0 by Ricardo de Roode"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14280
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
   ScaleWidth      =   952
   StartUpPosition =   3  'Windows Default
   Begin Project1.uButton cmdControls 
      Height          =   1410
      Index           =   0
      Left            =   1665
      TabIndex        =   10
      Top             =   1710
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   2487
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
      FocusVisible    =   0   'False
      Caption         =   "p"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrGetConnected 
      Interval        =   500
      Left            =   9180
      Top             =   915
   End
   Begin VB.PictureBox picConnected 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   180
      ScaleHeight     =   255
      ScaleWidth      =   390
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   810
      Width           =   420
   End
   Begin VB.Timer tmrCheckPortOpen 
      Interval        =   1000
      Left            =   8205
      Top             =   930
   End
   Begin VB.VScrollBar scrollTest 
      Height          =   3015
      LargeChange     =   20
      Left            =   9630
      Max             =   10
      TabIndex        =   5
      Top             =   3465
      Visible         =   0   'False
      Width           =   255
   End
   Begin Project1.uTextBox txtStatus 
      Height          =   300
      Left            =   90
      TabIndex        =   4
      Top             =   6330
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
   Begin Project1.uCheckBox chkCommOptions 
      Height          =   465
      Index           =   0
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
      CheckSelectionColor=   4210752
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
      Left            =   10080
      TabIndex        =   2
      Top             =   180
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   820
      BackgroundColor =   12632319
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
      FocusVisible    =   0   'False
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
      Width           =   5865
      _ExtentX        =   10345
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
      Left            =   7965
      Top             =   1665
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      OutBufferSize   =   1
      ParityReplace   =   0
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin Project1.uDropDown drpBaud 
      Height          =   465
      Left            =   6120
      TabIndex        =   1
      Top             =   180
      Width           =   3885
      _ExtentX        =   6853
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
      Height          =   810
      Left            =   10485
      TabIndex        =   6
      Top             =   2805
      Visible         =   0   'False
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   1429
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
   Begin Project1.uCheckBox chkCommOptions 
      Height          =   465
      Index           =   1
      Left            =   13230
      TabIndex        =   7
      Top             =   180
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   820
      BackgroundColor =   14737632
      BorderColor     =   8421504
      Caption         =   "RTS"
      CheckBorderColor=   8421504
      CheckSelectionColor=   4210752
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
   Begin Project1.uButton cmdZebro 
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   1170
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   661
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
      FocusVisible    =   0   'False
      Caption         =   "0"
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
   Begin Project1.uButton cmdControls 
      Height          =   1410
      Index           =   1
      Left            =   3150
      TabIndex        =   11
      Top             =   1710
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   2487
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
      FocusVisible    =   0   'False
      Caption         =   "q"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdControls 
      Height          =   1410
      Index           =   2
      Left            =   180
      TabIndex        =   12
      Top             =   1710
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   2487
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
      FocusVisible    =   0   'False
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdControls 
      Height          =   1410
      Index           =   3
      Left            =   1665
      TabIndex        =   13
      Top             =   3195
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   2487
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
      FocusVisible    =   0   'False
      Caption         =   "STOP"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdControls 
      Height          =   420
      Index           =   4
      Left            =   180
      TabIndex        =   14
      Top             =   3195
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
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
      FocusVisible    =   0   'False
      Caption         =   "Led 1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionOffsetTop=   -1
   End
   Begin Project1.uButton cmdControls 
      Height          =   420
      Index           =   5
      Left            =   180
      TabIndex        =   15
      Top             =   3690
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
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
      FocusVisible    =   0   'False
      Caption         =   "Led 2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionOffsetTop=   -1
   End
   Begin Project1.uButton cmdControls 
      Height          =   420
      Index           =   6
      Left            =   180
      TabIndex        =   16
      Top             =   4185
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
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
      FocusVisible    =   0   'False
      Caption         =   "Led 3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionOffsetTop=   -1
   End
   Begin Project1.uButton cmdControls 
      Height          =   420
      Index           =   7
      Left            =   3150
      TabIndex        =   17
      Top             =   3195
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
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
      FocusVisible    =   0   'False
      Caption         =   "Led 4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionOffsetTop=   -1
   End
   Begin Project1.uButton cmdControls 
      Height          =   420
      Index           =   8
      Left            =   3150
      TabIndex        =   18
      Top             =   3690
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
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
      FocusVisible    =   0   'False
      Caption         =   "Led 5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionOffsetTop=   -1
   End
   Begin Project1.uButton cmdControls 
      Height          =   420
      Index           =   9
      Left            =   3150
      TabIndex        =   19
      Top             =   4185
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
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
      FocusVisible    =   0   'False
      Caption         =   "Led 6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionOffsetTop=   -1
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

Dim picoSendCommand(0 To 19) As Boolean
Dim picoConnected(0 To 19) As Boolean


Dim lastMessageBytes() As Byte

Dim wmi As WbemScripting.SWbemServices

Dim serialDevices() As SerialDevice

Dim receiveBuffer As String

Dim errorMessages() As String
Const errorMessagesConst = ",M_ERROR,M_ERROR_NOT_CONNECTED,M_ERROR_BUFFER_OVERFLOW,M_ERROR_BUFFER_EMPTY,M_ERROR_UNKNOWN_COMMAND"

Private Sub chkCommOptions_ActivateNextState(Index As Integer, u_Cancel As Boolean, u_NewState As uCheckboxConstants)
    Dim newState As Boolean
    
    newState = (u_NewState = u_Checked)
    
    Select Case Index
    
        Case 0
            comm.DTREnable = newState
            
        Case 1
            comm.RTSEnable = newState
    End Select
    
    
End Sub

Private Sub cmdConnect_Click(Button As Integer, x As Single, y As Single)
    On Error GoTo notWorking
    
    If comm.PortOpen Then
        comm.PortOpen = False
        cmdConnect.Caption = "Connect"
        cmdConnect.BackgroundColor = &HC0C0FF
        
    Else
        cmdConnect.BackgroundColor = &HC0FFC0
        cmdConnect.Caption = "Disconnect"
        comm.PortOpen = True
        
        setStatus "Connected!"
    End If
    
Exit Sub
notWorking:
    cmdConnect.BackgroundColor = &HC0C0FF
    cmdConnect.Caption = "Connect"
    
    setStatus Err.Description, True, Err.Number
    
End Sub

Sub setStatus(msg As String, Optional isError As Boolean = False, Optional errorNumber As Long = 0)
    txtStatus.RedrawPause
    
    If isError Then
        txtStatus.Text = "[ ERROR " & errorNumber & " ] " & msg
        txtStatus.ForeColor = vbRed
    Else
        txtStatus.Text = msg
        txtStatus.ForeColor = vbBlack
    End If
    
    txtStatus.RedrawResume
End Sub

Private Sub cmdControls_Click(Index As Integer, Button As Integer, x As Single, y As Single)
    
    Dim i As Byte
    
    For i = 0 To UBound(picoSendCommand)
        If picoSendCommand(i) And picoConnected(i) Then
            Select Case Index
            
                Case 0
                    sendCommand i, 1, 32, 1, 255
            
                Case 3
                    sendCommand i, 1, 32, 0, 255
                
                Case 4
                    sendCommand i, 1, 33, Rnd * 7, 255
                    
            End Select
        End If
    Next i
    
End Sub

Private Sub cmdZebro_Click(Index As Integer, Button As Integer, x As Single, y As Single)
        
    picoSendCommand(Index) = Not picoSendCommand(Index)
    
    
    Dim i As Long
    
    For i = 0 To UBound(picoSendCommand)
        cmdZebro(i).BackgroundColor = IIf(picoSendCommand(i), &HFFC0C0, &HE0E0E0)
        cmdZebro(i).MouseOverBackgroundColor = IIf(picoSendCommand(i), &HFF8080, &HC0C0C0)
    Next i
    
End Sub

Private Sub comm_OnComm()
    'On Error Resume Next
    
    Select Case comm.CommEvent
    
        Case 2   ' comEvReceive event occured
        
            receiveBuffer = receiveBuffer & comm.Input
            
            If InStr(1, receiveBuffer, Chr(255)) > 0 Then
                processIncommingMessage
            End If
            
            'Debug.Print UBound(receiveBuffer)
            
'            Do While comm.InBufferCount > 0
'                txtReceived.SelStart = txtReceived.TextLength
'                txtReceived.SelLength = 0
'
'                txtReceived.AddCharAtCursor comm.Input
'            Loop
        
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

Sub processIncommingMessage()
    Dim i As Long
    Dim tmpBytes() As Long
    
    ReDim tmpBytes(0)
    
    printBuffer
    
    Dim msg As String
    Dim firstPlace As String
    firstPlace = InStr(1, receiveBuffer, Chr(255))
    If firstPlace = 0 Then Exit Sub
    
    msg = Left$(receiveBuffer, firstPlace)
    
    receiveBuffer = Right$(receiveBuffer, Len(receiveBuffer) - firstPlace)
    
    Select Case Len(msg)
        Case 22
            For i = 0 To 19
                picoConnected(i) = IIf(Asc(Mid$(msg, i + 2, 1)) = 1, True, False)
            Next i
            refreshConnected
        
        Case 2
            Dim errorValue As Long
            errorValue = Asc(Left$(msg, 1))
            setStatus "Arduino Error: " & IIf(errorValue > 0 And errorValue < 6, errorMessages(errorValue), "UNKNOWN_ERROR"), True, errorValue
            
    End Select
    
    
    Debug.Print UBound(tmpBytes)
    
    processIncommingMessage
End Sub

Sub printBuffer()
    Dim i As Long
    
    Dim tmp As String
    
    tmp = "receiveBuffer = {"
    
    For i = 1 To Len(receiveBuffer)
        tmp = tmp & "0x" & Hex(Asc(Mid$(receiveBuffer, i, 1))) & ", "
    Next i
    
    tmp = tmp & "}"
    
    Debug.Print tmp
End Sub


Private Sub drpBaud_ItemChange(ItemIndex As Long)
    comm.Settings = drpBaud.List(ItemIndex) & ",n,8,1"
End Sub

Private Sub drpCommports_ItemChange(ItemIndex As Long)
    On Error GoTo notWorking
    
    comm.commport = Replace(serialDevices(ItemIndex).commport, "COM", "")
    Exit Sub
notWorking:
    setStatus Err.Description, True, Err.Number
    
End Sub

Private Sub drpCommports_OnDropdown()
    fillCommportList
End Sub

Sub fillZebroButtons()
    Dim i As Long

    For i = 1 To 19
        Load cmdZebro(i)
        cmdZebro(i).Left = cmdZebro(i - 1).Left + cmdZebro(i - 1).Width + 5
        cmdZebro(i).Visible = True
        cmdZebro(i).Caption = i
        
        Load picConnected(i)
        picConnected(i).Left = picConnected(i - 1).Left + picConnected(i - 1).Width + 5
        picConnected(i).Visible = True
        picConnected(i).BackColor = &HC0C0FF
    Next i

    refreshConnected
End Sub

Sub refreshConnected()
    Dim i As Long

    For i = 0 To 19
        picConnected(i).BackColor = IIf(picoConnected(i), &HC0FFC0, &HC0C0FF)
    Next i
    
End Sub

Private Sub Form_Load()
    
    fillCommportList
    
    fillBaudList
    
    fillZebroButtons
    
    errorMessages = Split(errorMessagesConst, ",")
    
    comm.OutBufferSize = 5
    
    On Error Resume Next
    drpBaud.ListIndex = GetSetting("SerialConsole", "dropdown", "drpBaud.ListIndex", 0)
    drpCommports.ListIndex = GetSetting("SerialConsole", "dropdown", "drpCommports.ListIndex", 0)
    
    
    'txtReceived.ReCalculateMarkup
    'txtReceived.ReCalculateWords
    'txtReceived.Redraw
    
End Sub



Private Sub Form_Resize()
    txtReceived.Left = 0
    txtReceived.Width = Me.ScaleWidth
    txtReceived.Height = Me.ScaleHeight - txtReceived.Top - txtStatus.Height - 32
    
    txtStatus.Top = Me.ScaleHeight - txtStatus.Height - 12
    txtStatus.Left = 12
    txtStatus.Width = txtReceived.Width - 24
    
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
'On Error Resume Next
    
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
    
    Dim prevIndex As Long
    
    prevIndex = drpCommports.ListIndex
    
    drpCommports.Clear
    
    For i = 0 To itemCount - 1
        drpCommports.AddItem serialDevices(i).name & " (" & serialDevices(i).commport & ", " & serialDevices(i).baudMax & ")", i, , -1
    Next i
    
    drpCommports.ItemsVisible = itemCount
    
    If prevIndex <> -1 And drpCommports.ListCount < prevIndex Then
        drpCommports.ListIndex = prevIndex
    Else
        If itemCount <> 0 Then drpCommports.ListIndex = 0
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "SerialConsole", "dropdown", "drpBaud.ListIndex", drpBaud.ListIndex
    SaveSetting "SerialConsole", "dropdown", "drpCommports.ListIndex", drpCommports.ListIndex
    
End Sub

Private Sub scrollTest_Scroll()
    txtReceived.m_lScrollTop = scrollTest.Value
    txtReceived.Redraw
    
End Sub

Private Sub tmrCheckPortOpen_Timer()
    'txtStatus.Text = comm.CommEvent
End Sub

Private Sub tmrGetConnected_Timer()
    sendCommand 20, 0, 0, 0, 255
End Sub

Sub sendCommand(byte0 As Byte, byte1 As Byte, byte2 As Byte, byte3 As Byte, byte4 As Byte)
    Dim bytes(0 To 4) As Byte
    Dim variantBytes As Variant
    
    If Not comm.PortOpen Then Exit Sub
    
    bytes(0) = byte0
    bytes(1) = byte1
    bytes(2) = byte2
    bytes(3) = byte3
    bytes(4) = byte4
    
    lastMessageBytes = bytes
    variantBytes = bytes
    
    comm.Output = variantBytes
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
