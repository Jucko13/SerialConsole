VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "ZebroMote - V1.0 by Ricardo de Roode"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14775
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
   ScaleHeight     =   394
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   985
   StartUpPosition =   3  'Windows Default
   Begin Project1.uFrame frmColors 
      Height          =   615
      Left            =   180
      TabIndex        =   20
      Top             =   4665
      Visible         =   0   'False
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   1085
      BackgroundColor =   -2147483633
      BorderColor     =   8421504
      Caption         =   "Colors for Led 1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox picColors 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   420
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   225
         Width           =   450
      End
   End
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
      Left            =   9615
      Max             =   10
      TabIndex        =   5
      Top             =   2130
      Visible         =   0   'False
      Width           =   255
   End
   Begin Project1.uTextBox txtStatus 
      Height          =   300
      Left            =   180
      TabIndex        =   4
      Top             =   5445
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
      Left            =   12780
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
      Left            =   10575
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
      Width           =   6360
      _ExtentX        =   11218
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
      Left            =   6615
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
      Left            =   13725
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
      Caption         =   "Leds Off"
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
   Begin Project1.uButton cmdControls 
      Height          =   1410
      Index           =   10
      Left            =   4635
      TabIndex        =   22
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
      Caption         =   "u"
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
      Index           =   11
      Left            =   4635
      TabIndex        =   23
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
      Caption         =   "t"
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim picoSendCommand(0 To 19) As Boolean
Dim picoConnected(0 To 19) As Boolean


Dim lastMessageBytes() As Byte

Dim wmi As WbemScripting.SWbemServices

Dim serialDevices() As SerialDevice

Dim receiveBuffer As String

Dim ledCommand As Long

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
                Case 1
                    sendCommand i, 1, 32, 2, 255
                    
                Case 10
                    sendCommand i, 1, 32, 4, 255
                Case 11
                    sendCommand i, 1, 32, 3, 255
                    
                Case 3
                    sendCommand i, 1, 32, 0, 255
                
                Case 2
                    Dim j As Byte
                    For j = 0 To 5
                        sendCommand i, 1, 33 + j, 0, 255
                        DoEvents
                        wait 100
                    Next j
            End Select
        End If
    Next i
    
    
    Select Case Index
        Case 4 To 9
            
            If ledCommand = Index - 4 Then
                ledCommand = -1
                frmColors.Visible = False
            Else
                ledCommand = Index - 4
                frmColors.Caption = "Colors for Led " & ledCommand + 1
                frmColors.Visible = True
                
            End If
    End Select
    
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
            If errorValue = 0 Then
                setStatus "Command successfull!"
            Else
                setStatus "Arduino Error: " & IIf(errorValue > 0 And errorValue < 6, errorMessages(errorValue), "UNKNOWN_ERROR"), True, errorValue
            End If
            
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

Sub fillLedColors()
    Const strColors As String = "&h0, &HFF0000,&HFF00,&HFFFF00,&HFF,&HFF00FF,&HFFFF,&HFFFFFF"
    Dim splColors() As String
    
    splColors = Split(strColors, ",")
    
    Dim i As Long
    
    picColors(0).BackColor = CLng(splColors(0))
    
    For i = 1 To UBound(splColors)
        Load picColors(i)
        picColors(i).BackColor = CLng(splColors(i))
        picColors(i).Left = picColors(i - 1).Left + picColors(i - 1).Width + Screen.TwipsPerPixelX * 5
        
        picColors(i).Visible = True
    Next i
    
End Sub

Private Sub Form_Load()
    uDontDrawDots = True
    frmColors.Redraw
    
    fillCommportList
    
    fillBaudList
    
    fillZebroButtons
    
    fillLedColors
    
    errorMessages = Split(errorMessagesConst, ",")
    
    comm.OutBufferSize = 5
    
    On Error Resume Next
    drpBaud.ListIndex = GetSetting("SerialConsole", "dropdown", "drpBaud.ListIndex", 0)
    drpCommports.ListIndex = GetSetting("SerialConsole", "dropdown", "drpCommports.ListIndex", 0)
    
    ledCommand = -1
    
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
    
    Dim colItems As Object, objItem As Object
    
    'ReDim serialDevices(0 to colItems
    
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\WMI") '\\.\root\cimv2
    
    Set colItems = wmi.ExecQuery("SELECT * from MSSerial_PortName") ' WHERE Name LIKE '%(COM%' ''' 'Win32_SerialPort
    
    
    'Set objWMIService =
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_SerialPort")
    
    ReDim serialDevices(0)
    Dim itemCount As Long
    
    If colItems Is Nothing Then Exit Sub
    
    For Each objItem In colItems
        ReDim Preserve serialDevices(0 To itemCount)
        
        With serialDevices(itemCount)
            .commport = objItem.PortName
            .instanceName = objItem.instanceName
            fillDeviceProperties serialDevices(itemCount)
            
        End With
        itemCount = itemCount + 1
    Next
    
    Dim i As Long
    
    Dim prevIndex As Long
    
    prevIndex = drpCommports.ListIndex
    
    drpCommports.Clear
    
    For i = 0 To itemCount - 1
        drpCommports.AddItem serialDevices(i).friendlyName & " (" & serialDevices(i).commport & ", " & serialDevices(i).locationInformation & ")", i, , -1
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

Private Sub picColors_Click(Index As Integer)
    Dim i As Byte
    
    For i = 0 To UBound(picoSendCommand)
        If picoSendCommand(i) And picoConnected(i) Then
            sendCommand i, 1, 33 + ledCommand, CByte(Index), 255
        End If
    Next i
    
    ledCommand = -1
    frmColors.Visible = False
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

Function getDeviceErrorStatusMessage(errorCode As Long) As String
    Dim msg As String
    Select Case errorCode
        Case 0: msg = "This device is working properly."
        Case 1: msg = "This device is not configured correctly."
        Case 2: msg = "Windows cannot load the driver for this device."
        Case 3: msg = "The driver might be corrupted, or your system " & _
                       "may be running low on memory or other resources."
        Case 4: msg = "This device is not working properly. One of its " & _
                       "drivers or your registry might be corrupted."
        Case 5: msg = "The driver for this device needs a resource " & _
                       "that Windows cannot manage."
        Case 6: msg = "The boot configuration for this device " & _
                      "conflicts with other devices."
        Case 7: msg = "Cannot filter."
        Case 8: msg = "The driver loader for the device is missing."
        Case 9: msg = "This device is not working properly because" & _
                      "the controlling firmware is reporting the " & _
                      "resources for the device incorrectly."
        Case 10: msg = "This device cannot start."
        Case 11: msg = "This device failed."
        Case 12: msg = "This device cannot find enough free " & _
                       "resources that it can use."
        Case 13: msg = "Windows cannot verify this device's resources."
        Case 14: msg = "This device cannot work properly until " & _
                       "you restart your computer."
        Case 15: msg = "This device is not working properly because " & _
                       "there is probably a re-enumeration problem."
        Case 16: msg = "Windows cannot identify all the resources this device uses."
        Case 17: msg = "This device is asking for an unknown resource type."
        Case 18: msg = "Reinstall the drivers for this device."
        Case 19: msg = "Failure using the VXD loader."
        Case 20: msg = "Your registry might be corrupted."
        Case 21: msg = "System failure: Try changing the driver for this device. " & _
                       "If that does not work, see your hardware " & _
                       "documentation. Windows is removing this device."
        Case 22: msg = "This device is disabled."
        Case 23: msg = "System failure: Try changing the driver for " & _
                       "this device. If that doesn't work, see your " & _
                       "hardware documentation."
        Case 24: msg = "This device is not present, is not working " & _
                       "properly, or does not have all its drivers installed."
        Case 25: msg = "Windows is still setting up this device."
        Case 26: msg = "Windows is still setting up this device."
        Case 27: msg = "This device does not have valid log configuration."
        Case 28: msg = "The drivers for this device are not installed."
        Case 29: msg = "This device is disabled because the firmware of " & _
                       "the device did not give it the required resources."
        Case 30: msg = "This device is using an Interrupt Request (IRQ) " & _
                       "resource that another device is using."
        Case 31: msg = "This device is not working properly because Windows " & _
                       "cannot load the drivers required for this device."
    End Select


    getDeviceErrorStatusMessage = msg
End Function
