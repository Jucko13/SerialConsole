VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0024211E&
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   555
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1196
   StartUpPosition =   2  'CenterScreen
   Begin SerialConsole.uToolTip ttToolTip 
      Height          =   420
      Left            =   5130
      TabIndex        =   93
      Top             =   3420
      Visible         =   0   'False
      Width           =   780
      _ExtentX        =   7938
      _ExtentY        =   7938
   End
   Begin SerialConsole.uButton cmdConnect 
      Height          =   360
      Left            =   8295
      TabIndex        =   3
      Top             =   225
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      BackgroundColor =   4671472
      BorderColor     =   8421504
      ForeColor       =   16777215
      MouseOverBackgroundColor=   8882165
      FocusColor      =   12632256
      BackgroundColorDisabled=   14737632
      BorderColorDisabled=   8421504
      ForeColorDisabled=   0
      MouseOverBackgroundColorDisabled=   12632256
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   12632256
      FocusVisible    =   0   'False
      Caption         =   "Connect"
      Border          =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SerialConsole.uLoadBar loadReconnect 
      Height          =   450
      Left            =   8250
      TabIndex        =   91
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   794
      BackgroundColor =   4671472
      BarColor        =   2367774
      BarType         =   0
      BarWidth        =   0
      Border          =   0   'False
      Caption         =   ""
      CaptionType     =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LoadingSpeed    =   1
      Value           =   0
   End
   Begin SerialConsole.uFrame frmReconnectSettings 
      Height          =   1785
      Left            =   3765
      TabIndex        =   83
      Top             =   1230
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   3149
      BackgroundColor =   2367774
      BorderColor     =   14737632
      ForeColor       =   16777215
      Caption         =   "Reconnect Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin SerialConsole.uCheckBox chkComOptions 
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   84
         Top             =   240
         Width           =   2130
         _ExtentX        =   3572
         _ExtentY        =   344
         BackgroundColor =   2367774
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "Clear On Connect"
         CaptionOffsetLeft=   5
         CaptionOffsetTop=   2
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   0
         CheckOffsetLeft =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
      End
      Begin SerialConsole.uCheckBox chkComOptions 
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   85
         Top             =   510
         Width           =   2130
         _ExtentX        =   3387
         _ExtentY        =   344
         BackgroundColor =   2367774
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "Auto Disconnect"
         CaptionOffsetLeft=   5
         CaptionOffsetTop=   2
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   0
         CheckOffsetLeft =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
      End
      Begin SerialConsole.uCheckBox chkComOptions 
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   86
         Top             =   780
         Width           =   2130
         _ExtentX        =   2831
         _ExtentY        =   344
         BackgroundColor =   2367774
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "Auto Connect"
         CaptionOffsetLeft=   5
         CaptionOffsetTop=   2
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   0
         CheckOffsetLeft =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
      End
      Begin SerialConsole.uCheckBox chkComOptions 
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   87
         Top             =   1050
         Width           =   2130
         _ExtentX        =   3572
         _ExtentY        =   344
         BackgroundColor =   2367774
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "Auto Connect USB"
         CaptionOffsetLeft=   5
         CaptionOffsetTop=   2
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   0
         CheckOffsetLeft =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
      End
      Begin SerialConsole.uCheckBox chkComOptions 
         Height          =   375
         Index           =   6
         Left            =   90
         TabIndex        =   92
         Top             =   1320
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   661
         BackgroundColor =   2367774
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "Auto disconnect on"
         CaptionOffsetLeft=   5
         CaptionOffsetTop=   2
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   0
         CheckOffsetLeft =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
      End
   End
   Begin VB.PictureBox picConnectionSettings 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0024211E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   14040
      MouseIcon       =   "Form1.frx":1542
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":1694
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   82
      TabStop         =   0   'False
      ToolTipText     =   "Connect/Reconnect/Arduino Settings"
      Top             =   165
      Width           =   480
   End
   Begin VB.Timer tmrCheckUsbStillConnected 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   15525
      Top             =   855
   End
   Begin SerialConsole.uFrame frmWindow 
      Height          =   1485
      Index           =   2
      Left            =   13560
      TabIndex        =   70
      Top             =   1395
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   2619
      BackgroundColor =   2367774
      BorderColor     =   14737632
      ForeColor       =   16777215
      Caption         =   "Graph"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin SerialConsole.uGraph graphArduino 
         Height          =   705
         Left            =   90
         TabIndex        =   71
         Top             =   510
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   1244
      End
      Begin SerialConsole.uCheckBox chkEnableGraph 
         Height          =   165
         Left            =   90
         TabIndex        =   74
         ToolTipText     =   "Show the received data in hex (Hold: H)"
         Top             =   240
         Width           =   1905
         _ExtentX        =   1191
         _ExtentY        =   291
         BackgroundColor =   2367774
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "Enable Graph"
         CaptionOffsetLeft=   5
         CaptionOffsetTop=   1
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   0
         CheckOffsetLeft =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
         AutoSize        =   0   'False
      End
   End
   Begin SerialConsole.uFrame frmWindow 
      Height          =   885
      Index           =   1
      Left            =   15930
      TabIndex        =   61
      Top             =   1455
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1561
      BackgroundColor =   2367774
      BorderColor     =   14737632
      ForeColor       =   16777215
      Caption         =   "ArduinoWindow"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin SerialConsole.uListBox lstArduino 
         Height          =   405
         Left            =   90
         TabIndex        =   62
         Top             =   285
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   714
         BackgroundColor =   3551534
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
         ForeColor       =   16777215
         Text            =   "uFrame"
         SelectionBackgroundColor=   3551534
         SelectionBorderColor=   16777215
         SelectionForeColor=   12648384
         ItemHeight      =   2
         VisibleItems    =   15
      End
   End
   Begin VB.Timer tmrCheckForReconnect 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   14175
      Top             =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Index           =   1
      Left            =   1530
      TabIndex        =   54
      Top             =   900
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Index           =   0
      Left            =   270
      TabIndex        =   53
      Top             =   900
      Visible         =   0   'False
      Width           =   1230
   End
   Begin SerialConsole.uFrame frmTxtSettings 
      Height          =   435
      Left            =   195
      TabIndex        =   49
      Top             =   4320
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   767
      BackgroundColor =   2367774
      BorderColor     =   14737632
      ForeColor       =   16777215
      Caption         =   "Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin SerialConsole.uDropDown drpReceiveSpeed 
         Height          =   240
         Left            =   4485
         TabIndex        =   65
         Top             =   150
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   423
         BackgroundColor =   14322034
         BorderColor     =   14322034
         ForeColor       =   16777215
         SelectionBackgroundColor=   13592135
         SelectionBorderColor=   14322034
         BackgroundColorDisabled=   14737632
         BorderColorDisabled=   8421504
         SelectionBackgroundColorDisabled=   14737632
         SelectionBorderColorDisabled=   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Ultra Slow"
         Border          =   0   'False
         ItemHeight      =   19
         VisibleItems    =   8
         ScrollBarWidth  =   19
      End
      Begin SerialConsole.uCheckBox chkTxtSettings 
         Height          =   165
         Index           =   0
         Left            =   90
         TabIndex        =   51
         Top             =   195
         Width           =   1185
         _ExtentX        =   2196
         _ExtentY        =   291
         BackgroundColor =   2367774
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "AutoScroll"
         CaptionOffsetLeft=   5
         CaptionOffsetTop=   1
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   0
         CheckOffsetLeft =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
      End
      Begin SerialConsole.uCheckBox chkTxtSettings 
         Height          =   165
         Index           =   1
         Left            =   1350
         TabIndex        =   52
         Top             =   195
         Width           =   1365
         _ExtentX        =   2672
         _ExtentY        =   291
         BackgroundColor =   2367774
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "ConsoleColors"
         CaptionOffsetLeft=   5
         CaptionOffsetTop=   1
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   0
         CheckOffsetLeft =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
      End
      Begin SerialConsole.uCheckBox chkTxtSettings 
         Height          =   165
         Index           =   2
         Left            =   2790
         TabIndex        =   55
         Top             =   195
         Width           =   1170
         _ExtentX        =   1720
         _ExtentY        =   291
         BackgroundColor =   2367774
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "HexFont"
         CaptionOffsetLeft=   5
         CaptionOffsetTop=   1
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   0
         CheckOffsetLeft =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
      End
   End
   Begin VB.TextBox txtDataExchange 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5070
      TabIndex        =   47
      Top             =   5025
      Visible         =   0   'False
      Width           =   480
   End
   Begin SerialConsole.uFrame frmSearch 
      Height          =   645
      Left            =   210
      TabIndex        =   44
      Top             =   1455
      Visible         =   0   'False
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   1138
      BackgroundColor =   2367774
      BorderColor     =   14737632
      ForeColor       =   16777215
      Caption         =   "Search individual Characters"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin SerialConsole.uTextBox txtSearch 
         Height          =   330
         Left            =   75
         TabIndex        =   45
         Top             =   210
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         BorderColor     =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   2367774
         ConsoleColors   =   0   'False
      End
      Begin SerialConsole.uButton cmdSearch 
         Height          =   330
         Left            =   2145
         TabIndex        =   46
         Top             =   210
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         BackgroundColor =   4671472
         BorderColor     =   8421504
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8882165
         FocusColor      =   12632256
         BackgroundColorDisabled=   14737632
         BorderColorDisabled=   8421504
         ForeColorDisabled=   0
         MouseOverBackgroundColorDisabled=   12632256
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   12632256
         FocusVisible    =   0   'False
         Caption         =   "Find"
         Border          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionOffsetTop=   -1
      End
      Begin SerialConsole.uButton cmdSearchClose 
         Height          =   330
         Left            =   3000
         TabIndex        =   48
         Top             =   210
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         BackgroundColor =   3551534
         BorderColor     =   8421504
         ForeColor       =   16777215
         MouseOverBackgroundColor=   12632256
         FocusColor      =   12632256
         BackgroundColorDisabled=   14737632
         BorderColorDisabled=   8421504
         ForeColorDisabled=   0
         MouseOverBackgroundColorDisabled=   12632256
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   12632256
         FocusVisible    =   0   'False
         Caption         =   "X"
         Border          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionOffsetTop=   -1
      End
   End
   Begin VB.Timer tmrShowBuffer 
      Interval        =   1
      Left            =   15075
      Top             =   855
   End
   Begin VB.PictureBox picToolbar 
      Appearance      =   0  'Flat
      BackColor       =   &H0024211E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Index           =   2
      Left            =   0
      ScaleHeight     =   105
      ScaleWidth      =   1695
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1695
   End
   Begin SerialConsole.uFrame frmOutput 
      Height          =   960
      Left            =   75
      TabIndex        =   10
      Top             =   6450
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   1693
      BackgroundColor =   2367774
      BorderColor     =   14737632
      ForeColor       =   16777215
      Caption         =   "Send Information"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin SerialConsole.uOptionBox optInput 
         Height          =   315
         Index           =   1
         Left            =   1980
         TabIndex        =   13
         Top             =   180
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         BackgroundColor =   2367774
         Border          =   0   'False
         Caption         =   "HEX"
         CaptionOffsetTop=   2
         CheckBackgroundColor=   2367774
         CheckBorderColor=   4671472
         CheckBorderThickness=   2
         CheckSelectionColor=   8421631
         CheckSize       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4671472
      End
      Begin SerialConsole.uTextBox txtOutput 
         Height          =   330
         Left            =   90
         TabIndex        =   11
         Top             =   540
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         BorderColor     =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   2367774
         BorderThickness =   3
         ConsoleColors   =   0   'False
      End
      Begin SerialConsole.uButton cmdSend 
         Height          =   330
         Left            =   3585
         TabIndex        =   12
         Top             =   540
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   582
         BackgroundColor =   4671472
         BorderColor     =   8421504
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8882165
         FocusColor      =   12632256
         BackgroundColorDisabled=   8421504
         BorderColorDisabled=   8421504
         ForeColorDisabled=   14737632
         MouseOverBackgroundColorDisabled=   12632256
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   12632256
         FocusVisible    =   0   'False
         Caption         =   "Send"
         Border          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionOffsetTop=   -1
         Enabled         =   0   'False
      End
      Begin SerialConsole.uOptionBox optInput 
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         BackgroundColor =   2367774
         Border          =   0   'False
         Caption         =   "ANSII"
         CaptionOffsetTop=   2
         CheckBackgroundColor=   2367774
         CheckBorderColor=   13604352
         CheckBorderThickness=   2
         CheckSelectionColor=   16776960
         CheckSize       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   13604352
      End
      Begin SerialConsole.uCheckBox chkSend 
         Height          =   285
         Index           =   0
         Left            =   7410
         TabIndex        =   15
         ToolTipText     =   "Clear On Send"
         Top             =   165
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         BackgroundColor =   2367774
         Border          =   0   'False
         Caption         =   "COS"
         CaptionOffsetLeft=   5
         CaptionOffsetTop=   2
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   1
         CheckOffsetLeft =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
      End
      Begin SerialConsole.uOptionBox optInput 
         Height          =   315
         Index           =   2
         Left            =   1170
         TabIndex        =   16
         Top             =   180
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         BackgroundColor =   2367774
         Border          =   0   'False
         Caption         =   "BIN"
         CaptionOffsetTop=   2
         CheckBackgroundColor=   2367774
         CheckBorderColor=   1746682
         CheckBorderThickness=   2
         CheckSelectionColor=   8438015
         CheckSize       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   1746682
      End
      Begin SerialConsole.uOptionBox optInput 
         Height          =   315
         Index           =   3
         Left            =   2835
         TabIndex        =   17
         Top             =   165
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         BackgroundColor =   2367774
         Border          =   0   'False
         Caption         =   "DEC"
         CaptionOffsetTop=   2
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8500547
         CheckBorderThickness=   2
         CheckSelectionColor=   8454016
         CheckSize       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8500547
      End
      Begin SerialConsole.uOptionBox optInput 
         Height          =   315
         Index           =   4
         Left            =   3645
         TabIndex        =   18
         Top             =   165
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         BackgroundColor =   2367774
         Border          =   0   'False
         Caption         =   "OCTAL"
         CaptionOffsetTop=   2
         CheckBackgroundColor=   2367774
         CheckBorderColor=   14322034
         CheckBorderThickness=   2
         CheckSelectionColor=   16761024
         CheckSize       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14322034
      End
      Begin SerialConsole.uDropDown drpOnSend 
         Height          =   270
         Left            =   10020
         TabIndex        =   63
         Top             =   195
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   476
         BackgroundColor =   14322034
         BorderColor     =   14322034
         ForeColor       =   16777215
         SelectionBackgroundColor=   13592135
         SelectionBorderColor=   14322034
         BackgroundColorDisabled=   14737632
         BorderColorDisabled=   8421504
         SelectionBackgroundColorDisabled=   14737632
         SelectionBorderColorDisabled=   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "drpCommports met een erg lange zin er achter aan"
         Border          =   0   'False
         ItemHeight      =   19
         VisibleItems    =   8
         ScrollBarWidth  =   19
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OnSend:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Index           =   0
         Left            =   9030
         TabIndex        =   64
         Top             =   210
         Width           =   945
      End
   End
   Begin SerialConsole.uTextBox txtReceived 
      Height          =   1965
      Left            =   225
      TabIndex        =   0
      Top             =   2190
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   3466
      BackgroundColor =   3551534
      BorderColor     =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      LineNumbers     =   -1  'True
      LineNumberForeColor=   8421504
      LineNumberBackground=   2367774
      RowLines        =   -1  'True
      RowLineColor    =   4210752
      RowNumberOnEveryLine=   -1  'True
      WordWrap        =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1
   End
   Begin SerialConsole.uTextBox txtStatus 
      Height          =   420
      Left            =   6600
      TabIndex        =   5
      Top             =   7590
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   741
      BorderColor     =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   2367774
      BorderThickness =   3
      ConsoleColors   =   0   'False
   End
   Begin VB.PictureBox picToolbar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0024211E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   1
      Left            =   225
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   552
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7395
      Width           =   8280
      Begin SerialConsole.uFrame frmInOut 
         Height          =   750
         Left            =   3315
         TabIndex        =   66
         Top             =   -30
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1323
         BackgroundColor =   2367774
         BorderColor     =   14737632
         ForeColor       =   16777215
         Caption         =   "Bytes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H0024211E&
            Caption         =   "IN:  0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   210
            Index           =   2
            Left            =   90
            TabIndex        =   68
            Top             =   210
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H0024211E&
            Caption         =   "OUT: 0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   210
            Index           =   1
            Left            =   90
            TabIndex        =   67
            Top             =   465
            Width           =   720
         End
      End
      Begin SerialConsole.uFrame frmComStats 
         Height          =   750
         Left            =   4545
         TabIndex        =   56
         Top             =   -30
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1323
         BackgroundColor =   2367774
         BorderColor     =   14737632
         ForeColor       =   16777215
         Caption         =   "Com Stats"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label lblComStats 
            Alignment       =   2  'Center
            BackColor       =   &H0036312E&
            Caption         =   "RING"
            ForeColor       =   &H004747F0&
            Height          =   195
            Index           =   3
            Left            =   765
            TabIndex        =   60
            Top             =   465
            Width           =   600
         End
         Begin VB.Label lblComStats 
            Alignment       =   2  'Center
            BackColor       =   &H0036312E&
            Caption         =   "CD"
            ForeColor       =   &H004747F0&
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   59
            Top             =   465
            Width           =   600
         End
         Begin VB.Label lblComStats 
            Alignment       =   2  'Center
            BackColor       =   &H0036312E&
            Caption         =   "DSR"
            ForeColor       =   &H004747F0&
            Height          =   195
            Index           =   1
            Left            =   765
            TabIndex        =   58
            Top             =   195
            Width           =   600
         End
         Begin VB.Label lblComStats 
            Alignment       =   2  'Center
            BackColor       =   &H0036312E&
            Caption         =   "CTS"
            ForeColor       =   &H004747F0&
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   57
            Top             =   195
            Width           =   600
         End
      End
      Begin SerialConsole.uGraph graphDataInOut 
         Height          =   795
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1402
      End
   End
   Begin VB.PictureBox picFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   45
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   45
      Width           =   150
   End
   Begin VB.Timer tmrGetConnected 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   14625
      Top             =   855
   End
   Begin SerialConsole.uCheckBox chkComOptions 
      Height          =   450
      Index           =   0
      Left            =   10395
      TabIndex        =   4
      Top             =   180
      Width           =   870
      _ExtentX        =   1482
      _ExtentY        =   794
      BackgroundColor =   2367774
      Border          =   0   'False
      BorderColor     =   2367774
      Caption         =   "DTR"
      CaptionOffsetLeft=   5
      CaptionOffsetTop=   2
      CheckBackgroundColor=   2367774
      CheckBorderColor=   8421504
      CheckBorderThickness=   2
      CheckSelectionColor=   4210752
      CheckOffsetLeft =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12632256
   End
   Begin SerialConsole.uDropDown drpCommports 
      Height          =   450
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   794
      BackgroundColor =   14322034
      BorderColor     =   14322034
      ForeColor       =   16777215
      SelectionBackgroundColor=   13592135
      SelectionBorderColor=   14322034
      BackgroundColorDisabled=   8421504
      BorderColorDisabled=   8421504
      ForeColorDisabled=   14737632
      SelectionBackgroundColorDisabled=   14737632
      SelectionBorderColorDisabled=   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "drpCommports met een erg lange zin er achter aan"
      Border          =   0   'False
      ScrollBarWidth  =   30
   End
   Begin MSCommLib.MSComm comm 
      Left            =   4365
      Top             =   4965
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      InBufferSize    =   10
      InputLen        =   100
      OutBufferSize   =   1
      ParityReplace   =   0
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin SerialConsole.uDropDown drpBaud 
      Height          =   450
      Left            =   6615
      TabIndex        =   2
      Top             =   180
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   794
      BackgroundColor =   14322034
      BorderColor     =   14322034
      ForeColor       =   16777215
      SelectionBackgroundColor=   13592135
      SelectionBorderColor=   14322034
      BackgroundColorDisabled=   14737632
      BorderColorDisabled=   8421504
      SelectionBackgroundColorDisabled=   14737632
      SelectionBorderColorDisabled=   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "drpCommports met een erg lange zin er achter aan"
      ItemHeight      =   20
      ScrollBarWidth  =   30
   End
   Begin SerialConsole.uCheckBox chkComOptions 
      Height          =   450
      Index           =   1
      Left            =   11550
      TabIndex        =   6
      Top             =   180
      Width           =   870
      _ExtentX        =   1482
      _ExtentY        =   794
      BackgroundColor =   2367774
      Border          =   0   'False
      BorderColor     =   2367774
      Caption         =   "RTS"
      CaptionOffsetLeft=   5
      CaptionOffsetTop=   2
      CheckBackgroundColor=   2367774
      CheckBorderColor=   8421504
      CheckBorderThickness=   2
      CheckSelectionColor=   4210752
      CheckOffsetLeft =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12632256
   End
   Begin VB.PictureBox picToolbar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0024211E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   0
      Left            =   0
      ScaleHeight     =   810
      ScaleWidth      =   14730
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   14730
   End
   Begin SerialConsole.uFrame frmWindow 
      Height          =   5355
      Index           =   0
      Left            =   7380
      TabIndex        =   21
      Top             =   1275
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   9446
      BackgroundColor =   2367774
      BorderColor     =   14737632
      ForeColor       =   16777215
      Caption         =   "Zebro Controls"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin SerialConsole.uCheckBox chkRefreshZebro 
         Height          =   285
         Left            =   90
         TabIndex        =   39
         Top             =   4725
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   503
         BackgroundColor =   2367774
         Border          =   0   'False
         Caption         =   "Refresh connected Zebros"
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
      End
      Begin VB.PictureBox picConnected 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   90
         ScaleHeight     =   255
         ScaleWidth      =   390
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   195
         Width           =   420
      End
      Begin SerialConsole.uFrame frmColors 
         Height          =   615
         Left            =   90
         TabIndex        =   22
         Top             =   4050
         Visible         =   0   'False
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   1085
         BackgroundColor =   2367774
         BorderColor     =   14737632
         ForeColor       =   16777215
         Caption         =   "Led Color"
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
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   225
            Width           =   450
         End
      End
      Begin SerialConsole.uButton cmdControls 
         Height          =   1410
         Index           =   0
         Left            =   1575
         TabIndex        =   24
         Top             =   1095
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
      Begin SerialConsole.uButton cmdZebro 
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   26
         Top             =   555
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
      Begin SerialConsole.uButton cmdControls 
         Height          =   1410
         Index           =   1
         Left            =   3060
         TabIndex        =   27
         Top             =   1095
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
      Begin SerialConsole.uButton cmdControls 
         Height          =   1410
         Index           =   2
         Left            =   90
         TabIndex        =   28
         Top             =   1095
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
      Begin SerialConsole.uButton cmdControls 
         Height          =   1410
         Index           =   3
         Left            =   1575
         TabIndex        =   29
         Top             =   2580
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
      Begin SerialConsole.uButton cmdControls 
         Height          =   420
         Index           =   4
         Left            =   90
         TabIndex        =   30
         Top             =   2580
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
      Begin SerialConsole.uButton cmdControls 
         Height          =   420
         Index           =   5
         Left            =   90
         TabIndex        =   31
         Top             =   3075
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
      Begin SerialConsole.uButton cmdControls 
         Height          =   420
         Index           =   6
         Left            =   90
         TabIndex        =   32
         Top             =   3570
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
      Begin SerialConsole.uButton cmdControls 
         Height          =   420
         Index           =   7
         Left            =   3060
         TabIndex        =   33
         Top             =   2580
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
      Begin SerialConsole.uButton cmdControls 
         Height          =   420
         Index           =   8
         Left            =   3060
         TabIndex        =   34
         Top             =   3075
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
      Begin SerialConsole.uButton cmdControls 
         Height          =   420
         Index           =   9
         Left            =   3060
         TabIndex        =   35
         Top             =   3570
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
      Begin SerialConsole.uButton cmdControls 
         Height          =   1410
         Index           =   10
         Left            =   4545
         TabIndex        =   36
         Top             =   1095
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
      Begin SerialConsole.uButton cmdControls 
         Height          =   1410
         Index           =   11
         Left            =   4545
         TabIndex        =   37
         Top             =   2580
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
   Begin SerialConsole.uFrame uFrame1 
      Height          =   315
      Left            =   5730
      TabIndex        =   50
      Top             =   5130
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      BackgroundColor =   2367774
      BorderColor     =   14737632
      ForeColor       =   16777215
      Caption         =   "uFrame"
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
   Begin SerialConsole.uDropDown drpWindowType 
      Height          =   330
      Left            =   7290
      TabIndex        =   69
      Top             =   825
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   582
      BackgroundColor =   14322034
      BorderColor     =   14322034
      ForeColor       =   16777215
      SelectionBackgroundColor=   13592135
      SelectionBorderColor=   14322034
      BackgroundColorDisabled=   14737632
      BorderColorDisabled=   8421504
      SelectionBackgroundColorDisabled=   14737632
      SelectionBorderColorDisabled=   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "List"
      Border          =   0   'False
      ScrollBarWidth  =   30
   End
   Begin SerialConsole.uFrame frmWindow 
      Height          =   2655
      Index           =   3
      Left            =   13590
      TabIndex        =   75
      Top             =   3015
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   4683
      BackgroundColor =   2367774
      BorderColor     =   14737632
      ForeColor       =   16777215
      Caption         =   "Logs"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin SerialConsole.uCheckBox chkLogsEnable 
         Height          =   165
         Left            =   75
         TabIndex        =   76
         ToolTipText     =   "Show the received data in hex (Hold: H)"
         Top             =   210
         Width           =   1905
         _ExtentX        =   1191
         _ExtentY        =   291
         BackgroundColor =   2367774
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "Enable Logs"
         CaptionOffsetLeft=   5
         CaptionOffsetTop=   1
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   0
         CheckOffsetLeft =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
         AutoSize        =   0   'False
      End
      Begin SerialConsole.uButton cmdOpenLog 
         Height          =   330
         Left            =   135
         TabIndex        =   77
         Top             =   1950
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         BackgroundColor =   4671472
         BorderColor     =   8421504
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8882165
         FocusColor      =   12632256
         BackgroundColorDisabled=   14737632
         BorderColorDisabled=   8421504
         ForeColorDisabled=   0
         MouseOverBackgroundColorDisabled=   12632256
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   12632256
         FocusVisible    =   0   'False
         Caption         =   "Open Folder"
         Border          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionOffsetTop=   -1
      End
      Begin SerialConsole.uFrame frmLogsOnReconnect 
         Height          =   1290
         Left            =   135
         TabIndex        =   78
         Top             =   525
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   2275
         BackgroundColor =   2367774
         BorderColor     =   14737632
         ForeColor       =   16777215
         Caption         =   "On connect ..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin SerialConsole.uOptionBox optLogsReconnect 
            Height          =   300
            Index           =   0
            Left            =   90
            TabIndex        =   79
            Top             =   195
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            BackgroundColor =   2367774
            Border          =   0   'False
            Caption         =   "Append file"
            CheckBackgroundColor=   2367774
            CheckBorderColor=   8421504
            CheckBorderThickness=   2
            CheckSelectionColor=   8882165
            CheckSize       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   12632256
            Value           =   1
         End
         Begin SerialConsole.uOptionBox optLogsReconnect 
            Height          =   300
            Index           =   1
            Left            =   90
            TabIndex        =   80
            Top             =   555
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            BackgroundColor =   2367774
            Border          =   0   'False
            Caption         =   "Overwrite file"
            CheckBackgroundColor=   2367774
            CheckBorderColor=   8421504
            CheckBorderThickness=   2
            CheckSelectionColor=   8882165
            CheckSize       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   12632256
         End
         Begin SerialConsole.uOptionBox optLogsReconnect 
            Height          =   300
            Index           =   2
            Left            =   90
            TabIndex        =   81
            Top             =   915
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            BackgroundColor =   2367774
            Border          =   0   'False
            Caption         =   "Create new file"
            CheckBackgroundColor=   2367774
            CheckBorderColor=   8421504
            CheckBorderThickness=   2
            CheckSelectionColor=   8882165
            CheckSize       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   12632256
         End
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H0024211E&
         Caption         =   "Current Logfile Name: "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   88
         Top             =   2355
         Width           =   1620
      End
   End
   Begin SerialConsole.uFrame frmWindow 
      Height          =   1980
      Index           =   4
      Left            =   13665
      TabIndex        =   89
      Top             =   5940
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3493
      BackgroundColor =   2367774
      BorderColor     =   14737632
      ForeColor       =   16777215
      Caption         =   "History"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin SerialConsole.uListBox lstHistory 
         Height          =   1200
         Left            =   90
         TabIndex        =   90
         Top             =   540
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2117
         BackgroundColor =   3551534
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
         ForeColor       =   16777215
         Text            =   "uFrame"
         SelectionBackgroundColor=   3551534
         SelectionBorderColor=   16777215
         SelectionForeColor=   14322034
         ItemHeight      =   5
         VisibleItems    =   15
      End
      Begin SerialConsole.uCheckBox chkSendOnDoubleClick 
         Height          =   165
         Left            =   90
         TabIndex        =   94
         ToolTipText     =   "Show the received data in hex (Hold: H)"
         Top             =   240
         Width           =   2295
         _ExtentX        =   1191
         _ExtentY        =   291
         BackgroundColor =   2367774
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "Send on double click"
         CaptionOffsetLeft=   5
         CaptionOffsetTop=   1
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   0
         CheckOffsetLeft =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
         AutoSize        =   0   'False
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H0024211E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   7035
      MousePointer    =   9  'Size W E
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   825
      Width           =   180
   End
   Begin SerialConsole.uFrame frmWindow 
      Height          =   3990
      Index           =   5
      Left            =   9975
      TabIndex        =   96
      Top             =   1305
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7038
      BackgroundColor =   2367774
      BorderColor     =   14737632
      ForeColor       =   16777215
      Caption         =   "Label List"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin SerialConsole.uButton cmdClearLabels 
         Height          =   330
         Left            =   90
         TabIndex        =   97
         Top             =   3585
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   582
         BackgroundColor =   4671472
         BorderColor     =   8421504
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8882165
         FocusColor      =   12632256
         BackgroundColorDisabled=   14737632
         BorderColorDisabled=   8421504
         ForeColorDisabled=   0
         MouseOverBackgroundColorDisabled=   12632256
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   12632256
         FocusVisible    =   0   'False
         Caption         =   "Clear Labels"
         Border          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionOffsetTop=   -1
      End
      Begin SerialConsole.uCheckBox chkEnableLabelList 
         Height          =   165
         Left            =   90
         TabIndex        =   99
         ToolTipText     =   "Show the received data in hex (Hold: H)"
         Top             =   240
         Width           =   1950
         _ExtentX        =   1191
         _ExtentY        =   291
         BackgroundColor =   2367774
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "Enable Label List"
         CaptionOffsetLeft=   5
         CaptionOffsetTop=   1
         CheckBackgroundColor=   2367774
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   0
         CheckOffsetLeft =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
         AutoSize        =   0   'False
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H0024211E&
         Caption         =   "Label1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   98
         Top             =   555
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.Label lblCursorStats 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2145
      TabIndex        =   95
      Top             =   5040
      Width           =   570
   End
   Begin VB.Label LBLSplit 
      AutoSize        =   -1  'True
      BackColor       =   &H0024211E&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Index           =   1
      Left            =   6555
      TabIndex        =   73
      Top             =   4485
      Width           =   180
   End
   Begin VB.Label LBLSplit 
      AutoSize        =   -1  'True
      BackColor       =   &H0024211E&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Index           =   0
      Left            =   6750
      TabIndex        =   72
      Top             =   3975
      Width           =   180
   End
   Begin VB.Label lblCursorStats 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   3750
      TabIndex        =   43
      Top             =   5040
      Width           =   540
   End
   Begin VB.Label lblCursorStats 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1485
      TabIndex        =   42
      Top             =   5040
      Width           =   600
   End
   Begin VB.Label lblCursorStats 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   765
      TabIndex        =   41
      Top             =   5040
      Width           =   645
   End
   Begin VB.Label lblCursorStats 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   40
      Top             =   5040
      Width           =   600
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'TODO: verschillende programma instances aan elkaar kunnen linken dat de serial devices tegen elkaar praten (man in the middle)
'   - 2 output windows dan 2 vensters
'- Tegen zichzelf kunnen praten (loopback).
'- Thema's maken

Dim dragSplitStartX As Long
Dim dragSplitPercentage As Double
Dim dragSplit As Boolean

Dim serialDevices As CommPortList
Dim Timer As PerformanceTimer
Dim inputFilter As InputHandler


Private receiveBufferForShow As String * 5000
Private receiveBufferForShowLength As Long

Private receiveBufferArduino As String * 5000
Private receiveBufferArduinoLength As Long

Dim picoSendCommand(0 To 10) As Boolean
Dim picoConnected(0 To 10) As Boolean

Dim lastMessageBytes() As Byte

Dim receiveBuffer As String
Dim ledCommand As Long

Dim errorMessages() As String
Const errorMessagesConst = ",M_ERROR,M_ERROR_NOT_CONNECTED,M_ERROR_BUFFER_OVERFLOW,M_ERROR_BUFFER_EMPTY,M_ERROR_UNKNOWN_COMMAND"

Dim onSendCharacters() As String

Dim bitrateInbound As Long
Dim bitrateOutbound As Long
Dim bitsReceived As Long
Dim bitsSend As Long

Dim searchFor() As Byte

Dim ConsoleColors As Variant

Dim arduinoListHeaders() As String
Dim arduinoLabels As Collection

Dim logFileHandle As Long

Private WithEvents tmrCheckBitRate As SelfTimer
Attribute tmrCheckBitRate.VB_VarHelpID = -1

Private Sub chkComOptions_Changed(Index As Integer, u_NewState As uCheckboxConstants)
    On Error GoTo disconnectError:
    Dim newState As Boolean
    
    newState = (u_NewState = u_Checked)
    
    Select Case Index
    
        Case 0
            comm.DTREnable = newState
            If newState And chkComOptions(2).Value = u_Checked Then
                txtReceived.Clear
            End If
            
        Case 1
            comm.RTSEnable = newState
    End Select

    
disconnectError:
    If Err.number <> 0 Then
        Err.Clear
        If comm.PortOpen Then
            cmdConnect_Click 0, 0, 0
             
        End If
        
    End If
    
    SaveSetting "SerialConsole", "checkboxes", "chkComOptions(" & Index & ").Value", u_NewState

    setCheckColors chkComOptions(Index), newState

End Sub

Sub setCheckColors(chk As uCheckBox, newState As Boolean)
    With chk
        If newState = False Then
            .CheckBorderColor = &H808080
            .CheckBackgroundColor = .BackgroundColor
            .CheckSelectionColor = vbWhite
        Else
            .CheckBorderColor = &HDA8972
            .CheckBackgroundColor = &HDA8972
            .CheckSelectionColor = vbWhite
        End If
        
        .Redraw
    End With
End Sub

Private Sub chkEnableGraph_Changed(u_NewState As uCheckboxConstants)
    
    setCheckColors chkEnableGraph, u_NewState = u_Checked
End Sub

Private Sub chkEnableLabelList_Changed(u_NewState As uCheckboxConstants)
    SaveSetting "SerialConsole", "label", "chkEnableLabelList.Value", u_NewState
    
    setCheckColors chkEnableLabelList, u_NewState = u_Checked
End Sub

Private Sub chkLogsEnable_Changed(u_NewState As uCheckboxConstants)
    SaveSetting "SerialConsole", "logs", "chkLogsEnable.Value", u_NewState
    
    If u_NewState = u_Checked And comm.PortOpen = True And logFileHandle = -1 Then
        checkForAndOpenLogFile
    ElseIf u_NewState = u_unChecked And logFileHandle <> -1 Then
        Close logFileHandle
        logFileHandle = -1
    End If
    
    setCheckColors chkLogsEnable, u_NewState = u_Checked
End Sub

Private Sub chkRefreshZebro_Changed(u_NewState As uCheckboxConstants)
    setCheckColors chkRefreshZebro, u_NewState = u_Checked
    
    tmrGetConnected.Enabled = (u_NewState = u_Checked)
End Sub

Private Sub chkSend_Changed(Index As Integer, u_NewState As uCheckboxConstants)
    setCheckColors chkSend(Index), u_NewState = u_Checked
    
End Sub

Private Sub chkSendOnDoubleClick_Changed(u_NewState As uCheckboxConstants)

    SaveSetting "SerialConsole", "history", "chkSendOnDoubleClick.Value", u_NewState

    setCheckColors chkSendOnDoubleClick, u_NewState = u_Checked
End Sub

Private Sub chkTxtSettings_Changed(Index As Integer, u_NewState As uCheckboxConstants)
    Dim newState As Boolean
    
    newState = (u_NewState = u_Checked)
    
    Select Case Index
        Case 1
            txtReceived.ConsoleColors = newState
            
        Case 2
            If newState Then
                txtReceived.Font.Name = "CompendiumArcana Hexadecimal"
                txtReceived.Font.Bold = False
                txtReceived.Redraw
            Else
                txtReceived.Font.Name = "CompendiumArcana Ctrl Char Hex"
                txtReceived.Font.Bold = False
                txtReceived.Redraw
            End If

    End Select
    
    SaveSetting "SerialConsole", "checkboxes", "chkTxtSettings(" & Index & ").Value", u_NewState
    
    setCheckColors chkTxtSettings(Index), newState
End Sub


Private Sub cmdClearLabels_Click(Button As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim s As Variant
    
    For Each s In arduinoLabels
        If s(0) <> 0 Then
            Unload lblLabel(s(0))
        Else
            lblLabel(s(0)).Visible = False
        End If
    Next s
    
    Set arduinoLabels = New Collection
End Sub

Private Sub cmdConnect_MouseEnter()
    loadReconnect.BackgroundColor = cmdConnect.MouseOverBackgroundColor
End Sub

Private Sub cmdConnect_MouseLeave()
    loadReconnect.BackgroundColor = cmdConnect.BackgroundColor
End Sub


Private Sub cmdConnect_Click(Button As Integer, X As Single, Y As Single)
    On Error GoTo notWorking
    
    If comm.PortOpen Then
        comm.PortOpen = False
        cmdConnect.Caption = "Connect"
        cmdConnect.BackgroundColor = &H4747F0
        cmdConnect.MouseOverBackgroundColor = &H8787F5
        If cmdConnect.isMouseOverControl Then
            loadReconnect.BackgroundColor = &H8787F5
        Else
            loadReconnect.BackgroundColor = &H4747F0
        End If
        
        drpCommports.Enabled = True
        cmdSend.Enabled = False
        
        If logFileHandle <> -1 Then
            Close logFileHandle
            logFileHandle = -1
        End If
        
    Else
        comm.DTREnable = False
        receiveBufferForShowLength = 0
        chkRefreshZebro.Value = u_unChecked
        
        comm.PortOpen = True
        comm.InBufferCount = 0
        comm.OutBufferCount = 0
        comm.DTREnable = (chkComOptions(0).Value = u_Checked)
        If chkComOptions(2).Value = u_Checked Then txtReceived.Clear
        
        cmdSend.Enabled = True
        drpCommports.Enabled = False
        If chkLogsEnable.Value = u_Checked Then checkForAndOpenLogFile
        tmrCheckForReconnect.Enabled = False
        loadReconnect.Loading = False
        setStatus "Connected!"
        
        cmdConnect.Caption = "Disconnect"
        cmdConnect.BackgroundColor = &H81B543
        cmdConnect.MouseOverBackgroundColor = &HA4CB74
        If cmdConnect.isMouseOverControl Then
            loadReconnect.BackgroundColor = &HA4CB74
        Else
            loadReconnect.BackgroundColor = &H81B543
        End If
        
        tmrShowBuffer.Enabled = True
    End If
    
Exit Sub
notWorking:
    
    setStatus Err.Description, True, Err.number
    
End Sub

Sub setStatus(Msg As String, Optional isError As Boolean = False, Optional errorNumber As Long = 0)
    txtStatus.RedrawPause
    
    If isError Then
        txtStatus.Text = "[ ERROR " & errorNumber & " ] " & Msg
        txtStatus.ForeColor = &H4747F0
        txtStatus.BorderColor = &H4747F0
    Else
        txtStatus.Text = Msg
        txtStatus.ForeColor = &H81B543
        txtStatus.BorderColor = &H81B543
    End If
    
    txtStatus.RedrawResume
End Sub

Sub checkForAndOpenLogFile()
    initLogs
    
    Dim deviceName As String
    Dim fileName As String

    If serialDevices.Count = 0 Then Exit Sub
    
    deviceName = serialDevices.friendlyName(drpCommports.ListIndex)
    
    fileName = App.Path & "\logs\" & deviceName
    
    If Dir(fileName, vbDirectory) = "" Then
        MkDir fileName
    End If
    

    If optLogsReconnect(0).Value = u_Selected Or optLogsReconnect(1).Value = u_Selected Then 'append or overwrite
        Dim sFile As String
        Dim sFile2 As String
        sFile2 = Dir(fileName & "\*.log", vbNormal)
        sFile = sFile2
        
        Do While sFile2 <> ""
            sFile2 = Dir
            If sFile2 <> "" Then sFile = sFile2
        Loop
        
        logFileHandle = FreeFile
        
        If sFile = "" Then
            getRightFileName fileName
            Open fileName For Binary Access Write As logFileHandle
        Else
            fileName = fileName & "\" & sFile
            
            If optLogsReconnect(1).Value = u_Selected Then
                Kill fileName
            End If
            
            Open fileName For Binary Access Write As logFileHandle
            Seek logFileHandle, LOF(logFileHandle) + 1
        End If
    
    ElseIf optLogsReconnect(2).Value = u_Selected Then 'create new file
        getRightFileName fileName
        
        logFileHandle = FreeFile
    
        Open fileName For Binary Access Write As logFileHandle
    End If
    
    Dim lastName() As String
    If fileName <> "" Then
        lastName = Split(fileName, "\")
        lblInfo(3).Caption = "Current Logfile Name: " & lastName(UBound(lastName))
    Else
        lblInfo(3).Caption = "Current Logfile Name: ..."
    End If
    
        
    'If number = -1 Then
    '    MsgBox "The logs folder is full (total of 1000 files) please move them or clear the folder"
    'End If
    
End Sub


Sub getRightFileName(ByRef currentFilename As String)
    Dim i As Long
    Dim number As Long
    Dim today As String
    
    today = Format(Now, "yyyyMMdd_HH.mm.ss")
    currentFilename = currentFilename & "\" & today
    If Dir(currentFilename & ".log") <> "" Then
        currentFilename = currentFilename & "_"
        Do
            If Dir(currentFilename & i & ".log") = "" Then
                number = i
                Exit Do
            End If
            i = i + 1
        Loop While True
        
        currentFilename = currentFilename & number & ".log"
    Else
        currentFilename = currentFilename & ".log"
    End If
    
    

    
End Sub

Private Sub cmdControls_Click(Index As Integer, Button As Integer, X As Single, Y As Single)
    
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

Private Sub cmdOpenLog_Click(Button As Integer, X As Single, Y As Single)

Dim deviceName As String
    Dim fileName As String

    fileName = App.Path & "\logs\"
    
    If serialDevices.Count > 0 Then
        fileName = fileName & serialDevices.friendlyName(drpCommports.ListIndex)
    End If
    
    ShellExecute Me.hWnd, "OPEN", "explorer.exe", fileName, "", vbNormalFocus
    
End Sub

Private Sub cmdSearch_Click(Button As Integer, X As Single, Y As Single)
    Dim tmpOutput As Variant
    Dim bytes() As Byte
    
    searchFor = parseInputToBytes(txtSearch)
    
    fillReceivedTextColors 0
    
    txtReceived.RedrawResume
End Sub


Function parseInputToBytes(uTxt As uTextBox) As Byte()
    Dim splitStr() As String
    Dim str As String
    Dim i As Long, j As Long
    
    str = uTxt.Text
    splitStr = Split(str, " ")
    
    Dim forceFunction As Long
    forceFunction = -1
    
    For i = optInput.LBound To optInput.UBound
        If optInput(i).Value = u_Selected Then
            forceFunction = i
            Exit For
        End If
    Next i
    
    Dim parseBytes() As Byte
    Dim totalBytes() As Byte
    Dim totalBytesLength As Long
    
    If forceFunction = 0 Then
        parseBytes = inputFilter.parseString(str, forceFunction)
        
        If UBound(parseBytes) > -1 Then
                
            ReDim Preserve totalBytes(0 To totalBytesLength + UBound(parseBytes))
            
            For j = 0 To UBound(parseBytes)
                totalBytes(totalBytesLength + j) = parseBytes(j)
            Next j
            
            totalBytesLength = totalBytesLength + UBound(parseBytes) + 1
        End If
        
    Else
        For i = 0 To UBound(splitStr)
            parseBytes = inputFilter.parseString(splitStr(i), forceFunction)
            If UBound(parseBytes) > -1 Then
                
                ReDim Preserve totalBytes(0 To totalBytesLength + UBound(parseBytes))
                
                For j = 0 To UBound(parseBytes)
                    totalBytes(totalBytesLength + j) = parseBytes(j)
                Next j
                
                totalBytesLength = totalBytesLength + UBound(parseBytes) + 1
            Else
                Erase parseInputToBytes
                uTxt.BorderColor = &H4747F0
                uTxt.BackgroundColor = &H8080FF
                Exit Function
            End If
            
        Next i
    End If
    
    If drpOnSend.ListIndex > 0 Then
        j = Len(onSendCharacters(drpOnSend.ListIndex))
        ReDim Preserve totalBytes(0 To totalBytesLength - 1 + j)
    
        For i = 1 To j
            totalBytes(UBound(totalBytes) - j + i) = Asc(Mid$(onSendCharacters(drpOnSend.ListIndex), i, 1))
        Next i
        totalBytesLength = totalBytesLength + j
    End If
    
    
    parseInputToBytes = totalBytes
    
    Dim tmpStr As String
    
    For i = 0 To UBound(totalBytes)
        tmpStr = tmpStr & totalBytes(i) & " "
    Next i
    
    'MsgBox tmpStr
End Function

Private Sub cmdSearchClose_Click(Button As Integer, X As Single, Y As Single)
    frmSearch.Visible = False
    Erase searchFor
    txtReceived.ClearMarkup
    txtReceived.Redraw
    Form_Resize
    txtReceived.SetFocus
    'fillReceivedTextColors 0
End Sub

Private Sub cmdSend_Click(Button As Integer, X As Single, Y As Single)
    Dim tmpOutput As Variant
    Dim bytes() As Byte
    Dim tmpCombine As String
    Dim i As Long
    
    bytes = parseInputToBytes(txtOutput)
    
    If comm.PortOpen = True Then
        If (Not (Not bytes)) <> 0 Then
            tmpOutput = bytes
            commOut tmpOutput
        End If
        
    End If
    
    'log to history list
    If (Not (Not bytes)) <> 0 Then
        'For i = 0 To UBound(bytes)
        '    tmpCombine = tmpCombine & Chr(bytes(i))
        'Next i
        For i = 0 To lstHistory.ListCount - 1
            If lstHistory.List(i) = txtOutput.Text Then
                lstHistory.RemoveItem i
                Exit For
            End If
        Next i
        lstHistory.AddItem txtOutput.Text, getOutputOptionsAsLong, 0, -1, -1
    End If
    
    If chkSend(0).Value = u_Checked Then
        txtOutput.Clear
    End If
    
    

    
    
    'txtOutput.SetFocus
End Sub

Function getOutputOptionsAsLong() As Long
    Dim outputVal As Long
    Dim i As Long
    
    For i = 0 To optInput.UBound
        If optInput(i).Value = u_Selected Then
            outputVal = (outputVal Or (2 ^ i))
        End If
    Next i
    
    outputVal = outputVal + (2 ^ (drpOnSend.ListIndex + optInput.UBound + 1))
    
    getOutputOptionsAsLong = outputVal
    
End Function

Sub setOutputOptionsWithLong(inputVal As Long)
    Dim i As Long
    
    For i = 0 To optInput.UBound
        If (inputVal And (2 ^ i)) > 0 Then
            optInput(i).Value = u_Selected
        Else
            optInput(i).Value = u_UnSelected
        End If
    Next i
    
    For i = 0 To drpOnSend.ListCount
        If (inputVal And (2 ^ (i + optInput.UBound + 1))) > 0 Then
            drpOnSend.ListIndex = i
            Exit For
        End If
        
    Next i
End Sub

Private Sub cmdZebro_Click(Index As Integer, Button As Integer, X As Single, Y As Single)
        
    picoSendCommand(Index) = Not picoSendCommand(Index)
    
    
    Dim i As Long
    
    For i = 0 To UBound(picoSendCommand)
        cmdZebro(i).BackgroundColor = IIf(picoSendCommand(i), &HFFC0C0, &HE0E0E0)
        cmdZebro(i).MouseOverBackgroundColor = IIf(picoSendCommand(i), &HFF8080, &HC0C0C0)
    Next i
    
End Sub

Private Sub comm_OnComm()
    On Error GoTo showError:
    Static i As Long
    Dim RL As Long 'received length
    
    i = i + 1
    
    Select Case comm.CommEvent
    
        Case comEvReceive   'comEvReceive event occured
            Dim tmpReceived As String
            
            'bitrateInbound = bitrateInbound + comm.InBufferCount
            
            'Timer.StartTimer
            
            tmpReceived = comm.Input
            RL = Len(tmpReceived)
            
            If receiveBufferForShowLength + RL > 4000 Then
                tmrShowBuffer_Timer
            End If
            
            Mid$(receiveBufferForShow, receiveBufferForShowLength + 1, RL) = tmpReceived
            receiveBufferForShowLength = receiveBufferForShowLength + RL
            
            Mid$(receiveBufferArduino, receiveBufferArduinoLength + 1, RL) = tmpReceived
            receiveBufferArduinoLength = receiveBufferArduinoLength + RL
            
            'Timer.StopTimer
            'Debug.Print Timer.TimeElapsed(pvMilliSecond)
            
            
            'Debug.Print RL
            
            '########################################################################################################################################################################
            'Replace this concatenation of bullshit with a more memory friendly solution, like a buffered string of 5000 chars that will stay 5000 chars and will never be moved
            'this way there are no more intense memory usages
            '########################################################################################################################################################################
            
            'Clipboard.Clear
            'Clipboard.SetText receiveBuffer
            
            bitrateInbound = bitrateInbound + RL
            bitsReceived = bitsReceived + RL
            
            If chkRefreshZebro.Value = u_Checked Then receiveBuffer = receiveBuffer & tmpReceived
            'receiveBufferForShow = receiveBufferForShow & tmpReceived
        
        Case comEvCTS
            lblComStats(0).ForeColor = IIf(comm.CTSHolding, vbBlack, &H4747F0)
            lblComStats(0).BackColor = IIf(comm.CTSHolding, &H81B543, &H36312E)
            
        Case comEvDSR
            lblComStats(1).ForeColor = IIf(comm.DSRHolding, vbBlack, &H4747F0)
            lblComStats(1).BackColor = IIf(comm.DSRHolding, &H81B543, &H36312E)
            
        Case comEvCD
            lblComStats(2).ForeColor = IIf(comm.CDHolding, vbBlack, &H4747F0)
            lblComStats(2).BackColor = IIf(comm.CDHolding, &H81B543, &H36312E)
            
        Case comEvRing ', comEvEOF
            lblComStats(3).ForeColor = vbBlack
            lblComStats(3).BackColor = &H81B543
            
            
        Case comEvSend ' something is getting away
            
        
        Case Is > 1000
            
            ' The CommEvent property always returns a numerical value.
            ' Whenever the CommEvent property returns a number
            ' above 1000 then you know that an error occurred.
            txtStatus.Text = "Some ComPort Error occurred: " & comm.CommEvent
            
            
        Case Else
            Debug.Print "whatthefuck"
            ' What happened? It wasn't the arrival of data - and it wasn't
            ' an error. See the ' CommEvent property for a full listing
            ' of all the events and errors.
   End Select
   
   Exit Sub
showError:
   setStatus Err.Description, True, Err.number
   
End Sub

Sub processIncommingMessage()
    Dim i As Long
    Dim tmpBytes() As Long
    
    ReDim tmpBytes(0)
    
    'printBuffer
    
    
    Dim Msg As String
    Dim firstPlace As String
    firstPlace = InStr(1, receiveBuffer, Chr(255))
    If firstPlace = 0 Then Exit Sub
    
    Msg = Left$(receiveBuffer, firstPlace)
    
    receiveBuffer = Right$(receiveBuffer, Len(receiveBuffer) - firstPlace)
    
    Select Case Len(Msg)
        Case 22
            For i = 0 To 10
                picoConnected(i) = IIf(Asc(Mid$(Msg, i + 2, 1)) = 1, True, False)
            Next i
            refreshConnected
        
        Case 2
            Dim errorValue As Long
            errorValue = Asc(Left$(Msg, 1))
            If errorValue = 0 Then
                setStatus "Command successfull!"
            Else
                setStatus "Arduino Error: " & IIf(errorValue > 0 And errorValue < 6, errorMessages(errorValue), "UNKNOWN_ERROR"), True, errorValue
            End If
            
    End Select
    
    
    'Debug.Print UBound(tmpBytes)
    
    processIncommingMessage
End Sub

Sub printBuffer()
    Dim i As Long
    
    Dim tmp As String
    
    tmp = "receiveBufferForShow = {"
    
    For i = 1 To Len(receiveBufferForShow)
        tmp = tmp & "0x" & Hex(Asc(Mid$(receiveBufferForShow, i, 1))) & ", "
    Next i
    
    tmp = tmp & "}"
    
    Debug.Print tmp
End Sub


Private Sub showDummydata()
    receiveBufferForShow = "dit is een hele" & vbCrLf & " lange test om te kijken of de 0 langetestomtekijkenofde0langetestomtekijkenofde0" & vbCrLf & "enters enzo wel goed gaan"
    tmrShowBuffer_Timer
    
End Sub



Private Sub Command1_Click(Index As Integer)
'
'    Select Case Index
'        Case 0
'            txtReceived.AddCharAtCursor Chr(27) & "[45m" & Chr(27) & "[30mH"
'
'        Case 1
'            txtReceived.AddCharAtCursor Chr(27) & "[47m" & Chr(27) & "[31mK"
'
'    End Select
'
'    txtReceived.Redraw

'SetParent Me.hWnd, 332588

'checkForAndOpenLogFile

Dim i As Long
Dim s As String

txtReceived.Text = "hahah" & vbCrLf

For i = 0 To 500
    txtReceived.RedrawPause
    txtReceived.SelStart = txtReceived.TextLength
    txtReceived.AddCharAtCursor "h"
    txtReceived.RedrawResume
Next i


'Dim i As Long
'For i = 0 To 90
'    graphDataInOut.AddItem 0, Sin(i / 180 * 3.1415926535926) * 80, False
'    graphDataInOut.AddItem 1, Sin(i / 180 * 3.1415926535926) * 80, False
'Next i
'
'graphDataInOut.Redraw


End Sub

Private Sub drpBaud_ItemChange(ItemIndex As Long)
    comm.Settings = drpBaud.List(ItemIndex) & ",n,8,1"
End Sub

Private Sub drpCommports_ItemChange(ItemIndex As Long)
    On Error GoTo notWorking
    
    comm.commPort = Replace(serialDevices.commPort(ItemIndex), "COM", "")
    SaveSetting "SerialConsole", "dropdown", "selectedCommPort", serialDevices.commPort(ItemIndex)
    
    setCaption ItemIndex
    
    tmrCheckForReconnect.Enabled = False
    loadReconnect.Loading = False
    
    Exit Sub
notWorking:
    setStatus Err.Description, True, Err.number
    
End Sub

Sub setCaption(Optional Index As Long = -1)
    Dim capt As String
    
    If Index > -1 And Index < serialDevices.Count Then
        capt = serialDevices.commPort(Index) & " - "
    End If
    
    Me.Caption = capt & "SerialConsole - V1.0 by Ricardo de Roode"
End Sub

Sub fillZebroButtons()
    Dim i As Long

    For i = 1 To 10
        Load cmdZebro(i)
        cmdZebro(i).Left = cmdZebro(i - 1).Left + cmdZebro(i - 1).Width + 5 * Screen.TwipsPerPixelX
        
        cmdZebro(i).Visible = True
        cmdZebro(i).Caption = i
        
        Load picConnected(i)
        picConnected(i).Left = picConnected(i - 1).Left + picConnected(i - 1).Width + 5 * Screen.TwipsPerPixelX
        picConnected(i).Visible = True
        picConnected(i).BackColor = &HC0C0FF
    Next i

    refreshConnected
End Sub

Sub refreshConnected()
    Dim i As Long

    For i = 0 To 10
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

Private Sub drpCommports_OnDropdown(cancel As Boolean)
   If comm.PortOpen Then
        cancel = True
        setStatus "Cant change COMMPORT when connected!", True, -1
        Exit Sub
    Else
        If tmrCheckForReconnect.Enabled = True Then
            tmrCheckForReconnect.Enabled = False
            loadReconnect.Loading = False
        End If
    End If
    
    fillCommportList
End Sub

Private Sub drpReceiveSpeed_ItemChange(ItemIndex As Long)
    tmrShowBuffer.Interval = drpReceiveSpeed.ItemData(ItemIndex)
End Sub

Sub setFont()
    Dim i As Long
    
    Dim c As Control
    Dim newFontName As String
    Dim newFontSize As Long
    
    newFontName = "Small Fonts" '"Px437 ATI 8x14"
    newFontSize = 10
    
    Me.FontName = newFontName
    Me.FontSize = newFontSize
    
    
    For Each c In Me.Controls
        If c.Name <> "txtReceived" And c.Name <> "cmdControls" Then
            
            Select Case Left$(c.Name, 3)
                Case "lbl", "drp", "txt", "cmd", "opt", "chk"
                    c.Font.Name = newFontName
                    c.Font.Size = newFontSize
                    
                    If Left$(c.Name, 3) <> "txt" Then c.Font.Bold = True
                    
                    If TypeName(c) = "uCheckBox" Or TypeName(c) = "uOptionBox" Then
                        c.CaptionOffsetTop = 1
                    End If
                    
                    'Debug.Print c.Name
                    'c.FontSize = 10
                
                Case "frm"
                    'Debug.Print c.Name & " " & c.Container.Name
                    c.Font.Name = newFontName
                    c.Font.Size = 8
                    
            End Select
            
            Select Case Left$(c.Name, 3)
                Case "chk", "opt"
                    'Debug.Print "name: " & c.Name & " cap:" & c.Caption & " l:"
                    If c.Container.Name = Me.Name Then
                        c.Width = Me.TextWidth(c.Caption) + 35
                    Else
                        c.Width = (Me.TextWidth(c.Caption) + 35) * Screen.TwipsPerPixelX
                    End If
            End Select
            
            'Debug.Print TypeName(c)
            
            If Left$(TypeName(c), 1) = "u" Then
                c.Redraw
            End If
                    
        End If
    Next c
    
    
End Sub

Private Sub drpWindowType_ItemChange(ItemIndex As Long)
    Dim i As Long
    
    For i = 0 To frmWindow.UBound
        frmWindow(i).Visible = (i = ItemIndex)
    Next i
    
End Sub

Sub fillOnSendList()
    ReDim onSendCharacters(0 To 7) As String
    onSendCharacters(0) = ""
    onSendCharacters(1) = vbCr
    onSendCharacters(2) = vbLf
    onSendCharacters(3) = vbCrLf
    onSendCharacters(4) = vbLf & vbCr
    onSendCharacters(5) = Chr(0)
    onSendCharacters(6) = Chr(255)
    onSendCharacters(7) = ";"
    
    drpOnSend.Clear
    drpOnSend.AddItem "Nothing"
    drpOnSend.AddItem "Cr"
    drpOnSend.AddItem "Lf"
    drpOnSend.AddItem "Cr + LF"
    drpOnSend.AddItem "LF + Cr"
    drpOnSend.AddItem "0x00"
    drpOnSend.AddItem "0xFF"
    drpOnSend.AddItem ";"

    drpOnSend.ItemsVisible = drpOnSend.ListCount
    
End Sub

Sub fillReceiveSpeeds()
    drpReceiveSpeed.AddItem "Realtime", 1
    drpReceiveSpeed.AddItem "Normal", 20
    drpReceiveSpeed.AddItem "Slow", 75
    drpReceiveSpeed.AddItem "Ultra Slow", 200
    drpReceiveSpeed.ItemsVisible = drpReceiveSpeed.ListCount
    
End Sub

Sub fillWindowType()
    drpWindowType.Clear
    
    drpWindowType.AddItem "ZebroMote"
    drpWindowType.AddItem "List"
    drpWindowType.AddItem "Graph"
    drpWindowType.AddItem "Logs"
    drpWindowType.AddItem "History"
    drpWindowType.AddItem "Label List"
    
    drpWindowType.ItemsVisible = drpWindowType.ListCount
End Sub

Sub initLogs()

    If Dir(App.Path & "/logs", vbDirectory) = "" Then
        MkDir App.Path & "/logs"
    End If

End Sub

Sub initToolTips()
    ttToolTip.setForm Me

    'ttToolTip.Add txtReceived.hWnd, "test"
    ttToolTip.Add drpBaud.hWnd, "Select the used baudrate." & vbCrLf & "Can be changed on the fly."
    ttToolTip.Add cmdConnect.hWnd, "Connect/Disconnect." & vbCrLf & "Shows an animation when reconnect is pending."
    ttToolTip.Add chkComOptions(0).hWnd, "Data Terminal Ready." & vbCrLf & "Resets Arduino on connect and on rising toggle."

    ttToolTip.Add chkTxtSettings(1).hWnd, "Enable support for Ansii console colors"
    
    ttToolTip.Add chkTxtSettings(2).hWnd, "Show the received data in hex (Hold: H)"
    ttToolTip.Add chkTxtSettings(0).hWnd, "Automatically scrolldown in textbox when receiving data"
    ttToolTip.Add drpCommports.hWnd, "List of available comports." & vbCrLf & "Click to refresh list."
    ttToolTip.Add chkComOptions(1).hWnd, "Request To Send"
    ttToolTip.Add drpReceiveSpeed.hWnd, "Set the receive window refresh rate." & vbCrLf & "Fully utilized baudrates above 115200 requires lower speed."
    
    
    ttToolTip.Add chkComOptions(2).hWnd, "Clear receive window on connect" & vbCrLf & "or on rising edge of DTR."
    ttToolTip.Add chkComOptions(3).hWnd, "Disconnect when you want to upload" & vbCrLf & "with VisualMicro or the Arduino IDE."
    ttToolTip.Add chkComOptions(4).hWnd, "Automatically connect when disconnect" & vbCrLf & "was triggered by Auto Disconnect."
    ttToolTip.Add chkComOptions(5).hWnd, "Automatically connect when USB is" & vbCrLf & "plugged after it was disconnected."
    


    'ttToolTip.Add .hWnd, ""
    'ttToolTip.Add .hWnd, ""
    'ttToolTip.Add .hWnd, ""
    
    
    
    
    ttToolTip.StartTimer
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim lstCount As Long, totalCount As Long
    
    
    Set serialDevices = New CommPortList
    Set Timer = New PerformanceTimer
    Set inputFilter = New InputHandler
    
    Set tmrCheckBitRate = New SelfTimer
    
    Set arduinoLabels = New Collection
    
    tmrCheckBitRate.Interval = 500
    tmrCheckBitRate.Enabled = True
    
    Erase searchFor
    
    fillCommportList True
    
    fillBaudList
    
    fillZebroButtons
    
    fillLedColors
    
    fillOnSendList
    
    fillReceiveSpeeds
    
    fillWindowType
    
    initToolTips
    
    initLogs
    
    errorMessages = Split(errorMessagesConst, ",")
    
    comm.OutBufferSize = 5
    
    ledCommand = -1
    logFileHandle = -1
    
    'for focus loss of the dropdown menus
    picFocus.Width = 1
    picFocus.Height = 1
    picFocus.Left = -10
    picFocus.Top = -10
    
    graphDataInOut.LineColor(0) = vbRed
    graphDataInOut.LineVisible(0) = True
    graphDataInOut.LineThickness(0) = 1
    
    graphDataInOut.LineColor(1) = vbGreen
    graphDataInOut.LineVisible(1) = True
    graphDataInOut.LineThickness(1) = 1
    
    graphDataInOut.Redraw
    graphDataInOut.AddItem 0, 0, False
    graphDataInOut.AddItem 1, 0, True
    
    dragSplitPercentage = GetSetting("SerialConsole", "UI", "dragSplitPercentage", 0.4964)

    ConsoleColors = Array(vbBlack, vbRed, vbGreen, vbYellow, vbBlue, vbMagenta, vbCyan, vbWhite)
    
    fillArduinoListTestData
    
    setFont
    
    On Error Resume Next
    drpBaud.ListIndex = GetSetting("SerialConsole", "dropdown", "drpBaud.ListIndex", 0)
    drpOnSend.ListIndex = GetSetting("SerialConsole", "dropdown", "drpOnSend.ListIndex", 0)
    drpReceiveSpeed.ListIndex = GetSetting("SerialConsole", "dropdown", "drpReceiveSpeed.ListIndex", 0)
    drpWindowType.ListIndex = GetSetting("SerialConsole", "dropdown", "drpWindowType.ListIndex", 0)
    
    chkLogsEnable.Value = GetSetting("SerialConsole", "logs", "chkLogsEnable.Value", u_unChecked)
    chkSendOnDoubleClick.Value = GetSetting("SerialConsole", "history", "chkSendOnDoubleClick.Value", u_unChecked)
    chkEnableLabelList.Value = GetSetting("SerialConsole", "label", "chkEnableLabelList.Value", u_unChecked)
    
    'loading comport options
    For i = 0 To chkComOptions.UBound
        chkComOptions(i).Value = GetSetting("SerialConsole", "checkboxes", "chkComOptions(" & i & ").Value", u_unChecked)
    Next i
    For i = 0 To chkTxtSettings.UBound
        chkTxtSettings(i).Value = GetSetting("SerialConsole", "checkboxes", "chkTxtSettings(" & i & ").Value", u_unChecked)
    Next i
    
    'loading reconnect checkboxes
    For i = 0 To optLogsReconnect.UBound
        optLogsReconnect(i).Value = GetSetting("SerialConsole", "logs", "optLogsReconnect(" & i & ").Value", u_UnSelected)
        totalCount = totalCount + IIf(optLogsReconnect(i).Value = u_Selected, 1, 0)
    Next i
    If totalCount = 0 Then optLogsReconnect(0).Value = u_Selected
    
    
    'loading history
    lstCount = GetSetting("SerialConsole", "history", "lstHistory.ListCount", 0)
    For i = 0 To lstCount - 1
        lstHistory.AddItem GetSetting("SerialConsole", "history", "List(" & i & ")", ""), CLng(GetSetting("SerialConsole", "history", "ItemData(" & i & ")", 0))
    Next i
    
    'loading label
    lstCount = GetSetting("SerialConsole", "label", "labelCount", 0)
    For i = 0 To lstCount - 1
        AddLabel GetSetting("SerialConsole", "label", i, "")
    Next i
    
    Me.Width = Screen.TwipsPerPixelX * 862
    
    chkComOptions(6).Caption = "Auto disconnect on" & vbCrLf & "focus loss"
    
    'txtReceived.Text = txtReceived.FileToString("F:\Github\SerialConsole\changelog.txt")
End Sub


Private Sub Form_Click()
    picFocus.SetFocus
End Sub

Private Sub Form_LostFocus()
    Debug.Print "lost focus"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Debug.Print X; Y
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Dim i As Long
    
    Dim nominalOffsetX As Long
    Dim smallOffsetX As Long
    Dim firstFrame As Long
    

    nominalOffsetX = 12 * Screen.TwipsPerPixelX
    smallOffsetX = 7 * Screen.TwipsPerPixelX
    
    
    picToolbar(0).Width = Me.ScaleWidth
    picToolbar(1).Width = Me.ScaleWidth
    picToolbar(1).Top = Me.ScaleHeight - picToolbar(1).Height
    picToolbar(1).Left = 0
    
    drpOnSend.Top = smallOffsetX * 2
    txtOutput.Top = drpOnSend.Top + drpOnSend.Height + smallOffsetX
    
    frmOutput.Height = (txtOutput.Top + txtOutput.Height + smallOffsetX) / Screen.TwipsPerPixelX
    frmOutput.Left = 12
    frmOutput.Width = Me.ScaleWidth - frmOutput.Left * 2
    frmOutput.Top = picToolbar(1).Top - frmOutput.Height - 5
    
    
    cmdSend.Left = frmOutput.ScaleWidth - cmdSend.Width - smallOffsetX
    drpOnSend.Left = cmdSend.Left
    cmdSend.Top = drpOnSend.Top + drpOnSend.Height + smallOffsetX
    
    lblInfo(0).Left = drpOnSend.Left - lblInfo(0).Width - nominalOffsetX
    lblInfo(0).Top = drpOnSend.Top + drpOnSend.Height / 2 - lblInfo(0).Height / 2
    chkSend(0).Left = lblInfo(0).Left - chkSend(0).Width - nominalOffsetX
    chkSend(0).Top = drpOnSend.Top + drpOnSend.Height / 2 - chkSend(0).Height / 2
    
    
    txtOutput.Left = smallOffsetX
    txtOutput.Width = cmdSend.Left - txtOutput.Left - smallOffsetX
    
    optInput(0).Left = smallOffsetX
    optInput(0).Top = smallOffsetX * 2
    optInput(0).Height = txtOutput.Top - optInput(0).Top - smallOffsetX
    For i = 1 To optInput.UBound
        optInput(i).Top = optInput(0).Top
        optInput(i).Height = optInput(0).Height
        optInput(i).Left = optInput(i - 1).Left + optInput(i - 1).Width '+ smallOffsetX
    Next i
    
    
    picToolbar(2).Left = 0
    picToolbar(2).Width = Me.ScaleWidth
    picToolbar(2).Top = frmOutput.Top - picToolbar(2).Height
    
    
    'splitter
    txtReceived.Visible = True
    frmTxtSettings.Visible = True
    For i = 0 To lblCursorStats.UBound: lblCursorStats(i).Visible = True: Next i
    'For i = 0 To frmWindow.UBound: frmWindow(i).Visible = True: Next i
    drpWindowType.Visible = True
        
    If dragSplitPercentage = 0 Then
        picSplit.Left = 0
        txtReceived.Visible = False
        frmSearch.Visible = False
        frmTxtSettings.Visible = False
        For i = 0 To lblCursorStats.UBound: lblCursorStats(i).Visible = False: Next i
        drpWindowType_ItemChange drpWindowType.ListIndex
        
    ElseIf dragSplitPercentage = 1 Then
        picSplit.Left = Me.ScaleWidth - picSplit.Width
        For i = 0 To frmWindow.UBound: frmWindow(i).Visible = False: Next i
        drpWindowType.Visible = False
        
    Else
        picSplit.Left = Me.ScaleWidth * dragSplitPercentage
        drpWindowType_ItemChange drpWindowType.ListIndex
    End If
    picSplit.Top = picToolbar(0).Top + picToolbar(0).Height
    picSplit.Height = picToolbar(2).Top - picSplit.Top
    
    
    LBLSplit(0).Left = 0
    LBLSplit(1).Left = Me.ScaleWidth - LBLSplit(1).Width
    LBLSplit(0).Top = Me.ScaleHeight / 2 - LBLSplit(1).Height / 2
    LBLSplit(1).Top = LBLSplit(0).Top
    'end of splitter
    
    
    txtReceived.Left = 12
    firstFrame = picSplit.Left - txtReceived.Left
    
    If frmSearch.Visible Then
        frmSearch.Left = 12
        frmSearch.Top = picToolbar(0).Top + picToolbar(0).Height
        frmSearch.Width = picSplit.Left - frmSearch.Left
        txtReceived.Top = frmSearch.Top + frmSearch.Height + 12
        
        cmdSearchClose.Left = frmSearch.ScaleWidth - cmdSearchClose.Width - smallOffsetX
        cmdSearch.Left = cmdSearchClose.Left - cmdSearch.Width - smallOffsetX
        txtSearch.Left = smallOffsetX
        txtSearch.Width = cmdSearch.Left - txtSearch.Left - smallOffsetX
        
    Else
        txtReceived.Top = picToolbar(0).Top + picToolbar(0).Height
        
    End If
    
    resizeReceivedTextLabels
    
    frmTxtSettings.Top = lblCursorStats(0).Top - frmTxtSettings.Height - 12
    frmTxtSettings.Left = 12
    frmTxtSettings.Width = firstFrame
    
    chkTxtSettings(0).Left = smallOffsetX
    For i = 1 To chkTxtSettings.UBound
        chkTxtSettings(i).Left = smallOffsetX + chkTxtSettings(i - 1).Width + chkTxtSettings(i - 1).Left
    Next i
    
    drpReceiveSpeed.Left = frmTxtSettings.ScaleWidth - drpReceiveSpeed.Width - 3 * Screen.TwipsPerPixelX
    
    
    
    'chkTxtSettings(0).Left = 0
    
    'For i = 1 To chkTxtSettings.UBound
    '    chkTxtSettings(i).Left = chkTxtSettings(i - 1).Left + chkTxtSettings(i).Width
    'Next i
    
    
    
    txtReceived.Width = firstFrame
    txtReceived.Height = frmTxtSettings.Top - txtReceived.Top - 6
    
    
    drpWindowType.Left = picSplit.Left + picSplit.Width + 1
    drpWindowType.Top = picToolbar(0).Top + picToolbar(0).Height
    drpWindowType.Width = Me.ScaleWidth - drpWindowType.Left - 12
    
    frmWindow(0).Left = drpWindowType.Left
    frmWindow(0).Width = drpWindowType.Width
    frmWindow(0).Top = drpWindowType.Top + drpWindowType.Height
    frmWindow(0).Height = picToolbar(2).Top - frmWindow(0).Top
    
    For i = 1 To frmWindow.UBound
        frmWindow(i).Left = frmWindow(0).Left
        frmWindow(i).Top = frmWindow(0).Top
        frmWindow(i).Width = frmWindow(0).Width
        frmWindow(i).Height = frmWindow(0).Height
    Next i
    
    lstArduino.Left = smallOffsetX
    lstArduino.Top = smallOffsetX * 2
    lstArduino.Width = frmWindow(0).ScaleWidth - 2 * smallOffsetX
    lstArduino.Height = frmWindow(0).ScaleHeight - 3 * smallOffsetX
    
    If (Not (Not (arduinoListHeaders))) <> 0 Then
        Dim serperationWidth As Long
        If UBound(arduinoListHeaders) > 0 Then
            serperationWidth = lstArduino.Width / Screen.TwipsPerPixelX / UBound(arduinoListHeaders)
            For i = 0 To UBound(arduinoListHeaders)
                lstArduino.setTabStop i, serperationWidth * i
            Next i
        Else
            lstArduino.setTabStop 0, 0
        End If
        
        
    End If
    
    frmInOut.Left = graphDataInOut.Left + graphDataInOut.Width + 12
    frmInOut.Width = (MyMax(lblInfo(1).Width, lblInfo(2).Width) + 2 * smallOffsetX) / Screen.TwipsPerPixelX
    frmComStats.Left = frmInOut.Left + frmInOut.Width + 12
    
    txtStatus.Top = Me.ScaleHeight - txtStatus.Height - 12
    txtStatus.Left = frmComStats.Left + frmComStats.Width + 12
    txtStatus.Width = Me.ScaleWidth - txtStatus.Left - 12

    cmdClearLabels.Top = cmdClearLabels.Container.ScaleHeight - cmdClearLabels.Height - smallOffsetX
    cmdClearLabels.Left = smallOffsetX
    cmdClearLabels.Width = cmdClearLabels.Container.ScaleWidth - smallOffsetX * 2

    
    
    
    'upper toolbar
    
    picConnectionSettings.Left = Me.ScaleWidth - picConnectionSettings.Width - 12
    frmReconnectSettings.Left = Me.ScaleWidth - frmReconnectSettings.Width - 12
    frmReconnectSettings.Top = picToolbar(0).Top + picToolbar(0).Height
    chkComOptions(1).Left = picConnectionSettings.Left - chkComOptions(1).Width - 12
    chkComOptions(0).Left = chkComOptions(1).Left - chkComOptions(0).Width - 12
    
    'For i = chkComOptions.UBound - 1 To 0 Step -1
    '    chkComOptions(i).Left = chkComOptions(i + 1).Left - chkComOptions(i).Width - 12
    'Next i
    
    loadReconnect.Left = chkComOptions(0).Left - cmdConnect.Width - 12 - 7
    cmdConnect.Left = loadReconnect.Left + 3
    'cmdConnect.Left = chkComOptions(0).Left - cmdConnect.Width - 12
    drpBaud.Left = cmdConnect.Left - drpBaud.Width - 12
    drpCommports.Left = 12
    drpCommports.Width = drpBaud.Left - 12 - drpCommports.Left

    
    'side panels
    
    chkEnableGraph.Left = smallOffsetX
    chkEnableGraph.Top = nominalOffsetX + smallOffsetX
    chkEnableGraph.Width = graphArduino.Container.ScaleWidth - smallOffsetX * 2
    graphArduino.Left = chkEnableGraph.Left
    graphArduino.Top = nominalOffsetX + chkEnableGraph.Top + chkEnableGraph.Height
    graphArduino.Width = chkEnableGraph.Width
    graphArduino.Height = graphArduino.Container.ScaleHeight - graphArduino.Top - smallOffsetX
    
    
    chkLogsEnable.Left = smallOffsetX
    chkLogsEnable.Top = nominalOffsetX + smallOffsetX
    chkLogsEnable.Width = chkLogsEnable.Container.ScaleWidth - smallOffsetX * 2
    frmLogsOnReconnect.Left = smallOffsetX
    
    chkSendOnDoubleClick.Left = smallOffsetX
    chkSendOnDoubleClick.Top = nominalOffsetX + smallOffsetX
    chkSendOnDoubleClick.Width = lstHistory.Width - smallOffsetX * 2
    lstHistory.Left = smallOffsetX
    lstHistory.Top = nominalOffsetX + chkSendOnDoubleClick.Top + chkSendOnDoubleClick.Height
    lstHistory.Width = frmWindow(4).ScaleWidth - 2 * smallOffsetX
    lstHistory.Height = frmWindow(4).ScaleHeight - smallOffsetX - lstHistory.Top
    
    chkEnableLabelList.Left = smallOffsetX
    chkEnableLabelList.Top = nominalOffsetX + smallOffsetX
    chkEnableLabelList.Width = chkEnableLabelList.Container.ScaleWidth - smallOffsetX * 2
    
    'Debug.Print Me.Width
    'Debug.Print dragSplitPercentage
End Sub


Sub fillArduinoListTestData()
    
    ReDim arduinoListHeaders(0 To 4)
    arduinoListHeaders(0) = "PosX"
    arduinoListHeaders(1) = "PosY"
    arduinoListHeaders(2) = "PosZ"
    arduinoListHeaders(3) = "Time"
    
    Dim i As Long
    
    For i = 0 To 10
        lstArduino.AddItem (i * 80) & vbTab & (80 - i * 80) & vbTab & i & vbTab & Round(Rnd + i, 3)
    Next i
    'lstArduino.Redraw
    
End Sub

Function MyMax(ParamArray TheValues() As Variant) As Variant
Dim intLoop As Integer
Dim varCurrentMax As Variant
  varCurrentMax = TheValues(LBound(TheValues))
  For intLoop = LBound(TheValues) + 1 To UBound(TheValues)
    If TheValues(intLoop) > varCurrentMax Then
      varCurrentMax = TheValues(intLoop)
    End If
  Next intLoop
  
  MyMax = varCurrentMax
End Function




Sub fillBaudList()
drpBaud.Clear

Const bauds As String = "300,600,1200,2400,4800,9600,14400,19200,28800,38400,56000,57600,115200,128000,256000"
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




Sub fillCommportList(Optional initializeForm As Boolean = False)
'On Error Resume Next
    Dim newCommPortIndex As Long
    Dim prevCommPort As String
    Dim i As Long
    
    If initializeForm Then
        prevCommPort = GetSetting("SerialConsole", "dropdown", "selectedCommPort", "")
    Else
        If serialDevices.Count > 0 And drpCommports.ListIndex <> -1 Then
            prevCommPort = serialDevices.commPort(drpCommports.ListIndex)
        End If
    End If
    
    serialDevices.Refresh

    drpCommports.Clear
    newCommPortIndex = -1
    
    For i = 0 To serialDevices.Count - 1
        If prevCommPort <> "" And prevCommPort = serialDevices.commPort(i) Then newCommPortIndex = i
        drpCommports.AddItem serialDevices.friendlyName(i) & " (" & serialDevices.commPort(i) & ", " & serialDevices.locationInformation(i) & ")", i, , -1
    Next i
    
    drpCommports.ItemsVisible = IIf(serialDevices.Count < 10, serialDevices.Count, 10)
    
    If serialDevices.Count <> 0 Then
        If newCommPortIndex = -1 Then newCommPortIndex = 0
        drpCommports.ListIndex = newCommPortIndex
        Me.Caption = serialDevices.commPort(newCommPortIndex) & " - SerialConsole - V1.0 by Ricardo de Roode"
    End If
    
    If drpCommports.ListCount = 0 Then
        drpCommports.Text = "No Devices (click to refresh)"
        setCaption
    Else
        setCaption drpCommports.ListIndex
    End If
    
    
    
End Sub


Private Sub Form_Unload(cancel As Integer)
    SaveSetting "SerialConsole", "dropdown", "drpBaud.ListIndex", drpBaud.ListIndex
    SaveSetting "SerialConsole", "dropdown", "drpOnSend.ListIndex", drpOnSend.ListIndex
    SaveSetting "SerialConsole", "dropdown", "drpReceiveSpeed.ListIndex", drpReceiveSpeed.ListIndex
    SaveSetting "SerialConsole", "dropdown", "drpWindowType.ListIndex", drpWindowType.ListIndex
    SaveSetting "SerialConsole", "UI", "dragSplitPercentage", dragSplitPercentage
    
    Dim i As Long
    
    SaveSetting "SerialConsole", "history", "lstHistory.ListCount", lstHistory.ListCount
    
    For i = 0 To lstHistory.ListCount - 1
        SaveSetting "SerialConsole", "history", "List(" & i & ")", lstHistory.List(i)
        SaveSetting "SerialConsole", "history", "ItemData(" & i & ")", lstHistory.ItemData(i)
    Next i
    
    
    SaveSetting "SerialConsole", "label", "labelCount", arduinoLabels.Count
    
    Dim v As Variant
    For Each v In arduinoLabels
        SaveSetting "SerialConsole", "label", CLng(v(0)), CStr(v(1))
    Next v
    
    If comm.PortOpen Then comm.PortOpen = False
    DoEvents
End Sub

Private Sub lblInfo_DblClick(Index As Integer)
    Select Case Index
        Case 1
            bitsSend = 0
            changeBitsSendReceived
        Case 2
            bitsReceived = 0
            changeBitsSendReceived
    End Select
End Sub

Private Sub LBLSplit_Click(Index As Integer)
    
    Select Case Index
        Case 0
            dragSplitPercentage = 0
            
        Case 1
            dragSplitPercentage = 1
    
    End Select
    
    Form_Resize
End Sub

Private Sub lstHistory_DblClick()
    Dim i As Long
    
    i = lstHistory.ListIndex
    If i <> -1 Then
        setOutputOptionsWithLong lstHistory.ItemData(i)
        txtOutput.Text = lstHistory.List(i)
        If chkSendOnDoubleClick.Value = u_Checked Then cmdSend_Click 0, 0, 0
    End If
    
End Sub

Private Sub optInput_ActivateNextState(Index As Integer, u_Cancel As Boolean, u_NewState As uOptionBoxConstants)
    If optInput(Index).Value = u_Selected Then
        u_NewState = u_UnSelected
        u_Cancel = True
    End If
    txtOutput.SetFocus
    
End Sub

Private Sub optInput_Changed(Index As Integer, u_NewState As uOptionBoxConstants)
    txtOutput_Changed
    txtSearch_Changed
End Sub

Private Sub optLogsReconnect_Changed(Index As Integer, u_NewState As uOptionBoxConstants)
    Dim i As Long
    
    For i = 0 To optLogsReconnect.UBound
        SaveSetting "SerialConsole", "logs", "optLogsReconnect(" & i & ").Value", optLogsReconnect(i).Value
    Next i
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


Private Sub picConnectionSettings_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmReconnectSettings.Visible = Not frmReconnectSettings.Visible
End Sub

Private Sub picSplit_Click()
    If dragSplitPercentage = 0 Then
        dragSplitPercentage = 0.1
    ElseIf dragSplitPercentage = 1 Then
        dragSplitPercentage = 0.9
    End If
    
    Form_Resize
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dragSplitStartX = X
    dragSplit = True
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static isResizing As Boolean
    
    If isResizing Then Exit Sub
    isResizing = True
    
    If dragSplit = True Then
        Dim newLeft As Double
        
        newLeft = picSplit.Left - (dragSplitStartX - X)
        
        newLeft = 1# / Me.ScaleWidth * newLeft
        If newLeft < 0.1 Then newLeft = 0.1
        If newLeft > 0.9 Then newLeft = 0.9
        
        dragSplitPercentage = newLeft
        
        DoEvents
        
        Form_Resize
        
    End If
    
    isResizing = False
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dragSplit = False
End Sub

Private Sub tmrCheckBitRate_Timer(ByVal Seconds As Currency)
    Static timerPart As Boolean
    
    If Seconds = 0 Then Exit Sub
    
    Select Case timerPart
        Case True
            'Debug.Print "--- " & Seconds
            If comm.PortOpen And drpCommports.ListIndex > -1 Then
                If serialDevices.isCommAvailable(drpCommports.ListIndex) = False Then
                    cmdConnect_Click 0, 0, 0
                    setStatus "Device was removed unexpectedly!", True, -1
                    If chkComOptions(5).Value = u_Checked Then
                        tmrCheckForReconnect.Enabled = True
                        loadReconnect.Loading = True
                    End If
                End If
            End If
        
        Case False
            'Debug.Print "--> " & Seconds
            graphDataInOut.AddItem 0, CDbl(bitrateInbound), False
            graphDataInOut.AddItem 1, CDbl(bitrateOutbound), False
            
            bitrateInbound = 0
            bitrateOutbound = 0
            
            graphDataInOut.ScrollToLastItem 0, True
            graphDataInOut.Redraw
            
    End Select
    
    timerPart = Not timerPart
    
    If chkComOptions(6).Value = u_Checked Then
    
        If comm.PortOpen And GetForegroundWindow() <> Me.hWnd Then
            cmdConnect_Click 0, 0, 0
            setStatus "Disconnected by focus loss"
        End If
        
    End If
    'Debug.Print Me.ActiveControl.Name
    

    
    
End Sub



Private Sub tmrCheckForReconnect_Timer()
    If comm.PortOpen = False Then
        cmdConnect_Click 0, 0, 0
        If comm.PortOpen = False Then
            setStatus "Checking for automatic connect..."
        End If
    Else
        tmrCheckForReconnect.Enabled = False
        loadReconnect.Loading = False
        Exit Sub
    End If
    
End Sub

Private Sub txtDataExchange_Change()
    If txtDataExchange.Text = "" Then Exit Sub
    
    Dim strSplit() As String
    
    strSplit = Split(txtDataExchange.Text, " ")
    
    txtDataExchange.Text = ""
    
    If chkComOptions(3).Value <> u_Checked Then Exit Sub
    
    If UBound(strSplit) <> 1 Then
        MsgBox "Not a valid message!"
        Exit Sub
    End If
    
    Select Case strSplit(0)
        Case "CC"
            'MsgBox "close port " & strSplit(1)
            If comm.PortOpen = True Then
                If strSplit(1) = "{serial.port}" Or serialDevices.commPort(drpCommports.ListIndex) = strSplit(1) Then
                    cmdConnect_Click 0, 0, 0
                    If chkComOptions(4).Value = u_Checked Then
                        tmrCheckForReconnect.Enabled = False
                        tmrCheckForReconnect.Enabled = True
                        loadReconnect.Loading = True
                    End If
                    
                End If
            End If
            
        Case "OC"
            'MsgBox "open port " & strSplit(1)
            
            If comm.PortOpen = False Then
                If strSplit(1) = "{serial.port}" Or serialDevices.commPort(drpCommports.ListIndex) = strSplit(1) Then
                    cmdConnect_Click 0, 0, 0
                End If
            End If
            
            
            
    End Select
    
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
    
    commOut variantBytes
    'comm.Output = variantBytes
End Sub

Sub commOut(bytes As Variant)
   
    If comm.PortOpen = False Then Exit Sub
    
    bitrateOutbound = bitrateOutbound + UBound(bytes)
    bitsSend = bitsSend + UBound(bytes)
    changeBitsSendReceived
    
    On Error GoTo disconnectFromDevice
    comm.Output = bytes
    
Exit Sub
disconnectFromDevice:
    cmdConnect_Click 0, 0, 0
    setStatus Err.Description, True, Err.number
    If chkComOptions(5).Value = u_Checked Then
        tmrCheckForReconnect.Enabled = True
        loadReconnect.Loading = True
    End If
    
End Sub

Private Sub tmrShowBuffer_Timer()
    'receiveBufferForShow = "lolabcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ lolabcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    If receiveBufferForShowLength = 0 Then Exit Sub
    'Clipboard.Clear
    'Clipboard.SetText receiveBuffer
    'printBuffer
    
    'Timer.StartTimer
    
    Dim tmpSelStart As Long
    Dim tmpSelLength As Long
    
    tmpSelStart = txtReceived.SelStart
    tmpSelLength = txtReceived.SelLength
    
    txtReceived.RedrawPause
    txtReceived.SelStart = txtReceived.TextLength
    
    txtReceived.AddCharAtCursor Left$(receiveBufferForShow, receiveBufferForShowLength), True
    
    fillReceivedTextColors txtReceived.TextLength - Len(receiveBufferForShow)
    
    txtReceived.SelStart = tmpSelStart
    txtReceived.SelLength = tmpSelLength
    
    txtReceived.RedrawResume
    
    If chkTxtSettings(0).Value = u_Checked Then txtReceived.ScrollToEnd
    
    If chkLogsEnable.Value = u_Checked And logFileHandle <> -1 Then
        Put logFileHandle, , Mid$(receiveBufferForShow, 1, receiveBufferForShowLength)
    End If
    
    receiveBufferForShowLength = 0
    
    changeBitsSendReceived
    
    
    If chkRefreshZebro.Value = u_Checked Then
        processIncommingMessage
    End If
    
    
    
    
    'Timer.StopTimer
    
    'Debug.Print Timer.TimeElapsed(pvMilliSecond)
    
    '##################
    '## Arduino Part ##
    '##################
    
    
    If comm.PortOpen = False Then
        tmrShowBuffer.Enabled = False
    End If
    
    ProcessGraphData
    
End Sub


Sub ProcessGraphData()
    'On Error Resume Next
    Dim tmpSplit() As String, tmpValue() As String
    Dim i As Long, j As Long
    
    If (drpWindowType.ListIndex <> 2 And drpWindowType.ListIndex <> 5) Or _
        (chkEnableGraph.Value <> u_Checked And chkEnableLabelList.Value <> u_Checked) Then
        receiveBufferArduinoLength = 0
        Exit Sub
    End If
    
    
    If InStr(1, receiveBufferArduino, vbCrLf) > 0 Then
        If chkEnableGraph.Value = u_Checked Then
            
            tmpSplit = Split(Left$(receiveBufferArduino, receiveBufferArduinoLength), vbCrLf)
            
            
            For i = 0 To UBound(tmpSplit)
                If i = UBound(tmpSplit) Then
                    If Len(tmpSplit(UBound(tmpSplit))) > 0 Then
                        receiveBufferArduinoLength = Len(tmpSplit(UBound(tmpSplit)))
                        Mid$(receiveBufferArduino, 1, receiveBufferArduinoLength) = tmpSplit(UBound(tmpSplit))
                        Exit For
                    Else
                        receiveBufferArduinoLength = 0
                    End If
                End If
                
                tmpValue = Split(tmpSplit(i), " ")
                For j = 0 To UBound(tmpValue)
                    If Val(tmpValue(j)) = tmpValue(j) Then
                        graphArduino.LineVisible(j) = True
                        graphArduino.AddItem j, Val(tmpValue(j))
                    End If
                Next j
            Next i
            graphArduino.ScrollToLastItem 0, True
            
            graphArduino.Redraw
            
        ElseIf chkEnableLabelList.Value = u_Checked Then
            tmpSplit = Split(Left$(receiveBufferArduino, receiveBufferArduinoLength), vbCrLf)
            
            For i = 0 To UBound(tmpSplit)
                
                If Len(tmpSplit(i)) > 0 Then
                    If i = UBound(tmpSplit) Then
                        receiveBufferArduinoLength = Len(tmpSplit(UBound(tmpSplit)))
                        Mid$(receiveBufferArduino, 1, receiveBufferArduinoLength) = tmpSplit(UBound(tmpSplit))
                    Else
                        tmpValue = Split(tmpSplit(i), ":")
                        If UBound(tmpValue) > 0 Then
                            If Len(tmpValue(0)) > 0 Then
                                CheckIfLabelIsAddedOrAdd tmpValue(0), tmpSplit(i)
                            End If
                        End If
                    End If
                ElseIf i = UBound(tmpSplit) Then
                    'the end of the splitter
                    receiveBufferArduinoLength = 0
                End If

            Next i
        End If
    
    End If

End Sub

Sub CheckIfLabelIsAddedOrAdd(whatLabel As String, whatValue As String)
    On Error GoTo ExistsNonObjectErrorHandler
    Dim triedBefore As Boolean
    Dim Index As Long
    
    triedBefore = False
    
tryAgain:
    Index = arduinoLabels(whatLabel)(0)
    lblLabel(Index).Caption = whatValue
    
    Exit Sub
ExistsNonObjectErrorHandler:
    'not found, add here
    AddLabel whatLabel
    
    If triedBefore = False Then
        triedBefore = False
        GoTo tryAgain
    End If
End Sub


Sub AddLabel(labelDescription As String)
    Dim i As Long
    Dim a(0 To 1) As Variant
    
    a(0) = arduinoLabels.Count
    a(1) = labelDescription
    
    arduinoLabels.Add a, CStr(a(1))

    If a(0) > 0 Then
        Load lblLabel(a(0))
        
        lblLabel(a(0)).Top = lblLabel(a(0) - 1).Top + lblLabel(a(0) - 1).Height * 2
        lblLabel(a(0)).Left = lblLabel(a(0) - 1).Left
    End If
    
    lblLabel(a(0)).Visible = True
    lblLabel(a(0)).Caption = a(1) & ": "
    
End Sub




Sub changeBitsSendReceived()
    Dim bitsIn As String
    Dim bitsOut As String
    Dim mustResize As Boolean
    
    mustResize = False
    bitsIn = "IN:  " & bitsReceived
    bitsOut = "OUT: " & bitsSend
    
    If Len(lblInfo(1).Caption) <> Len(bitsOut) Then mustResize = True
    lblInfo(1).Caption = bitsOut

    If Len(lblInfo(2).Caption) <> Len(bitsIn) Then mustResize = True
    lblInfo(2).Caption = bitsIn
    
    If mustResize Then Form_Resize
End Sub

Sub parseInputColors(uTxt As uTextBox, uCmd As uButton, checkConnected As Boolean)
Dim i As Long, j As Long
    Dim lPlace As Long
    
    Dim str As String
    Dim splitStr() As String
    Dim splitLength As Long
    
    Dim partColor As Long
    Dim partType As ParseType
    
    str = uTxt.Text
    splitStr = Split(str, " ")
    
    uTxt.RedrawPause
    If checkConnected And comm.PortOpen Then uCmd.Enabled = True
    
    Dim forceFunction As Long
    forceFunction = -1
    
    For i = optInput.LBound To optInput.UBound
        If optInput(i).Value = u_Selected Then
            forceFunction = i
            Exit For
        End If
    Next i
    
    
    For i = 0 To Len(str) - 1
        uTxt.setCharForeColor i, -1
        uTxt.setCharBold i, 0
    Next i

    'txtInput.RedrawResume
    'Exit Sub
    
    For i = 0 To UBound(splitStr)
        splitLength = Len(splitStr(i))
        
        If splitLength > 0 Then
            partType = inputFilter.getTypeByString(splitStr(i), forceFunction)
            If partType <> ParseType.pNoColor Then
                partColor = inputFilter.getColorByType(partType)
            Else
                partColor = -1
                uCmd.Enabled = False
            End If
            
            For j = 0 To splitLength - 1
                uTxt.setCharForeColor j + lPlace, partColor
                uTxt.setCharBold j + lPlace, IIf(partColor <> -1, 1, 0)
                'txtInput.setCharForeColor j + lPlace, IIf(partColor = -1, vbBlack, vbWhite)
                
            Next j
        End If
        
        lPlace = lPlace + splitLength + 1
    Next i
    
    
    uTxt.RedrawResume
End Sub


Private Sub txtOutput_Changed()
    txtOutput_GotFocus
    
    parseInputColors txtOutput, cmdSend, True
    
    'l = Asc(Mid(firstText, i, 1))
    'tmpStr = tmpStr & IIf(l < 16, "0", "") & Hex(l) & " "
End Sub



Private Sub txtOutput_GotFocus()
    txtOutput.BorderColor = &H81B543
    txtOutput.BackgroundColor = vbWhite
End Sub

Private Sub txtOutput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSend_Click 0, 0, 0
        KeyCode = 0
        Shift = 0
    End If
    
End Sub

Private Sub txtOutput_LostFocus()
    txtOutput.BorderColor = &H808080
End Sub

Sub fillReceivedTextColors(startChar As Long)
    Dim i As Long, j As Long
    Dim s() As Byte
    Dim t As String
    
    Dim mayColor As Boolean
    
    If startChar < 0 Then Exit Sub

    s = txtReceived.RawText
    'txtReceived.RedrawPause
    
    If (Not (Not searchFor)) = 0 Then
        Exit Sub
    End If
    
    
    For i = startChar To UBound(s)
        If s(i) = searchFor(0) Then
            mayColor = True
            For j = 1 To UBound(searchFor)
                If i + j < UBound(s) Then
                    If s(i + j) <> searchFor(j) Then
                        mayColor = False
                        Exit For
                    End If
                Else
                    mayColor = False
                    Exit For
                End If
            Next j
            
            If mayColor Then
                For j = 0 To UBound(searchFor)
                    txtReceived.setCharBorderColor i + j, vbBlue
                Next j
                i = i + UBound(searchFor)
            Else
                txtReceived.setCharBorderColor i, -1
            End If
            
        Else
            txtReceived.setCharBorderColor i, -1
        End If
        
    Next i


    'txtReceived.RedrawResume
End Sub

Private Sub txtReceived_Click(ByVal charIndex As Long, ByVal charRow As Long)
    'Debug.Print charIndex & " " & charRow
End Sub

Private Sub txtReceived_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 And Shift = 2 Then
        frmSearch.Visible = True
        Form_Resize
        txtSearch.SetFocus
        txtSearch.SelStart = 0
        txtSearch.SelLength = txtSearch.TextLength
    ElseIf KeyCode = vbKeyH Then
        chkTxtSettings(2).Value = u_Checked
        KeyCode = 0
        Shift = 0
    End If
End Sub

Private Sub txtReceived_KeyPress(KeyAscii As Integer)
    'KeyAscii = 0
End Sub

Private Sub txtReceived_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyH Then
        chkTxtSettings(2).Value = u_unChecked
    End If
End Sub

Private Sub txtReceived_OnCursorPositionChanged(ByVal charIndex As Long, ByVal charRow As Long, ByVal charCol As Long, ByVal charVal As Byte)
    lblCursorStats(0).Caption = "index: " & charIndex
    lblCursorStats(1).Caption = "row: " & charRow
    lblCursorStats(2).Caption = "col: " & charCol
    lblCursorStats(3).Caption = "sel: " & txtReceived.SelLength
    lblCursorStats(4).Caption = "val: " & IIf(charIndex = txtReceived.TextLength, "-- (---)", charVal & "(0x" & Hex(charVal) & ")")
    resizeReceivedTextLabels
End Sub

Sub resizeReceivedTextLabels()
    Dim firstFrame As Long

    firstFrame = picSplit.Left - txtReceived.Left
        
    lblCursorStats(0).Top = picToolbar(2).Top - lblCursorStats(0).Height
    lblCursorStats(0).Left = 12
    lblCursorStats(0).Width = Fix(firstFrame / 5)
    
    lblCursorStats(1).Top = lblCursorStats(0).Top
    lblCursorStats(1).Left = lblCursorStats(0).Left + lblCursorStats(0).Width
    lblCursorStats(1).Width = Fix(firstFrame / 5)
    
    lblCursorStats(2).Top = lblCursorStats(0).Top
    lblCursorStats(2).Left = lblCursorStats(1).Left + lblCursorStats(1).Width
    lblCursorStats(2).Width = lblCursorStats(1).Width
    
    lblCursorStats(4).Top = lblCursorStats(0).Top
    'lblCursorStats(4).Width = firstFrame - (lblCursorStats(3).Width + lblCursorStats(3).Left)
    lblCursorStats(4).Left = firstFrame - lblCursorStats(4).Width + txtReceived.Left
    
    lblCursorStats(3).Top = lblCursorStats(0).Top
    lblCursorStats(3).Left = lblCursorStats(2).Left + lblCursorStats(2).Width
    lblCursorStats(3).Width = lblCursorStats(4).Left - lblCursorStats(3).Left
    


End Sub






Private Sub txtSearch_Changed()
    txtSearch_GotFocus
    parseInputColors txtSearch, cmdSearch, False
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.BorderColor = &H81B543
    txtSearch.BackgroundColor = vbWhite
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        Shift = 0
        cmdSearch_Click 0, 0, 0
    ElseIf KeyCode = vbKeyEscape Then
        cmdSearchClose_Click 0, 0, 0
        KeyCode = 0
        Shift = 0
    End If
End Sub

