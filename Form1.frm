VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0036312E&
   Caption         =   "ZebroMote - V1.0 by Ricardo de Roode"
   ClientHeight    =   7500
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
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   985
   StartUpPosition =   3  'Windows Default
   Begin Project1.uTextBox txtStatus 
      Height          =   420
      Left            =   270
      TabIndex        =   4
      Top             =   6885
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   741
      BorderColor     =   2367774
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   2367774
   End
   Begin VB.PictureBox picToolbar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0024211E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   1
      Left            =   0
      ScaleHeight     =   810
      ScaleWidth      =   14730
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6690
      Width           =   14730
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   45
      Width           =   150
   End
   Begin Project1.uFrame frmColors 
      Height          =   615
      Left            =   2835
      TabIndex        =   20
      Top             =   4755
      Visible         =   0   'False
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   1085
      BackgroundColor =   3551534
      BorderColor     =   14737632
      ForeColor       =   16777215
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
      Left            =   4320
      TabIndex        =   10
      Top             =   1800
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
      Left            =   13620
      Top             =   3780
   End
   Begin VB.PictureBox picConnected 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   2835
      ScaleHeight     =   255
      ScaleWidth      =   390
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   900
      Width           =   420
   End
   Begin VB.Timer tmrCheckPortOpen 
      Interval        =   1000
      Left            =   12660
      Top             =   3795
   End
   Begin VB.VScrollBar scrollTest 
      Height          =   3015
      LargeChange     =   20
      Left            =   14220
      Max             =   10
      TabIndex        =   5
      Top             =   1005
      Visible         =   0   'False
      Width           =   255
   End
   Begin Project1.uCheckBox chkCommOptions 
      Height          =   465
      Index           =   0
      Left            =   10395
      TabIndex        =   3
      Top             =   180
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   820
      BackgroundColor =   2367774
      Border          =   0   'False
      BorderColor     =   2367774
      Caption         =   "DTR"
      CaptionOffsetLeft=   5
      CheckBackgroundColor=   2367774
      CheckBorderColor=   8421504
      CheckBorderThickness=   2
      CheckSelectionColor=   4210752
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
   Begin Project1.uButton cmdConnect 
      Height          =   465
      Left            =   8055
      TabIndex        =   2
      Top             =   180
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   820
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
   Begin Project1.uDropDown drpCommports 
      Height          =   465
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   820
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
      ScrollBarWidth  =   30
   End
   Begin MSCommLib.MSComm comm 
      Left            =   12405
      Top             =   4530
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
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
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
   Begin Project1.uTextBox txtReceived 
      Height          =   810
      Left            =   12915
      TabIndex        =   6
      Top             =   1320
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
      Left            =   11565
      TabIndex        =   7
      Top             =   180
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   820
      BackgroundColor =   2367774
      Border          =   0   'False
      BorderColor     =   2367774
      Caption         =   "RTS"
      CaptionOffsetLeft=   5
      CheckBackgroundColor=   2367774
      CheckBorderColor=   8421504
      CheckBorderThickness=   2
      CheckSelectionColor=   4210752
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
   Begin Project1.uButton cmdZebro 
      Height          =   375
      Index           =   0
      Left            =   2835
      TabIndex        =   8
      Top             =   1260
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
      Left            =   5805
      TabIndex        =   11
      Top             =   1800
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
      Left            =   2835
      TabIndex        =   12
      Top             =   1800
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
      Left            =   4320
      TabIndex        =   13
      Top             =   3285
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
      Left            =   2835
      TabIndex        =   14
      Top             =   3285
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
      Left            =   2835
      TabIndex        =   15
      Top             =   3780
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
      Left            =   2835
      TabIndex        =   16
      Top             =   4275
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
      Left            =   5805
      TabIndex        =   17
      Top             =   3285
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
      Left            =   5805
      TabIndex        =   18
      Top             =   3780
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
      Left            =   5805
      TabIndex        =   19
      Top             =   4275
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
      Left            =   7290
      TabIndex        =   22
      Top             =   1800
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
      Left            =   7290
      TabIndex        =   23
      Top             =   3285
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
   Begin VB.PictureBox picToolbar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0024211E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   0
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14730
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   14730
   End
   Begin Project1.uFrame frmOutput 
      Height          =   960
      Left            =   180
      TabIndex        =   27
      Top             =   5460
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   1693
      BackgroundColor =   3551534
      BorderColor     =   14737632
      ForeColor       =   16777215
      Caption         =   "Send Information"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Project1.uOptionBox optInput 
         Height          =   315
         Index           =   1
         Left            =   1980
         TabIndex        =   30
         Top             =   180
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         BackgroundColor =   3551534
         Border          =   0   'False
         Caption         =   "HEX"
         CheckBackgroundColor=   3551534
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
      Begin Project1.uTextBox txtInput 
         Height          =   330
         Left            =   90
         TabIndex        =   28
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
      End
      Begin Project1.uButton cmdSend 
         Height          =   330
         Left            =   9405
         TabIndex        =   29
         Top             =   540
         Width           =   2130
         _ExtentX        =   3757
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
         Caption         =   "Send"
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
      Begin Project1.uOptionBox optInput 
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   31
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         BackgroundColor =   3551534
         Border          =   0   'False
         Caption         =   "ANSII"
         CheckBackgroundColor=   3551534
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
      Begin Project1.uCheckBox chkSend 
         Height          =   285
         Index           =   0
         Left            =   10695
         TabIndex        =   32
         Top             =   165
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
         BackgroundColor =   3551534
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "COS"
         CaptionOffsetLeft=   5
         CheckBackgroundColor=   3551534
         CheckBorderColor=   8421504
         CheckBorderThickness=   2
         CheckSelectionColor=   4210752
         CheckSize       =   1
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
      Begin Project1.uOptionBox optInput 
         Height          =   315
         Index           =   2
         Left            =   1170
         TabIndex        =   33
         Top             =   180
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         BackgroundColor =   3551534
         Border          =   0   'False
         Caption         =   "BIN"
         CheckBackgroundColor=   3551534
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
      Begin Project1.uOptionBox optInput 
         Height          =   315
         Index           =   3
         Left            =   2835
         TabIndex        =   34
         Top             =   165
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         BackgroundColor =   3551534
         Border          =   0   'False
         Caption         =   "DEC"
         CheckBackgroundColor=   3551534
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
      Begin Project1.uOptionBox optInput 
         Height          =   315
         Index           =   4
         Left            =   3645
         TabIndex        =   35
         Top             =   165
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         BackgroundColor =   3551534
         Border          =   0   'False
         Caption         =   "OCTAL"
         CheckBackgroundColor=   3551534
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


Dim serialDevices As CommPortList
Dim timer As PerformanceTimer
Dim inputFilter As InputHandler


Dim picoSendCommand(0 To 19) As Boolean
Dim picoConnected(0 To 19) As Boolean

Dim lastMessageBytes() As Byte

Dim receiveBuffer As String
Dim ledCommand As Long

Dim errorMessages() As String
Const errorMessagesConst = ",M_ERROR,M_ERROR_NOT_CONNECTED,M_ERROR_BUFFER_OVERFLOW,M_ERROR_BUFFER_EMPTY,M_ERROR_UNKNOWN_COMMAND"


Private Sub chkCommOptions_Changed(Index As Integer, u_NewState As uCheckboxConstants)
    Dim newState As Boolean
    
    newState = (u_NewState = u_Checked)
    
    Select Case Index
    
        Case 0
            comm.DTREnable = newState
            
        Case 1
            comm.RTSEnable = newState
    End Select
    
    setCheckColors chkCommOptions(Index), newState
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

Private Sub chkSend_Changed(Index As Integer, u_NewState As uCheckboxConstants)
    setCheckColors chkSend(Index), u_NewState = u_Checked
    
End Sub

Private Sub cmdConnect_Click(Button As Integer, X As Single, Y As Single)
    On Error GoTo notWorking
    
    If comm.PortOpen Then
        comm.PortOpen = False
        cmdConnect.Caption = "Connect"
        cmdConnect.BackgroundColor = &H4747F0
        
    Else
        cmdConnect.BackgroundColor = &H81B543
        cmdConnect.Caption = "Disconnect"
        comm.PortOpen = True
        
        setStatus "Connected!"
    End If
    
Exit Sub
notWorking:
    cmdConnect.BackgroundColor = &H4747F0
    cmdConnect.Caption = "Connect"
    
    setStatus Err.Description, True, Err.Number
    
End Sub

Sub setStatus(msg As String, Optional isError As Boolean = False, Optional errorNumber As Long = 0)
    txtStatus.RedrawPause
    
    If isError Then
        txtStatus.Text = "[ ERROR " & errorNumber & " ] " & msg
        txtStatus.ForeColor = &H4747F0
        txtStatus.BorderColor = &H4747F0
    Else
        txtStatus.Text = msg
        txtStatus.ForeColor = &H81B543
        txtStatus.BorderColor = &H81B543
    End If
    
    txtStatus.RedrawResume
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

Private Sub cmdSend_Click(Button As Integer, X As Single, Y As Single)
    Dim str As String
    Dim splitStr() As String
    Dim i As Long, j As Long
    
    
    str = txtInput.Text
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
    
    For i = 0 To UBound(splitStr)
        parseBytes = inputFilter.parseString(splitStr(i), forceFunction)
        If UBound(parseBytes) > -1 Then
            
            ReDim Preserve totalBytes(0 To totalBytesLength + UBound(parseBytes))
            
            For j = 0 To UBound(parseBytes)
                totalBytes(totalBytesLength + j) = parseBytes(j)
            Next j
            
            totalBytesLength = totalBytesLength + UBound(parseBytes) + 1
            
        End If
        
    Next i
    
    Dim tmpStr As String
    
    For i = 0 To totalBytesLength - 1
        tmpStr = tmpStr & totalBytes(i) & " "
    Next i
    
    MsgBox tmpStr
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
    
    comm.commPort = Replace(serialDevices.commPort(ItemIndex), "COM", "")
    SaveSetting "SerialConsole", "dropdown", "selectedCommPort", serialDevices.commPort(ItemIndex)
     
    Exit Sub
notWorking:
    setStatus Err.Description, True, Err.Number
    
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

Private Sub drpCommports_OnDropdown(cancel As Boolean)
   If comm.PortOpen Then
        cancel = True
        setStatus "Cant change COMMPORT when connected!", True, -1
        Exit Sub
    End If
    
    fillCommportList
End Sub

Private Sub Form_Click()
    picFocus.SetFocus
End Sub

Private Sub Form_Load()
    Set serialDevices = New CommPortList
    Set timer = New PerformanceTimer
    Set inputFilter = New InputHandler
    
    
    fillCommportList True
    
    fillBaudList
    
    fillZebroButtons
    
    fillLedColors
    
    errorMessages = Split(errorMessagesConst, ",")
    
    comm.OutBufferSize = 5
    
    On Error Resume Next
    drpBaud.ListIndex = GetSetting("SerialConsole", "dropdown", "drpBaud.ListIndex", 0)
    
    
    ledCommand = -1
    
    
    'for focus loss of the dropdown menus
    picFocus.Width = 1
    picFocus.Height = 1
    picFocus.Left = -10
    picFocus.Top = -10
    
    'global uControl settings
    
    'txtReceived.ReCalculateMarkup
    'txtReceived.ReCalculateWords
    'txtReceived.Redraw
    
End Sub



Private Sub Form_Resize()
On Error Resume Next

    Dim nominalOffsetX As Long
    nominalOffsetX = 6 * Screen.TwipsPerPixelX

    txtReceived.Left = 0
    txtReceived.Width = Me.ScaleWidth
    txtReceived.Height = Me.ScaleHeight - txtReceived.Top - txtStatus.Height - 32
    
    txtStatus.Top = Me.ScaleHeight - txtStatus.Height - 12
    txtStatus.Left = 12
    txtStatus.Width = Me.ScaleWidth - txtStatus.Left * 2
    
    picToolbar(0).Width = Me.ScaleWidth
    picToolbar(1).Width = Me.ScaleWidth
    picToolbar(1).Top = Me.ScaleHeight - picToolbar(1).Height
    
    
    frmOutput.Width = Me.ScaleWidth - frmOutput.Left * 2
    frmOutput.Top = picToolbar(1).Top - frmOutput.Height - 12
    
    
    chkSend(0).Left = frmOutput.ScaleWidth - chkSend(0).Width - nominalOffsetX
    cmdSend.Left = frmOutput.ScaleWidth - cmdSend.Width - nominalOffsetX
    
    txtInput.Width = cmdSend.Left - txtInput.Left - nominalOffsetX
    
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
    
    drpCommports.ItemsVisible = serialDevices.Count
    
    If serialDevices.Count <> 0 Then
        If newCommPortIndex <> -1 Then
            drpCommports.ListIndex = newCommPortIndex
        Else
             drpCommports.ListIndex = 0
        End If
    End If
    
End Sub


Private Sub Form_Unload(cancel As Integer)
    SaveSetting "SerialConsole", "dropdown", "drpBaud.ListIndex", drpBaud.ListIndex
End Sub

Private Sub optInput_ActivateNextState(Index As Integer, u_Cancel As Boolean, u_NewState As uOptionBoxConstants)
    If optInput(Index).Value = u_Selected Then
        u_NewState = u_UnSelected
        u_Cancel = True
    End If
End Sub

Private Sub optInput_Changed(Index As Integer, u_NewState As uOptionBoxConstants)
    txtInput_Changed
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
    
    'Debug.Print Me.ActiveControl.Name
    
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

Private Sub txtInput_Changed()
    Dim i As Long, j As Long
    Dim lPlace As Long
    
    
    Dim str As String
    Dim splitStr() As String
    Dim splitLength As Long
    
    Dim partColor As Long
    Dim partType As ParseType
    
    str = txtInput.Text
    splitStr = Split(str, " ")
    
    txtInput.RedrawPause
    
    Dim forceFunction As Long
    forceFunction = -1
    
    For i = optInput.LBound To optInput.UBound
        If optInput(i).Value = u_Selected Then
            forceFunction = i
            Exit For
        End If
    Next i
    
    
    For i = 0 To Len(str) - 1
        txtInput.setCharForeColor i, -1
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
            End If
            
            For j = 0 To splitLength - 1
                txtInput.setCharForeColor j + lPlace, partColor
                txtInput.setCharBold j + lPlace, partColor <> -1
                'txtInput.setCharForeColor j + lPlace, IIf(partColor = -1, vbBlack, vbWhite)
                
            Next j
        End If
        
        lPlace = lPlace + splitLength + 1
    Next i
    
    
    txtInput.RedrawResume
    
    'l = Asc(Mid(firstText, i, 1))
    'tmpStr = tmpStr & IIf(l < 16, "0", "") & Hex(l) & " "
End Sub



Private Sub txtInput_GotFocus()
    txtInput.BorderColor = &H81B543
End Sub

Private Sub txtInput_LostFocus()
    txtInput.BorderColor = &H808080
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

