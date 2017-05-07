VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0024211E&
   ClientHeight    =   7905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16485
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
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1099
   StartUpPosition =   2  'CenterScreen
   Begin SerialConsole.uFrame frmArduinoWindow 
      Height          =   2055
      Left            =   3975
      TabIndex        =   64
      Top             =   1035
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3625
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
         Height          =   960
         Left            =   300
         TabIndex        =   65
         Top             =   510
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   1693
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
         ItemHeight      =   4
         VisibleItems    =   15
      End
   End
   Begin VB.Timer tmrCheckForReconnect 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   14805
      Top             =   1035
   End
   Begin SerialConsole.uCheckBox chkCommOptions 
      Height          =   465
      Index           =   4
      Left            =   14700
      TabIndex        =   58
      ToolTipText     =   "Auto Connect (Edit Arduino platform.txt)"
      Top             =   180
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   820
      BackgroundColor =   2367774
      Border          =   0   'False
      BorderColor     =   2367774
      Caption         =   "AC"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Index           =   1
      Left            =   1530
      TabIndex        =   56
      Top             =   900
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Index           =   0
      Left            =   270
      TabIndex        =   55
      Top             =   930
      Visible         =   0   'False
      Width           =   1230
   End
   Begin SerialConsole.uFrame frmTxtSettings 
      Height          =   435
      Left            =   360
      TabIndex        =   51
      Top             =   4350
      Width           =   4560
      _ExtentX        =   8043
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
      Begin SerialConsole.uCheckBox chkTxtSettings 
         Height          =   165
         Index           =   0
         Left            =   90
         TabIndex        =   53
         ToolTipText     =   "Data Terminal Ready"
         Top             =   195
         Width           =   1185
         _ExtentX        =   2090
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
         TabIndex        =   54
         ToolTipText     =   "Data Terminal Ready"
         Top             =   195
         Width           =   1365
         _ExtentX        =   2408
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
         TabIndex        =   57
         ToolTipText     =   "Data Terminal Ready"
         Top             =   195
         Width           =   915
         _ExtentX        =   1614
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
   Begin SerialConsole.uCheckBox chkCommOptions 
      Height          =   465
      Index           =   3
      Left            =   13770
      TabIndex        =   49
      ToolTipText     =   "Auto Disconnect (Edit Arduino platform.txt)"
      Top             =   180
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   820
      BackgroundColor =   2367774
      Border          =   0   'False
      BorderColor     =   2367774
      Caption         =   "AD"
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
      Left            =   3825
      TabIndex        =   48
      Top             =   2385
      Visible         =   0   'False
      Width           =   480
   End
   Begin SerialConsole.uCheckBox chkCommOptions 
      Height          =   465
      Index           =   2
      Left            =   12975
      TabIndex        =   47
      ToolTipText     =   "Clear On Connect"
      Top             =   180
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   820
      BackgroundColor =   2367774
      Border          =   0   'False
      BorderColor     =   2367774
      Caption         =   "COC"
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
   Begin SerialConsole.uFrame frmSearch 
      Height          =   645
      Left            =   3045
      TabIndex        =   44
      Top             =   1515
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
         Left            =   3105
         TabIndex        =   50
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
      Interval        =   100
      Left            =   3900
      Top             =   3975
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
      Left            =   90
      TabIndex        =   10
      Top             =   5670
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   1693
      BackgroundColor =   2367774
      BorderColor     =   14737632
      ForeColor       =   16777215
      Caption         =   "Send Information"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
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
         Left            =   8400
         TabIndex        =   15
         Top             =   150
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
         BackgroundColor =   2367774
         Border          =   0   'False
         BorderColor     =   2367774
         Caption         =   "COS"
         CaptionOffsetLeft=   5
         CheckBackgroundColor=   2367774
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
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H0024211E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   9195
      MousePointer    =   9  'Size W E
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   825
      Width           =   150
   End
   Begin SerialConsole.uTextBox txtReceived 
      Height          =   2760
      Left            =   225
      TabIndex        =   0
      Top             =   1395
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   4868
      BackgroundColor =   3551534
      BorderColor     =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal_Ctrl+Hex"
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
      Left            =   6315
      TabIndex        =   5
      Top             =   6885
      Width           =   7650
      _ExtentX        =   12409
      _ExtentY        =   741
      BorderColor     =   2367774
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
      Left            =   0
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   982
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6675
      Width           =   14730
      Begin SerialConsole.uFrame frmComStats 
         Height          =   750
         Left            =   3255
         TabIndex        =   59
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
            TabIndex        =   63
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
            TabIndex        =   62
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
            TabIndex        =   61
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
            TabIndex        =   60
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
      Left            =   6120
      Top             =   3090
   End
   Begin SerialConsole.uCheckBox chkCommOptions 
      Height          =   465
      Index           =   0
      Left            =   10395
      TabIndex        =   4
      ToolTipText     =   "Data Terminal Ready"
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
   Begin SerialConsole.uButton cmdConnect 
      Height          =   465
      Left            =   8055
      TabIndex        =   3
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
   Begin SerialConsole.uDropDown drpCommports 
      Height          =   465
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   6300
      _ExtentX        =   11113
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
      Left            =   5145
      Top             =   4095
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      InBufferSize    =   10
      InputLen        =   1
      OutBufferSize   =   1
      ParityReplace   =   0
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin SerialConsole.uDropDown drpBaud 
      Height          =   465
      Left            =   6615
      TabIndex        =   2
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
   Begin SerialConsole.uCheckBox chkCommOptions 
      Height          =   465
      Index           =   1
      Left            =   11550
      TabIndex        =   6
      ToolTipText     =   "Request To Send"
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   14730
   End
   Begin SerialConsole.uFrame frmZebroControls 
      Height          =   5625
      Left            =   7305
      TabIndex        =   21
      Top             =   900
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   9922
      BackgroundColor =   3551534
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
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   503
         BackgroundColor =   3551534
         Border          =   0   'False
         Caption         =   "Refresh connected Zebros"
         CheckBackgroundColor=   3551534
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
      Left            =   5970
      TabIndex        =   52
      Top             =   4890
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
   Begin VB.Label lblCursorStats 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0024211E&
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
      Left            =   3285
      TabIndex        =   43
      Top             =   5040
      Width           =   1005
   End
   Begin VB.Label lblCursorStats 
      Alignment       =   2  'Center
      BackColor       =   &H0024211E&
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
      Left            =   2130
      TabIndex        =   42
      Top             =   5040
      Width           =   1005
   End
   Begin VB.Label lblCursorStats 
      Alignment       =   2  'Center
      BackColor       =   &H0024211E&
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
      Left            =   1095
      TabIndex        =   41
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label lblCursorStats 
      BackColor       =   &H0024211E&
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
      Left            =   180
      TabIndex        =   40
      Top             =   5040
      Width           =   855
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

Dim picoSendCommand(0 To 10) As Boolean
Dim picoConnected(0 To 10) As Boolean

Dim lastMessageBytes() As Byte

Dim receiveBuffer As String
Dim ledCommand As Long

Dim errorMessages() As String
Const errorMessagesConst = ",M_ERROR,M_ERROR_NOT_CONNECTED,M_ERROR_BUFFER_OVERFLOW,M_ERROR_BUFFER_EMPTY,M_ERROR_UNKNOWN_COMMAND"

Dim bitrateInbound As Long
Dim bitrateOutbound As Long

Dim searchFor() As Byte

Dim ConsoleColors As Variant

Dim arduinoListHeaders() As String

Private WithEvents tmrCheckBitRate As SelfTimer
Attribute tmrCheckBitRate.VB_VarHelpID = -1


Private Sub chkCommOptions_Changed(Index As Integer, u_NewState As uCheckboxConstants)
    Dim newState As Boolean
    
    newState = (u_NewState = u_Checked)
    
    Select Case Index
    
        Case 0
            comm.DTREnable = newState
            If newState And chkCommOptions(2).Value = u_Checked Then
                txtReceived.Clear
            End If
            
        Case 1
            comm.RTSEnable = newState
    End Select

    SaveSetting "SerialConsole", "checkboxes", "chkCommOptions(" & Index & ").Value", u_NewState

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

Private Sub chkRefreshZebro_Changed(u_NewState As uCheckboxConstants)
    setCheckColors chkRefreshZebro, u_NewState = u_Checked
    
    tmrGetConnected.Enabled = (u_NewState = u_Checked)
End Sub

Private Sub chkSend_Changed(Index As Integer, u_NewState As uCheckboxConstants)
    setCheckColors chkSend(Index), u_NewState = u_Checked
    
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
                txtReceived.Redraw
            Else
                txtReceived.Font.Name = "CompendiumArcana Ctrl Char Hex"
                txtReceived.Redraw
            End If

    End Select
    
    SaveSetting "SerialConsole", "checkboxes", "chkTxtSettings(" & Index & ").Value", u_NewState
    
    setCheckColors chkTxtSettings(Index), newState
End Sub

Private Sub cmdConnect_Click(Button As Integer, x As Single, y As Single)
    On Error GoTo notWorking
    
    If comm.PortOpen Then
        comm.PortOpen = False
        cmdConnect.Caption = "Connect"
        cmdConnect.BackgroundColor = &H4747F0
        
    Else
        comm.DTREnable = False
        
        cmdConnect.BackgroundColor = &H81B543
        cmdConnect.Caption = "Disconnect"
        comm.PortOpen = True
        comm.DTREnable = (chkCommOptions(0).Value = u_Checked)
        wait 10
        receiveBufferForShowLength = 0
        txtReceived.Clear
        
        If chkCommOptions(0).Value = u_Checked And chkCommOptions(2).Value = u_Checked Then txtReceived.Clear
            
        setStatus "Connected!"
        
        tmrShowBuffer.Enabled = True
    End If
    
Exit Sub
notWorking:
    cmdConnect.BackgroundColor = &H4747F0
    cmdConnect.Caption = "Connect"
    
    setStatus Err.Description, True, Err.Number
    
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

Private Sub cmdSearch_Click(Button As Integer, x As Single, y As Single)
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
    
    parseInputToBytes = totalBytes
    
    'Dim tmpStr As String
    
    'For i = 0 To totalBytesLength - 1
    '    tmpStr = tmpStr & totalBytes(i) & " "
    'Next i
    
    'Dim tmpOutput As Variant

    
    'MsgBox tmpStr
End Function

Private Sub cmdSearchClose_Click(Button As Integer, x As Single, y As Single)
    frmSearch.Visible = False
    Erase searchFor
    txtReceived.ClearMarkup
    txtReceived.Redraw
    Form_Resize
    txtReceived.SetFocus
    'fillReceivedTextColors 0
End Sub

Private Sub cmdSend_Click(Button As Integer, x As Single, y As Single)
    Dim tmpOutput As Variant
    Dim bytes() As Byte
    
    If comm.PortOpen = False Then Exit Sub
    
    bytes = parseInputToBytes(txtOutput)
    
    If (Not (Not bytes)) <> 0 Then
        tmpOutput = bytes
        commOut tmpOutput
    End If
    
    If chkSend(0).Value = u_Checked Then
        txtOutput.Clear
    End If
    
    'txtOutput.SetFocus
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
    Static i As Long
    Dim RL As Long 'received length
    
    i = i + 1
    
    Select Case comm.CommEvent
    
        Case comEvReceive   'comEvReceive event occured
            Dim tmpReceived As String
            
            'bitrateInbound = bitrateInbound + comm.InBufferCount
            
            
            tmpReceived = comm.Input
            RL = Len(tmpReceived)
            
            If receiveBufferForShowLength + RL > 4000 Then
                tmrShowBuffer_Timer
            End If
            
            Mid$(receiveBufferForShow, receiveBufferForShowLength + 1, RL) = tmpReceived
            receiveBufferForShowLength = receiveBufferForShowLength + RL
            
            'Debug.Print tmpReceived
            
            '########################################################################################################################################################################
            'Replace this concatenation of bullshit with a more memory friendly solution, like a buffered string of 5000 chars that will stay 5000 chars and will never be moved
            'this way there are no more intense memory usages
            '########################################################################################################################################################################
            
            'Clipboard.Clear
            'Clipboard.SetText receiveBuffer
            
            bitrateInbound = bitrateInbound + Len(tmpReceived)
            
            'If chkRefreshZebro.Value = u_Checked Then receiveBuffer = receiveBuffer & tmpReceived
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
            txtStatus.Text = "Some ComPort Error occurred"
            
            
        Case Else
            Debug.Print "whatthefuck"
            ' What happened? It wasn't the arrival of data - and it wasn't
            ' an error. See the ' CommEvent property for a full listing
            ' of all the events and errors.
   End Select
   
   
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
    
    Select Case Index
        Case 0
            txtReceived.AddCharAtCursor Chr(27) & "[45m" & Chr(27) & "[30mH"
            
        Case 1
            txtReceived.AddCharAtCursor Chr(27) & "[47m" & Chr(27) & "[31mK"
            
    End Select
    
    txtReceived.Redraw
End Sub

Private Sub drpBaud_ItemChange(ItemIndex As Long)
    comm.Settings = drpBaud.List(ItemIndex) & ",n,8,1"
End Sub

Private Sub drpCommports_ItemChange(ItemIndex As Long)
    On Error GoTo notWorking
    
    comm.commPort = Replace(serialDevices.commPort(ItemIndex), "COM", "")
    SaveSetting "SerialConsole", "dropdown", "selectedCommPort", serialDevices.commPort(ItemIndex)
     
    Me.Caption = serialDevices.commPort(ItemIndex) & " - SerialConsole - V1.0 by Ricardo de Roode"
     
    Exit Sub
notWorking:
    setStatus Err.Description, True, Err.Number
    
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
    End If
    
    fillCommportList
End Sub

Private Sub Form_Click()
    picFocus.SetFocus
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    Set serialDevices = New CommPortList
    Set Timer = New PerformanceTimer
    Set inputFilter = New InputHandler
    
    Set tmrCheckBitRate = New SelfTimer
    tmrCheckBitRate.Interval = 1000
    tmrCheckBitRate.Enabled = True
    
    
    'testTxt
    
    Erase searchFor
    
    fillCommportList True
    
    fillBaudList
    
    fillZebroButtons
    
    fillLedColors
    
    errorMessages = Split(errorMessagesConst, ",")
    
    comm.OutBufferSize = 5
    
    'On Error Resume Next
    drpBaud.ListIndex = GetSetting("SerialConsole", "dropdown", "drpBaud.ListIndex", 0)
    
    
    
    For i = 0 To chkCommOptions.UBound
        chkCommOptions(i).Value = GetSetting("SerialConsole", "checkboxes", "chkCommOptions(" & i & ").Value", u_unChecked)
    Next i
    For i = 0 To chkTxtSettings.UBound
        chkTxtSettings(i).Value = GetSetting("SerialConsole", "checkboxes", "chkTxtSettings(" & i & ").Value", u_unChecked)
    Next i
    
    
    ledCommand = -1
    
    
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
    
    graphDataInOut.Refresh
    graphDataInOut.AddItem 0, 0, False
    graphDataInOut.AddItem 1, 0, True
    
    dragSplitPercentage = 0.41
    
    
    
'    Dim str As String
'    For i = 0 To 255
'        str = str & Chr(i)
'    Next i
'
'    txtReceived.Text = str
'    txtReceived.setCharBorderColor 65, vbRed
'    txtReceived.setCharForeColor 65, vbGreen
    
    ConsoleColors = Array(vbBlack, vbRed, vbGreen, vbYellow, vbBlue, vbMagenta, vbCyan, vbWhite)
    
    
    'Dim str As String
    
    'str = "dit is text "
    'str = str & Chr(&H1B) & "[" & (6 + 30) & "m"
    'str = str & "dit is gekleurd"
    
    'str = parseAndAddText(txtReceived, str)
    
    
'    Dim i As Long
'    Dim j As Long
'
'    Dim str1(0 To 10) As String
'    Dim str2 As String
'
'    For j = 0 To 10
'        For i = 0 To 10
'            str1(j) = str1(j) & Chr(j)
'        Next i
'
'        txtReceived.AddCharAtCursor str1(j) & vbCrLf
'    Next j
    
    fillArduinoListTestData
    
    'txtReceived.Text = txtReceived.FileToString("F:\Github\SerialConsole\changelog.txt")
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
    

    optInput(0).Left = nominalOffsetX
    optInput(0).Top = nominalOffsetX
    For i = 1 To optInput.UBound
        optInput(i).Top = nominalOffsetX
        optInput(i).Left = optInput(i - 1).Left + optInput(i - 1).Width + nominalOffsetX
    Next i
    
    frmOutput.Height = (txtOutput.Top + txtOutput.Height + nominalOffsetX) / Screen.TwipsPerPixelX
    frmOutput.Left = 12
    frmOutput.Width = Me.ScaleWidth - frmOutput.Left * 2
    frmOutput.Top = picToolbar(1).Top - frmOutput.Height - 5
    
    chkSend(0).Left = frmOutput.ScaleWidth - chkSend(0).Width - nominalOffsetX
    cmdSend.Left = frmOutput.ScaleWidth - cmdSend.Width - nominalOffsetX
    
    txtOutput.Left = nominalOffsetX
    txtOutput.Width = cmdSend.Left - txtOutput.Left - nominalOffsetX
    
    picToolbar(2).Left = 0
    picToolbar(2).Width = Me.ScaleWidth
    picToolbar(2).Top = frmOutput.Top - picToolbar(2).Height
    
    
    
    picSplit.Left = Me.ScaleWidth * dragSplitPercentage
    picSplit.Top = picToolbar(0).Top + picToolbar(0).Height
    picSplit.Height = picToolbar(2).Top - picSplit.Top
    
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
        
    lblCursorStats(0).Top = picToolbar(2).Top - lblCursorStats(0).Height
    lblCursorStats(0).Left = 12
    lblCursorStats(0).Width = Fix(firstFrame / 4)
    
    lblCursorStats(1).Top = lblCursorStats(0).Top
    lblCursorStats(1).Left = lblCursorStats(0).Left + lblCursorStats(0).Width
    lblCursorStats(1).Width = Fix(firstFrame / 5)
    
    lblCursorStats(2).Top = lblCursorStats(0).Top
    lblCursorStats(2).Left = lblCursorStats(1).Left + lblCursorStats(1).Width
    lblCursorStats(2).Width = lblCursorStats(1).Width
    
    lblCursorStats(3).Top = lblCursorStats(0).Top
    lblCursorStats(3).Width = firstFrame - (lblCursorStats(2).Width + lblCursorStats(2).Left)
    lblCursorStats(3).Left = lblCursorStats(2).Width + lblCursorStats(2).Left
    
    
    frmTxtSettings.Top = lblCursorStats(0).Top - frmTxtSettings.Height - 12
    frmTxtSettings.Left = 12
    frmTxtSettings.Width = firstFrame
    
    'chkTxtSettings(0).Left = 0
    
    'For i = 1 To chkTxtSettings.UBound
    '    chkTxtSettings(i).Left = chkTxtSettings(i - 1).Left + chkTxtSettings(i).Width
    'Next i
    
    
    txtReceived.Left = 12
    txtReceived.Width = firstFrame
    txtReceived.Height = frmTxtSettings.Top - txtReceived.Top - 6
    
    
    
    frmZebroControls.Left = picSplit.Left + picSplit.Width
    frmZebroControls.Width = Me.ScaleWidth - frmZebroControls.Left - 12
    frmZebroControls.Top = picToolbar(0).Top + picToolbar(0).Height
    frmZebroControls.Height = picToolbar(2).Top - frmZebroControls.Top
    
    frmArduinoWindow.Left = frmZebroControls.Left
    frmArduinoWindow.Top = frmZebroControls.Top
    frmArduinoWindow.Width = frmZebroControls.Width
    frmArduinoWindow.Height = frmZebroControls.Height
    
    lstArduino.Left = smallOffsetX
    lstArduino.Top = smallOffsetX * 2
    lstArduino.Width = frmArduinoWindow.ScaleWidth - 2 * smallOffsetX
    lstArduino.Height = frmArduinoWindow.ScaleHeight - 3 * smallOffsetX
    
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
    
    frmComStats.Left = graphDataInOut.Left + graphDataInOut.Width + 12
    
    txtStatus.Top = Me.ScaleHeight - txtStatus.Height - 12
    txtStatus.Left = frmComStats.Left + frmComStats.Width + 12
    txtStatus.Width = Me.ScaleWidth - txtStatus.Left - 12

    
    
    
    'upper toolbar
    
    
    chkCommOptions(chkCommOptions.UBound).Left = Me.ScaleWidth - chkCommOptions(chkCommOptions.UBound).Width - 12
    For i = chkCommOptions.UBound - 1 To 0 Step -1
        chkCommOptions(i).Left = chkCommOptions(i + 1).Left - chkCommOptions(i).Width - 12
    Next i
    
    cmdConnect.Left = chkCommOptions(0).Left - cmdConnect.Width - 12
    drpBaud.Left = cmdConnect.Left - drpBaud.Width - 12
    drpCommports.Left = 12
    drpCommports.Width = drpBaud.Left - 12 - drpCommports.Left

    


End Sub


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
    txtOutput_Changed
    txtSearch_Changed
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


Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    dragSplitStartX = x
    dragSplit = True
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static isResizing As Boolean
    
    If isResizing Then Exit Sub
    isResizing = True
    
    If dragSplit = True Then
        Dim newLeft As Double
        
        newLeft = picSplit.Left - (dragSplitStartX - x)
        
        newLeft = 1 / Me.ScaleWidth * newLeft
        If newLeft < 0.1 Then newLeft = 0.1
        If newLeft > 0.9 Then newLeft = 0.9
        
        dragSplitPercentage = newLeft
        
        Form_Resize
        DoEvents
    End If
    
    isResizing = False
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    dragSplit = False
End Sub

Private Sub tmrCheckBitRate_Timer(ByVal Seconds As Currency)
    'txtStatus.Text = comm.CommEvent
    
    'Debug.Print Me.ActiveControl.Name
    
    graphDataInOut.AddItem 0, CDbl(bitrateInbound), False
    graphDataInOut.AddItem 1, CDbl(bitrateOutbound), False
    
    bitrateInbound = 0
    bitrateOutbound = 0
    
    graphDataInOut.ScrollToLastItem 0, True
    graphDataInOut.Refresh
    
    'Debug.Print Seconds
End Sub

Private Sub tmrCheckForReconnect_Timer()
    If comm.PortOpen = False Then
        cmdConnect_Click 0, 0, 0
    Else
        tmrCheckForReconnect.Enabled = False
    End If
    
End Sub

Private Sub txtDataExchange_Change()
    If txtDataExchange.Text = "" Then Exit Sub
    
    Dim strSplit() As String
    
    strSplit = Split(txtDataExchange.Text, " ")
    
    txtDataExchange.Text = ""
    
    If chkCommOptions(3).Value <> u_Checked Then Exit Sub
    
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
                    If chkCommOptions(4).Value = u_Checked Then
                        tmrCheckForReconnect.Enabled = False
                        tmrCheckForReconnect.Enabled = True
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
    comm.Output = bytes
End Sub

Private Sub tmrShowBuffer_Timer()
    'receiveBufferForShow = "lolabcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ lolabcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    If receiveBufferForShowLength = 0 Then Exit Sub
    'Clipboard.Clear
    'Clipboard.SetText receiveBuffer
    'printBuffer
    
    txtReceived.RedrawPause
    txtReceived.SelStart = txtReceived.TextLength
    
    'If chkTxtSettings(1).Value = u_Checked Then
    '    receiveBufferForShow = parseAndAddText(txtReceived, receiveBufferForShow)
    'Else
        txtReceived.AddCharAtCursor Left$(receiveBufferForShow, receiveBufferForShowLength), True
        
    'End If
    
    If chkTxtSettings(0).Value = u_Checked Then txtReceived.ScrollToEnd
    
    fillReceivedTextColors txtReceived.TextLength - Len(receiveBufferForShow)
    
    txtReceived.RedrawResume
    
    
    receiveBufferForShowLength = 0
    
    
    If comm.PortOpen = False Then
        tmrShowBuffer.Enabled = False
    End If
End Sub

Sub parseInputColors(uTxt As uTextBox)
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
        uTxt.setCharBold i, False
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
                uTxt.setCharForeColor j + lPlace, partColor
                uTxt.setCharBold j + lPlace, partColor <> -1
                'txtInput.setCharForeColor j + lPlace, IIf(partColor = -1, vbBlack, vbWhite)
                
            Next j
        End If
        
        lPlace = lPlace + splitLength + 1
    Next i
    
    
    uTxt.RedrawResume
End Sub


Private Sub txtOutput_Changed()
    txtOutput_GotFocus
    
    parseInputColors txtOutput
    
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
    
    If startChar < 0 Then Exit Sub

    s = txtReceived.RawText
    'txtReceived.RedrawPause
    
    If (Not (Not searchFor)) = 0 Then
        Exit Sub
    End If
    
    
    For i = startChar To UBound(s)
        txtReceived.setCharBackColor i, -1
        
        For j = 0 To UBound(searchFor)
            If s(i) = searchFor(j) Then
                txtReceived.setCharBorderColor i, vbBlue
                Exit For
            End If
        Next j
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
    lblCursorStats(3).Caption = "val: " & IIf(charIndex = txtReceived.TextLength, "-- (---)", charVal & "(0x" & Hex(charVal) & ")")
End Sub








Private Sub txtSearch_Changed()
    txtSearch_GotFocus
    parseInputColors txtSearch
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

