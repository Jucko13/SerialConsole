VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13815
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
   ScaleHeight     =   6960
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1065
      Left            =   7635
      TabIndex        =   2
      Top             =   3075
      Width           =   1560
   End
   Begin SerialConsole.uGraph uGraph1 
      Height          =   1290
      Left            =   2220
      TabIndex        =   1
      Top             =   2985
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   2275
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   780
      Top             =   5145
   End
   Begin SerialConsole.uTextBox txtReceived 
      Height          =   2760
      Left            =   165
      TabIndex        =   0
      Top             =   255
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   4868
      BackgroundColor =   3551534
      BorderColor     =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
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
      RowLineColor    =   8421504
      RowNumberOnEveryLine=   -1  'True
      WordWrap        =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Timer1.Enabled = False
End Sub

Private Sub Form_Load()
    uGraph1.LineColor(0) = vbRed
    uGraph1.LineVisible(0) = True
    
    uGraph1.LineColor(1) = vbGreen
    uGraph1.LineVisible(1) = True
    
    Me.Visible = True
    
    'Do While 1
    Dim i As Long
    For i = 0 To 200
        Timer1_Timer
        DoEvents
    Next i
    'Loop
    
End Sub

Private Sub Timer1_Timer()
    Static i As Double
    
    i = i + 0.01
    
    uGraph1.AddItem 0, Sin(i) * 100 + Rnd * 5 + 100, False
    uGraph1.AddItem 1, Sin(i + 0.5) * 100 + Rnd * 5 + 100, False
    'uGraph1.ScrollToLastItem 0, True
    uGraph1.Redraw
End Sub
