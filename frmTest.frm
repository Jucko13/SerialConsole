VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6630
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
   ScaleHeight     =   4110
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SerialConsole.uTextBox txtReceived 
      Height          =   2760
      Left            =   165
      TabIndex        =   0
      Top             =   255
      Width           =   5220
      _extentx        =   9208
      _extenty        =   4868
      backgroundcolor =   3551534
      bordercolor     =   8421504
      font            =   "frmTest.frx":0000
      forecolor       =   16777215
      linenumbers     =   -1  'True
      linenumberforecolor=   8421504
      linenumberbackground=   2367774
      rowlines        =   -1  'True
      rowlinecolor    =   8421504
      rownumberoneveryline=   -1  'True
      wordwrap        =   -1  'True
      multiline       =   -1  'True
      scrollbars      =   1
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
