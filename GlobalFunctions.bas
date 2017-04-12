Attribute VB_Name = "GlobalFunctions"
Option Explicit


Public Declare Function GetTickCount Lib "kernel32" () As Long


Sub main()
    uDontDrawDots = True
    frmMain.Show
End Sub
