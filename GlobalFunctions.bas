Attribute VB_Name = "GlobalFunctions"
Option Explicit


Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long

Sub Main()
    Dim commandProcessed As String
    
    uDontDrawDots = True
    
    commandProcessed = Replace(Command, Chr(34), "")
    
    If Len(commandProcessed) > 0 Then
        'MsgBox "'" & commandProcessed & "'"
        SendMessageToInstances commandProcessed
    Else
        frmMain.Show
        'frmTest.Show
    End If
    
    
End Sub




Public Sub wait(ByVal dblMilliseconds As Double)
    Dim dblStart As Double
    Dim dblEnd As Double
    Dim dblTickCount As Double
    
    dblTickCount = GetTickCount()
    dblStart = GetTickCount()
    dblEnd = GetTickCount + dblMilliseconds
    
    Do
    DoEvents
    dblTickCount = GetTickCount()
    Loop Until dblTickCount > dblEnd Or dblTickCount < dblStart
       
    
End Sub

