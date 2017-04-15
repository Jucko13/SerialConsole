VERSION 5.00
Begin VB.UserControl GraphChart 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13155
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   877
   Begin VB.Timer tmrTick 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   2415
   End
   Begin VB.PictureBox picChart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00008000&
      FillStyle       =   7  'Diagonal Cross
      ForeColor       =   &H00FF0000&
      Height          =   3195
      Left            =   780
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   808
      TabIndex        =   0
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "GraphChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim LastPointY           As Double
Dim LastPointX           As Double
Dim TotalValue           As Double
Dim GraphScaleY          As Double    ' = 10000
Dim GraphScaleX          As Double    ' = 3
Dim xLineEvery           As Double    ' = 10
Dim yLineEvery           As Double    ' = GraphScaleY / 4
Dim LineThickness        As Long    ' = 2
Dim GridThickness        As Long    ' = 1
Dim GraphLine_Color      As Long    ' = vbGreen
Dim GridLine_Color       As Long    ' = &H8000&
Dim GemiddeldeLine_Color As Long    ' = vbBlue
Dim picChartHeight       As Long
Dim HighestValue         As Double
Dim AllPoints()          As Double

'Private Declare Function ExtFloodFill _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal x As Long, _
                             ByVal y As Long, _
                             ByVal colorCode As Long, _
                             ByVal fillType As Long) As Long

Const FLOODFILLBORDER = 0
Const FLOODFILLSURFACE = 1

Private Const unitNamesConst As String = "mS,S,M,U,D"
Private Const unitDevidersConst As String = "1000,60,60,24"
Private unitNames() As String
Private unitDeviders() As String

Dim pTimer As PerformanceTimer


Sub FillIt()

    If UBound(AllPoints) > 0 Then
        'ExtFloodFill picChart.hdc, picChart.Width - 1, picChartHeight, vbGreen, FLOODFILLBORDER
        picChart.Picture = picChart.Image

    End If

End Sub

Sub tmrTick_Timer()
'
'    Static i As Long
'
'    i = i + 1
'
'    If i > 7200 Then
'        i = 0
'    End If
'
'
'    AddPoint ((Sin(i / 20) * 10000) + 5000) + (Rnd * 5000)

    pTimer.StopTimer

    AddPoint pTimer.TimeElapsed(pvMilliSecond)
    
    pTimer.StartTimer
    
End Sub

Sub DrawGemiddelde()
    Dim tmpValueY  As Double
    Dim Gemiddelde As Double

    Gemiddelde = TotalValue / (UBound(AllPoints) + 1)
    tmpValueY = picChartHeight - (picChartHeight / GraphScaleY * Gemiddelde)
    
    picChart.DrawWidth = LineThickness
    picChart.Line (0, tmpValueY)-(picChart.Width, tmpValueY), GemiddeldeLine_Color

End Sub

Sub DrawBorder()
    picChart.DrawWidth = 1

    picChart.Line (0, 0)-(0, picChart.Height - 1), GridLine_Color
    picChart.Line (0, 0)-(picChart.Width - 1, 0), GridLine_Color
    picChart.Line (0, picChart.Height - 1)-(picChart.Width - 1, picChart.Height - 1), GridLine_Color
    picChart.Line (picChart.Width - 1, 0)-(picChart.Width - 1, picChart.Height), GridLine_Color

End Sub

Sub DrawGrid()
    Dim x            As Double
    Dim y            As Double
    Dim Verschuiving As Double
    Dim tmpValue     As Double

    picChart.DrawWidth = GridThickness

    Verschuiving = (UBound(AllPoints) Mod (xLineEvery)) * GraphScaleX
    UserControl.Cls

    tmpValue = (UBound(AllPoints) Mod (xLineEvery))
    tmpValue = UBound(AllPoints) - tmpValue

    If picChart.Width <= 0 Or picChartHeight <= 0 Then Exit Sub

    For x = picChart.Width To 0 Step -(xLineEvery * GraphScaleX)
        picChart.Line (x - 1 - Verschuiving, 0)-(x - 1 - Verschuiving, picChartHeight), GridLine_Color

        If tmpValue >= 0 Then
            UserControl.CurrentX = x - 1 - Verschuiving - (TextWidth("" & tmpValue) / 2) + picChart.Left
            UserControl.CurrentY = UserControl.ScaleHeight - TextHeight("" & tmpValue)

            If UserControl.CurrentX > picChart.Left Then    '+ (TextWidth("" & tmpValue) / 2) Then
                UserControl.Print "" & (tmpValue)
                UserControl.AutoRedraw = True

            End If

        End If

        tmpValue = tmpValue - xLineEvery
    Next x

    tmpValue = 0

    UserControl.CurrentX = 0
    UserControl.CurrentY = picChartHeight + picChart.Top - (picChartHeight / GraphScaleY * tmpValue) - (TextHeight("H") / 2)
    UserControl.Print " " & Round(tmpValue, 0) & " Ms"
    tmpValue = tmpValue + yLineEvery

    UserControl.CurrentX = 0
    UserControl.CurrentY = picChartHeight + picChart.Top - (picChartHeight / GraphScaleY * tmpValue) - (TextHeight("H") / 2)
    UserControl.Print " " & getShortName(tmpValue)
        
        
    For y = picChartHeight - (picChartHeight / GraphScaleY * yLineEvery) To 0 Step -(picChartHeight / GraphScaleY * yLineEvery)
        picChart.Line (0, y - 1)-(picChart.Width, y - 1), GridLine_Color

        'MsgBox (Val(y) - Val(picChartHeight / GraphScaleY * yLineEvery))

        If y <> 0 And (Val(y) - Val(picChartHeight / GraphScaleY * yLineEvery)) <= 0 Then
            UserControl.CurrentX = 0
            UserControl.CurrentY = picChart.Top - (TextHeight("H") / 2)
            UserControl.Print " " & getShortName(GraphScaleY)

        End If

        tmpValue = tmpValue + yLineEvery

    Next y

End Sub

Sub DrawPoints()
    Dim tmpValueY  As Double
    Dim tmpValueX  As Double
    Dim tmpHighest As Double
    Dim i          As Long

    picChart.Picture = LoadPicture()

    DrawGrid

    picChart.DrawWidth = LineThickness

    If UBound(AllPoints) > 0 Then
        LastPointY = picChartHeight - (picChartHeight / GraphScaleY * AllPoints(UBound(AllPoints) - 1))
    Else
        LastPointY = picChartHeight

    End If

    LastPointX = 0

    For i = UBound(AllPoints) - 1 To 0 Step -1

        If tmpHighest < AllPoints(i) Then
            tmpHighest = AllPoints(i)

        End If

        tmpValueY = picChartHeight - (picChartHeight / GraphScaleY * AllPoints(i))

        tmpValueX = picChart.Width - LastPointX

        picChart.Line (tmpValueX, LastPointY)-(tmpValueX - GraphScaleX, tmpValueY), GraphLine_Color

        LastPointX = LastPointX + GraphScaleX
        LastPointY = tmpValueY

        If tmpValueX <= 0 Then
            Exit For

        End If

        'DoEvents
    Next i

    If tmpHighest < HighestValue Then
        HighestValue = tmpHighest
        ResizeMaximum tmpHighest
        Exit Sub

    End If

    LastPointX = GraphScaleX * (UBound(AllPoints))
    LastPointY = picChartHeight - (picChartHeight / GraphScaleY * AllPoints(0))
    tmpValueX = picChart.Width - LastPointX
    tmpValueY = picChartHeight

    If tmpValueX >= 0 Then
        picChart.DrawWidth = LineThickness
        picChart.Line (tmpValueX, LastPointY)-(tmpValueX, tmpValueY), GraphLine_Color
        picChart.DrawWidth = 1
        picChart.Line (tmpValueX + 1, picChartHeight)-(picChart.Width, picChartHeight), vbBlack
    Else
        picChart.DrawWidth = 1
        picChart.Line (1, picChartHeight)-(picChart.Width, picChartHeight), vbBlack

    End If

    FillIt
    DrawGemiddelde
    DrawBorder
    
    'UserControl.Refresh
End Sub

Sub AddPoint(newValue As Variant, Optional lRefresh As Boolean = True)
    Dim tmpDouble As Double

    tmpDouble = Val(newValue)

    If tmpDouble < 1 Then tmpDouble = 1

    'If tmpDouble = 0 Then Exit Sub

    AllPoints(UBound(AllPoints)) = tmpDouble
    ReDim Preserve AllPoints(0 To UBound(AllPoints) + 1) As Double

    If HighestValue < tmpDouble Then
        ResizeMaximum tmpDouble
        HighestValue = tmpDouble

    End If

    'If tmpDouble > GraphScaleY Then
    '    ResizeMaximum tmpDouble
    'End If

    TotalValue = TotalValue + tmpDouble

    If lRefresh = True Then
        Refresh

    End If

End Sub

Sub ResizeMaximum(NewMax As Double)
    GraphScaleY = NewMax
    yLineEvery = GraphScaleY
    DrawPoints

End Sub

Sub Refresh()
    DrawPoints
    DoEvents

End Sub

Function GetPoints() As Double()
    GetPoints = AllPoints

End Function

Function HowManyPoints()
    HowManyPoints = UBound(AllPoints) - 1

End Function

Sub Clear()
    UserControl_Initialize
    AllPoints(0) = 0
    DrawPoints

End Sub

Property Let BackGroundColor(newValue As OLE_COLOR)
    UserControl.BackColor = newValue

End Property

Property Get BackGroundColor() As OLE_COLOR
    BackGroundColor = UserControl.BackColor

End Property

Property Let FontColor(newValue As OLE_COLOR)
    UserControl.ForeColor = newValue

End Property

Property Get FontColor() As OLE_COLOR
    FontColor = UserControl.ForeColor

End Property

Private Sub UserControl_Initialize()
    GraphScaleY = 10000
    GraphScaleX = 3

    xLineEvery = 10
    yLineEvery = GraphScaleY

    LineThickness = 1
    GridThickness = 1

    GraphLine_Color = vbGreen
    GridLine_Color = &H4000&
    GemiddeldeLine_Color = vbBlue
    picChart.FillColor = &H8000&
    
    UserControl.ForeColor = vbWhite
    UserControl.FontName = "Consolas"
    HighestValue = 0
    TotalValue = 0

    ReDim Preserve AllPoints(0) As Double
    
    
    unitNames = Split(unitNamesConst, ",")
    unitDeviders = Split(unitDevidersConst, ",")
    'DrawPoints
    
    Set pTimer = New PerformanceTimer
    
    pTimer.StartTimer
End Sub

Function getShortName(ByVal valueMS As Double) As String
    Dim unitNr As Long
    
    While valueMS > CLng(unitDeviders(unitNr))
        valueMS = valueMS / 1000
        unitNr = unitNr + 1
    Wend
    
    getShortName = Round(valueMS, 1) & unitNames(unitNr)
End Function


Sub StartTest(Inter As Long)
    tmrTick.Interval = Inter
    tmrTick.Enabled = True

End Sub

Sub StopTest()
    tmrTick.Enabled = False

End Sub


Private Sub Usercontrol_Resize()
    On Error Resume Next

    picChart.Left = 40
    picChart.Width = UserControl.ScaleWidth - picChart.Left - 15
    picChart.Top = 10
    picChart.Height = UserControl.ScaleHeight - 25
    picChartHeight = picChart.Height - 1

    DrawPoints

End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        UserControl.BackColor = .ReadProperty("BackgroundColor", &HE18700)
        
    End With
    
    Refresh
    
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackgroundColor", UserControl.BackColor, &HE18700
    End With
End Sub
