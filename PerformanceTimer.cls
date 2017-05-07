VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerformanceTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PerformanceValue
    pvSecond = 1                's
    pvDeciSecond = 10           'ds
    pvCentiSecond = 100         'cs
    pvMilliSecond = 1000        'ms
    pvMicroSecond = 1000000     'µs
    pvNanoSecond = 1000000000   'ns
End Enum

Private m_CountsPerSecond As Currency
Private m_Start As Currency
Private m_Stop As Currency
Private m_ApiOverhead As Currency

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Sub Class_Initialize()
    'Does the system support a performance counter
    
    If QueryPerformanceFrequency(m_CountsPerSecond) Then
        Dim i As Long, TotalOverhead As Currency
        
        TotalOverhead = 0
        
        'Find out how long it takes the system to call the API function
        For i = 1 To 1000
            QueryPerformanceCounter m_Start
            QueryPerformanceCounter m_Stop
            TotalOverhead = TotalOverhead + m_Stop - m_Start
        Next i
        m_ApiOverhead = TotalOverhead / 1000
    Else
        m_CountsPerSecond = 1
    End If
    m_Start = 0
    m_Stop = 0
End Sub

Public Property Get Supported() As Boolean
    'Does the system support a performance counter
    Supported = QueryPerformanceCounter(0)
End Property

Public Sub StartTimer()
    'Get the start time
    QueryPerformanceCounter m_Start
    m_Stop = 0
End Sub

Public Sub StopTimer()
    'Get the end time
    QueryPerformanceCounter m_Stop
End Sub

Public Property Get TimeElapsed(ByVal ReturnAccuracy As PerformanceValue) As Double
    'Return the time taken
    If m_Start And m_Stop Then TimeElapsed = (m_Stop - m_Start - m_ApiOverhead) / m_CountsPerSecond * ReturnAccuracy
End Property
