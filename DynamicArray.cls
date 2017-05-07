VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynamicArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private arr() As Byte
'Private actualLength As Long
'
'
'
'Private Sub Class_Initialize()
'    ReDim arr(0)
'    actualLength = 0
'End Sub
'
'Private Sub grow(howMany As Long)
'    ReDim Preserve arr(0 To UBound(arr) + howMany)
'End Sub
'
'Public Function insert(ByRef data() As Byte, posStart As Long, posEnd As Long)
'    Dim lInsertLength As Long
'    Dim lLengthDifference As Long
'
'    lInsertLength = UBound(sChar)
'
'
'    If pos + howMany > UBound(arr) Then
'        If UBound(arr) * 2 < UBound(arr) + howMany Then
'            grow howMany + UBound(arr)
'        Else
'            grow UBound(arr)
'        End If
'    End If
'
'    actualLength = actualLength + howMany
'
'    CopyMemory CharMap(pos + howMany), arr(m_SelEnd), CursorToEnd * LenB(arr(0))
'End Function
'
'Public Function remove(pos As Long, howMany As Long)
'
'End Function
'
'Public Function bounds() As Long
'    bounds = actualLength
'End Function
'
'
