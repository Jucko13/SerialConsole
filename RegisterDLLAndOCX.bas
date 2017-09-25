Attribute VB_Name = "RegisterDLLAndOCX"
Option Explicit

Private Declare Function LoadLibraryRegister Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function CreateThreadForRegister Lib "kernel32" Alias "CreateThread" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetProcAddressRegister Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibraryRegister Lib "kernel32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
    
Function RegSvr32(ByVal FileName As String, bUnReg As Boolean) As Boolean

Dim lLib As Long
Dim lProcAddress As Long
Dim lThreadID As Long
Dim lSuccess As Long
Dim lExitCode As Long
Dim lThread As Long
Dim bAns As Boolean
Dim sPurpose As String

sPurpose = IIf(bUnReg, "DllUnregisterServer", "DllRegisterServer")

If Dir(FileName) = "" Then Exit Function

lLib = LoadLibraryRegister(FileName)
If lLib = 0 Then Exit Function

lProcAddress = GetProcAddressRegister(lLib, sPurpose)

If lProcAddress = 0 Then
   FreeLibraryRegister lLib
   Exit Function
Else
   lThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lProcAddress, ByVal 0&, 0&, lThread)
   If lThread Then
        lSuccess = (WaitForSingleObject(lThread, 10000) = 0)
        If Not lSuccess Then
           Call GetExitCodeThread(lThread, lExitCode)
           Call ExitThread(lExitCode)
           bAns = False
           Exit Function
        Else
           bAns = True
        End If
        CloseHandle lThread
        FreeLibraryRegister lLib
   End If
End If
    RegSvr32 = bAns
End Function
