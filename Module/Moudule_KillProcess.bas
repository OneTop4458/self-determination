Attribute VB_Name = "Module_KillProcess"
'---------------------------------------------------------------------------------------
' Module    : Module_KillProcess
' Author    : 모름
' Date      : 2019-05-12
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function Process32First Lib "kernel32.dll" ( _
    ByVal hSnapshot As Long, _
    ByRef lppe As PROCESSENTRY32 _
) As Long
Private Declare Function Process32Next Lib "kernel32.dll" ( _
    ByVal hSnapshot As Long, _
    ByRef lppe As PROCESSENTRY32 _
) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" ( _
    ByVal dwFlags As Long, _
    ByVal th32ProcessID As Long _
) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long _
) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" ( _
    ByVal hProcess As Long, _
    ByVal uExitCode As Long _
) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" ( _
    ByVal hObject As Long _
) As Long
Private Const TH32CS_SNAPTHREAD As Long = &H4
Private Const TH32CS_SNAPPROCESS As Long = &H2
Private Const TH32CS_SNAPMODULE As Long = &H8
Private Const TH32CS_SNAPHEAPLIST As Long = &H1
Private Const TH32CS_SNAPALL As Long = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const MAXIMUM_ALLOWED As Long = &H2000000
Private Const MAX_PATH As Long = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public Function GetProcess(ByVal ProcessName As String) As Long
    ' -- Thanks for end_sub (end_sub@naver.com)
    Dim A As Long
    Dim B As PROCESSENTRY32
    Dim C As Long
    Dim d As String
    Dim GetAs As Boolean
    B.dwSize = Len(B)
    A = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0)
    C = Process32First(A, B)
    Do While C <> 0 '--- 더 가져올 프로세스가 있을때만
        d = IIf(InStr(B.szExeFile, vbNullChar), Mid(B.szExeFile, 1, InStr(B.szExeFile, vbNullChar) - 1), B.szExeFile)
        If LCase(d) = LCase(ProcessName) Then
            GetAs = True
            Exit Do
        End If
        C = Process32Next(A, B) '---프로세스 다음 검색
    Loop
    
    If GetAs = True Then
        GetProcess = B.th32ProcessID
    Else
        GetProcess = 0
    End If
    
    CloseHandle A
    'CloseHandle B
    CloseHandle C
End Function

Public Function ProcessKill(ByVal pPID As Long) As Boolean
    Dim oPID As Long
    If pPID = 0 Then Exit Function
    oPID = OpenProcess(MAXIMUM_ALLOWED, 0, pPID)
    If oPID = 0 Then CloseHandle oPID: Exit Function
    TerminateProcess oPID, 0
    CloseHandle oPID
    ProcessKill = True
End Function


