Attribute VB_Name = "Module_BSBD"
'---------------------------------------------------------------------------------------
' Module    : Module_BSBD
' Author    : http://cafe.daum.net/_c21_/bbs_search_read?grpid=1EgTQ&fldid=4Yrp&datanum=67
' Date      : 2019-05-12
' Purpose   :
'---------------------------------------------------------------------------------------

'본 모듈은 프로그램 강제 종료시 블루스크린 발생 모듈 입니다
'본 모듈의 원리는 프로그램을 윈도우 중요프로그램으로 지정하여
'비정상적인 경로로 프로그램 종료시 블루스크린을 발생시킵니다
'부디 본소스를 악용하는일이 없도록 당부 부탁드립니다.
'사용 방법은 프로젝트 로드시 ProtectProcess 입력
'프로젝트 언로드시 RestoreProcess 입력 해주시면 되고
'비정상적인 경로로 종료하는 기준은 RestoreProcess 을 해주고 종료했는냐 여부로 판단됩니다

Option Explicit

' ### 특권 활성화 코드
Private Declare Function RtlAdjustPrivilege Lib "ntdll" ( _
    ByVal Privilege As Long, _
    ByVal bEnablePrivilege As Long, _
    ByVal IsThreadPrivilege As Long, _
    ByRef PreviousValue As Long _
) As Long

' ### 임계 프로세스 설정
Private Declare Function RtlSetProcessIsCritical Lib "ntdll" ( _
    ByVal NewValue As Long, _
    ByRef OldValue As Long, _
    ByVal IsWinlogon As Long _
) As Long

' ### 미처리 예외 핸들러 설정
Private Declare Function SetUnhandledExceptionFilter Lib "kernel32.dll" ( _
    ByVal lpTopLevelExceptionFilter As Long _
) As Long

' ### 스레드 관련 API
Private Declare Function CreateThread Lib "kernel32.dll" ( _
    ByRef lpThreadAttributes As Any, _
    ByVal dwStackSize As Long, _
    ByVal lpStartAddress As Long, _
    ByRef lpParameter As Any, _
    ByVal dwCreationFlags As Long, _
    ByRef lpThreadId As Long _
) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long _
) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" ( _
    ByVal hThread As Long, _
    ByRef lpExitCode As Long _
) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" ( _
    ByVal hObject As Long _
) As Long
Private Const INFINITE& = &HFFFFFFFF
Private Const WAIT_OBJECT_0& = 0&
Private Const SeDebugPrivilege& = 20&
Private OldSEH As Long, OldValue As Long, Protected As Boolean

Public Function ProtectProcess() As Boolean
    On Error GoTo Failed

    ' ### 이미 보호되어 있다면 탈출
    If Protected Then ProtectProcess = True: Exit Function

    ' ### IDE인 경우 탈출
    If App.LogMode = 0& Then Exit Function

    ' ### 디버그 특권을 얻는다.
    If RtlAdjustPrivilege(SeDebugPrivilege, 1&, 0&, 0&) >= 0& Then
        ' ### 미처리 예외 핸들러 설정
        ' ### (API 쓰다가 무슨 일이 생길지 모르므로...)
        OldSEH = SetUnhandledExceptionFilter(AddressOf SafeSEH)

        ' ### 임계 프로세스 설정
        If RtlSetProcessIsCritical(1&, OldValue, 0&) >= 0& Then
            Protected = True: ProtectProcess = True
        End If
    End If

Failed:
End Function

Public Sub RestoreProcess()
    On Error GoTo Failed
    ' ### 원상태로 복구
    SetUnhandledExceptionFilter OldSEH
    RtlSetProcessIsCritical OldValue, 0&, 0&
    Protected = False
Failed:
End Sub

Private Function SafeSEH(ByVal pvExceptPointer As Long) As Long
    ' ### 원상태로 복구
    RestoreProcess

    ' ### 정상적인 처리를 위해 이전 예외 처리기 함수를 호출해준다.
    If OldSEH Then
        Dim hThread As Long, retVal As Long
        hThread = CreateThread(ByVal 0&, 0&, OldSEH, ByVal pvExceptPointer, 0&, 0&)
        If WaitForSingleObject(hThread, INFINITE) = WAIT_OBJECT_0 Then
            If GetExitCodeThread(hThread, retVal) Then
                SafeSEH = retVal
            End If
        End If
        CloseHandle hThread
    End If
End Function

