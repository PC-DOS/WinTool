VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   315
      Left            =   6750
      TabIndex        =   2
      Top             =   4230
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4200
      Left            =   3165
      ScaleHeight     =   4140
      ScaleWidth      =   4920
      TabIndex        =   1
      Top             =   0
      Width           =   4980
      Begin VB.Label txtinfo 
         BackStyle       =   0  'Transparent
         Height          =   3525
         Left            =   960
         TabIndex        =   4
         Top             =   585
         Width           =   3870
      End
      Begin VB.Image imgImage 
         Height          =   720
         Left            =   75
         Top             =   105
         Width           =   720
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "����δ����ʱִ�е�Դ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   960
         TabIndex        =   3
         Top             =   105
         Width           =   3870
      End
   End
   Begin VB.ListBox List1 
      Height          =   4200
      ItemData        =   "Form2.frx":0000
      Left            =   0
      List            =   "Form2.frx":0037
      TabIndex        =   0
      Top             =   0
      Width           =   3150
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
On Error Resume Next
Unload Me
Form1.SetFocus
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
Form1.Show
Form1.SetFocus
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
Select Case Form1.mnuEnable.Checked
Case True
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
With Form1.mnuEnable
.Checked = True
.Enabled = False
End With
With Form1.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = False
End With
With Form1.mnuEnable
.Checked = False
.Enabled = True
End With
With Form1.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End Sub
Private Sub Form_Load()
On Error Resume Next
With Me
.Left = Screen.Width / 2 - .Width / 2
.Top = Screen.Height / 2 - .Height / 2
.Icon = LoadPicture()
End With
With Picture1
.BackColor = RGB(255, 255, 255)
End With
With lblTitle
.ForeColor = RGB(0, 0, 255)
.Caption = "    "
End With
With txtinfo
.Appearance = 0
.BorderStyle = 0
.Caption = "    "
End With
With imgImage
.Visible = True
.Picture = LoadPicture()
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
Select Case Form1.mnuEnable.Checked
Case True
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
With Form1.mnuEnable
.Checked = True
.Enabled = False
End With
With Form1.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = False
End With
With Form1.mnuEnable
.Checked = False
.Enabled = True
End With
With Form1.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Form1.Show
Form1.SetFocus
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
Select Case Form1.mnuEnable.Checked
Case True
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
With Form1.mnuEnable
.Checked = True
.Enabled = False
End With
With Form1.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = False
End With
With Form1.mnuEnable
.Checked = False
.Enabled = True
End With
With Form1.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Form1.Show
Form1.SetFocus
End Sub
Private Sub List1_Click()
On Error Resume Next
lblTitle.Caption = Trim(List1.List(List1.ListIndex))
Select Case List1.ListIndex
Case 0
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "�鿴������İ�����Ϣ"
Case 1
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "Window Tool��һ����԰��������ô��ڵĹ���,���԰������Զ�����Ի��Ĵ���"
Case 2
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "�鿴����Window Tool���ܵ���Ϣ"
Case 3
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "Window Tool���Զ���ȡ��ǰ����ڵ���Ϣ,����:" & vbCrLf & "-----���ھ��(hWnd);���ڱ���;�����������Ľ��̵�·��;�����豸�����ľ��(hDC);��������"
Case 4
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "����ť����Ŀ�괰�ڷ���һ��WM_CLOSE��Ϣ�Թرմ���,��ʱ�����û����Ӧ��������ʹ��"
Case 5
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "����ť���Ը�λ���д���'�޸Ĺ��Ĵ���'�б���д��ڵ�����,�����˳�ʱ���Զ�ѯ���Ƿ�ִ��"
Case 6
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "�鿴Window Tool����ѡ�����Ϣ"
Case 7
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "�����ѡ�����ѡ��,�����޷���Ŀ�괰�ڽ���,ֱ��������ѡ���ѡ��Ϊֹ"
Case 8
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "����ѡ������,��Ŀ�괰�����ۺ�ʱ������������������ǰ��"
Case 9
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "����ѡ������,��Ŀ�괰�ڿ���͸��,������ͨ�������������򵥻��·���ǩ����(��Χ0-255)"
Case 10
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "����ѡ������,��Ŀ�괰�ڲ��ܽ������/��С����ر�,��'�߼�'ѡ���������Ե�������������ѡ���������"
Case 11
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "����ѡ������,��Ŀ�괰�ڲ��ɼ�"
Case 12
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "����ѡ������,��Ŀ�괰�ڲ��ܵ�����С(�������/��С��)"
Case 13
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "����ѡ������,��Ŀ�괰�ڲ����ƶ�"
Case 14
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "����ѡ������,��Ŀ�괰�ڵı���������ɾ��(���ѡ����ܵ��²��ִ���ֹͣ��Ӧ,�����ʹ��)"
Case 15
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "�鿴Window Tool��������Ϣ"
Case 16
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "Window Tool��ͨ��ϵͳ�����޸Ĵ������õ�,��ĳЩ��������¿��ܵ���Ŀ�괰�ڱ�����Ӧ�ó�����ֹ(���ʺܵ͵��ǻ���żȻ����),���Ŀ�������Ҫ����,�����ñ��湤��." & vbCrLf & vbCrLf & "PC-DOS Workshop��Ʒ" & vbCrLf & "���ʹ��,ллѡ��!"
End Select
End Sub
