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
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "确定"
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
         Caption         =   "允许未解锁时执行电源操作"
         BeginProperty Font 
            Name            =   "宋体"
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
txtinfo.Caption = "查看本程序的帮助信息"
Case 1
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "Window Tool是一款可以帮助您设置窗口的工具,可以帮助您自定义个性化的窗口"
Case 2
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "查看关于Window Tool功能的信息"
Case 3
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "Window Tool会自动获取当前活动窗口的信息,包括:" & vbCrLf & "-----窗口句柄(hWnd);窗口标题;窗口所隶属的进程的路径;窗口设备上下文句柄(hDC);窗口类名"
Case 4
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "本按钮会向目标窗口发送一个WM_CLOSE消息以关闭窗口,有时如果它没有响应或出错可以使用"
Case 5
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "本按钮可以复位所有处于'修改过的窗口'列表框中窗口的设置,程序退出时会自动询问是否执行"
Case 6
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "查看Window Tool窗口选项的信息"
Case 7
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "如果不选定这个选项,则您无法和目标窗口交互,直到您重新选项该选项为止"
Case 8
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "若该选项启用,则目标窗口无论何时都保持在其它窗口最前端"
Case 9
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "若该选项启用,则目标窗口可以透明,您可以通过滑动滚动条或单击下方标签输入(范围0-255)"
Case 10
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "若该选项启用,则目标窗口不能进行最大化/最小化或关闭,在'高级'选项中您可以单独设置这三个选项启用与否"
Case 11
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "若该选项启用,则目标窗口不可见"
Case 12
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "若该选项启用,则目标窗口不能调整大小(包括最大化/最小化)"
Case 13
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "若该选项启用,则目标窗口不能移动"
Case 14
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "若该选项启用,则目标窗口的标题栏将被删除(这个选项可能导致部分窗口停止响应,请谨慎使用)"
Case 15
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "查看Window Tool的声明信息"
Case 16
Me.imgImage.Picture = Form1.Icon
txtinfo.Caption = "Window Tool是通过系统函数修改窗口设置的,在某些极端情况下可能导致目标窗口崩溃或应用程序终止(几率很低但是还是偶然发送),如果目标存有重要数据,请做好保存工作." & vbCrLf & vbCrLf & "PC-DOS Workshop出品" & vbCrLf & "免费使用,谢谢选择!"
End Select
End Sub
