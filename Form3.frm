VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Current Windows"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":08CA
   ScaleHeight     =   5370
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      Caption         =   "设置选中窗口(&S)"
      Enabled         =   0   'False
      Height          =   330
      Left            =   6405
      TabIndex        =   22
      Top             =   5010
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Cancel          =   -1  'True
      Caption         =   "退出本界面(&E)"
      Height          =   330
      Left            =   8055
      TabIndex        =   21
      Top             =   5010
      Width           =   1905
   End
   Begin VB.CheckBox Check1 
      Caption         =   "每隔60秒执行一次刷新(&E)"
      Height          =   270
      Left            =   2430
      TabIndex        =   20
      Top             =   5040
      Value           =   1  'Checked
      Width           =   2370
   End
   Begin VB.CommandButton Command1 
      Caption         =   "刷新(&R)"
      Height          =   330
      Left            =   4920
      TabIndex        =   19
      Top             =   5010
      Width           =   1410
   End
   Begin VB.CheckBox Check2 
      Caption         =   "保持在所有窗口之上(&K)"
      Height          =   315
      Left            =   45
      TabIndex        =   18
      Top             =   5025
      Value           =   1  'Checked
      Width           =   2205
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4395
      Top             =   2445
   End
   Begin VB.Frame Frame3 
      Caption         =   "窗口操作选项"
      Height          =   1125
      Left            =   5745
      TabIndex        =   13
      Top             =   3855
      Width           =   4230
      Begin VB.CommandButton Command8 
         Caption         =   "查看窗口信息(&I)..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   2310
         TabIndex        =   23
         Top             =   645
         Width           =   1800
      End
      Begin VB.CommandButton Command5 
         Caption         =   "关闭选定窗口进程(&T)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2220
         TabIndex        =   17
         Top             =   210
         Width           =   1890
      End
      Begin VB.CommandButton Command4 
         Caption         =   "最小化(&N)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1215
         TabIndex        =   16
         Top             =   645
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         Caption         =   "最大化(&M)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   150
         TabIndex        =   15
         Top             =   645
         Width           =   1020
      End
      Begin VB.CommandButton Command2 
         Caption         =   "关闭选定窗口(&C)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   150
         TabIndex        =   14
         Top             =   210
         Width           =   1890
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "信息"
      Height          =   3840
      Left            =   5745
      TabIndex        =   2
      Top             =   0
      Width           =   4230
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标题:"
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   12
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lpszCaption 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   630
         Left            =   75
         TabIndex        =   11
         Top             =   435
         Width           =   4065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "窗口句柄ID:"
         Height          =   180
         Index           =   1
         Left            =   75
         TabIndex        =   10
         Top             =   1095
         Width           =   990
      End
      Begin VB.Label hWinString 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   75
         TabIndex        =   9
         Top             =   1320
         Width           =   4065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "窗口类名信息:"
         Height          =   180
         Index           =   2
         Left            =   75
         TabIndex        =   8
         Top             =   1635
         Width           =   1170
      End
      Begin VB.Label lpszClass 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   75
         TabIndex        =   7
         Top             =   1845
         Width           =   4065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "隶属进程信息:"
         Height          =   180
         Index           =   3
         Left            =   75
         TabIndex        =   6
         Top             =   2145
         Width           =   1170
      End
      Begin VB.Label lpszThread 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   900
         Left            =   75
         TabIndex        =   5
         Top             =   2355
         Width           =   4065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "设备上下文ID:"
         Height          =   180
         Index           =   4
         Left            =   75
         TabIndex        =   4
         Top             =   3285
         Width           =   1170
      End
      Begin VB.Label hDCString 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   75
         TabIndex        =   3
         Top             =   3480
         Width           =   4065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "当前屏幕上的窗口"
      Height          =   4980
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   5715
      Begin VB.ListBox List1 
         Height          =   4740
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   5595
      End
   End
   Begin VB.Menu hid 
      Caption         =   "mnuHidden"
      Visible         =   0   'False
      Begin VB.Menu mnuInfo 
         Caption         =   "信息(&I)..."
      End
      Begin VB.Menu mnuClose 
         Caption         =   "关闭选定窗口(&C)"
      End
      Begin VB.Menu mnuTP 
         Caption         =   "关闭选定窗口隶属的进程(&T)"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMax 
         Caption         =   "最大化(&M)"
      End
      Begin VB.Menu mnuMIN 
         Caption         =   "最小化(&N)"
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "刷新窗口列表(&R)"
      End
   End
   Begin VB.Menu m 
      Caption         =   "1"
      Visible         =   0   'False
      Begin VB.Menu mR 
         Caption         =   "刷新窗口列表(&R)"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const SC_MINIMIZE = &HF020&
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Const SC_ICON = SC_MINIMIZE
Const SC_TASKLIST = &HF130&
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Dim bCodeUse As Boolean
Private Const WS_CAPTION = &HC00000
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Const PROCESS_ALL_ACCESS = &H1F0FFF
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
szExeFile As String * 1024
End Type
Const SC_RESTORE = &HF120&
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Dim lMeWinStyle As Long
Private Const SC_MOVE = &HF010&
Private Const SC_SIZE = &HF000&
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Const WS_EX_APPWINDOW = &H40000
Private Type WINDOWINFORMATION
hWindow As Long
hWindowDC As Long
hThreadProcess As Long
hThreadProcessID As Long
lpszCaption As String
lpszClassName As String
lpszThreadProcessName As String * 1024
lpszThreadProcessPath As String
lpszExe As String
lpszPath As String
End Type
Private Type WINDOWPARAM
bEnabled As Boolean
bHide As Boolean
bTrans As Boolean
bClosable As Boolean
bSizable As Boolean
bMinisizable As Boolean
bTop As Boolean
lpTransValue As Integer
End Type
Dim lpWindow As WINDOWINFORMATION
Dim lpWindowParam() As WINDOWPARAM
Dim lpCur As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Dim lpRtn As Long
Dim hWindow As Long
Dim lpLength As Long
Dim lpArray() As Byte
Dim lpArray2() As Byte
Dim lpBuff As String
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const LWA_COLORKEY = &H1
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private Const WS_SYSMENU = &H80000
Private Const GWL_STYLE = (-16)
Private Const MF_BYCOMMAND = &H0
Private Const SC_CLOSE = &HF060&
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Private Const MF_INSERT = &H0&
Private Const SC_MAXIMIZE = &HF030&
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Type WINDOWINFOBOXDATA
lpszCaption As String
lpszClass As String
lpszThread As String
lpszHandle As String
lpszDC As String
End Type
Dim dwWinInfo As WINDOWINFOBOXDATA
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Sub GetProcessName(ByVal processID As Long, szExeName As String, szPathName As String)
On Error Resume Next
Dim my As PROCESSENTRY32
Dim hProcessHandle As Long
Dim success As Long
Dim l As Long
l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If l Then
my.dwSize = 1060
If (Process32First(l, my)) Then
Do
If my.th32ProcessID = processID Then
CloseHandle l
szExeName = Left$(my.szExeFile, InStr(1, my.szExeFile, Chr$(0)) - 1)
For l = Len(szExeName) To 1 Step -1
If Mid$(szExeName, l, 1) = "\" Then
Exit For
End If
Next l
szPathName = Left$(szExeName, l)
Exit Sub
End If
Loop Until (Process32Next(l, my) < 1)
End If
CloseHandle l
End If
End Sub
Private Sub Check1_Click()
On Error Resume Next
Select Case Check1.Value
Case 0
With Me.Timer1
.Interval = 60000
.Enabled = False
End With
Case 1
With Me.Timer1
.Interval = 60000
.Enabled = True
End With
Case Else
With Me.Timer1
.Interval = 60000
.Enabled = False
End With
End Select
End Sub
Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF5 Then
On Error Resume Next
List1.Clear
List1.Clear
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command5.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
With List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf EnumWindowProc, 0
ElseIf KeyCode = vbKeyEscape Then
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
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
Unload Me
Form1.Show
Form1.SetFocus
Else
Exit Sub
End If
End Sub
Private Sub Check2_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
End Sub
Private Sub Command1_Click()
On Error Resume Next
List1.Clear
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Me.Command5.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
With List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf EnumWindowProc, 0
End Sub
Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF5 Then
On Error Resume Next
List1.Clear
List1.Clear
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command5.Enabled = False
With List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf EnumWindowProc, 0
ElseIf KeyCode = vbKeyEscape Then
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
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
Unload Me
Form1.Show
Form1.SetFocus
Else
Exit Sub
End If
End Sub
Private Sub Command2_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const WM_CLOSE = &H10
SetWindowPos Form3.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Dim ans As Integer
ans = MsgBox("确定关闭这个窗口吗?请注意保存数据.未保存的数据可能丢失!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
PostMessage Me.hWinString.Caption, WM_CLOSE, 0, 0
SendMessage Me.hWinString.Caption, WM_CLOSE, 0, 0
List1.Clear
EnumWindows AddressOf EnumWindowProc, 0
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command5.Enabled = False
Me.Command7.Enabled = False
Command8.Enabled = False
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
SetWindowPos Form3.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Else
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
SetWindowPos Form3.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Exit Sub
End If
End Sub
Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF5 Then
On Error Resume Next
List1.Clear
List1.Clear
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command5.Enabled = False
Me.Command7.Enabled = False
Command8.Enabled = False
With List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf EnumWindowProc, 0
ElseIf KeyCode = vbKeyEscape Then
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
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
Unload Me
Form1.Show
Form1.SetFocus
Else
Exit Sub
End If
End Sub
Private Sub Command3_Click()
On Error Resume Next
Const WM_SYSCOMMAND = &H112
SendMessage Me.hWinString.Caption, WM_SYSCOMMAND, SC_MAXIMIZE, 0
End Sub
Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF5 Then
On Error Resume Next
List1.Clear
List1.Clear
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command5.Enabled = False
Me.Command7.Enabled = False
Command8.Enabled = False
With List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf EnumWindowProc, 0
ElseIf KeyCode = vbKeyEscape Then
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
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
Unload Me
Form1.Show
Form1.SetFocus
Else
Exit Sub
End If
End Sub
Private Sub Command4_Click()
On Error Resume Next
Const WM_SYSCOMMAND = &H112
SendMessage Me.hWinString.Caption, WM_SYSCOMMAND, SC_MINIMIZE, 0
End Sub
Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF5 Then
On Error Resume Next
List1.Clear
List1.Clear
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command7.Enabled = False
Command5.Enabled = False
Me.Command7.Enabled = False
Command8.Enabled = False
With List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf EnumWindowProc, 0
ElseIf KeyCode = vbKeyEscape Then
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
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
Unload Me
Form1.Show
Form1.SetFocus
Else
Exit Sub
End If
End Sub
Private Sub Command5_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Dim ans As Integer
ans = MsgBox("是否关闭此进程?所有未保存数据将丢失", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Const WM_CLOSE = &H10
PostMessage Me.hWinString.Caption, WM_CLOSE, ByVal 0&, ByVal 0&
SendMessage Me.hWinString.Caption, WM_CLOSE, ByVal 0&, 0&
Dim hProcess As Long
hProcess = OpenProcess(PROCESS_ALL_ACCESS, True, lpWindow.hThreadProcessID)
TerminateProcess hProcess, PROCESS_ALL_ACCESS
List1.Clear
EnumWindows AddressOf EnumWindowProc, 0
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command5.Enabled = False
Me.Command7.Enabled = False
Me.lpszCaption.Caption = ""
Command8.Enabled = False
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
SetWindowPos Form3.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Else
SetWindowPos Form3.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
End If
End Sub
Private Sub Command5_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF5 Then
On Error Resume Next
List1.Clear
List1.Clear
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command5.Enabled = False
Me.Command7.Enabled = False
Command8.Enabled = False
With List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf EnumWindowProc, 0
ElseIf KeyCode = vbKeyEscape Then
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
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
Unload Me
Form1.Show
Form1.SetFocus
Else
Exit Sub
End If
End Sub
Private Sub Command6_Click()
On Error Resume Next
Unload Child1
Unload Child2
Unload Me
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
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
Private Sub Command7_Click()
On Error Resume Next
With Child1
.Left = Me.Left + Me.Width + 5
.Top = Me.Top
.Show
.hWinString.Caption = Me.hWinString.Caption
End With
End Sub
Private Sub Command8_Click()
On Error Resume Next
Dim rtn As Long
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const INFO_CAPTION = "窗口标题信息:"
Const INFO_HANDLE = "窗口句柄信息:"
Const INFO_CLASS = "窗口类名信息:"
Const INFO_DC = "窗口设备上下文信息:"
Const INFO_PROCESS = "窗口所隶属进程的信息:"
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
If lpszClass.Caption = "" Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Exit Sub
End If
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
If Me.lpszClass.Caption <> "" Then
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Else
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Exit Sub
End If
End Sub
Private Sub Command8_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF5 Then
On Error Resume Next
List1.Clear
List1.Clear
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command7.Enabled = False
Command5.Enabled = False
Me.Command7.Enabled = False
Command8.Enabled = False
With List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf EnumWindowProc, 0
ElseIf KeyCode = vbKeyEscape Then
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
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
Unload Me
Form1.Show
Form1.SetFocus
Else
Exit Sub
End If
End Sub
Private Sub Form_Activate()
On Error Resume Next
With Form1.MouseHook
.Enabled = False
.Interval = 1000
End With
End Sub
Private Sub Form_Click()
On Error Resume Next
With Form1.MouseHook
.Enabled = False
.Interval = 1000
End With
End Sub
Private Sub Form_DblClick()
On Error Resume Next
With Form1.MouseHook
.Enabled = False
.Interval = 1000
End With
End Sub
Private Sub Form_GotFocus()
On Error Resume Next
With Form1.MouseHook
.Enabled = False
.Interval = 1000
End With
End Sub
Private Sub Form_Initialize()
On Error Resume Next
With Form1.MouseHook
.Enabled = False
.Interval = 1000
End With
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF5 Then
On Error Resume Next
List1.Clear
List1.Clear
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command5.Enabled = False
Me.Command7.Enabled = False
Command8.Enabled = False
With List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf EnumWindowProc, 0
ElseIf KeyCode = vbKeyEscape Then
Unload Child1
Unload Child2
Unload Me
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
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
Unload Me
Form1.Show
Form1.SetFocus
Else
Exit Sub
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Form1.MouseHook
.Enabled = False
.Interval = 1000
End With
With Me.Timer1
.Interval = 60000
.Enabled = True
End With
With Me.List1
.Clear
End With
With Me.Check1
.Enabled = True
.Value = 1
End With
With Check2
.Value = 1
.Enabled = True
End With
On Error Resume Next
Select Case Check1.Value
Case 0
With Me.Timer1
.Interval = 60000
.Enabled = False
End With
Case 1
With Me.Timer1
.Interval = 60000
.Enabled = True
End With
Case Else
With Me.Timer1
.Interval = 60000
.Enabled = False
End With
With Me
.KeyPreview = True
End With
End Select
EnumWindows AddressOf EnumWindowProc, 0
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
End Sub
Private Function GetPassword(hwnd As Long) As String
On Error Resume Next
lpLength = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
If lpLength > 0 Then
ReDim lpArray(lpLength + 1) As Byte
ReDim lpArray2(lpLength - 1) As Byte
CopyMemory lpArray(0), lpLength, 2
SendMessage hwnd, WM_GETTEXT, lpLength + 1, lpArray(0)
CopyMemory lpArray2(0), lpArray(0), lpLength
GetPassword = StrConv(lpArray2, vbUnicode)
Else
GetPassword = ""
End If
End Function
Private Function GetWindowClassName(hwnd As Long) As String
On Error Resume Next
Dim lpszWindowClassName As String * 256
lpszWindowClassName = Space(256)
GetClassName hwnd, lpszWindowClassName, 256
lpszWindowClassName = Trim(lpszWindowClassName)
GetWindowClassName = lpszWindowClassName
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload Child1
Unload Child2
Unload Me
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
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
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Child1
Unload Child2
Unload Me
End Sub
Private Sub hDCString_Click()
On Error Resume Next
Dim rtn As Long
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const INFO_CAPTION = "窗口标题信息:"
Const INFO_HANDLE = "窗口句柄信息:"
Const INFO_CLASS = "窗口类名信息:"
Const INFO_DC = "窗口设备上下文信息:"
Const INFO_PROCESS = "窗口所隶属进程的信息:"
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
If lpszClass.Caption = "" Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Exit Sub
End If
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
If Me.lpszClass.Caption <> "" Then
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Else
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Exit Sub
End If
End Sub
Private Sub hWinString_Click()
On Error Resume Next
Dim rtn As Long
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const INFO_CAPTION = "窗口标题信息:"
Const INFO_HANDLE = "窗口句柄信息:"
Const INFO_CLASS = "窗口类名信息:"
Const INFO_DC = "窗口设备上下文信息:"
Const INFO_PROCESS = "窗口所隶属进程的信息:"
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
If lpszClass.Caption = "" Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Exit Sub
End If
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
If Me.lpszClass.Caption <> "" Then
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Else
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Exit Sub
End If
End Sub
Private Sub List1_Click()
On Error Resume Next
If List1.ListIndex >= 0 Then
With Me.Command2
.Enabled = True
End With
With Me.Command3
.Enabled = True
End With
With Me.Command4
.Enabled = True
End With
With Me.Command5
.Enabled = True
End With
Me.Command7.Enabled = True
Command8.Enabled = True
Me.lpszCaption.Caption = List1.List(List1.ListIndex)
With lpWindow
.hWindow = FindWindow(vbNullString, Me.lpszCaption.Caption)
End With
With lpWindow
.hWindowDC = GetWindowDC(.hWindow)
.lpszCaption = GetPassword(.hWindow)
If .lpszCaption <> Me.lpszCaption.Caption Then
With Me.Check1
.Enabled = True
.Value = 1
End With
End If
.hThreadProcessID = GetWindowThreadProcessId(.hWindow, 0)
.hThreadProcessID = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName Null, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszThreadProcessName, .lpszThreadProcessPath
.hThreadProcessID = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcessID, .lpszThreadProcessName, 256
Me.lpszThread.Caption = .lpszThreadProcessName
.lpszClassName = GetWindowClassName(.hWindow)
Dim lpszClsName As String * 256
GetClassName .hWindow, lpszClsName, 256
.lpszClassName = Trim(lpszClsName)
Me.lpszCaption = .lpszCaption
Me.lpszClass = .lpszClassName
Me.hDCString = .hWindowDC
Me.hWinString = .hWindow
End With
With lpWindow
.hThreadProcessID = GetWindowThreadProcessId(Me.hWinString.Caption, 0)
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszThreadProcessPath, .lpszThreadProcessPath
Me.lpszThread.Caption = Trim(.lpszThreadProcessName)
End With
With lpWindow
GetWindowThreadProcessId Me.hWinString.Caption, .hThreadProcessID
Debug.Print .hThreadProcessID
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszExe, .lpszPath
Me.lpszThread.Caption = Trim(.lpszExe)
Me.lpszThread.Caption = Me.lpszThread.Caption & vbCrLf & "进程所属PID:" & .hThreadProcessID
lpszThread.Caption = Me.lpszThread.Caption & vbCrLf & "进程所属句柄:" & .hThreadProcess
End With
With lpWindow
GetWindowThreadProcessId Me.hWinString.Caption, .hThreadProcessID
Debug.Print .hThreadProcessID
.hThreadProcess = OpenProcess(PROCESS_ALL_ACCESS, False, .hThreadProcessID)
GetModuleFileName .hThreadProcess, .lpszThreadProcessName, 1024
GetProcessName .hThreadProcessID, .lpszExe, .lpszPath
Me.lpszThread.Caption = "进程映像可执行文件名:" & Trim(.lpszExe)
Me.lpszThread.Caption = Me.lpszThread.Caption & vbCrLf & "进程所属PID:" & .hThreadProcessID
lpszThread.Caption = Me.lpszThread.Caption & vbCrLf & "进程所属句柄:" & .hThreadProcess
End With
Else
With Me.Command2
.Enabled = False
End With
With Me.Command3
.Enabled = False
End With
With Me.Command2
.Enabled = False
End With
End If
End Sub
Private Sub List1_DblClick()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const WM_CLOSE = &H10
SetWindowPos Form3.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Dim ans As Integer
ans = MsgBox("确定关闭这个窗口吗?请注意保存数据.未保存的数据可能丢失!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
PostMessage Me.hWinString.Caption, WM_CLOSE, 0, 0
SendMessage Me.hWinString.Caption, WM_CLOSE, 0, 0
List1.Clear
EnumWindows AddressOf EnumWindowProc, 0
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command5.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
SetWindowPos Form3.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Else
SetWindowPos Form3.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Exit Sub
End If
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF5 Then
On Error Resume Next
List1.Clear
List1.Clear
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command5.Enabled = False
Command8.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
With List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf EnumWindowProc, 0
ElseIf KeyCode = vbKeyEscape Then
On Error Resume Next
With Form1.MouseHook
.Interval = 1000
.Enabled = True
End With
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
Unload Me
Form1.Show
Form1.SetFocus
Else
Exit Sub
End If
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 2 Then
If List1.ListIndex >= 0 Then
PopupMenu Me.hid
Exit Sub
Else
PopupMenu Me.m
Exit Sub
End If
Else
Exit Sub
End If
End Sub
Private Sub lpszCaption_Click()
On Error Resume Next
Dim rtn As Long
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const INFO_CAPTION = "窗口标题信息:"
Const INFO_HANDLE = "窗口句柄信息:"
Const INFO_CLASS = "窗口类名信息:"
Const INFO_DC = "窗口设备上下文信息:"
Const INFO_PROCESS = "窗口所隶属进程的信息:"
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
If lpszClass.Caption = "" Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Exit Sub
End If
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
If Me.lpszClass.Caption <> "" Then
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Else
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Exit Sub
End If
End Sub
Private Sub lpszClass_Click()
On Error Resume Next
Dim rtn As Long
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const INFO_CAPTION = "窗口标题信息:"
Const INFO_HANDLE = "窗口句柄信息:"
Const INFO_CLASS = "窗口类名信息:"
Const INFO_DC = "窗口设备上下文信息:"
Const INFO_PROCESS = "窗口所隶属进程的信息:"
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
If lpszClass.Caption = "" Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Exit Sub
End If
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
If Me.lpszClass.Caption <> "" Then
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Else
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Exit Sub
End If
End Sub
Private Sub lpszThread_Click()
On Error Resume Next
Dim rtn As Long
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const INFO_CAPTION = "窗口标题信息:"
Const INFO_HANDLE = "窗口句柄信息:"
Const INFO_CLASS = "窗口类名信息:"
Const INFO_DC = "窗口设备上下文信息:"
Const INFO_PROCESS = "窗口所隶属进程的信息:"
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
If lpszClass.Caption = "" Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Exit Sub
End If
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
If Me.lpszClass.Caption <> "" Then
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Else
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Exit Sub
End If
End Sub
Private Sub mnuClose_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const WM_CLOSE = &H10
SetWindowPos Form3.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Dim ans As Integer
ans = MsgBox("确定关闭这个窗口吗?请注意保存数据.未保存的数据可能丢失!", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
PostMessage Me.hWinString.Caption, WM_CLOSE, 0, 0
SendMessage Me.hWinString.Caption, WM_CLOSE, 0, 0
List1.Clear
EnumWindows AddressOf EnumWindowProc, 0
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command5.Enabled = False
Command8.Enabled = False
Command7.Enabled = False
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
SetWindowPos Form3.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Else
SetWindowPos Form3.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Exit Sub
End If
End Sub
Private Sub mnuInfo_Click()
On Error Resume Next
Dim rtn As Long
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const INFO_CAPTION = "窗口标题信息:"
Const INFO_HANDLE = "窗口句柄信息:"
Const INFO_CLASS = "窗口类名信息:"
Const INFO_DC = "窗口设备上下文信息:"
Const INFO_PROCESS = "窗口所隶属进程的信息:"
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
If lpszClass.Caption = "" Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Exit Sub
End If
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
If Me.lpszClass.Caption <> "" Then
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Else
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Exit Sub
End If
End Sub
Private Sub mnuMax_Click()
On Error Resume Next
Const WM_SYSCOMMAND = &H112
SendMessage Me.hWinString.Caption, WM_SYSCOMMAND, SC_MAXIMIZE, 0
End Sub
Private Sub mnuMin_Click()
On Error Resume Next
Const WM_SYSCOMMAND = &H112
SendMessage Me.hWinString.Caption, WM_SYSCOMMAND, SC_MINIMIZE, 0
End Sub
Private Sub mnuRefresh_Click()
On Error Resume Next
List1.Clear
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Me.Command5.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
With List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf EnumWindowProc, 0
End Sub
Private Sub mnuTP_Click()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Dim ans As Integer
ans = MsgBox("是否关闭此进程?所有未保存数据将丢失", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Const WM_CLOSE = &H10
PostMessage Me.hWinString.Caption, WM_CLOSE, ByVal 0&, ByVal 0&
SendMessage Me.hWinString.Caption, WM_CLOSE, ByVal 0&, 0&
Dim hProcess As Long
hProcess = OpenProcess(PROCESS_ALL_ACCESS, True, lpWindow.hThreadProcessID)
TerminateProcess hProcess, PROCESS_ALL_ACCESS
List1.Clear
EnumWindows AddressOf EnumWindowProc, 0
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command5.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
SetWindowPos Form3.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
Else
SetWindowPos Form3.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Select Case Check2.Value
Case 1
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
Case 0
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 5745
.Width = 10080
End With
End Select
End If
End Sub
Private Sub mR_Click()
On Error Resume Next
List1.Clear
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Me.Command5.Enabled = False
With List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf EnumWindowProc, 0
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
List1.Clear
EnumWindows AddressOf EnumWindowProc, 0
End Sub
