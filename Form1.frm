VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Costom - PC-DOS Workshop"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8010
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   30
      Top             =   1455
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Frame Frame4 
      Caption         =   "高级选项"
      Height          =   1800
      Left            =   4740
      TabIndex        =   26
      Top             =   4860
      Width           =   3225
      Begin VB.CommandButton Command2 
         Caption         =   "结束此窗口隶属的进程(&P)"
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   1410
         Width           =   2955
      End
      Begin VB.CheckBox Check12 
         Caption         =   "删除标题栏(&D)"
         Height          =   300
         Left            =   120
         TabIndex        =   31
         Top             =   1050
         Width           =   1560
      End
      Begin VB.CheckBox Check11 
         Caption         =   "不允许最小化(&O)"
         Height          =   345
         Left            =   120
         TabIndex        =   29
         Top             =   735
         Width           =   1650
      End
      Begin VB.CheckBox Check10 
         Caption         =   "不允许最大化(&T)"
         Height          =   345
         Left            =   120
         TabIndex        =   28
         Top             =   450
         Width           =   1650
      End
      Begin VB.CheckBox Check9 
         Caption         =   "不允许关闭(&N)"
         Height          =   345
         Left            =   120
         TabIndex        =   27
         Top             =   165
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "修改过的窗口(双击还原默认)"
      Height          =   4695
      Left            =   4740
      TabIndex        =   22
      Top             =   120
      Width           =   3240
      Begin VB.ListBox List1 
         Height          =   4380
         Left            =   60
         TabIndex        =   23
         Top             =   225
         Width           =   3135
      End
   End
   Begin VB.Timer MouseHook 
      Interval        =   1000
      Left            =   1755
      Top             =   2610
   End
   Begin VB.Frame Frame2 
      Caption         =   "选项"
      Height          =   3195
      Left            =   30
      TabIndex        =   12
      Top             =   3465
      Width           =   4665
      Begin VB.CommandButton Command3 
         Caption         =   "修改窗口标题(&I)"
         Height          =   315
         Left            =   1785
         TabIndex        =   33
         Top             =   2805
         Width           =   2760
      End
      Begin 工程1.cSysTray cSysTray1 
         Left            =   3915
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   900
         InTray          =   0   'False
         TrayIcon        =   "Form1.frx":030A
         TrayTip         =   "Win Tool - 双击还原窗口"
      End
      Begin VB.CheckBox Check8 
         Caption         =   "该窗口不允许移动(&M)"
         Enabled         =   0   'False
         Height          =   270
         Left            =   150
         TabIndex        =   25
         Top             =   2460
         Width           =   4410
      End
      Begin VB.CommandButton Command1 
         Caption         =   "关闭此窗口(&C)"
         Enabled         =   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   24
         Top             =   2805
         Width           =   1500
      End
      Begin VB.CheckBox Check7 
         Caption         =   "该窗口不允许被调整大小(&S)"
         Height          =   300
         Left            =   150
         TabIndex        =   21
         Top             =   2160
         Width           =   4035
      End
      Begin VB.CheckBox Check6 
         Caption         =   "该窗口不允许被最小化/最大化/关闭(&A)"
         Height          =   300
         Left            =   150
         TabIndex        =   20
         Top             =   1890
         Width           =   4035
      End
      Begin VB.CheckBox Check4 
         Caption         =   "该窗口被隐藏(&H)"
         Height          =   300
         Left            =   150
         TabIndex        =   19
         Top             =   1620
         Width           =   4035
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   10
         Left            =   420
         Max             =   255
         SmallChange     =   5
         TabIndex        =   16
         Top             =   1005
         Width           =   4155
      End
      Begin VB.CheckBox Check3 
         Caption         =   "该窗口透明(&T)"
         Height          =   300
         Left            =   150
         TabIndex        =   15
         Top             =   690
         Width           =   4035
      End
      Begin VB.CheckBox Check2 
         Caption         =   "该窗口始终保持在所有窗口之上(&K)"
         Height          =   300
         Left            =   150
         TabIndex        =   14
         Top             =   435
         Width           =   4035
      End
      Begin VB.CheckBox Check1 
         Caption         =   "该窗口有效并允许用户与之交互(&E)"
         Height          =   300
         Left            =   150
         TabIndex        =   13
         Top             =   180
         Value           =   1  'Checked
         Width           =   4035
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "255"
         Height          =   240
         Left            =   930
         TabIndex        =   18
         Top             =   1305
         Width           =   3645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数值:"
         Height          =   180
         Left            =   435
         TabIndex        =   17
         Top             =   1335
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "活动窗口信息"
      Height          =   2760
      Left            =   15
      TabIndex        =   1
      Top             =   645
      Width           =   4680
      Begin VB.Label hDCString 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1290
         TabIndex        =   11
         Top             =   2430
         Width           =   3330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "设备上下文ID:"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   10
         Top             =   2475
         Width           =   1170
      End
      Begin VB.Label lpszThread 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   885
         Left            =   1290
         TabIndex        =   9
         Top             =   1515
         Width           =   3330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "隶属进程信息:"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   8
         Top             =   1545
         Width           =   1170
      End
      Begin VB.Label lpszClass 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1290
         TabIndex        =   7
         Top             =   1215
         Width           =   3330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "窗口类名信息:"
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   6
         Top             =   1260
         Width           =   1170
      End
      Begin VB.Label hWinString 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1290
         TabIndex        =   5
         Top             =   915
         Width           =   3330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "窗口句柄ID:"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   4
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lpszCaption 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   630
         Left            =   1290
         TabIndex        =   3
         Top             =   255
         Width           =   3330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标题:"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   300
         Width           =   450
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请使用鼠标单击一个窗口使其成为活动窗口,之后您就可以在下面设置它的属性了."
      Height          =   435
      Left            =   825
      TabIndex        =   0
      Top             =   75
      Width           =   3840
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":075C
      Top             =   75
      Width           =   480
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "显示主窗口(&S)"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "复位所有窗口设置(&R)"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "当前窗口信息(&C)..."
      End
      Begin VB.Menu mnuTaskT 
         Caption         =   "Windows 任务管理器(&T)..."
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&E)"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuRestore 
         Caption         =   "复位所有窗口设置(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnubar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMon 
         Caption         =   "监视活动窗口(&A)"
         Begin VB.Menu mnuEnable 
            Caption         =   "启用(&E)"
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuDisable 
            Caption         =   "禁用(&D)"
            Shortcut        =   {F9}
         End
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVPWWCW 
         Caption         =   "查看拥有子窗口的进程(&V)..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuWindView 
         Caption         =   "窗口查看器(&W)..."
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuTasks 
         Caption         =   "Windows 任务管理器(&T)..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewCurWin 
         Caption         =   "查看当前窗口信息(&U)..."
      End
      Begin VB.Menu mnuMini 
         Caption         =   "最小化到托盘(&M)"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "退出(&E)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuTop 
         Caption         =   "总在最前面(&T)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "窗口75%透明功能(&A)"
      End
      Begin VB.Menu mnuRefesh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpForm 
         Caption         =   "帮助(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于WinTool(&A)..."
      End
      Begin VB.Menu mnuBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfoA 
         Caption         =   "活动窗口信息(&I)..."
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Dim lpszCaptionNew As String
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
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOOWNERZORDER = &H200
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
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
Private Sub DisableClose(hwnd As Long, Optional ByVal MDIChild As Boolean)
On Error Resume Next
Exit Sub
Dim hSysMenu As Long
Dim nCnt As Long
Dim cID As Long
hSysMenu = GetSystemMenu(hwnd, False)
If hSysMenu = 0 Then
Exit Sub
End If
nCnt = GetMenuItemCount(hSysMenu)
If MDIChild Then
cID = 3
Else
cID = 1
End If
If nCnt Then
RemoveMenu hSysMenu, nCnt - cID, MF_BYPOSITION Or MF_REMOVE
RemoveMenu hSysMenu, nCnt - cID - 1, MF_BYPOSITION Or MF_REMOVE
DrawMenuBar hwnd
End If
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
Private Sub Check1_Click()
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
On Error GoTo ep
If bCodeUse = True Then
Exit Sub
End If
Dim lpReturn As Long
Select Case Check1.Value
Case 0
lpReturn = EnableWindow(lpWindow.hWindow, False)
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
If lpReturn <> 0 Then
Exit Sub
Else
If 1 = 245 Then
MsgBox "发生错误,无法设置窗口有效性", vbCritical, "Error"
End If
End If
Case 1
lpReturn = EnableWindow(lpWindow.hWindow, True)
If lpReturn <> 0 Then
Exit Sub
Else
If 1 = 245 Then
MsgBox "发生错误,无法设置窗口有效性", vbCritical, "Error"
End If
End If
End Select
Exit Sub
ep:
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
With MouseHook
.Enabled = False
.Interval = 1000
End With
MsgBox Err.Description, vbCritical, "Error"
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
End Sub
Private Sub Check1_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_GotFocus()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_KeyPress(KeyAscii As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_KeyUp(KeyCode As Integer, Shift As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_LostFocus()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
End Sub
Private Sub Check1_Validate(Cancel As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check10_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Dim lpTemp As Long
lpTemp = GetWindowLong(Me.hWinString.Caption, GWL_STYLE)
If bCodeUse = True Then
Exit Sub
End If
Select Case Check10.Value
Case 0
lpTemp = lpTemp Or WS_MAXIMIZEBOX
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If Check9.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
End If
If Check11.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
lpTemp = lpTemp And Not WS_MAXIMIZEBOX
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
End Select
End Sub
Private Sub Check11_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If bCodeUse = True Then
Exit Sub
End If
Dim lpTemp As Long
lpTemp = GetWindowLong(Me.hWinString.Caption, GWL_STYLE)
Select Case Check11.Value
Case 0
lpTemp = lpTemp Or WS_MINIMIZEBOX
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If Check10.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
End If
If Check9.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
lpTemp = lpTemp And Not WS_MINIMIZEBOX
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
End Select
End Sub
Private Sub Check12_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If bCodeUse = True Then
Exit Sub
End If
Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1
Const SWP_FRAMECHANGED = &H20
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOZORDER = &H4
Dim lpStyle As Long
lpStyle = GetWindowLong(Me.hWinString.Caption, GWL_STYLE)
Select Case Check12.Value
Case 1
lpStyle = lpStyle And Not WS_CAPTION
SetWindowLong hWinString.Caption, GWL_STYLE, lpStyle
SetWindowPos hWinString.Caption, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
Select Case Check2.Value
Case 0
SetWindowPos lpWindow.hWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Case 1
SetWindowPos lpWindow.hWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Select
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
Case 0
lpStyle = lpStyle Or WS_CAPTION
SetWindowLong hWinString.Caption, GWL_STYLE, lpStyle
SetWindowPos hWinString.Caption, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
Select Case Check2.Value
Case 0
SetWindowPos lpWindow.hWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Case 1
SetWindowPos lpWindow.hWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Select
End Select
End Sub
Private Sub Check2_Click()
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
On Error GoTo ep
If bCodeUse = True Then
Exit Sub
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim lpReturn As Long
Select Case Check2.Value
Case 1
lpReturn = SetWindowPos(lpWindow.hWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
If lpReturn <> 0 Then
Exit Sub
Else
If 1 = 245 Then
MsgBox "发生错误,无法设置窗口位置", vbCritical, "Error"
End If
End If
Case 0
lpReturn = SetWindowPos(lpWindow.hWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
If lpReturn <> 0 Then
Exit Sub
Else
If 1 = 245 Then
MsgBox "发生错误,无法设置窗口位置", vbCritical, "Error"
End If
End If
End Select
Exit Sub
ep:
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
MsgBox Err.Description, vbCritical, "Error"
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
End Sub
Private Sub Check3_Click()
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
On Error GoTo ep
If List1.ListCount >= 2 Then
For nTmp = 0 To List1.ListIndex
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
If bCodeUse = True Then
Exit Sub
End If
Dim lpReturn As Long
Select Case Check3.Value
Case 0
With Me.HScroll1
.Enabled = False
.Min = 0
.Max = 255
.LargeChange = 10
.SmallChange = 5
End With
With Me.Label9
.Caption = Me.HScroll1.Value
.Enabled = False
.BorderStyle = 1
.BackStyle = 0
.Alignment = 2
End With
With Me.Label8
.Enabled = False
End With
lpReturn = GetWindowLong(lpWindow.hWindow, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong lpWindow.hWindow, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes lpWindow.hWindow, 0, 255, LWA_ALPHA
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Case 1
With Me.HScroll1
.Enabled = True
.Min = 0
.Max = 255
.LargeChange = 10
.SmallChange = 5
End With
With Me.Label9
.Caption = Me.HScroll1.Value
.Enabled = True
.BorderStyle = 1
.BackStyle = 0
.Alignment = 2
End With
With Me.Label8
.Enabled = True
End With
lpReturn = GetWindowLong(lpWindow.hWindow, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong lpWindow.hWindow, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes lpWindow.hWindow, 0, HScroll1.Value, LWA_ALPHA
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Select
Exit Sub
ep:
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
MsgBox Err.Description, vbCritical, "Error"
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
End Sub
Private Sub Check3_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_GotFocus()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_KeyDown(KeyCode As Integer, Shift As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_KeyPress(KeyAscii As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_KeyUp(KeyCode As Integer, Shift As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_LostFocus()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check3_Validate(Cancel As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check4_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If bCodeUse = True Then
Exit Sub
End If
Dim lpReturn As Long
Dim lpTemp As Long
Select Case Check4.Value
Case 0
Const SWP_NOZORDER = &H4
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const HWND_NOTOPMOST = -2
GetSystemMenu hWinString.Caption, True
EnableWindow hWinString.Caption, True
Const SWP_FRAMECHANGED = &H20
Dim lpStyle As Long
lpStyle = GetWindowLong(hWinString.Caption, GWL_STYLE)
lpStyle = lpStyle Or WS_CAPTION
SetWindowLong hWinString.Caption, GWL_STYLE, lpStyle
SetWindowPos hWinString.Caption, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOACTIVATE
lpReturn = GetWindowLong(hWinString.Caption, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong hWinString.Caption, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes hWinString.Caption, 0, 255, LWA_ALPHA
SetWindowPos hWinString.Caption, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
ShowWindow hWinString.Caption, 1
GetSystemMenu hWinString.Caption, True
lpTemp = GetWindowLong(hWinString.Caption, GWL_STYLE)
lpTemp = lpTemp Or WS_MINIMIZEBOX
lpTemp = lpTemp Or WS_MAXIMIZEBOX
lpTemp = lpTemp Or WS_SYSMENU
SetWindowLong hWinString.Caption, GWL_STYLE, lpTemp
ShowWindow hWinString.Caption, 1
SetWindowPos hWinString.Caption, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
SetWindowPos hWinString.Caption, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
Dim lpLong As Long
lpLong = GetWindowLong(Me.hWinString.Caption, GWL_STYLE)
If 1 = 245 Then
lpLong = lpLong Or WS_ICONIC
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpLong
End If
If Check3.Value = 1 Then
GetSystemMenu Me.hWinString.Caption, True
lpReturn = GetWindowLong(lpWindow.hWindow, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong lpWindow.hWindow, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes lpWindow.hWindow, 0, HScroll1.Value, LWA_ALPHA
With Me.Label9
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = True
End With
SetWindowPos hWinString.Caption, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
Else
lpReturn = GetWindowLong(lpWindow.hWindow, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong lpWindow.hWindow, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes lpWindow.hWindow, 0, 255, LWA_ALPHA
With Me.Label9
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = True
End With
End If
If 245 <> 245 Then
Select Case Check9.Value
Case 0
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Check10.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
End If
If Check11.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
End Select
On Error Resume Next
Select Case Check10.Value
Case 0
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Check9.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
End If
If Check11.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
End Select
On Error Resume Next
Select Case Check11.Value
Case 0
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Check10.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
End If
If Check9.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
End Select
End If
If 245 <> 245 Then
GetSystemMenu hWinString.Caption, True
Select Case Check9.Value
Case 0
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Check10.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
End If
If Check11.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
End Select
On Error Resume Next
Select Case Check10.Value
Case 0
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Check9.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
End If
If Check11.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
End Select
On Error Resume Next
Select Case Check11.Value
Case 0
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Check10.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
End If
If Check9.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
End Select
On Error Resume Next
lpTemp = GetWindowLong(Me.hWinString.Caption, GWL_STYLE)
If bCodeUse = True Then
Exit Sub
End If
Select Case Check10.Value
Case 0
lpTemp = lpTemp Or WS_MAXIMIZEBOX
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Check9.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
End If
If Check11.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
lpTemp = lpTemp And Not WS_MAXIMIZEBOX
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
End If
End If
Dim cunt As Long
Dim addvar As Long
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
End If
Next
End Select
On Error Resume Next
If bCodeUse = True Then
Exit Sub
End If
lpTemp = GetWindowLong(Me.hWinString.Caption, GWL_STYLE)
Select Case Check11.Value
Case 0
lpTemp = lpTemp Or WS_MINIMIZEBOX
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Check10.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
End If
If Check9.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
lpTemp = lpTemp And Not WS_MINIMIZEBOX
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
End If
End If
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
End If
Next
End Select
End If
Case 1
If 1 = 245 Then
lpLong = lpLong And Not WS_ICONIC
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpLong
End If
lpReturn = GetWindowLong(lpWindow.hWindow, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong lpWindow.hWindow, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes lpWindow.hWindow, 0, 0, LWA_ALPHA
GetSystemMenu hWinString.Caption, True
EnableWindow hWinString.Caption, True
lpStyle = GetWindowLong(hWinString.Caption, GWL_STYLE)
lpStyle = lpStyle Or WS_CAPTION
SetWindowLong hWinString.Caption, GWL_STYLE, lpStyle
SetWindowPos hWinString.Caption, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOACTIVATE
lpReturn = GetWindowLong(hWinString.Caption, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong hWinString.Caption, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes hWinString.Caption, 0, 255, LWA_ALPHA
SetWindowPos hWinString.Caption, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
ShowWindow hWinString.Caption, 1
GetSystemMenu hWinString.Caption, True
lpTemp = GetWindowLong(hWinString.Caption, GWL_STYLE)
lpTemp = lpTemp Or WS_MINIMIZEBOX
lpTemp = lpTemp Or WS_MAXIMIZEBOX
lpTemp = lpTemp Or WS_SYSMENU
SetWindowLong hWinString.Caption, GWL_STYLE, lpTemp
ShowWindow hWinString.Caption, 1
SetWindowPos hWinString.Caption, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
SetWindowPos hWinString.Caption, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_HIDEWINDOW
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
With Me.Label9
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = True
End With
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
End If
End If
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
End If
Next
End Select
End Sub
Private Sub Check5_Click()
On Error Resume Next
If 1 = 1 Then
Exit Sub
End If
Dim addvar As Integer
Dim cunt As Integer
Select Case Check1.Value
Case 0
lMeWinStyle = GetWindowLong(hWinString.Caption, -16)
Call SetWindowLong(Me.hWinString.Caption, -16, lMeWinStyle)
Call ShowWindow(Me.hWinString.Caption, 1)
Case 1
lMeWinStyle = GetWindowLong(hWinString.Caption, -16)
Call SetWindowLong(Me.hWinString.Caption, -16, lMeWinStyle And &HFFF7FFFF)
Call ShowWindow(Me.hWinString.Caption, 1)
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
Exit Sub
DisableClose Me.hWinString.Caption
Dim hSysMenu As Long
Dim nCnt As Long
Dim cID As Long
hSysMenu = GetSystemMenu(Me.hWinString.Caption, False)
If hSysMenu = 0 Then
Exit Sub
End If
nCnt = GetMenuItemCount(hSysMenu)
If MDIChild Then
cID = 3
Else
cID = 1
End If
If nCnt Then
RemoveMenu hSysMenu, nCnt - cID, MF_BYPOSITION Or MF_REMOVE
RemoveMenu hSysMenu, nCnt - cID - 1, MF_BYPOSITION Or MF_REMOVE
DrawMenuBar hWinString.Caption
End If
Call SetWindowLong(Me.hWinString.Caption, -16, lMeWinStyle And &HFFF7FFFF)
Call ShowWindow(Me.hWinString.Caption, 1)
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
End Select
End Sub
Private Sub Check6_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If bCodeUse = True Then
Exit Sub
End If
Dim lpTemp As Long
lpTemp = GetWindowLong(Me.hWinString.Caption, GWL_STYLE)
Select Case Check6.Value
Case 1
lpTemp = lpTemp And Not WS_MINIMIZEBOX
lpTemp = lpTemp And Not WS_MAXIMIZEBOX
If 1 = 2 Then
lpTemp = lpTemp And Not WS_SYSMENU
End If
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
Check9.Value = 1
Check10.Value = 1
Check11.Value = 1
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
Exit Sub
If 1 = 0 Then
lpTemp = GetWindowLong(Me.hWinString.Caption, GWL_STYLE)
lpTemp = lpTemp And Not WS_MINIMIZEBOX
lpTemp = lpTemp And Not WS_MAXIMIZEBOX
lpTemp = lpTemp And Not WS_SYSMENU
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
End If
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
Case 0
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
Check9.Value = 0
Check10.Value = 0
Check11.Value = 0
lpTemp = lpTemp Or WS_MINIMIZEBOX
lpTemp = lpTemp Or WS_MAXIMIZEBOX
lpTemp = lpTemp Or WS_SYSMENU
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
Case 2
Select Case Check9.Value
Case 0
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If Check10.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
End If
If Check11.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
End Select
On Error Resume Next
Select Case Check10.Value
Case 0
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If Check9.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
End If
If Check11.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
End Select
On Error Resume Next
Select Case Check11.Value
Case 0
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If Check10.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
End If
If Check9.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
End Select
On Error Resume Next
lpTemp = GetWindowLong(Me.hWinString.Caption, GWL_STYLE)
If bCodeUse = True Then
Exit Sub
End If
Select Case Check10.Value
Case 0
lpTemp = lpTemp Or WS_MAXIMIZEBOX
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If Check9.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
End If
If Check11.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
lpTemp = lpTemp And Not WS_MAXIMIZEBOX
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
End If
End If
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
End If
Next
End Select
On Error Resume Next
If bCodeUse = True Then
Exit Sub
End If
lpTemp = GetWindowLong(Me.hWinString.Caption, GWL_STYLE)
Select Case Check11.Value
Case 0
lpTemp = lpTemp Or WS_MINIMIZEBOX
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If Check10.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
End If
If Check9.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
lpTemp = lpTemp And Not WS_MINIMIZEBOX
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpTemp
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
End If
End If
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
End If
Next
End Select
End Select
End Sub
Private Sub Check7_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If bCodeUse = True Then
Exit Sub
End If
Static hSystemSize As Long
If hSystemSize = 0 Then
hSystemSize = GetSystemMenu(hWinString.Caption, 0)
Debug.Print hSystemSize
Else
Debug.Print hSystemSize
End If
Select Case Check7.Value
Case 0
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_INSERT
RemoveMenu hSystemSize, SC_SIZE, MF_INSERT
GetSystemMenu hWinString.Caption, True
Case 1
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_RESTORE, MF_REMOVE
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
End Select
End Sub
Private Sub Check8_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If bCodeUse = True Then
Exit Sub
End If
Select Case Check8.Value
Case 0
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_INSERT
GetSystemMenu hWinString.Caption, True
Case 1
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
End Select
End Sub
Private Sub Check9_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If bCodeUse = True Then
Exit Sub
End If
Select Case Check9.Value
Case 0
GetSystemMenu hWinString.Caption, True
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_INSERT
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_INSERT
GetSystemMenu hWinString.Caption, True
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
If Check10.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MAXIMIZE, MF_REMOVE
End If
If Check11.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MINIMIZE, MF_REMOVE
End If
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
Case 1
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If 1 = 245 Then
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
End If
If Me.Check7.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_SIZE, MF_REMOVE
End If
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_CLOSE, MF_REMOVE
If (Check9.Value = 1) And (Check10.Value = 1) And (Check11.Value = 1) Then
Check6.Value = 1
Else
Check6.Value = 2
End If
If (Check9.Value = 0) And (Check10.Value = 0) And (Check11.Value = 0) Then
Check6.Value = 0
End If
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
End Select
End Sub
Private Sub Command1_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Dim ans As Integer
ans = MsgBox("是否关闭此窗口?所有未保存数据将丢失", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Const WM_CLOSE = &H10
Dim lpListNum As Long
For lpListNum = 0 To Me.List1.ListCount
If List1.List(lpListNum) = Me.hWinString.Caption Then
List1.RemoveItem lpListNum
End If
Next
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
PostMessage Me.hWinString.Caption, WM_CLOSE, ByVal 0&, ByVal 0&
SendMessage Me.hWinString.Caption, WM_CLOSE, ByVal 0&, 0&
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.Check10.Enabled = False
Me.Check11.Enabled = False
Me.Check12.Enabled = False
Me.Check9.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Me.Check1.Enabled = False
Me.Check2.Enabled = False
Me.Check3.Enabled = False
Me.Check4.Enabled = False
Me.Check6.Enabled = False
Check8.Enabled = False
Me.Check7.Enabled = False
Me.HScroll1.Enabled = False
Me.Label9.Enabled = False
Me.Label8.Enabled = False
Me.Command1.Enabled = False
Me.Check10.Enabled = False
Me.Check11.Enabled = False
Me.Check12.Enabled = False
Me.Check9.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command1.Enabled = False
Me.Check1.Enabled = False
Me.Check2.Enabled = False
Me.Check3.Enabled = False
Me.Check4.Enabled = False
Me.Check6.Enabled = False
Me.Check7.Enabled = False
Check8.Enabled = False
Me.hDCString.Caption = ""
Me.hWinString.Caption = ""
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
With Me.HScroll1
.Min = 0
.Max = 255
.SmallChange = 5
.LargeChange = 10
.Enabled = False
.Value = 255
End With
With Me.Label9
.Enabled = False
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = HScroll1.Value
End With
With Me.Label8
.Enabled = False
End With
Check1.Value = 1
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
With Me.Check1
.Enabled = False
.Value = 1
End With
With Me.Check2
.Enabled = False
.Value = 0
End With
With Me.Check3
.Enabled = False
.Value = 0
End With
With Me.HScroll1
.Enabled = False
.Max = 255
.Min = 0
End With
Me.Label9.Enabled = False
Me.Label8.Enabled = False
With Me.Check4
.Enabled = False
.Value = 0
End With
With Me.Check6
.Enabled = False
.Value = 0
End With
With Me.Check7
.Enabled = False
.Value = 0
End With
With Me.Check8
.Enabled = False
.Value = 0
End With
With Me.Check9
.Enabled = False
.Value = 0
End With
With Me.Check10
.Enabled = False
.Value = 0
End With
With Me.Check11
.Enabled = False
.Value = 0
End With
With Me.Check12
.Enabled = False
.Value = 0
End With
For lpListNum = 0 To Me.List1.ListCount
If List1.List(lpListNum) = Me.hWinString.Caption Then
List1.RemoveItem lpListNum
End If
Next
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Else
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
End If
End Sub
Private Sub Command1_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_GotFocus()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_KeyPress(KeyAscii As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_LostFocus()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Command2_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Dim ans As Integer
ans = MsgBox("是否关闭此进程?所有未保存数据将丢失", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Const WM_CLOSE = &H10
Dim lpListNum As Long
For lpListNum = 0 To Me.List1.ListCount
If List1.List(lpListNum) = Me.hWinString.Caption Then
List1.RemoveItem lpListNum
End If
Next
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
PostMessage Me.hWinString.Caption, WM_CLOSE, ByVal 0&, ByVal 0&
SendMessage Me.hWinString.Caption, WM_CLOSE, ByVal 0&, 0&
Dim hProcess As Long
hProcess = OpenProcess(PROCESS_ALL_ACCESS, True, lpWindow.hThreadProcessID)
TerminateProcess hProcess, PROCESS_ALL_ACCESS
Me.hWinString.Caption = ""
Me.hDCString.Caption = ""
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Me.Check1.Enabled = False
Me.Check2.Enabled = False
Me.Check3.Enabled = False
Me.Check4.Enabled = False
Me.Check6.Enabled = False
Me.Check10.Enabled = False
Me.Check11.Enabled = False
Me.Check12.Enabled = False
Me.Check9.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Check8.Enabled = False
Me.Check7.Enabled = False
Me.HScroll1.Enabled = False
Me.Label9.Enabled = False
Me.Label8.Enabled = False
Me.Command1.Enabled = False
Me.Check10.Enabled = False
Me.Check11.Enabled = False
Me.Check12.Enabled = False
Me.Check9.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command1.Enabled = False
Me.Check1.Enabled = False
Me.Check2.Enabled = False
Me.Check3.Enabled = False
Me.Check4.Enabled = False
Me.Check6.Enabled = False
Me.Check7.Enabled = False
Check8.Enabled = False
Me.hDCString.Caption = ""
Me.hWinString.Caption = ""
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
With Me.HScroll1
.Min = 0
.Max = 255
.SmallChange = 5
.LargeChange = 10
.Enabled = False
.Value = 255
End With
With Me.Label9
.Enabled = False
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = HScroll1.Value
End With
With Me.Label8
.Enabled = False
End With
Check1.Value = 1
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
With Me.Check1
.Enabled = False
.Value = 1
End With
With Me.Check2
.Enabled = False
.Value = 0
End With
With Me.Check3
.Enabled = False
.Value = 0
End With
With Me.HScroll1
.Enabled = False
.Max = 255
.Min = 0
End With
Me.Label9.Enabled = False
Me.Label8.Enabled = False
With Me.Check4
.Enabled = False
.Value = 0
End With
With Me.Check6
.Enabled = False
.Value = 0
End With
With Me.Check7
.Enabled = False
.Value = 0
End With
With Me.Check8
.Enabled = False
.Value = 0
End With
With Me.Check9
.Enabled = False
.Value = 0
End With
With Me.Check10
.Enabled = False
.Value = 0
End With
With Me.Check11
.Enabled = False
.Value = 0
End With
With Me.Check12
.Enabled = False
.Value = 0
End With
For lpListNum = 0 To Me.List1.ListCount
If List1.List(lpListNum) = Me.hWinString.Caption Then
List1.RemoveItem lpListNum
End If
Next
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Else
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
End If
End Sub
Private Sub Command3_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
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
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Dim ans As Integer
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
ans = MsgBox("修改窗口标题是不可逆操作,继续?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Dim lpszStringOld As String * 256
GetWindowText Me.hWinString.Caption, lpszStringOld, 256
lpszCaptionNew = InputBox$("请输入新的窗口标题,注意:空字符串不被接受", "Set Caption", lpszStringOld)
If lpszCaptionNew = "" Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
End If
If lpszCaptionNew <> "" Then
SetWindowText Me.hWinString.Caption, lpszCaptionNew
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
Else
MsgBox "请不要输入空字符串!", vbCritical, "Error"
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
End If
Else
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End If
End Sub
Private Sub Form_Activate()
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Me.SetFocus
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Me.SetFocus
End Sub
Private Sub Form_Deactivate()
On Error Resume Next
Exit Sub
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End Sub
Private Sub Form_Initialize()
On Error Resume Next
On Error Resume Next
Unload Child1
Unload Child2
Unload Form2
Unload Form3
Unload Form5
Unload frmAbout
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
bCodeUse = False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
.Left = 0
.Top = 0
End With
If App.PrevInstance = True Then
MsgBox "本程序不允许同时运行2个及以上实例,请单击'确定',终止应用程序.", vbCritical, "Error"
bCodeUse = False
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
.Left = 0
.Top = 0
End With
With Me.mnuTop
.Checked = True
End With
With Me.mnuTrans
.Checked = False
End With
List1.Clear
Me.Check10.Enabled = False
Me.Check11.Enabled = False
Me.Check12.Enabled = False
Me.Check9.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
Me.Check1.Enabled = False
Me.Check2.Enabled = False
Me.Check3.Enabled = False
Me.Check4.Enabled = False
Me.Check6.Enabled = False
Me.Check7.Enabled = False
Me.hDCString.Caption = ""
Me.hWinString.Caption = ""
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Check8.Enabled = False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
With Me.HScroll1
.Min = 0
.Max = 255
.SmallChange = 5
.LargeChange = 10
.Enabled = False
.Value = 255
End With
With Me.Label9
.Enabled = False
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = HScroll1.Value
End With
With Me.Label8
.Enabled = False
End With
Check1.Value = 1
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
With Me.Check1
.Enabled = False
.Value = 1
End With
With Me.Check2
.Enabled = False
.Value = 0
End With
With Me.Check3
.Enabled = False
.Value = 0
End With
With Me.HScroll1
.Enabled = False
.Max = 255
.Min = 0
End With
Me.Label9.Enabled = False
Me.Label8.Enabled = False
With Me.Check4
.Enabled = False
.Value = 0
End With
With Me.Check6
.Enabled = False
.Value = 0
End With
With Me.Check7
.Enabled = False
.Value = 0
End With
With Me.Check8
.Enabled = False
.Value = 0
End With
With Me.Check9
.Enabled = False
.Value = 0
End With
With Me.Check10
.Enabled = False
.Value = 0
End With
With Me.Check11
.Enabled = False
.Value = 0
End With
With Me.Check12
.Enabled = False
.Value = 0
End With
End
End If
Dim ans As Integer
ans = MsgBox("请慎重使用本程序,如果使用不当可能导致系统不稳定或程序错误" & vbCrLf & vbCrLf & "单击'确定'继续使用程序;" & vbCrLf & "单击'取消'以关闭程序", vbExclamation + vbOKCancel, "Alert")
If ans = vbOK Then
bCodeUse = False
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
.Left = 0
.Top = 0
End With
With Me.mnuTop
.Checked = True
End With
With Me.mnuTrans
.Checked = False
End With
List1.Clear
Me.Check10.Enabled = False
Me.Check11.Enabled = False
Me.Check12.Enabled = False
Me.Check9.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
Me.Check1.Enabled = False
Me.Check2.Enabled = False
Me.Check3.Enabled = False
Me.Check4.Enabled = False
Me.Check6.Enabled = False
Me.Check7.Enabled = False
Me.hDCString.Caption = ""
Me.hWinString.Caption = ""
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Check8.Enabled = False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
With Me.HScroll1
.Min = 0
.Max = 255
.SmallChange = 5
.LargeChange = 10
.Enabled = False
.Value = 255
End With
With Me.Label9
.Enabled = False
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = HScroll1.Value
End With
With Me.Label8
.Enabled = False
End With
Check1.Value = 1
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
With Me.Check1
.Enabled = False
.Value = 1
End With
With Me.Check2
.Enabled = False
.Value = 0
End With
With Me.Check3
.Enabled = False
.Value = 0
End With
With Me.HScroll1
.Enabled = False
.Max = 255
.Min = 0
End With
Me.Label9.Enabled = False
Me.Label8.Enabled = False
With Me.Check4
.Enabled = False
.Value = 0
End With
With Me.Check6
.Enabled = False
.Value = 0
End With
With Me.Check7
.Enabled = False
.Value = 0
End With
With Me.Check8
.Enabled = False
.Value = 0
End With
With Me.Check9
.Enabled = False
.Value = 0
End With
With Me.Check10
.Enabled = False
.Value = 0
End With
With Me.Check11
.Enabled = False
.Value = 0
End With
With Me.Check12
.Enabled = False
.Value = 0
End With
Else
bCodeUse = False
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
.Left = 0
.Top = 0
End With
With Me.mnuTop
.Checked = True
End With
With Me.mnuTrans
.Checked = False
End With
List1.Clear
Me.Check10.Enabled = False
Me.Check11.Enabled = False
Me.Check12.Enabled = False
Me.Check9.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
Me.Check1.Enabled = False
Me.Check2.Enabled = False
Me.Check3.Enabled = False
Me.Check4.Enabled = False
Me.Check6.Enabled = False
Me.Check7.Enabled = False
Me.hDCString.Caption = ""
Me.hWinString.Caption = ""
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Check8.Enabled = False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
With Me.HScroll1
.Min = 0
.Max = 255
.SmallChange = 5
.LargeChange = 10
.Enabled = False
.Value = 255
End With
With Me.Label9
.Enabled = False
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = HScroll1.Value
End With
With Me.Label8
.Enabled = False
End With
Check1.Value = 1
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
With Me.Check1
.Enabled = False
.Value = 1
End With
With Me.Check2
.Enabled = False
.Value = 0
End With
With Me.Check3
.Enabled = False
.Value = 0
End With
With Me.HScroll1
.Enabled = False
.Max = 255
.Min = 0
End With
Me.Label9.Enabled = False
Me.Label8.Enabled = False
With Me.Check4
.Enabled = False
.Value = 0
End With
With Me.Check6
.Enabled = False
.Value = 0
End With
With Me.Check7
.Enabled = False
.Value = 0
End With
With Me.Check8
.Enabled = False
.Value = 0
End With
With Me.Check9
.Enabled = False
.Value = 0
End With
With Me.Check10
.Enabled = False
.Value = 0
End With
With Me.Check11
.Enabled = False
.Value = 0
End With
With Me.Check12
.Enabled = False
.Value = 0
End With
End
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
bCodeUse = False
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
.Left = 0
.Top = 0
End With
With Me.mnuTop
.Checked = True
End With
With Me.mnuTrans
.Checked = False
End With
List1.Clear
Me.Check10.Enabled = False
Me.Check11.Enabled = False
Me.Check12.Enabled = False
Me.Check9.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
Me.Check1.Enabled = False
Me.Check2.Enabled = False
Me.Check3.Enabled = False
Me.Check4.Enabled = False
Me.Check6.Enabled = False
Me.Check7.Enabled = False
Me.hDCString.Caption = ""
Me.hWinString.Caption = ""
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
Check8.Enabled = False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
With Me.HScroll1
.Min = 0
.Max = 255
.SmallChange = 5
.LargeChange = 10
.Enabled = False
.Value = 255
End With
With Me.Label9
.Enabled = False
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = HScroll1.Value
End With
With Me.Label8
.Enabled = False
End With
Check1.Value = 1
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
With Me.Check1
.Enabled = False
.Value = 1
End With
With Me.Check2
.Enabled = False
.Value = 0
End With
With Me.Check3
.Enabled = False
.Value = 0
End With
With Me.HScroll1
.Enabled = False
.Max = 255
.Min = 0
End With
Me.Label9.Enabled = False
Me.Label8.Enabled = False
With Me.Check4
.Enabled = False
.Value = 0
End With
With Me.Check6
.Enabled = False
.Value = 0
End With
With Me.Check7
.Enabled = False
.Value = 0
End With
With Me.Check8
.Enabled = False
.Value = 0
End With
With Me.Check9
.Enabled = False
.Value = 0
End With
With Me.Check10
.Enabled = False
.Value = 0
End With
With Me.Check11
.Enabled = False
.Value = 0
End With
With Me.Check12
.Enabled = False
.Value = 0
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOZORDER = &H4
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Dim rtn As Long
Const HWND_NOTOPMOST = -2
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
With Me.MouseHook
.Enabled = False
.Interval = 1000
End With
If List1.ListCount > 0 Then
Dim ans As Integer
ans = MsgBox("是否复位所有窗口的设定?", vbQuestion + vbYesNoCancel, "Ask")
Select Case ans
Case vbYes
Dim i As Integer
For i = 0 To List1.ListCount
Const SWP_FRAMECHANGED = &H20
Dim lpStyle As Long
Dim lpTemp As Long
lpTemp = GetWindowLong(List1.List(i), GWL_STYLE)
lpTemp = lpTemp Or WS_MINIMIZEBOX
lpTemp = lpTemp Or WS_MAXIMIZEBOX
lpTemp = lpTemp Or WS_SYSMENU
SetWindowLong List1.List(i), GWL_STYLE, lpTemp
lpStyle = GetWindowLong(List1.List(i), GWL_STYLE)
lpStyle = lpStyle Or WS_CAPTION
SetWindowLong List1.List(i), GWL_STYLE, lpStyle
Const SWP_NOACTIVATE = &H10
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOACTIVATE
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOACTIVATE
EnableWindow List1.List(i), True
GetSystemMenu List1.List(i), True
Dim lpReturn As Long
lpReturn = GetWindowLong(List1.List(i), GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong List1.List(i), GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes List1.List(i), 0, 255, LWA_ALPHA
SetWindowPos List1.List(i), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
GetSystemMenu List1.List(i), True
lpTemp = GetWindowLong(List1.List(i), GWL_STYLE)
lpTemp = lpTemp Or WS_MINIMIZEBOX
lpTemp = lpTemp Or WS_MAXIMIZEBOX
lpTemp = lpTemp Or WS_SYSMENU
SetWindowLong List1.List(i), GWL_STYLE, lpTemp
ShowWindow List1.List(i), 1
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
Next
Cancel = 0
Case vbNo
Cancel = 0
Case Else
Cancel = 666 + 444
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Select
Else
With Me.cSysTray1
.InTray = False
.TrayTip = "Win Tool - 双击还原窗口"
End With
End
End If
End Sub
Private Sub Form_Terminate()
On Error Resume Next
With Me.cSysTray1
.InTray = False
End With
On Error Resume Next
End
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
End
End Sub
Private Sub Frame1_Click()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame1_DblClick()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame1_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame1_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame1_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame2_Click()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame2_DblClick()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame2_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame2_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame2_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame2_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame2_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame2_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame2_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame3_Click()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame3_DblClick()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame3_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame3_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame3_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame3_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame3_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame3_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame3_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame4_Click()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame4_DblClick()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame4_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame4_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame4_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame4_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame4_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame4_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Frame4_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_Change()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
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
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
If lpszCaption.Caption = "" Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
Else
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End If
End Sub
Private Sub hDCString_DblClick()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_LinkClose()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_LinkError(LinkErr As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_LinkNotify()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_LinkOpen(Cancel As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hDCString_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
On Error Resume Next
Dim lpReturn As Long
lpReturn = GetWindowLong(lpWindow.hWindow, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong lpWindow.hWindow, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes lpWindow.hWindow, 0, HScroll1.Value, LWA_ALPHA
With Me.Label9
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = True
End With
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
If List1.ListCount >= 2 Then
For nTmp = 0 To List1.ListIndex
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If List1.ListCount >= 2 Then
For nTmp = 0 To List1.ListIndex
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If List1.ListCount >= 2 Then
For nTmp = 0 To List1.ListIndex
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If List1.ListCount >= 2 Then
For nTmp = 0 To List1.ListIndex
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If List1.List(0) = "" Then
List1.RemoveItem 0
End If
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
If List1.List(0) = "" Then
List1.RemoveItem 0
End If
Exit Sub
End If
End If
If List1.List(0) = "" Then
List1.RemoveItem 0
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
If List1.ListCount >= 2 Then
For nTmp = 0 To List1.ListIndex
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If List1.List(0) = "" Then
List1.RemoveItem 0
End If
Exit Sub
End If
Next
If HScroll1.Value = 0 Then
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
If List1.ListCount >= 2 Then
For nTmp = 0 To List1.ListIndex
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If List1.List(0) = "" Then
List1.RemoveItem 0
End If
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
If List1.List(0) = "" Then
List1.RemoveItem 0
End If
Exit Sub
End If
End If
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
If List1.List(0) = "" Then
List1.RemoveItem 0
End If
Exit Sub
End If
Next
If HScroll1.Value = 0 Then
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
If List1.List(0) = "" Then
List1.RemoveItem 0
End If
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
If List1.List(0) = "" Then
List1.RemoveItem 0
End If
Exit Sub
End If
End If
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
If List1.List(0) = "" Then
List1.RemoveItem 0
End If
Exit Sub
End If
Next
End If
End If
End Sub
Private Sub HScroll1_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub HScroll1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub HScroll1_GotFocus()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub HScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub HScroll1_KeyPress(KeyAscii As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub HScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub HScroll1_LostFocus()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub HScroll1_Scroll()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub HScroll1_Validate(Cancel As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_Change()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
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
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
If lpszCaption.Caption = "" Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
Else
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End If
End Sub
Private Sub hWinString_DblClick()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_LinkClose()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_LinkError(LinkErr As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_LinkNotify()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_LinkOpen(Cancel As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub hWinString_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Image1_Click()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Image1_DblClick()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Image1_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Image1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Image1_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Image1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Image1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Image1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Image1_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Image1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_Change()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_Click()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_DblClick()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_LinkClose()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_LinkError(LinkErr As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_LinkNotify()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_LinkOpen(Cancel As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_Change(Index As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_Click(Index As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_DblClick(Index As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_LinkClose(Index As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_LinkError(Index As Integer, LinkErr As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_LinkNotify(Index As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_LinkOpen(Index As Integer, Cancel As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_OLECompleteDrag(Index As Integer, Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_OLEGiveFeedback(Index As Integer, Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_OLESetData(Index As Integer, Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label2_OLEStartDrag(Index As Integer, Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_Change()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_Click()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_DblClick()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_LinkClose()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_LinkError(LinkErr As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_LinkNotify()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_LinkOpen(Cancel As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label8_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_Change()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim lpReturn As Long
Dim lpValue As Integer
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
lpValue = Val(InputBox("请输入要设定的透明度,范围0-255", "Alpha", "128"))
If 0 <= lpValue And lpValue <= 255 Then
lpReturn = GetWindowLong(lpWindow.hWindow, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong lpWindow.hWindow, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes lpWindow.hWindow, 0, lpValue, LWA_ALPHA
With Me.HScroll1
.Value = lpValue
.Enabled = True
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
If HScroll1.Value = 0 Then
If List1.ListCount = 0 Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If List1.ListCount = 1 Then
If Me.hWinString.Caption <> List1.List(0) Then
List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
For cunt = 0 To List1.ListCount - 1
If Me.hWinString.Caption <> List1.List(cunt) Then
addvar = addvar + 1
If addvar = List1.ListCount Then
List1.AddItem hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
End If
Else
MsgBox "输入的数值有误", vbCritical, "Error"
lpValue = HScroll1.Value
lpReturn = GetWindowLong(lpWindow.hWindow, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong lpWindow.hWindow, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes lpWindow.hWindow, 0, lpValue, LWA_ALPHA
With Me.HScroll1
.Value = lpValue
.Enabled = True
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
If HScroll1.Value = 0 Then
List1.AddItem Me.hWinString.Caption
End If
End If
End Sub
Private Sub Label9_DblClick()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_LinkClose()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_LinkError(LinkErr As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_LinkNotify()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_LinkOpen(Cancel As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Label9_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_Click()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_DblClick()
On Error Resume Next
Dim rtn As Long
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim ans As Integer
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
If List1.ListIndex >= 0 Then
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
ans = MsgBox("还原这个窗口的设定吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
GetSystemMenu List1.List(List1.ListIndex), True
EnableWindow List1.List(List1.ListIndex), True
Const SWP_FRAMECHANGED = &H20
Dim lpStyle As Long
lpStyle = GetWindowLong(List1.List(List1.ListIndex), GWL_STYLE)
lpStyle = lpStyle Or WS_CAPTION
SetWindowLong List1.List(List1.ListIndex), GWL_STYLE, lpStyle
Const SWP_NOACTIVATE = &H10
SetWindowPos List1.List(List1.ListIndex), 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOACTIVATE
Dim lpReturn As Long
lpReturn = GetWindowLong(List1.List(List1.ListIndex), GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong List1.List(List1.ListIndex), GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes List1.List(List1.ListIndex), 0, 255, LWA_ALPHA
SetWindowPos List1.List(List1.ListIndex), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
ShowWindow List1.List(List1.ListIndex), 1
GetSystemMenu List1.List(List1.ListIndex), True
Dim lpTemp As Long
lpTemp = GetWindowLong(List1.List(List1.ListIndex), GWL_STYLE)
lpTemp = lpTemp Or WS_MINIMIZEBOX
lpTemp = lpTemp Or WS_MAXIMIZEBOX
lpTemp = lpTemp Or WS_SYSMENU
SetWindowLong List1.List(List1.ListIndex), GWL_STYLE, lpTemp
ShowWindow List1.List(List1.ListIndex), 1
SetWindowPos List1.List(List1.ListIndex), 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
If Me.hWinString.Caption = List1.List(List1.ListIndex) Then
bCodeUse = True
With Me.Check1
.Enabled = True
.Value = 1
End With
bCodeUse = True
With Me.Check10
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check11
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check12
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check2
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check3
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check4
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check6
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check7
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check8
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check9
.Enabled = True
.Value = 0
End With
With Me.HScroll1
.Max = 255
.Min = 0
.LargeChange = 10
.SmallChange = 5
.Enabled = False
End With
Me.Label8.Enabled = False
With Me.Label9
.Enabled = False
.Caption = HScroll1.Value
.Alignment = 2
End With
End If
List1.RemoveItem List1.ListIndex
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
If Me.mnuEnable.Checked = True Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Select Case mnuEnable.Checked
Case True
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Case False
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Case Else
With MouseHook
.Interval = 1000
.Enabled = True
End With
End Select
Else
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
With MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Select Case mnuEnable.Checked
Case True
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Case False
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Case Else
With MouseHook
.Interval = 1000
.Enabled = True
End With
End Select
End If
Else
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Select Case mnuEnable.Checked
Case True
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Case False
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Case Else
With MouseHook
.Interval = 1000
.Enabled = True
End With
End Select
If Me.mnuEnable.Checked = True Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Select Case mnuEnable.Checked
Case True
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Case False
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Case Else
With MouseHook
.Interval = 1000
.Enabled = True
End With
End Select
Else
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
With MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Select Case mnuEnable.Checked
Case True
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Case False
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Case Else
With MouseHook
.Interval = 1000
.Enabled = True
End With
End Select
End If
Exit Sub
End If
Else
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Select Case mnuEnable.Checked
Case True
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Case False
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Case Else
With MouseHook
.Interval = 1000
.Enabled = True
End With
End Select
If Me.mnuEnable.Checked = True Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Select Case mnuEnable.Checked
Case True
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Case False
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Case Else
With MouseHook
.Interval = 1000
.Enabled = True
End With
End Select
Else
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
With MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Select Case mnuEnable.Checked
Case True
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Case False
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Case Else
With MouseHook
.Interval = 1000
.Enabled = True
End With
End Select
End If
Exit Sub
End If
End Sub
Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)
On Error Resume Next
If Button = 2 Then
PopupMenu Me.mnuTray
Else
Exit Sub
End If
End Sub
Private Sub List1_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_GotFocus()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_ItemCheck(Item As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
On Error Resume Next
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Dim ans As Integer
If List1.ListIndex >= 0 Then
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
ans = MsgBox("还原这个窗口的设定吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
GetSystemMenu List1.List(List1.ListIndex), True
EnableWindow List1.List(List1.ListIndex), True
Const SWP_FRAMECHANGED = &H20
Dim lpStyle As Long
lpStyle = GetWindowLong(List1.List(List1.ListIndex), GWL_STYLE)
lpStyle = lpStyle Or WS_CAPTION
SetWindowLong List1.List(List1.ListIndex), GWL_STYLE, lpStyle
Const SWP_NOACTIVATE = &H10
SetWindowPos List1.List(List1.ListIndex), 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOACTIVATE
Dim lpReturn As Long
lpReturn = GetWindowLong(List1.List(List1.ListIndex), GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong List1.List(List1.ListIndex), GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes List1.List(List1.ListIndex), 0, 255, LWA_ALPHA
SetWindowPos List1.List(List1.ListIndex), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
GetSystemMenu List1.List(List1.ListIndex), True
ShowWindow List1.List(List1.ListIndex), 1
SetWindowPos List1.List(List1.ListIndex), 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
If Me.hWinString.Caption = List1.List(List1.ListIndex) Then
bCodeUse = True
With Me.Check1
.Enabled = True
.Value = 1
End With
bCodeUse = True
With Me.Check10
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check11
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check12
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check2
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check3
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check4
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check6
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check7
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check8
.Enabled = True
.Value = 0
End With
bCodeUse = True
With Me.Check9
.Enabled = True
.Value = 0
End With
With Me.HScroll1
.Max = 255
.Min = 0
.LargeChange = 10
.SmallChange = 5
.Enabled = False
End With
Me.Label8.Enabled = False
With Me.Label9
.Enabled = False
.Caption = HScroll1.Value
.Alignment = 2
End With
End If
List1.RemoveItem List1.ListIndex
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Else
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
If Me.mnuEnable.Checked = True Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Else
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
With MouseHook
.Interval = 1000
.Enabled = True
End With
End If
Exit Sub
End If
Else
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
If Me.mnuEnable.Checked = True Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
Else
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
With MouseHook
.Interval = 1000
.Enabled = True
End With
End If
Exit Sub
End If
End If
End Sub
Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_LostFocus()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_Scroll()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub List1_Validate(Cancel As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_Change()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
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
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
If lpszCaption.Caption = "" Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
Else
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End If
End Sub
Private Sub lpszCaption_DblClick()
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_LinkClose()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_LinkError(LinkErr As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_LinkNotify()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_LinkOpen(Cancel As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszCaption_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_Change()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
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
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
If lpszCaption.Caption = "" Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
Else
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End If
End Sub
Private Sub lpszClass_DblClick()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_LinkClose()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_LinkError(LinkErr As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_LinkNotify()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_LinkOpen(Cancel As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszClass_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_Change()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
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
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
If lpszCaption.Caption = "" Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
Else
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End If
End Sub
Private Sub lpszThread_DblClick()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_LinkClose()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_LinkError(LinkErr As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_LinkNotify()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_LinkOpen(Cancel As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_OLESetData(Data As DataObject, DataFormat As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub lpszThread_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub mnuAbout_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Me.Hide
frmAbout.Show
End Sub
Private Sub mnuDisable_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Sub
Private Sub mnuEnable_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
End Sub
Private Sub mnuEnd_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
On Error Resume Next
Dim rtn As Long
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
If List1.ListCount > 0 Then
Dim ans As Integer
ans = MsgBox("是否复位所有窗口的设定?", vbQuestion + vbYesNoCancel, "Ask")
Select Case ans
Case vbYes
Dim i As Integer
For i = 0 To List1.ListCount
EnableWindow List1.List(i), True
Dim lpStyle As Long
lpStyle = GetWindowLong(List1.List(i), GWL_STYLE)
lpStyle = lpStyle Or WS_CAPTION
SetWindowLong List1.List(i), GWL_STYLE, lpStyle
Const SWP_FRAMECHANGED = &H20
Const SWP_NOACTIVATE = &H10
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOACTIVATE
GetSystemMenu List1.List(i), True
Dim lpReturn As Long
lpReturn = GetWindowLong(List1.List(i), GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong List1.List(i), GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes List1.List(i), 0, 255, LWA_ALPHA
SetWindowPos List1.List(i), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
GetSystemMenu List1.List(i), True
Dim lpTemp As Long
lpTemp = GetWindowLong(List1.List(i), GWL_STYLE)
lpTemp = lpTemp Or WS_MINIMIZEBOX
lpTemp = lpTemp Or WS_MAXIMIZEBOX
lpTemp = lpTemp Or WS_SYSMENU
SetWindowLong List1.List(i), GWL_STYLE, lpTemp
ShowWindow List1.List(i), 1
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
Next
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
With Me.cSysTray1
.InTray = False
.TrayTip = "Win Tool - 双击还原窗口"
End With
End
Case vbNo
With Me.cSysTray1
.InTray = False
.TrayTip = "Win Tool - 双击还原窗口"
End With
End
Case Else
With Me.cSysTray1
.InTray = True
.TrayTip = "Win Tool - 双击还原窗口"
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
End Select
Else
With Me.cSysTray1
.InTray = False
.TrayTip = "Win Tool - 双击还原窗口"
End With
End
End If
End Sub
Private Sub mnuExit_Click()
On Error Resume Next
On Error Resume Next
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Dim rtn As Long
Const HWND_NOTOPMOST = -2
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
If List1.ListCount > 0 Then
Dim ans As Integer
ans = MsgBox("是否复位所有窗口的设定?", vbQuestion + vbYesNoCancel, "Ask")
Select Case ans
Case vbYes
Dim i As Integer
For i = 0 To List1.ListCount
Dim lpStyle As Long
lpStyle = GetWindowLong(List1.List(i), GWL_STYLE)
lpStyle = lpStyle Or WS_CAPTION
SetWindowLong List1.List(i), GWL_STYLE, lpStyle
Const SWP_FRAMECHANGED = &H20
Const SWP_NOACTIVATE = &H10
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOACTIVATE
EnableWindow List1.List(i), True
GetSystemMenu List1.List(i), True
Dim lpReturn As Long
lpReturn = GetWindowLong(List1.List(i), GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong List1.List(i), GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes List1.List(i), 0, 255, LWA_ALPHA
SetWindowPos List1.List(i), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
GetSystemMenu List1.List(i), True
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
Next
With Me.cSysTray1
.InTray = False
.TrayTip = "Win Tool - 双击还原窗口"
End With
End
Case vbNo
With Me.cSysTray1
.InTray = False
.TrayTip = "Win Tool - 双击还原窗口"
End With
End
Case Else
With Me.cSysTray1
.InTray = True
.TrayTip = "Win Tool - 双击还原窗口"
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
End Select
Else
With Me.cSysTray1
.InTray = False
.TrayTip = "Win Tool - 双击还原窗口"
End With
End
End If
End Sub
Private Sub mnuHelpForm_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Me.Hide
Form2.Show
End Sub
Private Sub mnuInfo_Click()
On Error Resume Next
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
MsgBox "窗口信息列表:" & vbCrLf & "标题:" & vbCrLf & .lpszCaption & vbCrLf & "类名:" & vbCrLf & .lpszClass & vbCrLf & "隶属于:" & vbCrLf & .lpszThread & vbCrLf & "窗口句柄ID:" & vbCrLf & .lpszHandle & vbCrLf & "设备上下文ID:" & vbCrLf & .lpszDC, vbInformation, "Info"
End With
End Sub
Private Sub mnuInfoA_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
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
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
If lpszCaption.Caption = "" Then
MsgBox "没有活动窗口", vbCritical, "Error"
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
Else
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End If
End Sub
Private Sub mnuMini_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Enabled = False
.Interval = 1000
End With
Me.Hide
With Me.cSysTray1
.InTray = True
.TrayTip = "Win Tool - 双击还原窗口"
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Visible = False
.Height = 7335
.Width = 8100
End With
End Sub
Private Sub mnuRefesh_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_FRAMECHANGED = &H20
Const SWP_NOZORDER = &H4
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
bCodeUse = False
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Me.Check10.Enabled = False
Me.Check11.Enabled = False
Me.Check12.Enabled = False
Me.Check9.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command1.Enabled = False
Me.Check1.Enabled = False
Me.Check2.Enabled = False
Me.Check3.Enabled = False
Me.Check4.Enabled = False
Me.Check6.Enabled = False
Me.Check7.Enabled = False
Check8.Enabled = False
Me.hDCString.Caption = ""
Me.hWinString.Caption = ""
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
With Me.HScroll1
.Min = 0
.Max = 255
.SmallChange = 5
.LargeChange = 10
.Enabled = False
.Value = 255
End With
With Me.Label9
.Enabled = False
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = HScroll1.Value
End With
With Me.Label8
.Enabled = False
End With
Check1.Value = 1
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
With Me.Check1
.Enabled = False
.Value = 1
End With
With Me.Check2
.Enabled = False
.Value = 0
End With
With Me.Check3
.Enabled = False
.Value = 0
End With
With Me.HScroll1
.Enabled = False
.Max = 255
.Min = 0
End With
Me.Label9.Enabled = False
Me.Label8.Enabled = False
With Me.Check4
.Enabled = False
.Value = 0
End With
With Me.Check6
.Enabled = False
.Value = 0
End With
With Me.Check7
.Enabled = False
.Value = 0
End With
With Me.Check8
.Enabled = False
.Value = 0
End With
With Me.Check9
.Enabled = False
.Value = 0
End With
With Me.Check10
.Enabled = False
.Value = 0
End With
With Me.Check11
.Enabled = False
.Value = 0
End With
With Me.Check12
.Enabled = False
.Value = 0
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
If List1.ListCount = 0 Then
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Me.Check10.Enabled = False
Me.Check11.Enabled = False
Me.Check12.Enabled = False
Me.Check9.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
Me.Check1.Enabled = False
Me.Check2.Enabled = False
Me.Check3.Enabled = False
Me.Check4.Enabled = False
Me.Check6.Enabled = False
Me.Check7.Enabled = False
Check8.Enabled = False
Me.hDCString.Caption = ""
Me.hWinString.Caption = ""
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
With Me.HScroll1
.Min = 0
.Max = 255
.SmallChange = 5
.LargeChange = 10
.Enabled = False
.Value = 255
End With
With Me.Label9
.Enabled = False
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = HScroll1.Value
End With
With Me.Label8
.Enabled = False
End With
Check1.Value = 1
With Me.Check1
.Enabled = False
.Value = 1
End With
With Me.Check2
.Enabled = False
.Value = 0
End With
With Me.Check3
.Enabled = False
.Value = 0
End With
With Me.HScroll1
.Enabled = False
.Max = 255
.Min = 0
End With
Me.Label9.Enabled = False
Me.Label8.Enabled = False
With Me.Check4
.Enabled = False
.Value = 0
End With
With Me.Check6
.Enabled = False
.Value = 0
End With
With Me.Check7
.Enabled = False
.Value = 0
End With
With Me.Check8
.Enabled = False
.Value = 0
End With
With Me.Check9
.Enabled = False
.Value = 0
End With
With Me.Check10
.Enabled = False
.Value = 0
End With
With Me.Check11
.Enabled = False
.Value = 0
End With
With Me.Check12
.Enabled = False
.Value = 0
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
If Me.mnuEnable.Checked = True Then
With Me.MouseHook
.Enabled = True
.Interval = 1000
End With
ElseIf Me.mnuDisable.Checked = True Then
With Me.MouseHook
.Enabled = False
.Interval = 1000
End With
Else
With Me.MouseHook
.Enabled = True
.Interval = 1000
End With
End If
Exit Sub
End If
Dim dwAnswer As Integer
dwAnswer = MsgBox("是否复位已经被修改的窗口?", vbYesNo + vbQuestion, "Ask")
If dwAnswer = vbYes Then
If List1.ListCount = 0 Then
If 1 = 245 Then
MsgBox "没有可以操作的内容", vbCritical, "Error"
End If
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
End If
Dim ans As Integer
ans = vbYes
Select Case ans
Case vbYes
Dim i As Integer
For i = 0 To List1.ListCount
EnableWindow List1.List(i), True
GetSystemMenu List1.List(i), True
Dim lpStyle As Long
lpStyle = GetWindowLong(List1.List(i), GWL_STYLE)
lpStyle = lpStyle Or WS_CAPTION
SetWindowLong List1.List(i), GWL_STYLE, lpStyle
Const SWP_NOACTIVATE = &H10
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOACTIVATE
Dim lpReturn As Long
lpReturn = GetWindowLong(List1.List(i), GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong List1.List(i), GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes List1.List(i), 0, 255, LWA_ALPHA
SetWindowPos List1.List(i), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
GetSystemMenu List1.List(i), True
Dim lpTemp As Long
lpTemp = GetWindowLong(List1.List(i), GWL_STYLE)
lpTemp = lpTemp Or WS_MINIMIZEBOX
lpTemp = lpTemp Or WS_MAXIMIZEBOX
lpTemp = lpTemp Or WS_SYSMENU
SetWindowLong List1.List(i), GWL_STYLE, lpTemp
ShowWindow List1.List(i), 1
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
Next
If 1 = 245 Then
MsgBox "复位操作成功完成", vbExclamation, "Info"
End If
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
For i = 0 To List1.ListCount
EnableWindow List1.List(i), True
GetSystemMenu List1.List(i), True
lpStyle = GetWindowLong(List1.List(i), GWL_STYLE)
lpStyle = lpStyle Or WS_CAPTION
SetWindowLong List1.List(i), GWL_STYLE, lpStyle
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOACTIVATE
lpReturn = GetWindowLong(List1.List(i), GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong List1.List(i), GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes List1.List(i), 0, 255, LWA_ALPHA
SetWindowPos List1.List(i), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
GetSystemMenu List1.List(i), True
ShowWindow List1.List(i), 1
Next
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Me.Check10.Enabled = False
Me.Check11.Enabled = False
Me.Check12.Enabled = False
Me.Check9.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
Me.Check1.Enabled = False
Me.Check2.Enabled = False
Me.Check3.Enabled = False
Me.Check4.Enabled = False
Me.Check6.Enabled = False
Me.Check7.Enabled = False
Check8.Enabled = False
Me.hDCString.Caption = ""
Me.hWinString.Caption = ""
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
With Me.HScroll1
.Min = 0
.Max = 255
.SmallChange = 5
.LargeChange = 10
.Enabled = False
.Value = 255
End With
With Me.Label9
.Enabled = False
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = HScroll1.Value
End With
With Me.Label8
.Enabled = False
End With
Check1.Value = 1
With Me.Check1
.Enabled = False
.Value = 1
End With
With Me.Check2
.Enabled = False
.Value = 0
End With
With Me.Check3
.Enabled = False
.Value = 0
End With
With Me.HScroll1
.Enabled = False
.Max = 255
.Min = 0
End With
Me.Label9.Enabled = False
Me.Label8.Enabled = False
With Me.Check4
.Enabled = False
.Value = 0
End With
With Me.Check6
.Enabled = False
.Value = 0
End With
With Me.Check7
.Enabled = False
.Value = 0
End With
With Me.Check8
.Enabled = False
.Value = 0
End With
With Me.Check9
.Enabled = False
.Value = 0
End With
With Me.Check10
.Enabled = False
.Value = 0
End With
With Me.Check11
.Enabled = False
.Value = 0
End With
With Me.Check12
.Enabled = False
.Value = 0
End With
List1.Clear
Else
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Me.Check10.Enabled = False
Me.Check11.Enabled = False
Me.Check12.Enabled = False
Me.Check9.Enabled = False
Command3.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
Me.Check1.Enabled = False
Me.Check2.Enabled = False
Me.Check3.Enabled = False
Me.Check4.Enabled = False
Me.Check6.Enabled = False
Me.Check7.Enabled = False
Check8.Enabled = False
Me.hDCString.Caption = ""
Me.hWinString.Caption = ""
Me.lpszCaption.Caption = ""
Me.lpszClass.Caption = ""
Me.lpszThread.Caption = ""
With Me.HScroll1
.Min = 0
.Max = 255
.SmallChange = 5
.LargeChange = 10
.Enabled = False
.Value = 255
End With
With Me.Label9
.Enabled = False
.Alignment = 2
.BorderStyle = 1
.BackStyle = 0
.Caption = HScroll1.Value
End With
With Me.Label8
.Enabled = False
End With
Check1.Value = 1
With Me.Check1
.Enabled = False
.Value = 1
End With
With Me.Check2
.Enabled = False
.Value = 0
End With
With Me.Check3
.Enabled = False
.Value = 0
End With
With Me.HScroll1
.Enabled = False
.Max = 255
.Min = 0
End With
Me.Label9.Enabled = False
Me.Label8.Enabled = False
With Me.Check4
.Enabled = False
.Value = 0
End With
With Me.Check6
.Enabled = False
.Value = 0
End With
With Me.Check7
.Enabled = False
.Value = 0
End With
With Me.Check8
.Enabled = False
.Value = 0
End With
With Me.Check9
.Enabled = False
.Value = 0
End With
With Me.Check10
.Enabled = False
.Value = 0
End With
With Me.Check11
.Enabled = False
.Value = 0
End With
With Me.Check12
.Enabled = False
.Value = 0
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End If
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End Sub
Private Sub mnuReset_Click()
On Error Resume Next
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
If 1 = 245 Then
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
End If
If List1.ListCount = 0 Then
MsgBox "没有可以操作的内容", vbCritical, "Error"
If 1 = 245 Then
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End If
Exit Sub
End If
Dim ans As Integer
ans = MsgBox("是否复位所有窗口的设定?", vbQuestion + vbYesNoCancel, "Ask")
Select Case ans
Case vbYes
Dim i As Integer
For i = 0 To List1.ListCount
Const SWP_FRAMECHANGED = &H20
Dim lpStyle As Long
lpStyle = GetWindowLong(List1.List(i), GWL_STYLE)
lpStyle = lpStyle Or WS_CAPTION
SetWindowLong List1.List(i), GWL_STYLE, lpStyle
Const SWP_NOACTIVATE = &H10
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOACTIVATE
EnableWindow List1.List(i), True
GetSystemMenu List1.List(i), True
Dim lpReturn As Long
lpReturn = GetWindowLong(List1.List(i), GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong List1.List(i), GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes List1.List(i), 0, 255, LWA_ALPHA
SetWindowPos List1.List(i), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
GetSystemMenu List1.List(i), True
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
If Me.hWinString.Caption = List1.List(i) Then
With Me.Check1
.Enabled = True
.Value = 1
End With
With Me.Check10
.Enabled = True
.Value = 0
End With
With Me.Check11
.Enabled = True
.Value = 0
End With
With Me.Check12
.Enabled = True
.Value = 0
End With
With Me.Check2
.Enabled = True
.Value = 0
End With
With Me.Check3
.Enabled = True
.Value = 0
End With
With Me.Check4
.Enabled = True
.Value = 0
End With
With Me.Check6
.Enabled = True
.Value = 0
End With
With Me.Check7
.Enabled = True
.Value = 0
End With
With Me.Check8
.Enabled = True
.Value = 0
End With
With Me.Check9
.Enabled = True
.Value = 0
End With
With Me.HScroll1
.Max = 255
.Min = 0
.LargeChange = 10
.SmallChange = 5
.Enabled = False
End With
Me.Label8.Enabled = False
With Me.Label9
.Enabled = False
.Caption = HScroll1.Value
.Alignment = 2
End With
End If
Next
MsgBox "复位操作成功完成", vbExclamation, "Info"
End Select
End Sub
Private Sub mnuRestore_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Const SWP_NOZORDER = &H4
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
If List1.ListCount = 0 Then
MsgBox "没有可以操作的内容", vbCritical, "Error"
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
End If
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Dim ans As Integer
ans = MsgBox("是否复位所有窗口的设定?", vbQuestion + vbYesNoCancel, "Ask")
Select Case ans
Case vbYes
Dim i As Integer
For i = 0 To List1.ListCount
EnableWindow List1.List(i), True
GetSystemMenu List1.List(i), True
Const SWP_FRAMECHANGED = &H20
Dim lpStyle As Long
lpStyle = GetWindowLong(List1.List(i), GWL_STYLE)
lpStyle = lpStyle Or WS_CAPTION
SetWindowLong List1.List(i), GWL_STYLE, lpStyle
Const SWP_NOACTIVATE = &H10
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOACTIVATE
Dim lpReturn As Long
lpReturn = GetWindowLong(List1.List(i), GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong List1.List(i), GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes List1.List(i), 0, 255, LWA_ALPHA
SetWindowPos List1.List(i), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
GetSystemMenu List1.List(i), True
Dim lpTemp As Long
lpTemp = GetWindowLong(List1.List(i), GWL_STYLE)
lpTemp = lpTemp Or WS_MINIMIZEBOX
lpTemp = lpTemp Or WS_MAXIMIZEBOX
lpTemp = lpTemp Or WS_SYSMENU
SetWindowLong List1.List(i), GWL_STYLE, lpTemp
ShowWindow List1.List(i), 1
SetWindowPos List1.List(i), 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
If Me.hWinString.Caption = List1.List(i) Then
With Me.Check1
.Enabled = True
.Value = 1
End With
With Me.Check10
.Enabled = True
.Value = 0
End With
With Me.Check11
.Enabled = True
.Value = 0
End With
With Me.Check12
.Enabled = True
.Value = 0
End With
With Me.Check2
.Enabled = True
.Value = 0
End With
With Me.Check3
.Enabled = True
.Value = 0
End With
With Me.Check4
.Enabled = True
.Value = 0
End With
With Me.Check6
.Enabled = True
.Value = 0
End With
With Me.Check7
.Enabled = True
.Value = 0
End With
With Me.Check8
.Enabled = True
.Value = 0
End With
With Me.Check9
.Enabled = True
.Value = 0
End With
With Me.HScroll1
.Max = 255
.Min = 0
.LargeChange = 10
.SmallChange = 5
.Enabled = False
End With
Me.Label8.Enabled = False
With Me.Label9
.Enabled = False
.Caption = HScroll1.Value
.Alignment = 2
End With
End If
Next
MsgBox "复位操作成功完成", vbExclamation, "Info"
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End Sub
Private Sub mnuShow_Click()
On Error Resume Next
With Me.MouseHook
.Enabled = True
.Interval = 1000
End With
Me.Show
Form1.Show
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
With Me.cSysTray1
.InTray = False
.TrayTip = "Win Tool - 双击还原窗口"
End With
Form1.WindowState = 0
Form1.Show
With Form1
.WindowState = 0
.Show
.Visible = True
End With
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Sub
Private Sub StartTaskMgr()
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Dim ans As Integer
Dim uMsg As String
StartProc:
On Error GoTo ep
Shell "Taskmgr.exe", vbNormalFocus
Exit Sub
ep:
On Error GoTo ep
uMsg = "发生系统错误:" & vbCrLf & Err.Description
ans = MsgBox(uMsg, vbCritical + vbAbortRetryIgnore, "Error")
Select Case ans
Case vbAbort
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
Case vbRetry
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
mnuTasks_Click
Case vbIgnore
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
Case Else
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
End Select
End Sub
Private Sub mnuTasks_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Dim ans As Integer
Dim uMsg As String
StartProc:
On Error GoTo ep
Shell "Taskmgr.exe", vbNormalFocus
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
ep:
On Error GoTo ep
uMsg = "发生系统错误:" & vbCrLf & Err.Description
ans = MsgBox(uMsg, vbCritical + vbAbortRetryIgnore, "Error")
Select Case ans
Case vbAbort
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
Case vbRetry
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
StartTaskMgr
Case vbIgnore
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
Case Else
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
End Select
End Sub
Private Sub mnuTaskT_Click()
Const SWP_NOZORDER = &H4
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Dim ans As Integer
Dim uMsg As String
StartProc:
On Error GoTo ep
Shell "Taskmgr.exe", vbNormalFocus
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
ep:
On Error GoTo ep
uMsg = "发生系统错误:" & vbCrLf & Err.Description
ans = MsgBox(uMsg, vbCritical + vbAbortRetryIgnore, "Error")
Select Case ans
Case vbAbort
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
Case vbRetry
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
StartTaskMgr
Case vbIgnore
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
Case Else
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Exit Sub
End Select
End Sub
Private Sub mnuTop_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
mnuTop.Checked = True
Case True
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
mnuTop.Checked = False
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub mnuTrans_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
mnuTrans.Checked = True
Exit Sub
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
mnuTrans.Checked = False
Exit Sub
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
End Sub
Private Sub mnuViewCurWin_Click()
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
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
If Me.hWinString.Caption = "" Then
MsgBox "没有选择窗口", vbCritical, "Error"
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
Exit Sub
End If
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
If lpszCaption.Caption = "" Then
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Exit Sub
Else
With dwWinInfo
.lpszCaption = Me.lpszCaption.Caption
.lpszClass = Me.lpszClass.Caption
.lpszDC = Me.hDCString.Caption
.lpszHandle = Me.hWinString.Caption
.lpszThread = Me.lpszThread.Caption
End With
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End If
End Sub
Private Sub mnuVPWWCW_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Me.Hide
Form5.Show
End Sub
Private Sub mnuWindView_Click()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
Me.Hide
Form3.Show
End Sub
Private Sub MouseHook_Timer()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
DoEvents
Dim hWndTmp As Long
hWndTmp = GetForegroundWindow()
If hWndTmp = Me.hwnd Then
Exit Sub
End If
If hWndTmp = Child1.hwnd Then
Exit Sub
End If
If hWndTmp = Child2.hwnd Then
Exit Sub
End If
If hWndTmp = Form1.hwnd Then
Exit Sub
End If
If hWndTmp = Form3.hwnd Then
Exit Sub
End If
If hWndTmp = Form2.hwnd Then
Exit Sub
End If
If hWndTmp = Form5.hwnd Then
Exit Sub
End If
If hWndTmp = frmAbout.hwnd Then
Exit Sub
End If
If hWndTmp = Child1.hwnd Then
Exit Sub
End If
With lpWindow
.hWindow = GetForegroundWindow()
End With
If lpWindow.hWindow = Form3.hwnd Then
Me.SetFocus
Exit Sub
End If
If lpWindow.hWindow <> Me.hWinString.Caption Then
bCodeUse = True
Command3.Enabled = True
With Me.Check1
.Enabled = True
.Value = 1
End With
With Me.Check2
.Enabled = True
.Value = 0
End With
With Me.Check3
.Enabled = True
.Value = 0
End With
With Me.HScroll1
.Enabled = False
.Max = 255
.Min = 0
End With
Me.Label9.Enabled = False
Me.Label8.Enabled = False
With Me.Check4
.Enabled = True
.Value = 0
End With
With Me.Check6
.Enabled = True
.Value = 0
End With
With Me.Check7
.Enabled = True
.Value = 0
End With
With Me.Check8
.Enabled = True
.Value = 0
End With
With Me.Check9
.Enabled = True
.Value = 0
End With
With Me.Check10
.Enabled = True
.Value = 0
End With
With Me.Check11
.Enabled = True
.Value = 0
End With
With Me.Check12
.Enabled = True
.Value = 0
End With
bCodeUse = False
End If
If lpWindow.hWindow <> Me.hwnd Then
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
Me.Check1.Enabled = True
Command1.Enabled = True
Check8.Enabled = True
Me.Check2.Enabled = True
Me.Check3.Enabled = True
Me.Check4.Enabled = True
Me.Check6.Enabled = True
Check7.Enabled = True
Me.Check10.Enabled = True
Me.Check11.Enabled = True
Me.Check12.Enabled = True
Me.Check9.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
DoEvents
Else
DoEvents
Exit Sub
End If
End Sub
Private Sub Form_Resize()
On Error Resume Next
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
If Me.WindowState = 1 Then
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Enabled = False
.Interval = 1000
End With
Me.Hide
With Me.cSysTray1
.InTray = True
.TrayTip = "Win Tool - 双击还原窗口"
End With
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Visible = False
.Height = 7335
.Width = 8100
End With
Else
With Me.cSysTray1
.InTray = False
.TrayTip = "Win Tool - 双击还原窗口"
End With
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End If
End Sub
Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
On Error Resume Next
If Button = 1 Then
With Me.MouseHook
.Enabled = True
.Interval = 1000
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Me.Show
Form1.Show
With Me.cSysTray1
.InTray = False
.TrayTip = "Win Tool - 双击还原窗口"
End With
Form1.WindowState = 0
Form1.Show
With Form1
.WindowState = 0
.Show
.Visible = True
End With
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
End If
End Sub
Private Sub Check2_DragDrop(Source As Control, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check2_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check2_GotFocus()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check2_KeyDown(KeyCode As Integer, Shift As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check2_KeyPress(KeyAscii As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check2_KeyUp(KeyCode As Integer, Shift As Integer)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check2_LostFocus()
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
Private Sub Check2_OLECompleteDrag(Effect As Long)
Exit Sub
On Error Resume Next
bCodeUse = False
If List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To List1.ListIndex
Dim lpszListData As String
lpszListData = List1.List(nTmp)
If Trim(lpszListData) = "" Then
List1.RemoveItem nTmp
End If
Next
End If
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
With Me.MouseHook
.Interval = 1000
.Enabled = True
Debug.Print .Enabled
End With
Select Case Me.mnuEnable.Checked
Case True
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = True
End With
With Me.mnuEnable
.Checked = True
.Enabled = False
End With
With Me.mnuDisable
.Checked = False
.Enabled = True
End With
Case False
On Error Resume Next
With Me.MouseHook
.Interval = 1000
.Enabled = False
End With
With Me.mnuEnable
.Checked = False
.Enabled = True
End With
With Me.mnuDisable
.Checked = True
.Enabled = False
End With
End Select
Select Case mnuTop.Checked
Case False
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
Case True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Height = 7335
.Width = 8100
End With
End Select
Dim rtn As Long
Select Case mnuTrans.Checked
Case False
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Case True
On Error Resume Next
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
End Select
End Sub
