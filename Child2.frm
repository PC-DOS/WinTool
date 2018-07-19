VERSION 5.00
Begin VB.Form Child2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Window Settings"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "Child2.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   345
      Left            =   3360
      TabIndex        =   17
      Top             =   4455
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "选项"
      Height          =   3600
      Left            =   30
      TabIndex        =   0
      Top             =   810
      Width           =   4665
      Begin VB.CheckBox Check9 
         Caption         =   "不允许关闭(&N)"
         Height          =   345
         Left            =   150
         TabIndex        =   14
         Top             =   2385
         Width           =   1500
      End
      Begin VB.CheckBox Check10 
         Caption         =   "不允许最大化(&T)"
         Height          =   345
         Left            =   150
         TabIndex        =   13
         Top             =   2670
         Width           =   1650
      End
      Begin VB.CheckBox Check11 
         Caption         =   "不允许最小化(&O)"
         Height          =   345
         Left            =   150
         TabIndex        =   12
         Top             =   2955
         Width           =   1650
      End
      Begin VB.CheckBox Check12 
         Caption         =   "删除标题栏(&D)"
         Height          =   300
         Left            =   150
         TabIndex        =   11
         Top             =   3270
         Width           =   1560
      End
      Begin VB.CheckBox Check1 
         Caption         =   "该窗口有效并允许用户与之交互(&E)"
         Height          =   300
         Left            =   150
         TabIndex        =   8
         Top             =   180
         Value           =   1  'Checked
         Width           =   4035
      End
      Begin VB.CheckBox Check2 
         Caption         =   "该窗口始终保持在所有窗口之上(&K)"
         Height          =   300
         Left            =   150
         TabIndex        =   7
         Top             =   435
         Width           =   4035
      End
      Begin VB.CheckBox Check3 
         Caption         =   "该窗口透明(&T)"
         Height          =   300
         Left            =   150
         TabIndex        =   6
         Top             =   690
         Width           =   4035
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   10
         Left            =   420
         Max             =   255
         SmallChange     =   5
         TabIndex        =   5
         Top             =   1005
         Width           =   4155
      End
      Begin VB.CheckBox Check4 
         Caption         =   "该窗口被隐藏(&H)"
         Height          =   300
         Left            =   150
         TabIndex        =   4
         Top             =   1305
         Width           =   4035
      End
      Begin VB.CheckBox Check6 
         Caption         =   "该窗口不允许被最小化/最大化/关闭(&A)"
         Height          =   300
         Left            =   150
         TabIndex        =   3
         Top             =   1575
         Width           =   4035
      End
      Begin VB.CheckBox Check7 
         Caption         =   "该窗口不允许被调整大小(&S)"
         Height          =   300
         Left            =   150
         TabIndex        =   2
         Top             =   1845
         Width           =   4035
      End
      Begin VB.CheckBox Check8 
         Caption         =   "该窗口不允许移动(&M)"
         Enabled         =   0   'False
         Height          =   270
         Left            =   150
         TabIndex        =   1
         Top             =   2145
         Width           =   4410
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数值:"
         Height          =   180
         Left            =   435
         TabIndex        =   10
         Top             =   1335
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "255"
         Height          =   240
         Left            =   930
         TabIndex        =   9
         Top             =   1305
         Visible         =   0   'False
         Width           =   3645
      End
   End
   Begin VB.Label hWinString 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   30
      TabIndex        =   16
      Top             =   285
      Width           =   4650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前选中窗口的句柄:"
      Height          =   180
      Left            =   30
      TabIndex        =   15
      Top             =   45
      Width           =   1710
   End
End
Attribute VB_Name = "Child2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOOWNERZORDER = &H200
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
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
On Error Resume Next
If Form1.List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To Form1.List1.ListIndex
Dim lpszListData As String
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
End If
Next
End If
If bCodeUse = True Then
Exit Sub
End If
Dim lpReturn As Long
Select Case Check1.Value
Case 0
lpReturn = EnableWindow(hWinString.Caption, False)
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
lpReturn = EnableWindow(hWinString.Caption, True)
If lpReturn <> 0 Then
Exit Sub
Else
If 1 = 245 Then
MsgBox "发生错误,无法设置窗口有效性", vbCritical, "Error"
End If
End If
End Select
End Sub
Private Sub Check10_Click()
On Error Resume Next
If Form1.List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To Form1.List1.ListIndex
Dim lpszListData As String
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
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
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
If Form1.List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To Form1.List1.ListIndex
Dim lpszListData As String
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
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
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
If Form1.List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To Form1.List1.ListIndex
Dim lpszListData As String
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
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
SetWindowPos hWinString.Caption, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Case 1
SetWindowPos hWinString.Caption, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Select
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
SetWindowPos hWinString.Caption, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Case 1
SetWindowPos hWinString.Caption, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Select
End Select
End Sub
Private Sub Check2_Click()
On Error Resume Next
If Form1.List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To Form1.List1.ListIndex
Dim lpszListData As String
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
End If
Next
End If
On Error Resume Next
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
lpReturn = SetWindowPos(hWinString.Caption, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
lpReturn = SetWindowPos(hWinString.Caption, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
If lpReturn <> 0 Then
Exit Sub
Else
If 1 = 245 Then
MsgBox "发生错误,无法设置窗口位置", vbCritical, "Error"
End If
End If
End Select
Exit Sub
End Sub
Private Sub Check3_Click()
If Form1.List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To Form1.List1.ListIndex
Dim lpszListData As String
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
End If
Next
End If
If Form1.List1.ListCount >= 2 Then
For nTmp = 0 To Form1.List1.ListIndex
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
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
lpReturn = GetWindowLong(hWinString.Caption, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong hWinString.Caption, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes hWinString.Caption, 0, 255, LWA_ALPHA
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
lpReturn = GetWindowLong(hWinString.Caption, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong hWinString.Caption, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes hWinString.Caption, 0, HScroll1.Value, LWA_ALPHA
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
End Select
Exit Sub
End Sub
Private Sub Check4_Click()
On Error Resume Next
If Form1.List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To Form1.List1.ListIndex
Dim lpszListData As String
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
End If
Next
End If
If bCodeUse = True Then
Exit Sub
End If
Dim lpReturn As Long
Select Case Check4.Value
Case 0
Dim lpLong As Long
lpLong = GetWindowLong(Me.hWinString.Caption, GWL_STYLE)
If 1 = 245 Then
lpLong = lpLong Or WS_ICONIC
SetWindowLong Me.hWinString.Caption, GWL_STYLE, lpLong
End If
If Check3.Value = 1 Then
GetSystemMenu Me.hWinString.Caption, True
lpReturn = GetWindowLong(hWinString.Caption, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong hWinString.Caption, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes hWinString.Caption, 0, HScroll1.Value, LWA_ALPHA
Const SWP_NOZORDER = &H4
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
SetWindowPos hWinString.Caption, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
With Me.Label9
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = True
End With
Else
lpReturn = GetWindowLong(hWinString.Caption, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong hWinString.Caption, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes hWinString.Caption, 0, 255, LWA_ALPHA
SetWindowPos hWinString.Caption, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
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
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
End If
End If
Dim cunt As Long
Dim addvar As Long
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
End If
End If
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
lpReturn = GetWindowLong(hWinString.Caption, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong hWinString.Caption, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes hWinString.Caption, 0, 0, LWA_ALPHA
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
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
End If
End If
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
End If
Next
End Select
End Sub
Private Sub Check6_Click()
On Error Resume Next
If Form1.List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To Form1.List1.ListIndex
Dim lpszListData As String
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
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
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
End If
End If
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
End If
End If
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
If Form1.List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To Form1.List1.ListIndex
Dim lpszListData As String
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
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
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
If Form1.List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To Form1.List1.ListIndex
Dim lpszListData As String
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
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
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
If Form1.List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To Form1.List1.ListIndex
Dim lpszListData As String
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
End If
Next
End If
If bCodeUse = True Then
Exit Sub
End If
Select Case Check9.Value
Case 0
bCodeUse = False
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
bCodeUse = False
Case 1
bCodeUse = False
If Me.Check8.Value = 1 Then
RemoveMenu GetSystemMenu(hWinString.Caption, 0), SC_MOVE, MF_REMOVE
End If
If 1 = 245 Then
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
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
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
Exit Sub
End If
End If
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
Exit Sub
End If
Next
bCodeUse = False
End Select
End Sub
Private Sub Command1_Click()
On Error Resume Next
On Error Resume Next
Form5.List1.Clear
Form5.lpszCaption.Caption = ""
Form5.lpszClass.Caption = ""
Form5.lpszThread.Caption = ""
Form5.hWinString.Caption = ""
Form5.hDCString.Caption = ""
Form5.Command2.Enabled = False
Form5.Command3.Enabled = False
Form5.Command4.Enabled = False
Form5.Command5.Enabled = False
Form5.Command7.Enabled = False
With Form5.List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf FindProcessesWithChildWindows, 0
Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo ep
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2
bCodeUse = False
With Me.Command1
.Enabled = True
.Default = True
.Cancel = True
End With
With Me.HScroll1
.Min = 0
.Max = 255
.LargeChange = 10
.SmallChange = 5
.Enabled = False
.Visible = True
.Value = 255
End With
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
If 1 = 2 Then
With Me.hWinString
.Alignment = 2
.BackStyle = 0
.BorderStyle = 1
.Caption = ""
.Visible = True
.Enabled = True
End With
End If
With Me.Label8
.Visible = False
.Enabled = False
.Height = 0
.Width = 0
.Left = 0
.Top = 0
End With
With Me.Label9
.Visible = False
.Enabled = False
.Height = 0
.Width = 0
.Left = 0
.Top = 0
End With
If 1 = 2 Then
With Me.hWinString
.Caption = Form5.hWinString.Caption
End With
End If
With Me.Check1
.Value = 1
.Enabled = True
End With
With Me.Check2
.Value = 0
.Enabled = True
End With
With Me.Check3
.Value = 0
.Enabled = True
End With
With Me.Check4
.Value = 0
.Enabled = True
End With
With Me.Check6
.Value = 0
.Enabled = True
End With
With Me.Check7
.Value = 0
.Enabled = True
End With
With Me.Check8
.Value = 0
.Enabled = True
End With
With Me.Check9
.Value = 0
.Enabled = True
End With
With Me.Check10
.Value = 0
.Enabled = True
End With
With Me.Check11
.Value = 0
.Enabled = True
End With
With Me.Check12
.Value = 0
.Enabled = True
End With
hWinString.Caption = Form5.hWinString.Caption
Exit Sub
ep:
SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
MsgBox Err.Description, vbCritical, "Error"
Unload Me
Exit Sub
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
On Error Resume Next
Form5.List1.Clear
Form5.Command7.Enabled = False
Form5.lpszCaption.Caption = ""
Form5.lpszClass.Caption = ""
Form5.lpszThread.Caption = ""
Form5.hWinString.Caption = ""
Form5.hDCString.Caption = ""
Form5.Command2.Enabled = False
Form5.Command3.Enabled = False
Form5.Command4.Enabled = False
Form5.Command5.Enabled = False
With Form5.List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf FindProcessesWithChildWindows, 0
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
On Error Resume Next
Form5.List1.Clear
Form5.Command7.Enabled = False
Form5.lpszCaption.Caption = ""
Form5.lpszClass.Caption = ""
Form5.lpszThread.Caption = ""
Form5.hWinString.Caption = ""
Form5.hDCString.Caption = ""
Form5.Command2.Enabled = False
Form5.Command3.Enabled = False
Form5.Command4.Enabled = False
Form5.Command5.Enabled = False
With Form5.List1
.Clear
.ListIndex = -1
End With
EnumWindows AddressOf FindProcessesWithChildWindows, 0
Unload Me
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
If Form1.List1.ListCount >= 2 Then
Dim nTmp As Long
For nTmp = 0 To Form1.List1.ListIndex
Dim lpszListData As String
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
End If
Next
End If
On Error Resume Next
Dim lpReturn As Long
lpReturn = GetWindowLong(hWinString.Caption, GWL_EXSTYLE)
lpReturn = lpReturn Or WS_EX_LAYERED
SetWindowLong hWinString.Caption, GWL_EXSTYLE, lpReturn
SetLayeredWindowAttributes hWinString.Caption, 0, HScroll1.Value, LWA_ALPHA
With Me.Label9
.Alignment = 2
.Caption = HScroll1.Value
.Enabled = True
End With
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
If Form1.List1.ListCount >= 2 Then
For nTmp = 0 To Form1.List1.ListIndex
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
End If
Next
End If
If Form1.List1.ListCount >= 2 Then
For nTmp = 0 To Form1.List1.ListIndex
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
End If
Next
End If
If Form1.List1.ListCount >= 2 Then
For nTmp = 0 To Form1.List1.ListIndex
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
End If
Next
End If
If Form1.List1.ListCount >= 2 Then
For nTmp = 0 To Form1.List1.ListIndex
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
End If
Next
End If
If Form1.List1.List(0) = "" Then
Form1.List1.RemoveItem 0
End If
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
If Form1.List1.List(0) = "" Then
Form1.List1.RemoveItem 0
End If
Exit Sub
End If
End If
If Form1.List1.List(0) = "" Then
Form1.List1.RemoveItem 0
End If
Dim addvar As Integer
Dim cunt As Integer
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
If Form1.List1.ListCount >= 2 Then
For nTmp = 0 To Form1.List1.ListIndex
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
End If
Next
End If
If Form1.List1.List(0) = "" Then
Form1.List1.RemoveItem 0
End If
Exit Sub
End If
Next
If HScroll1.Value = 0 Then
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
If Form1.List1.ListCount >= 2 Then
For nTmp = 0 To Form1.List1.ListIndex
lpszListData = Form1.List1.List(nTmp)
If Trim(lpszListData) = "" Then
Form1.List1.RemoveItem nTmp
End If
Next
End If
If Form1.List1.List(0) = "" Then
Form1.List1.RemoveItem 0
End If
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
If Form1.List1.List(0) = "" Then
Form1.List1.RemoveItem 0
End If
Exit Sub
End If
End If
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
If Form1.List1.List(0) = "" Then
Form1.List1.RemoveItem 0
End If
Exit Sub
End If
Next
If HScroll1.Value = 0 Then
If Form1.List1.ListCount = 0 Then
Form1.List1.AddItem Me.hWinString.Caption
If Form1.List1.List(0) = "" Then
Form1.List1.RemoveItem 0
End If
Exit Sub
End If
If Form1.List1.ListCount = 1 Then
If Me.hWinString.Caption <> Form1.List1.List(0) Then
Form1.List1.AddItem Me.hWinString.Caption
If Form1.List1.List(0) = "" Then
Form1.List1.RemoveItem 0
End If
Exit Sub
End If
End If
For cunt = 0 To Form1.List1.ListCount - 1
If Me.hWinString.Caption <> Form1.List1.List(cunt) Then
addvar = addvar + 1
If addvar = Form1.List1.ListCount Then
Form1.List1.AddItem Me.hWinString.Caption
addvar = 0
End If
Else
If Form1.List1.List(0) = "" Then
Form1.List1.RemoveItem 0
End If
Exit Sub
End If
Next
End If
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
If Form5.lpszClass.Caption = "" Then
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Exit Sub
End If
With dwWinInfo
.lpszCaption = Form5.lpszCaption.Caption
.lpszClass = Form5.lpszClass.Caption
.lpszDC = Form5.hDCString.Caption
.lpszHandle = Form5.hWinString.Caption
.lpszThread = Form5.lpszThread.Caption
End With
With dwWinInfo
.lpszCaption = Form5.lpszCaption.Caption
.lpszClass = Form5.lpszClass.Caption
.lpszDC = Form5.hDCString.Caption
.lpszHandle = Form5.hWinString.Caption
.lpszThread = Form5.lpszThread.Caption
End With
If Form5.lpszClass.Caption <> "" Then
With dwWinInfo
MsgBox INFO_CAPTION & vbCrLf & .lpszCaption & vbCrLf & INFO_HANDLE & vbCrLf & .lpszHandle & vbCrLf & INFO_CLASS & vbCrLf & .lpszClass & vbCrLf & INFO_DC & vbCrLf & .lpszDC & vbCrLf & INFO_PROCESS & vbCrLf & .lpszThread, vbInformation, "Info"
End With
End If
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub
