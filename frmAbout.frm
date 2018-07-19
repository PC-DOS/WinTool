VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   300
      Left            =   5325
      TabIndex        =   5
      Top             =   4695
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   1140
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmAbout.frx":0000
      Top             =   2730
      Width           =   5340
   End
   Begin VB.Image Image3 
      Height          =   75
      Index           =   1
      Left            =   0
      Picture         =   "frmAbout.frx":0133
      Stretch         =   -1  'True
      Top             =   1050
      Width           =   6600
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   6360
      Top             =   0
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   270
      Picture         =   "frmAbout.frx":059F
      Top             =   2085
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "基于Microsoft(R) Visual Studio(R) 6.00 制作"
      Height          =   180
      Index           =   3
      Left            =   1140
      TabIndex        =   3
      Top             =   2370
      Width           =   3870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PC-DOS Workshop开发"
      Height          =   180
      Index           =   2
      Left            =   1140
      TabIndex        =   2
      Top             =   1995
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "版本1.00"
      Height          =   180
      Index           =   1
      Left            =   1140
      TabIndex        =   1
      Top             =   1635
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   0
      Top             =   1275
      Width           =   90
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   270
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   0
      Picture         =   "frmAbout.frx":0E69
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   6480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Unload Me
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
Private Sub Form_Activate()
On Error Resume Next
Me.Command1.SetFocus
End Sub
Private Sub Form_Load()
On Error Resume Next
With Me.Image2(0)
.Stretch = True
.Picture = Form1.Icon
End With
With Label1(0)
.AutoSize = True
.BackStyle = 0
.BorderStyle = 0
.Caption = App.Title
End With
With Label1(1)
.AutoSize = True
.BackStyle = 0
.BorderStyle = 0
.Caption = "版本" & App.Major & "." & App.Minor & App.Revision
End With
With Me.Text1
.Locked = True
End With
With Me
.Left = Screen.Width / 2 - .Width / 2
.Top = Screen.Height / 2 - .Height / 2
.Icon = LoadPicture()
End With
With Me.Command1
.Cancel = True
.Default = True
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
