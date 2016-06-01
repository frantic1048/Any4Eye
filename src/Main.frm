VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm_Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Any 4 Eye"
   ClientHeight    =   1065
   ClientLeft      =   6765
   ClientTop       =   2865
   ClientWidth     =   2970
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   2970
   Begin VB.Timer tmrRemindFlicker 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Tag             =   "-1"
      Top             =   240
   End
   Begin VB.Timer tmrRemind 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2280
      Top             =   720
   End
   Begin VB.Timer tmrShutDown 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3000
      Tag             =   "-1"
      Top             =   240
   End
   Begin VB.Timer tmr_Chk 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3480
      Top             =   240
   End
   Begin VB.Image imgHelp 
      Height          =   615
      Left            =   3000
      Picture         =   "Main.frx":1F82
      Stretch         =   -1  'True
      ToolTipText     =   "帮助与说明"
      Top             =   240
      Width           =   615
   End
   Begin MSForms.Label lblMore 
      Height          =   855
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   375
      ForeColor       =   8438015
      BackColor       =   0
      Caption         =   "   >    >"
      Size            =   "661;1508"
      BorderStyle     =   1
      FontName        =   "黑体"
      FontEffects     =   1073741825
      FontHeight      =   210
      FontCharSet     =   0
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdLight 
      Height          =   855
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   1695
      ForeColor       =   8454016
      VariousPropertyBits=   19
      Caption         =   "开灯"
      PicturePosition =   327683
      Size            =   "2990;1508"
      Picture         =   "Main.frx":2BC6
      FontName        =   "黑体"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdDelight 
      Default         =   -1  'True
      Height          =   840
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2385
      ForeColor       =   8454016
      BackColor       =   4210752
      VariousPropertyBits=   268435483
      Caption         =   "       关灯"
      PicturePosition =   327683
      Size            =   "4207;1482"
      Picture         =   "Main.frx":5888
      FontName        =   "黑体"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CheckBox ChkAutoShutOff 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
      BackColor       =   0
      ForeColor       =   12648384
      DisplayStyle    =   4
      Size            =   "1931;661"
      Value           =   "0"
      Caption         =   "自动关机"
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.CheckBox Chk_AutoRemind 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
      BackColor       =   0
      ForeColor       =   12648384
      DisplayStyle    =   4
      Size            =   "1931;661"
      Value           =   "0"
      Caption         =   "自动提醒"
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.CheckBox chkChangeSysClr 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
      BackColor       =   0
      ForeColor       =   12648384
      DisplayStyle    =   4
      Size            =   "2566;661"
      Value           =   "0"
      Caption         =   "改变系统颜色"
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.ScrollBar ScoBar 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
      ForeColor       =   8454016
      BackColor       =   0
      Size            =   "873;1508"
      Min             =   100
      Max             =   235
      Position        =   150
      LargeChange     =   10
      Delay           =   40
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Const frmOHeight As Integer = 1485
Private Const frmOWidth As Integer = 3060
Private Const frmNHeight As Integer = 1890
Private Const frmNWidth As Integer = 3855
Private frmMaximum As Boolean

Private Sub Chk_AutoRemind_Click()
    If Chk_AutoRemind.value = True Then
        tmrRemind.Tag = -1
        tmrRemind.Enabled = True
    Else
        tmrRemind.Enabled = False
    End If
End Sub

Private Sub ChkAutoShutOff_Click()
    Dim tmpint
    If ChkAutoShutOff.value = True Then
        tmpint = MsgBox("是否为延时关机(倒计时关机),否则为定时关机", vbYesNo, "关机方式选择")
        If tmpint = vbYes Then
            DelayShutDown = True
            frmTimeDialog.Show
            frmTimeDialog.MaskEdBoxTime.Mask = "####"
            frmTimeDialog.Caption = "请输入关机时间,单位为分钟"
        ElseIf tmpint = vbNo Then
            DelayShutDown = False
            frmTimeDialog.Show
            frmTimeDialog.MaskEdBoxTime.Mask = "9#:9#"
            frmTimeDialog.Caption = "请输入关机时间,(二十四时制)"
        End If
    ElseIf ChkAutoShutOff.value = False Then
        tmrShutDown.Enabled = False
        frm_Main.Caption = "Any 4 Eye"
    End If
End Sub

Private Sub chkChangeSysClr_Click()
    Dim tempans
    If CSetInitialized = False Then
        Call CSetInitialize '初始化颜色设置
    End If
    If chkChangeSysClr.value = True Then
        If CheckThemesStatus = True Then
            tempans = MsgBox("是否切换为Windows经典样式,该样式下颜色设置将更加有效", vbYesNo, "样式切换")
            If tempans = vbYes Then StopThemesService
        End If
        SetSysColors 23, ColorCategories(1), NewColor(1) ' 使用设置的系统颜色
    Else
        SetSysColors 23, ColorCategories(1), OriginalColor(1) ' 改回原来系统颜色
        If (OriginalThemesStatus = CheckThemesStatus) = False Then StartThemesService
    End If
End Sub

Private Sub cmdLight_Click()
    Frm_BG.Hide
    ScoBar.Visible = False
    cmdLight.Visible = False
    cmdDelight.Visible = True
    tmr_Chk.Enabled = False
End Sub

Private Sub cmdDelight_Click()
    Frm_BG.Show
    ScoBar.Visible = True
    cmdLight.Visible = True
    cmdDelight.Visible = False
    tmr_Chk.Enabled = True
End Sub

Private Sub Form_Initialize()
    ScoBar.Visible = False
    cmdLight.Visible = False
    OriginalThemesStatus = CheckThemesStatus
End Sub

Private Sub Form_Load()
    preWinProc = GetWindowLong(Me.hWnd, GWL_WNDPROC)
    SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf WndProc
    RegisterHotKey Me.hWnd, 1, MOD_ALT, vbKeyN '装载时注册热键
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetWindowLong Me.hWnd, GWL_WNDPROC, preWinProc
    UnregisterHotKey Me.hWnd, 1
    SetSysColors 23, ColorCategories(1), OriginalColor(1) ' 改回系统颜色
    If (OriginalThemesStatus = CheckThemesStatus) = False Then StartThemesService
    End
End Sub

Private Sub imgHelp_Click()
    frmHelp.Show 1
End Sub

Private Sub lblMore_Click()
    Dim I As Integer, j As Integer
    If frmMaximum = False Then
        For I = frmOWidth To frmNWidth Step 2
            frm_Main.Width = I
        Next I
        For j = frmOHeight To frmNHeight Step 2
            frm_Main.Height = j
        Next j
        frmMaximum = True
    Else
        For j = frmNHeight To frmOHeight Step -2
            frm_Main.Height = j
        Next j
        For I = frmNWidth To frmOWidth Step -2
            frm_Main.Width = I
        Next I
        frmMaximum = False
    End If
End Sub

Private Sub ScoBar_Change()
    WindowTransparent HBG, CByte(ScoBar.value) ' 实时调整
End Sub

Private Sub Tmr_Chk_Timer()
    HFWnd = GetForegroundWindow '先获取前台活动窗口句柄
    'Debug.Print
    'Debug.Print "HFWnd = "; HFWnd '
    'Debug.Print "tyLastHFWnd = "; LastHFWnd
    'Debug.Print "Judge = "; Judge(LastHFWnd, HFWnd)
    If Judge(LastHFWnd, HFWnd) = True Then
        SetWindowPos HBG, GetForegroundWindow, 0, 0, 0, 0, (SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
    End If
    LastHFWnd = HFWnd
End Sub

Private Sub tmrRemind_Timer()
        tmrRemind.Tag = tmrRemind.Tag + 1
        If tmrRemind.Tag = 30 Then
            tmrRemind.Tag = -1
            tmrRemindFlicker.Tag = -1
            tmrRemindFlicker.Enabled = True
        End If
End Sub


Private Sub tmrRemindFlicker_Timer()
    tmrRemindFlicker.Tag = tmrRemindFlicker.Tag + 1
    WindowShowHide frmRemind
    If tmrRemindFlicker.Tag = 10 Then
        Unload frmRemind
        tmrRemindFlicker.Enabled = False
    End If
End Sub

Private Sub tmrShutDown_Timer()
    Select Case DelayShutDown
    Case True
        tmrShutDown.Tag = tmrShutDown.Tag + 1
        frm_Main.Caption = "Any 4 Eye _" & (ShutDownTime - tmrShutDown.Tag) & "分钟后关机"
        If (ShutDownTime - tmrShutDown.Tag) = 5 Then
            MsgBox "还有5分钟关机,请注意及时保存数据或取消关机", vbOKOnly, "提示"
            Exit Sub
        End If
        If tmrShutDown.Tag = ShutDownTime Then ShutDown
    Case False
        If ShutDownTime = FormatDateTime(Time, vbShortTime) Then
            ShutDown
        End If
    End Select
End Sub
