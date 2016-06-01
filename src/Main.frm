VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm_Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Any 4 Eye"
   ClientHeight    =   2430
   ClientLeft      =   6720
   ClientTop       =   2490
   ClientWidth     =   4785
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":1F82
   ScaleHeight     =   2430
   ScaleWidth      =   4785
   StartUpPosition =   3  '窗口缺省
   Begin VB.VScrollBar ScoBar 
      Height          =   495
      LargeChange     =   5
      Left            =   750
      Max             =   240
      Min             =   80
      TabIndex        =   5
      Top             =   960
      Value           =   150
      Visible         =   0   'False
      Width           =   520
   End
   Begin VB.Timer tmrTimeShow 
      Interval        =   30000
      Left            =   2400
      Top             =   360
   End
   Begin VB.Timer tmrRemind 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1680
      Top             =   360
   End
   Begin VB.Timer tmrShutDown 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2040
      Tag             =   "-1"
      Top             =   360
   End
   Begin VB.Image Img_Delight 
      Height          =   900
      Left            =   480
      MouseIcon       =   "Main.frx":277C6
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":27918
      Top             =   840
      Width           =   3600
   End
   Begin VB.Image Img_Light 
      Height          =   825
      Left            =   1560
      MouseIcon       =   "Main.frx":3221C
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":3236E
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Label lbl_Remind 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lbl_Time 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Image ImgMini 
      Height          =   375
      Left            =   3960
      MouseIcon       =   "Main.frx":3A242
      MousePointer    =   99  'Custom
      ToolTipText     =   "最小化"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image ImgEnd 
      Height          =   375
      Left            =   4440
      MouseIcon       =   "Main.frx":3A394
      MousePointer    =   99  'Custom
      ToolTipText     =   "退出程序"
      Top             =   0
      Width           =   255
   End
   Begin MSForms.CheckBox chkChangeSysClr 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "将您的系统颜色替换为一套柔和的适宜夜间使用的颜色"
      Top             =   2040
      Width           =   1455
      VariousPropertyBits=   746588179
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
   Begin VB.Image imgHelp 
      Height          =   495
      Left            =   4200
      MouseIcon       =   "Main.frx":3A4E6
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":3A638
      Stretch         =   -1  'True
      ToolTipText     =   "帮助与说明"
      Top             =   1800
      Width           =   495
   End
   Begin MSForms.CheckBox ChkAutoShutOff 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
      VariousPropertyBits=   746588179
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
      Left            =   3000
      TabIndex        =   1
      ToolTipText     =   "每小时疲劳提醒"
      Top             =   2040
      Width           =   1095
      VariousPropertyBits=   746588179
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
   Begin VB.Image Img_Barback 
      Enabled         =   0   'False
      Height          =   795
      Left            =   600
      Picture         =   "Main.frx":3B27C
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   780
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1


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
        lbl_Remind.Caption = " "
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

Private Sub Form_Load()
    BGpath = App.Path & "\BackGround.exe"
    OriginalThemesStatus = CheckThemesStatus
    WindowTransparent frm_Main.hwnd, , True
    preWinProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf WndProc
    RegisterHotKey Me.hwnd, 1, MOD_ALT, vbKeyN '装载时注册热键
    
    
    lbl_Time.Caption = "当前时间 : " + FormatDateTime(Time, vbShortTime)

    Img_Delight_Click
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragWindow Button, Me.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetWindowLong Me.hwnd, GWL_WNDPROC, preWinProc
    UnregisterHotKey Me.hwnd, 1
    SetSysColors 23, ColorCategories(1), OriginalColor(1) ' 改回系统颜色
    If (OriginalThemesStatus = CheckThemesStatus) = False Then StartThemesService
    If DelightStatus Then Light
    End
End Sub

Private Sub Img_Delight_Click()
    Delight
End Sub

Private Sub Img_Light_Click()
    Light
End Sub



Private Sub ImgEnd_Click()
    Unload frm_Main
End Sub

Private Sub imgHelp_Click()
    frmHelp.Show
End Sub

Private Sub ImgMini_Click()
    Me.WindowState = 1
End Sub

Private Sub ScoBar_Change()
    SetLayeredWindowAttributes FindWindow(vbNullString, "A4EBKGROUND"), 0, CByte(ScoBar.value), LWA_ALPHA
 ' 实时调整
End Sub


Private Sub tmrRemind_Timer()
        tmrRemind.Tag = tmrRemind.Tag + 1
        If tmrRemind.Tag = 60 Then
            tmrRemind.Tag = -1
            RemindCase = 1
            frm_PicRemind.Show
        End If
End Sub

Private Sub tmrShutDown_Timer()
    Select Case DelayShutDown
    Case True
        tmrShutDown.Tag = tmrShutDown.Tag + 1
        lbl_Remind.Caption = (ShutDownTime - tmrShutDown.Tag) & "分钟后关机"
        If (ShutDownTime - tmrShutDown.Tag) = 5 Then
            RemindCase = 2
            frm_PicRemind.Show
            Exit Sub
        End If
        If tmrShutDown.Tag = ShutDownTime Then ShutDown
    Case False
        If ShutDownTime = FormatDateTime(Time, vbShortTime) Then
            ShutDown
        End If
    End Select
End Sub


Private Sub tmrTimeShow_Timer()
    lbl_Time.Caption = "当前时间 : " + FormatDateTime(Time, vbShortTime)
End Sub
