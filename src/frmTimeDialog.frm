VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTimeDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "对话框"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3810
   ControlBox      =   0   'False
   Icon            =   "frmTimeDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTimeDialog.frx":21E62
   ScaleHeight     =   765
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSMask.MaskEdBox MaskEdBoxTime 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1200
      Top             =   0
   End
   Begin MSForms.CommandButton cmdCancel 
      Default         =   -1  'True
      Height          =   495
      Left            =   1718
      TabIndex        =   2
      Top             =   145
      Width           =   855
      VariousPropertyBits=   19
      Caption         =   "取消"
      PicturePosition =   262148
      Size            =   "1508;873"
      FontName        =   "微软雅黑"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOK 
      Height          =   495
      Left            =   2785
      TabIndex        =   1
      Top             =   145
      Width           =   855
      VariousPropertyBits=   19
      Caption         =   "确定"
      PicturePosition =   262148
      Size            =   "1508;873"
      FontName        =   "微软雅黑"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmTimeDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frm_Main.ChkAutoShutOff.value = False
    frm_Main.ChkAutoShutOff.Enabled = True
    Unload frmTimeDialog
End Sub

Private Sub cmdOK_Click()
    If DelayShutDown = False Then
        If IsDate(MaskEdBoxTime.Text) = True Then
            ShutDownTime = FormatDateTime(MaskEdBoxTime.Text, vbShortTime)
            MsgBox "将在" & ShutDownTime & "关闭计算机", vbOKOnly
            frm_Main.lbl_Remind.Caption = "将在" & ShutDownTime & "关机"
            frm_Main.tmrShutDown.Enabled = True
            frm_Main.ChkAutoShutOff.Enabled = True
            Unload frmTimeDialog
        Else
            MsgBox "请输入正确时间"
        End If
    ElseIf DelayShutDown = True Then
        If MaskEdBoxTime.Text = "" Then
            MsgBox "请输入数字", vbOKOnly
        Else
            ShutDownTime = Val(MaskEdBoxTime.Text)
            MsgBox "将在" & ShutDownTime & "分钟之后关机", vbOKOnly
            frm_Main.lbl_Remind.Caption = ShutDownTime & "分钟后关机"
            frm_Main.tmrShutDown.Tag = 0
            frm_Main.tmrShutDown.Enabled = True
            frm_Main.ChkAutoShutOff.Enabled = True
            Unload frmTimeDialog
        End If
    End If
End Sub

Private Sub Form_Load()
    frm_Main.ChkAutoShutOff.Enabled = False
    With Me
        .Left = (Screen.Width - .Width) * 0.5
        .Top = (Screen.Height - .Height) * 0.5
    End With
    AnimateWindow frmTimeDialog.hwnd, 1000, AW_BLEND Or AW_ACTIVATE
    Me.Refresh
End Sub

Private Sub Timer1_Timer()
    If DelayShutDown = False Then
            Me.Caption = "请输入关机时间(二十四时制)" & "目前时间:" & FormatDateTime(Time, vbShortTime)
    End If
End Sub
