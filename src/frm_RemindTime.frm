VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm_RemindTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置提醒时间(以分钟为单位)"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_RemindTime.frx":0000
   ScaleHeight     =   795
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtRemindTime 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   495
      Left            =   2785
      TabIndex        =   2
      Top             =   156
      Width           =   855
      VariousPropertyBits=   19
      Caption         =   "确定"
      Size            =   "1508;873"
      FontName        =   "微软雅黑"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   495
      Left            =   1718
      TabIndex        =   1
      Top             =   156
      Width           =   855
      VariousPropertyBits=   19
      Caption         =   "取消"
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
Attribute VB_Name = "frm_RemindTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frm_Main.Chk_AutoRemind.value = False
    frm_Main.Chk_AutoRemind.Enabled = True
    frm_Main.tmrRemind.Enabled = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsNumeric(Me.txtRemindTime.Text) Then
        frm_Main.tmrRemind.Tag = 60 * CInt(Val(Me.txtRemindTime.Text))
        tmpRemindTime = CInt(Val(Me.txtRemindTime.Text))
        frm_Main.tmrRemind.Enabled = True
        frm_Main.Chk_AutoRemind.Enabled = True
        Unload Me
    Else
        MsgBox "请输入正确的时间"
        txtRemindTime.Text = ""
    End If
End Sub

