VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTimeDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�Ի���"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3765
   ControlBox      =   0   'False
   Icon            =   "frmTimeDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   120
   End
   Begin MSMask.MaskEdBox MaskEdBoxTime 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSForms.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      Caption         =   "OK"
      PicturePosition =   327683
      Size            =   "1931;1085"
      FontName        =   "΢���ź�"
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
Private Sub cmdOK_Click()
    If DelayShutDown = False Then
        If IsDate(MaskEdBoxTime.Text) = True Then
            ShutDownTime = FormatDateTime(MaskEdBoxTime.Text, vbShortTime)
            MsgBox "����" & ShutDownTime & "�رռ����", vbOKOnly
            frm_Main.Caption = frm_Main.Caption & "_" & ShutDownTime & "�ػ�"
            frm_Main.tmrShutDown.Enabled = True
            Unload frmTimeDialog
        Else
            MsgBox "��������ȷʱ��"
        End If
    ElseIf DelayShutDown = True Then
        If MaskEdBoxTime.Text = "" Then
            MsgBox "����������", vbOKOnly
        Else
            ShutDownTime = Val(MaskEdBoxTime.Text)
            MsgBox "����" & ShutDownTime & "����֮��ػ�", vbOKOnly
            frm_Main.Caption = frm_Main.Caption & "_" & ShutDownTime & "���Ӻ�ػ�"
            frm_Main.tmrShutDown.Tag = 0
            frm_Main.tmrShutDown.Enabled = True
            Unload frmTimeDialog
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    If DelayShutDown = False Then
            Me.Caption = "������ػ�ʱ��,(��ʮ��ʱ��)" & "Ŀǰʱ��:" & FormatDateTime(Time, vbShortTime)
    End If
End Sub
