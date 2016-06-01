VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_PicRemind 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer tmr_AutoUnload 
      Interval        =   10000
      Left            =   7920
      Tag             =   "0"
      Top             =   3720
   End
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   2640
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   600
      ImageHeight     =   328
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_SDRemind.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_SDRemind.frx":90294
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frm_PicRemind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_Load()

     Select Case RemindCase
        Case 1
            Image1.Picture = ImgLst.ListImages(1).Picture
        Case 2
            Image1.Picture = ImgLst.ListImages(2).Picture
    End Select
    
    With Me
        .Left = (Screen.Width - .Width) * 0.5
        .Top = (Screen.Height - .Height) * 0.5
    End With
    
    AnimateWindow Me.hwnd, 1500, AW_BLEND
    Me.Refresh
    WindowTransparent Me.hwnd, , True
    
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragWindow Button, Me.hwnd
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragWindow Button, Me.hwnd
End Sub

Private Sub tmr_AutoUnload_Timer()

        Unload frm_PicRemind

End Sub
