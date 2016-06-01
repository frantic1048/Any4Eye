VERSION 5.00
Begin VB.Form frm_BG 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3825
   ClientLeft      =   510
   ClientTop       =   1035
   ClientWidth     =   3945
   Icon            =   "BackGround.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "Frm_BG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    HBG = Frm_BG.hWnd
    WindowTransparent HBG '启动时透明
    Frm_Maximum Frm_BG ' 覆盖整个工作区
End Sub
