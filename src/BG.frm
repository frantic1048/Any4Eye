VERSION 5.00
Begin VB.Form frm_BG 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3825
   ClientLeft      =   510
   ClientTop       =   1035
   ClientWidth     =   3945
   Enabled         =   0   'False
   Icon            =   "BG.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrChk 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Frm_BG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private HFWnd As Long '(Hwnd of Last Foreground Window)
Private LastHFWnd As Long '(Hwnd of Last Foreground Window)


Private Declare Function SetWindowPos Lib "user32" (ByVal _
                        hwnd As Long, ByVal _
                        hWndInsertAfter As Long, ByVal _
                        X As Long, ByVal _
                        Y As Long, ByVal _
                        cx As Long, ByVal _
                        cy As Long, ByVal _
                        wFlags As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal _
                        hwnd As Long, ByVal _
                        nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal _
                        hwnd As Long, ByVal _
                        nIndex As Long, ByVal _
                        dwNewLong As Long) As Long ' 要设置扩展样式
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal _
                        hwnd As Long, ByVal _
                        crKey As Long, ByVal _
                        bAlpha As Byte, ByVal _
                        dwFlags As Long) As Long               ' 透明用
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal _
                        uAction As Long, ByVal _
                        uParam As Long, ByRef _
                        lpvParam As Any, ByVal _
                        fuWinIni As Long) As Long ' 得到工作区用
                        
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1 'SetWindowPos
Private Const GWL_EXSTYLE = (-20) '
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20& 'SetWindowLong
Private Const LWA_ALPHA As Long = &H2
Private Const SPI_GETWORKAREA = 48 '取得工作区大小,用以将Frm_BG覆盖除了任务栏之外的整个区域


Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type  'SystemParametersInfo的lpvParam参数在SPI_GETWkarea下的结构要求


Private Sub Frm_Maximum(frm As Form)
    Dim wkarea As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, wkarea, 0
    With frm
        .Left = wkarea.Left * Screen.TwipsPerPixelX
        .Top = wkarea.Top * Screen.TwipsPerPixelY
        .Width = (wkarea.Right - wkarea.Left) * Screen.TwipsPerPixelX
        .Height = (wkarea.Bottom - wkarea.Top) * Screen.TwipsPerPixelY
    End With
End Sub
Private Function Judge(ByVal Lhwnd, ByVal Phwnd) As Boolean
    If (Lhwnd <> Phwnd) And _
        (Phwnd <> 0) And _
        (Lhwnd <> 0) And _
        (Phwnd <> Frm_BG.hwnd) Then
        Judge = True
    Else
    Judge = False
    End If
End Function
Private Sub Form_Load()

    If App.PrevInstance = True Then
        MsgBox "程序已启动 (请勿自行运行BackGround.exe启动关灯功能)!", vbOKOnly, "警告"
        End
    End If
    
    
    SetWindowLong Frm_BG.hwnd, GWL_EXSTYLE, (GetWindowLong(Frm_BG.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED Or WS_EX_TRANSPARENT)
    SetLayeredWindowAttributes Frm_BG.hwnd, 0, 150, LWA_ALPHA
    Frm_Maximum Frm_BG ' 覆盖整个工作区
End Sub

Private Sub tmrChk_Timer()
    HFWnd = GetForegroundWindow '先获取前台活动窗口句柄
    Debug.Print
    Debug.Print "HFWnd = "; HFWnd '
    Debug.Print "tyLastHFWnd = "; LastHFWnd
    Debug.Print "Judge = "; Judge(LastHFWnd, HFWnd)
    If Judge(LastHFWnd, HFWnd) = True Then
        SetWindowPos Frm_BG.hwnd, GetForegroundWindow, 0, 0, 0, 0, (SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
    End If
    LastHFWnd = HFWnd
End Sub
