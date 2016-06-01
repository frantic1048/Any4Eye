Attribute VB_Name = "GeneralModule"
'注释中的括号为命名说明
Option Explicit
Option Base 1

Public HFWnd As Long '(Hwnd of Last Foreground Window)
Public LastHFWnd As Long '(Hwnd of Last Foreground Window)
Public HBG As Long '(Hwnd of Frm_BG)
Public NewColor(23) As Long
Public OriginalColor(23) As Long '用来储存原本的系统颜色(Value of Original SysColor)
Public ColorCategories(23) As Long '用来存放改变的颜色类别
Public CSetInitialized As Boolean '标记颜色记录是否初始化,(ColorSetInitialized)
Public DelayShutDown As Boolean '用来标记时间对话框的显示类型,true为延时输入,false为标准时间格式输入
Public ShutDownTime  '用来表示关机时间
Public OriginalThemesStatus As Boolean  ' 用来标记原本用户是否使用了Themes服务
Public preWinProc As Long '存储原本窗口过程的地址

'-----------------------------------------------------------------------------------[API Void ]

Public Declare Function GetForegroundWindow Lib "user32" () As Long '获取工作窗口的句柄
Public Declare Function SetWindowPos Lib "user32" (ByVal _
                        hWnd As Long, ByVal _
                        hWndInsertAfter As Long, ByVal _
                        x As Long, ByVal _
                        y As Long, ByVal _
                        cx As Long, ByVal _
                        cy As Long, ByVal _
                        wFlags As Long) As Long ' 达到Frm_BG永远在工作窗口下
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal _
                        hWnd As Long, ByVal _
                        nIndex As Long) As Long ' 获取扩展样式Extend Style
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long ' 获得系统颜色
Public Declare Function SetSysColors Lib "user32" (ByVal _
                        nChanges As Long, _
                        lpSysColor As Long, _
                        lpColorValues As Long) As Long ' 设置系统颜色
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal _
                        hWnd As Long, ByVal _
                        nIndex As Long, ByVal _
                        dwNewLong As Long) As Long ' 要设置扩展样式
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal _
                        hWnd As Long, ByVal _
                        crKey As Long, ByVal _
                        bAlpha As Byte, ByVal _
                        dwFlags As Long) As Long               ' 透明用
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal _
                        uAction As Long, ByVal _
                        uParam As Long, ByRef _
                        lpvParam As Any, ByVal _
                        fuWinIni As Long) As Long ' 得到工作区用
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, _
                        ByVal dwReserved As Long) As Long '  关机,但还需要权限
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long ' 获得当前进程句柄
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal _
                        DesiredAccess As Long, _
                        TokenHandle As Long) As Long ' 修改进程的访问令牌
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal _
                        lpSystemName As String, ByVal _
                        lpName As String, _
                        lpLuid As LUID) As Long ' 取得关机权限对应的本地唯一标示符
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal _
                        TokenHandle As Long, ByVal _
                        DisableAllPrivileges As Long, _
                        NewState As TOKEN_PRIVILEGES, ByVal _
                        BufferLength As Long, _
                        PreviousState As TOKEN_PRIVILEGES, _
                        ReturnLength As Long) As Long '在进程的访问令牌中启用关机权限
Private Declare Function OpenSCManager Lib "advapi32" Alias "OpenSCManagerA" (ByVal _
                        lpMachineName As String, ByVal _
                        lpDatabaseName As String, ByVal _
                        dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32" Alias "OpenServiceA" (ByVal _
                        hSCManager As Long, ByVal _
                        lpServiceName As String, ByVal _
                        dwDesiredAccess As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal _
                        hService As Long, _
                        lpServiceStatus As SERVICE_STATUS) As Long '获取服务状态用
Private Declare Function ControlService Lib "advapi32" (ByVal _
                        hService As Long, ByVal _
                        dwControl As Long, _
                        lpServiceStatus As SERVICE_STATUS) As Long '停止服务用
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal _
                        hService As Long, ByVal _
                        dwNumServiceArgs As Long, Optional ByVal _
                        lpServiceArgVectors As Long) As Long '启动服务用
Private Declare Function CloseServiceHandle Lib "advapi32" (ByVal hSCObject As Long) As Long '对服务操作完后关闭句柄用
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal _
                        lpPrevWndFunc As Long, ByVal _
                        hWnd As Long, ByVal _
                        Msg As Long, ByVal _
                        wParam As Long, ByVal _
                        lParam As Long) As Long '对消息的热键信息判断后需要把消息传到原本的窗口进程
Public Declare Function RegisterHotKey Lib "user32" (ByVal _
                        hWnd As Long, ByVal _
                        ID As Long, ByVal _
                        fsModifiers As Long, ByVal _
                        vk As Long) As Long '向系统注册热键
Public Declare Function UnregisterHotKey Lib "user32" (ByVal _
                        hWnd As Long, ByVal _
                        ID As Long) As Long '解除已经注册的热键(系统不会解除注册)

'--------------------------------------------------------------------------------- [Constants]

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1 'SetWindowPos
Public Const GWL_WNDPROC = (-4)
Private Const GWL_EXSTYLE = (-20) 'GetWindowLong

Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20& 'SetWindowLong
Private Const LWA_ALPHA As Long = &H2 'SetLayeredWindowAttributes
Private Const SPI_GETWORKAREA = 48 '取得工作区大小,用以将Frm_BG覆盖除了任务栏之外的整个区域
Private Const SPI_SETGRADIENTCAPTIONS = &H1009 '用以启用窗口标题栏渐变效果
Private Const COLOR_BACKGROUND = 1 '桌面颜色
Private Const COLOR_WINDOW = 5 '一般窗口背景
Private Const COLOR_WINDOWTEXT = 8 '窗口文本
Private Const COLOR_APPWORKSPACE = 12 '多文档界面的应用程序工作区背景
Private Const COLOR_WINDOWFRAME = 6 '窗体框架
Private Const COLOR_BTNFACE = 15 '3D对象 按钮 对话框背景
Private Const COLOR_BTNTEXT = 18 '按钮文本颜色
Private Const COLOR_SCROLLBAR = 0 '滚动条灰色区域
Private Const COLOR_CAPTIONTEXT = 9 '标题文本颜色 滚动条箭头
Private Const COLOR_ACTIVEBORDER = 10 '活动窗口边框
Private Const COLOR_ACTIVECAPTION = 2 '活动窗口标题栏左侧颜色
Private Const COLOR_GRADIENTACTIVECAPTION = 27 '活动窗口标题栏右侧颜色
Private Const COLOR_INACTIVEBORDER = 11 '非活动窗口边框
Private Const COLOR_INACTIVECAPTION = 3 '非活动窗口标题栏左侧颜色
Private Const COLOR_GRADIENTINACTIVECAPTION = 28 '非活动窗口标题栏右侧
Private Const COLOR_INACTIVECAPTIONTEXT = 19 '非活动窗口标题栏文本
Private Const COLOR_INFOBK = 24 '工具提示背景
Private Const COLOR_INFOTEXT = 23 '工具提示文本
Private Const COLOR_MENU = 4 '菜单背景颜色
Private Const COLOR_MENUTEXT = 7 '菜单文本
Private Const COLOR_HIGHLIGHT = 13 '选定项目背景
Private Const COLOR_HIGHLIGHTTEXT = 14 '选定项目文本
Private Const COLOR_HOTLIGHT = 26 '超链接颜色 (Colors)

Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_FORCE As Long = 4 'ExitWindowsEx

Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8 'AdjustTokenPrivileges  需要的访问权限

Private Const SE_PRIVILEGE_ENABLED = &H2 ' 控制特权的可使用性

Private Const SC_MANAGER_CONNECT = &H1 '连接到服务控制管理器的需要
Private Const SERVICE_STOP = &H20 'ControlService停止服务时需要的访问权
Private Const SERVICE_START = &H10 'StartService需要的访问权
Private Const SERVICE_QUERY_STATUS = &H4 'QueryServiceStatus需要的访问权
Private Const SERVICE_CONTROL_STOP As Long = 1&
Private Const SERVICE_RUNNING = &H4

Public Const WM_HOTKEY = &H312 '热键消息常数
Public Const MOD_ALT = &H1
'--------------------------------------------------------------------------------- [Type]

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type  'SystemParametersInfo的lpvParam参数在SPI_GETWkarea下的结构要求

Private Type LUID
     UsedPart As Long
     IgnoredForNowHigh32BitPart As Long
End Type 'LookupPrivilegeValue

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Private Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type

'------------------------------------------------------------------------------------[Sub&Function]

Sub Main()
    If App.PrevInstance = True Then    '如果如果已经运行就自己退出
        MsgBox "程序已经运行!", vbOKOnly, "提示"
        End
    End If
    CSetInitialized = False
    frm_Main.Show
End Sub


Public Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then '如果是热键消息
        If wParam = 1 Then '如果是本程序定义的
            If frmHelp.Visible = False Then ' 帮助窗口也没有显示
                Call WindowShowHide(frm_Main)
                Exit Function
            Else
                Beep
            End If
        End If
    End If
    WndProc = CallWindowProc(preWinProc, hWnd, Msg, wParam, lParam)
End Function


Public Sub Frm_Maximum(frm As Form)
    Dim wkarea As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, wkarea, 0
    With frm
        .Left = wkarea.Left * Screen.TwipsPerPixelX
        .Top = wkarea.Top * Screen.TwipsPerPixelY
        .Width = (wkarea.Right - wkarea.Left) * Screen.TwipsPerPixelX
        .Height = (wkarea.Bottom - wkarea.Top) * Screen.TwipsPerPixelY
    End With
End Sub


Public Function Judge(ByVal Lhwnd, ByVal Phwnd) As Boolean
    If (Lhwnd <> Phwnd) And _
        (Lhwnd <> 0) And _
        (Phwnd <> 0) And _
        (Phwnd <> Frm_BG.hWnd) And _
        (Phwnd <> frm_Main.hWnd) Then
        Judge = True
    Else
    Judge = False
    End If
End Function


Public Sub CSetInitialize()
    On Error Resume Next
    Dim I  As Integer '(用于后面循环的形式变量)
    ColorCategories(1) = COLOR_BACKGROUND
    ColorCategories(2) = COLOR_WINDOW
    ColorCategories(3) = COLOR_APPWORKSPACE
    ColorCategories(4) = COLOR_WINDOWFRAME
    ColorCategories(5) = COLOR_BTNFACE
    ColorCategories(6) = COLOR_BTNTEXT
    ColorCategories(7) = COLOR_SCROLLBAR
    ColorCategories(8) = COLOR_CAPTIONTEXT
    ColorCategories(9) = COLOR_ACTIVEBORDER
    ColorCategories(10) = COLOR_ACTIVECAPTION
    ColorCategories(11) = COLOR_GRADIENTACTIVECAPTION
    ColorCategories(12) = COLOR_INACTIVEBORDER
    ColorCategories(13) = COLOR_INACTIVECAPTION
    ColorCategories(14) = COLOR_GRADIENTINACTIVECAPTION
    ColorCategories(15) = COLOR_INACTIVECAPTIONTEXT
    ColorCategories(16) = COLOR_INFOBK
    ColorCategories(17) = COLOR_INFOTEXT
    ColorCategories(18) = COLOR_MENU
    ColorCategories(19) = COLOR_MENUTEXT
    ColorCategories(20) = COLOR_HIGHLIGHT
    ColorCategories(21) = COLOR_HIGHLIGHTTEXT
    ColorCategories(22) = COLOR_HOTLIGHT
    ColorCategories(23) = COLOR_WINDOWTEXT '要改变的各项颜色类别
    NewColor(1) = RGB(0, 0, 0)
    NewColor(2) = RGB(0, 0, 0)
    NewColor(3) = RGB(0, 0, 0)
    NewColor(4) = RGB(0, 0, 0)
    NewColor(5) = RGB(0, 0, 0)
    NewColor(6) = RGB(179, 147, 92)
    NewColor(7) = RGB(0, 0, 0)
    NewColor(8) = RGB(83, 199, 255)
    NewColor(9) = RGB(0, 0, 0)
    NewColor(10) = RGB(0, 0, 0)
    NewColor(11) = RGB(128, 128, 128)
    NewColor(12) = RGB(0, 0, 0)
    NewColor(13) = RGB(0, 0, 0)
    NewColor(14) = RGB(0, 0, 0)
    NewColor(15) = RGB(192, 192, 192)
    NewColor(16) = RGB(0, 0, 0)
    NewColor(17) = RGB(195, 190, 152)
    NewColor(18) = RGB(0, 0, 0)
    NewColor(19) = RGB(179, 147, 92)
    NewColor(20) = RGB(0, 0, 0)
    NewColor(21) = RGB(195, 190, 152)
    NewColor(22) = RGB(179, 147, 92)
    NewColor(23) = RGB(179, 147, 92) '这是自己制定的用于夜间关灯使用的比较柔和的颜色
    SystemParametersInfo SPI_SETGRADIENTCAPTIONS, 0, True, 0 '启用窗口标题栏渐变效果
    For I = 1 To 23
        OriginalColor(I) = GetSysColor(ColorCategories(I))
    Next I '将用户原来的颜色设置记录下来
        CSetInitialized = True
End Sub


Public Sub ShutDown() ' 关机
    Dim tempHandle  As Long
    Dim temptp As TOKEN_PRIVILEGES
    Dim formtp As TOKEN_PRIVILEGES
    Dim formlength As Long '作AdjustTokenPrivileges的形式参数
    OpenProcessToken GetCurrentProcess, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), tempHandle
    temptp.PrivilegeCount = 1
    LookupPrivilegeValue "", "SeShutdownPrivilege", temptp.TheLuid
    temptp.Attributes = SE_PRIVILEGE_ENABLED
    AdjustTokenPrivileges tempHandle, False, temptp, Len(formtp), formtp, formlength
    ExitWindowsEx (EWX_SHUTDOWN Or EWX_FORCE), 0
End Sub


Public Sub WindowTransparent(hWnd, Optional value)  '改变`设置透明度
    If IsMissing(value) = True Then
        SetWindowLong hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED Or WS_EX_TRANSPARENT)
        SetLayeredWindowAttributes hWnd, 0, 150, LWA_ALPHA
    Else
        SetLayeredWindowAttributes hWnd, 0, value, LWA_ALPHA
    End If
End Sub


Public Function CheckThemesStatus() As Boolean
    Dim hSCManager As Long '用来接收服务控制管理器数据库的句柄
    Dim hService As Long  ' 服务句柄
    Dim Status As SERVICE_STATUS '接收服务状态
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    hService = OpenService(hSCManager, "Themes", SERVICE_QUERY_STATUS)
    QueryServiceStatus hService, Status
    If Status.dwCurrentState = SERVICE_RUNNING Then
        CheckThemesStatus = True
    Else
        CheckThemesStatus = False
    End If
End Function


Public Sub StartThemesService()
    Dim hSCManager As Long
    Dim hService As Long
     hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    hService = OpenService(hSCManager, "Themes", SERVICE_START)
    StartService hService, 0
    CloseServiceHandle hService
    CloseServiceHandle hSCManager
End Sub


Public Sub StopThemesService()
    Dim hSCManager As Long
    Dim hService As Long
    Dim Status As SERVICE_STATUS
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    hService = OpenService(hSCManager, "Themes", SERVICE_STOP)
    ControlService hService, SERVICE_CONTROL_STOP, Status
    CloseServiceHandle hService
    CloseServiceHandle hSCManager
End Sub


Public Sub WindowShowHide(Form As Form)  ' 用于隐藏显示窗口
    Select Case Form.Visible
        Case True
            Form.Hide
        Case False
            Form.Show
    End Select
End Sub


'未使用的
'***************************************************************
'Judge函数被舍弃的两个算法,原因:不清晰
'算法1
'Public Function Judge2(ByVal Lhwnd As Long, ByVal Phwnd As Long) As Boolean
'If Lhwnd <> Phwnd Then
'    If (Lhwnd <> 0) And (Phwnd <> 0) Then
'        If (Phwnd <> FrmBG.hWnd) And (Phwnd <> FrmCtrl.hWnd) Then
'            Judge2 = True
'            Exit Function
'        Else
'            Judge2 = False
'            Exit Function
'        End If
'    Else
'        Judge2 = False
'        Exit Function
'    End If
'Else
'    Judg2e = False
'    Exit Function
'End If
'End Function


''算法2
'Public Function Judge3(ByVal Lhwnd, ByVal Phwnd) As Boolean
'If Lhwnd = Phwnd Then
'Judge3 = False
'Exit Function
'End If
'If (Lhwnd = 0) Or (Phwnd = 0) Then
'Judge3 = False
'Exit Function
'End If
'If (Phwnd = FrmBG.hWnd) Or (Phwnd = FrmCtrl.hWnd) Then
'Judge3 = False
'Exit Function
'End If
'Judge3 = True
'End Function



'为了最大化frmBG写的,任务栏高度计算,后来发现思路有漏洞,如果仅仅用这种方式计算,当任务栏不在屏幕底部的时候,
'不仅会得到错误的任务栏高度,而且据此数据进行最大化也会出问题
'SystemParametersInfo SPI_GETWORKAREA, 0, wkarea, 0
''***************************思路
'wkarea.Bottom '工作区高 / Pixel
''TwipsPerPixel = Twips/Pixel
''==>Twips = Pixel*TwipsPerPixel
''==>Pixel =Twips/TwipsPerPixel
'TaskBarHeight = Screen.Height - wkarea.Bottom * Screen.TwipsPerPixelY 'Twip = Twip - Pixel*TwipsPerPixel ====>Twip = Twip - Twip
''****************************
''写成Function
'Public Function GetTaskBarHeight() As Long
'Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal _
'                        uAction As Long, ByVal _
'                        uParam As Long, ByVal _
'                        lpvParam As Any, ByVal _
'                        fuWinIni As Long) As Long
'Private Type RECT 'SystemParametersInfo的lpvParam参数在SPI_GETWORKAREA下的结构要求
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'Private TaskBar As Long
'Private WorkArea As RECT
'Private Const SPI_GETWORKAREA = 48 'SystemParametersInfo
'SystemParametersInfo SPI_GETWORKAREA, 0, WorkArea, 0
'GetTaskBarHeight = Screen.Height - WorkArea.Bottom * Screen.TwipsPerPixelY 'Twip = Twip - Pixel*TwipsPerPixel ====>Twip = Twip - Twip
'End Function
