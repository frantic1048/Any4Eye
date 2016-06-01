Attribute VB_Name = "GeneralModule"
'ע���е�����Ϊ����˵��
Option Explicit
Option Base 1

Public HFWnd As Long '(Hwnd of Last Foreground Window)
Public LastHFWnd As Long '(Hwnd of Last Foreground Window)
Public HBG As Long '(Hwnd of Frm_BG)
Public NewColor(23) As Long
Public OriginalColor(23) As Long '��������ԭ����ϵͳ��ɫ(Value of Original SysColor)
Public ColorCategories(23) As Long '������Ÿı����ɫ���
Public CSetInitialized As Boolean '�����ɫ��¼�Ƿ��ʼ��,(ColorSetInitialized)
Public DelayShutDown As Boolean '�������ʱ��Ի������ʾ����,trueΪ��ʱ����,falseΪ��׼ʱ���ʽ����
Public ShutDownTime  '������ʾ�ػ�ʱ��
Public OriginalThemesStatus As Boolean  ' �������ԭ���û��Ƿ�ʹ����Themes����
Public preWinProc As Long '�洢ԭ�����ڹ��̵ĵ�ַ

'-----------------------------------------------------------------------------------[API Void ]

Public Declare Function GetForegroundWindow Lib "user32" () As Long '��ȡ�������ڵľ��
Public Declare Function SetWindowPos Lib "user32" (ByVal _
                        hWnd As Long, ByVal _
                        hWndInsertAfter As Long, ByVal _
                        x As Long, ByVal _
                        y As Long, ByVal _
                        cx As Long, ByVal _
                        cy As Long, ByVal _
                        wFlags As Long) As Long ' �ﵽFrm_BG��Զ�ڹ���������
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal _
                        hWnd As Long, ByVal _
                        nIndex As Long) As Long ' ��ȡ��չ��ʽExtend Style
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long ' ���ϵͳ��ɫ
Public Declare Function SetSysColors Lib "user32" (ByVal _
                        nChanges As Long, _
                        lpSysColor As Long, _
                        lpColorValues As Long) As Long ' ����ϵͳ��ɫ
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal _
                        hWnd As Long, ByVal _
                        nIndex As Long, ByVal _
                        dwNewLong As Long) As Long ' Ҫ������չ��ʽ
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal _
                        hWnd As Long, ByVal _
                        crKey As Long, ByVal _
                        bAlpha As Byte, ByVal _
                        dwFlags As Long) As Long               ' ͸����
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal _
                        uAction As Long, ByVal _
                        uParam As Long, ByRef _
                        lpvParam As Any, ByVal _
                        fuWinIni As Long) As Long ' �õ���������
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, _
                        ByVal dwReserved As Long) As Long '  �ػ�,������ҪȨ��
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long ' ��õ�ǰ���̾��
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal _
                        DesiredAccess As Long, _
                        TokenHandle As Long) As Long ' �޸Ľ��̵ķ�������
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal _
                        lpSystemName As String, ByVal _
                        lpName As String, _
                        lpLuid As LUID) As Long ' ȡ�ùػ�Ȩ�޶�Ӧ�ı���Ψһ��ʾ��
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal _
                        TokenHandle As Long, ByVal _
                        DisableAllPrivileges As Long, _
                        NewState As TOKEN_PRIVILEGES, ByVal _
                        BufferLength As Long, _
                        PreviousState As TOKEN_PRIVILEGES, _
                        ReturnLength As Long) As Long '�ڽ��̵ķ������������ùػ�Ȩ��
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
                        lpServiceStatus As SERVICE_STATUS) As Long '��ȡ����״̬��
Private Declare Function ControlService Lib "advapi32" (ByVal _
                        hService As Long, ByVal _
                        dwControl As Long, _
                        lpServiceStatus As SERVICE_STATUS) As Long 'ֹͣ������
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal _
                        hService As Long, ByVal _
                        dwNumServiceArgs As Long, Optional ByVal _
                        lpServiceArgVectors As Long) As Long '����������
Private Declare Function CloseServiceHandle Lib "advapi32" (ByVal hSCObject As Long) As Long '�Է���������رվ����
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal _
                        lpPrevWndFunc As Long, ByVal _
                        hWnd As Long, ByVal _
                        Msg As Long, ByVal _
                        wParam As Long, ByVal _
                        lParam As Long) As Long '����Ϣ���ȼ���Ϣ�жϺ���Ҫ����Ϣ����ԭ���Ĵ��ڽ���
Public Declare Function RegisterHotKey Lib "user32" (ByVal _
                        hWnd As Long, ByVal _
                        ID As Long, ByVal _
                        fsModifiers As Long, ByVal _
                        vk As Long) As Long '��ϵͳע���ȼ�
Public Declare Function UnregisterHotKey Lib "user32" (ByVal _
                        hWnd As Long, ByVal _
                        ID As Long) As Long '����Ѿ�ע����ȼ�(ϵͳ������ע��)

'--------------------------------------------------------------------------------- [Constants]

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1 'SetWindowPos
Public Const GWL_WNDPROC = (-4)
Private Const GWL_EXSTYLE = (-20) 'GetWindowLong

Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20& 'SetWindowLong
Private Const LWA_ALPHA As Long = &H2 'SetLayeredWindowAttributes
Private Const SPI_GETWORKAREA = 48 'ȡ�ù�������С,���Խ�Frm_BG���ǳ���������֮�����������
Private Const SPI_SETGRADIENTCAPTIONS = &H1009 '�������ô��ڱ���������Ч��
Private Const COLOR_BACKGROUND = 1 '������ɫ
Private Const COLOR_WINDOW = 5 'һ�㴰�ڱ���
Private Const COLOR_WINDOWTEXT = 8 '�����ı�
Private Const COLOR_APPWORKSPACE = 12 '���ĵ������Ӧ�ó�����������
Private Const COLOR_WINDOWFRAME = 6 '������
Private Const COLOR_BTNFACE = 15 '3D���� ��ť �Ի��򱳾�
Private Const COLOR_BTNTEXT = 18 '��ť�ı���ɫ
Private Const COLOR_SCROLLBAR = 0 '��������ɫ����
Private Const COLOR_CAPTIONTEXT = 9 '�����ı���ɫ ��������ͷ
Private Const COLOR_ACTIVEBORDER = 10 '����ڱ߿�
Private Const COLOR_ACTIVECAPTION = 2 '����ڱ����������ɫ
Private Const COLOR_GRADIENTACTIVECAPTION = 27 '����ڱ������Ҳ���ɫ
Private Const COLOR_INACTIVEBORDER = 11 '�ǻ���ڱ߿�
Private Const COLOR_INACTIVECAPTION = 3 '�ǻ���ڱ����������ɫ
Private Const COLOR_GRADIENTINACTIVECAPTION = 28 '�ǻ���ڱ������Ҳ�
Private Const COLOR_INACTIVECAPTIONTEXT = 19 '�ǻ���ڱ������ı�
Private Const COLOR_INFOBK = 24 '������ʾ����
Private Const COLOR_INFOTEXT = 23 '������ʾ�ı�
Private Const COLOR_MENU = 4 '�˵�������ɫ
Private Const COLOR_MENUTEXT = 7 '�˵��ı�
Private Const COLOR_HIGHLIGHT = 13 'ѡ����Ŀ����
Private Const COLOR_HIGHLIGHTTEXT = 14 'ѡ����Ŀ�ı�
Private Const COLOR_HOTLIGHT = 26 '��������ɫ (Colors)

Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_FORCE As Long = 4 'ExitWindowsEx

Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8 'AdjustTokenPrivileges  ��Ҫ�ķ���Ȩ��

Private Const SE_PRIVILEGE_ENABLED = &H2 ' ������Ȩ�Ŀ�ʹ����

Private Const SC_MANAGER_CONNECT = &H1 '���ӵ�������ƹ���������Ҫ
Private Const SERVICE_STOP = &H20 'ControlServiceֹͣ����ʱ��Ҫ�ķ���Ȩ
Private Const SERVICE_START = &H10 'StartService��Ҫ�ķ���Ȩ
Private Const SERVICE_QUERY_STATUS = &H4 'QueryServiceStatus��Ҫ�ķ���Ȩ
Private Const SERVICE_CONTROL_STOP As Long = 1&
Private Const SERVICE_RUNNING = &H4

Public Const WM_HOTKEY = &H312 '�ȼ���Ϣ����
Public Const MOD_ALT = &H1
'--------------------------------------------------------------------------------- [Type]

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type  'SystemParametersInfo��lpvParam������SPI_GETWkarea�µĽṹҪ��

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
    If App.PrevInstance = True Then    '�������Ѿ����о��Լ��˳�
        MsgBox "�����Ѿ�����!", vbOKOnly, "��ʾ"
        End
    End If
    CSetInitialized = False
    frm_Main.Show
End Sub


Public Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then '������ȼ���Ϣ
        If wParam = 1 Then '����Ǳ��������
            If frmHelp.Visible = False Then ' ��������Ҳû����ʾ
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
    Dim I  As Integer '(���ں���ѭ������ʽ����)
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
    ColorCategories(23) = COLOR_WINDOWTEXT 'Ҫ�ı�ĸ�����ɫ���
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
    NewColor(23) = RGB(179, 147, 92) '�����Լ��ƶ�������ҹ��ص�ʹ�õıȽ���͵���ɫ
    SystemParametersInfo SPI_SETGRADIENTCAPTIONS, 0, True, 0 '���ô��ڱ���������Ч��
    For I = 1 To 23
        OriginalColor(I) = GetSysColor(ColorCategories(I))
    Next I '���û�ԭ������ɫ���ü�¼����
        CSetInitialized = True
End Sub


Public Sub ShutDown() ' �ػ�
    Dim tempHandle  As Long
    Dim temptp As TOKEN_PRIVILEGES
    Dim formtp As TOKEN_PRIVILEGES
    Dim formlength As Long '��AdjustTokenPrivileges����ʽ����
    OpenProcessToken GetCurrentProcess, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), tempHandle
    temptp.PrivilegeCount = 1
    LookupPrivilegeValue "", "SeShutdownPrivilege", temptp.TheLuid
    temptp.Attributes = SE_PRIVILEGE_ENABLED
    AdjustTokenPrivileges tempHandle, False, temptp, Len(formtp), formtp, formlength
    ExitWindowsEx (EWX_SHUTDOWN Or EWX_FORCE), 0
End Sub


Public Sub WindowTransparent(hWnd, Optional value)  '�ı�`����͸����
    If IsMissing(value) = True Then
        SetWindowLong hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED Or WS_EX_TRANSPARENT)
        SetLayeredWindowAttributes hWnd, 0, 150, LWA_ALPHA
    Else
        SetLayeredWindowAttributes hWnd, 0, value, LWA_ALPHA
    End If
End Sub


Public Function CheckThemesStatus() As Boolean
    Dim hSCManager As Long '�������շ�����ƹ��������ݿ�ľ��
    Dim hService As Long  ' ������
    Dim Status As SERVICE_STATUS '���շ���״̬
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


Public Sub WindowShowHide(Form As Form)  ' ����������ʾ����
    Select Case Form.Visible
        Case True
            Form.Hide
        Case False
            Form.Show
    End Select
End Sub


'δʹ�õ�
'***************************************************************
'Judge�����������������㷨,ԭ��:������
'�㷨1
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


''�㷨2
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



'Ϊ�����frmBGд��,�������߶ȼ���,��������˼·��©��,������������ַ�ʽ����,��������������Ļ�ײ���ʱ��,
'������õ�������������߶�,���Ҿݴ����ݽ������Ҳ�������
'SystemParametersInfo SPI_GETWORKAREA, 0, wkarea, 0
''***************************˼·
'wkarea.Bottom '�������� / Pixel
''TwipsPerPixel = Twips/Pixel
''==>Twips = Pixel*TwipsPerPixel
''==>Pixel =Twips/TwipsPerPixel
'TaskBarHeight = Screen.Height - wkarea.Bottom * Screen.TwipsPerPixelY 'Twip = Twip - Pixel*TwipsPerPixel ====>Twip = Twip - Twip
''****************************
''д��Function
'Public Function GetTaskBarHeight() As Long
'Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal _
'                        uAction As Long, ByVal _
'                        uParam As Long, ByVal _
'                        lpvParam As Any, ByVal _
'                        fuWinIni As Long) As Long
'Private Type RECT 'SystemParametersInfo��lpvParam������SPI_GETWORKAREA�µĽṹҪ��
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
