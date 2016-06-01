VERSION 5.00
Begin VB.Form frmHelp 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "关于我的应用程序"
   ClientHeight    =   4470
   ClientLeft      =   7875
   ClientTop       =   4320
   ClientWidth     =   5745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHelp.frx":0000
   ScaleHeight     =   3085.273
   ScaleMode       =   0  'User
   ScaleWidth      =   5394.852
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   780
      Left            =   120
      Picture         =   "frmHelp.frx":57498
      ScaleHeight     =   505.68
      ScaleMode       =   0  'User
      ScaleWidth      =   505.68
      TabIndex        =   0
      Top             =   120
      Width           =   780
   End
   Begin VB.Image img_SysInfo 
      Height          =   615
      Left            =   4440
      MouseIcon       =   "frmHelp.frx":5941A
      MousePointer    =   99  'Custom
      Top             =   2760
      Width           =   855
   End
   Begin VB.Image img_OK 
      Height          =   615
      Left            =   4440
      MouseIcon       =   "frmHelp.frx":5956C
      MousePointer    =   99  'Custom
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Frantic Black     5,May,2012"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2934
      TabIndex        =   5
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Label lbl_Link 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://blog.163.com/frantic_hao/"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1680
      MouseIcon       =   "frmHelp.frx":596BE
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5183.565
      Y1              =   1076.74
      Y2              =   1076.74
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "应用程序描述"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5183.565
      Y1              =   1076.74
      Y2              =   1076.74
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "版本:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4080
      TabIndex        =   3
      Top             =   1200
      Width           =   1365
   End
   Begin VB.Label lblTip 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "警告: ..."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2235
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   3855
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'此窗体的代码由VB生成,我只做了小部分的更改


Option Explicit

' 注册表关键字安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字 ROOT 类型...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' 独立的空的终结字符串
Const REG_DWORD = 4                      ' 32位数字

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"



Private Const URL = "http://blog.163.com/frantic_hao/"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long





Private Sub Form_Load()
    Me.Caption = "关于 " & App.Title
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblDescription.Caption = "关灯不是Tudou,Youku们才会!!"
    lblTip.Caption = "提示:" & vbCrLf & "→Alt + N 可以隐藏/显示本程序" & vbCrLf & "→延时关机为指定的分钟数之后关机" & vbCrLf & "→定时关机是在用户指定的时刻关机" & vbCrLf & vbCrLf & vbCrLf & ">>欢迎到我的博客留言提供问题反馈与建议<<"
    
    WindowTransparent frmHelp.hwnd, , True
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' 试图从注册表中获得系统信息程序的路径及名称...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' 试图仅从注册表中获得系统信息程序的路径...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' 已知32位文件版本的有效位置
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' 错误 - 文件不能被找到...
        Else
            GoTo SysInfoErr
        End If
    ' 错误 - 注册表相应条目不能被找到...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "此时系统信息不可用", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim I As Long                                           ' 循环计数器
    Dim rc As Long                                          ' 返回代码
    Dim hKey As Long                                        ' 打开的注册表关键字句柄
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' 注册表关键字数据类型
    Dim tmpVal As String                                    ' 注册表关键字值的临时存储器
    Dim KeyValSize As Long                                  ' 注册表关键自变量的尺寸
    '------------------------------------------------------------
    ' 打开 {HKEY_LOCAL_MACHINE...} 下的 RegKey
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册表关键字
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误...
    
    tmpVal = String$(1024, 0)                             ' 分配变量空间
    KeyValSize = 1024                                       ' 标记变量尺寸
    
    '------------------------------------------------------------
    ' 检索注册表关键字的值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' 获得/创建关键字值
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 外接程序空终结字符串...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null 被找到,从字符串中分离出来
    Else                                                    ' WinNT 没有空终结字符串...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null 没有被找到, 分离字符串
    End If
    '------------------------------------------------------------
    ' 决定转换的关键字的值类型...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' 搜索数据类型...
    Case REG_SZ                                             ' 字符串注册关键字数据类型
        KeyVal = tmpVal                                     ' 复制字符串的值
    Case REG_DWORD                                          ' 四字节的注册表关键字数据类型
        For I = Len(tmpVal) To 1 Step -1                    ' 将每位进行转换
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' 生成值字符。 By Char。
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' 转换四字节的字符为字符串
    End Select
    
    GetKeyValue = True                                      ' 返回成功
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
    Exit Function                                           ' 退出
    
GetKeyError:      ' 错误发生后将其清除...
    KeyVal = ""                                             ' 设置返回值到空字符串
    GetKeyValue = False                                     ' 返回失败
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragWindow Button, Me.hwnd
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbl_Link.FontItalic = False
End Sub

Private Sub img_OK_Click()
    Unload Me
End Sub

Private Sub img_SysInfo_Click()
    Call StartSysInfo
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragWindow Button, Me.hwnd
End Sub

Private Sub lbl_Link_Click()
    Dim Success As Long
    Success = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub

Private Sub lbl_Link_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbl_Link.FontItalic = True
End Sub

Private Sub lblDescription_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragWindow Button, Me.hwnd
End Sub

Private Sub lblTip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragWindow Button, Me.hwnd
End Sub


Private Sub lblVersion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragWindow Button, Me.hwnd
End Sub

Private Sub picIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     DragWindow Button, Me.hwnd
End Sub
