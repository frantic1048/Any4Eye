VERSION 5.00
Begin VB.Form frmRemind 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   1890
   ClientTop       =   1545
   ClientWidth     =   23205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRemind.frx":0000
   ScaleHeight     =   3720
   ScaleWidth      =   23205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "frmRemind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal _
                        nHeight As Long, ByVal _
                        nWidth As Long, ByVal _
                        nEscapement As Long, ByVal _
                        nOrientation As Long, ByVal _
                        fnWeight As Long, ByVal _
                        fdwItalic As Long, ByVal _
                        fdwUnderline As Long, ByVal _
                        fdwStrikeOut As Long, ByVal _
                        fdwCharSet As Long, ByVal _
                        fdwOutputPrecision As Long, ByVal _
                        fdwClipPrecision As Long, ByVal _
                        fdwQuality As Long, ByVal _
                        fdwPitchAndFamily As Long, ByVal _
                        lpszFace As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal _
                        hDC As Long, ByVal _
                        hObject As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal _
                        hDC As Long, ByVal _
                        nBkMode As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal _
                        hDC As Long, ByVal _
                        x As Long, ByVal _
                        y As Long, ByVal _
                        lpString As String, ByVal _
                        nCount As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal _
                        hWnd As Long, ByVal _
                        hRgn As Long, ByVal _
                        bRedraw As Boolean) As Long



Private Const TRANSPARENT = 1
Private Const ANSI_CHARSET = 0
Private Const FW_HEAVY = 900
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_SWISS = 32


Private Sub Form_Load()
    Dim hDC
    Dim wndRgn As Long
    Dim Font As Long
    Dim OldFont As Long
    Const tip As String = "你已经使用电脑一小时了,请注意休息"
    
    hDC = Me.hDC
    Font = CreateFont( _
                        80, 40, _
                        0, 0, _
                        FW_HEAVY, 0, 0, 0, _
                        ANSI_CHARSET, _
                        OUT_DEFAULT_PRECIS, _
                        CLIP_DEFAULT_PRECIS, _
                        DEFAULT_QUALITY, _
                       DEFAULT_PITCH Or FF_SWISS, _
                       "黑体")
    BeginPath hDC
    SetBkMode hDC, TRANSPARENT
    OldFont = SelectObject(hDC, Font)
    TextOut hDC, 0, 0, tip, LenB(StrConv(tip, vbFromUnicode))
    SelectObject hDC, OldFont
    EndPath hDC
    wndRgn = PathToRegion(hDC)
    SetWindowRgn Me.hWnd, wndRgn, True
End Sub
