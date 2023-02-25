Attribute VB_Name = "Module1"
'�ö�api
Public Declare Function SetWindowPos Lib _
    "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fskey_Modifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long

Public Const WM_HOTKEY = &H312
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const GWL_WNDPROC = (-4)

Public preWinProc As Long
Public Modifiers As Long, uVirtKey1 As Long, idHotKey As Long

Public Function SetHotKey()
    preWinProc = GetWindowLong(Form1.hwnd, GWL_WNDPROC)
    SetWindowLong Form1.hwnd, GWL_WNDPROC, AddressOf Keywndproc '��������ص�
    '=========
    RegisterHotKey Form1.hwnd, 1, MOD_CONTROL, vbKey1 'ע���ȼ�
    RegisterHotKey Form1.hwnd, 2, MOD_CONTROL, vbKey2 'ע���ȼ�
    RegisterHotKey Form1.hwnd, 3, MOD_CONTROL, vbKey3 'ע���ȼ�
    RegisterHotKey Form1.hwnd, 4, MOD_CONTROL, vbKey4 'ע���ȼ�
    RegisterHotKey Form1.hwnd, 5, MOD_CONTROL, vbKey5 'ע���ȼ�
    '=========
    '���ﻹ�������
End Function


Public Function Keywndproc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then
        Form1.Paste (wParam - 1) 'ճ��
        
        'Select Case wParam
        '    Case 1
        '        MsgBox "A1"
        '    Case 2
        '        MsgBox "A2"
        '    Case 3
        '        MsgBox "A3"
        '    Case 4
        '        MsgBox "A4"
        '    Case 5
        '        MsgBox "A5"
        'End Select
    End If
    Keywndproc = CallWindowProc(preWinProc, hwnd, Msg, wParam, lParam)
End Function





