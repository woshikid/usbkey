Attribute VB_Name = "Module1"
Option Explicit
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long

Public hook1 As Long
Public hook2 As Long
Public Const WH_KEYBOARD = 2
Public Const WH_MOUSE = 7

Public Sub UnHook()
    On Error Resume Next
    UnhookWindowsHookEx hook1
    UnhookWindowsHookEx hook2
End Sub
Public Function EnableHook()
    On Error Resume Next
    hook1 = SetWindowsHookEx(WH_KEYBOARD, AddressOf MyFunc, App.hInstance, 0)
    hook2 = SetWindowsHookEx(WH_MOUSE, AddressOf MyFunc, App.hInstance, 0)
End Function
Public Function MyFunc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    MyFunc = 1  '≥‘µÙ—∂œ¢
    Exit Function
End Function
