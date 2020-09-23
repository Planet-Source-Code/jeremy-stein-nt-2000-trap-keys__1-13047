Attribute VB_Name = "modTrapKeys"
Option Explicit

Public Const HC_ACTION = 0
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const VK_TAB = &H9
Public Const VK_CONTROL = &H11
Public Const VK_ESCAPE = &H1B
Public Const VK_LWIN = &H5B
Public Const VK_RWIN = &H5C
Public Const WH_KEYBOARD_LL = 13
Public Const LLKHF_ALTDOWN = &H20

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Dim k As KBDLLHOOKSTRUCT

Public Function KeyboardHookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   
   Dim TrapKey As Boolean
   
   If (nCode = HC_ACTION) Then
      If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Or wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
         CopyMemory k, ByVal lParam, Len(k)
         TrapKey = k.vkCode = VK_LWIN Or k.vkCode = VK_RWIN Or ((k.vkCode = VK_TAB) And ((k.flags And LLKHF_ALTDOWN) <> 0)) Or ((k.vkCode = VK_ESCAPE) And ((k.flags And LLKHF_ALTDOWN) <> 0)) Or ((k.vkCode = VK_ESCAPE) And ((GetKeyState(VK_CONTROL) And &H8000) <> 0))
        End If
    End If
    
    If TrapKey Then
        KeyboardHookProc = -1
    Else
        KeyboardHookProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
    End If
    
End Function

