Attribute VB_Name = "Module2"
Option Explicit

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Const WH_KEYBOARD = 13

Public hHook As Long

Public Type KBDLLHOOKSTRUCT

    vkCode As Long

    scanCode As Long

    flags As Long

    time As Long

    dwExtraInfo As Long

End Type

Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, lParam As KBDLLHOOKSTRUCT) As Long

With lParam




If nCode < 0 Or .dwExtraInfo = 33 Then

LowLevelKeyboardProc = CallNextHookEx(hHook, nCode, wParam, lParam)

Exit Function

End If

If .flags = 0 Then
If Form7.DoingA = False Then
If .vkCode = "8" Or .vkCode = "110" Then
If Len(Form7.TempKey) > 0 Then Form7.TempKey = "" & Left(Form7.TempKey, Len(Form7.TempKey) - 1) & ""
ElseIf .vkCode = "144" Then
LowLevelKeyboardProc = 1
    ElseIf .vkCode = "49" Or .vkCode = "97" Then
        LowLevelKeyboardProc = 1
        Form7.TempKey = "" & Form7.TempKey & "1"
                ElseIf .vkCode = "50" Or .vkCode = "98" Then
        LowLevelKeyboardProc = 1
        Form7.TempKey = "" & Form7.TempKey & "2"
                ElseIf .vkCode = "51" Or .vkCode = "99" Then
        LowLevelKeyboardProc = 1
        Form7.TempKey = "" & Form7.TempKey & "3"
                ElseIf .vkCode = "52" Or .vkCode = "100" Then
        LowLevelKeyboardProc = 1
        Form7.TempKey = "" & Form7.TempKey & "4"
                ElseIf .vkCode = "53" Or .vkCode = "101" Then
        LowLevelKeyboardProc = 1
        Form7.TempKey = "" & Form7.TempKey & "5"
                ElseIf .vkCode = "54" Or .vkCode = "102" Then
        LowLevelKeyboardProc = 1
        Form7.TempKey = "" & Form7.TempKey & "6"
                ElseIf .vkCode = "55" Or .vkCode = "103" Then
        LowLevelKeyboardProc = 1
        Form7.TempKey = "" & Form7.TempKey & "7"
                ElseIf .vkCode = "56" Or .vkCode = "104" Then
        LowLevelKeyboardProc = 1
        Form7.TempKey = "" & Form7.TempKey & "8"
                ElseIf .vkCode = "57" Or .vkCode = "105" Then
        LowLevelKeyboardProc = 1
        Form7.TempKey = "" & Form7.TempKey & "9"
                ElseIf .vkCode = "48" Or .vkCode = "96" Then
        LowLevelKeyboardProc = 1
        Form7.TempKey = "" & Form7.TempKey & "0"
    
    End If
Form7.Label3.Caption = Form7.TempKey
End If
End If
End With




LowLevelKeyboardProc = 0

End Function
