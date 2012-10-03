Public Class KeyboardWatcher

  Public Event KeyPressed()

  Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Integer) As Integer
  Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Integer, ByVal lpfn As KeyboardHookDelegate, ByVal hmod As Integer, ByVal dwThreadId As Integer) As Integer
  Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Integer, ByVal nCode As Integer, ByVal wParam As Integer, ByVal lParam As KBDLLHOOKSTRUCT) As Integer

  Private Delegate Function KeyboardHookDelegate(ByVal Code As Integer, ByVal wParam As Integer, ByRef lParam As KBDLLHOOKSTRUCT) As Integer

  Private Const WM_KEYUP As Integer = &H101
  Private Const WM_SYSKEYUP As Integer = &H105

  Private m_ptrKeyboardHandle As IntPtr = 0
  Private m_cbKeyboardHook As KeyboardHookDelegate = Nothing

  Public Structure KBDLLHOOKSTRUCT
    Public vkCode As Integer
    Public scanCode As Integer
    Public flags As Integer
    Public time As Integer
    Public dwExtraInfo As Integer
  End Structure

  Enum virtualKey
    K_Return = &HD
    K_Backspace = &H8
    K_Space = &H20
    K_Tab = &H9
    K_Esc = &H1B
  End Enum

  Public Sub HookKeyboard()
    m_cbKeyboardHook = New KeyboardHookDelegate(AddressOf KeyboardCallback)
    m_ptrKeyboardHandle = SetWindowsHookEx(13, m_cbKeyboardHook, Process.GetCurrentProcess.MainModule.BaseAddress, 0)
  End Sub

  Private Function Hooked()
    Return m_ptrKeyboardHandle <> 0
  End Function

  Public Function KeyboardCallback(ByVal Code As Integer, ByVal wParam As Integer, ByRef lParam As KBDLLHOOKSTRUCT) As Integer
    If wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
      RaiseEvent KeyPressed()
    End If

    Return CallNextHookEx(m_ptrKeyboardHandle, Code, wParam, lParam)
  End Function

  Public Sub UnhookKeyboard()
    If (Hooked()) Then
      If UnhookWindowsHookEx(m_ptrKeyboardHandle) <> 0 Then
        m_ptrKeyboardHandle = 0
      End If
    End If
  End Sub

End Class