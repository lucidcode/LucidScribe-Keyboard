Public Class PluginHandler
  Inherits lucidcode.LucidScribe.Interface.LucidPluginBase

  Private m_objKeyboardMonitor As KeyboardWatcher
  Private m_dblValue As Double = 0
    Private m_dblPreviousValue = 0
    Private m_intNoChange = 0

  Public Overrides ReadOnly Property Name() As String
    Get
      Return "Keyboard"
    End Get
  End Property

  Public Overrides ReadOnly Property Value() As Double
        Get
            If m_dblValue > 999 Then m_dblValue = 999
            If m_dblValue < 0 Then m_dblValue = 0

            ' Make sure the value is not stuck 
            If m_dblValue > 0 Then
                If m_dblPreviousValue = m_dblValue Then
                    m_intNoChange += 1
                Else
                    m_intNoChange = 0
                End If
                If m_intNoChange > 100 Then
                    m_dblValue = 0
                    m_intNoChange = 0
                End If
                m_dblPreviousValue = m_dblValue
            End If

            Return m_dblValue
        End Get
  End Property

  Public Overrides Function Initialize() As Boolean
    m_objKeyboardMonitor = New KeyboardWatcher
    m_objKeyboardMonitor.HookKeyboard()
    AddHandler m_objKeyboardMonitor.KeyPressed, AddressOf KeyboardWatcher_KeyPressed
    Return True
  End Function

  Private Sub KeyboardWatcher_KeyPressed()
    m_dblValue += 100
        If m_dblValue > 999 Then m_dblValue = 999

    ' Remove the 100 from the value after 1 second (on another thread)
    Dim objSubtractThread As New Threading.Thread(AddressOf SubtractValue)
    objSubtractThread.Start()
  End Sub

  Private Sub SubtractValue()
    Threading.Thread.Sleep(1000)
    m_dblValue -= 100
    If m_dblValue < 0 Then m_dblValue = 0
  End Sub

End Class