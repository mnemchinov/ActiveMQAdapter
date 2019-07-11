Friend Class V7Data

  Public Shared Property V7Object() As Object
    Get
      Return m_V7Object
    End Get
    Set(ByVal Value As Object)
      m_V7Object = Value
      ' Вызываем неявно QueryInterface
      m_ErrorInfo = CType(Value, IErrorLog)
      m_AsyncEvent = CType(Value, IAsyncEvent)
      m_StatusLine = CType(Value, IStatusLine)
    End Set
  End Property

  Public Shared ReadOnly Property ErrorLog() As IErrorLog
    Get
      Return m_ErrorInfo
    End Get
  End Property

  Public Shared ReadOnly Property AsyncEvent() As IAsyncEvent
    Get
      Return m_AsyncEvent
    End Get
  End Property

  Public Shared ReadOnly Property StatusLine() As IStatusLine
    Get
      Return m_StatusLine
    End Get
  End Property

  Private Shared m_V7Object As Object
  Private Shared m_ErrorInfo As IErrorLog
  Private Shared m_AsyncEvent As IAsyncEvent
  Private Shared m_StatusLine As IStatusLine

End Class
