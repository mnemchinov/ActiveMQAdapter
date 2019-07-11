Imports System.Runtime.InteropServices

<Guid("ab634004-f13d-11d0-a459-004095e1daea"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
Public Interface IAsyncEvent
  Sub SetEventBufferDepth(ByVal lDepth As Integer)
  Sub GetEventBufferDepth(ByRef plDepth As Integer)
  Sub ExternalEvent(ByVal bstrSource As String, ByVal bstrMessage As String, ByVal bstrData As String)
  Sub CleanBuffer()
End Interface
