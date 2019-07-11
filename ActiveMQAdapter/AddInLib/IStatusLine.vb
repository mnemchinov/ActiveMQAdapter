Imports System.Runtime.InteropServices

<Guid("ab634005-f13d-11d0-a459-004095e1daea"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
Public Interface IStatusLine
  Sub SetStatusLine(ByVal bstrStatusLine As String)
  Sub ResetStatusLine()
End Interface
