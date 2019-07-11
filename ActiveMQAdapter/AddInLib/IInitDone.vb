Imports System.Runtime.InteropServices

<Guid("AB634001-F13D-11d0-A459-004095E1DAEA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
Public Interface IInitDone
  Sub Init(<MarshalAs(UnmanagedType.IDispatch)> ByVal pConnection As Object)
  Sub Done()
  Sub GetInfo(ByRef pInfo() As Object)
End Interface
