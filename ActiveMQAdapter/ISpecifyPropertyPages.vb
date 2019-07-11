Imports System.Runtime.InteropServices

<Guid("B196B28B-BAB4-101A-B69C-00AA00341D07"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
Public Interface ISpecifyPropertyPages
  Sub GetPages(<Out()> ByRef pPages As CAUUID)
End Interface
'--------------------------------------------------------------------------
<StructLayout(LayoutKind.Sequential)> _
Public Structure CAUUID
  Public cElems As Integer
  Public pElems As IntPtr
End Structure
