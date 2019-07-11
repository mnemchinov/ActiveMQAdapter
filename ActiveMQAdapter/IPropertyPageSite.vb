Imports System.Runtime.InteropServices

<Guid("B196B28C-BAB4-101A-B69C-00AA00341D07"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
Public Interface IPropertyPageSite
  Sub OnStatusChange(ByVal dwFlags As Integer)
  Sub GetLocaleID(ByRef pLocaleID As Integer)
  Sub GetPageContainer(<MarshalAs(UnmanagedType.IUnknown)> ByRef ppUnk As Object)
  Sub TranslateAccelerator(ByRef pMsg As tagMSG)
End Interface
