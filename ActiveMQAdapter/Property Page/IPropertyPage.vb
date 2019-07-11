Imports System.Runtime.InteropServices

<Guid("B196B28D-BAB4-101A-B69C-00AA00341D07"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
Public Interface IPropertyPage
  Sub SetPageSite(<[In]()> ByVal pPageSite As IPropertyPageSite)
  Sub Activate(ByVal hwndParent As IntPtr, ByRef pRect As tagRECT, ByVal bModal As Boolean)
  Sub Deactivate()
  Sub GetPageInfo(<Out()> ByRef pPageInfo As propPageInfo)
  Sub SetObjects(ByVal cObjects As Integer, <MarshalAs(UnmanagedType.IUnknown)> ByRef ppUnk As Object)
  Sub Show(ByVal cmdShow As Integer)
  Sub Move(ByRef pRect As tagRECT)
  Sub IsPageDirty()
  Sub Apply()
  Sub Help(ByVal pszHelpDir As String)
  Sub TranslateAccelerator(ByRef pMsg As tagMSG)
End Interface
'----------------------------------------------------------------------------
<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
Public Structure propPageInfo
  Public cb As Integer
  Public pszTitle As String
  Public cx As Integer
  Public cy As Integer
  Public pszDocString As String
  Public pszHelpFile As String
  Public dwHelpContext As Integer
End Structure
'----------------------------------------------------------------------------
<StructLayout(LayoutKind.Sequential)> _
Public Structure tagRECT
  Public left As Integer
  Public top As Integer
  Public right As Integer
  Public bottom As Integer
End Structure
'---------------------------------------------------------------------------
<StructLayout(LayoutKind.Sequential)> _
Public Structure tagMSG
  Public hwnd As IntPtr
  Public message As Integer
  Public wParam As Integer
  Public lParam As Integer
  Public time As Integer
  Public x As Integer
  Public y As Integer
End Structure