Imports System.Runtime.InteropServices

Friend Class WinAPI
  Public Const SW_HIDE As Integer = 0

  <DllImport("User32.dll", EntryPoint:="SetParent")> _
  Public Shared Function SetParent(ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As IntPtr
  End Function
End Class
