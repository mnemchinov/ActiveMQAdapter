Imports System.Runtime.InteropServices

<Guid("3127CA40-446E-11CE-8135-00AA004BB851"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
Public Interface IErrorLog
  Sub AddError(ByVal pszPropName As String, ByRef pExepInfo As ExcepInfo)
End Interface
'----------------------------------------------------------
<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=8)> _
Public Structure ExcepInfo
    Public wCode As Short
    Public wReserved As Short
    <MarshalAs(UnmanagedType.BStr)> Public bstrSource As String
    <MarshalAs(UnmanagedType.BStr)> Public bstrDescription As String
    <MarshalAs(UnmanagedType.BStr)> Public bstrHelpFile As String
    Public dwHelpContext As Integer
    Public pvReserved As IntPtr
    Public pfnDereffered As IntPtr
    Public scode As Integer
End Structure