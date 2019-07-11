Imports System.Runtime.InteropServices

<Guid("AB634003-F13D-11d0-A459-004095E1DAEA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
Public Interface ILanguageExtender
    Sub RegisterExtensionAs(ByRef bstrExtensionName As String)
    Sub GetNProps(ByRef plProps As Integer)
    Sub FindProp(ByVal bstrPropName As String, ByRef plPropNum As Integer)
    Sub GetPropName(ByVal lPropNum As Integer, ByVal lPropAlias As Integer, ByRef pbstrPropName As String)
    Sub GetPropVal(ByVal lPropNum As Integer, ByRef pvarPropVal As Object)
    Sub SetPropVal(ByVal lPropNum As Integer, ByRef varPropVal As Object)
    Sub IsPropReadable(ByVal lPropNum As Integer, ByRef pboolPropRead As Boolean)
    Sub IsPropWritable(ByVal lPropNum As Integer, ByRef pboolPropWrite As Boolean)
    Sub GetNMethods(ByRef plMethods As Integer)
    Sub FindMethod(ByVal bstrMethodName As String, ByRef plMethodNum As Integer)
    Sub GetMethodName(ByVal lMethodNum As Integer, ByVal lMethodAlias As Integer, ByRef pbstrMethodName As String)
    Sub GetNParams(ByVal lMethodNum As Integer, ByRef plParams As Integer)
    Sub GetParamDefValue(ByVal lMethodNum As Integer, ByVal lParamNum As Integer, ByRef pvarParamDefValue As Object)
    Sub HasRetVal(ByVal lMethodNum As Integer, ByRef pboolRetValue As Boolean)
    Sub CallAsProc(ByVal lMethodNum As Integer, <MarshalAs(UnmanagedType.SafeArray)> ByRef paParams As System.Array)
    Sub CallAsFunc(ByVal lMethodNum As Integer, ByRef pvarRetValue As Object, <MarshalAs(UnmanagedType.SafeArray)> ByRef paParams As System.Array)
End Interface
