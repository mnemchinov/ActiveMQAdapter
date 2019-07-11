Imports System
Imports System.Xml
Imports System.Reflection

Public Class ConfigSettings
    Public Shared Function ReadSetting(ByVal key As String) As String
        Dim doc As XmlDocument = loadConfigDocument()
        Dim node As XmlNode = doc.SelectSingleNode("//appSettings")
        If node Is Nothing Then
            Throw New InvalidOperationException("appSettings section not found in config file.")
        End If
        Try
            Dim elem As XmlElement = DirectCast(node.SelectSingleNode(String.Format("//add[@key='{0}']", key)), XmlElement)
            If Not (elem Is Nothing) Then
                Return elem.GetAttribute("value")
            Else
                Return ""
            End If
        Catch
            Throw
        End Try
    End Function

    Public Shared Sub WriteSetting(ByVal key As String, ByVal value As String)
        Dim doc As XmlDocument = loadConfigDocument()
        Dim node As XmlNode = doc.SelectSingleNode("//appSettings")
        If node Is Nothing Then
            Throw New InvalidOperationException("appSettings section not found in config file.")
        End If
        Try
            Dim elem As XmlElement = DirectCast(node.SelectSingleNode(String.Format("//add[@key='{0}']", key)), XmlElement)
            If Not (elem Is Nothing) Then
                elem.SetAttribute("value", value)
            Else
                elem = doc.CreateElement("add")
                elem.SetAttribute("key", key)
                elem.SetAttribute("value", value)
                node.AppendChild(elem)
            End If
            doc.Save(getConfigFilePath())
        Catch
            Throw
        End Try
    End Sub

    Public Shared Sub RemoveSetting(ByVal key As String)
        Dim doc As XmlDocument = loadConfigDocument()
        Dim node As XmlNode = doc.SelectSingleNode("//appSettings")
        Try
            If node Is Nothing Then
                Throw New InvalidOperationException("appSettings section not found in config file.")
            Else
                node.RemoveChild(node.SelectSingleNode(String.Format("//add[@key='{0}']", key)))
                doc.Save(getConfigFilePath())
            End If
        Catch e As NullReferenceException
            Throw New Exception(String.Format("The key {0} does not exist.", key), e)
        End Try
    End Sub

    Private Shared Function loadConfigDocument() As XmlDocument
        Dim doc As XmlDocument = Nothing
        Try
            doc = New XmlDocument()
            doc.Load(getConfigFilePath())
            Return doc
        Catch e As System.IO.FileNotFoundException
            Throw New Exception("No configuration file found.", e)
        End Try
    End Function

    Private Shared Function getConfigFilePath() As String
        Return Assembly.GetExecutingAssembly().Location + ".config"
    End Function
End Class