Imports System.Runtime.InteropServices
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Diagnostics
Imports System.Threading
Imports Apache.NMS.ActiveMQ
Imports Apache.NMS
Imports System.IO
Imports System.Reflection

'Сгенерируйте уникальный идентификатор компоненты (меню Tools - Create GUID)  
'Уникальный идентификатор в пределах Вселенной за время ее существования. :-) 
'Укажите ProgID компоненты (по этому имени ее будет находить 1С).
'Пример регистрации компоненты в системном реестре, чтобы ее смогла найти 1С: 
'regasm.exe ActiveMQAdapter.dll /codebase 

<ComVisible(True), Guid("CFDB4244-120F-4D07-BCC3-42B2ECAA603A"), ProgId("AddIn.ActiveMQAdapter")> _
Public Class ActiveMQAdapter

    Implements IInitDone
    Implements ILanguageExtender
    Implements ISpecifyPropertyPages
    Const c_AddinName As String = "ActiveMQAdapter"

#Region "IInitDone implementation"
    '/////////////////////////////////////////////////////////////////////////////////
    Public Sub New() ' Обязательно для COM инициализации
        'Вызывается при начале работы внешней компоненты
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Private Sub Init(<MarshalAs(UnmanagedType.IDispatch)> ByVal pConnection As Object) Implements IInitDone.Init
        Trace.AutoFlush = True
        Trace.Listeners.Add(New TextWriterTraceListener("ActiveMQAdapter.log"))
        Trace.Listeners.Add(New EventLogTraceListener("ActiveMQAdapter"))
        Trace.TraceInformation("Начало работы ActiveMQAdapter (C) ООО ППВТИ Михаил Немчинов (mnemchinov@mail.ru).")

        CreateRes("Apache_NMS_ActiveMQ_dll", My.Resources.Apache_NMS_ActiveMQ_dll)
        CreateRes("Apache_NMS_dll", My.Resources.Apache_NMS_dll)
        CreateRes("Ionic_Zlib_dll", My.Resources.Ionic_Zlib_dll)

        ReadConfigSettings()
        V7Data.V7Object = pConnection
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Private Sub Done() Implements IInitDone.Done
        'Вызывается при завершении работы внешней компоненты
        WriteConfigSettings()
        V7Data.V7Object = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Trace.TraceInformation("Завершение работы ActiveMQAdapter (C) ООО ППВТИ Михаил Немчинов (mnemchinov@mail.ru).")
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Private Sub GetInfo(ByRef pInfo() As Object) Implements IInitDone.GetInfo
        pInfo.SetValue("2000", 0)
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub RegisterExtensionAs(ByRef bstrExtensionName As String) Implements ILanguageExtender.RegisterExtensionAs
        bstrExtensionName = c_AddinName
    End Sub

#End Region

#Region "Переменные"

    Public Shared sServer As String 'Сервер
    Public Shared sPort As String 'Порт
    Public Shared sNameQueueIn As String 'Имя очереди входящих сообщений
    Public Shared sNameQueueOut As String 'Имя очереди исходящих сообщений
    Public sMessage As String = "" 'Текст сообщения

    Dim conFactory As ConnectionFactory
    Dim connection As IConnection
    Dim session As ISession
    Dim inbox As IMessageConsumer
    Dim inboxQ As IQueue
    Dim outbox As IMessageProducer
    Dim outboxQ As IQueue

#End Region

#Region "Внутренние функции"

    Private Sub CreateRes(ByVal strResName As String, ByVal strResValue() As Byte)
        Dim strFileName = Replace(strResName, "_", ".")
        Dim tempFile As String = New FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName & "\" & strFileName
        If Dir$(tempFile) = vbNullString Then
            Using fs As New FileStream(tempFile, FileMode.Create)
                fs.Write(strResValue, 0, strResValue.Length)
                fs.Close()
            End Using
        End If
    End Sub

    Public Shared Sub ReadConfigSettings()
        sServer = ConfigSettings.ReadSetting("Server")
        sPort = ConfigSettings.ReadSetting("Port")
        sNameQueueIn = ConfigSettings.ReadSetting("NameQueueIn")
        sNameQueueOut = ConfigSettings.ReadSetting("NameQueueOut")
    End Sub

    Public Shared Sub WriteConfigSettings()
        ConfigSettings.WriteSetting("Server", sServer)
        ConfigSettings.WriteSetting("Port", sPort)
        ConfigSettings.WriteSetting("NameQueueIn", sNameQueueIn)
        ConfigSettings.WriteSetting("NameQueueOut", sNameQueueOut)
    End Sub

    Public Function ConnectionIsActive() As Boolean
        Try
            ConnectionIsActive = connection.IsStarted
        Catch ex As Exception 'Обработчик исключительных ситуаций (ошибок)
            ConnectionIsActive = False
        End Try
    End Function

    Public Function ConnectToActiveMQ(ByVal strServer As String, ByVal strPort As String, ByVal strQueueNameOut As String, ByVal strQueueNameIn As String) As String
        Try
            conFactory = New ConnectionFactory("tcp://" & strServer & ":" & strPort)
            connection = conFactory.CreateConnection()
            connection.Start()
            session = connection.CreateSession(AcknowledgementMode.Transactional)
            outboxQ = session.GetQueue(strQueueNameOut)
            outbox = session.CreateProducer(outboxQ)
            inboxQ = session.GetQueue(strQueueNameIn)
            inbox = session.CreateConsumer(inboxQ)
            'Trace.TraceInformation("Соединение с Active MQ установлено." & vbCrLf & "tcp://" & strServer & ":" & strPort)
            ConnectToActiveMQ = "OK!"
        Catch ex As Exception 'Обработчик исключительных ситуаций (ошибок)
            Raise1CException(ex.Message, "ConnectToActiveMQ()", ex.ToString)
            ConnectToActiveMQ = "ERR:ActiveMQAdapter:ConnectToActiveMQ:" & ex.Message
        End Try
    End Function

    Public Function DisConnectFromActiveMQ() As String
        Try
            connection.Stop()
            connection.Close()
            DisConnectFromActiveMQ = "OK!"
            'Trace.TraceInformation("Завершение работы с Active MQ...")
        Catch ex As Exception 'Обработчик исключительных ситуаций (ошибок)
            Raise1CException(ex.Message, "DisConnectFromActiveMQ()", ex.ToString)
            DisConnectFromActiveMQ = "ERR:ActiveMQAdapter:DisConnectFromActiveMQ:" & ex.Message
        End Try
    End Function

    Public Function GetMessage() As String
        Dim message As IMessage
        GetMessage = "No messages"
        Try
            message = inbox.Receive(TimeSpan.FromSeconds(1))
            If message IsNot Nothing Then
                If TypeOf message Is ITextMessage Then
                    GetMessage = DirectCast(message, ITextMessage).Text
                End If
            End If
            session.Commit()
        Catch ex As Exception 'Обработчик исключительных ситуаций (ошибок)
            Raise1CException(ex.Message, "GetMessage()", ex.ToString)
            GetMessage = "ERR:ActiveMQAdapter:GetMessage:" & ex.Message
        End Try
    End Function

    Public Function SendMessage(ByVal strMessage As String) As String
        Dim message As IMessage
        Try
            message = outbox.CreateTextMessage(strMessage)
            outbox.Send(message)
            session.Commit()
            SendMessage = "OK!"
        Catch ex As Exception 'Обработчик исключительных ситуаций (ошибок)
            Raise1CException(ex.Message, "SendMessage()", ex.ToString)
            SendMessage = "ERR:ActiveMQAdapter:SendMessage:" & ex.Message
        End Try
    End Function
#End Region

#Region "ISpecifyPropertyPage implementation"
    Public Sub GetPages(ByRef pPages As CAUUID) Implements ISpecifyPropertyPages.GetPages
        pPages.cElems = 1
        pPages.pElems = Marshal.AllocCoTaskMem(Marshal.SizeOf(GetType(Guid)))
        ' GUID совпадает с ClientPage!
        Marshal.StructureToPtr(New Guid("70523FDD-E0E3-4518-BC87-BEA3868FB977"), pPages.pElems, False)
    End Sub
#End Region

#Region "Свойства"
    '/////////////////////////////////////////////////////////////////////////////////
    Enum Props
        'Числовые идентификаторы свойств внешней компоненты
        propServer = 0 'Сервер
        propPort = 1 'Порт
        propNameQueueIn = 2 'Имя очереди входящих сообщений
        propNameQueueOut = 3 'Имя очереди исходящих сообщений
        propMessage = 4 'Текст сообщения
        propConnectionIsActive = 5 'Текст сообщения
        LastProp = 6
    End Enum

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetNProps(ByRef plProps As Integer) Implements ILanguageExtender.GetNProps
        'Здесь 1С получает количество доступных из ВК свойств
        plProps = Props.LastProp
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub FindProp(ByVal bstrPropName As String, ByRef plPropNum As Integer) Implements ILanguageExtender.FindProp
        'Здесь 1С ищет числовой идентификатор свойства по его текстовому имени
        Select Case bstrPropName
            Case "Server", "Сервер"
                plPropNum = Props.propServer
            Case "Port", "Порт"
                plPropNum = Props.propPort
            Case "NameQueueIn", "ИмяОчередиВходящихСообщений"
                plPropNum = Props.propNameQueueIn
            Case "NameQueueOut", "ИмяОчередиИсходящихСообщений"
                plPropNum = Props.propNameQueueOut
            Case "Message", "Сообщение"
                plPropNum = Props.propMessage
            Case "ConnectionIsActive", "СоединениеАктивно"
                plPropNum = Props.propConnectionIsActive
            Case Else
                plPropNum = -1
        End Select
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetPropName(ByVal lPropNum As Integer, ByVal lPropAlias As Integer, ByRef pbstrPropName As String) Implements ILanguageExtender.GetPropName
        'Здесь 1С (теоретически) узнает имя свойства по его идентификатору. lPropAlias - номер псевдонима
        pbstrPropName = ""
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    'Функция генерирует исключение в 1С
    Public Sub Raise1CException(ByVal strMessage As String, Optional ByVal strOwner As String = "", Optional ByVal strDetails As String = "")
        Dim ei As ExcepInfo
        Dim StrErrMessage As String
        ei.wCode = 1006 'Вид пиктограммы
        '1000 - нет значка
        '1001 - обычный значок
        '1002 - красный значок !
        '1003 - красный значок !!
        '1004 - красный значок !!!
        '1005 - зеленый значок i
        '1006 - красный значок err
        '1007 - Окно предупреждения "Внимание"
        '1008 - Окно предупреждения "Информация"
        '1009 - Окно предупреждения "Ошибка"

        StrErrMessage = "ERR:" & strOwner.Trim & ":" & strMessage
        Trace.TraceError(StrErrMessage & vbCrLf & strDetails)
        ei.scode = 1 'Генерируем ошибку времени исполнения
        ei.bstrDescription = StrErrMessage 'Сообщение
        ei.bstrSource = c_AddinName

        V7Data.ErrorLog.AddError(c_AddinName, ei)

    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetPropVal(ByVal lPropNum As Integer, ByRef pvarPropVal As Object) Implements ILanguageExtender.GetPropVal
        'Здесь 1С узнает значения свойств 

        Try
            pvarPropVal = Nothing
            Select Case lPropNum

                Case Props.propServer
                    pvarPropVal = sServer

                Case Props.propPort
                    pvarPropVal = sPort

                Case Props.propNameQueueIn
                    pvarPropVal = sNameQueueIn

                Case Props.propNameQueueOut
                    pvarPropVal = sNameQueueOut

                Case Props.propMessage
                    pvarPropVal = sMessage

                Case Props.propConnectionIsActive
                    pvarPropVal = ConnectionIsActive()

            End Select
        Catch ex As Exception 'Обработчик исключительных ситуаций (ошибок)
            Raise1CException(ex.Message, "GetPropVal()", ex.ToString)
        End Try
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub SetPropVal(ByVal lPropNum As Integer, ByRef varPropVal As Object) Implements ILanguageExtender.SetPropVal
        'Здесь 1С изменяет значения свойств 
        Select Case lPropNum
            Case Props.propServer
                sServer = CStr(varPropVal)

            Case Props.propPort
                sPort = CStr(varPropVal)

            Case Props.propNameQueueIn
                sNameQueueIn = CStr(varPropVal)

            Case Props.propNameQueueOut
                sNameQueueOut = CStr(varPropVal)

            Case Props.propMessage
                sMessage = CStr(varPropVal)

        End Select
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub IsPropReadable(ByVal lPropNum As Integer, ByRef pboolPropRead As Boolean) Implements ILanguageExtender.IsPropReadable
        'Здесь 1С узнает, какие свойства доступны для чтения

        pboolPropRead = True ' Все свойства доступны для чтения
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub IsPropWritable(ByVal lPropNum As Integer, ByRef pboolPropWrite As Boolean) Implements ILanguageExtender.IsPropWritable
        'Здесь 1С узнает, какие свойства доступны для записи
        Select Case lPropNum
            Case Props.propConnectionIsActive
                pboolPropWrite = False
            Case Else
                pboolPropWrite = True ' Все свойства доступны для записи
        End Select
    End Sub

#End Region

#Region "Методы"

    '/////////////////////////////////////////////////////////////////////////////////
    Enum Methods
        'Числовые идентификаторы методов (процедур или функций) внешней компоненты
        methGetMessage = 0 'Получить сообщение
        methSendMessage = 1 'Отправить сообщение
        methConnectToActiveMQ = 2 'Отправить сообщение
        methDisConnectFromActiveMQ = 3 'Отправить сообщение

        LastMethod = 4
    End Enum

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetNMethods(ByRef plMethods As Integer) Implements ILanguageExtender.GetNMethods
        plMethods = Methods.LastMethod
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub FindMethod(ByVal bstrMethodName As String, ByRef plMethodNum As Integer) Implements ILanguageExtender.FindMethod
        'Здесь 1С получает числовой идентификатор метода (процедуры или функции) по имени (названию) процедуры или функции

        plMethodNum = -1
        Select Case bstrMethodName
            Case "GetMessage", "ПолучитьСообщение"
                plMethodNum = Methods.methGetMessage
            Case "SendMessage", "ОтправитьСообщение"
                plMethodNum = Methods.methSendMessage
            Case "ConnectToActiveMQ", "НачатьСессию"
                plMethodNum = Methods.methConnectToActiveMQ
            Case "DisConnectFromActiveMQ", "ЗавершитьСессию"
                plMethodNum = Methods.methDisConnectFromActiveMQ
        End Select
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetMethodName(ByVal lMethodNum As Integer, ByVal lMethodAlias As Integer, ByRef pbstrMethodName As String) Implements ILanguageExtender.GetMethodName
        'Здесь 1С (теоретически) получает имя метода по его идентификатору. lMethodAlias - номер синонима.
        pbstrMethodName = ""
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetNParams(ByVal lMethodNum As Integer, ByRef plParams As Integer) Implements ILanguageExtender.GetNParams
        'Здесь 1С получает количество параметров у метода (процедуры или функции)

        Select Case lMethodNum
            Case Methods.methGetMessage
                plParams = 0
            Case Methods.methSendMessage
                plParams = 0
            Case Methods.methConnectToActiveMQ
                plParams = 0
            Case Methods.methDisConnectFromActiveMQ
                plParams = 0
        End Select
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetParamDefValue(ByVal lMethodNum As Integer, ByVal lParamNum As Integer, ByRef pvarParamDefValue As Object) Implements ILanguageExtender.GetParamDefValue
        'Здесь 1С получает значения параметров процедуры или функции по умолчанию

        pvarParamDefValue = Nothing 'Нет значений по умолчанию
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub HasRetVal(ByVal lMethodNum As Integer, ByRef pboolRetValue As Boolean) Implements ILanguageExtender.HasRetVal
        'Здесь 1С узнает, возвращает ли метод значение (т.е. является процедурой или функцией)

        pboolRetValue = True  'Все методы у нас будут функциями (т.е. будут возвращать значение). 
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub CallAsProc(ByVal lMethodNum As Integer, ByRef paParams As System.Array) Implements ILanguageExtender.CallAsProc
        'Здесь внешняя компонента выполняет код процедур. А процедур у нас нет.
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub CallAsFunc(ByVal lMethodNum As Integer, ByRef pvarRetValue As Object, _
    ByRef paParams As System.Array) _
    Implements ILanguageExtender.CallAsFunc

        'Здесь внешняя компонента выполняет код функций.

        Try
            pvarRetValue = 0 'Возвращаемое значение метода для 1С
            Select Case lMethodNum 'Порядковый номер метода
                '//////////////////////////////////////////////////////////
                Case Methods.methGetMessage  'Реализуем метод для получения сообщения
                    pvarRetValue = GetMessage()
                Case Methods.methSendMessage  'Реализуем метод отправки сообщения
                    If sMessage.Length = 0 Then
                        pvarRetValue = "ERR:ActiveMQAdapter:SendMessage:Не задан текст сообщения"
                        Raise1CException("Не задан текст сообщения", "SendMessage()")
                    Else
                        pvarRetValue = SendMessage(sMessage)
                    End If
                Case Methods.methConnectToActiveMQ  'Реализуем соединение с Active MQ
                    If sServer.Length = 0 Or sPort.Length = 0 Or sNameQueueOut.Length = 0 Or sNameQueueIn.Length = 0 Then
                        pvarRetValue = "ERR:ActiveMQAdapter:ConnectToActiveMQ:Не задан один из параметров соединения"
                        Raise1CException("Не задан один из параметров соединения" & vbCrLf & "Server = " & sServer & vbCrLf & "Port = " & sPort & vbCrLf & "NameQueueOut = " & sNameQueueOut & vbCrLf & "NameQueueIn = " & sNameQueueIn, "ConnectToActiveMQ()")
                    Else
                        pvarRetValue = ConnectToActiveMQ(sServer, sPort, sNameQueueOut, sNameQueueIn)
                    End If
                Case Methods.methDisConnectFromActiveMQ  'Реализуем отсоединение от Active MQ
                    pvarRetValue = DisConnectFromActiveMQ()
            End Select

        Catch ex As Exception 'Обрабатываем исключение (ошибку)
            Raise1CException(ex.Message, "CallAsFunc()", ex.ToString)
        End Try
    End Sub
#End Region
End Class
