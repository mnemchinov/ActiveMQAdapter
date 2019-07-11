Imports Apache.NMS.ActiveMQ
Imports Apache.NMS
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Diagnostics
Imports System.Threading


Module SendGetMessageFromToActiveMQ
    Private conFactory As ConnectionFactory
    Private connection As IConnection
    Private session As ISession
    Private inbox As IMessageConsumer
    Private inboxQ As IQueue
    Private outbox As IMessageProducer
    Private outboxQ As IQueue

    Public Function ConnectionIsActive() As Boolean
        Try
            ConnectionIsActive = connection.IsStarted
        Catch ex As Exception 'Обработчик исключительных ситуаций (ошибок)
            'Trace.TraceError("ERR:ConnectionIsActive(): " & ex.Message & vbCrLf & ex.ToString)
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
            ActiveMQAdapter.Raise1CException("Не задан один из параметров соединения", "ConnectToActiveMQ()", ex)
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
            Trace.TraceError("ERR:DisConnectFromActiveMQ(): " & ex.Message & vbCrLf & ex.ToString)
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
            MsgBox("ERR:GetMessage(): " & Trace.Listeners.Count.ToString)
            Trace.TraceError("ERR:GetMessage(): " & ex.Message & vbCrLf & ex.ToString)
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
            Trace.TraceError("ERR:SendMessage(): " & ex.Message & vbCrLf & ex.ToString)
            SendMessage = "ERR:ActiveMQAdapter:SendMessage:" & ex.Message
        End Try
    End Function
End Module
