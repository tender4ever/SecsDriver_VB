Imports System
Imports System.Text
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports SecsDriver_VB.SocketDriver
Imports System.Threading
Imports System.Collections.Concurrent

Namespace SecsDriver


#Region "Delegate 委派"

    ' 當 SECS 連線狀態改變時，外部程式使用此委派來取得資訊
    Public Delegate Sub delegateSecsConnectStateChange(ByRef aMessage As String)

    ' 當送出 Primary Message 時，外部程式使用此委派來取得 SecsTransaction
    Public Delegate Sub delegatePrimarySent(ByRef transaction As SecsTransaction)

    ' 當取得 Primary Message 時，外部程式使用此委派來取得 SecsTransaction
    Public Delegate Sub delegatePrimaryReceived(ByRef transaction As SecsTransaction)

    ' 當送出 Secondary Message 時，外部程式使用此委派來取得 SecsTransaction
    Public Delegate Sub delegateSecondarySent(ByRef transaction As SecsTransaction)

    ' 當取得 Secondary Message 時，外部程式使用此委派來取得 SecsTransaction
    Public Delegate Sub delegateSecondaryReceived(ByRef transaction As SecsTransaction)

    ' 當發生 SendMessage Error、Transaction Not Found 時，外部程式使用此委派來取得 SecsMessage 以及 Error Message
    Public Delegate Sub delegateMessageError(ByRef message As SecsMessage, ByVal errorMessage As String)

    ' 當發生 Message Not Found 時，外部程式使用此委派來取得 SecsMessage 以及 Error Message
    Public Delegate Sub delegateMessageInfo(ByRef message As SecsMessage, ByVal aMessage As String)

    ' 當發生錯誤時，外部程式使用此委派來取得錯誤訊息
    Public Delegate Sub delegateError(ByRef aMessage As String)

    ' 當發生 Timeout 時，外部程式使用此委派來取得訊息
    Public Delegate Sub delegateTimeout(ByRef aMessage As String)

#End Region


    ''' <summary>
    ''' SECS Driver
    ''' </summary>
    Public Class SecsWrapper : Implements IFMessageListener


#Region "Public 屬性"

        Public siniFile As IniFile                                  ' Ini File
        Public sSXML As SXML                                        ' SXML

#End Region


#Region "Private 屬性"

        Private sServer As Server                                   ' Socket Server
        Private sClient As Client                                   ' Socket Client

        Private sTimeout As Timeout                                 ' Timeout
        Private sTimeoutList As List(Of Timeout)                    ' TimeoutList

        Private sLog As Log                                         ' Log

        Private ReceiveMessageQueue As ConcurrentQueue(Of Byte())                                       ' Receive Message Queue
        Private PrimarySendQueue As ConcurrentQueue(Of SecsMessage)                                     ' Primary Message Send Queue
        Private SecondarySendQueue As ConcurrentQueue(Of SecsTransaction)                               ' Secondary Message Send Queue

        Private Shared SendTransactionMap As ConcurrentDictionary(Of String, SecsTransaction)           ' Send Transaction Map
        Private Shared ReceiveTransactionMap As ConcurrentDictionary(Of String, SecsTransaction)        ' Receive Transaction Map

        Private systemBytesCount As UInteger = Now.Millisecond                                          ' System Bytes 計數


        ''' <summary>
        ''' SECS 連線狀態
        ''' </summary>
        Private SocketState As enumSecsConnectionState



        Private Shared CheckSendTransactionMapLock As Object = New Object       ' Lock Object For Send Transaction Map
        Private Shared CheckReceiveTransactionMapLock As Object = New Object    ' Lock Object For Receive Transaction Map

#End Region


#Region "Event"

        Public Event OnSecsConnectStateChange As delegateSecsConnectStateChange     ' Event For Secs Connect State Change
        Public Event OnPrimarySent As delegatePrimarySent                           ' Event For Primary Sent
        Public Event OnPrimaryReceived As delegatePrimaryReceived                   ' Event For Primary Received
        Public Event OnSecondarySent As delegateSecondarySent                       ' Event For Secondary Sent
        Public Event OnSecondaryReceived As delegateSecondaryReceived               ' Event For Secondary Received
        Public Event OnMessageError As delegateMessageError                         ' Event For Message Error
        Public Event OnMessageInfo As delegateMessageInfo                           ' Event For Message Info
        Public Event OnError As delegateError                                       ' Event For On Error
        Public Event OnTimeout As delegateTimeout                                   ' Event For On Timeout

#End Region


#Region "建構子"

        ''' <summary>
        ''' 建構子
        ''' </summary>
        Public Sub New()

            ReceiveMessageQueue = New ConcurrentQueue(Of Byte())
            PrimarySendQueue = New ConcurrentQueue(Of SecsMessage)
            SecondarySendQueue = New ConcurrentQueue(Of SecsTransaction)
            SendTransactionMap = New ConcurrentDictionary(Of String, SecsTransaction)
            ReceiveTransactionMap = New ConcurrentDictionary(Of String, SecsTransaction)
            sTimeoutList = New List(Of Timeout)
            sLog = New Log()

        End Sub

#End Region


#Region "Public Method"

        ''' <summary>
        ''' 開始連線
        ''' </summary>
        Public Sub Open()

            ' IniFile Path
            Dim path As String = Application.StartupPath

            ' 載入 iniFile
            Try
                siniFile = New IniFile(path & "\config.ini")
            Catch ex As Exception

                RaiseEvent OnError("Load IniFile Error !")
            End Try

            ' 載入 SXML
            Try
                sSXML = New SXML(siniFile.SXMLPath)

                sSXML.LoadMessage()
            Catch ex As Exception

                RaiseEvent OnError("Load SXML Error !")
            End Try


            ' 開啟 SECS
            If siniFile IsNot Nothing And sSXML IsNot Nothing Then

                SocketState = enumSecsConnectionState.sNone
                RaiseEvent OnSecsConnectStateChange("None")
                WatchSecsConnectionStatus()

            End If

        End Sub


        ''' <summary>
        ''' 關閉連線
        ''' </summary>
        Public Sub Close()

            Try
                SocketState = enumSecsConnectionState.sSeparate

                ' 如果是 Active，則使用 Client 來關閉連線
                ' 如果是 Passive，則使用 Server 來關閉連線
                If Me.siniFile.Entity = enumSecsEntity.sActive Then
                    sClient.disconnect()
                Else
                    sServer.disconnect()
                End If

            Catch ex As Exception

                RaiseEvent OnError("Disconnect Fail !")

            End Try

        End Sub


        ''' <summary>
        ''' 使用 Message Name 找 Message
        ''' </summary>
        ''' <param name="aMessageName"></param>
        ''' <returns></returns>
        Public Function GetMessageByName(ByRef aMessageName As String) As SecsMessage

            Try
                ' 使用 Message Name 找 SecsMessage
                If sSXML.FindMessageByMessageName(aMessageName) IsNot Nothing Then

                    Return sSXML.FindMessageByMessageName(aMessageName)

                End If

                ' 找不到 Message 則 Return Nothing
                Return Nothing

            Catch ex As Exception

                Return Nothing

            End Try

        End Function


        ''' <summary>
        ''' 供外部使用的 Send Primary
        ''' </summary>
        ''' <param name="aSecsMessage"></param>
        Public Sub SendPrimary(ByRef aSecsMessage As SecsMessage)

            PrimarySendQueue.Enqueue(aSecsMessage)

        End Sub


        ''' <summary>
        ''' 供外部使用的 Send Secondary
        ''' </summary>
        ''' <param name="aSecsTransaction"></param>
        Public Sub SendSecondary(ByRef aSecsTransaction As SecsTransaction)

            SecondarySendQueue.Enqueue(aSecsTransaction)

        End Sub


#End Region


#Region "Private Method"

        ''' <summary>
        ''' 產生 Thread 來檢查連線狀況、SECS Message處理狀況、Timeout狀況、IniFile 是否有變動
        ''' </summary>
        Private Sub WatchSecsConnectionStatus()

            ' NEW 一個 Thread 來持續監看 SECS 連線狀態
            Dim sWatchSecsConnectThread As New Thread(New ThreadStart(AddressOf SecsConnect))
            sWatchSecsConnectThread.IsBackground = True
            sWatchSecsConnectThread.Start()

            ' New 一個 Thread 來處理 Queue
            Dim sScheduleThread As New Thread(New ThreadStart(AddressOf Schedule))
            sScheduleThread.IsBackground = True
            sScheduleThread.Start()

            ' New 一個 Thread 來持續檢查是否有 Timeout
            Dim sWatchTimeoutThread As New Thread(New ThreadStart(AddressOf CheckTimeout))
            sWatchTimeoutThread.IsBackground = True
            sWatchTimeoutThread.Start()

            ' New 一個 Thread 來持續檢查 IniFile 是否有變動
            Dim sWatchIniFileThread As New Thread(New ThreadStart(AddressOf CheckIniFile))
            sWatchIniFileThread.IsBackground = True
            sWatchIniFileThread.Start()

        End Sub


        ''' <summary>
        ''' 檢查 SECS 連線狀況
        ''' </summary>
        Private Sub SecsConnect()

            Dim sendSelectRequestCountdown As Integer = 0           ' 幾秒鐘後送出 Select Request
            Dim sendLinktestRequestCountdown As Integer = 0         ' 幾秒鐘後送出 Linktest Request

            Do
                Select Case SocketState

                    ' SECS 連線狀態 = None
                    Case enumSecsConnectionState.sNone

                        ' 如果是 Active，則使用 Client
                        ' 如果是 Passive，則使用 Server
                        If Me.siniFile.Entity = enumSecsEntity.sActive Then

                            sClient = New Client(Me, siniFile.IP, siniFile.Port)
                        Else
                            sServer = New Server(Me, siniFile.IP, siniFile.Port, False)
                        End If

                        Application.DoEvents()
                        Thread.Sleep(1)


                    ' SECS 連線狀態 = Not Connected
                    Case enumSecsConnectionState.sNotConnected

                        ' 如果是 Active，則由 T5 Timeout 來重新連線
                        ' 如果是 Passive，則使用 Server 來重新連線
                        If Me.siniFile.Entity = enumSecsEntity.sActive Then

                            ' 由 T5 Timeout 來重新連線
                        Else
                            sServer.Reconnect()
                        End If

                        Application.DoEvents()
                        Thread.Sleep(1000)


                    ' SECS 連線狀態 = Not Selected
                    Case enumSecsConnectionState.sNotSelected

                        If Me.siniFile.Entity = enumSecsEntity.sActive Then

                            ' NEW 一個 Select Request
                            Dim tempSecsMessage As SecsMessage = New SecsMessage()
                            tempSecsMessage.messageFormat.Length = BitConverter.GetBytes(10)                        ' Length
                            Array.Reverse(tempSecsMessage.messageFormat.Length)
                            tempSecsMessage.messageFormat.DeviceID(0) = Convert.ToByte(&HFF)                        ' DeviceID
                            tempSecsMessage.messageFormat.DeviceID(1) = Convert.ToByte(&HFF)
                            tempSecsMessage.messageFormat.HeaderByte(0) = Convert.ToByte(&H0)                       ' HeaderByte
                            tempSecsMessage.messageFormat.HeaderByte(1) = Convert.ToByte(&H0)
                            tempSecsMessage.messageFormat.PType = Convert.ToByte(&H0)                               ' PType
                            tempSecsMessage.messageFormat.SType = Convert.ToByte(&H1)                               ' SType
                            tempSecsMessage.messageFormat.SystemBytes = BitConverter.GetBytes(getSystemByte())      ' SystemBytes
                            Array.Reverse(tempSecsMessage.messageFormat.SystemBytes)
                            tempSecsMessage.secsItem = Nothing

                            SendPrimary(tempSecsMessage)

                        End If

                        ' 設定 T7 Timeout
                        SyncLock sTimeoutList

                            If sTimeoutList.Count > 0 Then

                                Dim temp As Integer = 0
                                For i As Integer = 0 To sTimeoutList.Count - 1
                                    If sTimeoutList(i).sTimeoutType = enumTimeout.T7 Then
                                        temp = temp + 1
                                    End If
                                Next

                                If temp = 0 Then
                                    sTimeout = New Timeout(siniFile.T7Timeout, enumTimeout.T7)
                                    sTimeoutList.Add(sTimeout)
                                End If
                            Else
                                sTimeout = New Timeout(siniFile.T7Timeout, enumTimeout.T7)
                                sTimeoutList.Add(sTimeout)
                            End If

                        End SyncLock

                        Application.DoEvents()
                        Thread.Sleep(1000)


                    ' SECS 連線狀態 = Selected
                    Case enumSecsConnectionState.sSelected

                        SyncLock sTimeoutList

                            ' 解除 T7 Timeout
                            If sTimeoutList.Count > 0 Then

                                For i As Integer = 0 To sTimeoutList.Count - 1 Step +1

                                    If sTimeoutList(i).sTimeoutType = enumTimeout.T7 Then

                                        sTimeoutList(i) = Nothing
                                        sTimeoutList.RemoveAt(i)

                                    End If
                                Next

                            End If

                        End SyncLock

                        If sendLinktestRequestCountdown = 0 Then

                            ' NEW 一個 Linktest Request
                            Dim tempSecsMessage As SecsMessage = New SecsMessage()
                            tempSecsMessage.messageFormat.Length = BitConverter.GetBytes(10)                        ' Length
                            Array.Reverse(tempSecsMessage.messageFormat.Length)
                            tempSecsMessage.messageFormat.DeviceID(0) = Convert.ToByte(&HFF)                        ' DeviceID
                            tempSecsMessage.messageFormat.DeviceID(1) = Convert.ToByte(&HFF)
                            tempSecsMessage.messageFormat.HeaderByte(0) = Convert.ToByte(&H0)                       ' HeaderByte
                            tempSecsMessage.messageFormat.HeaderByte(1) = Convert.ToByte(&H0)
                            tempSecsMessage.messageFormat.PType = Convert.ToByte(&H0)                               ' PType
                            tempSecsMessage.messageFormat.SType = Convert.ToByte(&H5)                               ' SType
                            tempSecsMessage.messageFormat.SystemBytes = BitConverter.GetBytes(getSystemByte())      ' SystemBytes
                            Array.Reverse(tempSecsMessage.messageFormat.SystemBytes)
                            tempSecsMessage.secsItem = Nothing
                            SendPrimary(tempSecsMessage)

                            sendLinktestRequestCountdown = 15
                        End If

                        sendLinktestRequestCountdown = sendLinktestRequestCountdown - 1

                        Application.DoEvents()
                        Thread.Sleep(1000)


                    ' SECS 連線狀態 = Separate
                    Case enumSecsConnectionState.sSeparate

                        Application.DoEvents()
                        Thread.Sleep(1000)

                End Select

            Loop While Not SocketState = enumSecsConnectionState.sSeparate

        End Sub


        ''' <summary>
        ''' 處理 Message Queue
        ''' </summary>
        Private Sub Schedule()

            Do
                ' 檢查 Primary Send Queue 是否 > 0
                ' 代表有 Primary Message 需要送出
                If PrimarySendQueue.Count > 0 Then

                    ' 送出一個 Primary Message
                    Dim aSecsMessage As Object = Nothing
                    PrimarySendQueue.TryDequeue(aSecsMessage)
                    SendPrimary(aSecsMessage)

                End If

                ' 檢查 Send Transaction Map
                CheckSendTransactionMap()

                ' 檢查 Secondary Send Queue 是否 > 0
                ' 代表有 Secondary Message 需要送出
                If SecondarySendQueue.Count > 0 Then

                    ' 送出一個 Secondary Message
                    Dim aSecsTransaction As Object = Nothing
                    SecondarySendQueue.TryDequeue(aSecsTransaction)
                    SendSecondary(aSecsTransaction)

                End If

                ' 檢查 Receive Transaction Map
                CheckReceiveTransactionMap()

                ' 檢查 Receive Message Queue 是否 > 0
                ' 代表有 Message 需要處理
                If ReceiveMessageQueue.Count > 0 Then

                    ' 處理一個 Secondary Message
                    Dim temp As Byte() = Nothing
                    ReceiveMessageQueue.TryDequeue(temp)
                    ParseMessage(temp)

                End If

                Application.DoEvents()
                Thread.Sleep(1)

            Loop While Not SocketState = enumSecsConnectionState.sSeparate

        End Sub


        ''' <summary>
        ''' Send Primary Message
        ''' </summary>
        ''' <param name="sSecsMessage"></param>
        Private Sub SendPrimary(ByVal sSecsMessage As Object)

            Dim aSecsMessage As SecsMessage = New SecsMessage
            aSecsMessage = CType(sSecsMessage, SecsMessage)
            Dim sTransaction As SecsTransaction

            Try
                ' 檢查 System Bytes
                aSecsMessage.messageFormat.SystemBytes = BitConverter.GetBytes(getSystemByte())
                Array.Reverse(aSecsMessage.messageFormat.SystemBytes)

                ' 檢查 Device ID
                If siniFile.Role = enumRole.sHost And aSecsMessage.MessageType = enumMessageType.sDataMessage Then

                    aSecsMessage.messageFormat.DeviceID(0) = Convert.ToByte(siniFile.DeviceID)

                ElseIf siniFile.Role = enumRole.sEquipment And aSecsMessage.MessageType = enumMessageType.sDataMessage Then

                    aSecsMessage.messageFormat.DeviceID(1) = Convert.ToByte(siniFile.DeviceID)

                End If

                ' 檢查 Messae 是否為不需回覆的 Message
                If sSXML.sDontNeedToReplyList.Contains(aSecsMessage.MessageName) = True Then

                    ' 紀錄 Log
                    If siniFile.BinaryLog = True Then
                        sLog.DoBinaryLog(aSecsMessage.ConvertToBinaryLog("SND"))
                    End If

                    If siniFile.TxLog = True Then
                        sLog.DoTxLog(aSecsMessage.ConvertToSML("PrimaryOut"))
                    End If

                    ' NEW 一個 SecsTransaction
                    sTransaction = New SecsTransaction()
                    sTransaction.Primary = aSecsMessage
                    RaiseEvent OnPrimarySent(sTransaction)

                    ' 將 Message 透過 Socket 送出
                    SocketSend(aSecsMessage)

                Else
                    ' New 一個 SecsTransaction，並把 Primary 放在裡面
                    sTransaction = New SecsTransaction()
                    sTransaction.Primary = aSecsMessage

                    ' SendTransactionMap 新增此 Transaction
                    Dim SystemByteString As String = Nothing

                    For i As Integer = 0 To aSecsMessage.messageFormat.SystemBytes.Length - 1
                        SystemByteString = SystemByteString & aSecsMessage.messageFormat.SystemBytes(i)
                    Next

                    ' 將 Transaction 加到 SendTransactionMap
                    SendTransactionMap.TryAdd(SystemByteString, sTransaction)

                End If

            Catch ex As Exception

                RaiseEvent OnMessageError(aSecsMessage, "Send Primary Error")

                ' 刪除 Primary
                If aSecsMessage IsNot Nothing Then
                    aSecsMessage.Dispose()
                End If

                ' 刪除 Transaction
                sTransaction = Nothing

            End Try

        End Sub


        ''' <summary>
        ''' 檢查 Send Transaction Map
        ''' </summary>
        Private Sub CheckSendTransactionMap()

            ' Clone 一份目前的 SendTransactionMap
            Dim CloneSendTransactionMap As Dictionary(Of String, SecsTransaction)
            CloneSendTransactionMap = New Dictionary(Of String, SecsTransaction)(SendTransactionMap)

            ' 檢查 Send Transaction Map 是否 > 0
            If CloneSendTransactionMap.Count > 0 Then

                For Each item In CloneSendTransactionMap

                    Select Case item.Value.State

                        Case enumSecsTransactionState.Create

                            ' 當 Transaction 狀態為 Create 時
                            ' 如果有 Primary Message，代表已送出 Primary Message
                            ' 並將狀態改為 PrimarySent

                            If item.Value.Primary IsNot Nothing Then

                                ' 將 Message 透過 Socket 送出
                                SocketSend(item.Value.Primary)

                                ' 紀錄 Log
                                If siniFile.BinaryLog = True Then
                                    sLog.DoBinaryLog(item.Value.Primary.ConvertToBinaryLog("SND"))
                                End If

                                If siniFile.TxLog = True Then
                                    sLog.DoTxLog(item.Value.Primary.ConvertToSML("PrimaryOut"))
                                End If

                                RaiseEvent OnPrimarySent(item.Value)

                                SendTransactionMap.Item(item.Key).State = enumSecsTransactionState.PrimarySent

                            End If

                            Exit Select

                        Case enumSecsTransactionState.PrimarySent

                            ' 當 Transaction 狀態為 PrimarySent 時
                            ' 代表已經送出 Primary Message
                            ' 如果有收到 Secondary Message 則將狀態改為 SecondaryReceived
                            ' 否則就設定 T3、T6 Timeout，並檢查是否發生 Timeout

                            If item.Value.Secondary IsNot Nothing Then

                                SendTransactionMap.Item(item.Key).sTimeoutList.Clear()
                                SendTransactionMap.Item(item.Key).State = enumSecsTransactionState.SecondaryReceived

                            Else
                                If item.Value.sTimeoutList.Count <= 0 Then

                                    ' 如果是 Data Message，則設定 T3 Timeout 
                                    ' 如果是 Control Message，則設定 T6 Timeout 
                                    If item.Value.Primary.MessageType = enumMessageType.sDataMessage Then

                                        SendTransactionMap.Item(item.Key).SetTimeout(enumTimeout.T3, siniFile)
                                    Else
                                        If item.Value.Primary.MessageType <> enumMessageType.sSeparateRequest Then

                                            SendTransactionMap.Item(item.Key).SetTimeout(enumTimeout.T6, siniFile)

                                        End If

                                    End If

                                Else
                                    ' 檢查各個 Timeout
                                    For i As Integer = 0 To item.Value.sTimeoutList.Count - 1 Step +1

                                        Select Case item.Value.sTimeoutList(i).sTimeoutType

                                            Case enumTimeout.T3

                                                ' 發生 T3 Timeout
                                                If item.Value.sTimeoutList(i).CheckTimeout = True Then

                                                    RaiseEvent OnTimeout("Sent : " & item.Value.Primary.MessageName & " T3 Timeout")

                                                    ' 回覆 S9F9 
                                                    Dim tempSecsMessage As SecsMessage = GetMessageByName("S9F9_TransactionTimerTimeout")
                                                    Dim tempByte As Byte() = ConvertToBytes(item.Value.Primary)
                                                    Dim MessageHeader As Byte() = New Byte(13) {}
                                                    Array.Copy(tempByte, MessageHeader, 14)
                                                    tempSecsMessage.secsItem.SetItemValue(MessageHeader)
                                                    SendPrimary(tempSecsMessage)

                                                    ' 清除 Primary Message
                                                    If SendTransactionMap.Item(item.Key).Primary IsNot Nothing Then
                                                        SendTransactionMap.Item(item.Key).Primary.Dispose()
                                                    End If

                                                    ' 清除所有 Timeout 設定，並移除 Transaction
                                                    SendTransactionMap.Item(item.Key).sTimeoutList(i) = Nothing
                                                    SendTransactionMap.Item(item.Key).sTimeoutList = Nothing
                                                    SendTransactionMap.TryRemove(item.Key, item.Value)

                                                End If

                                            Case enumTimeout.T6

                                                ' 發生 T6 Timeout
                                                If item.Value.sTimeoutList(i).CheckTimeout = True Then

                                                    RaiseEvent OnTimeout("Sent : Control Message T6 Timeout")

                                                    ' 回覆 S9F9
                                                    Dim tempSecsMessage As SecsMessage = GetMessageByName("S9F9_TransactionTimerTimeout")
                                                    Dim tempByte As Byte() = ConvertToBytes(item.Value.Primary)
                                                    Dim MessageHeader As Byte() = New Byte(13) {}
                                                    Array.Copy(tempByte, MessageHeader, 14)
                                                    tempSecsMessage.secsItem.SetItemValue(MessageHeader)
                                                    SendPrimary(tempSecsMessage)

                                                    ' 清除 Primary Message
                                                    If SendTransactionMap.Item(item.Key).Primary IsNot Nothing Then
                                                        SendTransactionMap.Item(item.Key).Primary.Dispose()
                                                    End If

                                                    ' 清除所有 Timeout 設定，並移除 Transaction
                                                    SendTransactionMap.Item(item.Key).sTimeoutList(i) = Nothing
                                                    SendTransactionMap.Item(item.Key).sTimeoutList = Nothing
                                                    SendTransactionMap.TryRemove(item.Key, item.Value)

                                                End If

                                        End Select

                                    Next

                                End If

                            End If

                            Exit Select

                        Case enumSecsTransactionState.SecondaryReceived

                            ' 當 Transaction 狀態為 SecondaryReceived 時
                            ' 代表已經收到 Secondary Message
                            ' 則解除所有 Timeout 設定， 並將狀態改為 WaitForDelete

                            If item.Value.sTimeoutList.Count > 0 Then

                                For i As Integer = 0 To item.Value.sTimeoutList.Count - 1 Step +1

                                    If item.Value.sTimeoutList(i).sTimeoutType = enumTimeout.T3 Then

                                        ' 如果是 Data Message，則解除 T3 Timeout 
                                        SendTransactionMap.Item(item.Key).sTimeoutList(i) = Nothing
                                        SendTransactionMap.Item(item.Key).sTimeoutList.RemoveAt(i)

                                    ElseIf item.Value.sTimeoutList(i).sTimeoutType = enumTimeout.T6 Then

                                        ' 如果是 Control Message，則解除 T6 Timeout
                                        SendTransactionMap.Item(item.Key).sTimeoutList(i) = Nothing
                                        SendTransactionMap.Item(item.Key).sTimeoutList.RemoveAt(i)

                                    End If

                                Next

                            End If

                            SendTransactionMap.Item(item.Key).State = enumSecsTransactionState.WaitForDelete

                            Exit Select

                        Case enumSecsTransactionState.WaitForDelete

                            ' 當 Transaction 狀態為 WaitForDelete
                            ' 代表 Transaction 已完成，並等待刪除

                            ' 清除 Primary Message
                            If SendTransactionMap.Item(item.Key).Primary IsNot Nothing Then
                                SendTransactionMap.Item(item.Key).Primary.Dispose()
                            End If

                            ' 清除 Secondary Message
                            If SendTransactionMap.Item(item.Key).Secondary IsNot Nothing Then
                                SendTransactionMap.Item(item.Key).Secondary.Dispose()
                            End If

                            ' 清除所有 Timeout 設定，並移除 Transaction
                            SendTransactionMap.Item(item.Key).sTimeoutList = Nothing
                            SendTransactionMap.TryRemove(item.Key, item.Value)

                            Exit Select

                    End Select
                Next
            End If

            CloneSendTransactionMap = Nothing

        End Sub


        ''' <summary>
        ''' Send Secondary Message
        ''' </summary>
        ''' <param name="sSecsTransaction"></param>
        Private Sub SendSecondary(ByVal sSecsTransaction As Object)

            Dim aSecsTransaction As SecsTransaction = New SecsTransaction
            aSecsTransaction = CType(sSecsTransaction, SecsTransaction)

            Try
                If ReceiveTransactionMap.Count > 0 Then

                    ' 找出 Transaction
                    Dim SystemByteString As String = Nothing

                    For i As Integer = 0 To aSecsTransaction.Primary.messageFormat.SystemBytes.Length - 1
                        SystemByteString = SystemByteString & aSecsTransaction.Primary.messageFormat.SystemBytes(i)
                    Next

                    Dim tempTransaction As SecsTransaction = ReceiveTransactionMap.Item(SystemByteString)

                    If tempTransaction IsNot Nothing Then

                        ' 檢查 System Bytes
                        aSecsTransaction.Secondary.messageFormat.SystemBytes = aSecsTransaction.Primary.messageFormat.SystemBytes

                        ' 檢查 Device ID
                        If siniFile.Role = enumRole.sHost And aSecsTransaction.Secondary.MessageType = enumMessageType.sDataMessage Then

                            aSecsTransaction.Secondary.messageFormat.DeviceID(0) = Convert.ToByte(siniFile.DeviceID)
                            aSecsTransaction.Secondary.messageFormat.DeviceID(1) = tempTransaction.Primary.messageFormat.DeviceID(1)

                        ElseIf siniFile.Role = enumRole.sEquipment And aSecsTransaction.Secondary.MessageType = enumMessageType.sDataMessage Then

                            aSecsTransaction.Secondary.messageFormat.DeviceID(0) = tempTransaction.Primary.messageFormat.DeviceID(0)
                            aSecsTransaction.Secondary.messageFormat.DeviceID(1) = Convert.ToByte(siniFile.DeviceID)

                        End If

                        tempTransaction.Secondary = aSecsTransaction.Secondary
                    Else
                        RaiseEvent OnMessageError(aSecsTransaction.Secondary, "Transaction Not Found")

                    End If

                Else
                    RaiseEvent OnMessageError(aSecsTransaction.Secondary, "Transaction Not Found")

                End If

            Catch ex As Exception

                RaiseEvent OnMessageError(aSecsTransaction.Secondary, "Send Secondary Error")

                ' 刪除 Secondary
                If aSecsTransaction.Secondary IsNot Nothing Then
                    aSecsTransaction.Secondary.Dispose()
                End If

            End Try

        End Sub


        ''' <summary>
        ''' 檢查 Receive Transaction Map
        ''' </summary>
        Private Sub CheckReceiveTransactionMap()

            ' Clone 一份目前的 SendTransactionMap
            Dim CloneReceiveTransactionMap As Concurrent.ConcurrentDictionary(Of String, SecsTransaction)
            CloneReceiveTransactionMap = New Concurrent.ConcurrentDictionary(Of String, SecsTransaction)(ReceiveTransactionMap)


            ' 檢查 Receive Transaction Map 是否 > 0
            If CloneReceiveTransactionMap.Count > 0 Then

                For Each item In CloneReceiveTransactionMap

                    Select Case item.Value.State

                        Case enumSecsTransactionState.Create

                            ' 當 Transaction 狀態為 Create 時
                            ' 如果有 Primary Message，代表已經收到 Primary Message
                            ' 並將狀態改為 PrimaryReceived 
                            If item.Value.Primary IsNot Nothing Then

                                ReceiveTransactionMap.Item(item.Key).State = enumSecsTransactionState.PrimaryReceived

                            End If

                            Exit Select

                        Case enumSecsTransactionState.PrimaryReceived

                            ' 當狀態為 PrimaryReceived 時
                            ' 代表已經收到 Primary Message
                            ' 如果有 Secondary 時， 則送出 Secondary Message， 並將狀態改為 SecondarySent
                            ' 否則就設定 T3、T6 Timeout，並檢查是否發生 Timeout
                            If item.Value.Secondary IsNot Nothing Then

                                If item.Value.Secondary.messageFormat Is Nothing Then
                                    Exit Select
                                End If

                                ' 將 Message 透過 Socket 送出
                                SocketSend(item.Value.Secondary)

                                ' 紀錄 Log
                                If siniFile.BinaryLog = True Then
                                    sLog.DoBinaryLog(item.Value.Secondary.ConvertToBinaryLog("SND"))
                                End If

                                If siniFile.TxLog = True Then
                                    sLog.DoTxLog(item.Value.Secondary.ConvertToSML("SecondaryOut"))
                                End If

                                RaiseEvent OnSecondarySent(item.Value)

                                ReceiveTransactionMap.Item(item.Key).State = enumSecsTransactionState.SecondarySent

                            Else
                                If item.Value.sTimeoutList.Count <= 0 Then

                                    If item.Value.Primary.MessageType = enumMessageType.sDataMessage Then

                                        ' 如果是 Data Message，則設定 T3 Timeout 
                                        ReceiveTransactionMap.Item(item.Key).SetTimeout(enumTimeout.T3, siniFile)
                                    Else
                                        If item.Value.Primary.MessageType <> enumMessageType.sSeparateRequest Then

                                            ' 如果是 Control Message，則設定 T6 Timeout 
                                            ReceiveTransactionMap.Item(item.Key).SetTimeout(enumTimeout.T6, siniFile)
                                        End If
                                    End If

                                Else
                                    ' 檢查各個 Timeout
                                    For i As Integer = 0 To item.Value.sTimeoutList.Count - 1 Step +1

                                        Select Case item.Value.sTimeoutList(i).sTimeoutType

                                            Case enumTimeout.T3

                                                ' 發生 T3 Timeout
                                                If item.Value.sTimeoutList(i).CheckTimeout = True Then

                                                    RaiseEvent OnTimeout("Received : " & item.Value.Primary.MessageName & " T3 Timeout")

                                                    ' 清除 Primary Message
                                                    If ReceiveTransactionMap.Item(item.Key).Primary IsNot Nothing Then
                                                        ReceiveTransactionMap.Item(item.Key).Primary.Dispose()
                                                    End If

                                                    ' 清除所有 Timeout 設定，並移除 Transaction
                                                    ReceiveTransactionMap.Item(item.Key).sTimeoutList(i) = Nothing
                                                    ReceiveTransactionMap.Item(item.Key).sTimeoutList = Nothing
                                                    ReceiveTransactionMap.TryRemove(item.Key, item.Value)

                                                End If

                                            Case enumTimeout.T6

                                                ' 發生 T6 Timeout
                                                If item.Value.sTimeoutList(i).CheckTimeout = True Then

                                                    RaiseEvent OnTimeout("Received : Control Message T6 Timeout")

                                                    ' 清除 Primary Message
                                                    If ReceiveTransactionMap.Item(item.Key).Primary IsNot Nothing Then
                                                        ReceiveTransactionMap.Item(item.Key).Primary.Dispose()
                                                    End If

                                                    ' 清除所有 Timeout 設定，並移除 Transaction
                                                    ReceiveTransactionMap.Item(item.Key).sTimeoutList(i) = Nothing
                                                    ReceiveTransactionMap.Item(item.Key).sTimeoutList = Nothing
                                                    ReceiveTransactionMap.TryRemove(item.Key, item.Value)

                                                End If

                                        End Select
                                    Next
                                End If
                            End If

                            Exit Select

                        Case enumSecsTransactionState.SecondarySent

                            ' 當 Transaction 狀態為 SecondarySent 時
                            ' 代表已經送出 Secondary Message
                            ' 則解除所有 Timeout 設定， 並將狀態改為 WaitForDelete
                            If item.Value.sTimeoutList.Count > 0 Then

                                For i As Integer = 0 To item.Value.sTimeoutList.Count - 1 Step +1

                                    If item.Value.sTimeoutList(i).sTimeoutType = enumTimeout.T3 Then

                                        ' 如果是 Data Message，則解除 T3 Timeout 
                                        ReceiveTransactionMap.Item(item.Key).sTimeoutList(i) = Nothing
                                        ReceiveTransactionMap.Item(item.Key).sTimeoutList.RemoveAt(i)

                                    ElseIf item.Value.sTimeoutList(i).sTimeoutType = enumTimeout.T6 Then

                                        ' 如果是 Control Message，則解除 T6 Timeout
                                        ReceiveTransactionMap.Item(item.Key).sTimeoutList(i) = Nothing
                                        ReceiveTransactionMap.Item(item.Key).sTimeoutList.RemoveAt(i)

                                    End If

                                Next

                            End If

                            ReceiveTransactionMap.Item(item.Key).State = enumSecsTransactionState.WaitForDelete

                            Exit Select

                        Case enumSecsTransactionState.WaitForDelete

                            ' 當 Transaction 狀態為 WaitForDelete
                            ' 代表 Transaction 已完成，並等待刪除

                            ' 清除 Primary Message
                            If ReceiveTransactionMap.Item(item.Key).Primary IsNot Nothing Then
                                ReceiveTransactionMap.Item(item.Key).Primary.Dispose()
                            End If

                            ' 清除 Secondary Message
                            If ReceiveTransactionMap.Item(item.Key).Secondary IsNot Nothing Then
                                ReceiveTransactionMap.Item(item.Key).Secondary.Dispose()
                            End If

                            ' 清除所有 Timeout 設定，並移除 Transaction
                            ReceiveTransactionMap.Item(item.Key).sTimeoutList = Nothing
                            ReceiveTransactionMap.TryRemove(item.Key, item.Value)

                            Exit Select

                    End Select
                Next
            End If

            CloneReceiveTransactionMap = Nothing

        End Sub


        ''' <summary>
        ''' 解析 SecsMessage
        ''' </summary>
        ''' <param name="message"></param>
        Private Sub ParseMessage(ByVal message As Object)

            Dim sSecsMessage As SecsMessage = Nothing               ' SecsMessage
            Dim sTransaction As SecsTransaction                     ' Transaction

            Try
                Dim temp() As Byte = CType(message, Byte())

                ' NEW 一個 SECS Message
                sSecsMessage = New SecsMessage(temp)

                ' 檢查 Stream
                If sSXML.sStreamList.Contains(sSecsMessage.Stream.ToString) = False Then

                    Dim tempSecsMessage As SecsMessage = GetMessageByName("S9F3_UnrecognizedStreamType")
                    Dim tempByte As Byte() = ConvertToBytes(sSecsMessage)
                    Dim MessageHeader As Byte() = New Byte(13) {}
                    Array.Copy(tempByte, MessageHeader, 14)
                    tempSecsMessage.secsItem.SetItemValue(MessageHeader)

                    SendPrimary(tempSecsMessage)

                    Exit Sub
                End If

                ' 檢查 Function
                If sSXML.sFunctionList.Contains(sSecsMessage.Stream.ToString & sSecsMessage.Function.ToString) = False Then

                    Dim tempSecsMessage As SecsMessage = GetMessageByName("S9F5_UnrecognizedFunctionType")
                    Dim tempByte As Byte() = ConvertToBytes(sSecsMessage)
                    Dim MessageHeader As Byte() = New Byte(13) {}
                    Array.Copy(tempByte, MessageHeader, 14)
                    tempSecsMessage.secsItem.SetItemValue(MessageHeader)

                    SendPrimary(tempSecsMessage)

                    Exit Sub
                End If

                ' 檢查 SECS Message 是哪種 Message Type
                Select Case sSecsMessage.CheckMessage(sSecsMessage)

                    Case enumMessageType.sSelectRequest

                        ' 收到的 Message 是 Select Request

                        ' 紀錄 Log
                        If siniFile.BinaryLog = True Then
                            sLog.DoBinaryLog(sSecsMessage.ConvertToBinaryLog("RCV"))
                        End If

                        If siniFile.TxLog = True Then
                            sLog.DoTxLog(sSecsMessage.ConvertToSML("PrimaryIn"))
                        End If

                        ' NEW 一個 SecsTransaction
                        sTransaction = New SecsTransaction()
                        sTransaction.Primary = sSecsMessage
                        RaiseEvent OnPrimaryReceived(sTransaction)

                        ' 設定 System Byte
                        Dim SystemByteString As String = Nothing
                        For i As Integer = 0 To sSecsMessage.messageFormat.SystemBytes.Length - 1
                            SystemByteString = SystemByteString & sSecsMessage.messageFormat.SystemBytes(i)
                        Next

                        ' 將 Transaction 加到 Receive Transaction Map
                        ReceiveTransactionMap.TryAdd(SystemByteString, sTransaction)

                        ' NEW 一個 Select Response
                        Dim tempSecsMessage As SecsMessage = New SecsMessage()
                        tempSecsMessage.messageFormat.Length = BitConverter.GetBytes(10)                    ' Length
                        Array.Reverse(tempSecsMessage.messageFormat.Length)
                        tempSecsMessage.messageFormat.DeviceID(0) = Convert.ToByte(&HFF)                    ' DeviceID
                        tempSecsMessage.messageFormat.DeviceID(1) = Convert.ToByte(&HFF)
                        tempSecsMessage.messageFormat.HeaderByte(0) = Convert.ToByte(&H0)                   ' HeaderByte
                        tempSecsMessage.messageFormat.HeaderByte(1) = Convert.ToByte(&H0)
                        tempSecsMessage.messageFormat.PType = Convert.ToByte(&H0)                           ' PType
                        tempSecsMessage.messageFormat.SType = Convert.ToByte(&H2)                           ' SType
                        tempSecsMessage.messageFormat.SystemBytes = sSecsMessage.messageFormat.SystemBytes  ' SystemBytes
                        tempSecsMessage.secsItem = Nothing

                        ' 將 Secondary Message 放到 Transaction
                        sTransaction.Secondary = tempSecsMessage

                        ' 設定 SocketState 為 Selected
                        SocketState = enumSecsConnectionState.sSelected
                        RaiseEvent OnSecsConnectStateChange("Selected")

                        Exit Select


                    Case enumMessageType.sSelectResponse

                        ' 收到的 Message 是 Select Response

                        ' 紀錄 Log
                        If siniFile.BinaryLog = True Then
                            sLog.DoBinaryLog(sSecsMessage.ConvertToBinaryLog("RCV"))
                        End If

                        If siniFile.TxLog = True Then
                            sLog.DoTxLog(sSecsMessage.ConvertToSML("SecondaryIn"))
                        End If


                        ' 從 SendTransactionMap 中，尋找 Transaction 
                        If SendTransactionMap.Count > 0 Then

                            ' 取出 System Byte
                            Dim SystemByteString As String = Nothing
                            For i As Integer = 0 To sSecsMessage.messageFormat.SystemBytes.Length - 1
                                SystemByteString = SystemByteString & sSecsMessage.messageFormat.SystemBytes(i)
                            Next

                            ' 找出 Transaction
                            Dim tempTransaction As SecsTransaction = SendTransactionMap.Item(SystemByteString)

                            If tempTransaction IsNot Nothing Then

                                tempTransaction.Secondary = sSecsMessage
                                tempTransaction.sTimeoutList.Clear()
                                RaiseEvent OnSecondaryReceived(tempTransaction)

                                ' 設定 SocketState 為 Selected
                                SocketState = enumSecsConnectionState.sSelected
                                RaiseEvent OnSecsConnectStateChange("Selected")
                            Else
                                RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                            End If

                        Else
                            RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                        End If

                        Exit Select


                    Case enumMessageType.sLinktestRequest

                        ' 收到的 Message 是 Linktest Request

                        ' 紀錄 Log
                        If siniFile.BinaryLog = True Then
                            sLog.DoBinaryLog(sSecsMessage.ConvertToBinaryLog("RCV"))
                        End If

                        If siniFile.TxLog = True Then
                            sLog.DoTxLog(sSecsMessage.ConvertToSML("PrimaryIn"))
                        End If

                        ' NEW 一個 SecsTransaction
                        sTransaction = New SecsTransaction()
                        sTransaction.Primary = sSecsMessage
                        RaiseEvent OnPrimaryReceived(sTransaction)

                        Dim SystemByteString As String = Nothing

                        For i As Integer = 0 To sSecsMessage.messageFormat.SystemBytes.Length - 1
                            SystemByteString = SystemByteString & sSecsMessage.messageFormat.SystemBytes(i)
                        Next

                        ' 將 Transaction 加到 Receive Transaction Map
                        ReceiveTransactionMap.TryAdd(SystemByteString, sTransaction)

                        ' NEW 一個 Linktest Response
                        Dim tempSecsMessage As SecsMessage = New SecsMessage()
                        tempSecsMessage.messageFormat.Length = BitConverter.GetBytes(10)                    ' Length
                        Array.Reverse(tempSecsMessage.messageFormat.Length)
                        tempSecsMessage.messageFormat.DeviceID(0) = Convert.ToByte(&HFF)                    ' DeviceID
                        tempSecsMessage.messageFormat.DeviceID(1) = Convert.ToByte(&HFF)
                        tempSecsMessage.messageFormat.HeaderByte(0) = Convert.ToByte(&H0)                   ' HeaderByte
                        tempSecsMessage.messageFormat.HeaderByte(1) = Convert.ToByte(&H0)
                        tempSecsMessage.messageFormat.PType = Convert.ToByte(&H0)                           ' PType
                        tempSecsMessage.messageFormat.SType = Convert.ToByte(&H6)                           ' SType
                        tempSecsMessage.messageFormat.SystemBytes = sSecsMessage.messageFormat.SystemBytes  ' SystemBytes
                        tempSecsMessage.secsItem = Nothing
                        sTransaction.Secondary = tempSecsMessage

                        Exit Select

                    ' 收到的 Message 是 Linktest Reqsponse
                    Case enumMessageType.sLinktestResponse

                        ' 紀錄 Log
                        If siniFile.BinaryLog = True Then
                            sLog.DoBinaryLog(sSecsMessage.ConvertToBinaryLog("RCV"))
                        End If

                        If siniFile.TxLog = True Then
                            sLog.DoTxLog(sSecsMessage.ConvertToSML("SecondaryIn"))
                        End If

                        ' 從 SendTransactionMap 中，尋找 Transaction 
                        If SendTransactionMap.Count > 0 Then

                            ' 取出 System Byte
                            Dim SystemByteString As String = Nothing
                            For i As Integer = 0 To sSecsMessage.messageFormat.SystemBytes.Length - 1
                                SystemByteString = SystemByteString & sSecsMessage.messageFormat.SystemBytes(i)
                            Next

                            Dim tempTransaction As SecsTransaction = SendTransactionMap.Item(SystemByteString)

                            If tempTransaction IsNot Nothing Then

                                tempTransaction.Secondary = sSecsMessage
                                tempTransaction.sTimeoutList.Clear()
                                RaiseEvent OnSecondaryReceived(tempTransaction)

                            Else
                                RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                            End If

                        Else
                            RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                        End If

                        Exit Select


                    Case enumMessageType.sDataMessage

                        ' 收到的 Message 是 Data Message

                        If (sSecsMessage.Function Mod 2) = 1 Then

                            ' --------------------------------------- 收到 Primary Message --------------------------------------------


                            ' 紀錄 Log
                            If siniFile.BinaryLog = True Then
                                sLog.DoBinaryLog(sSecsMessage.ConvertToBinaryLog("RCV"))
                            End If

                            If siniFile.TxLog = True Then
                                sLog.DoTxLog(sSecsMessage.ConvertToSML("PrimaryIn"))
                            End If


                            ' 從 SXML 中找出符合 SecsMessage 格式的 Message
                            If (sSXML.FindMessageByFormat(sSecsMessage)) = False Then

                                ' 在 SXML 中沒有找到符合格式的 SecsMessage


                                RaiseEvent OnMessageInfo(sSecsMessage, "Message Not Found")

                                ' 回覆 S9F7 訊息
                                Dim tempSecsMessage As SecsMessage = GetMessageByName("S9F7_IllegalData")
                                Dim tempByte As Byte() = ConvertToBytes(sSecsMessage)
                                Dim MessageHeader As Byte() = New Byte(13) {}
                                Array.Copy(tempByte, MessageHeader, 14)
                                tempSecsMessage.secsItem.SetItemValue(MessageHeader)
                                SendPrimary(tempSecsMessage)

                            Else

                                ' 在 SXML 中找到符合格式的 SecsMessage                          


                                ' NEW 一個 SecsTransaction
                                sTransaction = New SecsTransaction()
                                sTransaction.Primary = sSecsMessage
                                RaiseEvent OnPrimaryReceived(sTransaction)

                                ' 檢查是否為不需要回覆的訊息
                                If sSXML.sDontNeedToReplyList.Contains(sSecsMessage.MessageName) = True Then

                                    ' 透過 System Bytes 找出 Transaction
                                    Dim SystemByteString As String = Nothing

                                    If sSecsMessage.secsItem.ItemValue IsNot Nothing Then

                                        For i As Integer = 11 To sSecsMessage.secsItem.ItemValue.Count

                                            Dim tempByte As Byte() = sSecsMessage.secsItem.ItemValue(i)
                                            SystemByteString = SystemByteString & tempByte(0)

                                        Next

                                    End If

                                    Try
                                        Dim tempTransaction As SecsTransaction = SendTransactionMap.Item(SystemByteString)
                                        tempTransaction.Secondary = sSecsMessage

                                    Catch ex As Exception

                                        ' Do Nothing
                                    End Try


                                Else

                                    Dim SystemByteString As String = Nothing

                                    For i As Integer = 0 To sSecsMessage.messageFormat.SystemBytes.Length - 1
                                        SystemByteString = SystemByteString & sSecsMessage.messageFormat.SystemBytes(i)
                                    Next

                                    ReceiveTransactionMap.TryAdd(SystemByteString, sTransaction)

                                    ' 是否執行 AutoReply
                                    DoAutoReply(sTransaction)
                                End If

                            End If

                            Exit Select

                        Else

                            ' --------------------------------------- 收到 Secondary Message -----------------------------------------------


                            ' 紀錄 Log
                            If siniFile.BinaryLog = True Then
                                sLog.DoBinaryLog(sSecsMessage.ConvertToBinaryLog("RCV"))
                            End If

                            If siniFile.TxLog = True Then
                                sLog.DoTxLog(sSecsMessage.ConvertToSML("SecondaryIn"))
                            End If


                            ' 檢查 Device ID
                            If siniFile.Role = enumRole.sHost Then

                                ' Device ID 不同時
                                If sSecsMessage.messageFormat.DeviceID(0) <> siniFile.DeviceID Then

                                    ' 透過 System Bytes 找出 Transaction
                                    Dim SystemByteString As String = Nothing

                                    For i As Integer = 0 To sSecsMessage.messageFormat.SystemBytes.Length - 1
                                        SystemByteString = SystemByteString & sSecsMessage.messageFormat.SystemBytes(i)
                                    Next

                                    Dim tempTransaction As SecsTransaction = SendTransactionMap.Item(SystemByteString)

                                    If tempTransaction IsNot Nothing Then
                                        tempTransaction.Secondary = sSecsMessage
                                    Else
                                        RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                                    End If

                                    Dim tempSecsMessage As SecsMessage = GetMessageByName("S9F1_UnrecognizedDeviceID")
                                    Dim tempByte As Byte() = ConvertToBytes(sSecsMessage)
                                    Dim MessageHeader As Byte() = New Byte(13) {}
                                    Array.Copy(tempByte, MessageHeader, 14)
                                    tempSecsMessage.secsItem.SetItemValue(MessageHeader)

                                    SendPrimary(tempSecsMessage)

                                    Exit Sub

                                End If

                            ElseIf siniFile.Role = enumRole.sEquipment Then

                                ' Device ID 不同時
                                If sSecsMessage.messageFormat.DeviceID(1) <> siniFile.DeviceID Then

                                    ' 透過 System Bytes 找出 Transaction
                                    Dim SystemByteString As String = Nothing

                                    For i As Integer = 0 To sSecsMessage.messageFormat.SystemBytes.Length - 1
                                        SystemByteString = SystemByteString & sSecsMessage.messageFormat.SystemBytes(i)
                                    Next

                                    Dim tempTransaction As SecsTransaction = SendTransactionMap.Item(SystemByteString)

                                    If tempTransaction IsNot Nothing Then
                                        tempTransaction.Secondary = sSecsMessage
                                    Else
                                        RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                                    End If

                                    Dim tempSecsMessage As SecsMessage = GetMessageByName("S9F1_UnrecognizedDeviceID")
                                    Dim tempByte As Byte() = ConvertToBytes(sSecsMessage)
                                    Dim MessageHeader As Byte() = New Byte(13) {}
                                    Array.Copy(tempByte, MessageHeader, 14)
                                    tempSecsMessage.secsItem.SetItemValue(MessageHeader)

                                    SendPrimary(tempSecsMessage)

                                    Exit Sub

                                End If

                            End If

                            ' 從 SendTransactionMap 中，尋找 Transaction
                            If SendTransactionMap.Count > 0 Then

                                Dim SystemByteString As String = Nothing

                                For i As Integer = 0 To sSecsMessage.messageFormat.SystemBytes.Length - 1
                                    SystemByteString = SystemByteString & sSecsMessage.messageFormat.SystemBytes(i)
                                Next

                                Dim tempTransaction As SecsTransaction = SendTransactionMap.Item(SystemByteString)

                                If tempTransaction IsNot Nothing Then

                                    ' 從 SXML 中找出符合 SecsMessage 格式的 Message                              
                                    If (sSXML.FindMessageByFormat(sSecsMessage)) = False Then

                                        ' 在 SXML 中沒有找到符合格式的 SecsMessage


                                        RaiseEvent OnMessageInfo(sSecsMessage, "Message Not Found")

                                        ' 回覆 S9F7 訊息
                                        Dim tempSecsMessage As SecsMessage = GetMessageByName("S9F7_IllegalData")
                                        Dim tempByte As Byte() = ConvertToBytes(sSecsMessage)
                                        Dim MessageHeader As Byte() = New Byte(13) {}
                                        Array.Copy(tempByte, MessageHeader, 14)
                                        tempSecsMessage.secsItem.SetItemValue(MessageHeader)

                                        SendPrimary(tempSecsMessage)

                                    Else
                                        ' 在 SXML 中找到符合格式的 SecsMessage  

                                        tempTransaction.Secondary = sSecsMessage
                                        RaiseEvent OnSecondaryReceived(tempTransaction)

                                    End If

                                Else
                                    RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                                End If

                            Else
                                RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                            End If

                            Exit Select

                        End If

                End Select

            Catch ex As Exception

                RaiseEvent OnMessageError(sSecsMessage, "Message Error")
            End Try

        End Sub


        ''' <summary>
        ''' 檢察是否有 Timeout
        ''' </summary>
        Private Sub CheckTimeout()

            Do
                SyncLock sTimeoutList

                    ' 檢察系統 Timeout
                    If sTimeoutList.Count > 0 Then

                        For i As Integer = 0 To sTimeoutList.Count - 1 Step +1

                            Select Case sTimeoutList(i).sTimeoutType

                                Case enumTimeout.T5

                                    If sTimeoutList(i).CheckTimeout() = True Then

                                        RaiseEvent OnTimeout("T5 Timeout")

                                        ' 如果是 Active，則由 T5 Timeout 來重新連線
                                        sTimeoutList(i) = Nothing
                                        sTimeoutList.RemoveAt(i)
                                        sClient.Reconnect()

                                        Exit For

                                    End If

                                Case enumTimeout.T7

                                    If sTimeoutList(i).CheckTimeout() = True Then

                                        RaiseEvent OnTimeout("T7 Timeout")

                                        sTimeoutList(i) = Nothing
                                        sTimeoutList.RemoveAt(i)
                                        Exit For

                                    End If

                                Case enumTimeout.T8

                                    If sTimeoutList(i).CheckTimeout() = True Then

                                        RaiseEvent OnTimeout("T8 Timeout")

                                        sTimeoutList(i) = Nothing
                                        sTimeoutList.RemoveAt(i)

                                        If Me.siniFile.Entity = enumSecsEntity.sActive Then

                                            sClient.disconnect()
                                        Else
                                            sServer.disconnect()
                                        End If

                                        Exit For

                                    End If

                            End Select
                        Next

                    End If

                End SyncLock

                Application.DoEvents()
                Thread.Sleep(1000)

            Loop While Not SocketState = enumSecsConnectionState.sSeparate

        End Sub


        ''' <summary>
        ''' 檢查 IniFile
        ''' </summary>
        Private Sub CheckIniFile()
            Do
                SyncLock siniFile

                    ' 重新讀取 IniFile
                    siniFile.ReLoadIniFile()

                End SyncLock

                Application.DoEvents()
                Thread.Sleep(1000)

            Loop While Not SocketState = enumSecsConnectionState.sSeparate
        End Sub


        ''' <summary>
        ''' 產生 System Bytes
        ''' </summary>
        ''' <returns></returns>
        Private Function getSystemByte() As UInteger

            systemBytesCount = systemBytesCount + 1

            If systemBytesCount >= 4294967295 Then
                systemBytesCount = 0
            End If

            Return systemBytesCount
        End Function


        ''' <summary>
        ''' 將 Message 透過 Socket 送出
        ''' </summary>
        ''' <param name="aSecsMessage"></param>
        Private Sub SocketSend(ByRef aSecsMessage As SecsMessage)

            ' 如果是 Active，則使用 Client
            ' 如果是 Passive，則使用 Server
            If Me.siniFile.Entity = enumSecsEntity.sActive Then
                sClient.send(ConvertToBytes(aSecsMessage))
            Else
                sServer.send(ConvertToBytes(aSecsMessage))
            End If

        End Sub


        ''' <summary>
        ''' 是否執行 AutoReply
        ''' </summary>
        ''' <param name="aTransaction"></param>
        Private Sub DoAutoReply(ByRef aTransaction As SecsTransaction)

            Try
                If siniFile.AutoReply = True Then

                    If GetMessageByName(aTransaction.Primary.AutoReply) IsNot Nothing Then

                        Dim tempTransaction As SecsTransaction = New SecsTransaction
                        tempTransaction = aTransaction
                        tempTransaction.Secondary = GetMessageByName(aTransaction.Primary.AutoReply)
                        SendSecondary(tempTransaction)

                    End If

                End If
            Catch ex As Exception

                RaiseEvent OnMessageError(aTransaction.Primary, "AutoReply Fail")
            End Try

        End Sub


        ''' <summary>
        ''' 將 SecsMessage 轉換成 Bytes
        ''' </summary>
        ''' <param name="aSecsMessage"></param>
        ''' <returns></returns>
        Private Function ConvertToBytes(ByRef aSecsMessage As SecsMessage) As Byte()

            ' 將 SecsItem 轉換成 Bytes
            Dim ItemTempList As ArrayList = New ArrayList

            If aSecsMessage.secsItem IsNot Nothing Then

                ' 將 SecsItem 轉換成 Bytes
                ItemConvertToBytes(aSecsMessage.secsItem, ItemTempList)
            End If

            aSecsMessage.messageFormat.Length = BitConverter.GetBytes(ItemTempList.Count + 10)
            Array.Reverse(aSecsMessage.messageFormat.Length)

            ' 將 SecsMessage 轉換成 Bytes
            Dim MessageTempList As ArrayList = New ArrayList

            ' Length
            MessageTempList.Add(aSecsMessage.messageFormat.Length(0))
            MessageTempList.Add(aSecsMessage.messageFormat.Length(1))
            MessageTempList.Add(aSecsMessage.messageFormat.Length(2))
            MessageTempList.Add(aSecsMessage.messageFormat.Length(3))

            ' DeviceID
            If aSecsMessage.messageFormat.RBit = True Then
                MessageTempList.Add((aSecsMessage.messageFormat.DeviceID(0) Or &H80))
            Else
                MessageTempList.Add(aSecsMessage.messageFormat.DeviceID(0))
            End If
            MessageTempList.Add(aSecsMessage.messageFormat.DeviceID(1))

            ' Header Bytes
            If aSecsMessage.messageFormat.WBit = True Then
                MessageTempList.Add((aSecsMessage.messageFormat.HeaderByte(0) Or &H80))
            Else
                MessageTempList.Add(aSecsMessage.messageFormat.HeaderByte(0))
            End If
            MessageTempList.Add(aSecsMessage.messageFormat.HeaderByte(1))

            ' PType
            MessageTempList.Add(aSecsMessage.messageFormat.PType)

            ' SType
            MessageTempList.Add(aSecsMessage.messageFormat.SType)

            ' System Bytes
            MessageTempList.Add(aSecsMessage.messageFormat.SystemBytes(0))
            MessageTempList.Add(aSecsMessage.messageFormat.SystemBytes(1))
            MessageTempList.Add(aSecsMessage.messageFormat.SystemBytes(2))
            MessageTempList.Add(aSecsMessage.messageFormat.SystemBytes(3))

            ' 假如有 SecsItem
            If ItemTempList IsNot Nothing Then

                ' 將 SecsItem 的 Bytes加進要送出的 Bytes
                For i As Int64 = 0 To ItemTempList.Count - 1 Step +1

                    MessageTempList.Add(ItemTempList(i))
                Next

            End If

            ' 產生要送出的 SecsMessage 的 Bytes
            Dim tempBytes(MessageTempList.Count - 1) As Byte

            For i As Int64 = 0 To MessageTempList.Count - 1 Step +1

                tempBytes(i) = MessageTempList(i)
            Next

            Return tempBytes

        End Function


        ''' <summary>
        ''' 將 SecsItem 轉換成 Bytes
        ''' </summary>
        ''' <param name="item"></param>
        ''' <param name="temp"></param>
        Private Sub ItemConvertToBytes(ByRef item As SecsItem, ByRef temp As ArrayList)

            If item.ItemType = "L" Then

                ' -------------- 假如 Item = List -----------------------

                Dim tempByte() As Byte
                tempByte = item.ItemConvertToBytes()

                For i As Int64 = 0 To tempByte.Length - 1
                    temp.Add(tempByte(i))
                Next

                For i As Integer = 0 To item.ItemList.Count - 1 Step +1
                    ItemConvertToBytes(item.ItemList(i), temp)
                Next

            Else

                ' -------------- 假如 Item != List -----------------------

                Dim tempByte() As Byte
                tempByte = item.ItemConvertToBytes()

                For i As Int64 = 0 To tempByte.Length - 1
                    temp.Add(tempByte(i))
                Next
            End If

        End Sub

#End Region


#Region "實作 IFMessageListener 的 Method"

        ''' <summary>
        ''' SECS 的 TCP / IP 連線狀況
        ''' </summary>
        ''' <param name="sysMessage"></param>
        Private Sub sysMessage(ByVal sysMessage As String) Implements IFMessageListener.sysMessage

            If sysMessage = "NotConnected" Then

                ' 收到 Client / Server 的 Not Connected 訊息

                ' 如果是 Active，則設定 T5 Timeout，由 T5 Timeout 來重新連線 
                SyncLock sTimeoutList

                    If sTimeoutList.Count > 0 Then

                        Dim temp As Integer = 0
                        For i As Integer = 0 To sTimeoutList.Count - 1
                            If sTimeoutList(i).sTimeoutType = enumTimeout.T5 Then
                                temp = temp + 1
                            End If
                        Next

                        If temp = 0 Then
                            If Me.siniFile.Entity = enumSecsEntity.sActive Then
                                sTimeout = New Timeout(siniFile.T5Timeout, enumTimeout.T5)
                                sTimeoutList.Add(sTimeout)
                            End If
                        End If
                    Else
                        If Me.siniFile.Entity = enumSecsEntity.sActive Then
                            sTimeout = New Timeout(siniFile.T5Timeout, enumTimeout.T5)
                            sTimeoutList.Add(sTimeout)
                        End If
                    End If

                End SyncLock

                SocketState = enumSecsConnectionState.sNotConnected
                RaiseEvent OnSecsConnectStateChange("NotConnected")

                Application.DoEvents()
                Thread.Sleep(1)

            ElseIf sysMessage = "Connected" Then

                ' 收到 Client / Server 的 Not Selected 訊息

                SocketState = enumSecsConnectionState.sNotSelected
                RaiseEvent OnSecsConnectStateChange("NotSelected")

                Application.DoEvents()
                Thread.Sleep(1)

            ElseIf sysMessage = "Separate" Then

                SocketState = enumSecsConnectionState.sSeparate
                RaiseEvent OnSecsConnectStateChange("Separate")

                Application.DoEvents()
                Thread.Sleep(1)

            ElseIf sysMessage = "Set T8 Timeout" Then

                sTimeout = New Timeout(siniFile.T8Timeout, enumTimeout.T8)
                sTimeoutList.Add(sTimeout)

                Application.DoEvents()
                Thread.Sleep(1)

            End If

		End Sub


        ' 收到 SECS Message 
        Private Sub onMessage(ByRef message As Byte()) Implements IFMessageListener.onMessage

            ' 解除 T8 Timeout 的設定
            SyncLock sTimeoutList

                For i As Integer = 0 To sTimeoutList.Count - 1 Step +1

                    Select Case sTimeoutList(i).sTimeoutType

                        Case enumTimeout.T8

                            sTimeoutList(i) = Nothing
                            sTimeoutList.RemoveAt(i)

                    End Select

                Next

            End SyncLock

            ReceiveMessageQueue.Enqueue(message)

        End Sub

#End Region



        ' ----------------例外---------------------------------------

        Public Sub SendS9F1Message(ByRef aSecsMessage As SecsMessage)

            Dim sTransaction As SecsTransaction

            Try
                ' 檢查 System Bytes
                aSecsMessage.messageFormat.SystemBytes = BitConverter.GetBytes(getSystemByte())
                Array.Reverse(aSecsMessage.messageFormat.SystemBytes)


                ' 檢查 Messae 是否為不需回覆的 Message
                If sSXML.sDontNeedToReplyList.Contains(aSecsMessage.MessageName) = True Then

                    ' 紀錄 Log
                    If siniFile.BinaryLog = True Then
                        sLog.DoBinaryLog(aSecsMessage.ConvertToBinaryLog("SND"))
                    End If

                    If siniFile.TxLog = True Then
                        sLog.DoTxLog(aSecsMessage.ConvertToSML("PrimaryOut"))
                    End If

                    ' NEW 一個 SecsTransaction
                    sTransaction = New SecsTransaction()
                    sTransaction.Primary = aSecsMessage
                    RaiseEvent OnPrimarySent(sTransaction)

                    SocketSend(aSecsMessage)

                Else
                    ' New 一個 SecsTransaction，並把 Primary 放在裡面
                    sTransaction = New SecsTransaction()
                    sTransaction.Primary = aSecsMessage

                    ' SendTransactionMap 新增此 Transaction
                    Dim SystemByteString As String = Nothing

                    For i As Integer = 0 To aSecsMessage.messageFormat.SystemBytes.Length - 1
                        SystemByteString = SystemByteString & aSecsMessage.messageFormat.SystemBytes(i)
                    Next

                    SendTransactionMap.TryAdd(SystemByteString, sTransaction)

                End If

            Catch ex As Exception

                RaiseEvent OnMessageError(aSecsMessage, "Send Primary Error")

                ' 刪除 Primary
                If aSecsMessage IsNot Nothing Then
                    aSecsMessage.Dispose()
                End If

                ' 刪除 Transaction
                sTransaction = Nothing

            End Try

        End Sub


        Public Sub SendT8TimeoutMessage(ByRef aSecsMessage As SecsMessage)

            Dim sTransaction As SecsTransaction

            Try
                ' 檢查 System Bytes
                aSecsMessage.messageFormat.SystemBytes = BitConverter.GetBytes(getSystemByte())
                Array.Reverse(aSecsMessage.messageFormat.SystemBytes)

                ' 檢查 Device ID
                If siniFile.Role = enumRole.sHost And aSecsMessage.MessageType = enumMessageType.sDataMessage Then

                    aSecsMessage.messageFormat.DeviceID(0) = Convert.ToByte(siniFile.DeviceID)

                ElseIf siniFile.Role = enumRole.sEquipment And aSecsMessage.MessageType = enumMessageType.sDataMessage Then

                    aSecsMessage.messageFormat.DeviceID(1) = Convert.ToByte(siniFile.DeviceID)

                End If

                ' 檢查 Messae 是否為不需回覆的 Message
                If sSXML.sDontNeedToReplyList.Contains(aSecsMessage.MessageName) = True Then

                    ' 紀錄 Log
                    If siniFile.BinaryLog = True Then
                        sLog.DoBinaryLog(aSecsMessage.ConvertToBinaryLog("SND"))
                    End If

                    If siniFile.TxLog = True Then
                        sLog.DoTxLog(aSecsMessage.ConvertToSML("PrimaryOut"))
                    End If

                    ' NEW 一個 SecsTransaction
                    sTransaction = New SecsTransaction()
                    sTransaction.Primary = aSecsMessage
                    RaiseEvent OnPrimarySent(sTransaction)

                    SocketSend(aSecsMessage)

                Else
                    ' New 一個 SecsTransaction，並把 Primary 放在裡面
                    sTransaction = New SecsTransaction()
                    sTransaction.Primary = aSecsMessage

                    Dim tempByte As Byte() = ConvertToBytes(aSecsMessage)
                    Dim tempByteALength As Int16 = tempByte.Length / 2
                    Dim tempByteA As Byte() = New Byte(tempByteALength - 1) {}
                    Dim tempByteB As Byte() = New Byte(tempByteALength - 1) {}
                    Array.Copy(tempByte, 0, tempByteA, 0, tempByteALength)
                    Array.Copy(tempByte, tempByteALength, tempByteB, 0, tempByteALength)

                    If Me.siniFile.Entity = enumSecsEntity.sActive Then

                        sClient.send(tempByteA)

                    Else
                        sServer.send(tempByteA)

                    End If

                    RaiseEvent OnPrimarySent(sTransaction)
                    sTransaction.State = enumSecsTransactionState.PrimarySent

                    ' SendTransactionMap 新增此 Transaction
                    Dim SystemByteString As String = Nothing

                    For i As Integer = 0 To aSecsMessage.messageFormat.SystemBytes.Length - 1
                        SystemByteString = SystemByteString & aSecsMessage.messageFormat.SystemBytes(i)
                    Next

                    SendTransactionMap.TryAdd(SystemByteString, sTransaction)

                End If

            Catch ex As Exception

                RaiseEvent OnMessageError(aSecsMessage, "Send Primary Error")

                ' 刪除 Primary
                If aSecsMessage IsNot Nothing Then
                    aSecsMessage.Dispose()
                End If

                ' 刪除 Transaction
                sTransaction = Nothing

            End Try

        End Sub


        Private Sub TimerTick(ByVal state As Object)

        End Sub


    End Class

End Namespace

