Imports System
Imports System.Text
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports SecsDriver_VB.SocketDriver
Imports System.Threading

Namespace SecsDriver

    ' ----------------------------- Delegate ----------------------------------

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



    Public Class SecsWrapper : Implements IFMessageListener

		' Send Transaction Map
		Private SendTransactionMap As Dictionary(Of Byte(), SecsTransaction)

		' Receive Transaction Map
		Private ReceiveTransactionMap As Dictionary(Of Byte(), SecsTransaction)

        ' System Bytes
        Private systemBytesCount As UInteger = Now.Millisecond


        ' ---------------- 物件 -------------------------------

        Public siniFile As IniFile                                  ' Ini File
        Private sSXML As SXML                                       ' SXML
        Private sServer As Server                                   ' Socket Server
        Private sClient As Client                                   ' Socket Client
        Private sSecsMessage As SecsMessage                         ' SecsMessage
        Private sTransaction As SecsTransaction                     ' Transaction
        Private sTimeout As Timeout                                 ' Timeout
        Private sTimeoutList As List(Of Timeout)                    ' TimeoutList
        Private sBinaryLog As FileStream                            ' Binary Log
        Private sTxLog As FileStream                                ' Tx Log


        ' SECS 連線狀態
        Private SocketState As enumSecsConnectionState

        ' Lock 物件
        Private Shared thisLock As Object = New Object

        ' ----------------------------- Event --------------------------------- 
        Public Event OnSecsConnectStateChange As delegateSecsConnectStateChange
        Public Event OnPrimarySent As delegatePrimarySent
		Public Event OnPrimaryReceived As delegatePrimaryReceived
		Public Event OnSecondarySent As delegateSecondarySent
		Public Event OnSecondaryReceived As delegateSecondaryReceived
        Public Event OnMessageError As delegateMessageError
        Public Event OnMessageInfo As delegateMessageInfo
        Public Event OnError As delegateError
        Public Event OnTimeout As delegateTimeout


        ' ----------------- 可供外部使用的程式 ---------------------------

        ' 建構子
        Public Sub New()

            SendTransactionMap = New Dictionary(Of Byte(), SecsTransaction)
            ReceiveTransactionMap = New Dictionary(Of Byte(), SecsTransaction)
            sTimeoutList = New List(Of Timeout)

        End Sub


        ' 開啟連線
        Public Sub Open()

            ' IniFile Path
            Dim path As String = Application.StartupPath

            Try
                ' 載入 iniFile
                siniFile = New IniFile(path & "\config.ini")

                ' 建立 Log Txt 檔
                sBinaryLog = New FileStream(Application.StartupPath & "\SECSLog\BinaryLog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
                sBinaryLog.Close()
                sTxLog = New FileStream(Application.StartupPath & "\SECSLog\TxLog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
                sTxLog.Close()

            Catch ex As Exception

                RaiseEvent OnError("Load IniFile Error !")

            End Try

            Try
                ' 載入 SXML
                sSXML = New SXML()
                sSXML.FilePath = siniFile.SXMLPath
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


        ' 關閉連線
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


		' 使用 Message Name 找 Message
		Public Function GetMessageByName(ByVal aMessageName As String) As SecsMessage

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


		' Send Primary Message
		Public Sub SendPrimary(ByRef aSecsMessage As SecsMessage)

            Try
                SyncLock thisLock

                    ' 檢查 System Bytes
                    aSecsMessage.messageFormat.SystemBytes = BitConverter.GetBytes(getSystemByte())
                    Array.Reverse(aSecsMessage.messageFormat.SystemBytes)

                    ' New 一個 SecsTransaction，並把 Primary 放在裡面
                    Dim sTransaction As SecsTransaction = New SecsTransaction()
                    sTransaction.Primary = aSecsMessage

                    ' SendTransactionMap 新增此 Transaction
                    SyncLock SendTransactionMap
                        SendTransactionMap.Add(aSecsMessage.messageFormat.SystemBytes, sTransaction)
                    End SyncLock

                    ' 如果是 Data Message，則設定 T3 Timeout 
                    ' 如果是 Control Message，則設定 T6 Timeout 
                    If aSecsMessage.MessageType = enumMessageType.sDataMessage Then
                        sTransaction.SetTimeout(enumTimeout.T3, siniFile)
                    Else
                        If aSecsMessage.MessageType <> enumMessageType.sSeparateRequest Then
                            sTransaction.SetTimeout(enumTimeout.T6, siniFile)
                        End If
                    End If

                    ' 如果是 Active，則使用 Client
                    ' 如果是 Passive，則使用 Server
                    If Me.siniFile.Entity = enumSecsEntity.sActive Then
                        sClient.send(aSecsMessage.ConvertToBytes())
                    Else
                        sServer.send(aSecsMessage.ConvertToBytes())
                    End If

                    sTransaction.State = enumSecsTransactionState.PrimarySent
                    RaiseEvent OnPrimarySent(sTransaction)

                    ' 紀錄 Log
                    DoLog(sTransaction.Primary.ConvertToBinaryLog("SND"), sTransaction.Primary.ConvertToSML("PrimaryOut"))

                End SyncLock

            Catch ex As Exception

                RaiseEvent OnMessageError(sTransaction.Primary, "Send Primary Error")

                ' 刪除 Primary
                If sTransaction.Primary IsNot Nothing Then
                    sTransaction.Primary.Dispose()
                End If

                ' 刪除 Transaction
                sTransaction = Nothing

            End Try

        End Sub


        ' Send Secondary Message
        Public Sub SendSecondary(ByRef aTransaction As SecsTransaction)

            Try
                SyncLock thisLock

                    ' 如果是 Data Message，則解除 T3 Timeout 
                    ' 如果是 Control Message，則解除 T6 Timeout
                    If aTransaction.sTimeoutList.Count > 0 Then

                        For i As Integer = 0 To aTransaction.sTimeoutList.Count - 1 Step +1
                            If aTransaction.sTimeoutList(i).sTimeoutType = enumTimeout.T3 Then

                                aTransaction.sTimeoutList(i) = Nothing
                                aTransaction.sTimeoutList.RemoveAt(i)
                                Exit For

                            ElseIf aTransaction.sTimeoutList(i).sTimeoutType = enumTimeout.T6 Then

                                aTransaction.sTimeoutList(i) = Nothing
                                aTransaction.sTimeoutList.RemoveAt(i)
                                Exit For

                            End If
                        Next

                    End If

                    ' 如果是 Active，則使用 Client
                    ' 如果是 Passive，則使用 Server
                    If Me.siniFile.Entity = enumSecsEntity.sActive Then

                        sClient.send(aTransaction.Secondary.ConvertToBytes())
                    Else

                        sServer.send(aTransaction.Secondary.ConvertToBytes())
                    End If

                    aTransaction.State = enumSecsTransactionState.SecondarySent
                    aTransaction.State = enumSecsTransactionState.WaitForDelete
                    RaiseEvent OnSecondarySent(aTransaction)

                    ' 紀錄 Log
                    DoLog(aTransaction.Secondary.ConvertToBinaryLog("SND"), aTransaction.Secondary.ConvertToSML("SecondaryOut"))

                End SyncLock

            Catch ex As Exception

                RaiseEvent OnMessageError(aTransaction.Secondary, "Send Secondary Error")

                ' 刪除 Primary
                If aTransaction.Primary IsNot Nothing Then
                    aTransaction.Primary.Dispose()
                End If

                ' 刪除 Secondary
                If aTransaction.Secondary IsNot Nothing Then
                    aTransaction.Secondary.Dispose()
                End If

                ' 刪除 Transaction
                aTransaction = Nothing

            End Try

        End Sub



        ' ------------------內部使用的程式------------------------------------

        ' 檢查連線狀況、Timeout狀況
        Private Sub WatchSecsConnectionStatus()

			' NEW 一個 Thread 來持續監看 SECS 連線狀態
			Dim sWatchSecsConnectThread As New Thread(New ThreadStart(AddressOf SecsConnect))
			sWatchSecsConnectThread.IsBackground = True
			sWatchSecsConnectThread.Start()

            ' New 一個 Thread 來持續檢查是否有 Timeout
            Dim sWatchTimeoutThread As New Thread(New ThreadStart(AddressOf CheckTimeout))
            sWatchTimeoutThread.IsBackground = True
            sWatchTimeoutThread.Start()

            ' New 一個 Thread 來持續檢查 IniFile 是否有變動
            Dim sWatchIniFileThread As New Thread(New ThreadStart(AddressOf CheckIniFile))
            sWatchIniFileThread.IsBackground = True
            sWatchIniFileThread.Start()

            ' New 一個 Thread 來持續檢查 Log.txt是否過大
            Dim sWatchLogTxtThread As New Thread(New ThreadStart(AddressOf CheckLogTxt))
            sWatchLogTxtThread.IsBackground = True
            sWatchLogTxtThread.Start()

        End Sub


        ' 檢查連線狀況
        Private Sub SecsConnect()

            Dim sendSelectRequestCountdown As Integer = 0
            Dim sendLinktestRequestCountdown As Integer = 0

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


        ' 檢察是否有 Timeout
        Private Sub CheckTimeout()

            Do
                SyncLock thisLock

                    CheckSendTransactionMap()

                    CheckReceiveTransactionMap()

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

                                End Select
                            Next

                        End If

                    End SyncLock

                End SyncLock

                Application.DoEvents()
                Thread.Sleep(1)

            Loop While Not SocketState = enumSecsConnectionState.sSeparate
        End Sub


        ' 檢查 SendTransactionMap
        Private Sub CheckSendTransactionMap()

            SyncLock SendTransactionMap

                ' 檢查 SendTransactionMap
                If SendTransactionMap.Count > 0 Then

                    For Each item In SendTransactionMap

                        Select Case item.Value.State

                            Case enumSecsTransactionState.PrimarySent

                                ' 檢查各個 Timeout
                                If item.Value.sTimeoutList.Count > 0 Then

                                    For i As Integer = 0 To item.Value.sTimeoutList.Count - 1 Step +1

                                        Select Case item.Value.sTimeoutList(i).sTimeoutType

                                            Case enumTimeout.T3

                                                If item.Value.sTimeoutList(i).CheckTimeout = True Then

                                                    If item.Value.Primary.MessageName IsNot Nothing Then
                                                        RaiseEvent OnTimeout("Sent : " & item.Value.Primary.MessageName & " T3 Timeout")
                                                    Else
                                                        RaiseEvent OnTimeout("Sent : Control Message T3 Timeout")
                                                    End If

                                                    If item.Value.Primary IsNot Nothing Then
                                                        item.Value.Primary.Dispose()
                                                    End If
                                                    item.Value.sTimeoutList(i) = Nothing
                                                    item.Value.sTimeoutList = Nothing
                                                    SendTransactionMap.Remove(item.Key)

                                                    Exit Sub
                                                End If

                                            Case enumTimeout.T6

                                                If item.Value.sTimeoutList(i).CheckTimeout = True Then

                                                    If item.Value.Primary.MessageName IsNot Nothing Then
                                                        RaiseEvent OnTimeout("Sent : " & item.Value.Primary.MessageName & " T6 Timeout")
                                                    Else
                                                        RaiseEvent OnTimeout("Sent : Control Message T6 Timeout")
                                                    End If

                                                    If item.Value.Primary IsNot Nothing Then
                                                        item.Value.Primary.Dispose()
                                                    End If
                                                    item.Value.sTimeoutList(i) = Nothing
                                                    item.Value.sTimeoutList = Nothing
                                                    SendTransactionMap.Remove(item.Key)

                                                    Exit Sub
                                                End If

                                        End Select
                                    Next
                                End If

                            Case enumSecsTransactionState.WaitForDelete

                                If item.Value.Primary IsNot Nothing Then
                                    item.Value.Primary.Dispose()
                                End If
                                If item.Value.Secondary IsNot Nothing Then
                                    item.Value.Secondary.Dispose()
                                End If

                                item.Value.sTimeoutList = Nothing
                                SendTransactionMap.Remove(item.Key)

                                Exit Sub

                        End Select
                    Next
                End If
            End SyncLock

        End Sub


        ' 檢查 ReceiveTransactionMap
        Private Sub CheckReceiveTransactionMap()

            SyncLock ReceiveTransactionMap

                ' 檢查 ReceiveTransactionMap
                If ReceiveTransactionMap.Count > 0 Then

                    For Each item In ReceiveTransactionMap

                        Select Case item.Value.State

                            Case enumSecsTransactionState.PrimaryReceived

                                ' 檢查各個 Timeout
                                If item.Value.sTimeoutList.Count > 0 Then

                                    For i As Integer = 0 To item.Value.sTimeoutList.Count - 1 Step +1

                                        Select Case item.Value.sTimeoutList(i).sTimeoutType

                                            Case enumTimeout.T3

                                                If item.Value.sTimeoutList(i).CheckTimeout = True Then

                                                    If item.Value.Primary.MessageName IsNot Nothing Then
                                                        RaiseEvent OnTimeout("Received : " & item.Value.Primary.MessageName & " T3 Timeout")
                                                    Else
                                                        RaiseEvent OnTimeout("Received : Control Message T3 Timeout")
                                                    End If

                                                    If item.Value.Primary IsNot Nothing Then
                                                        item.Value.Primary.Dispose()
                                                    End If
                                                    item.Value.sTimeoutList(i) = Nothing
                                                    item.Value.sTimeoutList = Nothing
                                                    ReceiveTransactionMap.Remove(item.Key)

                                                    Exit Sub

                                                End If

                                            Case enumTimeout.T6

                                                If item.Value.sTimeoutList(i).CheckTimeout = True Then

                                                    If item.Value.Primary.MessageName IsNot Nothing Then
                                                        RaiseEvent OnTimeout("Received : " & item.Value.Primary.MessageName & " T6 Timeout")
                                                    Else
                                                        RaiseEvent OnTimeout("Received : Control Message T6 Timeout")
                                                    End If

                                                    If item.Value.Primary IsNot Nothing Then
                                                        item.Value.Primary.Dispose()
                                                    End If
                                                    item.Value.sTimeoutList(i) = Nothing
                                                    item.Value.sTimeoutList = Nothing
                                                    ReceiveTransactionMap.Remove(item.Key)

                                                    Exit Sub

                                                End If

                                        End Select
                                    Next
                                End If

                            Case enumSecsTransactionState.WaitForDelete

                                If item.Value.Primary IsNot Nothing Then
                                    item.Value.Primary.Dispose()
                                End If
                                If item.Value.Secondary IsNot Nothing Then
                                    item.Value.Secondary.Dispose()
                                End If

                                item.Value.sTimeoutList = Nothing
                                ReceiveTransactionMap.Remove(item.Key)

                                Exit Sub

                        End Select
                    Next
                End If
            End SyncLock

        End Sub


        ' 檢查 IniFile
        Private Sub CheckIniFile()
            Do
                SyncLock siniFile

                    siniFile.ReLoadIniFile()

                End SyncLock

                Application.DoEvents()
                Thread.Sleep(1000)

            Loop While Not SocketState = enumSecsConnectionState.sSeparate
        End Sub


        ' 檢查 Log.txt 檔的大小
        Private Sub CheckLogTxt()
            Do
                Dim tempDir As New DirectoryInfo(Application.StartupPath & "\SECSLog")
                Dim tempBinaryLogs As FileInfo() = tempDir.GetFiles("*BinaryLog.txt*")
                Dim tempTxLogs As FileInfo() = tempDir.GetFiles("TxLog.txt")
                Dim tempBinaryLog As FileInfo = tempBinaryLogs(0)
                Dim tempTxLog As FileInfo = tempTxLogs(0)

                If tempBinaryLog IsNot Nothing Then

                    If tempBinaryLog.Length > 500000000 Then

                        FileSystem.RenameFile(Application.StartupPath & "\SECSLog\BinaryLog.txt", "BinaryLog_" & Now.ToString("yyyyMMdd-hh.mm.ss") & ".txt")

                        ' 建立 Log Txt 檔
                        sBinaryLog = New FileStream(Application.StartupPath & "\SECSLog\BinaryLog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
                        sBinaryLog.Close()

                    End If
                End If

                If tempTxLog IsNot Nothing Then

                    If tempTxLog.Length > 500000000 Then

                        FileSystem.RenameFile(Application.StartupPath & "\SECSLog\TxLog.txt", "TxLog_" & Now.ToString("yyyyMMdd-hh.mm.ss") & ".txt")

                        ' 建立 Log Txt 檔
                        sTxLog = New FileStream(Application.StartupPath & "\SECSLog\TxLog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
                        sTxLog.Close()

                    End If
                End If

                Application.DoEvents()
                Thread.Sleep(1000)

            Loop While Not SocketState = enumSecsConnectionState.sSeparate
        End Sub


        ' TCP / IP 連線狀況
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

            End If

		End Sub


        ' 收到 SECS Message 
        Private Sub onMessage(ByRef message As Byte()) Implements IFMessageListener.onMessage

            Dim tempThread As Thread = New Thread(New ParameterizedThreadStart(AddressOf ParseMessage))
            tempThread.IsBackground = True
            tempThread.Start(message)

        End Sub


        ' 解析 SecsMessage
        Private Sub ParseMessage(ByVal message As Object)

            Try
                SyncLock thisLock

                    Dim temp() As Byte = CType(message, Byte())

                    ' NEW 一個 SECS Message
                    sSecsMessage = New SecsMessage(temp)


                    ' 檢查 SECS Message 是哪種 Message Type
                    Select Case sSecsMessage.CheckMessage(sSecsMessage)


                    ' 收到的 Message 是 Select Request
                        Case enumMessageType.sSelectRequest

                            ' NEW 一個 SecsTransaction
                            sTransaction = New SecsTransaction()
                            sTransaction.Primary = sSecsMessage
                            sTransaction.State = enumSecsTransactionState.PrimaryReceived

                            ' ReceiveTransactionMap 新增此 Transaction
                            SyncLock ReceiveTransactionMap
                                ReceiveTransactionMap.Add(sSecsMessage.messageFormat.SystemBytes, sTransaction)
                            End SyncLock

                            RaiseEvent OnPrimaryReceived(sTransaction)

                            ' 紀錄 Log
                            DoLog(sTransaction.Primary.ConvertToBinaryLog("RCV"), sTransaction.Primary.ConvertToSML("PrimaryIn"))


                            ' 設定 T6 Timeout
                            sTransaction.SetTimeout(enumTimeout.T6, siniFile)

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
                            sTransaction.Secondary = tempSecsMessage

                            SendSecondary(sTransaction)

                            ' 設定 SocketState 為 Selected
                            SocketState = enumSecsConnectionState.sSelected
                            RaiseEvent OnSecsConnectStateChange("Selected")

                            Exit Select

                    ' 收到的 Message 是 Select Response
                        Case enumMessageType.sSelectResponse

                            SyncLock SendTransactionMap

                                ' 從 SendTransactionMap 中，尋找 Transaction    
                                If SendTransactionMap.Count > 0 Then

                                    For Each item In SendTransactionMap

                                        If item.Key.ToString = sSecsMessage.messageFormat.SystemBytes.ToString Then

                                            ' 解除 T6 Timeout
                                            If item.Value.sTimeoutList.Count > 0 Then

                                                For i As Integer = 0 To item.Value.sTimeoutList.Count - 1 Step +1
                                                    If item.Value.sTimeoutList(i).sTimeoutType = enumTimeout.T6 Then

                                                        item.Value.sTimeoutList(i) = Nothing
                                                        item.Value.sTimeoutList.RemoveAt(i)
                                                        Exit For

                                                    End If
                                                Next

                                            End If

                                            item.Value.Secondary = sSecsMessage
                                            item.Value.State = enumSecsTransactionState.SecondaryReceived
                                            item.Value.State = enumSecsTransactionState.WaitForDelete

                                            RaiseEvent OnSecondaryReceived(item.Value)

                                            ' 紀錄 Log
                                            DoLog(item.Value.Secondary.ConvertToBinaryLog("RCV"), item.Value.Secondary.ConvertToSML("SecondaryIn"))

                                            ' 設定 SocketState 為 Selected
                                            SocketState = enumSecsConnectionState.sSelected
                                            RaiseEvent OnSecsConnectStateChange("Selected")

                                            Exit Select

                                        End If

                                    Next

                                    RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                                    Exit Select

                                Else
                                    RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                                    Exit Select
                                End If

                            End SyncLock


                    ' 收到的 Message 是 Linktest Request
                        Case enumMessageType.sLinktestRequest

                            ' NEW 一個 SecsTransaction
                            sTransaction = New SecsTransaction()
                            sTransaction.Primary = sSecsMessage
                            sTransaction.State = enumSecsTransactionState.PrimaryReceived

                            SyncLock ReceiveTransactionMap
                                ReceiveTransactionMap.Add(sSecsMessage.messageFormat.SystemBytes, sTransaction)
                            End SyncLock

                            RaiseEvent OnPrimaryReceived(sTransaction)

                            ' 紀錄 Log
                            DoLog(sTransaction.Primary.ConvertToBinaryLog("RCV"), sTransaction.Primary.ConvertToSML("PrimaryIn"))


                            ' 設定 T6 Timeout
                            sTransaction.SetTimeout(enumTimeout.T6, siniFile)

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

                            SendSecondary(sTransaction)

                            Exit Select

                    ' 收到的 Message 是 Linktest Reqsponse
                        Case enumMessageType.sLinktestResponse

                            SyncLock SendTransactionMap

                                ' 從 SendTransactionMap 中，尋找 Transaction 
                                If SendTransactionMap.Count > 0 Then

                                    For Each item In SendTransactionMap

                                        If item.Key.ToString = sSecsMessage.messageFormat.SystemBytes.ToString Then

                                            ' 解除 T6 Timeout
                                            If item.Value.sTimeoutList.Count > 0 Then

                                                For i As Integer = 0 To item.Value.sTimeoutList.Count - 1
                                                    If item.Value.sTimeoutList(i).sTimeoutType = enumTimeout.T6 Then

                                                        item.Value.sTimeoutList(i) = Nothing
                                                        item.Value.sTimeoutList.RemoveAt(i)
                                                        Exit For

                                                    End If
                                                Next

                                            End If

                                            item.Value.Secondary = sSecsMessage
                                            item.Value.State = enumSecsTransactionState.SecondaryReceived
                                            item.Value.State = enumSecsTransactionState.WaitForDelete

                                            RaiseEvent OnSecondaryReceived(item.Value)

                                            ' 紀錄 Log
                                            DoLog(item.Value.Secondary.ConvertToBinaryLog("RCV"), item.Value.Secondary.ConvertToSML("SecondaryIn"))

                                            Exit Select

                                        End If
                                    Next

                                    RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                                    Exit Select
                                Else
                                    RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                                    Exit Select
                                End If

                            End SyncLock


                    ' 收到的 Message 是 Data Message
                        Case enumMessageType.sDataMessage

                            If (sSecsMessage.Function Mod 2) = 1 Then

                                ' --------------------------------------- 收到 Primary Message --------------------------------------------

                                ' 從 SXML 中找出符合 SecsMessage 格式的 Message
                                If (sSXML.FindMessageByFormat(sSecsMessage)) = False Then

                                    ' 在 SXML 中沒有找到符合格式的 SecsMessage
                                    RaiseEvent OnMessageInfo(sSecsMessage, "Message Not Found")

                                Else

                                    ' 在 SXML 中找到符合格式的 SecsMessage

                                    ' NEW 一個 SecsTransaction
                                    sTransaction = New SecsTransaction()
                                    sTransaction.Primary = sSecsMessage
                                    sTransaction.State = enumSecsTransactionState.PrimaryReceived

                                    SyncLock ReceiveTransactionMap
                                        ReceiveTransactionMap.Add(sSecsMessage.messageFormat.SystemBytes, sTransaction)
                                    End SyncLock

                                    RaiseEvent OnPrimaryReceived(sTransaction)

                                    ' 紀錄 Log
                                    DoLog(sTransaction.Primary.ConvertToBinaryLog("RCV"), sTransaction.Primary.ConvertToSML("PrimaryIn"))


                                    ' 設定 T3 Timeout
                                    sTransaction.SetTimeout(enumTimeout.T3, siniFile)

                                    ' 是否執行 AutoReply
                                    DoAutoReply(sTransaction)

                                End If

                                Exit Select

                            Else

                                ' --------------------------------------- 收到 Secondary Message -----------------------------------------------

                                SyncLock SendTransactionMap

                                    ' 從 SendTransactionMap 中，尋找 Transaction
                                    If SendTransactionMap.Count > 0 Then

                                        For Each item In SendTransactionMap

                                            If item.Key.ToString = sSecsMessage.messageFormat.SystemBytes.ToString Then

                                                ' 從 SXML 中找出符合 SecsMessage 格式的 Message
                                                If (sSXML.FindMessageByFormat(sSecsMessage)) = False Then

                                                    RaiseEvent OnMessageInfo(sSecsMessage, "Message Not Found")
                                                Else

                                                    ' 解除 T3 Timeout
                                                    If item.Value.sTimeoutList.Count > 0 Then

                                                        For i As Integer = 0 To item.Value.sTimeoutList.Count - 1
                                                            If item.Value.sTimeoutList(i).sTimeoutType = enumTimeout.T3 Then

                                                                item.Value.sTimeoutList(i) = Nothing
                                                                item.Value.sTimeoutList.RemoveAt(i)
                                                                Exit For

                                                            End If
                                                        Next

                                                    End If

                                                    item.Value.Secondary = sSecsMessage
                                                    item.Value.State = enumSecsTransactionState.SecondaryReceived
                                                    item.Value.State = enumSecsTransactionState.WaitForDelete

                                                    RaiseEvent OnSecondaryReceived(item.Value)

                                                    ' 紀錄 Log
                                                    DoLog(item.Value.Secondary.ConvertToBinaryLog("RCV"), item.Value.Secondary.ConvertToSML("SecondaryIn"))

                                                End If

                                                Exit Select

                                            End If

                                        Next

                                        RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                                        Exit Select
                                    Else
                                        RaiseEvent OnMessageError(sSecsMessage, "Transaction Not Found")
                                        Exit Select
                                    End If

                                End SyncLock

                            End If

                    End Select

                End SyncLock

            Catch ex As Exception

                RaiseEvent OnMessageError(sSecsMessage, "Message Error")
            End Try

        End Sub


        ' 是否執行 AutoReply
        Private Sub DoAutoReply(ByRef aTransaction As SecsTransaction)

            Try
                If siniFile.AutoReply = True Then

                    If GetMessageByName(aTransaction.Primary.AutoReply) IsNot Nothing Then

                        aTransaction.Secondary = GetMessageByName(aTransaction.Primary.AutoReply)
                        aTransaction.Secondary.messageFormat.SystemBytes = aTransaction.Primary.messageFormat.SystemBytes
                        SendSecondary(aTransaction)

                    End If

                End If
            Catch ex As Exception

                RaiseEvent OnMessageError(aTransaction.Primary, "AutoReply Fail")
            End Try

		End Sub


		' 計算出 System Bytes
		Private Function getSystemByte() As UInteger

			systemBytesCount = systemBytesCount + 1

			If systemBytesCount >= 4294967295 Then
				systemBytesCount = 0
			End If

			Return systemBytesCount
		End Function


        ' 紀錄 Log
        Private Sub DoLog(ByRef aBinaryLogString As String, ByRef aTxLogString As String)

            ' 紀錄 Binary Log
            Try
                If siniFile.BinaryLog = True Then

                    Dim tempSW As StreamWriter = File.AppendText(Application.StartupPath & "\SECSLog\BinaryLog.txt")

                    tempSW.WriteLine(aBinaryLogString)
                    tempSW.Flush()
                    tempSW.Close()
                    tempSW.Dispose()
                End If

            Catch ex As Exception

                RaiseEvent OnError("Do BinaryLog Error !")

            End Try

            ' 紀錄 SML Log
            Try
                If siniFile.TxLog = True Then

                    Dim tempSW As StreamWriter = File.AppendText(Application.StartupPath & "\SECSLog\TxLog.txt")

                    tempSW.WriteLine(aTxLogString)
                    tempSW.Flush()
                    tempSW.Close()
                    tempSW.Dispose()
                End If
            Catch ex As Exception

                RaiseEvent OnError("Do BinaryLog Error !")

            End Try

        End Sub

    End Class

End Namespace

