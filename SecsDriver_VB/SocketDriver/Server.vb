Imports System
Imports System.Net.Sockets
Imports System.Threading


Namespace SocketDriver


    ''' <summary>
    ''' TCP Server
    ''' </summary>
    Public Class Server : Implements SocketDriver.IFTcpServer


#Region "Private 屬性"

        Private TCPIPAddress As System.Net.IPAddress                ' TCP IP Address 物件
        Private TCPPort As Integer                                  ' TCP Port
        Public TCPListener As TcpListener                           ' TCP Listener 物件
        Private TCPSocket As Socket                                 ' TCP Socket 物件

        Private listener As IFMessageListener                       ' Message Listener

        Private MutiConnect As Boolean                              ' 是否允許多 Client 連線

        Private startReceive As Thread                              ' 負責 Receive Message 的 Thread

        Private bufferList As ArrayList                             ' 暫存區

        Private sTimeout As Timeout                                 ' Timeout

#End Region


#Region "建構子"

        ''' <summary>
        ''' 建構子
        ''' </summary>
        ''' <param name="slistener"></param>
        ''' <param name="sMutiConnect"></param>
        Public Sub New(ByVal slistener As IFMessageListener, ByVal sMutiConnect As Boolean)

            TCPIPAddress = System.Net.IPAddress.Parse("127.0.0.1")
            TCPPort = 36000
            listener = slistener
            MutiConnect = sMutiConnect

            ' New ArrayList (暫存區)
            bufferList = New ArrayList

            ' 連線
            connect()

        End Sub


        ''' <summary>
        ''' 建構子
        ''' </summary>
        ''' <param name="slistener"></param>
        ''' <param name="sIPAddress"></param>
        ''' <param name="sPort"></param>
        ''' <param name="sMutiConnect"></param>
        Public Sub New(ByVal slistener As IFMessageListener, ByVal sIPAddress As String, ByVal sPort As Integer, ByVal sMutiConnect As Boolean)


            TCPIPAddress = System.Net.IPAddress.Parse(sIPAddress)   ' 設定 IP
            TCPPort = sPort                                         ' 設定 Port

            listener = slistener                                    ' 設定 Listener
            MutiConnect = sMutiConnect

            ' New ArrayList (暫存區)
            bufferList = New ArrayList

            ' 連線
            connect()

        End Sub


#End Region


#Region "Public Method"

        ''' <summary>
        ''' 連線
        ''' </summary>
        Public Sub connect() Implements IFTcpServer.connect

			Try
                TCPListener = New TcpListener(TCPIPAddress, TCPPort)

                TCPListener.Start()
                TCPSocket = TCPListener.AcceptSocket()
                TCPSocket.NoDelay = True

                If MutiConnect = True Then

                Else
                    ' 接受來自 Client 的連線後，便關閉 TCP Listener
                    TCPListener.Stop()
                End If

                listener.sysMessage("Connected")


                ' New Thread (負責 Receive Message 的 Thread)
                startReceive = New Thread(AddressOf receive)
                startReceive.IsBackground = True
                startReceive.Start()

			Catch e As Exception

                listener.sysMessage("NotConnected")
            End Try

		End Sub


        ''' <summary>
        ''' 重新連線
        ''' </summary>
        Public Sub Reconnect()

            Try
                TCPSocket.Close()
                TCPSocket.Dispose()
                TCPListener.Stop()

            Catch e As Exception

                listener.sysMessage("NotConnected")
            End Try

            ' 連線
            connect()

        End Sub


        ''' <summary>
        ''' 關閉連線
        ''' </summary>
        Public Sub disconnect() Implements IFTcpServer.disconnect

            startReceive.Abort()
            TCPSocket.Close()

		End Sub


        ''' <summary>
        ''' Send Message
        ''' </summary>
        ''' <param name="message"></param>
        Public Sub send(ByVal message As Byte()) Implements IFTcpServer.send

            Try
                If TCPSocket.Connected Then

                    TCPSocket.Send(message)
                Else

                    listener.sysMessage("NotConnected")
                End If

            Catch ex As Exception

                listener.sysMessage("NotConnected")
            End Try
        End Sub


        ''' <summary>
        ''' Receive Message
        ''' </summary>
        Public Sub receive() Implements IFTcpServer.receive

            ' 暫存變數 : 是否繼續 Receive 迴圈
            Dim IsOpen As Boolean = True

            ' 測試連線是否存在所使用的 Byte
            Dim testByte(1) As Byte

            ' Socket連線是否已經斷線
            Dim IsSocketClosed As Boolean = False

            While IsOpen

                Try

                    ' 檢查連線
                    If TCPSocket.Poll(1, SelectMode.SelectRead) = False Then

                        ' 使用 Peek 測試連線是否還存在
                        ' 如果對方斷線時，會在這邊處理
                        If TCPSocket.Receive(testByte, SocketFlags.Peek) = 0 Then

                            IsOpen = False
                            listener.sysMessage("NotConnected")
                            Exit Try

                        End If

                    End If

                    ' 取得 Socket 狀態
                    If TCPSocket.Connected Then

                        ' 當有收到資料時，才 Receive Message
                        If TCPSocket.Available > 0 Then

                            ' 暫存變數 : 將資料先儲存到此變數
                            Dim tempBytes(TCPSocket.Available - 1) As Byte

                            ' 當收到的資料位元組數為 0 時
                            If TCPSocket.Receive(tempBytes) = 0 Then

                                'Application.DoEvents()
                                Thread.Sleep(1000)

                            Else

                                ' 將資料收到暫存區
                                For i As Integer = 0 To tempBytes.Length - 1 Step +1

                                    bufferList.Add(tempBytes(i))
                                Next

                                If bufferList.Count < 4 Then
                                    Exit Try
                                End If

                                ' 取前四個 Bytes 來計算出 Message Length
                                Dim temp(3) As Byte
                                For i As Integer = 0 To 3 Step +1
                                    temp(i) = bufferList(i)
                                Next
                                Array.Reverse(temp)
                                Dim MessageLength As Integer = BitConverter.ToInt32(temp, 0)


                                ' 如果 Bytes 數量能組成 SecsMessage 時，便傳出 Message
                                ' 如果 Bytes 數量不能組成 SecsMessage 時，便設定 T8 Timeout
                                If bufferList.Count >= MessageLength + 4 Then

                                    Dim messageByte As Byte() = New Byte(MessageLength + 3) {}

                                    For i As Integer = 0 To MessageLength + 3 Step +1

                                        messageByte(i) = bufferList(0)
                                        bufferList.RemoveAt(0)
                                    Next

                                    listener.onMessage(messageByte)

                                Else
                                    listener.sysMessage("Set T8 Timeout")
                                End If

                            End If

                        End If

                    End If

                Catch e As Exception

                    IsOpen = False
                    listener.sysMessage("Separate")

                End Try

            End While

		End Sub

#End Region


    End Class

End Namespace


