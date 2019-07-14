Imports System

Namespace SocketDriver

    ''' <summary>
    ''' TcpServer 的 Interface
    ''' </summary>
    Public Interface IFTcpServer

        ''' <summary>
        ''' TCP Server 連線
        ''' </summary>
        Sub connect()

        ''' <summary>
        ''' TCP Server 斷線
        ''' </summary>
        Sub disconnect()

        ''' <summary>
        ''' TCP Server Send Message
        ''' </summary>
        ''' <param name="message"></param>
        Sub send(ByVal message As Byte())

        ''' <summary>
        ''' TCP Server Receive Message
        ''' </summary>
        Sub receive()

	End Interface

End Namespace


