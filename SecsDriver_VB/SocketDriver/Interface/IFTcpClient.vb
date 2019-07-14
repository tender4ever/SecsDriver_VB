Imports System

Namespace SocketDriver

    ''' <summary>
    ''' TcpClient 的 Interface
    ''' </summary>
    Public Interface IFTcpClient

        ''' <summary>
        ''' TCP Client 連線
        ''' </summary>
        Sub connect()

        ''' <summary>
        ''' TCP Client 斷線
        ''' </summary>
        Sub disconnect()

        ''' <summary>
        ''' TCP Client Send Message
        ''' </summary>
        ''' <param name="message"></param>
        Sub send(ByVal message As Byte())

        ''' <summary>
        ''' TCP Client Receive Message
        ''' </summary>
        Sub receive()

	End Interface

End Namespace


