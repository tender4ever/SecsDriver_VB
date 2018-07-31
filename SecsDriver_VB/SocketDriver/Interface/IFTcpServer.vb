Imports System

Namespace SocketDriver

	' TcpServer 的 Interface
	Public Interface IFTcpServer

		' TCP Server 連線
		Sub connect()

		' TCP Server 斷線
		Sub disconnect()

		' TCP Server Send Message
		Sub send(ByVal message As Byte())

		' TCP Server Receive Message
		Sub receive()

	End Interface

End Namespace


