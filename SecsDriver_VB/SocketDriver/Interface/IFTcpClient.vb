Imports System

Namespace SocketDriver

	' TcpClient 的 Interface
	Public Interface IFTcpClient

		' TCP Client 連線
		Sub connect()

		' TCP Client 斷線
		Sub disconnect()

		' TCP Client Send Message
		Sub send(ByVal message As Byte())

		' TCP Client Receive Message
		Sub receive()

	End Interface

End Namespace


