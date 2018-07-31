Imports System

Namespace SocketDriver

	' MessageListener 的 Interface
	Public Interface IFMessageListener

		' Message 訊息
		Sub onMessage(ByRef message As Byte())

		' 系統訊息
		Sub sysMessage(ByVal sysMessage As String)

	End Interface

End Namespace


