Imports System

Namespace SecsDriver

    Public Class MessageFormat : Implements ICloneable

        ' Length (4 Bytes)
        Protected sLength As Byte()                 ' Length - Length Bytes

		' HSMS Message Header (10 Bytes)
		Protected sRBit As Boolean                  ' HSMS Message Header - RBit
		Protected sDeviceID As Byte()               ' HSMS Message Header - DeviceID
		Protected sWBit As Boolean                  ' HSMS Message Header - WBit
		Protected sHeaderByte As Byte()             ' HSMS Message Header - Header Byte
		Protected sPType As Byte                    ' HSMS Message Header - PType
		Protected sSType As Byte                    ' HSMS Message Header - SType
		Protected sSystemBytes As Byte()            ' HSMS Message Header - System Bytes

        ' Lock 物件
        Private Shared thisLock As Object = New Object


        ' ---------------- Property ----------------------

        ' 存取 Length
        Public Property Length As Byte()

			Get
				Return sLength
			End Get

			Set(ByVal value As Byte())
				sLength = value
			End Set

		End Property


		' 存取 RBit
		Public Property RBit As Boolean

			Get
				Return sRBit
			End Get

			Set(ByVal value As Boolean)
				sRBit = value
			End Set

		End Property


		' 存取 DeviceID
		Public Property DeviceID As Byte()

			Get
				Return sDeviceID
			End Get

			Set(ByVal value As Byte())
				sDeviceID = value
			End Set

		End Property


		' 存取 WBit
		Public Property WBit As Boolean

			Get
				Return sWBit
			End Get

			Set(ByVal value As Boolean)
				sWBit = value
			End Set

		End Property


		' 存取 HeaderByte
		Public Property HeaderByte As Byte()

			Get
				Return sHeaderByte
			End Get

			Set(ByVal value As Byte())
				sHeaderByte = value
			End Set

		End Property


		' 存取 PType
		Public Property PType As Byte

			Get
				Return sPType
			End Get

			Set(ByVal value As Byte)
				sPType = value
			End Set

		End Property


		' 存取 SType
		Public Property SType As Byte

			Get
				Return sSType
			End Get

			Set(ByVal value As Byte)
				sSType = value
			End Set

		End Property


		' 存取 SystemBytes
		Public Property SystemBytes As Byte()

			Get
				Return sSystemBytes
			End Get

			Set(ByVal value As Byte())
				sSystemBytes = value
			End Set

		End Property


        ' ---------------- Method ----------------------

        ' 建構子，新增 未設值的 MessageFormat
        Public Sub New()

            Try
                ' Length (4 Bytes)
                Me.sLength = New Byte(3) {}             ' 4 Bytes 的 Length Bytes

                ' HSMS Message Header (10 Bytes)
                Me.sDeviceID = New Byte(1) {}           ' 2 Bytes Device ID(upper device ID、lower device ID)
                Me.sHeaderByte = New Byte(1) {}         ' 2 Bytes Header Byte(Header Byte2、Header Byte3)
                Me.sSystemBytes = New Byte(3) {}        ' 4 Bytes System Bytes
                Me.sRBit = False                        ' RBit
                Me.sWBit = False                        ' WBit
                Me.sPType = 0                           ' PType
                Me.sSType = 0                           ' SType

            Catch ex As Exception

                Me.sLength = Nothing
                Me.sDeviceID = Nothing
                Me.sHeaderByte = Nothing
                Me.sSystemBytes = Nothing

            End Try

        End Sub


        ' 建構子，收到 Bytes 時，使用此方法來新增 MessageFormat
        Public Sub New(ByRef aMessageFormat As Byte())

            Try
                ' Length (4 Bytes)
                Me.sLength = New Byte(3) {}             ' 4 Bytes 的 Length Bytes

                ' HSMS Message Header (10 Bytes)
                Me.sDeviceID = New Byte(1) {}           ' 2 Bytes Device ID(upper device ID、lower device ID)
                Me.sHeaderByte = New Byte(1) {}         ' 2 Bytes Header Byte(Header Byte2、Header Byte3)
                Me.sSystemBytes = New Byte(3) {}        ' 4 Bytes System Bytes
                Me.sRBit = False                        ' RBit
                Me.sWBit = False                        ' WBit
                Me.sPType = 0                           ' PType
                Me.sSType = 0                           ' SType

                Array.Copy(aMessageFormat, 0, Me.sLength, 0, 4)         ' Length Bytes
                Array.Copy(aMessageFormat, 4, Me.sDeviceID, 0, 2)       ' Device ID
                Array.Copy(aMessageFormat, 6, Me.sHeaderByte, 0, 2)     ' Header Byte
                Me.PType = aMessageFormat(8)                            ' PType
                Me.SType = aMessageFormat(9)                            ' SType
                Array.Copy(aMessageFormat, 10, Me.sSystemBytes, 0, 4)   ' System Bytes


                ' 當 (upper device) & 1000 0000 = 1000 0000 
                ' 代表  (RBit = 1)
                ' 而 upper device = (upper device) & 0111 1111
                If (sDeviceID(0) And &H80) = &H80 Then

                    ' 設定 HSMS Message Header - upper device ID
                    Me.sDeviceID(0) = CByte((sDeviceID(0) And &H7F))

                    Me.sRBit = True
                Else

                    Me.sRBit = False
                End If


                ' 當 (Header Byte 2) & 1000 0000 = 1000 0000
                ' 代表 (WBit = 1)
                ' 而 Header Byte 2 = (Header Byte 2) & 0111 1111
                If (sHeaderByte(0) And &H80) = &H80 Then

                    ' 設定 HSMS Message Header - Header Byte 2 
                    Me.sHeaderByte(0) = CByte((sHeaderByte(0) And &H7F))

                    Me.sWBit = True
                Else

                    Me.sWBit = False
                End If

            Catch ex As Exception

                Me.sLength = Nothing
                Me.sDeviceID = Nothing
                Me.sHeaderByte = Nothing
                Me.sSystemBytes = Nothing

            End Try

        End Sub


        ' Clone 
        Public Function Clone() As Object Implements ICloneable.Clone

            SyncLock thisLock

                ' 宣告 MessageFormat 物件
                Dim messageFormat As MessageFormat = New MessageFormat

                Try
                    Array.Copy(Me.sLength, messageFormat.sLength, 4)                ' Length Bytes
                    Array.Copy(Me.sDeviceID, messageFormat.sDeviceID, 2)            ' Device ID
                    Array.Copy(Me.sHeaderByte, messageFormat.sHeaderByte, 2)        ' Header Byte
                    Array.Copy(Me.sSystemBytes, messageFormat.sSystemBytes, 4)      ' System Bytes
                    messageFormat.sRBit = Me.sRBit                                  ' RBit
                    messageFormat.sWBit = Me.sWBit                                  ' WBit
                    messageFormat.sPType = Me.sPType                                ' PType
                    messageFormat.sSType = Me.sSType                                ' SType

                    Return messageFormat

                Catch ex As Exception

                    messageFormat.Finalize()
                    Return Nothing

                End Try

            End SyncLock

        End Function

    End Class

End Namespace


