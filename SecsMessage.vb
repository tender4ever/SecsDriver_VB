﻿Imports System
Imports System.Text

Namespace SecsDriver

    Public Class SecsMessage : Implements ICloneable, IDisposable

        Private sStream As UInteger                     ' SML Stream Code
        Private sFunction As UInteger                   ' SML Function Code
        Private sWBit As Boolean                        ' WBit
        Private sMessageType As enumMessageType         ' Message Type

        Public messageFormat As MessageFormat           ' Message Format
        Public secsItem As SecsItem                     ' SecsItem


        ' ----------------- Property ------------------------

        ' 存取 MessageName
        Public Property MessageName As String

        ' 存取 MessageDescription
        Public Property MessageDescription As String

        ' 存取 AutoReply
        Public Property AutoReply As String


        ' 存取 Stream
        Public Property Stream As UInteger

            Get
                sStream = Convert.ToUInt16(messageFormat.HeaderByte(0))

                Return sStream
            End Get

            Set(ByVal value As UInteger)

                messageFormat.HeaderByte(0) = Convert.ToByte(value)

            End Set

        End Property


        ' 存取 Function
        Public Property [Function] As UInteger

            Get
                sFunction = Convert.ToUInt16(messageFormat.HeaderByte(1))

                Return sFunction
            End Get

            Set(ByVal value As UInteger)

                messageFormat.HeaderByte(1) = Convert.ToByte(value)

            End Set

        End Property


        ' 存取 WBit
        Public Property WBit As Boolean

            Get
                sWBit = messageFormat.WBit
                Return sWBit
            End Get

            Set(ByVal value As Boolean)

                messageFormat.WBit = value

            End Set

        End Property


        ' 存取 Message Type
        Public Property MessageType As enumMessageType
            Get
                Select Case Convert.ToInt16(messageFormat.SType)

                    Case 0
                        sMessageType = enumMessageType.sDataMessage

                    Case 1
                        sMessageType = enumMessageType.sSelectRequest

                    Case 2
                        sMessageType = enumMessageType.sSelectResponse

                    Case 3
                        sMessageType = enumMessageType.sDeselectRequest

                    Case 4
                        sMessageType = enumMessageType.sDeselectResponse

                    Case 5
                        sMessageType = enumMessageType.sLinktestRequest

                    Case 6
                        sMessageType = enumMessageType.sLinktestResponse

                    Case 7
                        sMessageType = enumMessageType.sRejectRequest

                    Case 9
                        sMessageType = enumMessageType.sSeparateRequest

                End Select

                Return sMessageType
            End Get
            Set(value As enumMessageType)

                Select Case value

                    Case enumMessageType.sDataMessage
                        messageFormat.SType = Convert.ToByte(0)

                    Case enumMessageType.sSelectRequest
                        messageFormat.SType = Convert.ToByte(1)

                    Case enumMessageType.sSelectResponse
                        messageFormat.SType = Convert.ToByte(2)

                    Case enumMessageType.sDeselectRequest
                        messageFormat.SType = Convert.ToByte(3)

                    Case enumMessageType.sDeselectResponse
                        messageFormat.SType = Convert.ToByte(4)

                    Case enumMessageType.sLinktestRequest
                        messageFormat.SType = Convert.ToByte(5)

                    Case enumMessageType.sLinktestResponse
                        messageFormat.SType = Convert.ToByte(6)

                    Case enumMessageType.sRejectRequest
                        messageFormat.SType = Convert.ToByte(7)

                    Case enumMessageType.sSeparateRequest
                        messageFormat.SType = Convert.ToByte(9)
                End Select

            End Set
        End Property


        ' ----------------- Method ------------------------

        ' 建構子，新增 未設值得 SecsMessage
        Public Sub New()

            ' New MessageFormat
            messageFormat = New MessageFormat()
            ' New SecsItem
            secsItem = New SecsItem()

        End Sub


        ' 建構子，收到 Bytes 時，使用此方法來新增 SecsMessage
        Public Sub New(ByRef secsMessage As Byte())
            Try
                ' 把 Bytes 解成 Message Format
                Dim byteValue1 As Byte() = New Byte(13) {}
                Array.Copy(secsMessage, 0, byteValue1, 0, 14)
                messageFormat = New MessageFormat(byteValue1)

                ' 把 Bytes 解成 SecsItem
                If (secsMessage.Length - 14) > 0 Then

                    Dim byteValue2 As Byte() = New Byte(secsMessage.Length - 14 - 1) {}
                    Array.Copy(secsMessage, 14, byteValue2, 0, secsMessage.Length - 14)
                    secsItem = New SecsItem(byteValue2)

                Else
                    secsItem = Nothing

                End If

            Catch ex As Exception

                messageFormat = Nothing
                secsItem = Nothing

            End Try

        End Sub


        ' 檢查 SecsMessage 的格式
        Public Function CheckMessageFormat() As String

            Dim temp As String

            ' S?F?
            temp = "S" & Stream.ToString() & "F" + [Function].ToString()

            If WBit = True Then

                ' // S?F? + W
                temp += " [W] " + vbLf

            End If

            ' S?F? + W + Item
            If secsItem IsNot Nothing Then

                temp += CheckItemFormat(secsItem)

            End If

            Return temp

        End Function


        ' 檢查 SecsItem 的格式
        Private Function CheckItemFormat(ByRef item As SecsItem) As String

            Dim temp As String = ""

            If item.ItemType = "L" Then

                ' 當 item 為 List 時

                ' <ItemType [ItemNumber]
                temp += "<" & item.ItemType & vbLf

                For i As Integer = 0 To item.ItemList.Count - 1

                    temp += CheckItemFormat(item.ItemList(i))

                Next

                temp += ">"

            Else

                ' 當 item 不為 List 時

                ' <ItemType [ItemNumber] >
                temp += "<" & item.ItemType & ">" + vbLf

            End If

            Return temp

        End Function


        ' 將 SecsMessage 轉換成 Bytes
        Public Function ConvertToBytes() As Byte()

            ' 將 SecsItem 轉換成 Bytes
            Dim ItemTempList As ArrayList = New ArrayList

            If Me.secsItem IsNot Nothing Then

                ItemConvertToBytes(Me.secsItem, ItemTempList)
            End If

            Me.messageFormat.Length = BitConverter.GetBytes(ItemTempList.Count + 10)
            Array.Reverse(Me.messageFormat.Length)

            ' 將 SecsMessage 轉換成 Bytes
            Dim MessageTempList As ArrayList = New ArrayList

            ' Length
            MessageTempList.Add(Me.messageFormat.Length(0))
            MessageTempList.Add(Me.messageFormat.Length(1))
            MessageTempList.Add(Me.messageFormat.Length(2))
            MessageTempList.Add(Me.messageFormat.Length(3))

            ' DeviceID
            If Me.messageFormat.RBit = True Then
                MessageTempList.Add((Me.messageFormat.DeviceID(0) Or &H80))
            Else
                MessageTempList.Add(Me.messageFormat.DeviceID(0))
            End If
            MessageTempList.Add(Me.messageFormat.DeviceID(1))

            ' Header Bytes
            If Me.messageFormat.WBit = True Then
                MessageTempList.Add((Me.messageFormat.HeaderByte(0) Or &H80))
            Else
                MessageTempList.Add(Me.messageFormat.HeaderByte(0))
            End If
            MessageTempList.Add(Me.messageFormat.HeaderByte(1))

            ' PType
            MessageTempList.Add(Me.messageFormat.PType)

            ' SType
            MessageTempList.Add(Me.messageFormat.SType)

            ' System Bytes
            MessageTempList.Add(Me.messageFormat.SystemBytes(0))
            MessageTempList.Add(Me.messageFormat.SystemBytes(1))
            MessageTempList.Add(Me.messageFormat.SystemBytes(2))
            MessageTempList.Add(Me.messageFormat.SystemBytes(3))

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


        ' 將 SecsItem 轉換成 Bytes
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


        ' 檢查 Message 是哪種 Message Type
        Public Function CheckMessage(ByRef aMessage As SecsMessage) As enumMessageType

            ' 計算 Message Length
            Dim MessageLength As Integer = Nothing
            Array.Reverse(aMessage.messageFormat.Length)
            MessageLength = BitConverter.ToInt32(aMessage.messageFormat.Length, 0)
            Array.Reverse(aMessage.messageFormat.Length)


            If MessageLength = 10 And
                Convert.ToInt16(aMessage.messageFormat.DeviceID(0)) = 127 And
                Convert.ToInt16(aMessage.messageFormat.DeviceID(1)) = 255 And
                Convert.ToInt16(aMessage.messageFormat.HeaderByte(0)) = 0 And
                Convert.ToInt16(aMessage.messageFormat.HeaderByte(1)) = 0 And
                Convert.ToInt16(aMessage.messageFormat.PType) = 0 And
                Convert.ToInt16(aMessage.messageFormat.SType) = 1 Then

                ' Control Message - Select Request
                aMessage.MessageType = enumMessageType.sSelectRequest
                Return aMessage.MessageType

            ElseIf MessageLength = 10 And
                Convert.ToInt16(aMessage.messageFormat.DeviceID(0)) = 127 And
                Convert.ToInt16(aMessage.messageFormat.DeviceID(1)) = 255 And
                Convert.ToInt16(aMessage.messageFormat.HeaderByte(0)) = 0 And
                Convert.ToInt16(aMessage.messageFormat.HeaderByte(1)) = 0 And
                Convert.ToInt16(aMessage.messageFormat.PType) = 0 And
                Convert.ToInt16(aMessage.messageFormat.SType) = 2 Then

                ' Control Message - Select Response
                aMessage.MessageType = enumMessageType.sSelectResponse
                Return aMessage.MessageType

            ElseIf MessageLength = 10 And
                Convert.ToInt16(aMessage.messageFormat.DeviceID(0)) = 127 And
                Convert.ToInt16(aMessage.messageFormat.DeviceID(1)) = 255 And
                Convert.ToInt16(aMessage.messageFormat.HeaderByte(0)) = 0 And
                Convert.ToInt16(aMessage.messageFormat.HeaderByte(1)) = 0 And
                Convert.ToInt16(aMessage.messageFormat.PType) = 0 And
                Convert.ToInt16(aMessage.messageFormat.SType) = 5 Then

                ' Control Message - Linktest Request
                aMessage.MessageType = enumMessageType.sLinktestRequest
                Return aMessage.MessageType

            ElseIf MessageLength = 10 And
                Convert.ToInt16(aMessage.messageFormat.DeviceID(0)) = 127 And
                Convert.ToInt16(aMessage.messageFormat.DeviceID(1)) = 255 And
                Convert.ToInt16(aMessage.messageFormat.HeaderByte(0)) = 0 And
                Convert.ToInt16(aMessage.messageFormat.HeaderByte(1)) = 0 And
                Convert.ToInt16(aMessage.messageFormat.PType) = 0 And
                Convert.ToInt16(aMessage.messageFormat.SType) = 6 Then

                ' Control Message - Linktest Response
                aMessage.MessageType = enumMessageType.sLinktestResponse
                Return aMessage.MessageType

            ElseIf MessageLength = 10 And
                Convert.ToInt16(aMessage.messageFormat.DeviceID(0)) = 127 And
                Convert.ToInt16(aMessage.messageFormat.DeviceID(1)) = 255 And
                Convert.ToInt16(aMessage.messageFormat.HeaderByte(0)) = 0 And
                Convert.ToInt16(aMessage.messageFormat.HeaderByte(1)) = 0 And
                Convert.ToInt16(aMessage.messageFormat.PType) = 0 And
                Convert.ToInt16(aMessage.messageFormat.SType) = 9 Then

                ' Control Message - SeparateRequest
                aMessage.MessageType = enumMessageType.sSeparateRequest
                Return aMessage.MessageType

            Else
                ' Data Message
                aMessage.MessageType = enumMessageType.sDataMessage
                Return aMessage.MessageType

            End If

        End Function


        ' 將 SecsMessage 轉換成 SML
        Public Function ConvertToSML(ByRef aType As String) As String

            Dim temp As String

            ' Time
            temp = Now.ToString("yyyy/MM/dd hh:mm:ss.fff") & " [" & aType

            ' S?F?
            temp += " S" & Stream.ToString() & "F" + [Function].ToString()

            ' S?F? + W
            If WBit = True Then

                temp += " W]"
            Else
                temp += "]"

            End If

            ' S?F? + W + MessageName
            If MessageName IsNot Nothing Then

                temp += " " & MessageName + vbCrLf
            Else
                temp += vbCrLf
            End If

            ' Message Header
            temp += "[" & messageFormat.DeviceID(0).ToString("X2") & " " & messageFormat.DeviceID(1).ToString("X2") & " "
            temp += messageFormat.HeaderByte(0).ToString("X2") & " " & messageFormat.HeaderByte(1).ToString("X2") & " "
            temp += messageFormat.PType.ToString("X2") & " " & messageFormat.SType.ToString("X2") & " "
            temp += messageFormat.SystemBytes(0).ToString("X2") & " " & messageFormat.SystemBytes(1).ToString("X2") & " " & messageFormat.SystemBytes(2).ToString("X2") & " " & messageFormat.SystemBytes(3).ToString("X2") & "]" & vbCrLf

            ' S?F? + W + Item
            If secsItem IsNot Nothing Then

                temp += itemConvertToSML(secsItem, 0) & vbCrLf
            Else
                temp += vbCrLf
            End If

            Return temp

        End Function


        ' 將 SecsItem 轉換成 SML
        Private Function itemConvertToSML(ByRef item As SecsItem, ByRef tabCount As Integer) As String

            Dim temp As String = ""

            Dim tabString As String = ""

            If tabCount > 0 Then

                For i As Integer = 1 To tabCount Step +1
                    tabString += vbTab
                Next

            End If

            If item.ItemType = "L" Then

                ' ------------------------------- 當 item 為 List 時 -----------------------------------

                ' <ItemType [ItemNumber]
                temp += tabString & "<" & item.ItemType & " [" & item.ItemNumber.ToString() & "]"

                ' <ItemType [ItemNumber][ItemName]
                If item.ItemName IsNot Nothing Then
                    temp += "[" & item.ItemName & "]" & vbCrLf
                Else
                    temp += "[]" & vbCrLf
                End If

                ' 解 ItemList
                For i As Integer = 0 To item.ItemList.Count - 1

                    temp += itemConvertToSML(item.ItemList(i), tabCount + 1)
                Next

                temp += tabString & ">" & vbCrLf

            Else

                ' -------------------------------- 當 item 不為 List 時 ---------------------------------

                Dim itemString As String = Nothing

                If item.ItemValue IsNot Nothing Then

                    For i As Integer = 1 To item.ItemValue.Count
                        itemString = itemString + Convert.ToString(item.ItemValue(i)) + " "
                    Next

                End If

                ' <ItemType [ItemNumber][ItemName] 'ItemValue'>
                If item.ItemName IsNot Nothing Then

                    temp += tabString & "<" & item.ItemType & " [" & item.ItemNumber.ToString() & "][" & item.ItemName & "] '" & itemString & "'>" + vbCrLf
                Else

                    temp += tabString & "<" & item.ItemType & " [" & item.ItemNumber.ToString() & "] '" & itemString & "'>" + vbCrLf
                End If

            End If

            Return temp

        End Function


        ' 將 SecsMessage 轉換成 BinaryLog
        Public Function ConvertToBinaryLog(ByRef aType As String) As String

            Dim temp As String

            ' Time
            temp = Now.ToString("yyyy/MM/dd hh:mm:ss.fff") & " [" & aType & " S" & Stream.ToString() & "F" + [Function].ToString() & "] ["

            ' Message Header
            temp += messageFormat.DeviceID(0).ToString("X2") & " " & messageFormat.DeviceID(1).ToString("X2") & " "
            temp += messageFormat.HeaderByte(0).ToString("X2") & " " & messageFormat.HeaderByte(1).ToString("X2") & " "
            temp += messageFormat.PType.ToString("X2") & " " & messageFormat.SType.ToString("X2") & " "
            temp += messageFormat.SystemBytes(0).ToString("X2") & " " & messageFormat.SystemBytes(1).ToString("X2") & " " & messageFormat.SystemBytes(2).ToString("X2") & " " & messageFormat.SystemBytes(3).ToString("X2") & "] "

            ' SecsItem
            If secsItem IsNot Nothing Then

                temp += secsItem.ConvertToBinaryLog()
            End If

            Return temp

        End Function


        ' Clone 
        Public Function Clone() As Object Implements ICloneable.Clone

            Dim secsMessage As New SecsMessage

            secsMessage.MessageName = Me.MessageName
            secsMessage.MessageDescription = Me.MessageDescription
            secsMessage.sStream = Me.sStream
            secsMessage.sFunction = Me.sFunction
            secsMessage.sWBit = Me.sWBit
            secsMessage.AutoReply = Me.AutoReply
            secsMessage.messageFormat = Me.messageFormat.Clone()

            If Me.secsItem IsNot Nothing Then
                secsMessage.secsItem = Me.secsItem.Clone()
            Else
                secsMessage.secsItem = Nothing
            End If

            Return secsMessage

		End Function


        ' Dispose
        Public Sub Dispose() Implements IDisposable.Dispose

            Me.messageFormat = Nothing
            Me.secsItem = Nothing

            GC.SuppressFinalize(Me)

        End Sub

    End Class

End Namespace

