﻿Imports System
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Linq
Imports System.Xml.Linq

Namespace SecsDriver

	Public Class SXML

		Private sMessageList As List(Of SecsMessage)                    ' Message List
		Private sMessageMap As Dictionary(Of String, SecsMessage)       ' Message Map


        ' 建構子
        Public Sub New()

            sMessageList = New List(Of SecsMessage)()
            sMessageMap = New Dictionary(Of String, SecsMessage)

        End Sub


        ' -------------------- Property ------------------------

        ' 存取 File Path
        Public Property FilePath As String


        ' -------------------- Method ------------------------

        ' 讀取 SXML 的 Message 到 MessageList、MessageMap
        Public Sub LoadMessage()

            ' 讀取所有的 SecsMessage
            Dim allMessage = XDocument.Load(FilePath).Root.Element("MessageDictionary").Elements("Message")

            '解析每個 SecsMessage
            For Each message In allMessage

				' New SecsMessage
				Dim secsMessage As SecsMessage = New SecsMessage()

                ' 設定 MessageName
                secsMessage.MessageName = message.Attribute("Name").Value

                ' 設定 MessageDescription
                secsMessage.MessageDescription = message.Attribute("Description").Value

                ' 設定 Stream
                secsMessage.Stream = UInteger.Parse(message.Element("Stream"))

                ' 設定 Function
                secsMessage.[Function] = Integer.Parse(message.Element("Function"))

                ' 設定 WBit
                If message.Element("WBit") IsNot Nothing Then

                    If Convert.ToUInt16(message.Element("WBit").Value) = 1 Then
                        secsMessage.WBit = True
                    Else
                        secsMessage.WBit = False
                    End If
                End If

                ' 設定 AutoReply
                If message.Element("AutoReply") IsNot Nothing Then

					secsMessage.AutoReply = message.Element("AutoReply").Value
				End If


				' 設定 SecsItem
				If message.Element("Structure") IsNot Nothing Then

					If message.Element("Structure").Elements.Count > 0 Then

						Dim list As List(Of XNode) = message.Element("Structure").DescendantNodes().ToList()

                        ' 讀取 SXML 中各 Message 的 Item
                        secsMessage.secsItem = LoadItem(list)

                        ' 讀取 Item 後，加到 ItemMap
                        LoadItemMap(secsMessage.secsItem, secsMessage.secsItem.ItemMap)
					Else
						secsMessage.secsItem = Nothing
					End If

				End If

				sMessageList.Add(secsMessage)
				sMessageMap.Add(secsMessage.MessageName, secsMessage)
			Next

		End Sub


        ' 讀取 SXML 中各 Message 的 Item 
        Private Function LoadItem(ByRef list As List(Of XNode)) As SecsItem

			Try
                ' Root Item
                Dim RootItem As XElement = DirectCast(list(0), XElement)

                ' New SecsItem
                Dim aSecsItem As SecsItem = New SecsItem

				' 設定 Item Type 
				aSecsItem.ItemType = RootItem.Name.LocalName

                If aSecsItem.ItemType = "L" Then

                    ' ------- Item Type = List -------

                    ' 設定 Item Number
                    aSecsItem.ItemNumber = RootItem.Elements.Count

                    ' 設定 Item Name
                    If RootItem.Attribute("Name") Is Nothing Then
                        aSecsItem.ItemName = Nothing
                    Else
                        aSecsItem.ItemName = RootItem.Attribute("Name").Value
                    End If

                    ' 設定 Item Description
                    If RootItem.Attribute("Description") Is Nothing Then
                        aSecsItem.ItemDescription = Nothing
                    Else
                        aSecsItem.ItemDescription = RootItem.Attribute("Description").Value
                    End If

                    ' 讀取下一個 Item
                    Dim temp As XElement = RootItem.FirstNode

                    For i As Integer = 1 To aSecsItem.ItemNumber Step +1

                        aSecsItem.ItemList.Add(LoadItem(temp.DescendantNodesAndSelf().ToList()))

                        If i <> aSecsItem.ItemNumber Then
                            temp = temp.NextNode
                        End If
                    Next

                Else
                    ' ------- Item Type != List -------

                    ' 設定 Item Value
                    If RootItem.Element("Value") IsNot Nothing Then

                        If RootItem.Element("Value").Value <> "" Then

                            Dim splitChar As Char() = New Char() {" "c, ","c}
                            Dim temp As String() = RootItem.Element("Value").Value.Split(splitChar)
                            aSecsItem.SetItemValue(temp)

                        End If
                    End If

                    ' 設定 Item Name
                    If RootItem.Attribute("Name") Is Nothing Then
                        aSecsItem.ItemName = Nothing
                    Else
                        aSecsItem.ItemName = RootItem.Attribute("Name").Value
                    End If

                    ' Item Description
                    If RootItem.Attribute("Description") Is Nothing Then
                        aSecsItem.ItemDescription = Nothing
                    Else
                        aSecsItem.ItemDescription = RootItem.Attribute("Description").Value
                    End If

                    ' Item MapsType
                    If RootItem.Element("Maps") IsNot Nothing Then

                        If RootItem.Element("Maps").Element("Map") IsNot Nothing Then

                            Dim temp As XElement = RootItem.Element("Maps").FirstNode

                            For i As Integer = 1 To RootItem.Element("Maps").Elements.Count Step +1

                                aSecsItem.ItemMapsType.Add(temp.Value)

                                If i <> RootItem.Element("Maps").Elements.Count Then
                                    temp = temp.NextNode
                                End If

                            Next
                        End If
                    End If
                End If

                Return aSecsItem

			Catch ex As Exception

                Return Nothing

            End Try

		End Function


        ' 讀取 Item 後，加到 ItemMap
        Private Sub LoadItemMap(ByRef aSecsItem As SecsItem, ByRef aItemMap As Dictionary(Of String, SecsItem))

            Try
                If aSecsItem.ItemType = "L" Then

                    ' ------------------ 假如 ItemType = List ------------------------

                    If aSecsItem.ItemName IsNot Nothing Then
                        aItemMap.Add(aSecsItem.ItemName, aSecsItem)
                    End If

                    For i As Integer = 0 To aSecsItem.ItemNumber - 1 Step +1
                        LoadItemMap(aSecsItem.ItemList(i), aItemMap)
                    Next
                Else

                    ' ------------------ 假如 ItemType != List -------------------------

                    If aSecsItem.ItemName IsNot Nothing Then
                        aItemMap.Add(aSecsItem.ItemName, aSecsItem)
                    End If
                End If

            Catch ex As Exception

                aItemMap.Clear()

			End Try
		End Sub


        ' 當收到 Message 時，使用此方法找出 Message
        Public Function FindMessageByFormat(ByRef aSecsMessage As SecsMessage) As Boolean

            Try
                ' 比對後符合格式的的 SecsMessage，放到 List
                Dim tempMessageList As List(Of SecsMessage) = New List(Of SecsMessage)

                ' 比對 Message 格式、SecsItem 格式
                For i As Integer = 0 To Me.sMessageList.Count - 1 Step +1

                    If aSecsMessage.CheckMessageFormat() = Me.sMessageList(i).CheckMessageFormat() Then

                        If aSecsMessage.secsItem Is Nothing Then

                            ' 符合格式的 SecsMessage 加到 tempMessageList
                            tempMessageList.Add(Me.sMessageList(i))

                        Else
                            ' 比對 ItemMapsType
                            If FindItemByMapsType(aSecsMessage.secsItem, Me.sMessageList(i).secsItem) = True Then

                                ' 符合格式的 SecsMessage 加到 tempMessageList
                                tempMessageList.Add(Me.sMessageList(i))
                            End If

                        End If

                    End If
                Next

                If tempMessageList.Count = 0 Then

                    ' 沒有符合格式的 Message

                    Return False

                ElseIf tempMessageList.Count = 1 Then

                    ' 有一個符合格式的 Message

                    aSecsMessage.MessageName = tempMessageList(0).MessageName
                    aSecsMessage.MessageDescription = tempMessageList(0).MessageDescription
                    aSecsMessage.AutoReply = tempMessageList(0).AutoReply

                    FindItemByFormat(aSecsMessage.secsItem, tempMessageList(0).secsItem)
                    Return True

                Else

                    ' 有兩個以上符合格式的 Message

                    aSecsMessage.MessageName = tempMessageList(0).MessageName
                    aSecsMessage.MessageDescription = tempMessageList(0).MessageDescription
                    aSecsMessage.AutoReply = tempMessageList(0).AutoReply

                    FindItemByFormat(aSecsMessage.secsItem, tempMessageList(0).secsItem)
                    Return True

                End If
                Return False

            Catch ex As Exception

                Return False
			End Try

		End Function


		' 當收到 Message 時，使用此方法找出 Item
		Private Sub FindItemByFormat(ByRef aSecsItem As SecsItem, ByRef aSXMLItem As SecsItem)

            Try
                If aSecsItem.ItemList.Count > 0 Then

                    ' ItemList 的數量 > 0

                    aSecsItem.ItemName = aSXMLItem.ItemName
                    aSecsItem.ItemDescription = aSXMLItem.ItemDescription
                    aSecsItem.ItemMapsType = aSXMLItem.ItemMapsType

                    ' 取出 ItemList 中的 Item 
                    For i As Integer = 0 To aSecsItem.ItemList.Count - 1

                        FindItemByFormat(aSecsItem.ItemList(i), aSXMLItem.ItemList(i))

                    Next

                Else
                    ' ItemList 的數量 <= 0

                    aSecsItem.ItemName = aSXMLItem.ItemName
                    aSecsItem.ItemDescription = aSXMLItem.ItemDescription
                    aSecsItem.ItemMapsType = aSXMLItem.ItemMapsType
                End If

            Catch ex As Exception

            End Try

		End Sub


		' 當有兩個以上的 SXML Message 有相同格式時，使用 Item MapType 來找出 Item
		Private Function FindItemByMapsType(ByRef aSecsItem As SecsItem, ByRef aSXMLItem As SecsItem) As Boolean

            Try
                If aSXMLItem.ItemList.Count > 0 Then

                    For i As Integer = 0 To aSXMLItem.ItemList.Count - 1 Step +1

                        If aSXMLItem.ItemList(i).ItemType = "L" Then

                            ' 假如 ItemType = List

                            FindItemByMapsType(aSecsItem.ItemList(i), aSXMLItem.ItemList(i))

                        Else

                            ' 假如 ItemType 不為 List

                            ' 假如 ItemMapsType 的數量 > 0
                            If aSXMLItem.ItemList(i).ItemMapsType.Count > 0 Then

                                For j As Integer = 0 To aSXMLItem.ItemList(i).ItemMapsType.Count - 1 Step +1

                                    Dim tempItemValue As String = Nothing

                                    For k As Integer = 1 To aSecsItem.ItemList(i).ItemValue.Count Step +1

                                        tempItemValue = tempItemValue + aSecsItem.ItemList(i).ItemValue(k).ToString
                                    Next

                                    If tempItemValue = aSXMLItem.ItemList(i).ItemMapsType(j) Then

                                        Return True

                                    End If

                                Next
                                Return False

                            End If

                        End If

                    Next
                    Return True

                End If
                Return True

            Catch ex As Exception

                Return False
			End Try

		End Function


		' 使用 Message Name 找出 SecsMessage 
		Public Function FindMessageByMessageName(ByVal aMessageName As String) As SecsMessage

            Try
                If sMessageMap(aMessageName) IsNot Nothing Then

                    Return sMessageMap(aMessageName).Clone()
                End If
                Return Nothing

            Catch ex As Exception

                Return Nothing
			End Try

		End Function

	End Class

End Namespace

