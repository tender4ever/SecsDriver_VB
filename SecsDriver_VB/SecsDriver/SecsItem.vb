Imports System
Imports System.Collections.Generic
Imports System.Text

Namespace SecsDriver

    Public Class SecsItem : Implements ICloneable


#Region "Private 屬性"

        Private sItemType As String                                     ' Item Type
        Private sItemNumber As Int32                                    ' Item Number
        Private sItemValue As Collection                                ' Item Value

        Private sItemFormat As ItemFormat                               ' Item Format

#End Region


        ' Lock 物件
        Private Shared thisLock As Object = New Object


#Region "Public 屬性"

        ''' <summary>
        ''' ItemList
        ''' </summary>
        ''' <returns></returns>
        Public Property ItemList As List(Of SecsItem)

        ''' <summary>
        ''' ItemMap
        ''' </summary>
        ''' <returns></returns>
        Public Property ItemMap As Dictionary(Of String, SecsItem)


        ''' <summary>
        ''' ItemType
        ''' </summary>
        ''' <returns></returns>
        Public Property ItemType As String

            Get

                Select Case CType(Me.sItemFormat.ItemFormatCode, enumItemFormatCode)

                    Case enumItemFormatCode.sASCII
                        sItemType = "A"
                        Exit Select

                    Case enumItemFormatCode.sBinary
                        sItemType = "B"
                        Exit Select

                    Case enumItemFormatCode.sBoolean
                        sItemType = "Boolean"
                        Exit Select

                    Case enumItemFormatCode.sFloat4
                        sItemType = "F4"
                        Exit Select

                    Case enumItemFormatCode.sFloat8
                        sItemType = "F8"
                        Exit Select

                    Case enumItemFormatCode.sINT1
                        sItemType = "I1"
                        Exit Select

                    Case enumItemFormatCode.sINT2
                        sItemType = "I2"
                        Exit Select

                    Case enumItemFormatCode.sINT4
                        sItemType = "I4"
                        Exit Select

                    Case enumItemFormatCode.sINT8
                        sItemType = "I8"
                        Exit Select

                    Case enumItemFormatCode.sJIS8
                        sItemType = "J"
                        Exit Select

                    Case enumItemFormatCode.sList
                        sItemType = "L"
                        Exit Select

                    Case enumItemFormatCode.sTwoByteChar
                        sItemType = ""
                        Exit Select

                    Case enumItemFormatCode.sUINT1
                        sItemType = "U1"
                        Exit Select

                    Case enumItemFormatCode.sUINT2
                        sItemType = "U2"
                        Exit Select

                    Case enumItemFormatCode.sUINT4
                        sItemType = "U4"
                        Exit Select

                    Case enumItemFormatCode.sUINT8
                        sItemType = "U8"
                        Exit Select

                End Select

                Return sItemType

            End Get

            Set(value As String)

                Select Case value

                    Case "A"
                        sItemType = "A"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sASCII))
                        Exit Select

                    Case "B"
                        sItemType = "B"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sBinary))
                        Exit Select

                    Case "Boolean"
                        sItemType = "Boolean"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sBoolean))
                        Exit Select

                    Case "F4"
                        sItemType = "F4"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sFloat4))
                        Exit Select

                    Case "F8"
                        sItemType = "F8"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sFloat8))
                        Exit Select

                    Case "I1"
                        sItemType = "I1"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sINT1))
                        Exit Select

                    Case "I2"
                        sItemType = "I2"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sINT2))
                        Exit Select

                    Case "I4"
                        sItemType = "I4"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sINT4))
                        Exit Select

                    Case "I8"
                        sItemType = "I8"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sINT8))
                        Exit Select

                    Case "J"
                        sItemType = "J"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sJIS8))
                        Exit Select

                    Case "L"
                        sItemType = "L"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sList))
                        Exit Select

                    Case ""
                        sItemType = ""
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sTwoByteChar))
                        Exit Select

                    Case "U1"
                        sItemType = "U1"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sUINT1))
                        Exit Select

                    Case "U2"
                        sItemType = "U2"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sUINT2))
                        Exit Select

                    Case "U4"
                        sItemType = "U4"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sUINT4))
                        Exit Select

                    Case "U8"
                        sItemType = "U8"
                        sItemFormat.SetItemFormatCode(Convert.ToByte(enumItemFormatCode.sUINT8))
                        Exit Select

                End Select

            End Set

        End Property


        ''' <summary>
        ''' ItemNumber
        ''' </summary>
        ''' <returns></returns>
        Public Property ItemNumber As Int32

            Get

                If sItemFormat.LengthBytes Is Nothing Then

                    ' ------------ 假如 LengthBytes = Nothing --------------

                    ItemNumber = 0
                Else

                    ' ------------ 假如 LengthBytes != Nothing --------------

                    ' 計算出 DataBytes 的 Bytes 數量
                    Dim temp As Int64 = 0
                    If sItemFormat.LengthBytes.Length = 1 Then

                        temp = Convert.ToInt32(sItemFormat.LengthBytes(0))

                    ElseIf sItemFormat.LengthBytes.Length = 2 Then

                        temp = Convert.ToInt32(sItemFormat.LengthBytes(0)) * 255 + Convert.ToInt32(sItemFormat.LengthBytes(1))

                    ElseIf sItemFormat.LengthBytes.Length = 2 Then

                        temp = Convert.ToInt32(sItemFormat.LengthBytes(0)) * 255 * 255 + Convert.ToInt32(sItemFormat.LengthBytes(1)) * 255 + Convert.ToInt32(sItemFormat.LengthBytes(2))

                    End If

                    ' 透過 DataBytes 的 Bytes 數量，計算出 ItemNumber
                    If ItemType = "I2" Or ItemType = "U2" Then

                        sItemNumber = temp / 2

                    ElseIf ItemType = "I4" Or ItemType = "U4" Or ItemType = "F4" Then

                        sItemNumber = temp / 4

                    ElseIf ItemType = "I8" Or ItemType = "U8" Or ItemType = "F8" Then

                        sItemNumber = temp / 8

                    Else
                        sItemNumber = temp / 1

                    End If

                End If

                Return sItemNumber

            End Get

            Set(value As Int32)

                sItemNumber = value

                ' 計算出 DataBytes 的 Bytes 數量
                Dim temp As Int64

                If ItemType = "I2" Or ItemType = "U2" Then

                    temp = sItemNumber * 2

                ElseIf ItemType = "I4" Or ItemType = "U4" Or ItemType = "F4" Then

                    temp = sItemNumber * 4

                ElseIf ItemType = "I8" Or ItemType = "U8" Or ItemType = "F8" Then

                    temp = sItemNumber * 8

                Else
                    temp = sItemNumber * 1

                End If

                ' 透過 DataBytes 的 Bytes 數量，來設定 NumOfLengthByte、LengthBytes
                If 0 <= temp < 256 Then

                    ' 設定 NumOfLengthByte
                    sItemFormat.SetNumOfLengthByte(Convert.ToByte(1))

                    ' 設定 LengthBytes
                    Dim tempLengthBytes(0) As Byte
                    tempLengthBytes(0) = Convert.ToByte(temp)
                    sItemFormat.LengthBytes = tempLengthBytes

                ElseIf 255 < temp < 65536 Then

                    ' 設定 NumOfLengthByte
                    sItemFormat.SetNumOfLengthByte(Convert.ToByte(2))

                    ' 設定 LengthBytes
                    Dim tempLengthBytes(1) As Byte
                    tempLengthBytes(0) = Convert.ToByte(temp / 256)
                    tempLengthBytes(1) = Convert.ToByte(temp Mod 256)
                    sItemFormat.LengthBytes = tempLengthBytes

                ElseIf 65535 < temp < 16777216 Then

                    ' 設定 NumOfLengthByte
                    sItemFormat.SetNumOfLengthByte(Convert.ToByte(3))

                    ' 設定 LengthBytes
                    Dim tempLengthBytes(2) As Byte
                    tempLengthBytes(0) = Convert.ToByte(temp / 65536)
                    tempLengthBytes(1) = Convert.ToByte(temp Mod 65536 / 256)
                    tempLengthBytes(2) = Convert.ToByte(temp Mod 256)
                    sItemFormat.LengthBytes = tempLengthBytes

                End If

            End Set

        End Property


        ''' <summary>
        ''' ItemValue
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ItemValue As Collection

            Get
                If sItemFormat.DataBytes Is Nothing Then

                    ' ---------------------- 假如 DataBytes = Nothing -------------------------

                    sItemValue = Nothing

                Else

                    ' ---------------------- 假如 DataBytes = Nothing -------------------------

                    Select Case ItemType

                        Case "A"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            Dim temp As String = Encoding.ASCII.GetString(sItemFormat.DataBytes)
                            sItemValue.Add(temp)

                        Case "B"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            For i As Integer = 0 To ItemNumber - 1 Step +1
                                Dim temp(0) As Byte
                                Array.Copy(sItemFormat.DataBytes, i, temp, 0, 1)
                                sItemValue.Add(temp)
                            Next

                        Case "Boolean"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            For i As Integer = 0 To ItemNumber - 1 Step +1
                                Dim temp(0) As Byte
                                Array.Copy(sItemFormat.DataBytes, i, temp, 0, 1)
                                sItemValue.Add(BitConverter.ToBoolean(temp, 0))
                            Next

                        Case "F4"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            For i As Integer = 0 To ItemNumber - 1 Step +1
                                Dim temp(3) As Byte
                                Array.Copy(sItemFormat.DataBytes, i * 4, temp, 0, 4)
                                Array.Reverse(temp)
                                sItemValue.Add(BitConverter.ToSingle(temp, 0))
                            Next

                        Case "F8"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            For i As Integer = 0 To ItemNumber - 1 Step +1
                                Dim temp(7) As Byte
                                Array.Copy(sItemFormat.DataBytes, i * 8, temp, 0, 8)
                                Array.Reverse(temp)
                                sItemValue.Add(BitConverter.ToDouble(temp, 0))
                            Next

                        Case "I1"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            For i As Integer = 0 To ItemNumber - 1 Step +1
                                Dim temp(1) As Byte
                                Array.Copy(sItemFormat.DataBytes, i, temp, 0, 1)
                                sItemValue.Add(BitConverter.ToUInt16(temp, 0))
                            Next

                        Case "I2"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            For i As Integer = 0 To ItemNumber - 1 Step +1
                                Dim temp(1) As Byte
                                Array.Copy(sItemFormat.DataBytes, i * 2, temp, 0, 2)
                                Array.Reverse(temp)
                                sItemValue.Add(BitConverter.ToInt16(temp, 0))
                            Next

                        Case "I4"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            For i As Integer = 0 To ItemNumber - 1 Step +1
                                Dim temp(3) As Byte
                                Array.Copy(sItemFormat.DataBytes, i * 4, temp, 0, 4)
                                Array.Reverse(temp)
                                sItemValue.Add(BitConverter.ToInt32(temp, 0))
                            Next

                        Case "I8"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            For i As Integer = 0 To ItemNumber - 1 Step +1
                                Dim temp(7) As Byte
                                Array.Copy(sItemFormat.DataBytes, i * 8, temp, 0, 8)
                                Array.Reverse(temp)
                                sItemValue.Add(BitConverter.ToInt64(temp, 0))
                            Next

                        Case "J"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            For i As Integer = 0 To ItemNumber - 1 Step +1
                                Dim temp(0) As Byte
                                Array.Copy(sItemFormat.DataBytes, i, temp, 0, 1)
                                sItemValue.Add(BitConverter.ToString(temp, 0))
                            Next

                        Case "L"

                            ' 清除 sItemValue
                            sItemValue = Nothing

                        Case "U1"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            For i As Integer = 0 To ItemNumber - 1 Step +1
                                Dim temp(1) As Byte
                                Array.Copy(sItemFormat.DataBytes, i, temp, 0, 1)
                                sItemValue.Add(BitConverter.ToUInt16(temp, 0))
                            Next

                        Case "U2"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            For i As Integer = 0 To ItemNumber - 1 Step +1
                                Dim temp(1) As Byte
                                Array.Copy(sItemFormat.DataBytes, i * 2, temp, 0, 2)
                                Array.Reverse(temp)
                                sItemValue.Add(BitConverter.ToUInt16(temp, 0))
                            Next

                        Case "U4"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            For i As Integer = 0 To ItemNumber - 1 Step +1
                                Dim temp(3) As Byte
                                Array.Copy(sItemFormat.DataBytes, i * 4, temp, 0, 4)
                                Array.Reverse(temp)
                                sItemValue.Add(BitConverter.ToUInt32(temp, 0))
                            Next

                        Case "U8"

                            ' 清除 sItemValue
                            sItemValue.Clear()

                            ' 設定 sItemValue
                            For i As Integer = 0 To ItemNumber - 1 Step +1
                                Dim temp(7) As Byte
                                Array.Copy(sItemFormat.DataBytes, i * 8, temp, 0, 8)
                                Array.Reverse(temp)
                                sItemValue.Add(BitConverter.ToUInt64(temp, 0))
                            Next

                    End Select

                End If

                Return sItemValue

            End Get

        End Property


        ''' <summary>
        ''' ItemName
        ''' </summary>
        ''' <returns></returns>
        Public Property ItemName As String


        ''' <summary>
        ''' ItemDescription
        ''' </summary>
        ''' <returns></returns>
        Public Property ItemDescription As String


        ''' <summary>
        ''' ItemMapsType
        ''' </summary>
        ''' <returns></returns>
        Public Property ItemMapsType As List(Of String)

#End Region


        ' ----------------------- Method ------------------------------

        ' 建構子，新增 未設值的 SecsItem
        Public Sub New()

            ItemList = New List(Of SecsItem)
            ItemMap = New Dictionary(Of String, SecsItem)
            sItemValue = New Collection
            ItemMapsType = New List(Of String)
            sItemFormat = New ItemFormat()

        End Sub


        ' 建構子，收到 Bytes 時，使用此方法來新增 SecsItem
        Public Sub New(ByRef aBytes() As Byte)

            ItemList = New List(Of SecsItem)
            ItemMap = New Dictionary(Of String, SecsItem)
            sItemValue = New Collection
            ItemMapsType = New List(Of String)
            sItemFormat = New ItemFormat()

            Try
                Dim i As Int64 = 0

                ' 設定 FormatByte
                sItemFormat.FormatByte = aBytes(i)

                ' 設定 LengthBytes
                Dim tempLengthBytes(Convert.ToInt32(sItemFormat.NumOfLengthByte) - 1) As Byte
                For j As Integer = 0 To tempLengthBytes.Length - 1 Step +1

                    i = i + 1
                    tempLengthBytes(j) = aBytes(i)
                Next
                sItemFormat.LengthBytes = tempLengthBytes

                ' 設定 DataBytes
                If ItemType = "L" Then

                    ' ------------------ 假如 RootItem = List -----------------------

                    i = i + 1

                    ' 解析 Bytes 並轉成 SecsItem，然後將解出的 SecsItem 加到 ItemList
                    i = i + ParseBytesToSecsItem(aBytes.Skip(i).ToArray, Me, ItemNumber) + 1

                Else

                    ' ------------------ 假如 RootItem != List ---------------------

                    If ItemNumber = 0 Then

                        sItemFormat.DataBytes = Nothing
                    Else
                        ' 計算 Item 的 DataBytes 的 Bytes Length
                        Dim tempDataBytes(ExecuteItemBytesLength(ItemType, ItemNumber) - 1) As Byte

                        For k As Int64 = 0 To tempDataBytes.Count - 1 Step +1

                            i = i + 1
                            tempDataBytes(k) = aBytes(i)
                        Next
                        sItemFormat.DataBytes = tempDataBytes
                    End If

                End If
            Catch ex As Exception

                ItemList.Clear()
                ItemMap.Clear()
                sItemValue.Clear()
                ItemMapsType.Clear()
                sItemFormat = Nothing

            End Try

        End Sub


        ' 解析收到的 Bytes 並轉換成 SecsItem
        Private Function ParseBytesToSecsItem(ByRef aBytes() As Byte, ByRef RootSecsItem As SecsItem, ByRef aItemNumber As Integer) As Int64

            Dim i As Int64 = 0

            For j As Integer = 0 To aItemNumber - 1 Step +1

                ' New SecsItem
                Dim sSecsitem As New SecsItem()

                Try
                    ' 設定 FormatByte
                    sSecsitem.sItemFormat.FormatByte = aBytes(i)

                    ' 設定 LengthBytes
                    Dim tempLengthBytes(Convert.ToInt32(sSecsitem.sItemFormat.NumOfLengthByte) - 1) As Byte
                    For k As Integer = 0 To tempLengthBytes.Length - 1 Step +1

                        i = i + 1
                        tempLengthBytes(k) = aBytes(i)
                    Next
                    sSecsitem.sItemFormat.LengthBytes = tempLengthBytes

                    ' 設定 DataBytes
                    If sSecsitem.ItemType = "L" Then

                        ' ------------------------ 假如 Item = List ---------------------------------

                        i = i + 1

                        ' 解析 Bytes 並轉成 SecsItem
                        i = i + ParseBytesToSecsItem(aBytes.Skip(i).ToArray, sSecsitem, sSecsitem.ItemNumber)

                    Else

                        ' ------------------------ 假如 Item != List ---------------------------------

                        If sSecsitem.ItemNumber = 0 Then
                            sSecsitem.sItemFormat.DataBytes = Nothing
                        Else
                            ' 計算 Item 的 DataBytes 的 Bytes Length
                            Dim tempDataBytes(ExecuteItemBytesLength(sSecsitem.ItemType, sSecsitem.ItemNumber) - 1) As Byte

                            For k As Int64 = 0 To tempDataBytes.Count - 1 Step +1

                                i = i + 1
                                tempDataBytes(k) = aBytes(i)
                            Next

                            sSecsitem.sItemFormat.DataBytes = tempDataBytes

                        End If

                        i = i + 1

                    End If

                    RootSecsItem.ItemList.Add(sSecsitem)

                Catch ex As Exception

                    sSecsitem.ItemList.Clear()
                    sSecsitem.ItemMap.Clear()
                    sSecsitem.sItemValue.Clear()
                    sSecsitem.ItemMapsType.Clear()
                    sSecsitem.sItemFormat = Nothing

                End Try

            Next

            Return i

        End Function


        ' 計算 Item 的 DataBytes 的 Bytes Length 
        Private Function ExecuteItemBytesLength(ByRef aItemType As String, ByVal aItemNumber As Integer) As Int64

            Dim temp As Int64 = Nothing

            If aItemType = "I2" Or aItemType = "U2" Then

                temp = aItemNumber * 2

            ElseIf aItemType = "I4" Or aItemType = "U4" Or aItemType = "F4" Then

                temp = aItemNumber * 4

            ElseIf aItemType = "I8" Or aItemType = "U8" Or aItemType = "F8" Then

                temp = aItemNumber * 8

            Else
                temp = aItemNumber * 1

            End If

            Return temp

        End Function


        ' 將 SecsItem 轉換成 Bytes
        Public Function ItemConvertToBytes() As Byte()
            Try
                If ItemType = "L" Or ItemNumber = 0 Then

                    ' --------------------------------- 假如 Item = List 或 ItemNumber = 0 -----------------------------

                    If ItemNumber = 0 Then
                        ItemNumber = 0
                    End If

                    Dim tempBytes(Me.sItemFormat.LengthBytes.Length) As Byte

                    ' 設定 FormatByte
                    tempBytes(0) = Me.sItemFormat.FormatByte

                    ' 設定 LengthBytes
                    For i As Integer = 0 To Me.sItemFormat.LengthBytes.Length - 1 Step +1

                        tempBytes(i + 1) = Me.sItemFormat.LengthBytes(i)
                    Next
                    Return tempBytes

                ElseIf ItemNumber <> 0 And ItemType <> "L" Then

                    ' -------------------------------- 假如Item != List 且 ItemNumber != 0 ----------------------------------

                    Dim tempBytes(Me.sItemFormat.LengthBytes.Length + Me.sItemFormat.DataBytes.Length) As Byte

                    ' 設定 FormatByte
                    tempBytes(0) = Me.sItemFormat.FormatByte

                    ' 設定 LengthBytes
                    For i As Integer = 0 To Me.sItemFormat.LengthBytes.Length - 1 Step +1

                        tempBytes(i + 1) = Me.sItemFormat.LengthBytes(i)
                    Next

                    ' 設定 DataBytes
                    For i As Integer = 0 To Me.sItemFormat.DataBytes.Length - 1 Step +1

                        tempBytes(i + Me.sItemFormat.LengthBytes.Length + 1) = Me.sItemFormat.DataBytes(i)
                    Next
                    Return tempBytes

                End If

                Return Nothing

            Catch ex As Exception

                Return Nothing

            End Try

        End Function


        ' 將 SecsItem 轉換成 BinaryLog
        Public Function ConvertToBinaryLog() As String

            Dim temp As String = ""

            If Me.ItemType = "L" Then

                ' ------------------------------- 當 item 為 List 時 -----------------------------------

                ' FormatByte
                temp += Me.sItemFormat.FormatByte.ToString("X2") & " "

                ' LengthBytes
                For i As Integer = 0 To Me.sItemFormat.LengthBytes.Length - 1 Step +1

                    temp += Me.sItemFormat.LengthBytes(i).ToString("X2") & " "
                Next

                ' 解 ItemList
                For i As Integer = 0 To Me.ItemList.Count - 1

                    temp += Me.ItemList(i).ConvertToBinaryLog()
                Next

            Else

                ' -------------------------------- 當 item 不為 List 時 ---------------------------------

                ' FormatByte
                temp += Me.sItemFormat.FormatByte.ToString("X2") & " "

                ' LengthBytes
                If Me.sItemFormat.LengthBytes IsNot Nothing Then


                    For i As Integer = 0 To Me.sItemFormat.LengthBytes.Length - 1 Step +1

                        temp += Me.sItemFormat.LengthBytes(i).ToString("X2") & " "
                    Next

                End If


                ' DataBytes
                If Me.sItemFormat.DataBytes IsNot Nothing Then

                    For i As Integer = 0 To Me.sItemFormat.DataBytes.Length - 1 Step +1

                        temp += Me.sItemFormat.DataBytes(i).ToString("X2") & " "
                    Next

                End If

            End If

            Return temp

        End Function


        ' 使用 ItemName 取得 ItemMap 中的 SecsItem
        Public Function GetItemByName(ByRef ItemName As String) As SecsItem

            If Me.ItemMap.Count <> 0 Then

                Try
                    Dim item As SecsItem = Me.ItemMap(ItemName)

                    If item IsNot Nothing Then
                        Return item
                    End If

                Catch ex As Exception

                    Return Nothing
                End Try

            Else
                Return Nothing
            End If

        End Function


        ' 將 ItemValue 轉換成 String
        Public Function ItemValueString() As String

            Dim itemString As String = Nothing

            If Me.ItemValue IsNot Nothing Then

                For i As Integer = 1 To Me.ItemValue.Count

                    If Me.ItemType = "B" Then

                        Dim tempByte As Byte()

                        If i = Me.ItemValue.Count Then
                            tempByte = Me.ItemValue(i)
                            itemString = itemString & tempByte(0).ToString
                        Else
                            tempByte = Me.ItemValue(i)
                            itemString = itemString & tempByte(0).ToString + " "

                        End If

                    Else

                        If i = Me.ItemValue.Count Then

                            itemString = itemString + Convert.ToString(Me.ItemValue(i))
                        Else
                            itemString = itemString + Convert.ToString(Me.ItemValue(i)) + " "
                        End If

                    End If

                Next

                Return itemString

            End If

            Return Nothing

        End Function


        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aString() As String) As Boolean

            Try
                If sItemValue IsNot Nothing Then
                    ' 清除 sItemValue
                    sItemValue.Clear()
                Else
                    sItemValue = New Collection
                End If

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "A"

                        ' 設定 sItemValue
                        For i As Integer = 0 To aString.Count - 1 Step +1
                            sItemValue.Add(aString(i))
                        Next

                        ' 設定 DataBytes

                        sItemFormat.DataBytes = Encoding.ASCII.GetBytes(sItemValue(1))

                        ' 設定 ItemNumber
                        ItemNumber = sItemFormat.DataBytes.Length

                        Return True

                    Case "B", "I1", "U1"

                        ' 設定 sItemValue
                        For i As Integer = 0 To aString.Count - 1 Step +1
                            Dim temp As Byte
                            temp = Convert.ToByte(aString(i))
                            sItemValue.Add(temp)
                        Next

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count - 1) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(0) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Copy(temp, 0, insertByte, i - 1, 1)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                    Case "Boolean"

                        ' 設定 sItemValue
                        For i As Integer = 0 To aString.Count - 1 Step +1
                            Dim temp As Boolean
                            temp = Convert.ToBoolean(aString(i))
                            sItemValue.Add(temp)
                        Next

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count - 1) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(0) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Copy(temp, 0, insertByte, i - 1, 1)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                    Case "I2", "U2"

                        ' 設定 sItemValue
                        For i As Integer = 0 To aString.Count - 1 Step +1

                            If ItemType = "I2" Then
                                Dim temp As Int16
                                temp = Convert.ToInt16(aString(i))
                                sItemValue.Add(temp)

                            ElseIf ItemType = "U2" Then
                                Dim temp As UInt16
                                temp = Convert.ToUInt16(aString(i))
                                sItemValue.Add(temp)

                            End If
                        Next

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 2 - 1) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(1) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 2)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                    Case "F4", "I4", "U4"
                        ' 設定 sItemValue
                        For i As Integer = 0 To aString.Count - 1 Step +1

                            If ItemType = "F4" Then
                                Dim temp As Single
                                temp = Convert.ToSingle(aString(i))
                                sItemValue.Add(temp)

                            ElseIf ItemType = "I4" Then
                                Dim temp As Int32
                                temp = Convert.ToInt32(aString(i))
                                sItemValue.Add(temp)

                            ElseIf ItemType = "U4" Then
                                Dim temp As UInt32
                                temp = Convert.ToUInt32(aString(i))
                                sItemValue.Add(temp)
                            End If
                        Next

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 4) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(3) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 4)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                    Case "F8", "I8", "U8"
                        ' 設定 sItemValue
                        For i As Integer = 0 To aString.Count - 1 Step +1

                            If ItemType = "F8" Then
                                Dim temp As Double
                                temp = Convert.ToDouble(aString(i))
                                sItemValue.Add(temp)

                            ElseIf ItemType = "I8" Then
                                Dim temp As Int64
                                temp = Convert.ToInt64(aString(i))
                                sItemValue.Add(temp)

                            ElseIf ItemType = "U8" Then
                                Dim temp As UInt64
                                temp = Convert.ToUInt64(aString(i))
                                sItemValue.Add(temp)
                            End If
                        Next

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 8) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(7) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 8)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                    Case "L"

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aString As String) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                sItemValue.Add(aString)

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "A"

                        ' 設定 DataBytes
                        sItemFormat.DataBytes = Encoding.ASCII.GetBytes(aString)

                        ' 設定 ItemNumber
                        ItemNumber = sItemFormat.DataBytes.Length

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try

        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aBinary() As Byte) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                For i As Integer = 0 To aBinary.Count - 1 Step +1
                    sItemValue.Add(aBinary(i))
                Next

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "B"

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count - 1) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(0) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Copy(temp, 0, insertByte, i - 1, 1)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aBinary As Byte) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                sItemValue.Add(aBinary)

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "B"

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(0) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Copy(temp, 0, insertByte, i - 1, 1)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aBoolean() As Boolean) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                For i As Integer = 0 To aBoolean.Count - 1 Step +1
                    sItemValue.Add(aBoolean(i))
                Next

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "Boolean"

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(0) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Copy(temp, 0, insertByte, i - 1, 1)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aBoolean As Boolean) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                sItemValue.Add(aBoolean)

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "Boolean"

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(0) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Copy(temp, 0, insertByte, i - 1, 1)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aFloat() As Single) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                For i As Integer = 0 To aFloat.Count - 1 Step +1
                    sItemValue.Add(aFloat(i))
                Next

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "F4"

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 4) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(3) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 4)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aFloat As Single) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                sItemValue.Add(aFloat)

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "F4"

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 4) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(3) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 4)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aDouble() As Double) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                For i As Integer = 0 To aDouble.Count - 1 Step +1
                    sItemValue.Add(aDouble(i))
                Next

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "F8"
                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 8) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(7) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 8)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aDouble As Double) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                sItemValue.Add(aDouble)

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "F8"
                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 8) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(7) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 8)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aInt() As Int16) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                For i As Integer = 0 To aInt.Count - 1 Step +1
                    sItemValue.Add(aInt(i))
                Next

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "I1"

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(0) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Copy(temp, 0, insertByte, i - 1, 1)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                    Case "I2"
                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 2) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(1) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 2)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aInt As Int16) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                sItemValue.Add(aInt)

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "I1"

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(0) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Copy(temp, 0, insertByte, i - 1, 1)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                    Case "I2"

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 2) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(1) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 2)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aInt() As Int32) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                For i As Integer = 0 To aInt.Count - 1 Step +1
                    sItemValue.Add(aInt(i))
                Next

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "I4"

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 4) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(3) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 4)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aInt As Int32) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                sItemValue.Add(aInt)

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "I4"
                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 4) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(3) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 4)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aInt() As Int64) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                For i As Integer = 0 To aInt.Count - 1 Step +1
                    sItemValue.Add(aInt(i))
                Next

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "I8"
                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 8) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(7) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 8)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aInt As Int64) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                sItemValue.Add(aInt)

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "I8"
                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 8) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(7) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 8)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aUint() As UInt16) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                For i As Integer = 0 To aUint.Count - 1 Step +1
                    sItemValue.Add(aUint(i))
                Next

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "U1"
                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(0) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Copy(temp, 0, insertByte, i - 1, 1)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                    Case "U2"
                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 2) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(1) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 2)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aUint As UInt16) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                sItemValue.Add(aUint)

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "U1"
                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(0) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Copy(temp, 0, insertByte, i - 1, 1)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                    Case "U2"
                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 2) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(1) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 2)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aUint() As UInt32) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                For i As Integer = 0 To aUint.Count - 1 Step +1
                    sItemValue.Add(aUint(i))
                Next

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "U4"
                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 4) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(3) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 4)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aUint As UInt32) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                sItemValue.Add(aUint)

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "U4"
                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 4) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(3) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 4)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aUint() As UInt64) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                For i As Integer = 0 To aUint.Count - 1 Step +1
                    sItemValue.Add(aUint(i))
                Next

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "U8"

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 8) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(7) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 8)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function

        ' 設定 ItemValue
        Public Function SetItemValue(ByRef aUint As UInt64) As Boolean
            Try
                ' 清除 sItemValue
                sItemValue.Clear()

                ' 設定 sItemValue
                sItemValue.Add(aUint)

                ' 清除 DataBytes
                sItemFormat.DataBytes = Nothing

                Select Case ItemType

                    Case "U8"

                        ' 設定 DataBytes
                        Dim insertByte() As Byte = New Byte(sItemValue.Count * 8) {}
                        For i As Integer = 1 To sItemValue.Count Step +1
                            Dim temp(7) As Byte
                            temp = BitConverter.GetBytes(sItemValue(i))
                            Array.Reverse(temp)
                            temp.CopyTo(insertByte, (i - 1) * 8)
                        Next
                        sItemFormat.DataBytes = insertByte

                        ' 設定 ItemNumber
                        ItemNumber = sItemValue.Count

                        Return True

                End Select

                Return False

            Catch ex As Exception

                sItemValue.Clear()
                Return False

            End Try
        End Function



        ' 在 ItemList 後面加上新的 SecsItem
        Public Sub AddItem()

            If Me.ItemType <> "L" Then

                Me.ItemType = "L"
                Me.ItemNumber = 1
            Else
                Me.ItemNumber = Me.ItemNumber + 1
            End If

            Dim sSecsItem As New SecsItem
            Me.ItemList.Add(sSecsItem)

        End Sub


        ' 設定 ItemList 中 的 SecsItem
        Public Sub SetItemList(ByVal index As Integer, ByVal value As SecsItem)

            ItemList(index) = value

        End Sub


        ' 刪除 ItemList 中的 Item
        Public Sub DeleteItem(ByVal index As UInteger)

            Me.ItemNumber = Me.ItemNumber - 1
            Me.ItemList.RemoveAt(index)

        End Sub


        ' 在 ItemList 中插入新的 SecsItem
        Public Sub InsertItem(ByVal index As UInteger)

            If Me.ItemType <> "L" Then

                Me.ItemType = "L"

            End If
            Me.ItemNumber = Me.ItemNumber + 1
            Dim sSecsitem As New SecsItem
            Me.ItemList.Insert(index, sSecsitem)

        End Sub


        ' Clone 
        Public Function Clone() As Object Implements ICloneable.Clone

            SyncLock thisLock

                ' 宣告 SecsItem 物件
                Dim secsItem As SecsItem = New SecsItem

                Try
                    secsItem.sItemType = Me.sItemType
                    secsItem.sItemNumber = Me.sItemNumber

                    For Each item As Object In Me.sItemValue
                        secsItem.sItemValue.Add(item)
                    Next

                    secsItem.ItemName = Me.ItemName
                    secsItem.ItemDescription = Me.ItemDescription

                    For Each item As String In Me.ItemMapsType
                        secsItem.ItemMapsType.Add(item)
                    Next

                    For Each item As SecsItem In Me.ItemList
                        secsItem.ItemList.Add(item.Clone)
                    Next

                    secsItem.ItemMap = New Dictionary(Of String, SecsItem)(Me.ItemMap)

                    secsItem.sItemFormat = Me.sItemFormat.Clone()

                    Return secsItem

                Catch ex As Exception

                    secsItem.Finalize()
                    Return Nothing
                End Try

            End SyncLock

        End Function


    End Class

End Namespace

