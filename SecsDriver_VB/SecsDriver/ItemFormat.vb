Imports System

Namespace SecsDriver

    ''' <summary>
    ''' Item Format
    ''' </summary>
    Friend Class ItemFormat : Implements ICloneable


        ''' <summary>
        ''' Lock 物件
        ''' </summary>
        Private Shared thisLock As Object = New Object


        ' -------------------- Property ------------------------

        ' 存取 Format Byte - 1 Byte
        Public Property FormatByte() As Byte

        ' 存取 Length Bytes - 1 ~ 3 Bytes
        Public Property LengthBytes As Byte()

        ' 存取 Data Bytes - ? Bytes
        Public Property DataBytes As Byte()


        ' 讀取 Item Format Code
        Public ReadOnly Property ItemFormatCode As Byte

            ' FormatByte & 1111 1100 (因為 Item Format Code 占前面六個Bits)
            Get
                Return (FormatByte And &HFC)
            End Get

        End Property


        ' 讀取 Number of Length Bytes
        Public ReadOnly Property NumOfLengthByte As Byte

            ' FormatByte & 0000 0011 (因為 No of Length Bytes 占最後兩個Bits)
            Get
                Return (FormatByte And &H3)
            End Get

        End Property


        ' -------------------- Method ------------------------

        ' 建構子
        Public Sub New()

            FormatByte = Nothing                   ' SECS-II Item Header - Format Byte
            LengthBytes = Nothing                  ' SECS-II Item Header - Length Bytes	
            DataBytes = Nothing                    ' SECS-II Item Header - Data Byte

        End Sub


        ' 設定 Item Format Code
        Public Sub SetItemFormatCode(ByVal aItFormat As Byte)

            ' SECS-II Item Header - Format Byte = Item Format Code + Number of Length Bytes
            FormatByte = (aItFormat Or Me.NumOfLengthByte)
        End Sub


        ' 設定 Number of Length Bytes
        Public Sub SetNumOfLengthByte(ByVal aItNumOfLen As Byte)

            ' SECS-II Item Header - Format Byte = Item Format Code + No.of Length Bytes
            FormatByte = (Me.ItemFormatCode Or aItNumOfLen)

        End Sub


        ' Clone
        Public Function Clone() As Object Implements ICloneable.Clone

            SyncLock thisLock

                Dim itemFormat As ItemFormat = New ItemFormat

                Try
                    itemFormat.FormatByte = Me.FormatByte

                    itemFormat.LengthBytes = New Byte() {}
                    itemFormat.LengthBytes = Me.LengthBytes

                    itemFormat.DataBytes = New Byte() {}
                    itemFormat.DataBytes = Me.DataBytes

                    Return itemFormat

                Catch ex As Exception

                    itemFormat.Finalize()
                    Return Nothing

                End Try

            End SyncLock

        End Function


	End Class

End Namespace

