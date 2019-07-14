Imports System
Imports System.Text
Imports System.Runtime.InteropServices

Namespace SecsDriver

	Public Class IniFile


#Region "Private 屬性"

        Private sIniFilePath As String              ' iniFile 檔案路徑
        Private sSection As String                  ' Section

#End Region


#Region "Public 屬性"

        ' IniFilePath
        Public ReadOnly Property IniFilePath
            Get
                Return sIniFilePath
            End Get
        End Property

        ' SXMLPath
        Public Property SXMLPath As String

        ' Entity
        Public Property Entity As enumSecsEntity

        ' Role
        Public Property Role As enumRole

        ' LinkType
        Public Property LinkType As String

        ' DeviceName
        Public Property DeviceName As String

        ' DeviceID
        Public Property DeviceID As UInteger

        ' Retry
        Public Property Retry As UInt16

        ' IP
        Public Property IP As String

        ' Port
        Public Property Port As UInt16

        ' T1 Timeout
        Public Property T1Timeout As Double

        ' T2 Timeout
        Public Property T2Timeout As Double

        ' T3 Timeout
        Public Property T3Timeout As Double

        ' T4 Timeout
        Public Property T4Timeout As Double

        ' T5 Timeout 
        Public Property T5Timeout As Double

        ' T6 Timeout 
        Public Property T6Timeout As Double

        ' T7 Timeout
        Public Property T7Timeout As Double

        ' T8 Timeout
        Public Property T8Timeout As Double

        ' Linktest Timeout
        Public Property CATimeout As Double

        ' Interleave Enabled
        Public Property InterleaveEnabled As Boolean

        ' Heartbeat Enabled
        Public Property HeartbeatEnabled As Boolean

        'Heartbeat Interval
        Public Property HeartbeatInterval As UInteger

        ' Heartbeat Tag
        Public Property HeartbeatTag As String

        ' AutoReply
        Public Property AutoReply As Boolean

        ' BinaryLog
        Public Property BinaryLog As Boolean

        ' TxLog
        Public Property TxLog As Boolean

#End Region


#Region "Public Method"

        ''' <summary>
        ''' 建構子
        ''' </summary>
        ''' <param name="aIniFilePath"></param>
        Protected Friend Sub New(ByVal aIniFilePath As String)

            sIniFilePath = aIniFilePath

            ' New TextHandler
            Dim aTextHandler As TextHandler.TextHandler
            aTextHandler = New TextHandler.TextHandler(sIniFilePath, "SECS")

            ' SXMLPath
            SXMLPath = aTextHandler.GetData("SXMLFile", "SecsWrapperX_AMHS.sxml")

            ' Entity
            Dim aEntity As String = aTextHandler.GetData("Entity", "Passive")

            If aEntity = "Passive" Then
                Entity = enumSecsEntity.sPassive
            Else
                Entity = enumSecsEntity.sActive
            End If

            ' Role
            Dim aRole As String = aTextHandler.GetData("Role", "Equipment")

            If aRole = "Equipment" Then
                Role = enumRole.sEquipment
            Else
                Role = enumRole.sHost
            End If

            ' LinkType
            LinkType = aTextHandler.GetData("LinkType", "HSMS")

            ' DeviceName
            DeviceName = aTextHandler.GetData("DeviceName", "HSMS")

            ' Device ID
            Dim aDeviceID As String = aTextHandler.GetData("DeviceID", "0x0000")
            DeviceID = Convert.ToUInt16(aDeviceID, 16)

            ' 讀取 Retry
            Dim aRetry As String = aTextHandler.GetData("Retry", "3")
            Retry = Convert.ToInt16(aRetry)

            ' 讀取 IP
            IP = aTextHandler.GetData("IP", "127.0.0.1")

            ' 讀取 Port
            Dim aPort As String = aTextHandler.GetData("Port", "5000")
            Port = Convert.ToInt32(aPort)

            ' 讀取 Timeout 設定值
            Dim aT1 As String = aTextHandler.GetData("T1", "1")
            T1Timeout = Convert.ToInt32(aT1)


            Dim aT2 As String = aTextHandler.GetData("T2", "10")
            T2Timeout = Convert.ToInt32(aT2)

            Dim aT3 As String = aTextHandler.GetData("T3", "45")
            T3Timeout = Convert.ToInt32(aT3)

            Dim aT4 As String = aTextHandler.GetData("T4", "10")
            T4Timeout = Convert.ToInt32(aT4)

            Dim aT5 As String = aTextHandler.GetData("T5", "10")
            T5Timeout = Convert.ToInt32(aT5)

            Dim aT6 As String = aTextHandler.GetData("T6", "5")
            T6Timeout = Convert.ToInt32(aT6)

            Dim aT7 As String = aTextHandler.GetData("T7", "10")
            T7Timeout = Convert.ToInt32(aT7)

            Dim aT8 As String = aTextHandler.GetData("T8", "5")
            T8Timeout = Convert.ToInt32(aT8)

            Dim aCA As String = aTextHandler.GetData("CA", "30")
            CATimeout = Convert.ToInt32(aCA)

            ' 讀取 Heartbeat 設定值
            Dim aInterleaveEnabled As String = aTextHandler.GetData("InterleaveEnabled", "True")

            If aInterleaveEnabled = "False" Then
                InterleaveEnabled = False
            Else
                InterleaveEnabled = True
            End If

            Dim aHeartbeatEnabled As String = aTextHandler.GetData("HeartbeatEnabled", "False")

            If aHeartbeatEnabled = "False" Then
                HeartbeatEnabled = False
            Else
                HeartbeatEnabled = True
            End If

            Dim aHeartbeatInterval As String = aTextHandler.GetData("HeartbeatInterval", "30")
            HeartbeatInterval = Convert.ToUInt32(aHeartbeatInterval)

            HeartbeatTag = aTextHandler.GetData("HeartbeatTag", " ")

            ' 讀取 AutoReply 設定值
            Dim aAutoReply As String = aTextHandler.GetData("AutoReply", "False")

            If aAutoReply = "False" Then
                AutoReply = False
            Else
                AutoReply = True
            End If

            ' 讀取 BinaryLog 設定值
            Dim aBinaryLog As String = aTextHandler.GetData("AutoReply", "False")

            If aBinaryLog = "False" Then
                BinaryLog = False
            Else
                BinaryLog = True
            End If

            ' 讀取 TxLog 設定值
            Dim aTxLog As String = aTextHandler.GetData("TxLog", "False")

            If aTxLog = "False" Then
                TxLog = False
            Else
                TxLog = True
            End If

        End Sub


        ''' <summary>
        ''' 重新讀取 Ini File
        ''' </summary>
        Protected Friend Sub ReLoadIniFile()

            ' New TextHandler
            Dim aTextHandler As TextHandler.TextHandler
            aTextHandler = New TextHandler.TextHandler(sIniFilePath, "SECS")

            ' SXMLPath
            SXMLPath = aTextHandler.GetData("SXMLFile", "SecsWrapperX_AMHS.sxml")

            ' Entity
            Dim aEntity As String = aTextHandler.GetData("Entity", "Passive")

            If aEntity = "Passive" Then
                Entity = enumSecsEntity.sPassive
            Else
                Entity = enumSecsEntity.sActive
            End If

            ' Role
            Dim aRole As String = aTextHandler.GetData("Role", "Equipment")

            If aRole = "Equipment" Then
                Role = enumRole.sEquipment
            Else
                Role = enumRole.sHost
            End If

            ' LinkType
            LinkType = aTextHandler.GetData("LinkType", "HSMS")

            ' DeviceName
            DeviceName = aTextHandler.GetData("DeviceName", "HSMS")

            ' Device ID
            Dim aDeviceID As String = aTextHandler.GetData("DeviceID", "0x0000")
            DeviceID = Convert.ToUInt16(aDeviceID, 16)

            ' 讀取 Retry
            Dim aRetry As String = aTextHandler.GetData("Retry", "3")
            Retry = Convert.ToInt16(aRetry)

            ' 讀取 IP
            IP = aTextHandler.GetData("IP", "127.0.0.1")

            ' 讀取 Port
            Dim aPort As String = aTextHandler.GetData("Port", "5000")
            Port = Convert.ToInt32(aPort)

            ' 讀取 Timeout 設定值
            Dim aT1 As String = aTextHandler.GetData("T1", "1")
            T1Timeout = Convert.ToInt32(aT1)


            Dim aT2 As String = aTextHandler.GetData("T2", "10")
            T2Timeout = Convert.ToInt32(aT2)

            Dim aT3 As String = aTextHandler.GetData("T3", "45")
            T3Timeout = Convert.ToInt32(aT3)

            Dim aT4 As String = aTextHandler.GetData("T4", "10")
            T4Timeout = Convert.ToInt32(aT4)

            Dim aT5 As String = aTextHandler.GetData("T5", "10")
            T5Timeout = Convert.ToInt32(aT5)

            Dim aT6 As String = aTextHandler.GetData("T6", "5")
            T6Timeout = Convert.ToInt32(aT6)

            Dim aT7 As String = aTextHandler.GetData("T7", "10")
            T7Timeout = Convert.ToInt32(aT7)

            Dim aT8 As String = aTextHandler.GetData("T8", "5")
            T8Timeout = Convert.ToInt32(aT8)

            Dim aCA As String = aTextHandler.GetData("CA", "30")
            CATimeout = Convert.ToInt32(aCA)

            ' 讀取 Heartbeat 設定值
            Dim aInterleaveEnabled As String = aTextHandler.GetData("InterleaveEnabled", "True")

            If aInterleaveEnabled = "False" Then
                InterleaveEnabled = False
            Else
                InterleaveEnabled = True
            End If

            Dim aHeartbeatEnabled As String = aTextHandler.GetData("HeartbeatEnabled", "False")

            If aHeartbeatEnabled = "False" Then
                HeartbeatEnabled = False
            Else
                HeartbeatEnabled = True
            End If

            Dim aHeartbeatInterval As String = aTextHandler.GetData("HeartbeatInterval", "30")
            HeartbeatInterval = Convert.ToUInt32(aHeartbeatInterval)

            HeartbeatTag = aTextHandler.GetData("HeartbeatTag", " ")

            ' 讀取 AutoReply 設定值
            Dim aAutoReply As String = aTextHandler.GetData("AutoReply", "False")

            If aAutoReply = "False" Then
                AutoReply = False
            Else
                AutoReply = True
            End If

            ' 讀取 BinaryLog 設定值
            Dim aBinaryLog As String = aTextHandler.GetData("AutoReply", "False")

            If aBinaryLog = "False" Then
                BinaryLog = False
            Else
                BinaryLog = True
            End If

            ' 讀取 TxLog 設定值
            Dim aTxLog As String = aTextHandler.GetData("TxLog", "False")

            If aTxLog = "False" Then
                TxLog = False
            Else
                TxLog = True
            End If

        End Sub


        ''' <summary>
        ''' 儲存 Ini File
        ''' </summary>
        Protected Friend Sub SaveIniFile()

            Dim buffer As String

            ' New TextHandler
            Dim aTextHandler As TextHandler.TextHandler
            aTextHandler = New TextHandler.TextHandler(sIniFilePath, "SECS")

            ' SXMLPath
            aTextHandler.SetData("SXMLFile", SXMLPath)

            ' Entity
            If Me.Entity = enumSecsEntity.sPassive Then
                aTextHandler.SetData("Entity", "Passive")
            Else
                aTextHandler.SetData("Entity", "Active")
            End If

            ' Role
            If Me.Role = enumRole.sEquipment Then
                aTextHandler.SetData("Role", "Equipment")
            Else
                aTextHandler.SetData("Role", "Host")
            End If

            ' LinkType
            aTextHandler.SetData("LinkType", LinkType)

            ' DeviceName
            aTextHandler.SetData("DeviceName", DeviceName)

            ' Device ID
            buffer = "0x" + Me.DeviceID.ToString("X4")
            aTextHandler.SetData("DeviceID", buffer)

            ' Retry
            aTextHandler.SetData("Retry", Retry.ToString)

            ' IP
            aTextHandler.SetData("IP", IP)

            ' Port
            aTextHandler.SetData("Port", Port.ToString)

            ' Timeout 設定值
            aTextHandler.SetData("T1", T1Timeout.ToString)
            aTextHandler.SetData("T2", T2Timeout.ToString)
            aTextHandler.SetData("T3", T3Timeout.ToString)
            aTextHandler.SetData("T4", T4Timeout.ToString)
            aTextHandler.SetData("T5", T5Timeout.ToString)
            aTextHandler.SetData("T6", T6Timeout.ToString)
            aTextHandler.SetData("T7", T7Timeout.ToString)
            aTextHandler.SetData("T8", T8Timeout.ToString)
            aTextHandler.SetData("CA", CATimeout.ToString)

            ' Heartbeat 設定值
            aTextHandler.SetData("InterleaveEnabled", InterleaveEnabled.ToString)
            aTextHandler.SetData("HeartbeatEnabled", HeartbeatEnabled.ToString)
            aTextHandler.SetData("HeartbeatInterval", HeartbeatInterval.ToString)
            aTextHandler.SetData("HeartbeatTag", HeartbeatTag)

        End Sub

#End Region


    End Class

End Namespace

