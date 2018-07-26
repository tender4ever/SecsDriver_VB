Imports System
Imports System.Text
Imports System.Runtime.InteropServices

Namespace SecsDriver

	Public Class IniFile

		' 引用外部程式
		<DllImport("kernel32.dll", EntryPoint:="GetPrivateProfileString")>
		Private Shared Function GetIniData(ByVal aAppName As String, ByVal aKeyName As String, ByVal aDefault As String, ByVal aReturnedString As StringBuilder, ByVal aSize As Integer, ByVal aFileName As String) As Integer

		End Function

		' 引用外部程式
		<DllImport("kernel32.dll", EntryPoint:="WritePrivateProfileString")>
		Private Shared Function SetIniData(ByVal aAppName As String, ByVal aKeyName As String, ByVal aString As String, ByVal aFileName As String)

		End Function


		Private sIniFilePath As String              ' iniFile 檔案路徑
        Private sSection As String                  ' Section


        ' 建構子
        Protected Friend Sub New(ByVal aIniFilePath As String)

            Dim chars As Integer = 256
            Dim buffer As StringBuilder = New StringBuilder(chars)

            ' IniFilePath
            sIniFilePath = aIniFilePath

            ' Section
            sSection = "SECS"

            ' SXMLPath
            GetIniData(sSection, "SXMLFile", "SecsWrapperX_AMHS.sxml", buffer, chars, sIniFilePath)
            SXMLPath = buffer.ToString

            ' Entity
            GetIniData(sSection, "Entity", "Passive", buffer, chars, sIniFilePath)

            If buffer.ToString = "Passive" Then
                Entity = enumSecsEntity.sPassive
            Else
                Entity = enumSecsEntity.sActive
            End If

            ' Role
            GetIniData(sSection, "Role", "Equipment", buffer, chars, sIniFilePath)

            If buffer.ToString = "Equipment" Then
                Role = enumRole.sEquipment
            Else
                Role = enumRole.sHost
            End If

            ' LinkType
            GetIniData(sSection, "LinkType", "HSMS", buffer, chars, sIniFilePath)
            LinkType = buffer.ToString

            ' DeviceName
            GetIniData(sSection, "DeviceName", "HSMS", buffer, chars, sIniFilePath)
            DeviceName = buffer.ToString

            ' Device ID
            GetIniData(sSection, "DeviceID", "0x0000", buffer, chars, sIniFilePath)
            DeviceID = Convert.ToUInt16(buffer.ToString, 16)

            ' 讀取 Retry
            GetIniData(sSection, "Retry", "3", buffer, chars, sIniFilePath)
            Retry = Convert.ToInt16(buffer.ToString)

            ' 讀取 IP
            GetIniData(sSection, "IP", "127.0.0.1", buffer, chars, sIniFilePath)
            IP = buffer.ToString

            ' 讀取 Port
            GetIniData(sSection, "Port", "5000", buffer, chars, sIniFilePath)
            Port = Convert.ToInt32(buffer.ToString)

            ' 讀取 Timeout 設定值
            GetIniData(sSection, "T1", "1", buffer, chars, sIniFilePath)
            T1Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T2", "10", buffer, chars, sIniFilePath)
            T2Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T3", "45", buffer, chars, sIniFilePath)
            T3Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T4", "10", buffer, chars, sIniFilePath)
            T4Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T5", "10", buffer, chars, sIniFilePath)
            T5Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T6", "5", buffer, chars, sIniFilePath)
            T6Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T7", "10", buffer, chars, sIniFilePath)
            T7Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T8", "5", buffer, chars, sIniFilePath)
            T8Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "CA", "30", buffer, chars, sIniFilePath)
            CATimeout = Convert.ToInt32(buffer.ToString)

            ' 讀取 Heartbeat 設定值
            GetIniData(sSection, "InterleaveEnabled", "True", buffer, chars, sIniFilePath)

            If buffer.ToString = "False" Then
                InterleaveEnabled = False
            Else
                InterleaveEnabled = True
            End If

            GetIniData(sSection, "HeartbeatEnabled", "False", buffer, chars, sIniFilePath)

            If buffer.ToString = "False" Then
                HeartbeatEnabled = False
            Else
                HeartbeatEnabled = True
            End If

            GetIniData(sSection, "HeartbeatInterval", "30", buffer, chars, sIniFilePath)
            HeartbeatInterval = Convert.ToUInt32(buffer.ToString)

            GetIniData(sSection, "HeartbeatTag", " ", buffer, chars, sIniFilePath)
            HeartbeatTag = buffer.ToString

            ' 讀取 AutoReply 設定值
            GetIniData(sSection, "AutoReply", "False", buffer, chars, sIniFilePath)

            If buffer.ToString = "False" Then
                AutoReply = False
            Else
                AutoReply = True
            End If

            ' 讀取 BinaryLog 設定值
            GetIniData(sSection, "BinaryLog", "False", buffer, chars, sIniFilePath)

            If buffer.ToString = "False" Then
                BinaryLog = False
            Else
                BinaryLog = True
            End If

            ' 讀取 TxLog 設定值
            GetIniData(sSection, "TxLog", "False", buffer, chars, sIniFilePath)

            If buffer.ToString = "False" Then
                TxLog = False
            Else
                TxLog = True
            End If

        End Sub


        ' 重新讀取 Ini File
        Protected Friend Sub ReLoadIniFile()

            Dim chars As Integer = 256
            Dim buffer As StringBuilder = New StringBuilder(chars)

            ' Section
            sSection = "SECS"

            ' SXMLPath
            GetIniData(sSection, "SXMLFile", "SecsWrapperX_AMHS.sxml", buffer, chars, sIniFilePath)
            SXMLPath = buffer.ToString

            ' Entity
            GetIniData(sSection, "Entity", "Passive", buffer, chars, sIniFilePath)

            If buffer.ToString = "Passive" Then
                Entity = enumSecsEntity.sPassive
            Else
                Entity = enumSecsEntity.sActive
            End If

            ' Role
            GetIniData(sSection, "Role", "Equipment", buffer, chars, sIniFilePath)

            If buffer.ToString = "Equipment" Then
                Role = enumRole.sEquipment
            Else
                Role = enumRole.sHost
            End If

            ' LinkType
            GetIniData(sSection, "LinkType", "HSMS", buffer, chars, sIniFilePath)
            LinkType = buffer.ToString

            ' DeviceName
            GetIniData(sSection, "DeviceName", "HSMS", buffer, chars, sIniFilePath)
            DeviceName = buffer.ToString

            ' Device ID
            GetIniData(sSection, "DeviceID", "0x0000", buffer, chars, sIniFilePath)
            DeviceID = Convert.ToUInt16(buffer.ToString, 16)

            ' 讀取 Retry
            GetIniData(sSection, "Retry", "3", buffer, chars, sIniFilePath)
            Retry = Convert.ToInt16(buffer.ToString)

            ' 讀取 IP
            GetIniData(sSection, "IP", "127.0.0.1", buffer, chars, sIniFilePath)
            IP = buffer.ToString

            ' 讀取 Port
            GetIniData(sSection, "Port", "5000", buffer, chars, sIniFilePath)
            Port = Convert.ToInt32(buffer.ToString)

            ' 讀取 Timeout 設定值
            GetIniData(sSection, "T1", "1", buffer, chars, sIniFilePath)
            T1Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T2", "10", buffer, chars, sIniFilePath)
            T2Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T3", "45", buffer, chars, sIniFilePath)
            T3Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T4", "10", buffer, chars, sIniFilePath)
            T4Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T5", "10", buffer, chars, sIniFilePath)
            T5Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T6", "5", buffer, chars, sIniFilePath)
            T6Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T7", "10", buffer, chars, sIniFilePath)
            T7Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "T8", "5", buffer, chars, sIniFilePath)
            T8Timeout = Convert.ToInt32(buffer.ToString)
            GetIniData(sSection, "CA", "30", buffer, chars, sIniFilePath)
            CATimeout = Convert.ToInt32(buffer.ToString)

            ' 讀取 Heartbeat 設定值
            GetIniData(sSection, "InterleaveEnabled", "True", buffer, chars, sIniFilePath)

            If buffer.ToString = "False" Then
                InterleaveEnabled = False
            Else
                InterleaveEnabled = True
            End If

            GetIniData(sSection, "HeartbeatEnabled", "False", buffer, chars, sIniFilePath)

            If buffer.ToString = "False" Then
                HeartbeatEnabled = False
            Else
                HeartbeatEnabled = True
            End If

            GetIniData(sSection, "HeartbeatInterval", "30", buffer, chars, sIniFilePath)
            HeartbeatInterval = Convert.ToUInt32(buffer.ToString)

            GetIniData(sSection, "HeartbeatTag", " ", buffer, chars, sIniFilePath)
            HeartbeatTag = buffer.ToString

            ' 讀取 AutoReply 設定值
            GetIniData(sSection, "AutoReply", "False", buffer, chars, sIniFilePath)

            If buffer.ToString = "False" Then
                AutoReply = False
            Else
                AutoReply = True
            End If

            ' 讀取 BinaryLog 設定值
            GetIniData(sSection, "BinaryLog", "False", buffer, chars, sIniFilePath)

            If buffer.ToString = "False" Then
                BinaryLog = False
            Else
                BinaryLog = True
            End If

            ' 讀取 TxLog 設定值
            GetIniData(sSection, "TxLog", "False", buffer, chars, sIniFilePath)

            If buffer.ToString = "False" Then
                TxLog = False
            Else
                TxLog = True
            End If

        End Sub


        ' 儲存 Ini File
        Protected Friend Sub SaveIniFile()

            Dim buffer As String

            ' SXMLPath
            SetIniData(sSection, "SXMLFile", SXMLPath, sIniFilePath)

            ' Entity
            If Me.Entity = enumSecsEntity.sPassive Then
                buffer = "Passive"
            Else
                buffer = "Active"
            End If
            SetIniData(sSection, "Entity", buffer, sIniFilePath)

            ' Role
            If Me.Role = enumRole.sEquipment Then
                buffer = "Equipment"
            Else
                buffer = "Host"
            End If
            SetIniData(sSection, "Role", buffer, sIniFilePath)

            ' LinkType
            SetIniData(sSection, "LinkType", LinkType, sIniFilePath)

            ' DeviceName
            SetIniData(sSection, "DeviceName", DeviceName, sIniFilePath)

            ' Device ID
            buffer = "0x" + Me.DeviceID.ToString("X4")
            SetIniData(sSection, "DeviceID", buffer, sIniFilePath)

            ' Retry
            buffer = Retry.ToString
            SetIniData(sSection, "Retry", buffer, sIniFilePath)

            ' IP
            SetIniData(sSection, "IP", IP, sIniFilePath)

            ' Port
            buffer = Port.ToString
            SetIniData(sSection, "Port", buffer, sIniFilePath)

            ' Timeout 設定值
            SetIniData(sSection, "T1", T1Timeout.ToString, sIniFilePath)
            SetIniData(sSection, "T2", T2Timeout.ToString, sIniFilePath)
            SetIniData(sSection, "T3", T3Timeout.ToString, sIniFilePath)
            SetIniData(sSection, "T4", T4Timeout.ToString, sIniFilePath)
            SetIniData(sSection, "T5", T5Timeout.ToString, sIniFilePath)
            SetIniData(sSection, "T6", T6Timeout.ToString, sIniFilePath)
            SetIniData(sSection, "T7", T7Timeout.ToString, sIniFilePath)
            SetIniData(sSection, "T8", T8Timeout.ToString, sIniFilePath)
            SetIniData(sSection, "CA", CATimeout.ToString, sIniFilePath)

            ' Heartbeat 設定值
            SetIniData(sSection, "InterleaveEnabled", InterleaveEnabled.ToString, sIniFilePath)
            SetIniData(sSection, "HeartbeatEnabled", HeartbeatEnabled.ToString, sIniFilePath)
            SetIniData(sSection, "HeartbeatInterval", HeartbeatInterval.ToString, sIniFilePath)
            SetIniData(sSection, "HeartbeatTag", HeartbeatTag, sIniFilePath)

        End Sub


        ' --------------------- Property ----------------------------

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


    End Class

End Namespace

