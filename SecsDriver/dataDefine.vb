Namespace SecsDriver

	' Secs Entity
	' STKC = Passive
	' MCS = Active
	Public Enum enumSecsEntity

		sActive
		sPassive

	End Enum

	' Role
	' STKC = Equipment
	' MCS = Host
	Public Enum enumRole

		sHost
		sEquipment

	End Enum

	' Secs Connection State
	Public Enum enumSecsConnectionState

		sNone
		sNotConnected
		sNotSelected
		sSelected
		sSeparate

	End Enum

	' Message Type
	Public Enum enumMessageType

		sDataMessage
		sSelectRequest
		sSelectResponse
		sDeselectRequest
		sDeselectResponse
		sLinktestRequest
		sLinktestResponse
		sRejectRequest
		sSeparateRequest

	End Enum

	' Item Format Code
	Public Enum enumItemFormatCode

		sList = &H0
		sBinary = &H20
		sBoolean = &H24
		sASCII = &H40
		sJIS8 = &H44
		sTwoByteChar = &H48
		sINT1 = &H64
		sINT2 = &H68
		sINT4 = &H70
		sINT8 = &H60
		sFloat4 = &H90
		sFloat8 = &H80
		sUINT1 = &HA4
		sUINT2 = &HA8
		sUINT4 = &HB0
		sUINT8 = &HA0

	End Enum

    ' Secs Transaction State
    Public Enum enumSecsTransactionState

        Create
        PrimarySent
        PrimaryReceived
        SecondarySent
        SecondaryReceived
        WaitForDelete
        'WaitForSendingPrimary
        'SendingPrimary
        'WaitForSendingSecondary
        'SendingSecondary
        'WaitForReceivingSecondary
        'SendPrimaryFailed
        'SendSecondaryFailed

    End Enum

    ' Timeout Type
    Public Enum enumTimeout

        T1
        T2
        T3
        T4
        T5
        T6
        T7
        T8

    End Enum

End Namespace
