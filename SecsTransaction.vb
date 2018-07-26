Imports System
Imports System.Collections.Generic

Namespace SecsDriver

	Public Class SecsTransaction


        Private sTransactionStartTime As DateTime

        Public sTimeoutList As List(Of Timeout)


        ' ----------------------- Property ------------------------------

        ' 存取 Primary Message
        Public Property Primary As SecsMessage


        ' 存取 Secondary Message 
        Public Property Secondary As SecsMessage


        ' 存取 Transaction State
        Public Property State As enumSecsTransactionState


        ' ----------------------- Property ------------------------------

        ' 建構子
        Public Sub New()

            State = enumSecsTransactionState.Create

            sTransactionStartTime = System.DateTime.Now

            sTimeoutList = New List(Of Timeout)

        End Sub


        ' 設定 Timeout
        Public Sub SetTimeout(ByRef sTimeoutType As enumTimeout, ByRef siniFile As IniFile)

            Select Case sTimeoutType

                Case enumTimeout.T1
                    Dim sTimeout As Timeout = New Timeout(siniFile.T1Timeout, sTimeoutType)
                    SyncLock sTimeoutList
                        sTimeoutList.Add(sTimeout)
                    End SyncLock

                Case enumTimeout.T2
                    Dim sTimeout As Timeout = New Timeout(siniFile.T2Timeout, sTimeoutType)
                    SyncLock sTimeoutList
                        sTimeoutList.Add(sTimeout)
                    End SyncLock

                Case enumTimeout.T3
                    Dim sTimeout As Timeout = New Timeout(siniFile.T3Timeout, sTimeoutType)
                    SyncLock sTimeoutList
                        sTimeoutList.Add(sTimeout)
                    End SyncLock

                Case enumTimeout.T4
                    Dim sTimeout As Timeout = New Timeout(siniFile.T4Timeout, sTimeoutType)
                    SyncLock sTimeoutList
                        sTimeoutList.Add(sTimeout)
                    End SyncLock

                Case enumTimeout.T5
                    Dim sTimeout As Timeout = New Timeout(siniFile.T5Timeout, sTimeoutType)
                    SyncLock sTimeoutList
                        sTimeoutList.Add(sTimeout)
                    End SyncLock

                Case enumTimeout.T6
                    Dim sTimeout As Timeout = New Timeout(siniFile.T6Timeout, sTimeoutType)
                    SyncLock sTimeoutList
                        sTimeoutList.Add(sTimeout)
                    End SyncLock

                Case enumTimeout.T7
                    Dim sTimeout As Timeout = New Timeout(siniFile.T7Timeout, sTimeoutType)
                    SyncLock sTimeoutList
                        sTimeoutList.Add(sTimeout)
                    End SyncLock

                Case enumTimeout.T8
                    Dim sTimeout As Timeout = New Timeout(siniFile.T8Timeout, sTimeoutType)
                    SyncLock sTimeoutList
                        sTimeoutList.Add(sTimeout)
                    End SyncLock

            End Select
        End Sub


    End Class

End Namespace

