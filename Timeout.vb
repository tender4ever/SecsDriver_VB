Imports System

Namespace SecsDriver

    Public Class Timeout

        ' Timeout 值
        Private sTimeoutValue As Double

        ' 開始啟動 Timeout 的時間
        Private sStartTime As DateTime

        ' Timeout Type
        Public sTimeoutType As enumTimeout


        ' 建構子
        Public Sub New(ByVal aTimeoutValue As Double, ByVal aTimeoutType As enumTimeout)

            sTimeoutValue = aTimeoutValue
            sStartTime = System.DateTime.Now
            sTimeoutType = aTimeoutType

        End Sub


        ' 檢查是否發生 Timeout
        Public Function CheckTimeout() As Boolean

            ' 目前時間
            Dim NowTime As DateTime = System.DateTime.Now

            ' 計算距離啟動 Timeout 
            Dim span As TimeSpan = NowTime - sStartTime

            If span.TotalSeconds > Me.sTimeoutValue Then

                Return True

            End If

            Return False

        End Function

    End Class

End Namespace

