Imports System

Namespace SecsDriver

    ''' <summary>
    ''' Timeout 物件
    ''' </summary>
    Public Class Timeout


#Region "Private 屬性"

        ''' <summary>
        ''' 發生 Timeout 的時間值(秒)
        ''' </summary>
        Private sTimeoutValue As Double

        ''' <summary>
        ''' 開始啟動 Timeout 的時間
        ''' </summary>
        Private sStartTime As DateTime

#End Region


#Region "Public 屬性"

        ' Timeout Type
        Public sTimeoutType As enumTimeout

#End Region


#Region "建構子"

        ''' <summary>
        ''' 建構子
        ''' </summary>
        ''' <param name="aTimeoutValue"></param>
        ''' <param name="aTimeoutType"></param>
        Public Sub New(ByVal aTimeoutValue As Double, ByVal aTimeoutType As enumTimeout)

            sTimeoutValue = aTimeoutValue
            sStartTime = System.DateTime.Now
            sTimeoutType = aTimeoutType

        End Sub

#End Region


#Region "Public Method"

        ''' <summary>
        ''' 查是否發生 Timeout
        ''' </summary>
        ''' <returns></returns>
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

#End Region


    End Class

End Namespace

