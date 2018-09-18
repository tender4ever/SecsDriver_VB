Imports System
Imports System.Text
Imports System.Runtime.InteropServices

Namespace TextHandler

    Public Class TextHandler

        ' 引用外部程式
        <DllImport("kernel32.dll", EntryPoint:="GetPrivateProfileString")>
        Private Shared Function GetIniData(ByVal aAppName As String, ByVal aKeyName As String, ByVal aDefault As String, ByVal aReturnedString As StringBuilder, ByVal aSize As Integer, ByVal aFileName As String) As Integer

        End Function

        ' 引用外部程式
        <DllImport("kernel32.dll", EntryPoint:="WritePrivateProfileString")>
        Private Shared Function SetIniData(ByVal aAppName As String, ByVal aKeyName As String, ByVal aString As String, ByVal aFileName As String)

        End Function

        Public sFilePath As String               ' File 檔案路徑
        Public sSection As String               ' Section


        ' 建構子
        Public Sub New(ByRef aFilePath As String, ByRef aSection As String)

            sFilePath = aFilePath
            sSection = aSection

        End Sub


        ' 取得資料
        Public Function GetData(ByRef aKeyName As String, ByRef aDefault As String) As String

            Dim chars As Integer = 256
            Dim buffer As StringBuilder = New StringBuilder(chars)

            GetIniData(sSection, aKeyName, aDefault, buffer, chars, sFilePath)

            Return buffer.ToString

        End Function


        ' 設定資料
        Public Sub SetData(ByRef aKeyName As String, ByRef aValue As String)

            SetIniData(sSection, aKeyName, aValue, sFilePath)

        End Sub

    End Class

End Namespace

