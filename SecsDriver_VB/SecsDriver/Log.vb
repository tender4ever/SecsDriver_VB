Imports System.Threading
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports SecsDriver_VB.SocketDriver

Namespace SecsDriver

    Public Class Log

        Private BinaryLogQueue As Concurrent.ConcurrentQueue(Of String)
        Private TxLogQueue As Concurrent.ConcurrentQueue(Of String)

        Private sBinaryLog As FileStream                            ' Binary Log
        Private sTxLog As FileStream                                ' Tx Log


        Public Sub New()


            BinaryLogQueue = New Concurrent.ConcurrentQueue(Of String)
            TxLogQueue = New Concurrent.ConcurrentQueue(Of String)

            ' 建立 Log Txt 檔
            sBinaryLog = New FileStream(Application.StartupPath & "\SECSLog\BinaryLog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
            sBinaryLog.Close()
            sTxLog = New FileStream(Application.StartupPath & "\SECSLog\TxLog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
            sTxLog.Close()

            ' New 一個 Thread 來持續檢查 Log.txt是否過大
            Dim sWatchLogTxtThread As New Thread(New ThreadStart(AddressOf CheckLogTxt))
            sWatchLogTxtThread.IsBackground = True
            sWatchLogTxtThread.Start()

            ' NEW 一個 Thread 來持續紀錄 Log
            Dim sDoLogThread As New Thread(New ThreadStart(AddressOf DoLog))
            sDoLogThread.IsBackground = True
            sDoLogThread.Start()

        End Sub


        ' 紀錄 BinaryLog
        Public Sub DoBinaryLog(ByVal aBinaryLogString As String)

            BinaryLogQueue.Enqueue(aBinaryLogString)

        End Sub


        ' 紀錄 TxLog
        Public Sub DoTxLog(ByVal aTxLogString As String)

            TxLogQueue.Enqueue(aTxLogString)

        End Sub


        ' 檢查 Log.txt 檔的大小
        Private Sub CheckLogTxt()
            Do
                Dim tempDir As New DirectoryInfo(Application.StartupPath & "\SECSLog")
                Dim tempBinaryLogs As FileInfo() = tempDir.GetFiles("*BinaryLog.txt*")
                Dim tempTxLogs As FileInfo() = tempDir.GetFiles("TxLog.txt")
                Dim tempBinaryLog As FileInfo = tempBinaryLogs(0)
                Dim tempTxLog As FileInfo = tempTxLogs(0)

                If tempBinaryLog IsNot Nothing Then

                    If tempBinaryLog.Length > 500000000 Then

                        FileSystem.RenameFile(Application.StartupPath & "\SECSLog\BinaryLog.txt", "BinaryLog_" & Now.ToString("yyyyMMdd-hh.mm.ss") & ".txt")

                        ' 建立 Log Txt 檔
                        sBinaryLog = New FileStream(Application.StartupPath & "\SECSLog\BinaryLog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
                        sBinaryLog.Close()

                    End If
                End If

                If tempTxLog IsNot Nothing Then

                    If tempTxLog.Length > 500000000 Then

                        FileSystem.RenameFile(Application.StartupPath & "\SECSLog\TxLog.txt", "TxLog_" & Now.ToString("yyyyMMdd-hh.mm.ss") & ".txt")

                        ' 建立 Log Txt 檔
                        sTxLog = New FileStream(Application.StartupPath & "\SECSLog\TxLog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
                        sTxLog.Close()

                    End If
                End If

                Application.DoEvents()
                Thread.Sleep(1000)

            Loop While True
        End Sub


        ' 執行記錄 Log
        Private Sub DoLog()

            Do
                If BinaryLogQueue.Count > 0 Then

                    Dim tempBinaryLogString As String = Nothing
                    BinaryLogQueue.TryDequeue(tempBinaryLogString)

                    ' 紀錄 Binary Log
                    Try
                        Dim tempSW As StreamWriter = File.AppendText(Application.StartupPath & "\SECSLog\BinaryLog.txt")

                        tempSW.WriteLine(tempBinaryLogString)
                        tempSW.Flush()
                        tempSW.Close()
                        tempSW.Dispose()

                    Catch ex As Exception

                        Dim tempSW As StreamWriter = File.AppendText(Application.StartupPath & "\SECSLog\BinaryLog.txt")

                        tempSW.WriteLine("Do BinaryLog Error !")
                        tempSW.Flush()
                        tempSW.Close()
                        tempSW.Dispose()

                        'RaiseEvent OnError("Do BinaryLog Error !")

                    End Try
                End If

                If TxLogQueue.Count > 0 Then

                    Dim tempTxLogString As String = Nothing
                    TxLogQueue.TryDequeue(tempTxLogString)

                    ' 紀錄 SML Log
                    Try
                        Dim tempSW As StreamWriter = File.AppendText(Application.StartupPath & "\SECSLog\TxLog.txt")

                        tempSW.WriteLine(tempTxLogString)
                        tempSW.Flush()
                        tempSW.Close()
                        tempSW.Dispose()

                    Catch ex As Exception

                        Dim tempSW As StreamWriter = File.AppendText(Application.StartupPath & "\SECSLog\TxLog.txt")

                        tempSW.WriteLine("Do SMLLog Error !")
                        tempSW.Flush()
                        tempSW.Close()
                        tempSW.Dispose()

                        'RaiseEvent OnError("Do SMLLog Error !")

                    End Try

                End If

                If BinaryLogQueue.Count = 0 And TxLogQueue.Count = 0 Then

                    Thread.Sleep(1000)

                End If

            Loop While True

        End Sub

    End Class

End Namespace