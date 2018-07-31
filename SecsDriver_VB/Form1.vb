Imports System
Imports System.Text
Imports SecsDriver_VB.SecsDriver
Imports System.Threading


Public Class Form1

    Dim aWrapper As SecsWrapper

    Private startSend As Thread

    Private cleanTextBox As Thread

    Private StopSend As Boolean
    Private StopClean As Boolean

    Private aPrimarySent As Integer = 0
    Private aPrimaryReceived As Integer = 0
    Private aSecondarySent As Integer = 0
    Private aSecondaryReceived As Integer = 0


    Public Delegate Sub UpdateGUI(ByVal sysMessage As String, ByRef control As Control)

    Public Delegate Sub CleanGUI(ByRef control As Control)


    Public Sub New()

        ' 設計工具需要此呼叫。
        InitializeComponent()

        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。

        btnClean.Text = "Start Clean"

    End Sub


    Private Sub btnConnect_Click(sender As Object, e As EventArgs) Handles btnConnect.Click

        aWrapper = New SecsWrapper()

        Dim sDelegateSecsConnect As New delegateSecsConnectStateChange(AddressOf SecsConnect)
        Dim sDelegatePrimaryReceive As New delegatePrimaryReceived(AddressOf PrimaryReceive)
        Dim sDelegatePrimarySent As New delegatePrimarySent(AddressOf PrimarySent)
        Dim sDelegateSecondaryReceive As New delegateSecondaryReceived(AddressOf SecondaryReceive)
        Dim sDelegateSecondarySent As New delegateSecondarySent(AddressOf SecondarySent)
        Dim sDelegateMessageError As New delegateMessageError(AddressOf ErrorMessage)
        Dim sDelegateMessageInfo As New delegateMessageInfo(AddressOf MessageInfo)
        Dim sDelegateSecsError As New delegateError(AddressOf SecsError)
        Dim sDelegateTimeout As New delegateTimeout(AddressOf HandleTimeout)

        AddHandler aWrapper.OnSecsConnectStateChange, sDelegateSecsConnect
        AddHandler aWrapper.OnPrimaryReceived, sDelegatePrimaryReceive
        AddHandler aWrapper.OnPrimarySent, sDelegatePrimarySent
        AddHandler aWrapper.OnSecondaryReceived, sDelegateSecondaryReceive
        AddHandler aWrapper.OnSecondarySent, sDelegateSecondarySent
        AddHandler aWrapper.OnMessageError, sDelegateMessageError
        AddHandler aWrapper.OnMessageInfo, sDelegateMessageInfo
        AddHandler aWrapper.OnError, sDelegateSecsError
        AddHandler aWrapper.OnTimeout, sDelegateTimeout

        aWrapper.Open()

        If aWrapper.siniFile IsNot Nothing Then
            lblEntity.Text = aWrapper.siniFile.Entity.ToString
        End If

    End Sub


    Private Sub btnDisconnect_Click(sender As Object, e As EventArgs) Handles btnDisconnect.Click

        StopSend = True
        aWrapper.Close()

    End Sub


    Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click

        startSend = New Thread(AddressOf SendMessage)
        startSend.IsBackground = True
        startSend.Start()

    End Sub


    Private Sub btnStopSend_Click(sender As Object, e As EventArgs) Handles btnStopSend.Click

        StopSend = True

    End Sub


    Private Sub btnClean_Click(sender As Object, e As EventArgs) Handles btnClean.Click

        If btnClean.Text = "Start Clean" Then

            cleanTextBox = New Thread(AddressOf CleanTbx)
            cleanTextBox.IsBackground = True
            cleanTextBox.Start()

            btnClean.Text = "Stop Clean"
        Else
            StopClean = True

            btnClean.Text = "Start Clean"

        End If


    End Sub


    Private Sub SendMessage()

        StopSend = False

        Dim sendcount As Integer

        If tbxTotalSend.Text <> "" Then

            sendcount = Convert.ToInt64(tbxTotalSend.Text)

        End If

        Do
            If tbxTotalSend.Text = "" Then

                Dim temp As SecsMessage = aWrapper.GetMessageByName(random())
                aWrapper.SendPrimary(temp)

                Application.DoEvents()
                Dim tempTime As Integer = Convert.ToInt16(tbxThreadSleep.Text)
                Thread.Sleep(tempTime)

            ElseIf sendcount = 0 Then

                StopSend = True

            Else

                Dim temp As SecsMessage = aWrapper.GetMessageByName(random())
                aWrapper.SendPrimary(temp)

                Application.DoEvents()
                Dim tempTime As Integer = Convert.ToInt16(tbxThreadSleep.Text)
                Thread.Sleep(tempTime)

                sendcount = sendcount - 1

            End If

        Loop Until StopSend = True

    End Sub


    Private Sub CleanTbx()

        StopClean = False

        Do
            Dim temp As CleanGUI = New CleanGUI(AddressOf GUIClean)
            Invoke(temp, RichTextBox1)

            Application.DoEvents()
            Thread.Sleep(10000)

        Loop Until StopClean = True

    End Sub


    Private Function random() As String

        Dim temp As List(Of String) = New List(Of String)

        temp.Add("S1F1_AreYouThere")
        temp.Add("S1F3_SelectedEquipmentStatusRequest")
        temp.Add("S1F11_StatusVariableNamelistRequest")
        temp.Add("S1F13_CommMCSRequest")
        temp.Add("S1F13_CommStkcRequest")
        temp.Add("S1F15_MCSOffnlineRequest")
        temp.Add("S1F17_OnlineRequestFromMCS")
        temp.Add("S2F13_EquipmentConstantRequest")
        temp.Add("S2F15_NewEquipmentConstantRequest")
        temp.Add("S2F17_DateTimeRequest")
        temp.Add("S2F21_RemoteCommandRequest")
        temp.Add("S2F25_LoopbackDiagnosticRequest")
        temp.Add("S2F29_EquipmentConstantNamelistRequest")
        temp.Add("S2F31_DateTimeSetRequest")
        temp.Add("S2F33_DefineReport")
        temp.Add("S2F35_LinkEventReport")
        temp.Add("S2F37_EnableDisableEventReport")
        temp.Add("S2F41_CarrierTypeSetCommand")
        temp.Add("S2F41_RemoteAbortCommand")
        temp.Add("S2F41_RemoteBlockCommand")
        temp.Add("S2F41_RemoteCancelCommand")
        temp.Add("S2F41_RemoteIDReaderModeChangeCommand")
        temp.Add("S2F41_RemoteInstallCommand")
        temp.Add("S2F41_RemoteLocateCommand")
        temp.Add("S2F49_RemoteTransferCommand")
        temp.Add("S5F1_AlarmReport")
        temp.Add("S5F3_EnableAlarmSend")
        temp.Add("S5F5_ListAlarmRequest")
        temp.Add("S5F7_ListEnabledAlarmRequest")
        temp.Add("S6F11_ActiveCarrierReq")
        temp.Add("S6F11_Alarm")
        temp.Add("S6F11_AlarmClear")
        temp.Add("S6F11_AlarmSet")

        Dim a As Random = New Random
        Dim b As Integer = a.Next(0, temp.Count)
        Return temp(b)

    End Function


    Public Sub SecsConnect(ByRef aMessage As String)

        Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
        Invoke(temp, "Secs Connect State : " & aMessage, RichTextBox1)

    End Sub


    Public Sub PrimaryReceive(ByRef aTransaction As SecsTransaction)

        Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
        Invoke(temp, "Primary Receive : " & aTransaction.Primary.MessageName, RichTextBox1)

        aPrimaryReceived = aPrimaryReceived + 1
        Invoke(temp, "PrimaryReceived", lblPrimaryReceived)

    End Sub


    Public Sub PrimarySent(ByRef aTransaction As SecsTransaction)

        Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
        Invoke(temp, "Primary Sent : " & aTransaction.Primary.MessageName, RichTextBox1)

        aPrimarySent = aPrimarySent + 1
        Invoke(temp, "PrimarySent", lblPrimarySent)

    End Sub


    Public Sub SecondaryReceive(ByRef aTransaction As SecsTransaction)

        Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
        Invoke(temp, "Secondary Receive : " & aTransaction.Secondary.MessageName, RichTextBox1)
        aSecondaryReceived = aSecondaryReceived + 1
        Invoke(temp, "SecondaryReceived", lblSecondaryReceived)

    End Sub


    Public Sub SecondarySent(ByRef aTransaction As SecsTransaction)

        Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
        Invoke(temp, "Secondary Sent : " & aTransaction.Secondary.MessageName, RichTextBox1)
        aSecondarySent = aSecondarySent + 1
        Invoke(temp, "SecondarySent", lblSecondarySent)

    End Sub


    Public Sub ErrorMessage(ByRef message As SecsMessage, ByVal errorMessage As String)

        Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)

        If message.MessageName Is Nothing Then
            Invoke(temp, "Control Message " & errorMessage, RichTextBox1)
        Else
            Invoke(temp, message.MessageName & " : " & errorMessage, RichTextBox1)
        End If

    End Sub


    Public Sub MessageInfo(ByRef message As SecsMessage, ByVal aMessage As String)

        Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
        Invoke(temp, message.MessageName & " : " & aMessage, RichTextBox1)

    End Sub


    Public Sub SecsError(ByRef aMessage As String)

        Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
        Invoke(temp, aMessage, RichTextBox1)

    End Sub


    Public Sub HandleTimeout(ByRef aMessage As String)

        Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
        Invoke(temp, aMessage, RichTextBox1)

    End Sub


    Public Sub GUIUpdate(ByVal sysMessage As String, ByRef control As Control)

        If sysMessage = "PrimarySent" Then
            lblPrimarySent.Text = aPrimarySent.ToString

        ElseIf sysMessage = "PrimaryReceived" Then
            lblPrimaryReceived.Text = aPrimaryReceived.ToString

        ElseIf sysMessage = "SecondarySent" Then
            lblSecondarySent.Text = aSecondarySent.ToString

        ElseIf sysMessage = "SecondaryReceived" Then
            lblSecondaryReceived.Text = aSecondaryReceived.ToString

        Else
            RichTextBox1.Text = RichTextBox1.Text & Now.ToString("yyyy/MM/dd hh:mm:ss.fff") & " : " & sysMessage & vbLf
        End If

    End Sub


    Public Sub GUIClean(ByRef control As Control)

        RichTextBox1.Text = ""

    End Sub



End Class
