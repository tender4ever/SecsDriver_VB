Imports System
Imports System.Text
Imports SecsDriver_VB.SecsDriver
Imports System.Threading


Public Class Form1

    Dim aWrapper As SecsWrapper                     ' SecsWrapper


    Private startSend As Thread                     ' Send Message Thread

    Private cleanTextBox As Thread                  ' Clean Thread

    Private StopSend As Boolean
    Private StopClean As Boolean
    Private ShowTextBox As Boolean = True
    Private ShowFunction As Boolean = False

    Private Shared aPrimarySent As Integer = 0
    Private Shared aPrimaryReceived As Integer = 0
    Private Shared aSecondarySent As Integer = 0
    Private Shared aSecondaryReceived As Integer = 0

    Dim SecsMessage As SecsMessage = Nothing
    Dim secsItemList As List(Of SecsItem) = New List(Of SecsItem)

    Public Delegate Sub UpdateGUI(ByVal sysMessage As String, ByRef control As Control)

    Public Delegate Sub CleanGUI(ByRef control As Control)


    Public Sub New()

        ' 設計工具需要此呼叫。
        InitializeComponent()

        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。

        If ShowFunction = True Then

            Me.Size = New Size(1000, 700)
        Else
            Me.Size = New Size(550, 700)
        End If

    End Sub


    ' 連線
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
            lblIP.Text = aWrapper.siniFile.IP.ToString
            lblPort.Text = aWrapper.siniFile.Port.ToString

            Dim messageList As ArrayList = New ArrayList()
            For i As Integer = 0 To aWrapper.sSXML.sMessageList.Count - 1
                messageList.Add(aWrapper.sSXML.sMessageList(i).MessageName)
            Next
            lbxMessageList.DataSource = messageList
            lbxMessageList.DisplayMember = "MessageName"
        End If

    End Sub


    ' 斷線
    Private Sub btnDisconnect_Click(sender As Object, e As EventArgs) Handles btnDisconnect.Click

        StopSend = True
        aWrapper.Close()

    End Sub


    ' Send Message
    Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click

        startSend = New Thread(AddressOf SendMessage)
        startSend.IsBackground = True
        startSend.Start()

    End Sub

    ' Stop Send Message
    Private Sub btnStopSend_Click(sender As Object, e As EventArgs) Handles btnStopSend.Click

        StopSend = True

    End Sub


    ' S9F1
    Private Sub btnS9F1Event_Click(sender As Object, e As EventArgs) Handles btnS9F1Event.Click

        Dim temp As SecsMessage = aWrapper.GetMessageByName("S1F1_AreYouThere")
        temp.messageFormat.DeviceID(0) = 99
        temp.messageFormat.DeviceID(1) = 99
        aWrapper.SendS9F1Message(temp)

    End Sub


    ' S9F3
    Private Sub btnS9F3Event_Click(sender As Object, e As EventArgs) Handles btnS9F3Event.Click

        Dim temp As SecsMessage = aWrapper.GetMessageByName("S1F1_AreYouThere")
        temp.MessageName = "S99F99_ErrorMessage"
        temp.Stream = Convert.ToUInt32(99)
        aWrapper.SendPrimary(temp)

    End Sub


    ' S9F5
    Private Sub btnS9F5Event_Click(sender As Object, e As EventArgs) Handles btnS9F5Event.Click

        Dim temp As SecsMessage = aWrapper.GetMessageByName("S1F1_AreYouThere")
        temp.MessageName = "S1F99_ErrorMessage"
        temp.Function = Convert.ToUInt32(99)
        aWrapper.SendPrimary(temp)

    End Sub


    ' S9F7
    Private Sub btnS9F7Event_Click(sender As Object, e As EventArgs) Handles btnS9F7Event.Click

        Dim temp As SecsMessage = aWrapper.GetMessageByName("S1F3_SelectedEquipmentStatusRequest")

        temp.secsItem.AddItem()
        temp.secsItem.ItemList(0).ItemType = "A"
        temp.secsItem.ItemList(0).SetItemValue("abcdefg")

        'temp.secsItem.InsertItem(1)
        'temp.secsItem.ItemList(1).ItemType = "Boolean"
        'Dim tempBoolean As Boolean() = New Boolean(1) {True, False}
        'temp.secsItem.ItemList(1).SetItemValue(tempBoolean)

        'temp.secsItem.DeleteItem(1)

        'Dim temptest As UInt16 = 1
        'temp.secsItem.GetItemByName("SVID").SetItemValue(temptest)

        aWrapper.SendPrimary(temp)

    End Sub


    ' T8 Timeout
    Private Sub btnT8_Click(sender As Object, e As EventArgs) Handles btnT8.Click

        Dim temp As SecsMessage = aWrapper.GetMessageByName("S1F1_AreYouThere")
        aWrapper.SendT8TimeoutMessage(temp)

    End Sub


    ' Hide
    Private Sub btnHide_Click(sender As Object, e As EventArgs) Handles btnHide.Click

        If ShowTextBox = True Then

            ShowTextBox = False
            RichTextBox1.Visible = False
        Else
            ShowTextBox = True
            RichTextBox1.Visible = True
        End If

    End Sub


    ' Function
    Private Sub btnFunction_Click(sender As Object, e As EventArgs) Handles btnFunction.Click

        If ShowFunction = True Then

            ShowFunction = False
            btnFunction.Text = "Close"
            Me.Size = New Size(550, 700)

        Else
            ShowFunction = True
            btnFunction.Text = "Open"
            Me.Size = New Size(1200, 700)

        End If

    End Sub


    ' Message SML
    Private Sub lbxMessageList_Click(sender As Object, e As EventArgs) Handles lbxMessageList.Click

        secsItemList.Clear()

        Dim Message As ArrayList = New ArrayList()

        SecsMessage = aWrapper.GetMessageByName(lbxMessageList.SelectedItem)

        Dim tempSML As String = SecsMessage.ConvertToSML("PrimaryIn")

        Dim tempSplitList As String() = tempSML.Split(ControlChars.CrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        Dim SplitList As List(Of String) = New List(Of String)

        For i As Integer = 0 To tempSplitList.Length - 1

            If tempSplitList(i).Contains("<") = True Then

                SplitList.Add(tempSplitList(i))
            End If
        Next

        For i As Integer = 0 To SplitList.Count - 1

            Message.Add(SplitList(i))
        Next

        If SecsMessage.secsItem IsNot Nothing Then

            AddSecsItem(SecsMessage.secsItem)
        End If

        lbxMessage.DataSource = Message

        tbxItemType2.Text = ""
        tbxItemValue2.Text = ""

    End Sub

    Private Sub AddSecsItem(ByRef aSecsItem As SecsItem)

        If aSecsItem.ItemType = "L" Then

            secsItemList.Add(aSecsItem)

            For i As Integer = 0 To aSecsItem.ItemList.Count - 1
                AddSecsItem(aSecsItem.ItemList(i))
            Next
        Else
            secsItemList.Add(aSecsItem)
        End If

    End Sub


    ' Message Item Type、Item Value
    Private Sub lbxMessage_Click(sender As Object, e As EventArgs) Handles lbxMessage.Click

        tbxItemType2.Text = secsItemList(lbxMessage.SelectedIndex).ItemType
        tbxItemValue2.Text = secsItemList(lbxMessage.SelectedIndex).ItemValueString

    End Sub


    Private Sub btnModify_Click(sender As Object, e As EventArgs) Handles btnModify.Click

        If tbxItemType2.Text = "" Or tbxItemValue2.Text = "" Or lbxMessage Is Nothing Then

            ' Do Nothing
        Else
            secsItemList(lbxMessage.SelectedIndex).ItemType = tbxItemType2.Text
            secsItemList(lbxMessage.SelectedIndex).SetItemValue(tbxItemValue2.Text.Split(","))
        End If

        secsItemList.Clear()

        Dim Message As ArrayList = New ArrayList()

        Dim tempSML As String = SecsMessage.ConvertToSML("PrimaryIn")

        Dim tempSplitList As String() = tempSML.Split(ControlChars.CrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        Dim SplitList As List(Of String) = New List(Of String)

        For i As Integer = 0 To tempSplitList.Length - 1

            If tempSplitList(i).Contains("<") = True Then

                SplitList.Add(tempSplitList(i))
            End If
        Next

        For i As Integer = 0 To SplitList.Count - 1

            Message.Add(SplitList(i))
        Next

        AddSecsItem(SecsMessage.secsItem)

        lbxMessage.DataSource = Message

        tbxItemType2.Text = ""
        tbxItemValue2.Text = ""

    End Sub


    Private Sub btnAddItem_Click(sender As Object, e As EventArgs) Handles btnAddItem.Click

        If tbxItemType.Text = "" Or tbxItemValue.Text = "" Or lbxMessage Is Nothing Then

            ' Do Nothing
        Else
            secsItemList(lbxMessage.SelectedIndex).AddItem()

            If secsItemList(lbxMessage.SelectedIndex).ItemType = "L" Then
                secsItemList(lbxMessage.SelectedIndex).ItemList(secsItemList(lbxMessage.SelectedIndex).ItemList.Count - 1).ItemType = tbxItemType.Text
                secsItemList(lbxMessage.SelectedIndex).ItemList(secsItemList(lbxMessage.SelectedIndex).ItemList.Count - 1).SetItemValue(tbxItemValue.Text.Split(","))
            Else
                secsItemList(lbxMessage.SelectedIndex).ItemList(0).ItemType = tbxItemType.Text
                secsItemList(lbxMessage.SelectedIndex).ItemList(0).SetItemValue(tbxItemValue.Text.Split(","))
            End If


        End If

        secsItemList.Clear()

        Dim Message As ArrayList = New ArrayList()

        Dim tempSML As String = SecsMessage.ConvertToSML("PrimaryIn")

        Dim tempSplitList As String() = tempSML.Split(ControlChars.CrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        Dim SplitList As List(Of String) = New List(Of String)

        For i As Integer = 0 To tempSplitList.Length - 1

            If tempSplitList(i).Contains("<") = True Then

                SplitList.Add(tempSplitList(i))
            End If
        Next

        For i As Integer = 0 To SplitList.Count - 1

            Message.Add(SplitList(i))
        Next

        AddSecsItem(SecsMessage.secsItem)

        lbxMessage.DataSource = Message


    End Sub


    Private Sub btnInsertItem_Click(sender As Object, e As EventArgs) Handles btnInsertItem.Click

        If tbxItemType.Text = "" Or tbxItemValue.Text = "" Or lbxMessage Is Nothing Or tbxInsertIndex.Text = "" Then

            ' Do Nothing
        Else
            Dim index As UInteger = Convert.ToUInt16(tbxInsertIndex.Text)


            If index > secsItemList(lbxMessage.SelectedIndex).ItemList.Count Then
                Return
            End If

            If secsItemList(lbxMessage.SelectedIndex).ItemType = "L" Then
                secsItemList(lbxMessage.SelectedIndex).InsertItem(index)
                secsItemList(lbxMessage.SelectedIndex).ItemList(index).ItemType = tbxItemType.Text
                secsItemList(lbxMessage.SelectedIndex).ItemList(index).SetItemValue(tbxItemValue.Text.Split(","))
            Else
                secsItemList(lbxMessage.SelectedIndex).AddItem()
                secsItemList(lbxMessage.SelectedIndex).ItemList(0).ItemType = tbxItemType.Text
                secsItemList(lbxMessage.SelectedIndex).ItemList(0).SetItemValue(tbxItemValue.Text.Split(","))
            End If

        End If

        secsItemList.Clear()

        Dim Message As ArrayList = New ArrayList()

        Dim tempSML As String = SecsMessage.ConvertToSML("PrimaryIn")

        Dim tempSplitList As String() = tempSML.Split(ControlChars.CrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        Dim SplitList As List(Of String) = New List(Of String)

        For i As Integer = 0 To tempSplitList.Length - 1

            If tempSplitList(i).Contains("<") = True Then

                SplitList.Add(tempSplitList(i))
            End If
        Next

        For i As Integer = 0 To SplitList.Count - 1

            Message.Add(SplitList(i))
        Next

        AddSecsItem(SecsMessage.secsItem)

        lbxMessage.DataSource = Message

    End Sub


    Private Sub btnDeleteItem_Click(sender As Object, e As EventArgs) Handles btnDeleteItem.Click

        If tbxInsertIndex.Text = "" Then

            ' Do Nothing
        Else
            Dim index As UInteger = Convert.ToUInt16(tbxInsertIndex.Text)

            If secsItemList(lbxMessage.SelectedIndex).ItemType = "L" Then
                secsItemList(lbxMessage.SelectedIndex).DeleteItem(index)
            Else
                ' Do Nothing
            End If

        End If

        secsItemList.Clear()

        Dim Message As ArrayList = New ArrayList()

        Dim tempSML As String = SecsMessage.ConvertToSML("PrimaryIn")

        Dim tempSplitList As String() = tempSML.Split(ControlChars.CrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        Dim SplitList As List(Of String) = New List(Of String)

        For i As Integer = 0 To tempSplitList.Length - 1

            If tempSplitList(i).Contains("<") = True Then

                SplitList.Add(tempSplitList(i))
            End If
        Next

        For i As Integer = 0 To SplitList.Count - 1

            Message.Add(SplitList(i))
        Next

        AddSecsItem(SecsMessage.secsItem)

        lbxMessage.DataSource = Message

    End Sub


    Private Sub btnSend2_Click(sender As Object, e As EventArgs) Handles btnSend2.Click

        aWrapper.SendPrimary(SecsMessage)

    End Sub


    '-----------------------------------------------------------------------------------------------------

    Private Sub SendMessage()

        StopSend = False

        Dim sendcount As Integer
        Dim sleepTime As Integer = Convert.ToInt16(tbxThreadSleep.Text)

        If tbxTotalSend.Text <> "" Then

            sendcount = Convert.ToInt64(tbxTotalSend.Text)

        End If

        Do
            If tbxTotalSend.Text = "" Then

                Dim temp As SecsMessage = aWrapper.GetMessageByName(random())
                aWrapper.SendPrimary(temp)

                Dim tempTime As Integer = Convert.ToInt16(tbxThreadSleep.Text)
                Thread.Sleep(tempTime)

            ElseIf sendcount = 0 Then

                StopSend = True

            Else

                Dim temp As SecsMessage = aWrapper.GetMessageByName(random())
                aWrapper.SendPrimary(temp)

                sendcount = sendcount - 1

                Thread.Sleep(sleepTime)

            End If

        Loop Until StopSend = True

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
        'temp.Add("S6F11_ActiveCarrierReq")
        'temp.Add("S6F11_Alarm")
        'temp.Add("S6F11_AlarmClear")
        'temp.Add("S6F11_AlarmSet")

        Dim a As Random = New Random
        Dim b As Integer = a.Next(0, temp.Count)
        Return temp(b)

    End Function


    Public Sub SecsConnect(ByRef aMessage As String)

        If ShowTextBox = True Then

            Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
            Invoke(temp, "Secs Connect State : " & aMessage, RichTextBox1)

        End If

        If aMessage = "NotConnected" Then
            StopSend = True
        End If

    End Sub


    Public Sub PrimaryReceive(ByRef aTransaction As SecsTransaction)

        If ShowTextBox = True Then

            Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
            Invoke(temp, "Primary Receive : " & aTransaction.Primary.MessageName, RichTextBox1)

        End If

        If aTransaction.Primary.MessageName IsNot Nothing Then

            aPrimaryReceived = aPrimaryReceived + 1
            Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
            Invoke(temp, "PrimaryReceived", lblPrimaryReceived)

        End If

    End Sub


    Public Sub PrimarySent(ByRef aTransaction As SecsTransaction)

        If ShowTextBox = True Then

            Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
            Invoke(temp, "Primary Sent : " & aTransaction.Primary.MessageName, RichTextBox1)

        End If

        If aTransaction.Primary.MessageName IsNot Nothing Then

            aPrimarySent = aPrimarySent + 1
            Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
            Invoke(temp, "PrimarySent", lblPrimarySent)

        End If

    End Sub


    Public Sub SecondaryReceive(ByRef aTransaction As SecsTransaction)

        If ShowTextBox = True Then

            Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
            Invoke(temp, "Secondary Receive : " & aTransaction.Secondary.MessageName, RichTextBox1)

        End If

        If aTransaction.Secondary.MessageName IsNot Nothing Then

            aSecondaryReceived = aSecondaryReceived + 1
            Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
            Invoke(temp, "SecondaryReceived", lblSecondaryReceived)

        End If

    End Sub


    Public Sub SecondarySent(ByRef aTransaction As SecsTransaction)

        If ShowTextBox = True Then

            Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
            Invoke(temp, "Secondary Sent : " & aTransaction.Secondary.MessageName, RichTextBox1)

        End If

        If aTransaction.Secondary.MessageName IsNot Nothing Then

            aSecondarySent = aSecondarySent + 1
            Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
            Invoke(temp, "SecondarySent", lblSecondarySent)

        End If

    End Sub


    Public Sub ErrorMessage(ByRef message As SecsMessage, ByVal errorMessage As String)

        If ShowTextBox = True Then

            Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)

            If message.MessageName Is Nothing Then
                Invoke(temp, "Control Message " & errorMessage, RichTextBox1)
            Else
                Invoke(temp, message.MessageName & " : " & errorMessage, RichTextBox1)
            End If

        End If

    End Sub


    Public Sub MessageInfo(ByRef message As SecsMessage, ByVal aMessage As String)

        If ShowTextBox = True Then

            Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
            Invoke(temp, message.MessageName & " : " & aMessage, RichTextBox1)

        End If

    End Sub


    Public Sub SecsError(ByRef aMessage As String)

        If ShowTextBox = True Then

            Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
            Invoke(temp, aMessage, RichTextBox1)

        End If

    End Sub


    Public Sub HandleTimeout(ByRef aMessage As String)

        If ShowTextBox = True Then

            Dim temp As UpdateGUI = New UpdateGUI(AddressOf GUIUpdate)
            Invoke(temp, aMessage, RichTextBox1)

        End If

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

        If RichTextBox1.TextLength > 1600 Then

            Dim temp As CleanGUI = New CleanGUI(AddressOf GUIClean)
            Invoke(temp, RichTextBox1)

        End If

    End Sub


    Public Sub GUIClean(ByRef control As Control)

        RichTextBox1.Text = RichTextBox1.Text.Substring(RichTextBox1.Lines(0).Length + 1)

    End Sub


End Class
