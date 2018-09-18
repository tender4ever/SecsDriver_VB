<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnConnect = New System.Windows.Forms.Button()
        Me.btnDisconnect = New System.Windows.Forms.Button()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.btnSend = New System.Windows.Forms.Button()
        Me.lblThreadSleep = New System.Windows.Forms.Label()
        Me.lblEntity = New System.Windows.Forms.Label()
        Me.tbxThreadSleep = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblPrimarySent = New System.Windows.Forms.Label()
        Me.lblSecondaryReceived = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblPrimaryReceived = New System.Windows.Forms.Label()
        Me.lblSecondarySent = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tbxTotalSend = New System.Windows.Forms.TextBox()
        Me.btnStopSend = New System.Windows.Forms.Button()
        Me.btnS9F3Event = New System.Windows.Forms.Button()
        Me.btnS9F5Event = New System.Windows.Forms.Button()
        Me.btnS9F7Event = New System.Windows.Forms.Button()
        Me.btnHide = New System.Windows.Forms.Button()
        Me.btnS9F1Event = New System.Windows.Forms.Button()
        Me.btnT8 = New System.Windows.Forms.Button()
        Me.lblIP = New System.Windows.Forms.Label()
        Me.lblPort = New System.Windows.Forms.Label()
        Me.btnT3 = New System.Windows.Forms.Button()
        Me.btnT6 = New System.Windows.Forms.Button()
        Me.btnFunction = New System.Windows.Forms.Button()
        Me.lbxMessageList = New System.Windows.Forms.ListBox()
        Me.lbxMessage = New System.Windows.Forms.ListBox()
        Me.lblItemType = New System.Windows.Forms.Label()
        Me.tbxItemType = New System.Windows.Forms.TextBox()
        Me.lblItmeValue = New System.Windows.Forms.Label()
        Me.tbxItemValue = New System.Windows.Forms.TextBox()
        Me.btnAddItem = New System.Windows.Forms.Button()
        Me.btnInsertItem = New System.Windows.Forms.Button()
        Me.btnDeleteItem = New System.Windows.Forms.Button()
        Me.btnSend2 = New System.Windows.Forms.Button()
        Me.tbxInsertIndex = New System.Windows.Forms.TextBox()
        Me.lblMessageList = New System.Windows.Forms.Label()
        Me.lblSML = New System.Windows.Forms.Label()
        Me.lblItemType2 = New System.Windows.Forms.Label()
        Me.tbxItemType2 = New System.Windows.Forms.TextBox()
        Me.lblItemValue2 = New System.Windows.Forms.Label()
        Me.tbxItemValue2 = New System.Windows.Forms.TextBox()
        Me.btnModify = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnConnect
        '
        Me.btnConnect.Location = New System.Drawing.Point(12, 29)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(75, 23)
        Me.btnConnect.TabIndex = 6
        Me.btnConnect.Text = "連線"
        Me.btnConnect.UseVisualStyleBackColor = True
        '
        'btnDisconnect
        '
        Me.btnDisconnect.Location = New System.Drawing.Point(12, 58)
        Me.btnDisconnect.Name = "btnDisconnect"
        Me.btnDisconnect.Size = New System.Drawing.Size(75, 23)
        Me.btnDisconnect.TabIndex = 7
        Me.btnDisconnect.Text = "斷線"
        Me.btnDisconnect.UseVisualStyleBackColor = True
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(12, 243)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(510, 400)
        Me.RichTextBox1.TabIndex = 8
        Me.RichTextBox1.Text = ""
        '
        'btnSend
        '
        Me.btnSend.Location = New System.Drawing.Point(108, 29)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.Size = New System.Drawing.Size(100, 23)
        Me.btnSend.TabIndex = 9
        Me.btnSend.Text = "Send Message"
        Me.btnSend.UseVisualStyleBackColor = True
        '
        'lblThreadSleep
        '
        Me.lblThreadSleep.AutoSize = True
        Me.lblThreadSleep.Location = New System.Drawing.Point(225, 34)
        Me.lblThreadSleep.Name = "lblThreadSleep"
        Me.lblThreadSleep.Size = New System.Drawing.Size(66, 12)
        Me.lblThreadSleep.TabIndex = 10
        Me.lblThreadSleep.Text = "Thread Sleep"
        '
        'lblEntity
        '
        Me.lblEntity.AutoSize = True
        Me.lblEntity.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblEntity.Location = New System.Drawing.Point(12, 9)
        Me.lblEntity.Name = "lblEntity"
        Me.lblEntity.Size = New System.Drawing.Size(39, 12)
        Me.lblEntity.TabIndex = 11
        Me.lblEntity.Text = "Entity"
        '
        'tbxThreadSleep
        '
        Me.tbxThreadSleep.Location = New System.Drawing.Point(297, 31)
        Me.tbxThreadSleep.Name = "tbxThreadSleep"
        Me.tbxThreadSleep.Size = New System.Drawing.Size(50, 22)
        Me.tbxThreadSleep.TabIndex = 12
        Me.tbxThreadSleep.Text = "500"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 95)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 12)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "PrimarySent :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 125)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(103, 12)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "SecondaryReceived :"
        '
        'lblPrimarySent
        '
        Me.lblPrimarySent.AutoSize = True
        Me.lblPrimarySent.ForeColor = System.Drawing.Color.Red
        Me.lblPrimarySent.Location = New System.Drawing.Point(126, 95)
        Me.lblPrimarySent.Name = "lblPrimarySent"
        Me.lblPrimarySent.Size = New System.Drawing.Size(11, 12)
        Me.lblPrimarySent.TabIndex = 15
        Me.lblPrimarySent.Text = "0"
        '
        'lblSecondaryReceived
        '
        Me.lblSecondaryReceived.AutoSize = True
        Me.lblSecondaryReceived.ForeColor = System.Drawing.Color.Red
        Me.lblSecondaryReceived.Location = New System.Drawing.Point(126, 125)
        Me.lblSecondaryReceived.Name = "lblSecondaryReceived"
        Me.lblSecondaryReceived.Size = New System.Drawing.Size(11, 12)
        Me.lblSecondaryReceived.TabIndex = 16
        Me.lblSecondaryReceived.Text = "0"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 157)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(91, 12)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "PrimaryReceived :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 188)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(83, 12)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "SecondarySent : "
        '
        'lblPrimaryReceived
        '
        Me.lblPrimaryReceived.AutoSize = True
        Me.lblPrimaryReceived.ForeColor = System.Drawing.Color.Red
        Me.lblPrimaryReceived.Location = New System.Drawing.Point(126, 157)
        Me.lblPrimaryReceived.Name = "lblPrimaryReceived"
        Me.lblPrimaryReceived.Size = New System.Drawing.Size(11, 12)
        Me.lblPrimaryReceived.TabIndex = 19
        Me.lblPrimaryReceived.Text = "0"
        '
        'lblSecondarySent
        '
        Me.lblSecondarySent.AutoSize = True
        Me.lblSecondarySent.ForeColor = System.Drawing.Color.Red
        Me.lblSecondarySent.Location = New System.Drawing.Point(126, 188)
        Me.lblSecondarySent.Name = "lblSecondarySent"
        Me.lblSecondarySent.Size = New System.Drawing.Size(11, 12)
        Me.lblSecondarySent.TabIndex = 20
        Me.lblSecondarySent.Text = "0"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(225, 63)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 12)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "Total Send"
        '
        'tbxTotalSend
        '
        Me.tbxTotalSend.Location = New System.Drawing.Point(297, 60)
        Me.tbxTotalSend.Name = "tbxTotalSend"
        Me.tbxTotalSend.Size = New System.Drawing.Size(50, 22)
        Me.tbxTotalSend.TabIndex = 22
        Me.tbxTotalSend.Text = "0"
        '
        'btnStopSend
        '
        Me.btnStopSend.Location = New System.Drawing.Point(108, 58)
        Me.btnStopSend.Name = "btnStopSend"
        Me.btnStopSend.Size = New System.Drawing.Size(100, 23)
        Me.btnStopSend.TabIndex = 23
        Me.btnStopSend.Text = "Stop Send"
        Me.btnStopSend.UseVisualStyleBackColor = True
        '
        'btnS9F3Event
        '
        Me.btnS9F3Event.Location = New System.Drawing.Point(227, 129)
        Me.btnS9F3Event.Name = "btnS9F3Event"
        Me.btnS9F3Event.Size = New System.Drawing.Size(90, 23)
        Me.btnS9F3Event.TabIndex = 24
        Me.btnS9F3Event.Text = "S9F3 Event"
        Me.btnS9F3Event.UseVisualStyleBackColor = True
        '
        'btnS9F5Event
        '
        Me.btnS9F5Event.Location = New System.Drawing.Point(227, 162)
        Me.btnS9F5Event.Name = "btnS9F5Event"
        Me.btnS9F5Event.Size = New System.Drawing.Size(90, 23)
        Me.btnS9F5Event.TabIndex = 25
        Me.btnS9F5Event.Text = "S9F5 Event"
        Me.btnS9F5Event.UseVisualStyleBackColor = True
        '
        'btnS9F7Event
        '
        Me.btnS9F7Event.Location = New System.Drawing.Point(227, 196)
        Me.btnS9F7Event.Name = "btnS9F7Event"
        Me.btnS9F7Event.Size = New System.Drawing.Size(90, 23)
        Me.btnS9F7Event.TabIndex = 26
        Me.btnS9F7Event.Text = "S9F7 Event"
        Me.btnS9F7Event.UseVisualStyleBackColor = True
        '
        'btnHide
        '
        Me.btnHide.Location = New System.Drawing.Point(12, 215)
        Me.btnHide.Name = "btnHide"
        Me.btnHide.Size = New System.Drawing.Size(40, 23)
        Me.btnHide.TabIndex = 28
        Me.btnHide.Text = "Hide"
        Me.btnHide.UseVisualStyleBackColor = True
        '
        'btnS9F1Event
        '
        Me.btnS9F1Event.Location = New System.Drawing.Point(227, 95)
        Me.btnS9F1Event.Name = "btnS9F1Event"
        Me.btnS9F1Event.Size = New System.Drawing.Size(90, 23)
        Me.btnS9F1Event.TabIndex = 29
        Me.btnS9F1Event.Text = "S9F1 Event"
        Me.btnS9F1Event.UseVisualStyleBackColor = True
        '
        'btnT8
        '
        Me.btnT8.Location = New System.Drawing.Point(323, 162)
        Me.btnT8.Name = "btnT8"
        Me.btnT8.Size = New System.Drawing.Size(90, 23)
        Me.btnT8.TabIndex = 30
        Me.btnT8.Text = "T8 Timeout"
        Me.btnT8.UseVisualStyleBackColor = True
        '
        'lblIP
        '
        Me.lblIP.AutoSize = True
        Me.lblIP.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblIP.Location = New System.Drawing.Point(106, 9)
        Me.lblIP.Name = "lblIP"
        Me.lblIP.Size = New System.Drawing.Size(17, 12)
        Me.lblIP.TabIndex = 31
        Me.lblIP.Text = "IP"
        '
        'lblPort
        '
        Me.lblPort.AutoSize = True
        Me.lblPort.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblPort.Location = New System.Drawing.Point(225, 9)
        Me.lblPort.Name = "lblPort"
        Me.lblPort.Size = New System.Drawing.Size(28, 12)
        Me.lblPort.TabIndex = 32
        Me.lblPort.Text = "Port"
        '
        'btnT3
        '
        Me.btnT3.Location = New System.Drawing.Point(323, 95)
        Me.btnT3.Name = "btnT3"
        Me.btnT3.Size = New System.Drawing.Size(90, 23)
        Me.btnT3.TabIndex = 33
        Me.btnT3.Text = "T3 Timeout"
        Me.btnT3.UseVisualStyleBackColor = True
        '
        'btnT6
        '
        Me.btnT6.Location = New System.Drawing.Point(323, 129)
        Me.btnT6.Name = "btnT6"
        Me.btnT6.Size = New System.Drawing.Size(90, 23)
        Me.btnT6.TabIndex = 34
        Me.btnT6.Text = "T6 Timeout"
        Me.btnT6.UseVisualStyleBackColor = True
        '
        'btnFunction
        '
        Me.btnFunction.Location = New System.Drawing.Point(447, 214)
        Me.btnFunction.Name = "btnFunction"
        Me.btnFunction.Size = New System.Drawing.Size(75, 23)
        Me.btnFunction.TabIndex = 35
        Me.btnFunction.Text = "Function"
        Me.btnFunction.UseVisualStyleBackColor = True
        '
        'lbxMessageList
        '
        Me.lbxMessageList.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lbxMessageList.FormattingEnabled = True
        Me.lbxMessageList.ItemHeight = 16
        Me.lbxMessageList.Location = New System.Drawing.Point(550, 28)
        Me.lbxMessageList.Name = "lbxMessageList"
        Me.lbxMessageList.Size = New System.Drawing.Size(422, 212)
        Me.lbxMessageList.TabIndex = 36
        '
        'lbxMessage
        '
        Me.lbxMessage.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lbxMessage.FormattingEnabled = True
        Me.lbxMessage.ItemHeight = 12
        Me.lbxMessage.Location = New System.Drawing.Point(550, 267)
        Me.lbxMessage.Name = "lbxMessage"
        Me.lbxMessage.Size = New System.Drawing.Size(622, 376)
        Me.lbxMessage.TabIndex = 37
        '
        'lblItemType
        '
        Me.lblItemType.AutoSize = True
        Me.lblItemType.Location = New System.Drawing.Point(978, 28)
        Me.lblItemType.Name = "lblItemType"
        Me.lblItemType.Size = New System.Drawing.Size(53, 12)
        Me.lblItemType.TabIndex = 38
        Me.lblItemType.Text = "Item Type"
        '
        'tbxItemType
        '
        Me.tbxItemType.Location = New System.Drawing.Point(1037, 24)
        Me.tbxItemType.Name = "tbxItemType"
        Me.tbxItemType.Size = New System.Drawing.Size(50, 22)
        Me.tbxItemType.TabIndex = 39
        '
        'lblItmeValue
        '
        Me.lblItmeValue.AutoSize = True
        Me.lblItmeValue.Location = New System.Drawing.Point(978, 58)
        Me.lblItmeValue.Name = "lblItmeValue"
        Me.lblItmeValue.Size = New System.Drawing.Size(56, 12)
        Me.lblItmeValue.TabIndex = 40
        Me.lblItmeValue.Text = "Item Value"
        '
        'tbxItemValue
        '
        Me.tbxItemValue.Location = New System.Drawing.Point(1037, 52)
        Me.tbxItemValue.Name = "tbxItemValue"
        Me.tbxItemValue.Size = New System.Drawing.Size(140, 22)
        Me.tbxItemValue.TabIndex = 41
        '
        'btnAddItem
        '
        Me.btnAddItem.Location = New System.Drawing.Point(1012, 84)
        Me.btnAddItem.Name = "btnAddItem"
        Me.btnAddItem.Size = New System.Drawing.Size(75, 23)
        Me.btnAddItem.TabIndex = 42
        Me.btnAddItem.Text = "Add Item"
        Me.btnAddItem.UseVisualStyleBackColor = True
        '
        'btnInsertItem
        '
        Me.btnInsertItem.Location = New System.Drawing.Point(1012, 115)
        Me.btnInsertItem.Name = "btnInsertItem"
        Me.btnInsertItem.Size = New System.Drawing.Size(75, 23)
        Me.btnInsertItem.TabIndex = 43
        Me.btnInsertItem.Text = "Insert Item"
        Me.btnInsertItem.UseVisualStyleBackColor = True
        '
        'btnDeleteItem
        '
        Me.btnDeleteItem.Location = New System.Drawing.Point(1093, 115)
        Me.btnDeleteItem.Name = "btnDeleteItem"
        Me.btnDeleteItem.Size = New System.Drawing.Size(75, 23)
        Me.btnDeleteItem.TabIndex = 44
        Me.btnDeleteItem.Text = "Delete Item"
        Me.btnDeleteItem.UseVisualStyleBackColor = True
        '
        'btnSend2
        '
        Me.btnSend2.Location = New System.Drawing.Point(1097, 238)
        Me.btnSend2.Name = "btnSend2"
        Me.btnSend2.Size = New System.Drawing.Size(75, 23)
        Me.btnSend2.TabIndex = 45
        Me.btnSend2.Text = "Send"
        Me.btnSend2.UseVisualStyleBackColor = True
        '
        'tbxInsertIndex
        '
        Me.tbxInsertIndex.Location = New System.Drawing.Point(981, 115)
        Me.tbxInsertIndex.Name = "tbxInsertIndex"
        Me.tbxInsertIndex.Size = New System.Drawing.Size(25, 22)
        Me.tbxInsertIndex.TabIndex = 46
        '
        'lblMessageList
        '
        Me.lblMessageList.AutoSize = True
        Me.lblMessageList.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblMessageList.Location = New System.Drawing.Point(547, 9)
        Me.lblMessageList.Name = "lblMessageList"
        Me.lblMessageList.Size = New System.Drawing.Size(101, 16)
        Me.lblMessageList.TabIndex = 47
        Me.lblMessageList.Text = "Message List"
        '
        'lblSML
        '
        Me.lblSML.AutoSize = True
        Me.lblSML.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblSML.Location = New System.Drawing.Point(547, 248)
        Me.lblSML.Name = "lblSML"
        Me.lblSML.Size = New System.Drawing.Size(41, 16)
        Me.lblSML.TabIndex = 48
        Me.lblSML.Text = "SML"
        '
        'lblItemType2
        '
        Me.lblItemType2.AutoSize = True
        Me.lblItemType2.Location = New System.Drawing.Point(978, 173)
        Me.lblItemType2.Name = "lblItemType2"
        Me.lblItemType2.Size = New System.Drawing.Size(53, 12)
        Me.lblItemType2.TabIndex = 49
        Me.lblItemType2.Text = "Item Type"
        '
        'tbxItemType2
        '
        Me.tbxItemType2.Location = New System.Drawing.Point(1037, 167)
        Me.tbxItemType2.Name = "tbxItemType2"
        Me.tbxItemType2.Size = New System.Drawing.Size(50, 22)
        Me.tbxItemType2.TabIndex = 50
        '
        'lblItemValue2
        '
        Me.lblItemValue2.AutoSize = True
        Me.lblItemValue2.Location = New System.Drawing.Point(978, 203)
        Me.lblItemValue2.Name = "lblItemValue2"
        Me.lblItemValue2.Size = New System.Drawing.Size(56, 12)
        Me.lblItemValue2.TabIndex = 51
        Me.lblItemValue2.Text = "Item Value"
        '
        'tbxItemValue2
        '
        Me.tbxItemValue2.Location = New System.Drawing.Point(1037, 198)
        Me.tbxItemValue2.Name = "tbxItemValue2"
        Me.tbxItemValue2.Size = New System.Drawing.Size(140, 22)
        Me.tbxItemValue2.TabIndex = 52
        '
        'btnModify
        '
        Me.btnModify.Location = New System.Drawing.Point(1097, 167)
        Me.btnModify.Name = "btnModify"
        Me.btnModify.Size = New System.Drawing.Size(75, 23)
        Me.btnModify.TabIndex = 53
        Me.btnModify.Text = "Modify"
        Me.btnModify.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1184, 662)
        Me.Controls.Add(Me.btnModify)
        Me.Controls.Add(Me.tbxItemValue2)
        Me.Controls.Add(Me.lblItemValue2)
        Me.Controls.Add(Me.tbxItemType2)
        Me.Controls.Add(Me.lblItemType2)
        Me.Controls.Add(Me.lblSML)
        Me.Controls.Add(Me.lblMessageList)
        Me.Controls.Add(Me.tbxInsertIndex)
        Me.Controls.Add(Me.btnSend2)
        Me.Controls.Add(Me.btnDeleteItem)
        Me.Controls.Add(Me.btnInsertItem)
        Me.Controls.Add(Me.btnAddItem)
        Me.Controls.Add(Me.tbxItemValue)
        Me.Controls.Add(Me.lblItmeValue)
        Me.Controls.Add(Me.tbxItemType)
        Me.Controls.Add(Me.lblItemType)
        Me.Controls.Add(Me.lbxMessage)
        Me.Controls.Add(Me.lbxMessageList)
        Me.Controls.Add(Me.btnFunction)
        Me.Controls.Add(Me.btnT6)
        Me.Controls.Add(Me.btnT3)
        Me.Controls.Add(Me.lblPort)
        Me.Controls.Add(Me.lblIP)
        Me.Controls.Add(Me.btnT8)
        Me.Controls.Add(Me.btnS9F1Event)
        Me.Controls.Add(Me.btnHide)
        Me.Controls.Add(Me.btnS9F7Event)
        Me.Controls.Add(Me.btnS9F5Event)
        Me.Controls.Add(Me.btnS9F3Event)
        Me.Controls.Add(Me.btnStopSend)
        Me.Controls.Add(Me.tbxTotalSend)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.lblSecondarySent)
        Me.Controls.Add(Me.lblPrimaryReceived)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lblSecondaryReceived)
        Me.Controls.Add(Me.lblPrimarySent)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbxThreadSleep)
        Me.Controls.Add(Me.lblEntity)
        Me.Controls.Add(Me.lblThreadSleep)
        Me.Controls.Add(Me.btnSend)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.btnDisconnect)
        Me.Controls.Add(Me.btnConnect)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnConnect As Button
    Friend WithEvents btnDisconnect As Button
    Friend WithEvents RichTextBox1 As RichTextBox
    Friend WithEvents btnSend As Button
    Friend WithEvents lblThreadSleep As Label
    Friend WithEvents lblEntity As Label
    Friend WithEvents tbxThreadSleep As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents lblPrimarySent As Label
    Friend WithEvents lblSecondaryReceived As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents lblPrimaryReceived As Label
    Friend WithEvents lblSecondarySent As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents tbxTotalSend As TextBox
    Friend WithEvents btnStopSend As Button
    Friend WithEvents btnS9F3Event As Button
    Friend WithEvents btnS9F5Event As Button
    Friend WithEvents btnS9F7Event As Button
    Friend WithEvents btnHide As Button
    Friend WithEvents btnS9F1Event As Button
    Friend WithEvents btnT8 As Button
    Friend WithEvents lblIP As Label
    Friend WithEvents lblPort As Label
    Friend WithEvents btnT3 As Button
    Friend WithEvents btnT6 As Button
    Friend WithEvents btnFunction As Button
    Friend WithEvents lbxMessageList As ListBox
    Friend WithEvents lbxMessage As ListBox
    Friend WithEvents lblItemType As Label
    Friend WithEvents tbxItemType As TextBox
    Friend WithEvents lblItmeValue As Label
    Friend WithEvents tbxItemValue As TextBox
    Friend WithEvents btnAddItem As Button
    Friend WithEvents btnInsertItem As Button
    Friend WithEvents btnDeleteItem As Button
    Friend WithEvents btnSend2 As Button
    Friend WithEvents tbxInsertIndex As TextBox
    Friend WithEvents lblMessageList As Label
    Friend WithEvents lblSML As Label
    Friend WithEvents lblItemType2 As Label
    Friend WithEvents tbxItemType2 As TextBox
    Friend WithEvents lblItemValue2 As Label
    Friend WithEvents tbxItemValue2 As TextBox
    Friend WithEvents btnModify As Button
End Class
