<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
	Inherits System.Windows.Forms.Form

	'Form 覆寫 Dispose 以清除元件清單。
	<System.Diagnostics.DebuggerNonUserCode()> _
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
	<System.Diagnostics.DebuggerStepThrough()> _
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
        Me.btnClean = New System.Windows.Forms.Button()
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
        Me.RichTextBox1.Location = New System.Drawing.Point(12, 159)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(620, 484)
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
        Me.lblThreadSleep.Location = New System.Drawing.Point(106, 97)
        Me.lblThreadSleep.Name = "lblThreadSleep"
        Me.lblThreadSleep.Size = New System.Drawing.Size(66, 12)
        Me.lblThreadSleep.TabIndex = 10
        Me.lblThreadSleep.Text = "Thread Sleep"
        '
        'lblEntity
        '
        Me.lblEntity.AutoSize = True
        Me.lblEntity.Location = New System.Drawing.Point(12, 9)
        Me.lblEntity.Name = "lblEntity"
        Me.lblEntity.Size = New System.Drawing.Size(37, 12)
        Me.lblEntity.TabIndex = 11
        Me.lblEntity.Text = "Label1"
        '
        'tbxThreadSleep
        '
        Me.tbxThreadSleep.Location = New System.Drawing.Point(178, 92)
        Me.tbxThreadSleep.Name = "tbxThreadSleep"
        Me.tbxThreadSleep.Size = New System.Drawing.Size(50, 22)
        Me.tbxThreadSleep.TabIndex = 12
        Me.tbxThreadSleep.Text = "500"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(248, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 12)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "PrimarySent :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(248, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(103, 12)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "SecondaryReceived :"
        '
        'lblPrimarySent
        '
        Me.lblPrimarySent.AutoSize = True
        Me.lblPrimarySent.ForeColor = System.Drawing.Color.Red
        Me.lblPrimarySent.Location = New System.Drawing.Point(397, 20)
        Me.lblPrimarySent.Name = "lblPrimarySent"
        Me.lblPrimarySent.Size = New System.Drawing.Size(11, 12)
        Me.lblPrimarySent.TabIndex = 15
        Me.lblPrimarySent.Text = "0"
        '
        'lblSecondaryReceived
        '
        Me.lblSecondaryReceived.AutoSize = True
        Me.lblSecondaryReceived.ForeColor = System.Drawing.Color.Red
        Me.lblSecondaryReceived.Location = New System.Drawing.Point(397, 55)
        Me.lblSecondaryReceived.Name = "lblSecondaryReceived"
        Me.lblSecondaryReceived.Size = New System.Drawing.Size(11, 12)
        Me.lblSecondaryReceived.TabIndex = 16
        Me.lblSecondaryReceived.Text = "0"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(248, 91)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(91, 12)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "PrimaryReceived :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(248, 126)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(83, 12)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "SecondarySent : "
        '
        'lblPrimaryReceived
        '
        Me.lblPrimaryReceived.AutoSize = True
        Me.lblPrimaryReceived.ForeColor = System.Drawing.Color.Red
        Me.lblPrimaryReceived.Location = New System.Drawing.Point(397, 90)
        Me.lblPrimaryReceived.Name = "lblPrimaryReceived"
        Me.lblPrimaryReceived.Size = New System.Drawing.Size(11, 12)
        Me.lblPrimaryReceived.TabIndex = 19
        Me.lblPrimaryReceived.Text = "0"
        '
        'lblSecondarySent
        '
        Me.lblSecondarySent.AutoSize = True
        Me.lblSecondarySent.ForeColor = System.Drawing.Color.Red
        Me.lblSecondarySent.Location = New System.Drawing.Point(397, 125)
        Me.lblSecondarySent.Name = "lblSecondarySent"
        Me.lblSecondarySent.Size = New System.Drawing.Size(11, 12)
        Me.lblSecondarySent.TabIndex = 20
        Me.lblSecondarySent.Text = "0"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(106, 132)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 12)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "Total Send"
        '
        'tbxTotalSend
        '
        Me.tbxTotalSend.Location = New System.Drawing.Point(178, 126)
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
        'btnClean
        '
        Me.btnClean.Location = New System.Drawing.Point(557, 132)
        Me.btnClean.Name = "btnClean"
        Me.btnClean.Size = New System.Drawing.Size(75, 23)
        Me.btnClean.TabIndex = 24
        Me.btnClean.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(644, 655)
        Me.Controls.Add(Me.btnClean)
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
    Friend WithEvents btnClean As Button
End Class
