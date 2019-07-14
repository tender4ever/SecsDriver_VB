# SECS-Driver

This program uses vb.net to practice SEMI standard.

SEMI standard defines the format of the message to pass between the host and the equipment

---

## SEMI Standard

| Standard | Name |
|----------|------|
| E23 | Cassette Transfer Parallel I/O Interface |
| E84 | Reinforce E23 
| E4 | SECS-I |
| E37 | HSMS |
| E5 | SECS-II |
| E30 | GEM |
| E82 | IB-SEM |
| E88 | STK-SEM |

---

## SXML 

The SXML file is defined the message for host and equipment communicating

### Format

```xml
<SXML>
	<DataDictionary>
	</DataDictionary>
	<MessageDictionary>
	
		<Message Name="S2F49_RemoteTransferCommand" Description="">					
			<Stream>2</Stream>
	        	<Function>49</Function>
			<WBit>1</WBit>
            		<AutoReply>S2F50_RemoteTransferCommandAck</AutoReply>
            		<Structure>
		        	<L Name="NameList" Description="Setting List">
	                		<A Name="NameASCII" Description="Setting ASCII">
                        			<Value></Value>
                        			<Maps Type="4096"></Maps>
                    			</A>
                    			<B Name="NameBinary" Description="Setting Binary">
                        			<Value>00</Value>
                        			<Maps Type="1024"></Maps>
                    			</B>
                    			<Boolean Name="NameBoolean" Description="Setting Boolean">
                        			<Value>True</Value>
                        			<Maps Type="2048"></Maps>
                    			</Boolean>
	                		<I1 Name="NameInt1" Description="Setting Int1">
		                		<Value>1</Value>
                        			<Maps Type="1"></Maps>
					</I1>
					<I2 Name="NameInt2" Description="Setting Int2">
		                		<Value></Value>
                        			<Maps Type="2"></Maps>
					</I2>
					<I4 Name="NameInt4" Description="Setting Int4">
		                		<Value></Value>
                        			<Maps Type="4"></Maps>
					</I4>
					<I8 Name="NameInt8" Description="Setting Int8">
		                		<Value></Value>
                        			<Maps Type="8"></Maps>
					</I8>
					<U1 Name="NameUInt1" Description="Setting UInt1">
	                    			<Value>11</Value>
	                    			<Maps Type="16"></Maps>
                    			</U1>
                    			<U2 Name="NameUInt2" Description="Setting UInt2">
                        			<Value>0</Value>
                        			<Maps Type="32"></Maps>
                    			</U2>
                    			<U4 Name="NameUInt4" Description="Setting UInt4">
	                    			<Value></Value>
	                    			<Maps Type="64"></Maps>
                    			</U4>
		            		<U8 Name="NameUInt8" Description="Setting UInt8">
	                    			<Value></Value>
	                    			<Maps Type="128"></Maps>
                    			</U8>
                		</L>
            		</Structure>
        	</Message>

	</MessageDictionary>
</SXML>
``` 

---

## Config.ini

The config.ini file is defined for SECS communication

The config.ini file, use **ANSI encode**

```xml
; Host = Active
; Equipment = Passive
Entity = 

; IP Address
IP =

; Port No
Port =

; Timeout Value For SECS Communication
T1 = 
T2 = 
T3 =
T4 = 
T5 =
T6 =
T7 =
T8 =
CA =

; SXML File Name
SXMLFIle =

; Device Name
DeviceName =

; Device ID
DeviceID =

; TCP/IP = HSMS
; RS232 = SECS-I
LinkType =

; Host = Host
; Equipment = Equipment 
Role =

; How many time For SECS to retry to connect
Retry =

; Log Setting
; Open Log = True
; Close Log = Flase
BinaryLog =
TxLog =

```

---

## Host to Use

```vb
Dim aWrapper As SecsWrapper             ' Secs Driver Wrapper

' New Secs Driver 物件
aWrapper = New SecsWrapper()

' New Secs Driver 委派
Dim sDelegateSecsConnect As New delegateSecsConnectStateChange(AddressOf SecsConnect)
Dim sDelegatePrimaryReceive As New delegatePrimaryReceived(AddressOf PrimaryReceive)
Dim sDelegatePrimarySent As New delegatePrimarySent(AddressOf PrimarySent)
Dim sDelegateSecondaryReceive As New delegateSecondaryReceived(AddressOf SecondaryReceive)
Dim sDelegateSecondarySent As New delegateSecondarySent(AddressOf SecondarySent)
Dim sDelegateMessageError As New delegateMessageError(AddressOf ErrorMessage)
Dim sDelegateMessageInfo As New delegateMessageInfo(AddressOf MessageInfo)
Dim sDelegateSecsError As New delegateError(AddressOf SecsError)
Dim sDelegateTimeout As New delegateTimeout(AddressOf HandleTimeout)

' 設定 Secs Driver Event 的處理
AddHandler aWrapper.OnSecsConnectStateChange, sDelegateSecsConnect
AddHandler aWrapper.OnPrimaryReceived, sDelegatePrimaryReceive
AddHandler aWrapper.OnPrimarySent, sDelegatePrimarySent
AddHandler aWrapper.OnSecondaryReceived, sDelegateSecondaryReceive
AddHandler aWrapper.OnSecondarySent, sDelegateSecondarySent
AddHandler aWrapper.OnMessageError, sDelegateMessageError
AddHandler aWrapper.OnMessageInfo, sDelegateMessageInfo
AddHandler aWrapper.OnError, sDelegateSecsError
AddHandler aWrapper.OnTimeout, sDelegateTimeout

' 開始連線
aWrapper.Open()

' Send SECS Message
Dim temp As SecsMessage = aWrapper.GetMessageByName("S1F1_AreYouThere")
aWrapper.SendPrimary(temp)

' 關閉連線
aWrapper.Close()
```

---

## SecsWrapper 物件

### Public Attribute

* siniFile : config.ini Object
* sSXML : SXML Object

### Public Method

Open()
> 開始 SECS 連線

Close()
> 關閉 SECS 連線

GetMessageByName(MessageName)
> 使用 Message Name 找 Message
> MessageName as String

SendPrimary(SecsMessage)
> Send Primary Message
> SecsMessage as SecsMessage Object

SendSecondary(SecsTransaction)
> Send Secondary Message
> SecsTransaction as SecsTransaction Object

---

## Secs Driver Tester

![](https://i.imgur.com/nmzxhje.png)

連線
> 會依照 config.ini 檔的設定，進行 SECS 連線

斷線
> 關閉 SECS 連線

Send Message
> 開始使用 Thread 方式持續丟 SecsMessage
> 比需先設定 **Thread Sleep**、**Total Send**

Stop Send
> 停止送出 SecsMessage

Hide
> 關閉即時訊息視窗

Function
> 開啟其他功能

S9F1 Event
> 送出會導致 S9F1 的訊息

S9F3
> 送出會導致 S9F3 的訊息

S9F5
> 送出會導致 S9F5 的訊息

S9F7
> 送出會導致 S9F7 的訊息

![](https://i.imgur.com/jmuOHUj.png)

Item Type
> 設定 Item Type

Item Value
> 設定 Item Value
> 可以設定成多組資料
> 如 A,B,C

Add Item
> 將設定的 SecsItem 加到目前所選取的位置上

Insert Item
> 將設定的 SecsItem，依照 index 設定加到位置上

Delete Item
> 刪除 index 設定上的 SecsItem

Modify
> 修改 SecsItem 的 ItemType、ItemValue

