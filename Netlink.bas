Attribute VB_Name = "Netlink"
Option Explicit
Enum NETSTATES
    NETLINK_DOWN
    NETLINK_CONNECTING
    NETLINK_VERIFYING
    NETLINK_UP
    NETLINK_ERROR
    End Enum

Enum NETACCESS
    ACCESS_DENIED
    ACCESS_IN
    ACCESS_OUT
    ACCESS_ALL
    End Enum
    
Type NL_OBJECT
    wasAutoConnected As Boolean
    southwest As Boolean          'is it from southwest?
    name As String
    password As String
    site As String
    port As Integer
    state As NETSTATES
    room As String
    version As String
    inpstr As String
    access As NETACCESS
    autoConnect As Boolean
    allowIn As Boolean
    allowOut As Boolean
    bytesIn As Long
    bytesOut As Long
    line As Integer
    End Type

Public Const MAX_NETLINKS = 20

Global connectTiming As Integer
Global madeNetlinks() As Boolean
Global net() As NL_OBJECT
Global netlinkStates(5) As String

Sub dropNetlink(line As Integer)
'this takes the actual socket as the argument
'***NOT*** the Netlink ID number for the array
Dim nn As Integer, count As Integer
nn = s2n(line)
net(nn).wasAutoConnected = False
For count = 1 To UBound(user)
    If user(count).atNetlink = line Then
        send "~OL~FRThe Netlink carrying you has been closed" & CRLF, count
        returnedFromNetlink count
        End If
    If user(count).netlinkType And user(count).netlinkFrom = line Then
        writeRoom "", "~FM" & user(count).name & " drifts back across the ethers" & CRLF
        removeUser (count)
        End If
    Next count
If line < LBound(madeNetlinks) Or line > UBound(madeNetlinks) Then
    net(line).state = NETLINK_DOWN
    Exit Sub
    End If
mainForm.Netlink(line).Flush
If mainForm.Netlink(line).Connected Then
    mainForm.Netlink(line).Disconnect
    End If
mainForm.Netlink(line).Flush
Unload mainForm.Netlink(line)
madeNetlinks(line) = False
If nn > -1 Then
    If net(nn).state > NETLINK_VERIFYING Then
        If net(nn).name = "" Then
            net(nn).access = ACCESS_DENIED
            Else
                writeSyslog "~FB" & net(nn).name & "~RS has disconnected"
                writeRoom "", "~OL~FYSYSTEM: ~RSDisconnecting from service " & net(nn).name & " in the " & net(nn).room & CRLF
                End If
        End If
    End If
If nn > -1 Then
    net(nn).inpstr = ""
    net(nn).bytesIn = 0
    net(nn).bytesOut = 0
    net(nn).state = NETLINK_DOWN
    net(nn).version = "3.3.3"
    net(nn).line = -1
    End If
mainForm.updateActiveNetlinks
updateNetstat
End Sub
Function loadNetlink() As Boolean
Dim linein As LOAD_OBJECT, FromFile As String
Dim parse As Boolean, count As Integer, file As String
lighter "Loading Netlinks"
netlinkStates(0) = "Down"           'These are the
netlinkStates(1) = "Connecting"     'names that users
netlinkStates(2) = "Verifying"      'will see when
netlinkStates(3) = "Operational"    'using the client.
netlinkStates(4) = "Closing"
netlinkStates(5) = "Error"


'Resize the Netlinks and MadeNetlinks arrays
count = 0
file = Dir$(App.Path & "\Netlinks\*.S")
Do While Not file = ""
    count = count + 1
    file = Dir$
    Loop
ReDim Preserve net(count), madeNetlinks(count)

For count = LBound(net) To UBound(net)
    If net(count).state > NETLINK_DOWN Then
        lighter "There are active Netlinks, cannot reload"
        loadNetlink = False
        Exit Function
        End If
    Next count
loadNetlink = True

For count = LBound(net) To UBound(net)
    With net(count)
        .name = ""
        .allowIn = False
        .allowOut = False
        .access = NETACCESS.ACCESS_DENIED
        .autoConnect = False
        .bytesIn = 0
        .bytesOut = 0
        .inpstr = ""
        .line = -1
        .room = ""
        .state = NETSTATES.NETLINK_DOWN
        .version = "3.3.3"
        .password = ""
        End With
    Next count
file = Dir(App.Path & "\Netlinks\*.S")
count = -1
Do While count < 21 And Not file = ""
    count = count + 1
    net(count).name = Left$(file, Len(file) - 2)
    Open App.Path & "\Netlinks\" & file For Input As #1
    Do While Not EOF(1)
        parse = True
        Line Input #1, FromFile
        FromFile = Replace(FromFile, Chr$(9), " ")
        FromFile = Trim$(FromFile)
        If FromFile = "" Then
            parse = False
            Else
                If Left$(FromFile, 1) = ";" Then
                    parse = False
                    End If
            End If
        If parse Then
            linein = spliceLoad(FromFile)
            Select Case linein.specifier
                Case "site"
                    net(count).site = linein.value
                Case "netport"
                    net(count).port = Int(linein.value)
                Case "password"
                    net(count).password = linein.value
                Case "autoconnect"
                    net(count).autoConnect = TF(linein.value)
                Case "incoming"
                    net(count).allowIn = TF(linein.value)
                Case "outgoing"
                    net(count).allowOut = TF(linein.value)
                Case "room"
                    net(count).room = linein.value
                    End Select
            End If
        Loop
    Close #1
    If net(count).allowIn Then
        net(count).access = NETACCESS.ACCESS_IN
        End If
    If net(count).allowOut Then
        If net(count).allowIn Then
            net(count).access = NETACCESS.ACCESS_ALL
            Else
                net(count).access = NETACCESS.ACCESS_OUT
                End If
        End If
    file = Dir
    Loop
If Not BOOTING Then
    loadViewer
    End If
updateNetstat
End Function

Sub netout(message As String, netnum As Integer)
Dim nn As Integer
nn = s2n(netnum)
If nn < 0 Then
    Exit Sub
    End If
If netnum < LBound(madeNetlinks) Or netnum > UBound(madeNetlinks) Then
    Exit Sub
    End If
If Len(message) > MAX_DATA_LEN Then
    message = Left(message, MAX_DATA_LEN)
    End If
mainForm.Netlink(netnum).Flush
message = Replace(message, CRLF, LF)
If madeNetlinks(netnum) Then
    mainForm.Netlink(netnum).SendLen = Len(message)
    net(nn).bytesOut = net(nn).bytesOut + mainForm.Netlink(netnum).SendLen
    If mainForm.Netlink(netnum).Connected Then
        On Error Resume Next
        mainForm.Netlink(netnum).SendData = message
        If Err Then
            dropNetlink netnum
            End If
        Else
            dropNetlink netnum
            End If
    End If
updateNetstat
'DoEvents   both good and bad may come from this
End Sub
Public Function NetToUser(UserName As String) As Integer
Dim count As Integer
For count = 1 To UBound(user)
    If user(count).name = UserName Then
        NetToUser = count
        Exit Function
        End If
    Next count
NetToUser = 0
End Function

Sub msgout(netnum As Integer, UserName As String, text As String)
Dim message As String
If Len(text) > MAX_DATA_LEN Then
    message = Left(message, MAX_DATA_LEN)
    End If
text = Replace(text, CRLF, LF)
text = "MSG " & UserName & LF & LF & text & " EMSG" & LF
If madeNetlinks(netnum) Then
    mainForm.Netlink(netnum).SendLen = Len(text)
    net(s2n(netnum)).bytesOut = net(s2n(netnum)).bytesOut + mainForm.Netlink(netnum).SendLen
    mainForm.Netlink(netnum).SendData = text
    End If
updateNetstat
End Sub

Public Sub updateNetstat()
If Not BOOTING Then
    If UBound(net) > 0 Then
        If mainForm.tree.Nodes("NETLINKS").Expanded Then
            If Not mainForm.tree.SelectedItem.Key = "NETLINKS" Then
                mainForm.treeLoad
                End If
            End If
        End If
    End If
End Sub

Public Function n2s(netnum As Integer) As Integer
n2s = net(netnum).line
End Function

Public Function s2n(socknum As Integer) As Integer
Dim count As Integer
s2n = -1
For count = LBound(net) To UBound(net)
    If net(count).line = socknum Then
        s2n = count
        Exit For
        End If
    Next count
End Function

Sub autoConnectNetlinks()
Dim count As Integer
For count = LBound(net) To UBound(net)
    If net(count).autoConnect Then
        
        End If
    Next count
End Sub

Function connectNetlink(netnum As Integer) As Boolean
'Finds an open spot for the Netlink and registers it
Dim sock As Integer, count As Integer, found As Boolean
If Not n2s(netnum) = -1 Or Not net(netnum).state = NETLINK_DOWN Then
    connectNetlink = False
    Exit Function
    End If
For sock = LBound(net) To UBound(net)
    If net(sock).line = -1 Then
        For count = 1 To UBound(madeNetlinks)
            If Not madeNetlinks(count) Then
                net(netnum).line = count
                madeNetlinks(count) = True
                Load mainForm.Netlink(count)
                found = True
                Exit For
                End If
            Next count
        If found Then
            Exit For
            End If
        End If
    Next sock
sock = n2s(netnum)
If sock = -1 Then
    connectNetlink = False
    Exit Function
    End If
net(netnum).state = NETLINK_CONNECTING
net(netnum).line = sock
madeNetlinks(sock) = True
mainForm.Netlink(sock).AddressFamily = AF_INET
mainForm.Netlink(sock).Protocol = IPPROTO_TCP
mainForm.Netlink(sock).SocketType = SOCK_STREAM
mainForm.Netlink(sock).Blocking = False
mainForm.Netlink(sock).LocalPort = 0
mainForm.Netlink(sock).HostName = net(netnum).site
mainForm.Netlink(sock).RemotePort = net(netnum).port
mainForm.Netlink(sock).Timeout = 1500
mainForm.Netlink(sock).Connect
updateNetstat
End Function

Sub netPacket(msg As String)
Dim lenToSpace As Integer, UserName As String, un As Integer
msg = Replace(msg, LF, " ", , 1)
msg = stripOne(msg) 'strip off MSG header word one
lenToSpace = InStr(msg, " ")
If lenToSpace > 1 Then ' get username
    UserName = Left$(msg, lenToSpace - 1)
    un = getUser(UserName)
    Else   ' it is an erroneous packet
        Exit Sub
        End If
If un <= 0 Then 'bad username
    Exit Sub
    End If
If Not Right$(msg, 5) = LF & "EMSG" Then 'bad packet
    Exit Sub
    End If
If Len(msg) > 5 Then
    msg = Left$(msg, Len(msg) - 4) & " EMSG"
    msg = stripOne(stripLast(msg)) 'expose payload
    msg = Replace(msg, LF, CRLF)
    End If
send msg, un
End Sub

Sub loadConnectNetlinks()
Dim count As Integer, found As Boolean
For count = LBound(net) To UBound(net)
    If net(count).autoConnect Then
        connectNetlink count
        net(count).wasAutoConnected = True
        found = True
        End If
    Next count
If found Then
    writeSyslog "Automaticlly connecting Netlinks"
    End If
End Sub

Function returnedFromNetlink(usernum As Integer)
writeRoom user(usernum).room, "~FM" & user(usernum).name & " returns from the Netlink of " & net(s2n(user(usernum).atNetlink)).name
send "~OL~FMYou drift back across the Ethers" & CRLF, usernum
user(usernum).room = user(usernum).oldRoom
user(usernum).listening = True
user(usernum).atNetlink = -1
look usernum
End Function
