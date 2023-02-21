Attribute VB_Name = "Commands"
Option Explicit
Option Compare Binary
Option Base 0

Sub samesite(usernum As Integer, inpstr As String)
Dim usr As String, usrComp As String, msg As String, count As Integer
Dim site As String
If wordCount(inpstr) < 2 Or Not (LCase$(word(1)) = "user" Or LCase$(word(1)) = "site") Then
    send "Usage: samesite user/site <user>/<site> [all]" & CRLF, usernum
    Exit Sub
    End If
If Not userExists(word(2)) And LCase$(word(1)) = "user" Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
msg = FANCY_BAR
If LCase$(word(1)) = "user" Then
    msg = msg & "Users with the same site as " & userCap(word(2)) & " ~FT["
    If userIsOnline(word(2)) Then
        site = user(getUser(word(2))).site
        Else
            user(0).name = word(2)
            loadUserData 0, False
            site = user(0).site
            End If
    msg = msg & site & "]" & CRLF
    Else
        site = word(2)
        msg = msg & "Users with the site of "
        msg = msg & Replace$(Replace$(site, "*", "~FY*~RS"), "?", "~FR?~RS") & CRLF
        End If
usr = Dir$(App.Path & "\Users\*.D")
If LCase$(word(3)) = "all" Then
    Do While Not usr = vbNullString
        usr = Left$(usr, Len(usr) - 2)
        user(0).name = usr
        loadUserData 0, False
        usrComp = user(0).site
        If (LCase$(usrComp) Like LCase$(site)) And Not userIsOnline(user(0).name) Then
            msg = msg & Space$(5) & userCap(user(0).name) & " " & user(0).desc & CRLF
            End If
        usr = Dir$
        Loop
    End If
For count = 1 To UBound(user)
    usrComp = user(count).site
    If user(count).operational And (LCase$(usrComp) Like LCase$(site)) Then
        msg = msg & Space(5) & userCap(user(count).name) & " " & user(count).desc & CRLF
        End If
    Next count
msg = msg & FANCY_BAR
send msg, usernum
End Sub
Sub age(usernum As Integer, inpstr As String)
If Not wordCount(inpstr) = 1 Then
    send "Usage: age <age>" & CRLF, usernum
    Exit Sub
    End If
If containsCorruptNumsOnly(word(1)) Then
    send "Your age must be a number" & CRLF, usernum
    Exit Sub
    End If
If Val(word(1)) > 120 Then
    send "Now be reasonable" & CRLF, usernum
    Exit Sub
    End If
user(usernum).age = word(1)
send "Age set to " & word(1) & CRLF, usernum
End Sub

Sub bcast(usernum As Integer, inpstr As String, beeper As Boolean)
If user(usernum).muzzled Then
    send MSG_USER_MUZZLED & CRLF, usernum
    Exit Sub
    End If
If wordCount(inpstr) < 1 Then
    If beeper Then
        send "Usage: bbcast <message>" & CRLF, usernum
        Exit Sub
        Else
            send "Usage: bcast <message>" & CRLF, usernum
            Exit Sub
            End If
    End If
Dim msg As String, count As Integer
msg = "~OL~FR--< ~RS~OL" & inpstr & "~OL~FR >--" & CRLF
If beeper Then
    msg = BELL & msg
    End If
For count = 1 To UBound(user)
    If user(count).state > STATE_LOGIN3 Then
        send msg, count
        End If
    Next count
End Sub

Sub cbuff(usernum As Integer)
Dim count As Integer, roomnum As Integer
roomnum = getRoom(user(usernum).room)
For count = 1 To REVBUFF_SIZE
    rooms(roomnum).buffer(count) = ""
    Next count
send "Review buffer cleared" & CRLF, usernum
End Sub

Sub clearScreen(usernum As Integer)
Dim count As Integer, msg As String
For count = 1 To user(usernum).pager + 5
    msg = msg & CRLF
    Next count
send msg, usernum
End Sub

Sub clearline(usernum As Integer, inpstr As String)
'This will clear a login by line number
Dim line As Integer
If wordCount(inpstr) <> 1 Then
    send "Ussage: clearline <line>" & CRLF, usernum
    Exit Sub
    End If
If containsCorruptNumsOnly(inpstr) Then
    send "Enter a line ~FYnumber~RS only" & CRLF, usernum
    Exit Sub
    End If
line = Val(inpstr)
If line < 1 Or line > UBound(user) Then
    send "Invalid line" & CRLF, usernum
    Exit Sub
    End If
If user(line).operational And user(line).state < STATE_NORMAL Then
    send "Clearing line " & Trim$(line) & CRLF, usernum
    send "Your line is being cleared" & CRLF, line
    killUser line
    Else
        send "You cannot clear the line of a logged in user" & CRLF, usernum
        End If
End Sub

Sub cmdlist(usernum As Integer)
Dim count As Integer, count2 As Integer, count3 As Integer
Dim msg As String, totalcommands As Integer, thispass As Boolean
Const BUFF = 12

msg = CRLF & FANCY_BAR
For count = 0 To user(usernum).rank
    msg = RTrim$(msg) & "~FR" & ranks(count) & Space(BUFF - Len(ranks(count))) & "~RS"
    msg = msg & "~FT"
    For count2 = 1 To UBound(cmds)
        If cmds(count2).rank = count Then
            msg = msg & Space(BUFF - Len(cmds(count2).name)) & cmds(count2).name
            thispass = True
            count3 = count3 + 1
            totalcommands = totalcommands + 1
            If count3 = 5 Then
                msg = msg & CRLF
                If Len(msg) > SEND_CHOP Then
                    send msg, usernum
                    msg = ""
                    End If
                msg = msg & Space(BUFF)
                count3 = 0
                thispass = False
                End If
            End If
        Next count2
    If thispass Then
        msg = msg & CRLF
        End If
    count3 = 0
    Next count
msg = RTrim$(msg) & FANCY_BAR & "Total of ~FM~OL" & LTrim$(Str$(totalcommands)) & "~RS commands." & CRLF & CRLF
send msg, usernum
End Sub

Sub colorDisplay(usernum As Integer)
Dim msg As String
msg = "~OLOL: SOUTHWEST VIDEO TEST~RS Bold" & CRLF
msg = msg & "~ULUL: SOUTHWEST VIDEO TEST~RS Underline" & CRLF
msg = msg & "~LILI: SOUTHWEST VIDEO TEST~RS Blinking" & CRLF
msg = msg & "~RVRV: SOUTHWEST VIDEO TEST~RS Reverse" & CRLF
msg = msg & "~FKFK: SOUTHWEST VIDEO TEST~RS Foreground Black" & CRLF
msg = msg & "~FRFR: SOUTHWEST VIDEO TEST~RS Foreground Red" & CRLF
msg = msg & "~FGFG: SOUTHWEST VIDEO TEST~RS Foreground Green" & CRLF
msg = msg & "~FYFY: SOUTHWEST VIDEO TEST~RS Foreground Yellow" & CRLF
msg = msg & "~FBFB: SOUTHWEST VIDEO TEST~RS Foreground Blue" & CRLF
msg = msg & "~FMFM: SOUTHWEST VIDEO TEST~RS Foreground Magenta" & CRLF
msg = msg & "~FTFT: SOUTHWEST VIDEO TEST~RS Foreground Turquiose" & CRLF
msg = msg & "~FWFW: SOUTHWEST VIDEO TEST~RS Foreground White" & CRLF
send msg, usernum
msg = vbNullString
msg = msg & "~BKBK: SOUTHWEST VIDEO TEST~RS Background Black" & CRLF
msg = msg & "~BRBR: SOUTHWEST VIDEO TEST~RS Background Red" & CRLF
msg = msg & "~BGBG: SOUTHWEST VIDEO TEST~RS Background Green" & CRLF
msg = msg & "~BYBY: SOUTHWEST VIDEO TEST~RS Background Yellow" & CRLF
msg = msg & "~BBBB: SOUTHWEST VIDEO TEST~RS Background Blue" & CRLF
msg = msg & "~BMBM: SOUTHWEST VIDEO TEST~RS Background Magenta" & CRLF
msg = msg & "~BTBT: SOUTHWEST VIDEO TEST~RS Background Turquiose" & CRLF
msg = msg & "~BWBW: SOUTHWEST VIDEO TEST~RS Background White" & CRLF
send msg, usernum
End Sub



Sub demote(usernum As Integer, inpstr As String)
Dim count As Integer
Dim UserName As String
If Not Len(word(1)) > 0 Then
    send "Usage: demote <user>" & CRLF, usernum
    Exit Sub
    End If
If Not wordCount(inpstr) = 1 Then
    send "Usage: demote <user>" & CRLF, usernum
    Exit Sub
    End If
UserName = word(1)
If Not userExists(UserName) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
For count = 1 To UBound(user)
    If UCase$(user(count).name) = UCase$(word(1)) And user(count).state > STATE_LOGIN3 Then
        If user(count).name = user(usernum).name Then
            send "You cannot demote yourself" & CRLF, usernum
            Exit Sub
            End If
        If user(count).rank <= LBound(ranks) Then
            user(count).rank = LBound(ranks)
            Exit Sub
            End If
        If user(count).rank >= user(usernum).rank Then
            send "You cannot demote a user of equal or greater rank" & CRLF, usernum
            send "~BR~FK" & user(usernum).name & " has tried to demote you!" & CRLF, count
            Exit Sub
            End If
        If user(count).rank = 0 Then
            send "They are about as low as they're going" & CRLF, usernum
            Exit Sub
            End If
        user(count).rank = user(count).rank - 1
        writeHistory user(count).name, "~FR~OLDEMOTED~RS by " & user(usernum).name & " to " & ranks(user(count).rank)
        writeSyslog "~FB" & user(usernum).name & "~RS demoted " & user(count).name & " to " & ranks(user(count).rank)
        writeRoom "", "~BR~FK" & user(usernum).name & " demoted " & user(count).name & " to " & ranks(user(count).rank) & CRLF
        Exit Sub
        End If
    Next count
user(0).name = word(1)
loadUserData 0
If user(0).rank >= user(usernum).rank Then
    send "You cannot demote a user of equal or greater rank" & CRLF, usernum
    Exit Sub
    End If
If user(0).rank = 0 Then
    send "They are about as low as they're going" & CRLF, usernum
    Exit Sub
    End If
user(0).rank = user(0).rank - 1
writeHistory user(0).name, "~FRDEMOTED~RS by " & user(usernum).name & " to " & ranks(user(count).rank)
saveUserData user(0)
writeSyslog "~FB" & user(usernum).name & "~RS demoted " & user(0).name & " to ~FM" & ranks(user(0).rank) & CRLF
End Sub

Sub desc(usernum As Integer, inpstr As String)
If Len(inpstr) > 0 Then
    If cLen(inpstr) > 30 Then
        send "Description too long" & CRLF, usernum
        Exit Sub
        End If
    user(usernum).desc = inpstr
    send "Description set" & CRLF, usernum
    Else
        send "Your description is: " & user(usernum).desc & CRLF, usernum
        End If

End Sub

Sub emote(usernum As Integer, inpstr As String)
Dim emotion As String
If Left(inpstr, 1) = ";" Then
    If Len(inpstr) > 1 Then
        emotion = Right$(inpstr, Len(inpstr) - 1)
        Else
            send "Emote what?" & CRLF, usernum
            Exit Sub
            End If
    Else
        emotion = inpstr
        End If
If Len(emotion) = 0 Then
    send "Emote what?" & CRLF, usernum
    Exit Sub
    End If
If Not InStr(Space(10), Left$(emotion, 8)) = 0 Then
    send "So, erm... Whats with all the spaces?" & CRLF, usernum
    Exit Sub
    End If
If Not Left$(emotion, 2) = "'s" Then
    emotion = " " & emotion
    End If
If user(usernum).visible Then
    writeRoom user(usernum).room, user(usernum).name & emotion & CRLF
    Else
        writeRoom user(usernum).room, "A shadow" & emotion & CRLF
        End If
writeRoomBuff user(usernum).name & emotion, getRoom(user(usernum).room)
End Sub

Sub entpro(usernum As Integer)
'Remote users cannot use this command
If user(usernum).netlinkType Then
    send MSG_NO_NETLINK & CRLF, usernum
    Exit Sub
    End If
user(usernum).editorType = EDITSTATE_ENTPRO
writeRoomExcept user(usernum).room, "~FT" & user(usernum).name & " starts to write a profile" & CRLF, user(usernum).name
lineEditor usernum, user(usernum).inpstr
End Sub

Sub gender(usernum As Integer, inpstr As String)
Select Case LCase$(word(1))
    Case "m", "male"
        user(usernum).gender = "Male"
    Case "f", "female"
        user(usernum).gender = "Female"
    Case "n", "neither"
        user(usernum).gender = "Neither"
    Case Else
        send "Usage: gender male/female/neither" & CRLF, usernum
        Exit Sub
        End Select
send "Gender set to " & user(usernum).gender & CRLF, usernum
End Sub

Sub examine(usernum As Integer, inpstr As String)
Dim msg As String, un As Integer, ticks As DHMS_OBJECT
Dim count As Integer, found As Boolean, datehold As Date
If Not wordCount(inpstr) = 1 Then
    send "Usage: examine <username>" & CRLF, usernum
    Exit Sub
    End If
If Not userExists(inpstr) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
If userIsOnline(inpstr) Then
    datehold = num2date(user(un).lastLogin)
    un = getUser(inpstr)
    msg = "~FYName~FB:~RS        " & user(un).name & " " & user(un).desc & "~RS"
    msg = mold(msg, (CLIENT_WIDTH - 1) - (6 + Len(ranks(user(un).rank))))
    msg = msg & "~FYRank~FB:~RS " & ranks(user(un).rank) & CRLF
    msg = msg & "~FYOn Since~FB:~RS    " & Format$(datehold, "dddd d") & getOrdinal(Int(Format$(datehold, "d"))) & Format$(datehold, " mmmm yyyy" & " at " & Format$(datehold, "hh:nn")) & CRLF
    msg = msg & "~FYOn For~FB:~RS      " & deriveTimeString(spliceTime(user(un).timeon), False) & CRLF
    msg = msg & "~FYIdle For~FB~RS     " & deriveTimeString(spliceTime(CDbl(user(un).idle * 60)), False) & CRLF
    msg = msg & "~FYTotal Login~FB:~RS " & deriveTimeString(spliceTime(user(un).totalTime), False) & CRLF
    msg = msg & "~FYSite~FB:~RS        " & user(un).site & CRLF
    Else
        un = 0
        user(0).name = inpstr
        loadUserData 0
        datehold = num2date(user(0).lastLogin)
        msg = "~FYName~FB:~RS        " & user(0).name & " " & user(0).desc & "~RS"
        msg = mold(msg, (CLIENT_WIDTH - 1) - (6 + Len(ranks(user(0).rank))))
        msg = msg & "~FYRank~FB:~RS " & ranks(user(0).rank) & CRLF
        msg = msg & "~FYLast Login~FB:~RS  " & Format$(datehold, "dddd d") & getOrdinal(Int(Format$(datehold, "d"))) & Format$(datehold, " mmmm yyyy" & " at " & Format$(datehold, "hh:nn")) & CRLF
        msg = msg & "~FYWhich was~FB:~RS   " & deriveTimeString(spliceTime(date2num(Now) - user(0).lastLogin)) & CRLF
        msg = msg & "~FYWas On For~FB:~RS  " & deriveTimeString(spliceTime(user(0).timeon), False) & CRLF
        msg = msg & "~FYTotal Login~FB:~RS " & deriveTimeString(spliceTime(user(0).totalTime), False) & CRLF
        msg = msg & "~FYLast Site~FB:~RS   " & user(0).site & CRLF
        End If
msg = FANCY_BAR & msg & embedBar("Profile")
For count = LBound(user(un).profile) To UBound(user(un).profile)
    If Not user(un).profile(count) = "" Then
        msg = msg & user(un).profile(count) & CRLF
        found = True
        End If
    Next count
If Not found Then
    msg = msg & "This user has not written a profile" & CRLF
    End If
msg = msg & FANCY_BAR
send msg, usernum
End Sub

Sub go_user(ByVal usernum As Integer, ByVal inpstr As String)
Dim temp As String, pass As String, count As Integer, usrCount As Integer
Dim goNetlink As Boolean, old_room As String, transString As String, count2 As Integer
If inpstr = vbNullString Then
    If user(usernum).room = rooms(1).name Then
        send "You are already in the main room" & CRLF, usernum
        Exit Sub
        End If
        old_room = user(usernum).room
        user(usernum).room = rooms(1).name
        look usernum
        writeRoom old_room, user(usernum).name & " " & user(usernum).exitMsg & " " & user(usernum).room & CRLF
        writeRoomExcept user(usernum).room, user(usernum).name & " " & user(usernum).enterMsg & " " & old_room & CRLF, user(usernum).name
    Exit Sub
    End If
'Get last word as password
For count = LBound(word) To UBound(word)
    If word(count) = "" Then
        Exit For
        End If
    pass = word(count)
    Next count
'Netlink names are allowed to have spaces in them. This causes a
'few headaches but we should do just fine. For the .go command,
'the spaces in Netlink names are replaced with underscores.
'See if there is a Netlink in that user's room
For count = LBound(net) To UBound(net)
    If UCase$(inpstr) = UCase$(net(count).name) Then
        goNetlink = True
        transString = "TRANS " & user(usernum).name & " " & user(usernum).password
        Exit For
        ElseIf UCase$(stripLast(inpstr)) = UCase$(net(count).name) And Not stripLast(inpstr) = "" Then
            goNetlink = True
            transString = "TRANS " & user(usernum).name & " " & crypt(pass)
            Exit For
            End If
    Next count
If goNetlink Then
    If Not net(count).state = NETSTATES.NETLINK_UP Then
        send "This netlink is not connected" & CRLF, usernum
        Exit Sub
        End If
    If user(usernum).netlinkPending Then
        send "Already processing a Netlink login for you" & CRLF, usernum
        End If
    user(usernum).netlinkPending = True
    netout transString & " " & Trim$(user(usernum).rank) & " " & user(usernum).desc & LF, net(count).line
    Exit Sub
    End If

If wordCount(inpstr) > 2 Then
    send "Usage: go <room/netlink> [netlink password]" & CRLF, usernum
    Exit Sub
    End If
For count = LBound(rooms) To UBound(rooms)
    If InStr(LCase$(rooms(count).name), LCase$(word(1))) = 1 Then
        If user(usernum).room = rooms(count).name Then
            send "You are already in the " & user(usernum).room & CRLF, usernum
            Exit Sub
            End If
        If user(usernum).rank < staffLevel And InStr(UCase$(rooms(getRoom(user(usernum).room)).allExits), UCase$(word(1))) = 0 Then
            send "The " & userCap(rooms(getRoom(word(1))).name) & " is not an exit to this room." & CRLF, usernum
            Exit Sub
            End If
        If rooms(count).access = ROOM_STAFF And user(usernum).rank < staffLevel Then
            send "This room is for staff members only" & CRLF, usernum
            Exit Sub
            End If
        If rooms(count).access = ROOM_PRIVATE And Not isUserInvited(usernum, rooms(count).name) And (system.gatecrash = False Or user(usernum).rank < system.gatecrashLevel) Then
            send "The " & rooms(count).name & " is currently private" & CRLF, usernum
            Exit Sub
            End If
        old_room = user(usernum).room
        user(usernum).room = rooms(count).name
        look usernum
        writeRoom old_room, user(usernum).name & " " & user(usernum).exitMsg & " " & user(usernum).room & CRLF
        writeRoomExcept user(usernum).room, user(usernum).name & " " & user(usernum).enterMsg & " " & old_room & CRLF, user(usernum).name
        'Count users, if private and usrnum now =< 1 then return to public
        For count2 = 1 To UBound(user)
            If user(count2).room = old_room Then
                usrCount = usrCount + 1
                End If
            Next count2
        If usrCount < 2 And rooms(getRoom(old_room)).access = ROOM_PRIVATE Then
            user(0).room = old_room
            set_public 0
            End If
        Exit Sub
        End If
    Next count
send MSG_ROOM_NOT_EXIST & CRLF, usernum
End Sub

Sub help(usernum As Integer, inpstr As String)
If inpstr = "" Then
    cmdlist usernum
    Exit Sub
    End If
Dim BadString As String, FromFile As String
If InStr(inpstr, ".") Or InStr(inpstr, "\") Or InStr(inpstr, ":") Then
    send "Im afraid I cant do that" & CRLF, usernum
    Exit Sub
    End If
If Dir$(App.Path & "\Help\" & word(1) & ".HF") = "" Then
    send "There is no help available on that topic" & CRLF, usernum
    Exit Sub
    End If
Open App.Path & "\Help\" & word(1) & ".HF" For Input As #1
Do While Not EOF(1)
    Line Input #1, FromFile
    send FromFile & CRLF, usernum
    Loop
Close #1
End Sub

Sub look(usernum As Integer)
Dim count, count2, found  As Boolean, msg As String, rn As Integer
'First we tell them the name of the room
msg = CRLF & "~FTRoom: ~FG" & userCap(user(usernum).room) & CRLF & CRLF
'Now we get the room description, if there is one
If Not Dir$(App.Path & "\Rooms\" & user(usernum).room & ".R") = "" Then
    Open App.Path & "\Rooms\" & user(usernum).room & ".R" For Input As #1
    Dim RoomDesc As String
    Do While Not EOF(1)
        Line Input #1, RoomDesc
        msg = msg & RoomDesc & CRLF
        Loop
    Close #1
    End If
msg = msg & CRLF
found = False
'Show the room's exits
For count = LBound(rooms) To UBound(rooms)
    If LCase$(rooms(count).name) = LCase$(user(usernum).room) Then
        For count2 = LBound(rooms(LBound(rooms)).exits) To UBound(rooms(LBound(rooms)).exits)
            If Not rooms(count).exits(count2) = "" Then
                If found = False Then
                    msg = msg & "~FTExits are: ~RS"
                    found = True
                    End If
                msg = msg & "~FG" & rooms(count).exits(count2) & "  ~RS"
                End If
            Next count2
        End If
    Next count
If found = False Then
    msg = msg & "~FGThere are no exits~RS"
    End If
msg = msg & CRLF
'Display Netlinked rooms
found = False
For count = LBound(net) To UBound(net)
    If UCase$(net(count).room) = UCase$(user(usernum).room) Then
        If Not found Then
            msg = msg & "~FTNetlinks are: "
            Else
                msg = msg & "  "
                End If
        Select Case net(count).state
            Case NETSTATES.NETLINK_DOWN
                msg = msg & "~FR" & net(count).name
            Case NETSTATES.NETLINK_UP
                msg = msg & "~FG" & net(count).name
            Case Else
                msg = msg & "~FY" & net(count).name
                End Select
        found = True
        End If
    Next count
msg = msg & CRLF
If found Then
    msg = msg & CRLF
    End If

'Lets see who is online
found = False
For count = 1 To UBound(user)
    If UCase$(user(count).room) = UCase$(user(usernum).room) And user(count).state > STATE_LOGIN3 And Not count = usernum Then
        If user(count).visible Or user(usernum).rank >= user(count).rank Then
            If found = False Then
                msg = msg & "~FG~OLYou can see:~RS" & CRLF
                found = True
                End If
            If user(count).visible Then
                msg = msg & "      " & user(count).name & " " & user(count).desc & CRLF
                Else
                    msg = msg & "     ~FR*~RS" & user(count).name & " " & user(count).desc & CRLF
                    End If
            End If
        End If
    Next count
For count = LBound(clones) To UBound(clones)
    If rooms(clones(count).room).name = user(usernum).room Then
        If found = False Then
            msg = msg & "~FG~OLYou can see:~RS" & CRLF
            found = True
            End If
        msg = msg & "      " & user(clones(count).owner).name & " ~BR~FW(CLONE)" & CRLF
        End If
    Next count
If Not found Then
    msg = msg & "~FGYou are all alone here" & CRLF
    End If
msg = msg & CRLF
'Show accesses
rn = getRoom(user(usernum).room)
If rooms(rn).access = ROOM_STAFF Then
    msg = msg & "This room is for staff access only"
    Else
        If rooms(rn).locked Then
            msg = msg & "Access is ~FRfixed~RS to "
            Else
                msg = msg & "Access is set to "
                End If
        If rooms(rn).access = ROOM_PRIVATE Then
            msg = msg & "~FRPRIVATE~RS"
            Else
                msg = msg & "~FGPUBLIC~RS"
                End If
        End If
msg = msg & " and there are ~OL~FM" & getMessageCount(App.Path & "\Rooms\" & user(usernum).room & ".B") & "~RS messages on the board." & CRLF
'Bring on the topics
If rooms(getRoom(user(usernum).room)).topic = "" Then
    msg = msg & "~OL~FGCurrent topic: ~RS" & MSG_TOPIC_NOT_SET & CRLF
        Else
        msg = msg & "~OL~FGCurrent topic: ~RS" & rooms(getRoom(user(usernum).room)).topic & CRLF
        End If
send msg, usernum
End Sub

Sub map(usernum As Integer)
If Dir$(App.Path & "\Rooms\Map.S") = "" Then
    send "There is no map" & CRLF, usernum
    Else
        Open App.Path & "\Rooms\Map.S" For Input As #1
        Dim MapLine As String
        Do While Not EOF(1)
            Line Input #1, MapLine
            send MapLine & CRLF, usernum
            Loop
        Close #1
        End If
End Sub

Sub moveUser(usernum As Integer, inpstr As String)
word(1) = completeUsername(word(1))
If wordCount(inpstr) < 1 Or wordCount(inpstr) > 2 Then
    send "Usage: move <user> [room]" & CRLF, usernum
    Exit Sub
    End If
Dim count As Integer, un As Integer, rn As Integer
If Not userExists(word(1)) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
If Not userIsOnline(word(1)) Then
    send MSG_USER_NOT_ONLINE & CRLF, usernum
    Exit Sub
    End If
un = getUser(word(1))
If usernum = un Then
    send "You cannot move yourself" & CRLF, usernum
    Exit Sub
    End If
If user(un).rank >= user(usernum).rank Then
    send "You cannot move a user of equal or greater rank than yourself" & CRLF, usernum
    Exit Sub
    End If
If wordCount(inpstr) > 1 Then
    rn = getRoom(word(2))
    Else
        rn = getRoom(user(usernum).room)
        End If
If rn = 0 Then
    send MSG_ROOM_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
If user(un).room = rooms(rn).name Then
    If wordCount(inpstr) = 1 Then
        send user(un).name & " is already here" & CRLF, usernum
        Else
            send user(un).name & " is already there" & CRLF, usernum
            End If
    Exit Sub
    End If
send user(usernum).name & " has pulled you into the " & rooms(rn).name & CRLF, un
writeRoomExcept user(un).room, "~FT" & user(un).name & " is pulled into the cold black night" & CRLF, user(un).name
user(un).room = rooms(rn).name
send "You move " & user(un).name & " to the " & user(un).room & CRLF, usernum
writeRoomExcept user(un).room, "~FT" & user(un).name & " is thrown into the room" & CRLF, user(un).name
look (un)
End Sub

Sub murder(usernum As Integer, inpstr As String)
Dim UserName As String
Dim count As Integer
inpstr = completeUsername(inpstr)
word(1) = completeUsername(word(1))
If Not Len(word(1)) > 0 Then
    send "Usage: kill <user>" & CRLF, usernum
    Exit Sub
    End If
If Not wordCount(inpstr) = 1 Then
    send "Usage: kill <user>" & CRLF, usernum
    Exit Sub
    End If
UserName = word(1)
If Not userExists(UserName) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
For count = 1 To UBound(user)
    If UCase$(user(count).name) = UCase$(word(1)) And user(count).state > STATE_LOGIN3 Then
        If user(count).name = user(usernum).name Then
            send "Suicide is not the awnser, man." & CRLF, usernum
            Exit Sub
            End If
        If user(count).rank >= user(usernum).rank Then
            send "You cannot kill a user of equal or greater rank" & CRLF, usernum
            send "~BR~FK" & user(usernum).name & " has tried to kill you!" & CRLF, count
            Exit Sub
            End If
        writeRoom "", "~FGA beam of green light shoots from " & userCap(user(usernum).name) & "'s finger and strikes " & userCap(UserName) & ".~RS" & CRLF
        writeSyslog "~FG" & user(usernum).name & "~RS killed ~FB" & user(count).name
        writeHistory user(count).name, "~FR~OLKILLED~RS by " & user(usernum).name
        killUser (count)
        writeRoom "", "~FG" & userCap(UserName) & " vaporizes!" & CRLF
        Exit Sub
        End If
    Next count
    send "That user is not logged in." & CRLF, usernum
End Sub

Sub muzzle(usernum As Integer, inpstr As String)
inpstr = completeUsername(inpstr)
word(1) = completeUsername(word(1))
If Not wordCount(inpstr) = 1 Then
    send "Usage: muzzle <user>" & CRLF, usernum
    Exit Sub
    End If
Dim user_to_muzzle As Integer
user_to_muzzle = getUser(word(1))
If user_to_muzzle <= 0 Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
If usernum = user_to_muzzle Then
    send "Self-discipline is good but don't you think this is going a little too far?" & CRLF, usernum
    Exit Sub
    End If
If user(user_to_muzzle).muzzled Then
    send user(user_to_muzzle).name & " is already muzzled" & CRLF, usernum
    Exit Sub
    End If
If user(user_to_muzzle).rank >= user(usernum).rank Then
    send "You cannot muzzle a user of equal or greater rank" & CRLF, usernum
    Exit Sub
    End If
user(user_to_muzzle).muzzled = True
writeHistory user(user_to_muzzle).name, "~FR~OLMUZZLED~RS by " & user(usernum).name
send "~OL~BB~FWYou have been muzzled!" & CRLF, user_to_muzzle
send "~OL~BB~FWYou have placed a muzzle on " & user(user_to_muzzle).name & CRLF, usernum
End Sub

Sub nuke(usernum As Integer, inpstr As String)
Dim count As Integer
Dim filename, UserName As String

If Not Len(word(1)) > 0 Then
    send "Usage: nuke <user>" & CRLF, usernum
    Exit Sub
    End If
If Not wordCount(inpstr) = 1 Then
    send "Usage: nuke <user>" & CRLF, usernum
    Exit Sub
    End If
UserName = word(1)
If Not userExists(UserName) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
If userIsOnline(UserName) Then
    send "You cannot nuke a user that is online" & CRLF, usernum
    Exit Sub
    End If
user(0).name = UserName
loadUserData 0
If user(0).rank >= user(usernum).rank Then
    send "You cannot erase the account of a user of equal or greater rank" & CRLF, usernum
    Exit Sub
    End If
deleteAccount user(count).name
writeSyslog "~FR" & user(usernum).name & " nuked " & UserName
send "~BR~FKUser " & userCap(UserName) & " nuked~RS" & CRLF, usernum
End Sub

Sub passwd(inpstr As String, usernum As Integer)
If Not wordCount(inpstr) = 2 Then
    send "Usage: passwd <old password> <new password>" & CRLF, usernum
    Exit Sub
    End If
If Not crypt(word(1)) = user(usernum).password Then
    send "Passwords do not match" & CRLF, usernum
    Exit Sub
    End If
If Not Len(word(2)) > 3 Then
    send "New password too short" & CRLF, usernum
    Exit Sub
    End If
user(usernum).password = crypt(word(2))
clearScreen (usernum)
send "Password changed" & CRLF, usernum
End Sub

Sub promote(usernum As Integer, inpstr As String)
'My name is Ozymandias, King of kings.
'Look on my works, ye mighty, and despair.
'                    -Percy Byshe Shelley:
'                     Ozymandias
Dim count As Integer
Dim UserName As String
If Not Len(word(1)) > 0 Then
    send "Usage: promote <user>" & CRLF, usernum
    Exit Sub
    End If
If Not wordCount(inpstr) = 1 Then
    send "Usage: promote <user>" & CRLF, usernum
    Exit Sub
    End If
UserName = word(1)
If Not userExists(UserName) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
For count = 1 To UBound(user)
    If UCase$(user(count).name) = UCase$(word(1)) And user(count).state > STATE_LOGIN3 Then
        If user(count).name = user(usernum).name Then
            send "You cannot promote yourself" & CRLF, usernum
            Exit Sub
            End If
        If user(count).rank >= user(usernum).rank Then
            send "You cannot promote a user of equal or greater rank" & CRLF, usernum
            Exit Sub
            End If
        If user(count).rank >= UBound(ranks) Then
            user(count).rank = UBound(ranks)
            Exit Sub
            End If
        user(count).rank = user(count).rank + 1
        writeSyslog "~FB" & user(usernum).name & "~RS promoted ~FB" & user(count).name & "~RS to ~FM" & ranks(user(count).rank)
        writeRoom "", "~BG~FK" & user(usernum).name & " promoted " & user(count).name & " to " & ranks(user(count).rank) & CRLF
        If user(usernum).name = "A server administrator" Then
            user(usernum).name = LCase$(user(usernum).name)
            End If
        writeHistory user(count).name, "~FG~OLPROMOTED~RS by " & user(usernum).name & " to " & ranks(user(count).rank)
        Exit Sub
        End If
    Next count
user(0).name = word(1)
loadUserData 0
        If user(0).rank >= user(usernum).rank Then
            send "You cannot promote a user of equal or greater rank" & CRLF, usernum
            Exit Sub
            End If
user(0).rank = user(0).rank + 1
saveUserData user(0)
writeHistory user(0).name, "~FG~OLPROMOTED~RS by " & user(usernum).name & " to " & ranks(user(0).rank)
writeSyslog "~FB" & user(usernum).name & "~RS promoted ~FB" & user(0).name & "~RS to ~FM" & ranks(user(0).rank)
End Sub

Sub quit(usernum As Integer)
killUser usernum
End Sub

Sub read_board(usernum As Integer)
Dim file As String
file = App.Path & "\Rooms\" & user(usernum).room & ".B"
If getMessageCount(file) = 0 Then
    send "There are no messages on the message board" & CRLF, usernum
    Exit Sub
    End If
Dim FromFile As String, BoardFull As String
Open file For Input As #1
Do While Not EOF(1)
    Line Input #1, FromFile
    BoardFull = BoardFull & FromFile & CRLF
    Loop
send BoardFull & CRLF, usernum
Close #1
writeRoomExcept user(usernum).room, user(usernum).name & " reads the message board", user(usernum).name
End Sub

Sub reboot(usernum As Integer, inpstr As String)
'Shut 'er down, sir!
'Remote users cannot use this command
If user(usernum).netlinkType Then
    send MSG_NO_NETLINK & CRLF, usernum
    Exit Sub
    End If
If Not system.shutdownType = SHUTDOWN_NONE Then
    If system.shutdownType = SHUTDOWN_SHUTDOWN Then
        send "A shutdown is already in affect" & CRLF, usernum
        Else
            If word(1) = "cancel" Then
            mainForm.Shutdown_Timer.Enabled = False
            system.shutdownCount = 0
            system.shutdownType = SHUTDOWN_NONE
            writeSyslog "~FB" & user(usernum).name & "~RS has canceled the reboot"
                writeRoom "", "~FGReboot has been canceled" & CRLF
                Else
                    send "A reboot is already in affect" & CRLF, usernum
                    End If
            End If
    Exit Sub
    End If
If word(1) = "cancel" Then
    send "A reboot is not in affect" & CRLF, usernum
    Exit Sub
    End If
If containsCorruptNumsOnly(word(1)) Then
    send "Usage: reboot [countdown]|cancel" & CRLF, usernum
    Exit Sub
    End If
If Len(word(1)) > 4 Then
    send "Countdown too big" & CRLF, usernum
    Exit Sub
    End If
system.shutdownCount = Val(word(1))
If system.shutdownCount = 0 Then
    writeSyslog "~FB" & user(usernum).name & "~RS has initiated a reboot"
    Else
        system.shutdownCount = system.shutdownCount + 1
        End If
user(usernum).state = STATE_OPTION
user(usernum).options = OPTION_REBOOT
send "~BR~FWAre you sure you want to reboot the talker?~RS", usernum
End Sub

Sub rmail(usernum As Integer)
If getMessageCount(App.Path & "\Users\" & user(usernum).name & ".M") = 0 Then
    send "You have no mail" & CRLF, usernum
    Exit Sub
    End If
Dim MailFromFile, AllMail As String
Dim count As Integer
Open App.Path & "\Users\" & user(usernum).name & ".M" For Input As #1
Do While Not EOF(1)
    For count = 1 To user(usernum).pager
        If Not EOF(1) Then
            Line Input #1, MailFromFile
            AllMail = AllMail & MailFromFile & CRLF
            Else
                Exit For
                End If
        Next count
        send AllMail, usernum
        AllMail = ""
    Loop
Close #1
user(usernum).unread = False
send CRLF, usernum
End Sub

Sub say(usernum As Integer)
Dim act_type, text As String
If user(usernum).muzzled Then
    send MSG_USER_MUZZLED & CRLF, usernum
    Exit Sub
    End If
If Len(user(usernum).inpstr) < 1 Then
    send "Usage: say <text>" & CRLF, usernum
    Exit Sub
    End If
Select Case Right$(user(usernum).inpstr, 1)
    Case "?"
        act_type = "asks"
    Case "!"
        act_type = "exclaims"
    Case Else
        act_type = "says"
        End Select
Select Case Right$(user(usernum).inpstr, 2)
    Case ":)", "=)"
        act_type = "smiles"
    Case ";)"
        act_type = "winks"
    Case ":("
        act_type = "frowns"
    End Select
If user(usernum).visible Then
    text = "~FG" & user(usernum).name & " " & act_type & ": " & "~RS" & user(usernum).inpstr & CRLF
    Else
        text = "~FGA shadow " & act_type & ": " & "~RS" & user(usernum).inpstr & CRLF
        End If
writeRoomExcept user(usernum).room, text, user(usernum).name
send "~FGYou " & Left$(act_type, Len(act_type) - 1) & ":~RS " & user(usernum).inpstr & CRLF, usernum
writeRoomBuff "~FG" & user(usernum).name & " " & act_type & ": " & "~RS" & user(usernum).inpstr, getRoom(user(usernum).room)
End Sub

Sub set_outmsg(usernum As Integer, inpstr As String)
If Len(inpstr) > 40 Then
    send "Exit message is too big" & CRLF, usernum
    Exit Sub
    End If
If Len(inpstr) = 0 Then
    send "Usage: outmsg <enter message>" & CRLF, usernum
    Exit Sub
    End If
user(usernum).exitMsg = inpstr
send "Exit message sent" & CRLF, usernum
End Sub

Sub suicide(usernum As Integer, inpstr As String)
If user(usernum).netlinkType Then
    send MSG_NO_NETLINK & CRLF, usernum
    Exit Sub
    End If
If Not wordCount(inpstr) = 1 Then
    send "Usage: suicide <your password>" & CRLF, usernum
    Exit Sub
    End If
If Not crypt(word(1)) = user(usernum).password Then
    send "Incorrect password" & CRLF, usernum
    Exit Sub
    End If
send "Are you sure you want to do this?", usernum
user(usernum).state = STATE_OPTION
user(usernum).options = OPTION_SUICIDE
End Sub
Sub Vis(usernum As Integer)
If user(usernum).visible Then
    send "You are already visible" & CRLF, usernum
    Else
        writeRoom user(usernum).room, user(usernum).name & " " & MSG_USER_GOES_VIS & CRLF
        user(usernum).visible = True
        End If
End Sub

Sub invis(usernum As Integer)
If Not user(usernum).visible Then
    send "You are already invisible" & CRLF, usernum
    Else
        writeRoom user(usernum).room, user(usernum).name & " " & MSG_USER_GOES_INVIS & CRLF
        user(usernum).visible = False
        End If
End Sub

Sub make_invis(usernum As Integer, inpstr As String)
Dim un As Integer
inpstr = completeUsername(inpstr)
word(1) = completeUsername(word(1))
If Not wordCount(inpstr) = 1 Then
    send "Usage: makeinvis <username>" & CRLF, usernum
    Exit Sub
    End If
If Not userExists(word(1)) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
If Not userIsOnline(word(1)) Then
    send MSG_USER_NOT_ONLINE & CRLF, usernum
    Exit Sub
    End If
un = getUser(word(1))
If user(un).rank >= user(usernum).rank Then
    send "You cannot force a user of equal or greater rank into the shadows" & CRLF, usernum
    Exit Sub
    End If
If Not user(un).visible Then
    send user(un).name & " is already invisible" & CRLF, usernum
    Exit Sub
    End If
user(un).visible = False
writeRoom user(un).room, user(usernum).name & " pushes " & user(un).name & " into the shadows" & CRLF
End Sub

Sub make_vis(usernum As Integer, inpstr As String)
Dim un As Integer
inpstr = completeUsername(inpstr)
word(1) = completeUsername(word(1))
If Not wordCount(inpstr) = 1 Then
    send "Usage: makevis <username>" & CRLF, usernum
    Exit Sub
    End If
If Not userExists(word(1)) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
If Not userIsOnline(word(1)) Then
    send MSG_USER_NOT_ONLINE & CRLF, usernum
    Exit Sub
    End If
un = getUser(word(1))
If user(un).rank >= user(usernum).rank Then
    send "You cannot force a user of equal or greater rank out of the shadows" & CRLF, usernum
    Exit Sub
    End If
If user(un).visible Then
    send user(un).name & " is already visible" & CRLF, usernum
    Exit Sub
    End If
user(un).visible = True
writeRoom user(un).room, user(usernum).name & " pushes " & user(un).name & " out of the shadows" & CRLF
End Sub

Sub set_inmsg(usernum As Integer, inpstr As String)
If Len(inpstr) > 40 Then
    send "Enter message is too big" & CRLF, usernum
    Exit Sub
    End If
If Len(inpstr) = 0 Then
    send "Usage: inmsg <enter message>" & CRLF, usernum
    Exit Sub
    End If
user(usernum).enterMsg = inpstr
send "Enterance message sent" & CRLF, usernum
End Sub

Sub shout(usernum As Integer, inpstr As String)
If user(usernum).muzzled Then
    send "You are muzzled and cannot shout" & CRLF, usernum
    Exit Sub
    End If
If inpstr = "" Then
    send "Usage: shout <message>" & CRLF, usernum
    Exit Sub
    End If
Dim count As Integer
For count = 1 To UBound(user)
    If user(count).listening And Not usernum = count Then
        send "~OL" & user(usernum).name & " shouts: ~RS" & inpstr & CRLF, count
        End If
    Next count
send "~OLYou shout: ~RS" & inpstr & CRLF, usernum
End Sub

Sub site(usernum As Integer, inpstr As String)
Dim un As Integer
If wordCount(inpstr) < 1 Then
    send "Usage: site <user>" & CRLF, usernum
    Exit Sub
    End If
If Not userExists(word(1)) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
If userIsOnline(word(1)) Then
    un = getUser(word(1))
    send user(un).name & " is logged in from " & user(un).site & CRLF, usernum
    Else
        user(0).name = word(1)
        loadUserData 0
        send user(0).name & " was last logged in from " & user(0).site & CRLF, usernum
        End If
End Sub

Sub shutdown(usernum As Integer, inpstr As String)
'Shut 'er down, sir!

'Remote users cannot use this command
If user(usernum).netlinkType Then
    send MSG_NO_NETLINK & CRLF, usernum
    Exit Sub
    End If
If Not system.shutdownType = SHUTDOWN_NONE Then
    If system.shutdownType = SHUTDOWN_SHUTDOWN Then
        If word(1) = "cancel" Then
            mainForm.Shutdown_Timer.Enabled = False
            system.shutdownCount = 0
            system.shutdownType = SHUTDOWN_NONE
            writeRoom "", "~FGShutdown has been canceled" & CRLF
            writeSyslog "~FB" & user(usernum).name & "~RS has canceled the shutdown"
            Else
                send "A shutdown is already in affect" & CRLF, usernum
                End If
        Else
            send "A reboot is already in affect" & CRLF, usernum
            End If
    Exit Sub
    End If
If word(1) = "cancel" Then
    send "A shutdown is not in affect" & CRLF, usernum
    Exit Sub
    End If
If containsCorruptNumsOnly(word(1)) Then
    send "Usage: shutdown [countdown]|cancel" & CRLF, usernum
    Exit Sub
    End If
If Len(word(1)) > 4 Then
    send "Countdown too big" & CRLF, usernum
    Exit Sub
    End If
system.shutdownCount = Val(word(1))
If Val(system.shutdownCount) = 0 Then
    writeSyslog "~FB" & user(usernum).name & "~RS has initiated a shutdown"
    Else
        system.shutdownCount = system.shutdownCount + 1
        End If
user(usernum).state = STATE_OPTION
user(usernum).options = OPTION_SHUTDOWN
send "~BR~FWAre you sure you want to shut down the talker?~RS", usernum
End Sub

Sub sing(usernum As Integer, inpstr As String)
If user(usernum).muzzled Then
    send MSG_USER_MUZZLED & CRLF, usernum
    Exit Sub
    End If
If inpstr = "" Then
    send "Usage: sing <lyrics>" & CRLF, usernum
    Exit Sub
    End If
writeRoom user(usernum).room, "~FG" & user(usernum).name & " sings: ~FTo/~ ~RS" & inpstr & "~FT o/~" & CRLF
End Sub

Sub smail(usernum As Integer, inpstr As String)
'Remote users cannot use this command
If user(usernum).netlinkType Then
    send MSG_NO_NETLINK & CRLF, usernum
    Exit Sub
    End If
If Not wordCount(inpstr) = 1 Then
    send "Usage: smail <user>" & CRLF, usernum
    Exit Sub
    End If

Dim founduser As Integer, count As Integer
For count = 1 To UBound(user)
    If LCase$(user(usernum).outMail.receiver) = LCase$(user(count).name) Then
        founduser = True
        Exit For
        End If
    Next count

If founduser Or userExists(word(1)) Then
    user(usernum).outMail.receiver = word(1)
    writeRoomExcept user(usernum).room, "~FT" & user(usernum).name & " starts to write some mail" & CRLF, user(usernum).name
    user(usernum).editorType = EDITSTATE_SMAIL
    lineEditor usernum, user(usernum).inpstr
    Else
        send MSG_USER_NOT_EXIST & CRLF, usernum
        End If
End Sub

Sub tell(usernum As Integer, inpstr As String)
If user(usernum).muzzled Then
    send MSG_USER_MUZZLED & CRLF, usernum
    Exit Sub
    End If
word(1) = completeUsername(word(1))
If wordCount(inpstr) < 2 Then
    send "Usage: tell <user> <message>" & CRLF, usernum
    Exit Sub
    End If
If Not userIsOnline(word(1)) Then
    send "That user is currently not online" & CRLF, usernum
    Exit Sub
    End If
If getUser(word(1)) = usernum Then
    send "Talking to yourself again?" & CRLF, usernum
    Exit Sub
    End If
inpstr = stripOne(inpstr)
send "~FT" & user(usernum).name & " tells you: ~RS" & inpstr & CRLF, getUser(word(1))
send "~FT" & "You tell " & userCap(word(1)) & ": ~RS" & inpstr & CRLF, usernum
End Sub

Sub unmuzzle(usernum As Integer, inpstr As String)
inpstr = completeUsername(inpstr)
word(1) = completeUsername(word(1))
If Not wordCount(inpstr) = 1 Then
    send "Usage: unmuzzle <user>" & CRLF, usernum
    Exit Sub
    End If
Dim user_to_muzzle As Integer
user_to_muzzle = getUser(word(1))
If user_to_muzzle <= 0 Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
If usernum = user_to_muzzle Then
    send "Im afraid you can't unmuzzle yourself" & CRLF, usernum
    Exit Sub
    End If
If Not user(user_to_muzzle).muzzled Then
    send user(user_to_muzzle).name & " is not muzzled" & CRLF, usernum
    Exit Sub
    End If
If user(user_to_muzzle).rank >= user(usernum).rank Then
    send "You cannot unmuzzle a user of equal or greater rank" & CRLF, usernum
    Exit Sub
    End If
user(user_to_muzzle).muzzled = False
writeHistory user(user_to_muzzle).name, "~FG~OLUNMUZZLED~RS by " & user(usernum).name
send "~OL~BB~FWYou have been unmuzzled!" & CRLF, user_to_muzzle
send "~OL~BB~FWYou unmuzzled " & user(user_to_muzzle).name & CRLF, usernum
End Sub

Sub wake(usernum As Integer, inpstr As String)
inpstr = completeUsername(inpstr)
word(1) = completeUsername(word(1))
If user(usernum).muzzled Then
    send MSG_USER_MUZZLED & CRLF, usernum
    Exit Sub
    End If
If Not wordCount(inpstr) = 1 Then
    send "Usage: wake <username>" & CRLF, usernum
    Exit Sub
    End If
Dim user_to_wake As Integer
user_to_wake = getUser(word(1))
If user_to_wake < 1 Then
    send "That user is not online" & CRLF, usernum
    Exit Sub
    End If
If usernum = user_to_wake Then
    send "You cannot wake yourself" & CRLF, usernum
    Exit Sub
    End If
send "~BR~OL~FWFrom: " & user(usernum).name & " >>>WAKE UP<<<" & BELL & CRLF, user_to_wake
send "Wake up call sent" & CRLF, usernum
End Sub

Sub who(usernum As Integer)
'Probably the most famous talker command: The who. This
'command will tell the user who is online and some info
'about them (like rank, idle time, ect).
Dim count As Integer, msg As String, specString As String
Dim strout As String, everyOther As Integer
msg = CRLF & FANCY_BAR
msg = msg & "~FMCurrent users on " & Format$(Now, "dddd d") & getOrdinal(Int(Format$(Now, "d"))) & Format$(Now, " mmmm yyyy" & " at " & Format$(Now, "hh:nn")) & CRLF
msg = msg & "~FYName                                    " & "Rank      " & "Room           " & "Tm/Id" & CRLF
msg = msg & FANCY_BAR
For count = 1 To UBound(user)
    If user(count).state > STATE_LOGIN3 Then
        If user(count).visible Or user(count).rank <= user(usernum).rank Then
            specString = "  "
            If Not user(count).visible And Not user(count).netlinkType Then
                    specString = "~FR *"
                    If everyOther = 0 Then
                        specString = specString & "~RS~FT"
                        Else
                            specString = specString & "~RS"
                            End If
                ElseIf user(count).visible And user(count).netlinkType Then
                    specString = "~FR@ "
                    If everyOther = 0 Then
                        specString = specString & "~RS~FT"
                        Else
                            specString = specString & "~RS"
                            End If
                ElseIf Not user(count).visible And user(count).netlinkFrom Then
                    specString = "~FR@*"
                    If everyOther = 0 Then
                        specString = specString & "~RS~FT"
                        Else
                            specString = specString & "~RS"
                            End If
                    End If
            strout = specString & user(count).name & " " & user(count).desc
            If everyOther = 0 Then
                strout = strout & "~RS~FT"
                Else
                    strout = strout & "~RS"
                    End If
            strout = strout & Space$(40 - ((Len(user(count).name) + _
            cLen(user(count).desc)) + 3)) & ranks(user(count).rank) & _
            Space$(10 - Len(ranks(user(count).rank))) & user(count).room & _
            Space$(15 - Len(user(count).room)) & user(count).timeon / 60 & _
            "/" & user(count).idle & CRLF
            If everyOther = 0 Then
                strout = "~FT" & strout & "~RS"
                everyOther = 1
                Else
                    everyOther = 0
                    End If
            msg = msg & strout
            If Len(msg) > SEND_CHOP Then
                send msg, usernum
                msg = ""
                End If
            End If
        End If
    Next count
msg = msg & FANCY_BAR & CRLF
send msg, usernum
End Sub

Sub write_board(usernum As Integer)
    writeRoomExcept user(usernum).room, "~FT" & user(usernum).name & " starts to write something on the message board" & "~RS" & CRLF, user(usernum).name
    user(usernum).editorType = EDITSTATE_BOARD
    lineEditor usernum, user(usernum).inpstr
End Sub

Sub system_info_show(usernum As Integer)
Const TWOCOL = 24
Dim stuff As String, msgout As String
msgout = FANCY_BAR & "~FBSouthWest ~FWSystem Information" & CRLF
msgout = msgout & FANCY_BAR & "   IP Address:   "
stuff = mainForm.Socket1.LocalAddress
If Len(stuff) = 0 Then
    stuff = "Error"
    End If
If Len(stuff) > TWOCOL Then
    stuff = Left$(stuff, TWOCOL)
    End If
msgout = msgout & stuff & Space(TWOCOL - Len(stuff)) & " Max Netlinks: " & UBound(net) & CRLF
stuff = mainForm.Socket1.LocalName
If Len(stuff) > TWOCOL Then
    stuff = Left$(stuff, TWOCOL)
    End If
msgout = msgout & "   Hostname:     " & stuff & Space(TWOCOL - Len(stuff)) & " Max Users: " & maxUsers & CRLF
stuff = mainForm.Socket1.version
If Len(stuff) > TWOCOL Then
    stuff = Left$(stuff, TWOCOL)
    End If
msgout = msgout & "   Winsock Ver:  " & stuff & Space(TWOCOL - Len(stuff)) & " Current Users: " & Trim$(usersOnline) & CRLF
stuff = mainForm.Socket1.LocalPort
If Len(stuff) > TWOCOL Then
    stuff = Left$(stuff, TWOCOL)
    End If
msgout = msgout & "   Local Port:   " & stuff & Space(TWOCOL - Len(stuff)) & " Netlink Port: " & Trim$(mainForm.Netlink(0).LocalPort) & CRLF
send msgout, usernum
End Sub

Sub ranksShow(usernum As Integer)
Dim count As Integer, msg As String, count2 As Integer
Dim thislev As Integer, sofar As Integer
msg = CRLF & "~FTRanks at " & system.talkerName & CRLF & "~FR"
msg = msg & String$(54, "-")
msg = msg & "~RS" & CRLF
For count = LBound(ranks) To UBound(ranks)
    If Not count = LBound(ranks) Then
        msg = msg & CRLF
        End If
    thislev = 0
    For count2 = 1 To NUM_OF_COMMANDS
        If cmds(count2).rank = count Then
            thislev = thislev + 1
            sofar = sofar + 1
            End If
        Next count2
    If user(usernum).rank = count Then
        msg = msg & "~FT"
        End If
    msg = msg & ranks(count) & Space(10 - Len(ranks(count)))
    msg = msg & " : Lev " & Str$(count) & Space(3 - Len(Str$(count)))
    msg = msg & " : " & mold(Trim$(sofar), 3) & " cmds total"
    msg = msg & " : " & Str$(thislev) & " this level"
    If user(usernum).rank = count Then
        msg = msg & "~RS"
        End If
    Next count
send msg & CRLF & CRLF, usernum
End Sub

Sub review(usernum As Integer, inpstr As String)
Dim count As Integer, roomnum As Integer, msg As String
Dim roomname As String
roomname = word(1)
If wordCount(inpstr) > 0 Then
    roomnum = getRoom(roomname)
    If roomnum = 0 Then
        send MSG_ROOM_NOT_EXIST, usernum
        Exit Sub
        End If
    'We only let the top brass review rooms that
    'they are not in. They dont even have to know
    'that this little feature exists.
    If user(usernum).rank < system.gatecrashLevel Then
        send "Usage: review" & CRLF, usernum
        Exit Sub
        End If
    Else
        roomnum = getRoom(user(usernum).room)
        End If
send "~BB~FG Review buffer for the " & rooms(roomnum).name & " " & CRLF, usernum
For count = 1 To REVBUFF_SIZE
    If Not rooms(roomnum).buffer(count) = "" Then
        msg = msg & rooms(roomnum).buffer(count) & CRLF
        End If
    Next count
If msg = "" Then
    msg = CRLF
    End If
send msg & "~BB~FG End of review buffer for the " & rooms(roomnum).name & " " & CRLF, usernum
End Sub

Sub home(usernum As Integer)
If user(usernum).atNetlink = -1 Then
    send "This command is for Netlink users only" & CRLF, usernum
    Exit Sub
    End If
netout "REL " & user(usernum).name & LF, user(usernum).atNetlink
returnedFromNetlink usernum
End Sub

Sub afk(usernum As Integer, inpstr As String)
If user(usernum).muzzled Then
    inpstr = vbNullString
    End If
user(usernum).afk = True
writeRoom user(usernum).room, "~FT" & user(usernum).name & " goes afk... " & inpstr & CRLF
End Sub

Sub connect_netlink(usernum As Integer, inpstr As String)
Dim msg As String, count As Integer, first As Boolean
Dim openSocket As Integer, netnum As Integer
If inpstr = "" Then
    msg = "Usage: connect <netlink>" & CRLF & "Services: "
    first = True
    For count = 0 To UBound(net)
        If Not net(count).name = "" Then
            If first Then
                msg = msg & net(count).name & CRLF
                first = False
                Else
                    msg = msg & Space(10) & net(count).name & CRLF
                    End If
            End If
        Next count
    send msg, usernum
    Exit Sub
    End If
netnum = -1
For count = 0 To UBound(net)
    If UCase$(net(count).name) = UCase$(inpstr) Then
        If Not net(count).state = NETLINK_DOWN Then
            send "This service active" & CRLF, usernum
            Exit Sub
            Else
                netnum = count
                Exit For
                End If
        End If
    Next count
If netnum = -1 Then
    msg = "Unknown service" & CRLF & "Services: "
    first = True
    For count = 0 To UBound(net)
        If Not net(count).name = "" Then
            If first Then
                msg = msg & net(count).name & CRLF
                first = False
                Else
                    msg = msg & Space(10) & net(count).name & CRLF
                    End If
            End If
        Next count
        send msg, usernum
    Exit Sub
    End If
connectNetlink netnum
End Sub

Sub clone(usernum As Integer, inpstr As String)
Dim clone_count As Integer, open_clone As Integer, count As Integer
Dim roomnum As Integer
If wordCount(inpstr) <> 1 Then
    send "Usage: clone <room>" & CRLF, usernum
    Exit Sub
    End If
roomnum = getRoom(word(1))
resizeClones True
For count = LBound(clones) To UBound(clones)
    If clones(count).owner = usernum And clones(count).active Then
        clone_count = clone_count + 1
        End If
    Next count
If clone_count >= MAX_USER_CLONES Then
    send "You are only allowed " & MAX_USER_CLONES & " clones" & CRLF, usernum
    Exit Sub
    End If
open_clone = -1
For count = LBound(clones) To UBound(clones)
    If Not clones(count).active Then
        open_clone = count
        Exit For
        End If
    Next count
For count = LBound(clones) To UBound(clones)
    If clones(count).owner = usernum And clones(count).room = roomnum Then
        send "You already have a clone in the " & rooms(roomnum).name & CRLF, usernum
        Exit Sub
        End If
    Next count
If open_clone = -1 Then
    send "Unable to make clone object" & CRLF, usernum
    Exit Sub
    End If
If roomnum <= 0 Then
    send MSG_ROOM_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
clones(open_clone).active = True
clones(open_clone).owner = usernum
clones(open_clone).room = roomnum
writeRoomExcept rooms(clones(open_clone).room).name, "~FMA clone of " & user(usernum).name & " appears before you" & CRLF, user(usernum).name
send "~FMYou create a clone in the " & rooms(roomnum).name & CRLF, usernum
End Sub

Sub destroy(usernum As Integer, inpstr As String)
Dim count As Integer, roomnum As Integer
If wordCount(inpstr) <> 1 Then
    send "Usage: destroy <room>" & CRLF, usernum
    Exit Sub
    End If
roomnum = getRoom(word(1))
If roomnum <= 0 Then
    send MSG_ROOM_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
For count = LBound(clones) To UBound(clones)
    If clones(count).owner = usernum And clones(count).room = roomnum And clones(count).active Then
        clones(count).active = False
        clones(count).owner = 0
        clones(count).room = 0
        writeRoomExcept "~FM" & rooms(roomnum).name, "A clone of " & user(usernum).name & " vanishes" & CRLF, user(usernum).name
        send "~FMYou destroy your clone in the " & rooms(roomnum).name & CRLF, usernum
        Exit Sub
        End If
    Next count
send "You do not have a clone in the " & rooms(roomnum).name & CRLF, usernum
resizeClones
End Sub

Sub cact(act_type As Boolean, usernum As Integer, inpstr As String)
Dim count As Integer, roomnum As Integer
If user(usernum).muzzled Then
    send MSG_USER_MUZZLED & CRLF, usernum
    Exit Sub
    End If
roomnum = getRoom(word(1))
If wordCount(inpstr) <= 1 Then
    If act_type Then
        send "Usage: csay <room> <message>" & CRLF, usernum
        Else
            send "Usage: cemote <room> <emotion>" & CRLF, usernum
            End If
    Exit Sub
    End If
If roomnum <= 0 Then
    send MSG_ROOM_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
For count = LBound(clones) To UBound(clones)
    If clones(count).owner = usernum And clones(count).room = roomnum And clones(count).active Then
        If act_type Then
            writeRoomExcept rooms(clones(count).room).name, "Clone of " & user(clones(count).owner).name & " says: " & stripOne(inpstr) & CRLF, user(usernum).name
            writeRoomBuff "Clone of " & user(clones(count).owner).name & " says: " & stripOne(inpstr), clones(count).room
            Else
                writeRoomExcept rooms(clones(count).room).name, "Clone of " & user(clones(count).owner).name & " " & stripOne(inpstr) & CRLF, user(usernum).name
                writeRoomBuff "Clone of " & user(clones(count).owner).name & " " & stripOne(inpstr), clones(count).room
                End If
        Exit Sub
        End If
    Next count
send "You do not have a clone in the " & rooms(roomnum).name & CRLF, usernum
End Sub

Sub disconnect_netlink(usernum As Integer, inpstr As String)
Dim count As Integer, msg As String, first As Boolean
If inpstr = "" Then
    msg = "Usage: disconnect <service name>" & CRLF & "Services: "
    first = True
    For count = 0 To UBound(net)
        If Not net(count).name = "" Then
            If first Then
                msg = msg & net(count).name & CRLF
                first = False
                Else
                    msg = msg & Space(10) & net(count).name & CRLF
                    End If
            End If
        Next count
    send msg, usernum
    Exit Sub
    End If
For count = 0 To UBound(net)
    If UCase$(net(count).name) = UCase$(inpstr) Then
        If net(count).state = NETLINK_DOWN Then
            send "This service is not active" & CRLF, usernum
            Exit Sub
            Else
                send "Disconnecting service " & net(count).name & "..." & CRLF, usernum
                End If
        dropNetlink n2s(count)
        Exit Sub
        End If
    Next count

msg = "Unknown service" & CRLF & "Services: "
first = True
For count = 0 To UBound(net)
    If Not net(count).name = "" Then
        If first Then
            msg = msg & net(count).name & CRLF
            first = False
            Else
                msg = msg & Space(10) & net(count).name & CRLF
                End If
        End If
    Next count
    send msg, usernum
End Sub

Sub echo_stuff(usernum As Integer, inpstr As String)
Dim count As Integer
If user(usernum).muzzled Then
    send MSG_USER_MUZZLED & CRLF, usernum
    Exit Sub
    End If
If wordCount(inpstr) = 0 Then
    send "Usage: echo <text>" & CRLF, usernum
    Exit Sub
    End If
End Sub

Sub show(usernum As Integer, inpstr As String)
If user(usernum).muzzled Then
    send MSG_USER_MUZZLED & CRLF, usernum
    Exit Sub
    End If
If wordCount(inpstr) < 1 Then
    send "Usage: show <text>" & CRLF, usernum
    Exit Sub
    End If
writeRoom user(usernum).room, "~OL~FTType--> ~RS" & inpstr & CRLF
End Sub
Sub think(usernum As Integer, inpstr As String)
If user(usernum).muzzled Then
    send MSG_USER_MUZZLED & CRLF, usernum
    Exit Sub
    End If
If wordCount(inpstr) < 1 Then
    send "Usage: think <thought>" & CRLF, usernum
    Exit Sub
    End If
writeRoomExcept user(usernum).room, user(usernum).name & " thinks .oO( " & inpstr & " ~RS)" & CRLF, user(usernum).name
send "You think .oO( " & inpstr & " ~RS)" & CRLF, usernum
End Sub

Sub semote(usernum As Integer, inpstr As String)
If user(usernum).muzzled Then
    send "You are muzzled and cannot shout emotes" & CRLF, usernum
    Exit Sub
    End If
If wordCount(inpstr) < 1 Then
    send "Usage: semote <emotion>" & CRLF, usernum
    Exit Sub
    End If
writeRoom "", "~OL!!~RS " & user(usernum).name & " " & inpstr & CRLF
End Sub

Sub verify(usernum As Integer, inpstr As String)
'Users enter the verification code they receive in their email
'and if it matches the record, forwarding is enabled.
If Not wordCount(inpstr) = 1 Then
    send "Usage: verify <verification code>" & CRLF, usernum
    Exit Sub
    End If
If user(usernum).sfVerifyed Then
    send "You have already verifyed" & CRLF, usernum
    Exit Sub
    End If
If Trim$(Str$(user(usernum).sfVercode)) = Trim$(word(1)) Then
    send "~OL~FG~LIVerification success!~RS" & CRLF & _
    " You may now use the .forward command to enable your smail forwarding to" & CRLF & _
    " your email address. Use .email to set your email address. When you reset" & CRLF & _
    " email addresses you will have to reverify." & CRLF, usernum
    user(usernum).sfVerifyed = True
    user(usernum).sfRec = True
    Else
        send "~OL~FRVerification failure~RS" & CRLF, usernum
        user(usernum).sfRec = False
        End If
End Sub

Sub forwarding(usernum As Integer)
If Not user(usernum).sfVerifyed Then
    send "You must use sendver and verify before receiving email forwards" & CRLF, usernum
    Exit Sub
    End If
If user(usernum).sfRec Then
    user(usernum).sfRec = False
    send "Your smail forwarding has been set to ~FR~OLOFF~RS" & CRLF, usernum
    Else
        user(usernum).sfRec = True
        send "Your smail forwarding has been set to ~FG~OLON~RS" & CRLF, usernum
        End If
End Sub

Sub sendver(usernum As Integer)
If user(usernum).sfVerifyed Then
    send "You have already verifyed" & CRLF, usernum
    Exit Sub
    End If
Dim mail_num As Integer
mail_num = Setup_Address(user(usernum).name)
If mail_num < 0 Then
    Select Case mail_num
        Case -1
            send "There was an error loading your user" & CRLF, usernum
        Case -2
            send "Mail system too busy... Try again later" & CRLF, usernum
        Case -3
            send "Invalid email address: " & user(usernum).email & CRLF, usernum
        Case Else
            send "There was an error preparing the validation" & CRLF, usernum
        End Select
    Exit Sub
    Else
        send "Sending verification to " & user(usernum).email & CRLF, usernum
        End If
With mail(mail_num)
    .inuse = True
    .u_email = user(usernum).email
    .u_to = user(usernum).name
    .u_from = system.emailAddress
    .userid = "Verification System"
    .timestamp = Now
    .message = "This is a verification message for your " & _
    "email forwarding services on " & system.talkerName & ". " & _
    CRLF & "Your verification code is: " & user(usernum).sfVercode & _
    CRLF & CRLF & "When you are on " & system.talkerName & " next " & _
    "time you should use this code with the verify command." & CRLF & _
    "Example: verify " & user(usernum).sfVercode & CRLF & CRLF & _
    "Thank you," & CRLF & "The " & system.talkerName & " Admin" & CRLF
    End With
End Sub

Sub email(usernum As Integer, inpstr As String)
If Not inpstr = "" Then
    If Not InStr(inpstr, "@") > 1 And InStr(inpstr, "@") < Len(inpstr) Then
        send "Bad email address" & CRLF, usernum
        Exit Sub
        End If
    If inpstr = user(usernum).email Then
        send "Your address is already set to " & user(usernum).email & CRLF, usernum
        Exit Sub
        End If
    user(usernum).sfRec = False
    user(usernum).sfVerifyed = False
    user(usernum).sfVercode = gen_ver_code
    user(usernum).email = inpstr
    send "Your email address has been set to " & inpstr & CRLF & "You may now use the .sendver command to send a verification code to this address. Once you receive the code, use .verify so that you can receive email forwarding of all your smail." & CRLF, usernum
    Else
        send "Your email address is set to " & user(usernum).email & CRLF, usernum
        End If
End Sub

Sub mailsys(usernum As Integer)
Dim msg As String, count As Integer, t As String, found As Boolean
Dim everyOther As Boolean
everyOther = True
msg = msg & "--------------------------------------------------------------------" & CRLF
msg = msg & "~FROwner~RS        | ~FRQ#~RS | ~FRTo~RS           | ~FREmail~RS                              " & CRLF
For count = LBound(mail) To UBound(mail)
    If everyOther Then
        msg = msg & "~FT"
        Else
            msg = msg & "~RS"
            End If
    If mail(count).inuse Then
        found = True
        If Len(mail(count).userid) > 12 Then
            t = Left$(mail(count).userid, 12)
            Else
                t = mail(count).userid & Space$(12 - Len(mail(count).userid))
                End If
        msg = msg & t & "~RS | "
        If everyOther Then
            msg = msg & "~FT"
            Else
                msg = msg & "~RS"
                End If
        t = Trim$(Str$(count))
        If Len(t) > 2 Then
            t = Left$(t, 2)
            Else
                t = t & Space$(2 - Len(t))
                End If
        msg = msg & t & "~RS | "
        If everyOther Then
            msg = msg & "~FT"
            Else
                msg = msg & "~RS"
                End If
        If Len(mail(count).u_to) > 12 Then
            t = Left$(mail(count).u_to, 12)
            Else
                t = mail(count).u_to & Space$(12 - Len(mail(count).u_to))
                End If
        msg = msg & t & "~RS | "
        If everyOther Then
            msg = msg & "~FT"
            Else
                msg = msg & "~RS"
                End If
        If Len(mail(count).u_email) > 33 Then
            t = Left$(mail(count).u_email, 33)
            Else
                t = mail(count).u_email & Space$(33 - Len(Str$(count)))
                End If
        msg = msg & t & CRLF
        If everyOther Then
            everyOther = False
            Else
                send msg, usernum
                msg = ""
                everyOther = True
                End If
        End If
    Next count
If found = False Then
    msg = "~RSThe mail queue is empty" & CRLF
    Else
        msg = msg & "~RS--------------------------------------------------------------------" & CRLF
        End If
send msg, usernum
End Sub

Sub kill_job(usernum As Integer, inpstr As String)
If Not wordCount(inpstr) = 1 Then
    send "Usage: kjob <email queue number>" & CRLF, usernum
    Exit Sub
    End If
If Val(word(1)) < 0 Or Val(word(1)) > MAX_MAIL_SLOTS Then
    send "Invalid job number" & CRLF, usernum
    Exit Sub
    End If
If mail(Val(word(1))).inuse Then
    mail(Val(word(1))).inuse = False
    send "Mail job " & Trim$(Str$(Val(word(1)))) & " killed" & CRLF, usernum
    Else
        send "Job not active" & CRLF, usernum
        End If
End Sub

Sub topic(usernum As Integer, inpstr As String)
If wordCount(inpstr) = 0 Then
    send "Usage: topic <new topic>" & CRLF, usernum
    Exit Sub
    End If
If Len(inpstr) < 40 Then
    rooms(getRoom(user(usernum).room)).topic = inpstr
    writeRoom user(usernum).room, "Topic set to: " & inpstr & CRLF
    Else
        send "Topic too long" & CRLF, usernum
        End If
End Sub

Sub purge(usernum As Integer)
Dim purger As PURGE_OBJECT
send CRLF & "~OL~LI~FRPurging userfiles..." & CRLF, usernum
purger = runPurge(purger)
send "~OL~UL~FGPurge Completed" & CRLF & Trim$(Str$( _
    purger.usersRemoved)) & " users were expunged" & CRLF & "There are now " & _
    Trim$(Str$(purger.usersNow)) & " user accounts" & CRLF & CRLF, usernum
End Sub

Sub rstat(usernum As Integer, inpstr As String)
Dim count As Integer, msg As String, first As Boolean
If inpstr = "" Then
    msg = "Usage: rstat <service name>" & CRLF & "Services: "
    first = True
    For count = 0 To UBound(net)
        If Not net(count).name = "" Then
            If first Then
                msg = msg & net(count).name & CRLF
                first = False
                Else
                    msg = msg & Space(10) & net(count).name & CRLF
                    End If
            End If
        Next count
    send msg, usernum
    Exit Sub
    End If
For count = 0 To UBound(net)
    If UCase$(net(count).name) = UCase$(inpstr) Then
        If net(count).state = NETLINK_DOWN Then
            send "This service is not active" & CRLF, usernum
            Exit Sub
            Else
                send "Requesting remote statistics..." & CRLF, usernum
                End If
        netout "RSTAT " & user(usernum).name & LF, count + 1
        Exit Sub
        End If
    Next count

msg = "Unknown service" & CRLF & "Services: "
first = True
For count = 0 To UBound(net)
    If Not net(count).name = "" Then
        If first Then
            msg = msg & net(count).name & CRLF
            first = False
            Else
                msg = msg & Space(10) & net(count).name & CRLF
                End If
        End If
    Next count
    send msg, usernum
End Sub

Sub add_history(usernum As Integer, inpstr As String)
If wordCount(inpstr) < 2 Then
    send "Usage: addhist <username> <comment>" & CRLF, usernum
    Exit Sub
    End If
If Not userExists(word(1)) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
'If Not userIsOnline(word(1)) Then
'    send MSG_USER_NOT_ONLINE & CRLF, usernum
'    Exit Sub
'    End If
If UCase$(user(usernum).name) = UCase$(word(1)) Then
    send "You cannot add to your own history" & CRLF, usernum
    Exit Sub
    End If
writeHistory word(1), "~FB" & user(usernum).name & " adds: ~RS" & stripOne(inpstr)
send "You add a comment to " & userCap(word(1)) & "'s history file" & CRLF, usernum
End Sub

Sub swban(usernum As Integer)
If system.swearing >= swbanLevels.SWEAR_MAX Then
    system.swearing = swbanLevels.SWEAR_OFF
    Else
        system.swearing = system.swearing + 1
        End If
writeRoom "", "Swearing ban now set to " & swearNames(system.swearing) & CRLF
End Sub

Sub version(usernum As Integer)
Dim verinfo As String
verinfo = "~OL~LI~FTSouthWest Version " & App.Major & "." & App.Minor & "." & _
    App.Revision & "~RS" & CRLF & "Programmed by Scott Lloyd" & _
    CRLF & "Inspired by the " & "work of Neil Robertson and Andrew Collington" & CRLF
send CRLF & verinfo & CRLF, usernum
End Sub

Sub charecho(usernum As Integer)
Dim msg As String
user(usernum).charEchoing = Not user(usernum).charEchoing
If user(usernum).charEchoing Then
    msg = "~FGON"
    Else
        msg = "~FROFF"
        End If
send "Your charictor echoing mode is now set to " & msg & CRLF, usernum
End Sub

Sub list_rooms(usernum As Integer)
Dim msg As String, count As Integer, five As Byte
For count = 0 To UBound(rooms)
    If rooms(count).name = "" Then
        Exit For
        End If
    msg = msg & rooms(count).name & Space(15 - Len(rooms(count).name))
    If rooms(count).locked Then
        msg = msg & mold("~FRFIXED~RS", 10)
        Else
            msg = msg & mold("~FGUNFIXED~RS", 10)
            End If
    Select Case rooms(count).access
        Case ROOM_PRIVATE
            msg = msg & "~FRPRIVATE "
        Case ROOM_PUBLIC
            msg = msg & "~FGPUBLIC  "
            Case Else
                msg = msg & "~FYSTAFF   "
                End Select
    msg = msg & "~RS"
    msg = msg & rooms(count).topic & CRLF
    five = five + 1
    If five = 5 Then
        five = 0
        send msg, usernum
        msg = ""
        End If
    Next count
If Not five = 0 Then
    send msg, usernum
    End If
End Sub

Sub set_public(usernum As Integer)
'This function may be called automaticly from go
Dim uroom As Integer, count As Integer
uroom = getRoom(user(usernum).room)
If rooms(uroom).locked Then
    If usernum > 0 Then
        send "This room's access is fixed" & CRLF, usernum
        End If
    Else
        If rooms(uroom).access = ROOM_PUBLIC Then
            If usernum > 0 Then
                send "This room is already set to ~FGPUBLIC" & CRLF, usernum
                End If
            Else
                writeRoom user(usernum).room, "Room set to ~FGPUBLIC" & CRLF
                rooms(uroom).access = ROOM_PUBLIC
                cbuff 0
                For count = 1 To UBound(user)
                    If isUserInvited(count, user(usernum).room) Then
                        uninvite 0, user(count).name
                        End If
                    Next count
                End If
        End If
End Sub

Sub set_private(usernum As Integer)
Dim uroom As Integer
uroom = getRoom(user(usernum).room)
If rooms(uroom).locked = True Then
    send "This room's access is fixed" & CRLF, usernum
    Else
        If rooms(uroom).access = ROOM_PRIVATE Then
            send "This room is already set to ~FRPRIVATE" & CRLF, usernum
            Else
                writeRoom user(usernum).room, "Room set to ~FRPRIVATE" & CRLF
                rooms(uroom).access = ROOM_PRIVATE
                End If
        End If
End Sub

Sub fix_room(usernum As Integer)
Dim uroom As Integer
uroom = getRoom(user(usernum).room)
If rooms(uroom).locked Then
    send "This room's access is already fixed" & CRLF, usernum
    Else
        rooms(uroom).locked = True
        writeRoom user(usernum).room, "Access for this room has been ~FRFIXED" & CRLF
        End If
End Sub

Sub unfix_room(usernum As Integer)
Dim uroom As Integer
uroom = getRoom(user(usernum).room)
If Not rooms(uroom).locked Then
    send "This room's access is already unfixed" & CRLF, usernum
    Else
        rooms(uroom).locked = False
        writeRoom user(usernum).room, "Access for this room has been ~FGUNFIXED" & CRLF
        End If
End Sub

Sub view_history(usernum As Integer, inpstr As String)
Dim five As Byte, FromFile As String, msg As String
If wordCount(inpstr) <> 1 Then
    send "Usage: history <user>" & CRLF, usernum
    Exit Sub
    End If
If containsCorrupt(inpstr) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
If Dir$(App.Path & "\Users\" & word(1) & ".His") = "" Then
    send "Could not find a history file for " & userCap(word(1)) & CRLF, usernum
    Exit Sub
    End If
send CRLF & "~FYThe History of " & userCap(word(1)) & CRLF & String$(15 + Len(word(1)), "=") & CRLF, usernum
Open App.Path & "\Users\" & word(1) & ".His" For Input As #1
Do While Not EOF(1)
    Line Input #1, FromFile
    msg = msg & FromFile & CRLF
    five = five + 1
    If five = 5 Then
        send msg, usernum
        msg = ""
        five = 0
        End If
    Loop
Close #1
If five > 0 Then
    msg = msg & CRLF
    Else
        msg = msg & CRLF & CRLF
        End If
send msg, usernum
End Sub

Sub netstat(usernum As Integer)
'Lists netlink connections
Dim msg As String, count As Integer, everyOther As Boolean, found As Boolean
Dim count2 As Integer, linkUsers As Integer
msg = "~FRName         Status      Access IU OU SW  Version Site~RS" & CRLF
For count = 0 To UBound(net)
    If Not net(count).name = "" Then
        If everyOther Then
            msg = msg & "~RS"
            Else
                msg = msg & "~FT"
                End If
        everyOther = Not everyOther
        msg = msg & mold(net(count).name, 12) & " " & mold(netlinkStates(Int(net(count).state)), 11) & " "
        Select Case net(count).access
            Case NETACCESS.ACCESS_IN
                msg = msg & "In     "
            Case NETACCESS.ACCESS_ALL
                msg = msg & "Both   "
            Case NETACCESS.ACCESS_OUT
                msg = msg & "Out    "
            Case NETACCESS.ACCESS_DENIED
                msg = msg & "Denied "
            End Select
        linkUsers = 0
        For count2 = 1 To UBound(user)
            If user(count2).operational And user(count2).netlinkType And s2n(user(count2).netlinkFrom) = count Then
                linkUsers = linkUsers + 1
                End If
            Next count2
        msg = msg & mold(Trim$(linkUsers), 2) & " "
        linkUsers = 0
        For count2 = 1 To UBound(user)
            If user(count2).operational And user(count2).atNetlink >= 0 Then
                If user(count2).atNetlink = n2s(count) Then
                    linkUsers = linkUsers + 1
                    End If
                End If
            Next count2
        msg = msg & mold(Trim$(linkUsers), 2) & " "
        If net(count).state = NETLINK_UP Then
            If net(count).southwest Then
                msg = msg & "Yes "
                Else
                    msg = msg & "No  "
                    End If
            Else
                msg = msg & "N/A "
                End If
        If net(count).state = NETLINK_UP Then
            msg = msg & mold(net(count).version, 7) & " "
            Else
                msg = msg & "N/A     "
                End If
        msg = msg & RTrim$(mold(net(count).site & ":" & Trim$(Str$(net(count).port)), 32))
        send msg & CRLF, usernum
        msg = ""
        found = True
        End If
    Next count
If Not found Then
    send "There are no Netlinks on this system" & CRLF, usernum
    Else
        send CRLF, usernum
        End If
End Sub

Sub figlet(usernum As Integer, inpstr As String)
Dim count As Integer, msg As String, count2 As Integer
If inpstr = "" Or Trim$(LCase$(inpstr)) = "(none)" Then
    msg = "Usage: figlet <figlet font>" & CRLF & "~FTFiglets loaded~FR: ~RS"
    For count = mainForm.FigletList.LBound To mainForm.FigletList.UBound
        If count > mainForm.FigletList.LBound Then
            msg = msg & Space(16)
            End If
        msg = msg & mainForm.FigletList(count).Caption
        If mainForm.FigletList(count).Checked Then
            msg = msg & " ~FY(~FBcurrent~FY)"
            End If
        msg = msg & CRLF
        Next count
    send msg, usernum
    Exit Sub
    End If
For count = mainForm.FigletList.LBound To mainForm.FigletList.UBound
    If UCase$(inpstr) = UCase$(mainForm.FigletList(count).Caption) Then
        For count2 = mainForm.FigletList.LBound To mainForm.FigletList.UBound
            mainForm.FigletList(count2).Checked = False
            Next count2
        mainForm.FigletList(count).Checked = True
        system.figlet = mainForm.FigletList(count).Caption
        loadFiglets
        send "Now using the " & mainForm.FigletList(count).Caption & " figlet font for .greet" & CRLF, usernum
        Exit Sub
        End If
    Next count
send "This figlet font was not found." & CRLF & "Use ~FT.figlet~RS to get a list of installed figlet fonts." & CRLF, usernum
End Sub

Sub pager(usernum As Integer, inpstr As String)
Dim newPager As Integer
If wordCount(inpstr) = 0 Then
    send "Your pager is currently set to ~FM~OL" & Trim$(user(usernum).pager) & CRLF, usernum
    Exit Sub
    End If
newPager = Int(Val(inpstr))
If newPager < 12 Or newPager > 99 Then
    send "Pager value must be greater than 11 and less than 100 " & CRLF, usernum
    Exit Sub
    End If
user(usernum).pager = newPager
send "Pager now set to ~FM~OL" & Trim$(Str$(newPager)) & CRLF, usernum
End Sub

Sub ban(usernum As Integer, inpstr As String)
Dim banMethod As banTypes, file As String, free As Integer
If wordCount(inpstr) <> 2 Then
    send "Usage: ban user/site/new <user/site>" & CRLF, usernum
    Exit Sub
    End If
Select Case LCase$(word(1))
    Case "user"
        banMethod = BAN_USER
    Case "site"
        banMethod = BAN_SITE
    Case "new"
        banMethod = BAN_NEW
    Case Else
        send "Usage: ban user/site/new <user/site>" & CRLF, usernum
        Exit Sub
        End Select
If isBanned(word(2), banMethod) Then
    If banMethod = BAN_USER Then
        send "This user is already banned" & CRLF, usernum
        Exit Sub
        Else
            send "This site is already banned" & CRLF, usernum
            Exit Sub
            End If
    End If
file = getBanFile(banMethod)
free = FreeFile
If Dir$(App.Path & "\Misc", vbDirectory) = "" Then
    MkDir (App.Path & "\Misc")
    End If
Open file For Append As #free
Print #free, word(2) & CRLF;
Close #free
Select Case banMethod
    Case BAN_USER
        send "User, " & userCap(word(2)) & ", banned" & CRLF, usernum
    Case BAN_SITE
        send "All users from " & word(2) & " banned" & CRLF, usernum
    Case BAN_NEW
        send "New users from " & word(2) & " banned" & CRLF, usernum
    End Select
End Sub

Sub unban(usernum As Integer, inpstr As String)
Dim banMethod As banTypes, file As String, free As Integer, found As Boolean
Dim lineCount As Integer, fileLines() As String, FromFile
If wordCount(inpstr) <> 2 Then
    send "Usage: unban user/site/new <user/site>" & CRLF, usernum
    Exit Sub
    End If
Select Case LCase$(word(1))
    Case "user"
        banMethod = BAN_USER
    Case "site"
        banMethod = BAN_SITE
    Case "new"
        banMethod = BAN_NEW
    Case Else
        send "Usage: unban user/site/new <user/site>" & CRLF, usernum
        Exit Sub
        End Select
file = getBanFile(banMethod)
If Not fileExists(file) Then
    If banMethod = BAN_USER Then
        send userCap(word(2)) & " is not banned" & CRLF, usernum
        Else
            send word(2) & " is not banned" & CRLF, usernum
            End If
    Exit Sub
    End If
free = FreeFile
Open file For Input As #free
Do While Not EOF(free)
    Line Input #free, FromFile
    lineCount = lineCount + 1
    Loop
Close #free
ReDim fileLines(lineCount)
lineCount = 0
Open file For Input As #free
Do While Not EOF(free)
    Line Input #free, FromFile
    If LCase$(FromFile) Like LCase$(word(2)) Then
        found = True
        Else
            fileLines(lineCount) = FromFile
            lineCount = lineCount + 1
            End If
    Loop
If Not found Then
    If banMethod = BAN_USER Then
        send userCap(word(2)) & " is not banned" & CRLF, usernum
        Else
            send word(2) & " is not banned" & CRLF, usernum
            End If
    Exit Sub
    End If
Close #free
Open file For Output As #free
For lineCount = LBound(fileLines) To UBound(fileLines)
    If Len(fileLines(lineCount)) > 0 Then
        Print #free, fileLines(lineCount) & CRLF;
        End If
    Next lineCount
Close #free
Select Case banMethod
    Case BAN_USER
        send "User, " & userCap(word(2)) & ", unbanned" & CRLF, usernum
    Case BAN_SITE
        send "All users from " & word(2) & " unbanned" & CRLF, usernum
    Case BAN_NEW
        send "New users from " & word(2) & " unbanned" & CRLF, usernum
    End Select
End Sub

Sub lban(usernum As Integer, inpstr As String)
Dim file As String, msg As String, banCount As Integer, free As Integer, tmp As String
Dim FromFile As String, everyOther As Boolean, last As String, banMethod As banTypes
If wordCount(inpstr) <> 1 Then
    send "Usage: lban users/sites/new" & CRLF, usernum
    Exit Sub
    End If
Select Case LCase$(word(1))
    Case "users"
        banMethod = BAN_USER
    Case "sites"
        banMethod = BAN_SITE
    Case "new"
        banMethod = BAN_NEW
    Case Else
        send "Usage: lban user/site/new" & CRLF, usernum
        Exit Sub
        End Select
file = getBanFile(banMethod)
If Not fileExists(file) Then
    Select Case banMethod
        Case BAN_USER
            send "No users are currently banned" & CRLF, usernum
        Case BAN_SITE
            send "No sites are currently being blocked" & CRLF, usernum
        Case BAN_NEW
            send "No sites are currently restricted" & CRLF, usernum
        End Select
    Exit Sub
    End If
msg = FANCY_BAR & "~OL~FT "
Select Case banMethod
    Case BAN_USER
        msg = msg & "List of banned users"
    Case BAN_SITE
        msg = msg & "List of blocked sites"
    Case BAN_NEW
        msg = msg & "List of restricted"
    End Select
msg = msg & CRLF & FANCY_BAR
free = FreeFile
Open file For Input As #free
Do While Not EOF(free)
    Line Input #free, FromFile
    banCount = banCount + 1
    If Len(FromFile) > 31 Then
        FromFile = Left$(FromFile, 31)
        End If
    tmp = Replace(FromFile, "*", "~FR*~RS")
    tmp = Replace(tmp, "?", "~FY?~RS")
    If everyOther Then
        If banMethod = BAN_USER Then
            msg = msg & Space(32 - Len(last)) & userCap(Trim$(tmp))
            Else
                msg = msg & Space(32 - Len(last)) & Trim$(tmp)
                End If
        Else
            If banMethod = BAN_USER Then
                msg = msg & Space(5) & userCap(Trim$(tmp))
                Else
                    msg = msg & Space(5) & Trim$(tmp)
                    End If
            End If
    If everyOther Then
        msg = msg & CRLF
        End If
    everyOther = Not everyOther
    last = FromFile
    If Len(msg) > SEND_CHOP Then
        send msg, usernum
        msg = ""
        End If
    Loop
If everyOther Then
    msg = msg & CRLF
    End If
Close #free
msg = msg & "Total of ~FM~OL" & Trim$(banCount) & "~RS "
If banMethod = BAN_USER Then
    msg = msg & "users banned" & CRLF
    Else
        msg = msg & "sites banned" & CRLF
        End If
If banCount = 0 Then
    Select Case banMethod
        Case BAN_USER
            msg = "No users are currently banned" & CRLF
        Case BAN_SITE
            msg = "No sites are currently being blocked" & CRLF
        Case BAN_NEW
            msg = "No sites are currently restricted" & CRLF
            End Select
    End If
send msg, usernum
End Sub

Sub rules(usernum As Integer)
Dim free As Integer, msg As String, FromFile As String
If Not fileExists(App.Path & "\Misc\Rules.S") Then
    send "Could not locate the rules file on this system" & CRLF, usernum
    Exit Sub
    End If
free = FreeFile
Open App.Path & "\Misc\Rules.S" For Input As #free
Do While Not EOF(free)
    Input #free, FromFile
    msg = msg & FromFile & CRLF
    If Len(msg) > SEND_CHOP Then
        send msg, usernum
        msg = vbNullString
        End If
    Loop
Close #free
If Len(msg) > 0 Then
    send msg, usernum
    End If
End Sub

Sub greet(usernum As Integer, text As String)
Dim maxFigletChars As Integer
maxFigletChars = CLIENT_WIDTH / (figletWidth + 1) + 1
If Len(text) < 1 Then
    send "Usage: greet <message>" & CRLF, usernum
    Exit Sub
    End If
If Len(text) > maxFigletChars Then
    send "Message must be " & maxFigletChars & " charitors or less" & CRLF, usernum
    Exit Sub
    End If
If Len(figbar) = 0 Then
    send "No figlet font is loaded" & CRLF, usernum
    Exit Sub
    End If
Dim count As Integer, charloc As Integer, chardraw() As String
Dim selchar As String, figouts(), count2 As Integer, msg As String
ReDim figouts(1 To figletHeight)
ReDim chardraw(1 To figletHeight)
text = UCase$(text)
For count = 1 To Len(text)
    selchar = Mid$(text, count, 1)
    charloc = InStr(figbar, selchar)
    If Not charloc = 0 Then
        For count2 = 1 To figletHeight
            figouts(count2) = figouts(count2) & figlets(charloc - 1, count2) & "  "
            Next count2
        Else
            If selchar = " " Then
                For count2 = 1 To figletHeight
                    figouts(count2) = figouts(count2) & Space(figletWidth)
                    Next count2
                End If
            End If
    Next count
If Len(figouts(1)) > 0 Then
    For count = 1 To figletHeight
        msg = msg & figouts(count) & CRLF
        If Len(msg) > SEND_CHOP Then
            writeRoom "", msg
            msg = ""
            End If
        Next count
    If Len(msg) > 0 Then
        writeRoom "", msg
        End If
    Else
        send "Nothing readable was entered" & CRLF, usernum
        End If
End Sub

Sub ustat(usernum As Integer, ByVal inpstr As String)
Dim msg As String, un As Integer
'If they don't give an argument, make the argument their
'own name (like on Amnuts210)
If wordCount(inpstr) = 0 Then
    word(1) = user(usernum).name
    inpstr = user(usernum).name
    End If
un = getUser(word(1))
If Not userExists(word(1)) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
msg = embedBar("User Info", "~FY", "~FB", "~FY")
msg = msg & mold("~FYName~FB:~RS " & user(un).name & " " & user(un).desc, 50) & _
    "~FYRank~FB:~RS " & ranks(user(un).rank) & CRLF
msg = msg & mold("~FYGender~FB:~RS " & user(un).gender, 50) & "~FYAge~FB:~RS " & user(un).age & CRLF
msg = msg & mold("~FYEmail~FB:~RS " & user(un).email, 50) & "~FYICQ~FB:~RS " & user(un).ICQ & CRLF & "~FYHomepage~FB:~RS " & user(un).url & CRLF
msg = msg & "~FYTotal Login~FB:~RS " & deriveTimeString(spliceTime(user(un).totalTime), False) & CRLF
msg = msg & embedBar("General Info", "~FY", "~FB", "~FY")
msg = msg & "~FYEnter Message~FB:~RS " & user(un).enterMsg & CRLF & "~FYExit Message~FB:~RS " & user(un).exitMsg & CRLF
msg = msg & mold("~FYNew Mail~FB:~RS " & bool2YN(user(un).unread), 20) & CRLF
send CRLF & msg, usernum 'Send in segments so as to not overflow the
msg = vbNullString       'buffers of some clients.
msg = msg & embedBar("Personal Info", "~FY", "~FB", "~FY")
msg = msg & mold("~FYPager~FB:~RS " & user(un).pager, 20) & mold("~FYAutofwd~FB:~RS " & bool2YN(user(un).sfRec), 20)
msg = msg & mold("~FYVerified~FB:~RS " & bool2YN(user(un).sfVerifyed), 20) & CRLF
msg = msg & mold("~FYChar echo~FB:~RS " & bool2YN(user(un).charEchoing), 20)
msg = msg & "~FYLogins~FB:~RS " & user(un).logins & CRLF
If user(usernum).rank >= staffLevel Then
    msg = msg & embedBar("Staff Only Info", "~OL~FY", "~FB", "~FY")
    msg = msg & "~FYLast site~FB:~RS " & user(un).site & CRLF
    msg = msg & mold("~FYArrested~FB:~RS " & bool2YN(user(un).arrested), 20)
    msg = msg & mold("~FYVisible~FB:~RS " & bool2YN(user(un).visible), 20)
    If user(un).arrested Then
        msg = msg & "~FYUnarrest Lev~FB~RS " & user(un).unarrestLevel & CRLF
        End If
    End If
    msg = msg & embedBar(, "~FY", "~FB", "~FY")
    send msg & CRLF, usernum
End Sub

Sub setIcq(usernum As Integer, inpstr As String)
If wordCount(inpstr) <> 1 Then
    send "Usage: icq unset/<ICQ number>" & CRLF, usernum
    Exit Sub
    End If
If LCase$(inpstr) = "unset" Then
    user(usernum).ICQ = "Unset"
    send "Your ICQ number is now unset" & CRLF, usernum
    Exit Sub
    End If
If containsCorruptNumsOnly(inpstr) Then
    send "Your ICQ number may be a number only" & CRLF, usernum
    Exit Sub
    End If
If Len(inpstr) > 9 Then
    send "ICQ numbers may only be up to nine digits" & CRLF, usernum
    Exit Sub
    End If
user(usernum).ICQ = Trim$(inpstr)
send "Your ICQ number has been set to " & user(usernum).ICQ & CRLF, usernum
End Sub

Sub setHomepage(usernum As Integer, inpstr As String)
If wordCount(inpstr) <> 1 Then
    send "Usage: homepage <url>" & CRLF, usernum
    Exit Sub
    End If
If Len(inpstr) > 60 Then
    send "URL too long" & CRLF, usernum
    Exit Sub
    End If
user(usernum).url = inpstr
send "Your homepage has been set to " & user(usernum).url & CRLF, usernum
End Sub

Sub expire(usernum As Integer, inpstr As String)
Dim un As Integer
If wordCount(inpstr) <> 1 Then
    send "Usage: expire <user>" & CRLF, usernum
    Exit Sub
    End If
If userIsOnline(inpstr) Then
    un = getUser(inpstr)
    user(un).expires = Not user(un).expires
    Else
        un = 0
        user(0).name = userCap(inpstr)
        loadUserData un
        user(un).expires = Not user(un).expires
        saveUserData user(0)
        End If
If user(un).expires Then
    send "You have set it so that " & user(un).name & " ~OL~FMWILL~RS expire with purge" & CRLF, usernum
    writeHistory user(un).name, user(usernum).name & " ~OL~FMENABLES~RS expiration with purge"
    Else
       send "You have set it so that " & user(un).name & " ~OL~FMWILL NOT~RS expire with purge" & CRLF, usernum
       writeHistory user(un).name, user(usernum).name & " ~OL~FMDISABLES~RS expiration with purge"
       End If
End Sub

Sub getTime(usernum As Integer)
send "~OL~FTTime: ~RS" & Format$(Now, "dddd d") & getOrdinal(Int(Format$(Now, "d"))) & Format$(Now, " mmmm yyyy" & " at " & Format$(Now, "hh:nn")) & CRLF, usernum
End Sub

Sub myClones(usernum As Integer)
Dim count As Integer, msg As String, found As Boolean, everyOther As Boolean
Dim col As String
msg = "~FRYou have clones in:" & CRLF
For count = LBound(clones) To UBound(clones)
    If clones(count).owner = usernum Then
        found = True
        If everyOther Then
            col = "~FW"
            Else
                col = "~FT"
                End If
        everyOther = Not everyOther
        msg = msg & Space(5) & col & rooms(clones(count).room).name & "~RS" & CRLF
        End If
    Next count
If Not found Then
    msg = "You have no active clones" & CRLF
    End If
send msg, usernum
End Sub

Sub allClones(usernum As Integer)
Dim count As Integer, msg As String, found As Boolean, everyOther As Boolean
Dim col As String
msg = "~FRThere are clones in:~RS" & CRLF
For count = LBound(clones) To UBound(clones)
    If clones(count).active Then
        found = True
        If everyOther Then
            col = "~FW"
            Else
                col = "~FT"
                End If
        everyOther = Not everyOther
        msg = msg & Space(5) & col & mold(user(clones(count).owner).name, MAX_NAME_LEN) & rooms(clones(count).room).name & "~RS" & CRLF
        End If
    Next count
If Not found Then
    msg = "There are no active clones" & CRLF
    End If
send msg, usernum
End Sub

Sub dmail(usernum As Integer, inpstr As String)
Dim file As String
file = App.Path & "\Users\" & user(usernum).name & ".M"
If Not fileExists(file) Then
    send "You do not have any mail" & CRLF, usernum
    Exit Sub
    End If
If Not deleteMessages(file, inpstr) Then
    send "Usage: dmail all" & CRLF & _
         "Usage: dmail <#>" & CRLF & _
         "Usage: dmail to <#>" & CRLF & _
         "Usage: dmail from <#> to <#>" & CRLF, usernum
    Else
        send "Done!" & CRLF, usernum
        End If
End Sub

Sub wipe(usernum As Integer, inpstr As String)
Dim file As String
file = App.Path & "\Rooms\" & user(usernum).room & ".B"
If Not fileExists(file) Then
    send "There are no messages on this board" & CRLF, usernum
    Exit Sub
    End If
If Not deleteMessages(file, inpstr) Then
    send "Usage: wipe all" & CRLF & _
         "Usage: wipe <#>" & CRLF & _
         "Usage: wipe to <#>" & CRLF & _
         "Usage: wipe from <#> to <#>" & CRLF, usernum
    Else
        writeRoomExcept user(usernum).room, user(usernum).name & " wipes some messages from the board" & CRLF, user(usernum).name
        send "Done!" & CRLF, usernum
        End If
End Sub

Sub arrest(usernum As Integer, inpstr As String)
Dim un As Integer
If Not wordCount(inpstr) = 1 Then
    send "Usage: arrest <user>" & CRLF, usernum
    Exit Sub
    End If
If Not userExists(inpstr) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
inpstr = userCap(inpstr)
If userIsOnline(inpstr) Then
    un = getUser(inpstr)
    Else
        un = 0
        user(0).name = inpstr
        loadUserData 0
        End If
If user(un).name = user(usernum).name Then
    send "You cannot arrest yourself" & CRLF, usernum
    Exit Sub
    End If
If user(un).arrested Then
    send inpstr & " is already under arrest" & CRLF, usernum
    Exit Sub
    End If
If user(un).rank >= user(usernum).rank Then
    send "You cannot arrest a user of an equal or greater rank" & CRLF, usernum
    If Not un = 0 Then
        send "~FR~OL" & user(usernum).name & " tried to arrest you~RS" & CRLF, un
        End If
    Exit Sub
    End If
user(un).unarrestLevel = user(un).rank
user(un).rank = 0
user(un).arrested = True
user(un).room = rooms(0).name
send "~FM" & user(un).name & " has been placed under arrest~RS" & CRLF, usernum
If un > 0 Then
    writeRoomExcept "", "~FM" & user(un).name & " has been arrested by " & user(usernum).name & "~RS" & CRLF, user(usernum).name
    look un
    End If
End Sub

Sub unarrest(usernum As Integer, inpstr As String)
Dim un As Integer
If Not wordCount(inpstr) = 1 Then
    send "Usage: unarrest <user>" & CRLF, usernum
    Exit Sub
    End If
If Not userExists(inpstr) Then
    send MSG_USER_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
inpstr = userCap(inpstr)
If userIsOnline(inpstr) Then
    un = getUser(inpstr)
    Else
        un = 0
        user(0).name = inpstr
        loadUserData 0
        End If
If user(un).name = user(usernum).name Then
    send "You cannot unarrest yourself" & CRLF, usernum
    Exit Sub
    End If
If Not user(un).arrested Then
    send user(un).name & " is not under arrest" & CRLF, usernum
    Exit Sub
    End If
user(un).rank = user(un).unarrestLevel
user(un).unarrestLevel = 0
user(un).arrested = False
user(un).room = rooms(1).name
send "~FM" & user(un).name & " has been freed from arrest" & CRLF, usernum
If un > 0 Then
    writeRoomExcept "", "~FM~OL" & user(un).name & " has been freed from arrest" & CRLF, user(usernum).name
    look un
    End If
End Sub

Sub invite(usernum As Integer, inpstr As String)
Dim un As Integer
inpstr = completeUsername(inpstr)
un = getUser(inpstr)
If un = 0 Then
    send MSG_USER_NOT_ONLINE & CRLF, usernum
    Exit Sub
    End If
If UCase$(user(usernum).name) = UCase$(inpstr) Then
    send "You cannot invite yourself" & CRLF, usernum
    Exit Sub
    End If
If Not rooms(getRoom(user(usernum).room)).access = ROOM_PRIVATE Then
    send "You cannot invite a user into a non-private room" & CRLF, usernum
    Exit Sub
    End If
If InStr(user(un).invitations, user(usernum).room) Then
    send user(un).name & " is already invited" & CRLF, usernum
    Exit Sub
    End If
If user(un).room = user(usernum).room Then
    send user(un).name & " is already in the room" & CRLF, usernum
    Exit Sub
    End If
user(un).invitations = user(un).invitations & " " & UCase$(user(usernum).room)
send "You have been invited by " & user(usernum).name & " into the " & user(usernum).room & CRLF, un
send "You invite " & user(un).name & " into the room" & CRLF, usernum
End Sub

Sub uninvite(usernum As Integer, inpstr As String)
inpstr = completeUsername(inpstr)
word(1) = completeUsername(word(1))
Dim un As Integer
un = getUser(inpstr)
If un = 0 Then
    send MSG_USER_NOT_ONLINE & CRLF, usernum
    Exit Sub
    End If
If UCase$(user(usernum).name) = UCase$(inpstr) Then
    send "You cannot uninvite yourself" & CRLF, usernum
    Exit Sub
    End If
If Not isUserInvited(un, user(usernum).room) Then
    send user(un).name & " is not invited" & CRLF, usernum
    Exit Sub
    End If
user(un).invitations = Replace(user(un).invitations, " " & UCase$(user(usernum).room), vbNullString)
If usernum < 1 Then
    Exit Sub
    End If
send "Your invitation to the " & user(usernum).room & " has been revoked" & CRLF, un
send "You uninvite " & user(un).name & " from the room" & CRLF, usernum
End Sub

Sub picture(usernum As Integer, inpstr As String, Optional showToRoom As Boolean = True)
Dim piclist() As String, file As String, msg As String
Dim free As Integer, FromFile As String, count As Integer
Dim count2 As Integer, everyOther As Boolean
'First a little security check
If containsCorrupt(inpstr) Then
    send "Illegal picture name" & CRLF, usernum
    Exit Sub
    End If
file = App.Path & "\Pictures\" & inpstr & ".TXT"
'Display the pic if an argument is given
If fileExists(file) Then
    free = FreeFile
    If showToRoom Then
        msg = "~BB~FT " & user(usernum).name & " shows a picture to the room ~RS" & CRLF & CRLF
        End If
    Open file For Input As #free
    Do While Not EOF(free)
        Line Input #free, FromFile
        msg = msg & FromFile & CRLF
        If Len(msg) > SEND_CHOP Then
            If showToRoom Then
                writeRoom user(usernum).room, msg
                Else
                    send msg, usernum
                    End If
            msg = vbNullString
            End If
        Loop
    Close #free
    If Len(msg) > 0 Then
        If showToRoom Then
            writeRoom user(usernum).room, msg
            Else
                send msg, usernum
                End If
        End If
    Exit Sub
    End If
'If no argument is given, show a list of all available pics
file = Dir$(App.Path & "\Pictures\*.txt")
If file = vbNullString Then
    send "There are no pictures on this server" & CRLF, usernum
    Exit Sub
    End If
Do
    If Not count = 0 Then
        file = Dir$
        End If
    If Len(file) > 4 Then
        ReDim Preserve piclist(count)
        piclist(count) = Left$(file, Len(file) - 4)
        count = count + 1
        End If
    Loop Until file = ""
ShellSort piclist
If count = 0 Then
    send "No pictures were found" & CRLF, usernum
    Exit Sub
    End If
msg = embedBar("Available Pictures") & "~FT"
For count = LBound(piclist) To UBound(piclist)
    msg = msg & mold(piclist(count), 14)
    count2 = count2 + 1
    If count2 = 5 Then
        If everyOther Then
            msg = msg & CRLF & "~FT"
            Else
                msg = msg & CRLF & "~RS"
                End If
        everyOther = Not everyOther
        count2 = 0
        End If
    If Len(msg) > SEND_CHOP Then
        send msg, usernum
        msg = ""
        End If
    Next count
msg = msg & CRLF
If count > 0 Then
    msg = msg & CRLF
    End If
send msg, usernum
End Sub

Sub syslogView(usernum As Integer)
Dim msg As String, found As Boolean, count As Integer
For count = logpos To UBound(logbook)
    If Not logbook(count) = "" Then
        msg = msg & logbook(count) & CRLF
        End If
    If Len(msg) > SEND_CHOP Then
        send msg, usernum
        msg = ""
        found = True
        End If
    Next count
For count = LBound(logbook) To logpos
    If count < logpos And Not logbook(count) = "" Then
        msg = msg & logbook(count) & CRLF
        End If
    If Len(msg) > SEND_CHOP Then
        send msg, usernum
        msg = ""
        found = True
        End If
    Next count
If Len(msg) < 3 And Not found Then
    msg = "The system logbook is empty."
    End If
If Len(msg) > 0 Then
    send msg, usernum
    End If
End Sub

Sub sayto(usernum As Integer, inpstr As String)
Dim un As Integer
word(1) = completeUsername(word(1))
inpstr = stripOne(inpstr)
If Not userIsOnline(word(1)) Then
    send MSG_USER_NOT_ONLINE & CRLF, usernum
    Exit Sub
    End If
un = getUser(word(1))
If Not user(usernum).room = user(un).room Then
    send user(un).name & " is not in this room" & CRLF, usernum
    Exit Sub
    End If
writeRoomExcept user(usernum).room, "~FT" & user(usernum).name & " says (to " & word(1) & "): ~RS" & inpstr & CRLF, user(usernum).name
send "~FTYou say to " & user(un).name & ": ~RS" & inpstr & CRLF, usernum
End Sub

Sub knock(usernum As Integer, inpstr As String)
Dim rn As Integer
rn = getRoom(word(1))
If wordCount(inpstr) < 1 Then
    send "Usage: knock <room>" & CRLF, usernum
    Exit Sub
    End If
If rn = 0 Then
    send MSG_ROOM_NOT_EXIST & CRLF, usernum
    Exit Sub
    End If
If user(usernum).room = rooms(rn).name Then
    send "You are already in the " & rooms(rn).name & CRLF, usernum
    Exit Sub
    End If
If isUserInvited(usernum, word(1)) Then
    send "You are already invited into this room" & CRLF, usernum
    Exit Sub
    End If
If Not rooms(rn).access = ROOM_PRIVATE Then
    send "The " & rooms(rn).name & " is not currently private" & CRLF, usernum
    Exit Sub
    End If
If user(usernum).rank < staffLevel And InStr(UCase$(rooms(getRoom(user(usernum).room)).allExits), UCase$(word(1))) = 0 Then
    send "The " & rooms(rn).name & " is not adjacent to this room" & CRLF, usernum
    Exit Sub
    End If
writeRoom rooms(rn).name, "~FM" & user(usernum).name & " knocks, asking to be let into the room" & CRLF
send "You ask to be let into the " & rooms(rn).name & CRLF, usernum
End Sub

Sub pemote(usernum As Integer, inpstr As String)
Dim un As Integer
word(1) = completeUsername(word(1))
If wordCount(inpstr) < 2 Then
    send "Usage: pemote <user> <message>" & CRLF, usernum
    Exit Sub
    End If
If Not userIsOnline(word(1)) Then
    send MSG_USER_NOT_ONLINE & CRLF, usernum
    Exit Sub
    End If
un = getUser(word(1))
inpstr = stripOne(inpstr)
send "~FG(=>" & user(un).name & ")~RS " & user(usernum).name & " " & inpstr & CRLF, usernum
send "~FG(Private)~RS " & user(usernum).name & " " & inpstr & CRLF, un
End Sub

Sub mutter(usernum As Integer, inpstr As String)
Dim un As Integer
word(1) = completeUsername(word(1))
If wordCount(inpstr) < 2 Then
    send "Usage: mutter <user> <message>" & CRLF, usernum
    Exit Sub
    End If
If Not userIsOnline(word(1)) Then
    send MSG_USER_NOT_ONLINE & CRLF, usernum
    Exit Sub
    End If
un = getUser(word(1))
If un = usernum Then
    send "Talking about yourself behind your own back is rather silly, is it not?" & CRLF, usernum
    Exit Sub
    End If
If Not user(un).room = user(usernum).room Then
    send user(un).name & " is not in this room" & CRLF, usernum
    Exit Sub
    End If
writeRoomExcept user(un).room, "~FT" & user(usernum).name & " mutters: ~RS" & stripOne(inpstr) & " ~FY(To all but ~OL" & user(un).name & "~RS~FY)~RS" & CRLF, user(un).name
End Sub

Sub people(usernum As Integer)
Dim msg As String, count As Integer, everyOther As Boolean
msg = embedBar("People", "~FY", "~FB", "~RS") & "~FY~OLName            Level Line Ignall Vis Idle Mins Site/Service" & CRLF
For count = 1 To UBound(user)
    If user(count).operational Then
        If user(count).state > STATE_LOGIN3 Then
            If everyOther Then
                msg = msg & "~FW"
                Else
                    msg = msg & "~FT"
                    End If
            everyOther = Not everyOther
            msg = msg & mold(user(count).name, 15) & " " & _
                mold(ranks(user(count).rank), 5) & " "
            If Not user(count).netlinkType Then
                msg = msg & mold(Trim$(user(count).line), 4)
                Else
                    msg = msg & "N/A "
                    End If
            msg = msg & " " & mold(bool2YN(user(count).listening), 6) & _
                " " & mold(bool2YN(Not user(count).visible), 3) & " " & _
                mold(Trim$(user(count).idle), 4) & " " & _
                mold(Trim$(user(count).timeon / 60), 4) & " "
            If user(count).netlinkType Then
                msg = msg & net(s2n(user(count).netlinkFrom)).name
                Else
                    msg = msg & user(count).site
                    End If
            msg = msg & CRLF
            Else
                msg = msg & "~FY~OL[Login Stage " & Trim$(user(count).state + 1) & "] ----- " & user(count).line & "    ------ --- ---- " & mold(Trim$(user(count).timeon), 4) & " " & user(count).site & CRLF
                End If
        End If
    If Len(msg) > SEND_CHOP Then
        send msg, usernum
        msg = vbNullString
        End If
    Next count
msg = msg & embedBar(, "~FY", "~FB", "~RS")
send msg, usernum
End Sub
