Attribute VB_Name = "html"
Global http_connections As Integer
Global http_requests As Integer
Global http_sites(30) As String
Global http_sitepos_pointer As Integer

Function html_error() As String
html_error = "HTTP/1.1" & CRLF
html_error = html_errer & "<HTML><HEAD><TITLE> " & system.talkerName & " </TITLE>" & _
       "</HEAD><BODY><H1> Error </H1><P>" & CRLF
End Function

Function html_who() As String
Dim numfound As Integer
html_who = "<HTML><HEAD><TITLE> " & system.talkerName & " Who Port </TITLE>" & _
       "</HEAD><BODY><H1> " & system.talkerName & " Who Port" & _
       "</H1><P>" & CRLF
For count = 1 To UBound(user)
    If user(count).state > STATE_LOGIN3 And user(count).operational And user(count).visible Then
        numfound = numfound + 1
        html_who = html_who & "<A HREF=" & doubleQuote
        html_who = html_who & "/&" & user(count).name & doubleQuote & ">" & user(count).name & "</A>" & "<BR>" & CRLF
        End If
    Next count
If numfound = 0 Then
    html_who = html_who & "No users online at the moment<BR>"
    End If
html_who = html_who & "<BR><A HREF=" & doubleQuote & "/" & doubleQuote & ">Back</A>" & "</P></BODY></HTML>"
End Function

Function html_ex_user(UserName As String) As String
If Len(word(1)) > MAX_NAME_LEN Then
    html_ex_user = html_error
    Exit Function
    End If
If Not userExists(UserName) Then
    html_ex_user = "<HTML><HEAD><TITLE> " & system.talkerName & " </TITLE>" & _
                   "</HEAD><BODY><H2>User not found</H2></BODY></HTML>"
    Exit Function
    End If

Dim h As String
h = "<HTML><HEAD><TITLE> " & system.talkerName & " </TITLE></HEAD>" & _
    "<BODY><H2>" & userCap(word(1)) & "'s User Data" & "</H3>"
If Not userIsOnline(word(1)) Then
    user(0).name = word(1)
    loadUserData 0
    datehold = num2date(user(0).lastLogin)
    h = h & "<STRONG>Desc:</STRONG> " & user(0).desc & "<BR>" & _
           "<STRONG>Rank: </STRONG>" & ranks(user(0).rank) & "<BR>" & _
           "<STRONG>Total Login: </STRONG>" & deriveTimeString(spliceTime(user(0).totalTime), False) & CRLF & "<BR>" & _
           "<STRONG>Last Login: </STRONG>" & Format$(datehold, "dddd d") & getOrdinal(Int(Format$(datehold, "d"))) & Format$(datehold, " mmmm yyyy" & " at " & Format$(datehold, "hh:nn")) & "<BR>" & _
           "<STRONG>Which Was: </STRONG>" & deriveTimeString(spliceTime(date2num(Now) - user(0).lastLogin)) & "<BR>" & _
           "<STRONG>Was On For: </STRONG>" & deriveTimeString(spliceTime(user(0).timeon), False) & "<BR>" & _
           "<STRONG>Enter Message: </STRONG>" & user(0).enterMsg & "<BR>" & _
           "<STRONG>Exit Message: </STRONG>" & user(0).exitMsg & "<BR>" & _
           "<STRONG>Last Site: </STRONG>" & user(0).site & "<BR>" & _
           "<STRONG>Email: </STRONG>" & IIf(user(0).email = "Unset", "Unset", "<a href=mailto:" & user(0).email & ">" & user(0).email & "</a>") & "<BR>" & _
           "<STRONG>ICQ: </STRONG>" & user(0).icq & "<BR>" & _
           "<STRONG>Logins: </STRONG>" & user(0).logins & "<BR>" & _
           "<STRONG>Gender: </STRONG>" & user(0).gender & "<BR>" & _
           "<STRONG>Age: </STRONG>" & IIf(Val(user(0).age) > 0, user(0).age, "Unset") & "<BR>" & _
           "<STRONG>Arrested: </STRONG>" & bool2YN(user(0).arrested) & "<BR>" & _
           "<STRONG>Visible: </STRONG>" & bool2YN(user(0).visible) & "<BR>" & _
           "<STRONG>Muzzled: </STRONG>" & bool2YN(user(0).muzzled) & "<BR>" & _
           "<STRONG>Expires: </STRONG>" & bool2YN(user(0).expires) & "<BR>" & _
           "<STRONG>Email Fwd: </STRONG>" & bool2YN(user(0).sfRec) & "<BR>"
        Else
            un = getUser(word(0))
            h = h & "<STRONG>Desc:</STRONG> " & user(un).desc & "<BR>" & _
                   "<STRONG>Rank: </STRONG>" & ranks(user(un).rank) & "<BR>" & _
                   "<STRONG>Total Login: </STRONG>" & deriveTimeString(spliceTime(user(un).totalTime), False) & "<BR>" & _
                   "<STRONG>On For: </STRONG>" & deriveTimeString(spliceTime(user(un).timeon), False) & "<BR>" & _
                   "<STRONG>Idle For: </STRONG>" & deriveTimeString(spliceTime(user(un).idle * 60), False) & "<BR>" & _
                   "<STRONG>Room: </STRONG>" & user(un).room & "<BR>"
            If user(un).atNetlink >= 0 Then
                h = h & "<STRONG>At Netlink: </STRONG>" & net(user(un).atNetlink).name & "<BR>"
                End If
            If user(un).netlinkType Then
                h = h & "<STRONG>From Netlink: </STRONG>" & net(user(un).netlinkFrom).name & "<BR>"
                End If
            h = h & "<STRONG>Enter Message: </STRONG>" & user(un).enterMsg & "<BR>" & _
                   "<STRONG>Exit Message: </STRONG>" & user(un).exitMsg & "<BR>" & _
                   "<STRONG>Site: </STRONG>" & user(un).site & "<BR>" & _
                   "<STRONG>Email: </STRONG>" & IIf(user(un).email = "Unset", "Unset", "<a href=mailto:" & user(un).email & ">" & user(un).email & "</a>") & "<BR>" & _
                   "<STRONG>ICQ: </STRONG>" & user(un).icq & "<BR>" & _
                   "<STRONG>Logins: </STRONG>" & user(un).logins & "<BR>" & _
                   "<STRONG>Gender: </STRONG>" & user(un).gender & "<BR>" & _
                   "<STRONG>Age: </STRONG>" & IIf(Int(user(0).age) > 0, user(0).age, "Unset") & "<BR>" & _
                   "<STRONG>Arrested: </STRONG>" & bool2YN(user(un).arrested) & "<BR>" & _
                   "<STRONG>Visible: </STRONG>" & bool2YN(user(un).visible) & "<BR>" & _
                   "<STRONG>Muzzled: </STRONG>" & bool2YN(user(un).muzzled) & "<BR>" & _
                   "<STRONG>Expires: </STRONG>" & bool2YN(user(un).expires) & "<BR>" & _
                   "<STRONG>Email Fwd: </STRONG>" & bool2YN(user(un).sfRec)
                End If
html_ex_user = cStrip(h)
End Function

Function html_index() As String
html_index = "<HTML><HEAD><TITLE> " & system.talkerName & " </TITLE></HEAD>" & _
    CRLF & "<BODY><H3>" & system.talkerName & " Web Interface</H3><P>" & system.talkerName & _
    " provides the following services: <BR>" & "<A HREF=" & doubleQuote & _
    "/who" & doubleQuote & ">Who</A><BR><A HREF=" & doubleQuote & _
    "/examine" & doubleQuote & ">Examine</A><BR>" & _
    "<A HREF=" & doubleQuote & "telnet://" & mainForm.Socket1.LocalName & _
    ":" & mainForm.Socket1.LocalPort & doubleQuote & ">Enter " & system.talkerName & "</A></P></BODY></HTML>"
End Function

Function html_examine() As String
Dim h As String, count As Integer, userhold() As String, uc As Integer
h = "<HTML><HEAD><TITLE> List of Users </TITLE></HEAD><BODY><P>This is a list of all user " & _
    " accounts that exist on " & system.talkerName & ". You will need a table-capable browser " & _
    "to view this page. If your browser does not have the ability to display tables then you should " & _
    "get one. Microsoft's Internet Explorer and Netscape's Navigator are both good browsers.<BR><BR>" & _
    "This page was designed for <I>Microsoft Internet Explorer 4.0</I><BR><BR>" & _
    "<TABLE BORDER=" & doubleQuote & "0" & doubleQuote & "Width=" & doubleQuote & "100%" & doubleQuote & ">"
userfile = Dir$(App.Path & "\Users\*.D")
count = -1
Do While Not userfile = ""
    userfile = Dir$
    count = count + 1
    Loop
If count < 0 Then
    h = "<HTML><HEAD><TITLE> List of Users </TITLE></HEAD><BODY><P>No userfiles were found on this system<BR><A HREF=" & doubleQuote & "/" & doubleQuote & ">Back</A>" & "</P></BODY></HTML>"
    html_examine = h
    Exit Function
    End If
ReDim userhold(count)
count = 0
userfile = Dir$(App.Path & "\Users\*.D")
Do While Not userfile = ""
    If Len(userfile) > 2 Then
        userfile = userCap(Left$(userfile, Len(userfile) - 2))
        userhold(count) = userfile
        userfile = Dir$
        count = count + 1
        End If
    Loop
'The shell sort algorythm
Dim swapped As Boolean, gap As Integer, temp As String
gap = Int(UBound(userhold) / 2)
Do
    Do
        swapped = False
        For count = LBound(userhold) To UBound(userhold) - gap
            If userhold(count) > userhold(count + gap) Then
                temp = userhold(count)
                userhold(count) = userhold(count + gap)
                userhold(count + gap) = temp
                swapped = True
                End If
            Next count
        Loop While swapped
    gap = Int(gap / 2)
    Loop While gap > 0
    
count = 0
For uc = LBound(userhold) To UBound(userhold)
    If count = 0 Then
        h = h & "<TR>"
        End If
    If Len(userhold(uc)) > 2 Then
        userfile = Left$(userhold(uc), Len(userhold(uc)) - 2)
        h = h & CRLF & "<TD WIDTH=" & doubleQuote & "20%" & doubleQuote & "><A HREF=" & doubleQuote & "/&" & userhold(uc) & doubleQuote & ">" & userhold(uc) & "</A>" & "</TD>"
        End If
    count = count + 1
    If count >= 5 Then
        count = 0
        h = h & "</TR>"
        End If
    Next uc
h = h & "</TABLE><BR><A HREF=" & doubleQuote & "/" & doubleQuote & ">Back</A>" & "</P></BODY></HTML>"
html_examine = h
End Function

