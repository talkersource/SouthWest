Attribute VB_Name = "SMTP"
Option Explicit
'Designed For ESMTP Servers
'======== === ===== =======
' This should work with most others but I'll admit that it
' isnt the strictest implement of the RFC.

'SMTP Response Codes
'220 Service is ready
'250 Can be ignored / Message
'251 Will forward
'354 Client should complete command
'550 Mailbox doesnt exist
'551 Wont forward

Type ML_OBJECT
    address As String   'Email address
    inuse As Boolean    'Process during mail run
    message As String   'Message body
    Server As String    'Server to connect to
    timestamp As Date   'Time message was typed
    u_email As String   'To set as reply address
    u_from As String    'User that typed message
    u_to As String      'Intended receiver
    userid As String    'Email username
    success As Boolean  'Success of operation
    End Type

'Error messages returned by the setup_address function
Public Const NO_USER = -1
Public Const BAD_EMAIL_NAME = -2

'Mail States
Global SMTP_STATE As Integer
Public Const EMAIL_NOT_CONNECTED = 0
Public Const EMAIL_CONNECTING = 1
Public Const EMAIL_HELO = 2
Public Const EMAIL_WAIT_GREET = 3
Public Const EMAIL_WAIT_MYID_ACK = 4
Public Const EMAIL_WAIT_SET_RECEIVER_ACK = 5
Public Const EMAIL_WAIT_EDIT = 6
Public Const EMAIL_EDITING = 7
Public Const EMAIL_END_EDIT = 8
Public Const EMAIL_QUIT = 9

'Normally you would just love to be able to put in a big
'number that wouldnt be overflowed. Thing is there are
'naughty people out there that might abuse the buffer.
Public Const MAX_MAIL_SLOTS = 7
Global mail(MAX_MAIL_SLOTS) As ML_OBJECT
Global mail_out  As ML_OBJECT

Function Setup_Address(UserName As String) As Integer
Dim cur_block As Integer, count As Integer, email As String
If Not userExists(UserName) Then
    Setup_Address = -1
    Exit Function
    End If
cur_block = -1
For count = 0 To MAX_MAIL_SLOTS
    If Not mail(count).inuse Then
        mail(count).inuse = True
        cur_block = count
        Exit For
        End If
    Next count
If cur_block = -1 Then
    Setup_Address = -2
    Exit Function
    End If
If userIsOnline(UserName) Then
    email = user(getUser(UserName)).email
    Else
        user(0).name = UserName
        loadUserData (0)
        email = user(0).email
        End If
If Not InStr(email, "@") > 1 And InStr(email, "@") < Len(email) Then
    Setup_Address = -3
    Exit Function
    End If
mail(cur_block).u_email = email
mail(cur_block).Server = Right$(email, Len(email) - InStr(email, "@"))
mail(cur_block).userid = Left$(email, InStr(email, "@") - 1)
Setup_Address = cur_block
End Function
 
Function gen_ver_code() As String
'This function will generate the verification code that
'will be sent to the user via email to verify that the
'account is active and owned/used by the user.
Dim k As Integer, count As Integer
Randomize Timer
For count = 0 To 5
    gen_ver_code = gen_ver_code & Chr$(((Rnd * 9) + 48))
    Next count
End Function

Sub smtp_out()
mainForm.mailsock.AddressFamily = AF_INET
mainForm.mailsock.Protocol = IPPROTO_TCP
mainForm.mailsock.SocketType = SOCK_STREAM
mainForm.mailsock.Binary = True
mainForm.mailsock.Blocking = False
mainForm.mailsock.bufferSize = 4000
mainForm.mailsock.HostName = system.smtpServer
mainForm.mailsock.RemotePort = 25
mainForm.mailsock.Connect
End Sub

Function Get_Code(msg As String) As Integer
If Len(msg) >= 3 Then
    Get_Code = Val(Left$(msg, 3))
    End If
End Function

Function drop_smtp()
SMTP_STATE = EMAIL_NOT_CONNECTED
If mainForm.mailsock.Connected Then
    mainForm.mailsock.Disconnect
    End If
If mail_out.success Then
    writeSyslog "~A mail message sent ~FGsuccessfully"
    Else
        writeSyslog "~FRA mail message was unsuccessfully sent"
        End If
End Function

Sub smtp_send(message As String)
If mainForm.mailsock.Connected Then
    mainForm.mailsock.SendLen = Len(message)
    mainForm.mailsock.SendData = message
    End If
End Sub

Sub loadSpool()
Dim mailfile As String, count As Integer, FromFile As Variant
Dim mail_num As Integer, count2 As Integer
lighter "Loading spooled mail"
mailfile = Dir$(App.Path & "\Spool\*.S")
If mailfile = "" Then
    Exit Sub
    End If
For count = 0 To MAX_MAIL_SLOTS
    If mailfile = "" Then
        Exit For
        End If
    If Len(mailfile) > 2 Then
        writeSyslog "Loading " & Left$(mailfile, Len(mailfile) - 2) & " from spool"
        Else
            writeSyslog "~FRSpooling canceled because of bad filename in mail spool"
            Exit Sub
            End If
    mail_num = -1
    For count2 = 0 To MAX_MAIL_SLOTS
        If Not mail(count2).inuse Then
            mail_num = count2
            Exit For
            End If
        Next count2
    If mail_num = -1 Then
        Exit For
        End If
    Open App.Path & "\Spool\" & mailfile For Input As #1
    mail(mail_num).inuse = True
    Line Input #1, mail(mail_num).u_email
    Line Input #1, mail(mail_num).u_to
    Line Input #1, mail(mail_num).u_from
    Line Input #1, mail(mail_num).userid
    Line Input #1, FromFile
    mail(mail_num).timestamp = FromFile
    If EOF(1) Then
        mail(mail_num).inuse = False
        End If
    Do While Not EOF(1)
        Line Input #1, FromFile
        mail(mail_num).message = mail(mail_num).message & FromFile & CRLF
        Loop
    Close #1
    Kill App.Path & "\Spool\" & mailfile
    mailfile = Dir$()
    Next count
End Sub

Sub save_spool()
'Saves the mail to a spooler file under a random name
Dim count As Integer, filename As String
For count = 0 To MAX_MAIL_SLOTS
    If mail(count).inuse Then
        Do
            filename = gen_ver_code
            Loop While Not Dir$(App.Path & "\Spool\" & filename & ".S") = ""
        writeSyslog "Spooling " & mail(count).userid & " as " & filename
        Open App.Path & "\Spool\" & filename & ".S" For Output As #1
        Print #1, mail(count).u_email & CRLF;
        Print #1, mail(count).u_to & CRLF;
        Print #1, mail(count).u_from & CRLF;
        Print #1, mail(count).userid & CRLF;
        Print #1, mail(count).timestamp & CRLF;
        Print #1, mail(count).message;
        Close #1
        End If
    Next count
End Sub
