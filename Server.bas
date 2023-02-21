Attribute VB_Name = "Server"
Option Explicit
Option Compare Binary
Option Base 0
'              ____
'              |  | Important : You must increment this value for
'              |  | =========   every new command you make since
'             _|  |_            this value is used to determine
'             \    /            the size of the array that holds
'              \  /             all of the commands.
'               \/
Public Const NUM_OF_COMMANDS = 110




'These two pre-compiler directives are not used by version one of
'SouthWest but have been added for future support. If you have
'converted SouthWest to run in VB5, you should change the VB_VERSION
'constant to 5.
#Const VB_VERSION = 6
#Const SW_VERSION = 1

'********API Declarations, Constants, and Structures**********
'This is system stuff that I didn't have too much to do with. The
'following block of code is uncommented and cramped because there
'is nothing that you would really get out of this anyways.
Declare Function xcrypt Lib "SWcrypt.dll" (ByRef plain As String) As String
Declare Function ICQSetKey Lib "icqmapi.dll" Alias "ICQAPICall_SetLicenseKey" (ByVal pszName As String, ByVal pszPassword As String, ByVal pszLicense As String) As Long
Declare Function ICQSendMessage Lib "icqmapi.dll" Alias "ICQAPICall_SendMessage" (ByVal iUIN As Long, ByVal pszMessage As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal h%, ByVal hb%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Public Const WM_LBUTTONDBLCLICK = &H203, WM_LBUTTONUP = &H202
Public Const WM_RBUTTONUP = &H205, WM_MOUSEMOVE = &H200
Public Const NIM_ADD = &H0, NIM_DELETE = &H2, NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2, NIF_TIP = &H4, MF_BYPOSITION = &H400&
Public Const MF_REMOVE = &H1000&, REVBUFF_SIZE = 15, MIIM_TYPE = &H10
Public Const MFT_RADIOCHECK = &H200&

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uId As Long
  uFlags As Long
  ucallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type
  
'This is the end of the API declarations, constants, and structures.
'What remains is the SouthWest source code. This module has been
'broken up into components like enums, structures, constants, globals,
'and functions. The code has been heavily commented to help you.

'************* Enumerated Constants***************
' These settings will tell us how to filter words in the swear list.
Enum swbanLevels
    SWEAR_OFF       ' No filtering (fastest)
    SWEAR_MIN       ' Censor swear words
    SWEAR_MAX       ' Reject lines with swearing (idiotic)
    End Enum
    
' Tells us what to AutoGrapher should be displaying
Enum graph
    GRAPH_NULL              ' Graph nothing
    GRAPH_ACTIONS           ' Graph user actions
    GRAPH_USER_LOGINS       ' Graph user logins
    GRAPH_HTTP_CONNECTIONS  ' Graph different HTTP servings
    GRAPH_HTTP_REQUESTS     ' Graph all HTTP servings
    End Enum

' Very important. This tells what the user is doing.
Enum userStates
    STATE_LOGIN1            ' Enter your name
    STATE_LOGIN2            ' Enter your password
    STATE_LOGIN3            ' Retype password (only for new users)
    STATE_NORMAL            ' Normal
    STATE_EDITOR            ' Line editor
    STATE_EDPICK            ' Save, View, Redo, Abord
    STATE_OPTION            ' Other prompts
    End Enum
    
' This will tell us why the user is using the line editor.
Enum editorStates
    EDITSTATE_ENTPRO        ' Entering a profile
    EDITSTATE_SMAIL         ' Sending mail
    EDITSTATE_BOARD         ' Writing on the message board
    End Enum

' If a user is in the STATE_OPTION mode, this will help us narrow
' down just what it is that he is doing. Ussually this STATE_OPTION
' mode is used for answering a question or prompt.
Enum stateOptions
    OPTION_NONE             ' User is not in STATE_OPTION
    OPTION_SHUTDOWN         ' Deciding if he wants to shut it down
    OPTION_REBOOT           ' Deciding if he wants to reboot
    OPTION_SUICIDE          ' Contemplating suicide
    End Enum
    
' The shutdown types for the talker
Enum shutdownTypes
    SHUTDOWN_NONE               ' Server is operating normally
    SHUTDOWN_USERCHOOSING_SHUT  ' User deciding on shutdown
    SHUTDOWN_USERCHOOSING_REBT  ' User deciding on reboot
    SHUTDOWN_SHUTDOWN           ' Shutdown (or countdown)
    SHUTDOWN_REBOOT             ' Reboot (or countdown)
    End Enum

' Determines who can enter what rooms
Enum roomAccesses
    ROOM_PUBLIC             ' Everyone may enter
    ROOM_PRIVATE            ' Only those already there or invited
    ROOM_STAFF              ' Staff only
    End Enum

' For calls to functions about banishment
Enum banTypes
    BAN_USER                ' User banned
    BAN_SITE                ' Site/domain banned
    BAN_NEW                 ' No more new users from site
    End Enum

'**************** User-Defined Data Types ****************

Public Type CM_OBJECT           ' **commmand object**
    name As String              ' command name
    rank As Integer             ' minimum rank
    End Type

Public Type RM_OBJECT           ' **room object**
    access As roomAccesses      ' the room's access
    allExits As String          ' a list of all the room's exits.
                                ' uses commas as delimiters
    buffer(1 To REVBUFF_SIZE) As String ' room's review buffer
    exits(30) As String         ' an array containing the room's exits
    name As String              ' the room's name
    topic As String             ' the current topic
    locked As Boolean           ' is the access changable?
    End Type

Public Type SM_OBJECT           ' **smail object**
    receiver As String          ' who is receiving the message?
    message(1 To 15) As String  ' the actual message
    End Type

Public Type UR_OBJECT           ' **user object**
    afk As Boolean              ' is user away from keyboard?
    age As String               ' the user's age
    arrested As Boolean         ' is the user arrested?
    atNetlink As Integer        ' holds Netlink socket number if the
                                ' user is at a Netlink or -1 if not.
    charEchoing As Boolean      ' send keystrokes back to user?
    cloneCount As Integer       ' Number of clones the user has active
    desc As String              ' textual description
    Editor(1 To 15) As String   ' private line editor buffer
    editorType As editorStates  ' what is he editing?
    editorPos As Integer        ' what line is he editing
    email As String             ' the user's email address
    enterMsg As String          ' room entrance announcement
    exitMsg As String           ' room exit announcement
    expires As Boolean          ' does the user expire with purge
    gender As String            ' the user's gender
    ICQ As String               ' the user's ICQ number
    idle As Integer             ' how long he has been idle (minutes)
    inpstr As String            ' holds what the user has typed
    Index As Integer            ' ussually the socket to which the
                                ' user is connected. Mind you that
                                ' there are Netlink users as well.
    invitations As String       ' rooms to which he is invited
    lastLogin As Double         ' last date user was here (UNIX)
    site As String              ' the user's login site
    line As Integer             ' much like index
    listening As Boolean        ' is the user listening?
    listing As Integer          ' user's position in the connections
                                ' list box
    logins As Integer           ' number of times that the user has
                                ' logged into the server
    muzzled As Boolean          ' can he speak?
    name As String              ' the user's name
    netlinkFrom As Integer      ' user's Netlink socket
    netlinkType As Boolean      ' is he from a Netlink?
    netlinkPending As Boolean   ' is he trying to enter a Netlink?
    newUser As Boolean          ' is the user a new user?
    oldInpstr As String         ' last command issued
    oldRoom As String           ' room the user was in before this one
    operational As Boolean      ' are we using this user?
    options As stateOptions     ' are we waiting for the user to answer
                                ' a question or prompt?
    outBuffer As String         ' buffer for sending user text
    outMail As SM_OBJECT        ' holds smail that he's sending
    pager As Integer            ' height of user's screen
    password As String          ' user's password (encrypted F-Crypt)
    profile(1 To 15) As String  ' their profile
    rank As Integer             ' the user's rank
    room As String              ' current room location
    sfVercode As Long           ' smail-to-email verification code
    sfRec As Boolean            ' should we forward their mail?
    sfVerifyed As Boolean       ' have they verified?
    state As userStates         ' the user's state (very important)
    timeon As Double            ' time the user has spent on here
    totalTime As Double         ' total login time
    unarrestLevel As Integer    ' level at which he should be returned
                                ' after being unarrested
    unread As Boolean           ' does he have unread mail?
    url As String               ' his homepage
    visible As Boolean          ' is he visible?
    End Type

Public Type SYS_OBJECT          ' **system object**
    emailAddress As String      ' system's email addy
    shutdownCount As Integer    ' shutdown/reboot countdown
    shutdownType As shutdownTypes   ' shutdown type
    rebooted As Boolean         ' did a reboot load this session?
    icqHook As Boolean          ' did we hooked into ICQ successfully?
    mainPort As Integer         ' talker's main port
    netlinkPort As Integer      ' Netlink port
    httpPort As Integer         ' HTTP Server port
    maxIdle As Integer          ' maximum user idle (minutes)
    swearing As swbanLevels     ' swearing filter setting
    autoConnect As Boolean      ' automaticly connect set Netlinks
    gatecrash As Boolean        ' allow gatecrashing?
    figlet As String            ' current figlet font
    gatecrashLevel As Integer   ' gatecrash minlevel
    talkerName As String        ' the name of the talker
    smtpServer As String        ' SMTP server for sending email
    purgeLength As Integer      ' accounts expire in this many days
    timeoutMaxLevel As Integer  ' maximum level at which users timeout
    siteAtStartup As Boolean    ' show site at startup
    End Type

Public Type DHMS_OBJECT         ' **day/hour/min/sec time object**
    days As Long                ' days part
    hours As Byte               ' hours part
    minutes As Byte             ' minutes part
    seconds As Byte             ' seconds part
    End Type

Public Type CL_OBJECT           ' **clone object**
    active As Boolean           ' is it active?
    owner As Integer            ' who is the owner/controller
    room As Integer             ' room in which the clone is talking
    End Type
    
Public Type PURGE_OBJECT        ' **purge object**
    usersRemoved As Integer     ' number of users removed
    usersNow As Integer         ' number of users now on the system
    End Type

Public Type VI_OBJECT           ' **viewer object**
    vDate As String             ' date or null for tooltip
    vRest As String             ' the nondate remainder
    End Type

Public Type LOAD_OBJECT         ' **load object**
    specifier As String         ' the L-value (setting name)
    value As String             ' the value
    End Type

Public Type PLUGIN_OBJECT       ' **plugin object**
    name As String              ' the plugin's name
    version As String           ' version of the plugin
    hooks As String             ' hooks
    exeptr As String            ' executable pointer
    inuse As Boolean            ' is it in use
    menuPos As Integer          ' it's position in the menu
    End Type

'******************* Constants *******************

'For always on top
Public Const SWP_NOMOVE = 2, SWP_NOSIZE = 1
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1, HWND_NOTOPMOST = -2

'Server Maximums
Public Const MAX_DATA_LEN = 1024    ' max send length (bytes)
Public Const MAX_PLUGINS = 20       ' max number of plugins
Public Const MAX_ARGS = 20          ' max argument words
Public Const MAX_SYSLOG = 199       ' syslog size (lines)
Public Const MAX_USER_CLONES = 3    ' maximum clones per user
Public Const MAX_NAME_LEN = 11      ' max user name length (chars)

'Messages
Public Const MSG_USER_NOT_EXIST = "There is no such beast"
Public Const MSG_ROOM_NOT_EXIST = "There is no such room"
Public Const MSG_TOPIC_NOT_SET = "Topic has not been set"
Public Const MSG_USER_MUZZLED = "You are muzzled and can not do this"
Public Const MSG_NO_NETLINK = "This command is unavailable to remote users"
Public Const MSG_USER_GOES_VIS = "steps out from the shadows"
Public Const MSG_USER_GOES_INVIS = "steps into the shadows"
Public Const MSG_USER_NOT_ONLINE = "User is not online at the moment"

'Other Constants
Public Const CRLF = vbCrLf          ' Carriage return/Line feed
Public Const CR = vbCr              ' Carriage return
Public Const LF = vbLf              ' Line feed
Public Const PHYSICALMAXCONNECTIONS = 150   ' max sockets alloc'd
Public Const CLIENT_WIDTH = 73      ' default client width (chars)
'the default separater bar
Public Const FANCY_BAR = "~FR+~FG-----------------------------------------------------------------------~FR+~RS" & CRLF
'The higher the number, the greater the potential speed but this
'also means data loss for Netlinks and slow connections. Setting
'this number too low can cause major speed degradations in certain
'commands. Use this variable wisely. I have found 725 to be pretty
'much optimal. Never set this above 950 and never set this below 150
'or else the server may send data out of sequence or not send data
'properly. This is basicly the speed control for the talker.
Public Const SEND_CHOP = 725

'******************* Globals *******************
Public rooms() As RM_OBJECT         ' holds all room info
Public user() As UR_OBJECT          ' holds all user info
Public clones() As CL_OBJECT        ' holds all clone info
Public colorHex(20) As Long         ' color hex values
Public colorShorts(20) As String    ' NUTS markup codes
Public colorValues(20) As String    ' telnet values
Public cmds(NUM_OF_COMMANDS) As CM_OBJECT   ' all commands
Public system As SYS_OBJECT         ' system information
Public staffLevel As Integer        ' level at which a user is
                                    ' considered to be a staffie
Public figletWidth As Byte          ' width of the current figlet
Public figletHeight As Byte         ' height of the current figlet
Public maxUsers As Integer          ' max simult users
Public noSplash As Boolean          ' no splashscreen at startup
Public easterEgg As Boolean         ' for the hidden "easter egg"
Public grapher() As Integer         ' holds autograph points
Public madeSockets(PHYSICALMAXCONNECTIONS) As Boolean 'shows which
                                    ' sockets are loaded and which
                                    ' are not loaded.
Public BOOTING As Boolean           ' are we still booting?
Public whatToGraph As graph         ' what is in the grapher array?
Public BELL As String               ' chr$(7), the system bell
Public actup As Boolean             ' is the grapher visible?
Public doubleQuote As String        ' double quote chartactar
Public echoOff As String            ' telnet echo off command
Public echoOn As String             ' telnet echo on command
Public MOTD As String               ' the login screen
Public ranks() As String            ' rank names
Public word(MAX_ARGS) As String     ' used for everything (important)
Public actions As Integer           ' current minute actions
Public httpActions As Integer       ' current minute http acts
Public userLoginsGraph As Integer   ' current minute logins
Public actPos As Integer            ' position in actions array
Public acts(1 To 120) As Integer    ' holds past actions
Public logins_graph(1 To 120) As Integer ' holds logins
Public http_acts(1 To 120) As Integer    ' holds http draws
Public httpConnects(1 To 120) As Integer ' holds http connects
Public figlets() As String          ' holds active figlet font
Public figbar As String             ' figlet charactars
Public logbook(MAX_SYSLOG)          ' the system logbook
Public logpos As Integer            ' logbook position
Public roomAccessLevels(2)          ' room access level names
Public swearNames(2)                ' swear filter switch names
Public swears() As String           ' a list of words to be filtered
Public VBGTray As NOTIFYICONDATA    ' for systray
Public autoPatchFile As String      ' the autopatch file we're doing
Public plugs(MAX_PLUGINS) As PLUGIN_OBJECT ' holds plugin data
Public lastprog As Integer          ' for AP statbar

Function countMessages(UserName As String) As Integer
'Counts the number of messages on the message board

'Does the file exist?
If Dir$(App.Path & "\Users\" & UserName & ".M", 0) = "" Then
    countMessages = 0
    Exit Function
    End If
'Open the file
Open App.Path & "\Users\" & UserName & ".M" For Input As #1
Dim FromFile As String, found As Integer
'Loop until the end of the file is reached, counting the lines
Do While Not EOF(1)
    Line Input #1, FromFile
    If FromFile = "" Then
        found = found + 1
        End If
    Loop
Close #1
End Function
Sub Commands(inpstr As String, usernum As Integer)
Dim count, lenToSpace, cmdFound As Integer
'This is the big command parser, when a user inputs something
'that starts with a '.' it comes here. Lets go!

'First we strip the period out of there
If Len(inpstr) > 1 Then
    inpstr = Right(inpstr, (Len(inpstr) - 1))
    End If
spliceWords (inpstr)
'Look up the command
cmdFound = -1
For count = 0 To NUM_OF_COMMANDS
    If InStr(cmds(count).name, LCase$(word(0))) = 1 Then
        If user(usernum).rank >= cmds(count).rank Then
            cmdFound = count
            Exit For
            Else
                Exit For
                End If
        End If
    Next count
If cmdFound = -1 Then
    send "Invalid command" & CRLF, usernum
    Exit Sub
    End If
'Some commands that use prompts go very not well over a
'netlink so we will stop them right here
Dim c As Integer
If c = 8 Or c = 19 Or c = 23 Or c = 29 Or c = 31 And user(usernum).netlinkType Then
    send "This command cannot be executed over a netlink" & CRLF, usernum
    Exit Sub
    End If
    
inpstr = stripOne(inpstr)
Select Case cmdFound
    Case 0   ' quit
        quit usernum
    Case 1   ' emote
        emote usernum, inpstr
    Case 2   ' who
        who usernum
    Case 3   ' say
        say usernum
    Case 4   ' desc
        desc usernum, inpstr
    Case 5   ' commands
        cmdlist (usernum)
    Case 6   ' passwd
        passwd inpstr, usernum
    Case 7   ' cls
        clearScreen (usernum)
    Case 8   ' entpro
        entpro usernum
    Case 9   ' promote
        promote usernum, inpstr
    Case 10  ' demote
        demote usernum, inpstr
    Case 11  ' kill
        murder usernum, inpstr
    Case 12  ' nuke
        nuke usernum, inpstr
    Case 13  ' age
        age usernum, inpstr
    Case 14  ' gender
        gender usernum, inpstr
    Case 15  ' examine
        examine usernum, inpstr
    Case 16  ' look
        look usernum
    Case 17  ' version
        version usernum
    Case 18  ' go
        go_user usernum, inpstr
    Case 19  ' smail
        smail usernum, inpstr
    Case 20  ' read
        read_board usernum
    Case 21  ' move
        moveUser usernum, inpstr
    Case 22  ' wake
        wake usernum, inpstr
    Case 23  ' write
        write_board usernum
    Case 24  ' rmail
        rmail usernum
    Case 25  ' muzzle
        muzzle usernum, inpstr
    Case 26  ' unmuzzle
        unmuzzle usernum, inpstr
    Case 27  ' color
        colorDisplay usernum
    Case 28  ' help
        help usernum, inpstr
    Case 29  ' shutdown
        shutdown usernum, inpstr
    Case 30  ' sing
        sing usernum, inpstr
    Case 31  ' reboot
        reboot usernum, inpstr
    Case 32  ' tell
        tell usernum, inpstr
    Case 33  ' system
        system_info_show usernum
    Case 34  ' shout
        shout usernum, inpstr
    Case 35  ' ranks
        ranksShow usernum
    Case 36  ' clearline
        clearline usernum, inpstr
    Case 37  ' bcast
        bcast usernum, inpstr, False
    Case 38  ' bbcast
        bcast usernum, inpstr, True
    Case 39  ' inmsg
        set_inmsg usernum, inpstr
    Case 40  ' outmsg
        set_outmsg usernum, inpstr
    Case 41  ' vis
        Vis usernum
    Case 42  ' invis
        invis usernum
    Case 43  ' makevis
        make_vis usernum, inpstr
    Case 44  ' makeinvis
        make_invis usernum, inpstr
    Case 45  ' suicide
        suicide usernum, inpstr
    Case 46  ' review
        review usernum, inpstr
    Case 47  ' cbuff
        cbuff usernum
    Case 48  ' home
        home usernum
    Case 49  ' system
        system_info_show usernum
    Case 50  ' afk
        afk usernum, inpstr
    Case 51  ' connect
        connect_netlink usernum, inpstr
    Case 52  ' clone
        clone usernum, inpstr
    Case 53  ' destroy
        destroy usernum, inpstr
    Case 54  ' csay
        cact True, usernum, inpstr
    Case 55  ' cemote
        cact False, usernum, inpstr
    Case 56  ' disconnect
        disconnect_netlink usernum, inpstr
    Case 57  ' echo
        echo_stuff usernum, inpstr
    Case 58  ' show
        show usernum, inpstr
    Case 59  ' think
        think usernum, inpstr
    Case 60  ' semote
        semote usernum, inpstr
    Case 61  ' verify
        verify usernum, inpstr
    Case 62  ' forwarding
        forwarding usernum
    Case 63  ' sendver
        sendver usernum
    Case 64  ' email
        email usernum, inpstr
    Case 65  ' samesite
        samesite usernum, inpstr
    Case 66  ' mailsys
        mailsys usernum
    Case 67  ' kjob
        kill_job usernum, inpstr
    Case 68  ' topic
        topic usernum, inpstr
    Case 69  ' purge
        purge usernum
    Case 70  ' rstat
        rstat usernum, inpstr
    Case 71  ' addhist
        add_history usernum, inpstr
    Case 72  ' swban
        swban usernum
    Case 73  ' map
        map usernum
    Case 74  ' charecho
        charecho usernum
    Case 75  ' public
        set_public usernum
    Case 76  ' private
        set_private usernum
    Case 77  ' rooms
        list_rooms usernum
    Case 78  ' fix
        fix_room usernum
    Case 79  ' unfix
        unfix_room usernum
    Case 80  ' history
        view_history usernum, inpstr
    Case 81  ' version
        pager usernum, inpstr
    Case 82  ' greet
        greet usernum, inpstr
    Case 83  ' site
        site usernum, inpstr
    Case 84  ' netstat
        netstat usernum
    Case 85  ' figlet
        figlet usernum, inpstr
    Case 86  ' ban
        ban usernum, inpstr
    Case 87  ' unban
        unban usernum, inpstr
    Case 88  ' lban
        lban usernum, inpstr
    Case 89  ' rules
        rules usernum
    Case 90  ' ustat
        ustat usernum, inpstr
    Case 91  ' homepage
        setHomepage usernum, inpstr
    Case 92  ' icq
        setIcq usernum, inpstr
    Case 93  ' expire
        expire usernum, inpstr
    Case 94  ' time
        getTime usernum
    Case 95  ' myclones
        myClones usernum
    Case 96  ' allclones
        allClones usernum
    Case 97  ' dmail
        dmail usernum, inpstr
    Case 98  ' wipe
        wipe usernum, inpstr
    Case 99  ' arrest
        arrest usernum, inpstr
    Case 100 ' unarrest
        unarrest usernum, inpstr
    Case 101 ' invite
        invite usernum, inpstr
    Case 102 ' uninvite
        uninvite usernum, inpstr
    Case 103 ' picture
        picture usernum, inpstr
    Case 104 ' preview
        picture usernum, inpstr, False
    Case 105 ' syslog
        syslogView usernum
    Case 106 ' sayto
        sayto usernum, inpstr
    Case 107 ' knock
        knock usernum, inpstr
    Case 108 ' pemote
        pemote usernum, inpstr
    Case 109 ' mutter
        mutter usernum, inpstr
    Case 110 ' people
        people usernum
    Case Else
        send "Invalid command" & CRLF, usernum
        End Select
End Sub

Function containsCorrupt(tocheck As String) As Boolean
'Does this string contain low-ASCII codes?
Dim count As Integer
Dim foo As String
For count = 1 To Len(tocheck)
    foo = Mid$(tocheck, count, 1)
    If Asc(foo) < 97 Or Asc(foo) > 122 Then
        If Asc(foo) < 90 Or Asc(foo) > 65 Then
            If Asc(foo) < 48 Or Asc(foo) > 57 Then
                containsCorrupt = True
                Exit Function
                End If
            End If
        End If
    Next count
End Function

Function containsCorruptNoNums(tocheck As String) As Integer
'Does this string contain low-ASCII codes (numbers are also
'illegal here)
Dim count As Integer
Dim foo As String
For count = 1 To Len(tocheck)
    foo = Mid$(tocheck, count, 1)
    If Asc(foo) < 97 Or Asc(foo) > 122 Then
        If Asc(foo) < 90 Or Asc(foo) > 65 Then
            containsCorruptNoNums = True
            Exit Function
            End If
        End If
    Next count
End Function

Function containsCorruptNumsOnly(tocheck As String) As Integer
'Does this string contain anything besides numbers?
Dim count As Integer
Dim foo As String
For count = 1 To Len(tocheck)
    foo = Mid$(tocheck, count, 1)
    If Asc(foo) < 48 Or Asc(foo) > 57 Then
        containsCorruptNumsOnly = True
        Exit Function
        End If
    Next count
containsCorruptNumsOnly = False
End Function

Sub createNewAccount(ByRef u As UR_OBJECT, usernum As Integer)
'Makes a new user account
Dim filename As String
filename = App.Path & "\USERS\" & LCase$(u.name) & ".D"
'To avoid the "Recreating user" message that we will get when we
'ask saveUserData to save the file, we will create the file
Open filename For Output As #1
Close #1
writeSyslog "Creating new user, ~FB" & u.name
writeHistory u.name, "~FB~OLAccount was initially created~RS"
saveUserData u
loadViewer
End Sub

Sub editorOptions(usernum As Integer, inpstr As String)
Dim count As Integer
Select Case LCase$(inpstr)
    Case "s", "save"
        Select Case user(usernum).editorType
            Case EDITSTATE_ENTPRO
            For count = 1 To user(usernum).editorPos
                user(usernum).profile(count) = user(usernum).Editor(count)
                Next count
                If user(usernum).editorPos < 15 Then
                    For count = (user(usernum).editorPos + 1) To 15
                        user(usernum).profile(count) = ""
                        Next count
                    End If
                writeRoom user(usernum).room, user(usernum).name & " has finished entering his profile" & CRLF
                send "You have finished entering your profile" & CRLF, usernum
            Case EDITSTATE_SMAIL
            For count = 1 To user(usernum).editorPos
                user(usernum).outMail.message(count) = user(usernum).Editor(count)
                Next count
                If user(usernum).editorPos < 15 Then
                    For count = (user(usernum).editorPos + 1) To 15
                        user(usernum).outMail.message(count) = ""
                        Next count
                   smailOut usernum
                    End If
                writeRoom user(usernum).room, user(usernum).name & " has finished writing a mail message" & CRLF
                send "You have finished writing a mail message" & CRLF, usernum
            Case EDITSTATE_BOARD
            Open App.Path & "\Rooms\" & user(usernum).room & ".B" For Append As #1
            Print #1, "From: " & userCap(user(usernum).name) & " [ " & Format$(Now, "dddd d mmmm yyyy") & " at " & Format$(Now, "hh:nn") & " ]" & CRLF;
            For count = 1 To user(usernum).editorPos
                If Not user(usernum).Editor(count) = "" Then
                    Print #1, user(usernum).Editor(count) & CRLF;
                    End If
                Next count
            Close #1
                writeRoom user(usernum).room, user(usernum).name & " writes a message on the message board" & CRLF
                send "You have finished writing a message on the message board" & CRLF, usernum
            End Select
        user(usernum).state = STATE_NORMAL
        user(usernum).listening = True
        user(usernum).editorPos = 0
    Case "v", "view"
        For count = 1 To user(usernum).editorPos
            send user(usernum).Editor(count) & CRLF, usernum
            Next count
        send "~FG(S)ave, " & "~FY(V)iew, " & "~FB(R)edo, " & "~FR(A)bort~RS ", usernum
    Case "r", "redo"
        user(usernum).state = STATE_EDITOR
        user(usernum).editorPos = 0
        lineEditor usernum, user(usernum).inpstr
    Case "a", "abort"
        user(usernum).state = STATE_NORMAL
        user(usernum).listening = True
        user(usernum).editorPos = 0
        send "Done!" & CRLF, usernum
    Case Else
        send "~FG(S)ave, " & "~FY(V)iew, " & "~FB(R)edo, " & "~FR(A)bort~RS ", usernum
        End Select
End Sub

Function fileExists(filename As String) As Boolean
If Not Dir$(filename) = "" Then
    fileExists = True
    Else
        fileExists = False
        End If
End Function

Function getRoom(RoomToSeek As String) As Integer
Dim count As Integer
If RoomToSeek = "" Then
    getRoom = 0
    Exit Function
    End If
For count = LBound(rooms) To UBound(rooms)
    If InStr(UCase$(rooms(count).name), UCase$(RoomToSeek)) = 1 Then
        getRoom = count
        Exit Function
        End If
    Next count
getRoom = 0
End Function

Function getUser(UserName As String) As Integer
Dim count As Integer
For count = 1 To UBound(user)
    If UCase$(user(count).name) = UCase$(UserName) And user(count).state > STATE_LOGIN3 Then
        getUser = count
        Exit Function
        End If
    Next count
getUser = 0
End Function

Sub killUser(Index As Integer)
If user(Index).netlinkType Then
    netout "REMVD " & user(Index).name & LF, user(Index).netlinkFrom
    Else
        mainForm.Socket2(Index).Action = SOCKET_DISCONNECT
        End If
removeUser Index
End Sub

Sub lineEditor(usernum As Integer, inpstr As String)
Dim count As Integer
user(usernum).state = STATE_EDITOR
user(usernum).listening = False
If user(usernum).editorPos = 0 Then
    For count = 1 To 15
        user(usernum).Editor(count) = ""
        Next count
    send "~BB~FW Starting Line Editor ~RS" & CRLF, usernum
    send "~FTMaximum of 15 lines, end with a '.' on a line by itself.~RS" & CRLF & CRLF, usernum
    user(usernum).editorPos = 1
    send "~FB" & Str$(user(usernum).editorPos) & "> ~RS", usernum
    Exit Sub
    End If
If Len(inpstr) < 1 Then
    send "~FB" & Str$(user(usernum).editorPos) & "> ~RS", usernum
    Exit Sub
    End If
If Len(inpstr) > CLIENT_WIDTH Then
    send "Line too long, redo" & CRLF, usernum
    Exit Sub
    End If
If user(usernum).inpstr = "ENDMAIL" Or user(usernum).inpstr = "EMSG" Then
    send "Netlink codes cannot be written in messages" & CRLF, usernum
    Exit Sub
    End If
If user(usernum).inpstr = "." Then
    user(usernum).state = STATE_EDPICK
    send "~FG(S)ave, " & "~FY(V)iew, " & "~FB(R)edo, " & "~FR(A)bort~RS ", usernum
    Exit Sub
    End If
user(usernum).Editor(user(usernum).editorPos) = inpstr
user(usernum).editorPos = user(usernum).editorPos + 1
If user(usernum).editorPos = 15 Then
    user(usernum).state = STATE_EDPICK
    send "~FG(S)ave, " & "~FY(V)iew, " & "~FB(R)edo, " & "~FR(A)bort~RS ", usernum
    Exit Sub
    End If
send "~FB" & Str$(user(usernum).editorPos) & "> ~RS", usernum
End Sub

Sub loadCommands()
lighter "Loading commands"
cmds(0).name = "quit":      cmds(0).rank = 0
cmds(1).name = "emote":     cmds(1).rank = 0
cmds(2).name = "who":       cmds(2).rank = 1
cmds(3).name = "say":       cmds(3).rank = 0
cmds(4).name = "desc":      cmds(4).rank = 1
cmds(5).name = "commands":  cmds(5).rank = 0
cmds(6).name = "passwd":    cmds(6).rank = 1
cmds(7).name = "cls":       cmds(7).rank = 1
cmds(8).name = "entpro":    cmds(8).rank = 1
cmds(9).name = "promote":   cmds(9).rank = 3
cmds(10).name = "demote":   cmds(10).rank = 4
cmds(11).name = "kill":     cmds(11).rank = 3
cmds(12).name = "nuke":     cmds(12).rank = 4
cmds(13).name = "age":      cmds(13).rank = 1
cmds(14).name = "gender":   cmds(14).rank = 1
cmds(15).name = "examine":  cmds(15).rank = 1
cmds(16).name = "look":     cmds(16).rank = 1
cmds(17).name = "version":  cmds(73).rank = 1
cmds(18).name = "go":       cmds(18).rank = 1
cmds(19).name = "smail":    cmds(19).rank = 1
cmds(20).name = "read":    cmds(20).rank = 1
cmds(21).name = "move":     cmds(21).rank = 3
cmds(22).name = "wake":     cmds(22).rank = 2
cmds(23).name = "write":    cmds(23).rank = 2
cmds(24).name = "rmail":     cmds(24).rank = 1
cmds(25).name = "muzzle":   cmds(25).rank = 4
cmds(26).name = "unmuzzle": cmds(26).rank = 4
cmds(27).name = "color":    cmds(27).rank = 1
cmds(28).name = "help":     cmds(28).rank = 0
cmds(29).name = "shutdown": cmds(29).rank = 5
cmds(30).name = "sing":     cmds(30).rank = 2
cmds(31).name = "reboot":   cmds(31).rank = 5
cmds(32).name = "tell":     cmds(32).rank = 1
cmds(33).name = "system":   cmds(33).rank = 3
cmds(34).name = "shout":    cmds(34).rank = 2
cmds(35).name = "ranks":    cmds(35).rank = 1
cmds(36).name = "clearline": cmds(36).rank = 4
cmds(37).name = "bcast":    cmds(37).rank = 3
cmds(38).name = "bbcast":   cmds(38).rank = 4
cmds(39).name = "inmsg":    cmds(39).rank = 2
cmds(40).name = "outmsg":   cmds(40).rank = 2
cmds(41).name = "visible":      cmds(41).rank = 3
cmds(42).name = "invis":    cmds(42).rank = 3
cmds(43).name = "makevis":  cmds(43).rank = 4
cmds(44).name = "makeinvis": cmds(44).rank = 4
cmds(45).name = "suicide":  cmds(45).rank = 0
cmds(46).name = "review":   cmds(46).rank = 2
cmds(47).name = "cbuff":    cmds(47).rank = 2
cmds(48).name = "home":     cmds(48).rank = 1
cmds(49).name = "system":   cmds(49).rank = 2
cmds(50).name = "afk":      cmds(50).rank = 2
cmds(51).name = "connect":  cmds(51).rank = 5
cmds(52).name = "clone":    cmds(52).rank = 3
cmds(53).name = "destroy":  cmds(53).rank = 3
cmds(54).name = "csay":     cmds(54).rank = 3
cmds(55).name = "cemote":   cmds(55).rank = 3
cmds(56).name = "disconnect": cmds(56).rank = 5
cmds(57).name = "echo":     cmds(57).rank = 2
cmds(58).name = "show":     cmds(58).rank = 2
cmds(59).name = "think":    cmds(59).rank = 2
cmds(60).name = "semote":   cmds(60).rank = 2
cmds(61).name = "verify":   cmds(61).rank = 2
cmds(62).name = "forwarding": cmds(62).rank = 2
cmds(63).name = "sendver":  cmds(63).rank = 2
cmds(64).name = "email":    cmds(64).rank = 2
cmds(65).name = "samesite": cmds(65).rank = 3
cmds(66).name = "mailsys":  cmds(66).rank = 4
cmds(67).name = "kjob":     cmds(67).rank = 4
cmds(68).name = "topic":    cmds(68).rank = 2
cmds(69).name = "purge":    cmds(69).rank = 5
cmds(70).name = "rstat":    cmds(70).rank = 4
cmds(71).name = "addhist":  cmds(71).rank = 3
cmds(72).name = "swban":    cmds(72).rank = 5
cmds(73).name = "map":      cmds(17).rank = 1
cmds(74).name = "charecho": cmds(74).rank = 0
cmds(75).name = "public":   cmds(75).rank = 2
cmds(76).name = "private":  cmds(76).rank = 2
cmds(77).name = "rooms":    cmds(77).rank = 2
cmds(78).name = "fix":      cmds(78).rank = 5
cmds(79).name = "unfix":    cmds(79).rank = 5
cmds(80).name = "history":  cmds(80).rank = 4
cmds(81).name = "pager":    cmds(81).rank = 1
cmds(82).name = "greet":    cmds(82).rank = 3
cmds(83).name = "site":     cmds(83).rank = 3
cmds(84).name = "netstat":  cmds(84).rank = 4
cmds(85).name = "figlet":   cmds(85).rank = 3
cmds(86).name = "ban":      cmds(86).rank = 4
cmds(87).name = "unban":    cmds(87).rank = 4
cmds(88).name = "lban":     cmds(88).rank = 3
cmds(89).name = "rules":    cmds(89).rank = 0
cmds(90).name = "ustat":    cmds(90).rank = 2
cmds(91).name = "homepage": cmds(91).rank = 2
cmds(92).name = "icq":   cmds(92).rank = 2
cmds(93).name = "expire":   cmds(93).rank = 5
cmds(94).name = "time":     cmds(94).rank = 2
cmds(95).name = "myclones": cmds(95).rank = 3
cmds(96).name = "allclones": cmds(96).rank = 3
cmds(97).name = "dmail":    cmds(97).rank = 2
cmds(98).name = "wipe":     cmds(98).rank = 4
cmds(99).name = "arrest":   cmds(99).rank = 3
cmds(100).name = "unarrest": cmds(100).rank = 3
cmds(101).name = "invite":  cmds(101).rank = 2
cmds(102).name = "uninvite": cmds(102).rank = 2
cmds(103).name = "picture": cmds(103).rank = 3
cmds(104).name = "preview": cmds(104).rank = 3
cmds(105).name = "syslog":  cmds(105).rank = 3
cmds(106).name = "sayto":   cmds(106).rank = 1
cmds(107).name = "knock":   cmds(107).rank = 2
cmds(108).name = "pemote":  cmds(108).rank = 1
cmds(109).name = "mutter":  cmds(109).rank = 2
cmds(110).name = "people":  cmds(110).rank = 3
End Sub

Sub loadGlobals()
Dim count As Integer
lighter "Loading globals"
system.gatecrash = UBound(ranks)
staffLevel = 3
system.maxIdle = 15

roomAccessLevels(0) = "~FGPUBLIC"
roomAccessLevels(1) = "~FRPRIVATE"
roomAccessLevels(2) = "~FYSTAFF"

swearNames(0) = "~FGOFF"
swearNames(1) = "~FYMIN"
swearNames(2) = "~FRMAX"

BELL = Chr$(7): doubleQuote = Chr$(34)
echoOn = Chr$(255) & Chr$(252) & Chr$(1)
echoOff = Chr$(255) & Chr$(251) & Chr$(1)
'This will format the user states
For count = 0 To UBound(user)
    user(count).state = -1
    user(count).listing = -2
    Next count
For count = LBound(clones) To UBound(clones)
    clones(count).room = -1
    Next count
actPos = UBound(acts)   'They are read backwards
For count = LBound(acts) To UBound(acts)
    acts(count) = -1
    http_acts(count) = -1
    httpConnects(count) = -1
    logins_graph(count) = -1
    Next count
count = 0
For count = 1 To UBound(user)
    user(count).state = -1
    Next count
loadSwears
End Sub

Sub loadSwears()
Dim count As Integer
lighter "Loading swearban list"
ReDim swears(0)
If Not Dir$(App.Path & "\Misc\Swears.S") = "" Then
    Open App.Path & "\Misc\Swears.S" For Input As #1
    Do While Not EOF(1)
        Line Input #1, swears(0)
        count = count + 1
        Loop
    Close #1
    ReDim swears(count - 1)
    count = 0
    Open App.Path & "\Misc\Swears.S" For Input As #1
    Do While Not EOF(1)
        Line Input #1, swears(count)
        count = count + 1
        Loop
    Close #1
    End If
End Sub

Sub loadRooms()
Dim count, count2 As Integer, temp As String
lighter "Loading rooms"
roomsResize
For count = 1 To UBound(rooms)
    With rooms(count)
        .name = ""
        .topic = ""
        .buffer(LBound(rooms(count).buffer)) = ""
        End With
    Next count
rooms(0).name = "Jail"
rooms(0).access = ROOM_STAFF
rooms(0).locked = True
If Dir$(App.Path & "\Rooms", 16) = "" Then
    MkDir App.Path & "\Rooms"
    End If
If Dir$(App.Path & "\Rooms\Rooms.S") = "" Then
    writeSyslog "~FRRoom manifest file not found"
    Exit Sub
    End If
    
Dim linein As LOAD_OBJECT, FromFile As String
Dim rn As Integer, curBlock As Byte, exitCount As Byte
rn = 0 'Bottom room to start at - 1
Open App.Path & "\Rooms\Rooms.S" For Input As #1
Do While Not EOF(1)
    Line Input #1, FromFile
    'Strip out the tabs
    FromFile = Replace(FromFile, Chr$(9), vbNullString)
    FromFile = Trim$(FromFile)
    'Start room block
    If Left$(FromFile, 6) = "[ROOM " Then
        rn = rn + 1
        spliceWords (FromFile)
        If Len(word(1)) <= 1 Then
            MsgBox "Invalid room name in Rooms.S", vbCritical, "SouthWest Error"
            End
            Else
                rooms(rn).name = Left$(word(1), Len(word(1)) - 1)
                End If
        Do
            Line Input #1, FromFile
            FromFile = Trim$(FromFile)
            FromFile = Replace(FromFile, Chr$(9), vbNullString)
            linein = spliceLoad(FromFile)
            Select Case FromFile
                Case "[Settings]"
                    curBlock = 1
                Case "[Exits]"
                    curBlock = 2
                    exitCount = LBound(rooms(LBound(rooms)).exits)
                Case "[END ROOM]"
                    'Do nothing.. This is here so it isnt parsed
                    'with the case else
                Case Else
                    Select Case curBlock
                        Case 1
                        Select Case linein.specifier
                            Case "access"
                                Select Case UCase$(linein.value)
                                    Case "PRIVATE"
                                        rooms(rn).access = ROOM_PRIVATE
                                    Case "STAFF"
                                        rooms(rn).access = ROOM_STAFF
                                    Case Else
                                        rooms(rn).access = ROOM_PUBLIC
                                    End Select
                            Case "fixed"
                                rooms(rn).locked = TF(linein.value)
                            End Select
                        Case 2
                            rooms(rn).exits(exitCount) = FromFile
                            rooms(rn).allExits = LTrim$(rooms(rn).allExits & " " & FromFile)
                            exitCount = exitCount + 1
                        End Select
                End Select
            Loop While Not FromFile = "[END ROOM]"
            'End room block
        End If
    Loop
Close #1
End Sub

Sub loadUserData(u As Integer, Optional validate As Boolean = True)
Dim FromFile As String, parse As Boolean
Dim linein As LOAD_OBJECT, filename As String
user(u).name = userCap(user(u).name)
filename = App.Path & "\Users\" & user(u).name & ".D"
If validate Then
    If Dir$(filename) = "" Then
        Exit Sub
        End If
    End If
Open filename For Input As #1
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
        Select Case LCase$(linein.specifier)
            Case "password"
                user(u).password = linein.value
            Case "lastlogin"
                user(u).lastLogin = Val(linein.value)
            Case "lastsite"
                user(u).site = linein.value
            Case "totaltime"
                user(u).totalTime = Val(linein.value)
            Case "wasonfor"
                user(u).timeon = Val(linein.value)
            Case "desc"
                user(u).desc = linein.value
            Case "muzzled"
                user(u).muzzled = Int(linein.value)
            Case "visible"
                user(u).visible = Int(linein.value)
            Case "gender"
                user(u).gender = linein.value
            Case "age"
                user(u).age = linein.value
            Case "rank"
                user(u).rank = Int(linein.value)
            Case "email"
                user(u).email = linein.value
            Case "forward_email"
                user(u).sfRec = Int(linein.value)
            Case "verified_email"
                user(u).sfVerifyed = Int(linein.value)
            Case "verification_code"
                user(u).sfVercode = Int(linein.value)
            Case "icq"
                user(u).ICQ = linein.value
            Case "logins"
                user(u).logins = Int(linein.value)
            Case "pager"
                user(u).pager = Int(linein.value)
            Case "enter_msg"
                user(u).enterMsg = linein.value
            Case "exit_msg"
                user(u).exitMsg = linein.value
            Case "arrested"
                user(u).arrested = Int(linein.value)
            Case "unread"
                user(u).unread = Int(linein.value)
            Case "charecho"
                user(u).charEchoing = Int(linein.value)
            Case "visible"
                user(u).visible = Int(linein.value)
            Case "expires"
                user(u).expires = Int(linein.value)
            Case "unarrest_lev"
                user(u).unarrestLevel = Int(linein.value)
                End Select
        End If
    Loop
Close #1
End Sub

Function loadUserPassword(UserName As String) As String
Dim filename As String, FromFile As String
Dim linein As LOAD_OBJECT
filename = App.Path & "\USERS\" & LCase$(UserName) & ".D"
If Dir$(filename) = "" Then
    loadUserPassword = ""
    Exit Function
    End If
Open filename For Input As #1
Do While Not EOF(1)
    Line Input #1, FromFile
    linein = spliceLoad(FromFile)
    If LCase$(linein.specifier) = "password" Then
        loadUserPassword = linein.value
        Exit Do
        End If
    Loop
Close #1
End Function

Sub loadColors()
lighter "Loading colors"
colorShorts(0) = "RS": colorValues(0) = Chr$(27) & "[0m"
colorShorts(1) = "OL": colorValues(1) = Chr$(27) & "[1m"
colorShorts(2) = "UL": colorValues(2) = Chr$(27) & "[1m"
colorShorts(3) = "LI": colorValues(3) = Chr$(27) & "[5m"
colorShorts(4) = "RV": colorValues(4) = Chr$(27) & "[7m"
colorShorts(5) = "FK": colorValues(5) = Chr$(27) & "[30m"
colorShorts(6) = "FR": colorValues(6) = Chr$(27) & "[31m"
colorShorts(7) = "FG": colorValues(7) = Chr$(27) & "[32m"
colorShorts(8) = "FY": colorValues(8) = Chr$(27) & "[33m"
colorShorts(9) = "FB": colorValues(9) = Chr$(27) & "[34m"
colorShorts(10) = "FM": colorValues(10) = Chr$(27) & "[35m"
colorShorts(11) = "FT": colorValues(11) = Chr$(27) & "[36m"
colorShorts(12) = "FW": colorValues(12) = Chr$(27) & "[37m"
colorShorts(13) = "BK": colorValues(13) = Chr$(27) & "[40m"
colorShorts(14) = "BR": colorValues(14) = Chr$(27) & "[41m"
colorShorts(15) = "BG": colorValues(15) = Chr$(27) & "[42m"
colorShorts(16) = "BY": colorValues(16) = Chr$(27) & "[43m"
colorShorts(17) = "BB": colorValues(17) = Chr$(27) & "[44m"
colorShorts(18) = "BM": colorValues(18) = Chr$(27) & "[45m"
colorShorts(19) = "BT": colorValues(19) = Chr$(27) & "[46m"
colorShorts(20) = "BW": colorValues(20) = Chr$(27) & "[47m"
colorHex(5) = &H0&
colorHex(6) = &HFF&
colorHex(7) = &H8000&
colorHex(8) = &HFFFF&
colorHex(9) = &HFF0000
colorHex(10) = &HFF00FB
colorHex(11) = &HFFFF00
colorHex(12) = &HFFFFFF
End Sub

Function parseColors(text As String) As String
Dim count As Integer
text = Replace(text, CRLF, "~RS" & CRLF)
For count = LBound(colorValues) To UBound(colorValues)
    text = Replace(text, "~" & colorShorts(count), colorValues(count))
    Next count
parseColors = text
End Function

Sub processNormal(usernum As Integer)
'This is the main input check loop.
user(usernum).inpstr = Trim$(user(usernum).inpstr)
If Len(user(usernum).inpstr) > 1 Then
    If Asc(Left(user(usernum).inpstr, 1)) = 1 Then
        user(usernum).inpstr = Right(user(usernum).inpstr, Len(user(usernum).inpstr) - 1)
        End If
    End If
If user(usernum).inpstr = "" Then
    If user(usernum).atNetlink > -1 Then
        netout "ACT " & user(usernum).name & " NL" & LF, user(usernum).atNetlink
        End If
    Exit Sub
    End If
If Asc(user(usernum).inpstr) = 13 Then
    user(usernum).inpstr = ""
    Exit Sub
    End If

'Ohhh... You were alive all this time? I know it seems like
'a redundant check but this sub is also called by netlinkers.
If user(usernum).afk Then
    user(usernum).afk = False
    writeRoom user(usernum).room, user(usernum).name & " shakes their head and wakes up" & CRLF
    End If

'NUTS Netlinks do things the annoying way.
If user(usernum).netlinkType Then
    If Len(word(2)) > 0 Then
        If Not Left$(user(usernum).inpstr, 1) = "." Then
            user(usernum).inpstr = "." & user(usernum).inpstr
            End If
        End If
    End If
    
'Dont ask, you really dont want to know.
If Asc(user(usernum).inpstr) = 1 Then
    If Len(user(usernum).inpstr) > 1 Then
        user(usernum).inpstr = Right$(user(usernum).inpstr, Len(user(usernum).inpstr) - 1)
        Else
            user(usernum).inpstr = ""
            Exit Sub
            End If
    End If
    
Dim firstchar As String
firstchar = Left(user(usernum).inpstr, 1)
'This will check and see whether the user
'is typing a command/shortcut instead of
'plain speech.
user(usernum).inpstr = swearFilter(user(usernum).inpstr)
If user(usernum).inpstr = Chr$(1) Then
    send "Swearing is not allowed here" & CRLF, usernum
    Exit Sub
    End If
'Make life prettier for the poor saps using charictor echo clients
If user(usernum).charEchoing Then
    send CRLF, usernum
    End If
'For our outgoing netlink users
If user(usernum).atNetlink > -1 Then
    If user(usernum).inpstr = ".home" Then
        home usernum
        End If
    Select Case firstchar
        Case "."
            user(usernum).inpstr = Right$(user(usernum).inpstr, Len(user(usernum).inpstr) - 1)
            netout "ACT " & user(usernum).name & " " & user(usernum).inpstr & LF, user(usernum).atNetlink
        Case ";"
            user(usernum).inpstr = Right$(user(usernum).inpstr, Len(user(usernum).inpstr) - 1)
            netout "ACT " & user(usernum).name & " emote " & user(usernum).inpstr & LF, user(usernum).atNetlink
        Case ":"
            user(usernum).inpstr = Right$(user(usernum).inpstr, Len(user(usernum).inpstr) - 1)
            netout "ACT " & user(usernum).name & " emote " & user(usernum).inpstr & LF, user(usernum).atNetlink
        Case "@"
            user(usernum).inpstr = Right$(user(usernum).inpstr, Len(user(usernum).inpstr) - 1)
            netout "ACT " & user(usernum).name & " who" & LF, user(usernum).atNetlink
        Case Else
            netout "ACT " & user(usernum).name & " say " & user(usernum).inpstr & LF, user(usernum).atNetlink
            End Select
        Exit Sub
    End If

Select Case firstchar
    Case "."
        Commands user(usernum).inpstr, usernum
    Case ";"
        emote usernum, user(usernum).inpstr
    Case ":"
        emote usernum, user(usernum).inpstr
    Case "@"
        who usernum
    Case Else
        say usernum
        End Select
End Sub

Public Sub statbarUsersUpdate()
Dim count As Integer, foundUsers As Integer
For count = 1 To UBound(user)
    If user(count).operational Then
        foundUsers = foundUsers + 1
        End If
    Next count
mainForm.StatusBar.Panels(3).text = Trim$(foundUsers)
End Sub

Public Sub statbarNetlinksUpdate()
Dim count As Integer, foundNetlinks As Integer
For count = 1 To UBound(net)
    If net(count).state > NETSTATES.NETLINK_UP Then
        foundNetlinks = foundNetlinks + 1
        End If
    Next count
mainForm.StatusBar.Panels(4).text = Trim$(foundNetlinks)
End Sub

Sub removeUser(usernum As Integer)
'I am become death, the destroyer of worlds
'                          -Bhagavid-Gita
Dim count As Integer
mainForm.dumpCheck = False
If user(usernum).state > STATE_LOGIN3 Then
    writeRoomExcept "", "~OL~FTLeaving:~RS " & user(usernum).name & " " & user(usernum).desc & CRLF, user(usernum).name
    user(usernum).lastLogin = date2num(Now)
    saveUserData user(usernum)
    If user(usernum).netlinkType Then
        writeSyslog "~FB" & user(usernum).name & "~RS returns to ~FG" & net(s2n(user(usernum).netlinkFrom)).name
        Else
            writeSyslog "~FB" & user(usernum).name & "~RS has disconnected"
            End If
    End If

'Remove all of a user's clones
For count = LBound(clones) To UBound(clones)
    If clones(count).owner = usernum And clones(count).active Then
        writeRoom rooms(clones(count).room).name, "~FMA clone of " & user(usernum).name & " vanishes" & CRLF
        clones(count).active = False
        End If
    Next count

'Ok, this is the rather dirty way to take the user off of the list. There
'are probably more efficient ways to do this but hey, I dont have the time.
'I dont know why you have to check it but I do know that it crashes if you
'dont during attacks, so best do it.
If user(usernum).listing > -1 And mainForm.connectionsList.ListCount = 1 Then
    mainForm.connectionsList.Clear
    ElseIf Not mainForm.connectionsList.ListCount = 0 Then
        mainForm.connectionsList.RemoveItem user(usernum).listing
        End If

'Do the grand shift of the user listings on the connection list
For count = 1 To UBound(user)
    If user(count).listing > user(usernum).listing Then
        user(count).listing = user(count).listing - 1
        End If
    Next count
    
If usernum <= UBound(user) Then
    madeSockets(usernum) = 0
    End If
If Not user(usernum).netlinkType Then
    Unload mainForm.Socket2(usernum)
    End If
cleanUser user(usernum)
userResize
statbarUsersUpdate
If mainForm.tree.SelectedItem.Key Like "UD *" Then
    mainForm.treeLoad
    End If
End Sub

Sub saveUserData(u As UR_OBJECT)
'This function will save the user data to the file. The
'first part does the .D file (for user data) and the second
'part does the profile.
Dim filename As String
Dim count As Integer, msg As String
filename = App.Path & "\USERS\" & LCase$(u.name) & ".D"
If Dir$(filename) = "" Then
    writeSyslog "Recreating user ~FB" & u.name
    End If
Open filename For Output As #1
msg = "password  = " & u.password & CRLF & _
          "lastlogin = " & u.lastLogin & CRLF & _
          "lastsite  = " & u.site & CRLF & _
          "totalTime = " & u.totalTime & CRLF & _
          "wasonfor  = " & u.timeon & CRLF & _
          "desc      = " & u.desc & CRLF & _
          "rank      = " & u.rank & CRLF & _
          "muzzled   = " & Int(u.muzzled) & CRLF & _
          "visible   = " & Int(u.visible) & CRLF & _
          "gender    = " & u.gender & CRLF & _
          "age       = " & u.age & CRLF & _
          "email     = " & u.email & CRLF & _
          "icq       = " & u.ICQ & CRLF & _
          "logins    = " & u.logins & CRLF & _
          "pager     = " & u.pager & CRLF & _
          "enter_msg = " & u.enterMsg & CRLF & _
          "exit_msg  = " & u.exitMsg & CRLF
msg = msg & "arrested  = " & Int(u.arrested) & CRLF & _
          "rank      = " & Int(u.rank) & CRLF & _
          "unread    = " & Int(u.unread) & CRLF & _
          "charecho  = " & Int(u.charEchoing) & CRLF & _
          "visible   = " & Int(u.visible) & CRLF & _
          "expires   = " & Int(u.expires) & CRLF & _
          "forward_email     = " & Int(u.sfRec) & CRLF & _
          "verified_email    = " & Int(u.sfVerifyed) & CRLF & _
          "verification_code = " & Trim$(Str$(u.sfVercode)) & CRLF & _
          "unarrest_lev = " & Trim$(Str$(u.unarrestLevel))
Print #1, msg;
Close #1
msg = vbNullString
If Not u.profile(1) = "" Then
    filename = App.Path & "\USERS\" & LCase$(u.name) & ".P"
    Open filename For Output As #1
    For count = 1 To 15
        If Not u.profile(count) = "" Then
            msg = msg & u.profile(count) & CRLF
            End If
        Next count
    If Len(msg) > 0 Then
        Print #1, msg;
        End If
    Close #1
    End If
End Sub

Sub send(text As String, linenum As Integer)
text = parseColors(text)
If Len(text) > MAX_DATA_LEN Then
    text = Left(text, MAX_DATA_LEN)
    End If
If user(linenum).netlinkType Then
    If Right$(text, 2) = CRLF Then
        Dim newtext As String
        newtext = Replace(text, CRLF, LF)
        If Len(newtext) > 0 Then
            If Right$(newtext, 1) = LF Then
                newtext = Left$(newtext, Len(newtext) - 1)
                End If
            End If
        If newtext = "" Then
            Exit Sub
            End If
        user(linenum).outBuffer = user(linenum).outBuffer & newtext
        netout "MSG " & user(linenum).name & LF & user(linenum).outBuffer & LF & "EMSG" & LF, user(linenum).netlinkFrom
        user(linenum).outBuffer = ""
        Else
            user(linenum).outBuffer = user(linenum).outBuffer & text
            End If
    Else
        If user(linenum).operational And madeSockets(linenum) Then
            If mainForm.Socket2(linenum).Connected Then
                On Error Resume Next
                mainForm.Socket2(linenum).SendLen = Len(text)
                mainForm.Socket2(linenum).SendData = text
                Else
                    removeUser (linenum)
                    Exit Sub
                    End If
            End If
        End If
End Sub
Sub smailOut(usernum As Integer)
Dim count As Integer, msg As String
user(0).name = user(usernum).outMail.receiver
loadUserData 0
For count = 1 To 15
    If Not user(usernum).outMail.message(count) = "" Then
        msg = msg & user(usernum).outMail.message(count) & CRLF
        End If
    Next count
If userIsOnline(user(usernum).outMail.receiver) Then
    send "~FT~OL~LIYOU HAVE NEW MAIL!" & CRLF, getUser(user(usernum).outMail.receiver)
    user(getUser(user(usernum).outMail.receiver)).unread = True
    Else
        user(0).unread = True
        saveUserData user(0)
        If user(0).sfRec Then
            If user(usernum).sfVerifyed Then
                send "You have already verifyed" & CRLF, usernum
                Exit Sub
                End If
            Dim mail_num As Integer
            mail_num = Setup_Address(user(0).name)
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
                End If
            With mail(mail_num)
                .inuse = True
                .u_to = user(0).name
                .u_from = system.emailAddress
                .userid = user(usernum).name
                .timestamp = Now
                .message = cStrip(msg)
                End With
            End If
        End If
Open App.Path & "\Users\" & user(usernum).outMail.receiver & ".M" For Append As #1
Print #1, "From: " & userCap(user(usernum).name) & " [ " & Format$(Now, "dddd d mmmm yyyy") & " at " & Format$(Now, "hh:nn") & " ]" & CRLF;
For count = 1 To 15
    If Not user(usernum).outMail.message(count) = "" Then
        Print #1, user(usernum).outMail.message(count) & CRLF;
        End If
    Next count
Print #1, CRLF;
Close #1
End Sub

Function stripOne(ByVal inpstr As String) As String
If InStr(inpstr, " ") = 0 Then
    stripOne = ""
    Exit Function
    End If
stripOne = Right$(inpstr, Len(inpstr) - InStr(inpstr, " "))
End Function

Function userCap(UserName As String) As String
userCap = StrConv(UserName, vbProperCase)
End Function

Function userExists(UserName As String) As Boolean
Dim count As Integer
If Len(UserName) = 0 Then
    userExists = False
    Exit Function
    End If
If LCase$(Dir$(App.Path & "\Users\" & UserName & ".D")) = LCase$(UserName) & ".d" Then
    userExists = True
    Exit Function
    Else
        userExists = False
        End If
For count = 1 To UBound(user)
    If UCase$(user(count).name) = UCase$(UserName) Then
        userExists = True
        Exit Function
        End If
    Next count
End Function

Function userIsOnline(UserName As String) As Boolean
Dim count As Integer
If UserName = "" Then
    userIsOnline = False
    Exit Function
    End If
For count = 1 To UBound(user)
    If UCase$(user(count).name) = UCase$(UserName) And user(count).state > STATE_LOGIN3 Then
        userIsOnline = True
        Exit Function
        End If
    Next count
userIsOnline = False
End Function

Function wordCount(thing As String) As Integer
Dim foo As String, bar  As String, stringer As String
Dim spacecount As Integer, count As Integer, lastwasspace As Integer
stringer = thing
If stringer = "" Then
    wordCount = 0
    Exit Function
    End If
If Left$(stringer, 1) = " " Then
    wordCount = -1
    Exit Function
    End If
spacecount = 1
lastwasspace = False
For count = 1 To Len(stringer)
    foo = Left$(stringer, 1)
    If foo = " " And lastwasspace = False Then
        spacecount = spacecount + 1
        lastwasspace = True
        Else
            lastwasspace = False
            End If
    If Len(stringer) > 1 Then
        stringer = Right$(stringer, Len(stringer) - 1)
        End If
    Next count
wordCount = spacecount
End Function

Sub writeRoom(room As String, inpstr As String)
Dim count As Integer
If room = "" Then
    For count = 1 To UBound(user)
        If user(count).listening And user(count).operational Then
            send inpstr, count
            End If
        Next count
    Else
        For count = 1 To UBound(user)
            If user(count).room = room And user(count).listening Then
                send inpstr, count
                End If
            Next count
        For count = LBound(clones) To UBound(clones)
            If user(clones(count).owner).listening And rooms(clones(count).room).name = room And Not room = "" Then
                send "~FT[ " & rooms(clones(count).room).name & " ]: ~RS" & inpstr, clones(count).owner
                End If
            Next count
        End If
End Sub

Sub writeRoomExcept(room As String, inpstr As String, UserName As String)
Dim count As Integer
If room = "" Then
    For count = 1 To UBound(user)
    If user(count).listening = True And Not user(count).name = UserName And user(count).operational Then
        send inpstr, count
        End If
    Next count
    Else
        For count = 1 To UBound(user)
            If user(count).listening = True And Not user(count).name = UserName And user(count).room = room Then
                send inpstr, count
                End If
            Next count
        For count = LBound(clones) To UBound(clones)
            If user(clones(count).owner).listening And rooms(clones(count).room).name = room And Not room = "" Then
                send "~FT[ " & rooms(clones(count).room).name & " ]: ~RS" & inpstr, clones(count).owner
                End If
            Next count
        End If
End Sub

Sub writeSyslog(LogEntry As String)
Dim stripped As String
stripped = cStrip(LogEntry)
mainForm.Syslog.AddItem stripped
If mainForm.Syslog.ListCount >= 7 Then
    mainForm.Syslog.RemoveItem 0
    End If
If Not BOOTING Then
    lighter stripped
    End If
writeRawSyslog LogEntry
End Sub

Sub writeRawSyslog(LogEntry As String)
If mainForm.LoggingFile.Checked Then
    LogEntry = Format$(Now, "mm-dd-yyyy hh:nnam/pm  ") & LogEntry
    logbook(logpos) = LogEntry
    logpos = logpos + 1
    If logpos > UBound(logbook) Then
        logpos = LBound(logbook)
        End If
    End If
If Not BOOTING Then
    If mainForm.tree.SelectedItem.Key = "SYSLOG" Then
        mainForm.treeLoad
        End If
    End If
End Sub

Sub lighter(text As String)
mainForm.StatusBar.Panels(2).text = text
mainForm.lightbulbOn True
If mainForm.light_time.Enabled Then
    mainForm.light_time.Enabled = False
    End If
mainForm.light_time.Enabled = True
End Sub

Sub writeRoomBuff(message As String, roomnum As Integer)
Dim count As Integer
For count = 2 To REVBUFF_SIZE
    rooms(roomnum).buffer(count - 1) = rooms(roomnum).buffer(count)
    Next count
rooms(roomnum).buffer(REVBUFF_SIZE) = message
End Sub
Function completeUsername(name As String) As String
Dim count As Integer
For count = 1 To UBound(user)
    If UCase$(user(count).name) = UCase$(name) Then
        completeUsername = user(count).name
        Exit Function
        End If
    Next count
For count = 1 To UBound(user)
    If InStr(UCase$(user(count).name), UCase$(name)) = 1 Then
        completeUsername = user(count).name
        Exit Function
        End If
    Next count
End Function

Function crypt(pass As String) As String
Dim password As String, passhold As String
If Len(pass) <= 1 Then
    crypt = ""
    Exit Function
    End If
'I know it looks funky but xcrypt erases the password
'given to it for security reasons
passhold = pass
password = xcrypt(passhold)
crypt = Left$(password, InStr(password, Chr$(0)) - 1)
End Function

Function alreadyLoggedOn(u As UR_OBJECT) As UR_OBJECT
'This will check and see if a user is already logged onto the server
Dim count As Integer, uCopy As UR_OBJECT
For count = 1 To UBound(user)
    If user(count).state > STATE_LOGIN3 And UCase$(user(count).name) = UCase$(u.name) And Not u.line = user(count).line Then
        uCopy = u
        u = user(count)
        send "Switching sessions..." & CRLF, count
        u.line = uCopy.line
        u.Index = uCopy.Index: u.state = STATE_NORMAL: u.oldInpstr = ""
        u.inpstr = "": u.afk = False: u.idle = 0
        u.listing = uCopy.listing: u.netlinkType = False
        u.Index = uCopy.Index: u.line = uCopy.line
        u.listening = True
        u.site = uCopy.site
        send "Already logged in..." & CRLF, u.line
        killUser count
        alreadyLoggedOn = u
        Exit Function
        End If
    Next count
alreadyLoggedOn = u
End Function

Function spliceTime(ByVal secs As Double) As DHMS_OBJECT
spliceTime.days = Int(secs / 86400)
secs = secs Mod 86400
spliceTime.hours = Int(secs / 3600)
secs = secs Mod 3600
spliceTime.minutes = Int(secs / 60)
spliceTime.seconds = secs Mod 60
End Function

Function runPurge(Data As PURGE_OBJECT) As PURGE_OBJECT
'It may seem odd that I have broken the delete and find sections
'into two parts but the reason for this is that the deleteAccount
'sub also calls the Dir function with arguments so it erases the
'Dir functions' track of our file search.
Dim count As Integer, usr As String, diff As Long, delme As String
Dim hold As String
usr = Dir$(App.Path & "\Users\*.D")
If system.purgeLength < 1 Then     'Do a default if it has
    system.purgeLength = 45        'not been set
    End If
Do While Not usr = ""
    If Len(usr) > 2 Then
        usr = Left$(usr, Len(usr) - 2)
        End If
    user(0).name = usr
    loadUserData 0, False
    diff = Abs(DateAdd("d", num2date(user(0).lastLogin), -Now))
    If diff > system.purgeLength And user(0).expires Then
        If userIsOnline(usr) Then
            runPurge.usersNow = runPurge.usersNow + 1
            Else
                runPurge.usersRemoved = runPurge.usersRemoved + 1
                If delme = vbNullString Then
                    delme = usr & ","
                    Else
                        delme = delme & usr & ","
                        End If
                End If
        Else
            runPurge.usersNow = runPurge.usersNow + 1
            End If
    usr = Dir$()
    Loop
If runPurge.usersRemoved = 0 Then
    Exit Function
    End If
For count = 1 To runPurge.usersRemoved
    deleteAccount Left$(delme, InStr(delme, ",") - 1), False
    delme = Right$(delme, Len(delme) - InStr(delme, ","))
    Next count
loadViewer
End Function

Sub spliceWords(ByVal foo As String)
Dim lenToSpace As Integer, count As Integer
For count = 0 To MAX_ARGS
    lenToSpace = InStr(foo, " ")
    If lenToSpace > 0 Then
        word(count) = Left$(foo, lenToSpace - 1)
        foo = Right$(foo, Len(foo) - lenToSpace)
        Else
            word(count) = foo
            foo = ""
            End If
    Next count
End Sub

Sub writeHistory(UserName As String, entry As String)
entry = Format$(Now, "mm-dd-yyyy hh:nnam/pm  ") & entry
Open App.Path & "\Users\" & UserName & ".His" For Append As #1
Print #1, entry & CRLF;
Close #1
If mainForm.tree.SelectedItem.Key = "UH " & UCase$(UserName) Then
    mainForm.treeLoad
    End If
End Sub

Function swearFilter(text As String) As String
Dim count As Integer, tText As String
'Note that it even checks for color-code bypasses in MAX mode
Select Case system.swearing
    Case swbanLevels.SWEAR_MIN
        For count = LBound(swears) To UBound(swears)
            text = Replace(text, swears(count), String$(Len(swears(count)), "*"))
            Next count
    Case swbanLevels.SWEAR_MAX
        tText = cStrip(text)
        For count = LBound(swears) To UBound(swears)
            If InStr(tText, swears(count)) And Not swears(count) = "" Then
                text = Chr$(1)
                End If
            Next count
    End Select
swearFilter = text
End Function

Function stripLast(ByVal text As String) As String
Dim iPos As Integer, temp As String
text = RTrim$(text)
Do
    If iPos > 0 Then
        iPos = InStr(iPos + 1, text, " ")
        Else
            iPos = InStr(text, " ")
            End If
    If Not iPos = 0 Then
        stripLast = RTrim$(Left$(text, iPos))
        End If
    Loop While iPos > 0
If stripLast = text Then
    stripLast = text
    End If
End Function

Function reline(text As String) As String
Dim count As Integer
If text = "" Then Exit Function
text = Replace(text, LF & colorValues(0), "")
text = Replace(text, LF, LF & "~RS")
text = Replace(text, LF, "")
text = Replace(text, LF, CRLF & colorValues(0))
reline = text
End Function

Function usersOnline() As Integer
Dim count As Integer
For count = 1 To UBound(user)
    If user(count).operational And user(count).state > STATE_LOGIN3 Then
        usersOnline = usersOnline + 1
        End If
    Next count
End Function

Sub loadFiglets()
Dim count As Integer, FromFile As String, count2 As Integer
lighter "Loading figlet fonts"
If Dir$(App.Path & "\Figlets\" & system.figlet & ".F") = "" Or Trim$(LCase$(system.figlet)) = "(none)" Then
    figbar = ""
    ReDim figlets(0, 0)
    figletHeight = 0
    figletWidth = 0
    Exit Sub
    End If
Open App.Path & "\Figlets\" & system.figlet & ".F" For Input As #1
Line Input #1, FromFile
figletHeight = Int(FromFile)
Line Input #1, FromFile
figletWidth = Int(FromFile)
Line Input #1, figbar
ReDim figlets(Len(figbar), figletHeight)
For count = 0 To Len(figbar)
    For count2 = 1 To figletHeight
        If EOF(1) Then Exit For
        Line Input #1, FromFile
        Select Case Len(FromFile)
            Case Is < figletWidth
                FromFile = FromFile & Space(figletWidth - Len(FromFile))
            Case Is > figletWidth
                FromFile = Left$(FromFile, figletWidth)
                End Select
        figlets(count, count2) = FromFile
        Next count2
        If EOF(1) Then
            Exit For
            End If
        Line Input #1, FromFile 'Dummy line
    Next count
Close #1
End Sub

Public Function TF(boolstr As String) As Boolean
boolstr = UCase$(boolstr)
If boolstr = "TRUE" Or boolstr = "YES" Or boolstr = "ON" Then
    TF = True
    End If
End Function

Public Function spliceLoad(tosplice As String) As LOAD_OBJECT
Dim toequal As Integer
toequal = InStr(tosplice, "=")
If toequal <= 1 Or toequal = Len(tosplice) Then
    Exit Function
    End If
spliceLoad.specifier = LCase$(Trim$(Left$(tosplice, toequal - 1)))
spliceLoad.value = Trim$(Right$(tosplice, Len(tosplice) - toequal))
End Function

Function BoolYN(boolin) As String
If boolin Then
    BoolYN = "Yes"
    Else
        BoolYN = "No"
        End If
End Function

Function Bin2YN(inp As Boolean) As String
If inp Then
    Bin2YN = "YES"
    Else
        Bin2YN = "NO"
        End If
End Function

Function int2sw(intp As Integer) As String
Select Case system.swearing
    Case SWEAR_MIN
        int2sw = "MIN"
    Case SWEAR_MAX
        int2sw = "MAX"
    Case Else
        int2sw = "OFF"
        End Select
End Function

Public Function bin2TF(intc As Boolean) As String
If intc Then
    bin2TF = "TRUE"
    Else
        bin2TF = "FALSE"
        End If
End Function

Sub postLoadup(ByRef u As UR_OBJECT)
u.state = STATE_NORMAL
masterLogin u
u.site = mainForm.Socket2(u.line).PeerName
End Sub

Sub masterLogin(ByRef u As UR_OBJECT)
If u.netlinkType Then
    u.site = "Netlink"
    End If
u.timeon = 0
u.logins = u.logins + 1
u.lastLogin = date2num(Now)
u.listening = True
userLoginsGraph = userLoginsGraph + 1
statbarUsersUpdate
If mainForm.tree.SelectedItem.Key Like "UD *" Then
    mainForm.treeLoad
    End If
If u.arrested Then
    u.room = "Jail"
    End If
End Sub

Sub deleteAccount(UserName As String, Optional refresh As Boolean = True)
Dim filename As String
filename = App.Path & "\USERS\" & LCase$(UserName) & ".D"
If Not Dir$(filename) = "" Then
    Kill filename
    End If
filename = App.Path & "\USERS\" & LCase$(UserName) & ".P"
If Not Dir$(filename) = "" Then
    Kill filename
    End If
filename = App.Path & "\USERS\" & LCase$(UserName) & ".M"
If Not Dir$(filename) = "" Then
    Kill filename
    End If
filename = App.Path & "\USERS\" & LCase$(UserName) & ".His"
If Not Dir$(filename) = "" Then
    Kill filename
    End If
If refresh Then
    loadViewer
    End If
mainForm.tree_NodeClick mainForm.tree.Nodes("USER DATA")
End Sub

Function mold(ByVal strin As String, size As Integer, Optional fill As String = " ") As String
mold = strin
Select Case cLen(mold)
    Case Is < size
        mold = mold & String$(size - cLen(mold), fill)
    Case Is > size
        mold = Left$(strin, size + (Len(strin) - cLen(strin)))
    End Select
End Function

Function num2date(DN As Double) As Date
num2date = DateAdd("s", DN, #1/1/1970#)
End Function

Function date2num(DN As Date) As Double
date2num = DateDiff("s", #1/1/1970#, DN)
End Function

Function cStrip(ByVal text As String) As String
Dim count As Integer
For count = LBound(colorShorts) To UBound(colorShorts)
    text = Replace(text, "~" & colorShorts(count), "")
    Next count
cStrip = text
End Function

Function cLen(ByVal text As String) As Long
cLen = Len(cStrip(text))
End Function

Public Sub ShellSort(ByRef Arr() As String)
'I did not write this so do not take it out on me. I mearly reformated
'the code a little. A is the array, Lb the lower boundry and Ub is
'the upper boundry.
Dim N As Long, h As Long, i As Long, j As Long, t As Variant, lb As Integer, ub As Integer
lb = LBound(Arr)
ub = UBound(Arr)
N = ub - lb + 1
h = 1
If (N < 14) Then
    h = 1
    Else
        Do While h < N
            h = 3 * h + 1
            Loop
        h = h \ 3
        h = h \ 3
        End If

Do While h > 0
    ' sort by insertion in increments of h
    For i = lb + h To ub
        t = Arr(i)
        For j = i - h To lb Step -h
            If Arr(j) <= t Then
                Exit For
                End If
            Arr(j + h) = Arr(j)
            Next j
        Arr(j + h) = t
        Next i
    h = h \ 3
    Loop
End Sub

Sub cleanUser(ByRef u As UR_OBJECT, Optional volatile As Boolean = True)
u.afk = False
u.age = "Unknown"
u.arrested = False
u.atNetlink = -1
u.charEchoing = False
u.cloneCount = 0
u.desc = "is new here"
u.email = "Unset"
u.enterMsg = "enters from the"
u.exitMsg = "goes off to the"
u.expires = True
u.gender = "Neither"
u.ICQ = "Unset"
u.idle = 0
u.Index = -1
u.inpstr = vbNullString
u.invitations = vbNullString
u.listening = True
u.logins = 0
u.muzzled = False
u.netlinkType = False
u.netlinkPending = False
u.oldInpstr = vbNullString
u.outBuffer = vbNullString
u.pager = 32
u.rank = 1
u.room = rooms(1).name
u.site = "Nobody knows"
u.timeon = 0
u.unread = False
u.url = "Unset"
u.visible = True
u.sfVercode = gen_ver_code
If volatile Then
    u.totalTime = 0
    u.state = STATE_LOGIN1
    u.operational = False
    u.name = vbNullString
    End If
End Sub

Function isNameValid(testName As String) As Boolean
Dim count As Integer
'Check length
If Len(testName) < 3 Or Len(testName) > 12 Then
    isNameValid = False
    Exit Function
    Else
        isNameValid = True
        End If
'Check for symbols and numbers and punctuation and such
If containsCorruptNoNums(testName) Then
    isNameValid = False
    End If
'Check name for swearing
For count = LBound(swears) To UBound(swears)
    If InStr(1, testName, swears(count), vbTextCompare) Then
        isNameValid = False
        End If
    Next count
End Function

Sub userResize(Optional addUser As Boolean = False, Optional downsizeOk As Boolean = False)
'This function is extremly important. Until almost the end of the
'SouthWest project, the user array was fixed to the size of MAX_USERS
'which by default was 250. Can you immagine how much ram this was
'taking up and the wasted processor time spent constantly searching
'through a primarily empty array! This function will automaticly
'resize the array to it's most efficient setting.
Dim tempUser() As UR_OBJECT, newsize As Integer, count As Integer
Dim prospectFree As Boolean, foundFreeSpace As Boolean, found As Boolean
Dim upper As Integer
Const USER_ARR_MIN_SIZE = 2
'Find the size we need automaticly. Remember, users are relative to
'their sockets, so we can't be too cheap and restack them.
For count = 1 To UBound(user)
    If user(count).operational Then
        newsize = count
        found = True
        If prospectFree Then
            foundFreeSpace = True
            End If
        Else
            prospectFree = True
            End If
    Next count
If addUser And Not foundFreeSpace And found Then
    newsize = newsize + 1
    End If
If newsize > maxUsers And Not maxUsers < 2 Then
    newsize = maxUsers
    End If
If newsize < USER_ARR_MIN_SIZE Then
    newsize = USER_ARR_MIN_SIZE
    End If
If newsize = UBound(user) Then
    Exit Sub
    End If
upper = UBound(user)
If newsize > upper Or downsizeOk Then
    ReDim Preserve user(newsize)
    End If
For count = 1 To UBound(user)
    If Not user(count).operational Then
        cleanUser user(count)
        End If
    Next count
End Sub

Function isBanned(check As String, banMethod As banTypes) As Boolean
Dim free As Integer, FromFile As String, file As String
free = FreeFile
file = getBanFile(banMethod)
If fileExists(file) Then
    Open file For Input As #free
    Do While Not EOF(free)
        Line Input #free, FromFile
        If LCase$(check) Like LCase$(Trim$(FromFile)) Then
            isBanned = True
            Close #free
            Exit Function
            End If
        Loop
    Close #free
    End If
End Function

Function getBanFile(banMethod As banTypes) As String
Select Case banMethod
    Case BAN_USER
        getBanFile = App.Path & "\Misc\BannedUsers.S"
    Case BAN_SITE
        getBanFile = App.Path & "\Misc\BannedSites.S"
    Case BAN_NEW
        getBanFile = App.Path & "\Misc\BannedNews.S"
        End Select
End Function

Function roomsResize() As Integer
Dim count As Integer, FromFile As String, free As Integer
count = 1
If Not fileExists(App.Path & "\Rooms\Rooms.S") Then
    ReDim rooms(-1 To count)
    Exit Function
    End If
free = FreeFile
Open App.Path & "\Rooms\Rooms.S" For Input As #free
Do While Not EOF(free)
    Line Input #free, FromFile
    count = count + 1
    Loop
Close #free
ReDim Preserve rooms(-1 To count)
End Function

Function deriveTimeString(t As DHMS_OBJECT, Optional addAgo As Boolean = True)
Dim hi As Integer, count As Integer
If t.days > 0 Then
    deriveTimeString = deriveTimeString & Trim$(t.days) & " days"
    End If
If t.hours > 0 Then
    If Len(deriveTimeString) > 0 Then
        deriveTimeString = deriveTimeString & ", "
        End If
    deriveTimeString = deriveTimeString & Trim$(t.hours) & " hours"
    End If
If t.minutes > 0 Then
    If Len(deriveTimeString) > 0 Then
        deriveTimeString = deriveTimeString & ", "
        End If
    deriveTimeString = deriveTimeString & Trim$(t.minutes) & " minutes"
    End If
count = 1
Do
    count = InStr(count + 1, deriveTimeString, ",")
    hi = hi + 1
    Loop While Not count = 0
If hi > 2 Then
    deriveTimeString = Replace(deriveTimeString, ",", "[")
    deriveTimeString = Replace(deriveTimeString, "[", ",", , hi - 2)
    deriveTimeString = Replace(deriveTimeString, "[", ", and")
    ElseIf hi = 2 Then
        deriveTimeString = Replace(deriveTimeString, ",", " and")
        End If
If deriveTimeString = "" Then
    deriveTimeString = "Just a few seconds"
    End If
deriveTimeString = Replace(deriveTimeString, " 1 minutes", " 1 minute")
deriveTimeString = Replace(deriveTimeString, " 1 hours", " 1 hour")
deriveTimeString = Replace(deriveTimeString, " 1 days", " 1 day")
If addAgo Then
    deriveTimeString = deriveTimeString & " ago"
    End If
End Function

Function getOrdinal(num As Integer) As String
Dim strNum As String, strNum2 As String
strNum = Trim$(num)
If Len(strNum) > 2 Then
    strNum2 = Right$(strNum, 2)
    End If
If Len(strNum) > 1 Then
    strNum = Right$(strNum, 1)
    End If
If strNum Like "1?" Then
    getOrdinal = "th"
    Exit Function
    End If
Select Case strNum
    Case 1
        getOrdinal = "st"
    Case 2
        getOrdinal = "nd"
    Case 3
        getOrdinal = "rd"
    Case Else
        getOrdinal = "th"
        End Select
End Function

Function embedBar(Optional text As String, Optional plusCol As String = "~FR", Optional barCol As String = "~FG", Optional textCol As String = "~FY")
embedBar = plusCol & "+" & barCol & "------" & textCol & text & barCol
embedBar = embedBar & String$(CLIENT_WIDTH - (Len(text) + 8), "-") & plusCol & "+~RS" & CRLF
End Function

Function bool2YN(bool As Boolean) As String
If bool Then
    bool2YN = "Yes"
    Else
        bool2YN = "No"
        End If
End Function

Sub resizeClones(Optional add As Boolean)
'Resizes the clone array or adds an entry
Dim count As Integer, peekClone As Integer, emptyFound As Boolean
For count = LBound(clones) To UBound(clones)
    If clones(count).active Then
        peekClone = count
        Else
            emptyFound = True
            End If
    Next count
If Not emptyFound Then
    peekClone = peekClone + 1
    End If
If peekClone < 2 Then
    peekClone = 2
    End If
If Not peekClone = UBound(clones) Then
    ReDim Preserve clones(count)
    End If
End Sub

Function getMessageCount(file As String) As Integer
'Counts how many messages there are in a file. This function can
'be used for both mail and board messages since their formats are
'roughly the same.
Dim numfound As Integer, inpbuff As String, free As Integer
'If it doesn't exist, there are no entries
If Not fileExists(file) Then
    getMessageCount = 0
    Exit Function
    End If
free = FreeFile
Open file For Input As #free
'Loop until the end of the file
Do
    Line Input #free, inpbuff
    'If a non-null line is found, loop down until the end of the
    'message.
    If Not inpbuff = vbNullString Then
        numfound = numfound + 1
        'Loop until the end of the file or a null line is found
        Do Until EOF(free) Or inpbuff = vbNullString
            Line Input #free, inpbuff
            Loop
        End If
    Loop While Not EOF(free)
Close #free
getMessageCount = numfound
End Function

Function deleteMessages(file As String, dele As String) As Boolean
Dim count As Integer, num1 As Integer, num2 As Integer
If Not fileExists(file) Then
    deleteMessages = False
    Exit Function
    End If
dele = LCase$(dele)
spliceWords dele
num1 = Int(Val(word(1)))
num2 = Int(Val(word(3)))
Select Case word(0)
    Case "all"
        Kill file
    Case "to"
        If num1 = 0 Or num1 > getMessageCount(file) Then
            deleteMessages = False
            Exit Function
            End If
        For count = num1 To 1 Step -1
            deleteMessageByNumber file, count
            Next count
    Case "from"
        If num1 = 0 Or num2 > getMessageCount(file) Or num2 < num1 Then
            deleteMessages = False
            Exit Function
            End If
        For count = num2 To num1 Step -1
            deleteMessageByNumber file, count
            Next count
    Case Else
        num1 = Int(Val(word(0)))
        If num1 < 1 Or num1 > getMessageCount(file) Then
            deleteMessages = False
            Exit Function
            End If
        deleteMessageByNumber file, num1
        End Select
deleteMessages = True
End Function

Sub deleteMessageByNumber(file As String, num As Integer)
'Deletes a board or mail message by number
Dim curMsg As Integer, free As Integer, done As Boolean
Dim FromFile As String, newBuff As String, Top As Integer
Top = getMessageCount(file) ' - 1
If num = Top Then
    Top = Top - 1
    End If
free = FreeFile
Open file For Input As #free
'Loop until we hit the message, recording everything that is not
'in that message into a buffer
Do
    Line Input #free, FromFile
    If FromFile = vbNullString Then
        If Not curMsg = num And curMsg < Top Then
            newBuff = newBuff & CRLF
            End If
        Else
            curMsg = curMsg + 1
            Do
                If Not curMsg = num Then
                    newBuff = newBuff & FromFile & CRLF
                    End If
                Line Input #free, FromFile
                Loop Until FromFile = vbNullString Or EOF(free)
            If Not curMsg = num And curMsg < Top Then
                newBuff = newBuff & CRLF
                End If
            End If
    Loop Until done Or EOF(free)
Close #free
'Save the new buffer
Open file For Output As #free
Print #1, newBuff & CRLF;
Close #free
If getMessageCount(file) = 0 Then
    Kill file
    End If
End Sub

Function stripLastChar(inpstr As String) As String
If Len(inpstr) > 1 Then
    stripLastChar = Left$(inpstr, Len(inpstr) - 1)
    Else
        stripLastChar = ""
        End If
End Function

Function isUserInvited(usernum, roomname As String) As Boolean
If InStr(user(usernum).invitations, UCase$(roomname)) Then
    isUserInvited = True
    End If
End Function

Public Function apEngine() As Boolean
Dim count As Integer, numoflines As Long, DummyVar As String
Dim FromFile As String, message As String, count2 As Integer
Dim firstchar As String, fileopen As Boolean, argstr As String
Dim winDir As String
patcherForm.cop.Caption = "Initializing"
lastprog = 0
If Not fileExists(App.Path & autoPatchFile) And autoPatchFile = "\master.ap" Then
    MsgBox "AutoPatch could not locate the file 'master.ap' that is needed to run SouthWest Autopatch.", vbCritical
    apEngine = False
    Exit Function
    ElseIf Not fileExists(App.Path & autoPatchFile) Then
        Exit Function
        End If
apEngine = True
Open App.Path & autoPatchFile For Input As #1
If LOF(1) = 0 Then
    MsgBox "The setup file contains no instructions."
    End
    Exit Function
    End If
Do While Not EOF(1)
    numoflines = numoflines + 1
    Line Input #1, DummyVar
    Loop
Close #1
Open App.Path & autoPatchFile For Input As #1
For count = 1 To numoflines
    DoEvents
    Line Input #1, FromFile
    'If it is a null line, we ignore it, if not we process it
    If Len(FromFile) > 0 Then
        firstchar = Left$(FromFile, 1)
        percentage (count / numoflines * 100)
        spliceWords (FromFile)
            If firstchar = ">" Then
                argstr = stripOne(FromFile)
                Else
                    argstr = stripLastChar(stripOne(FromFile))
                    End If
            Select Case firstchar
                Case "!"
                'This is a comment and we will ignore it
                Case "["
                'Telling us to perform an operation
                Select Case word(0)
                    Case "[NAME"
                        patcherForm.setupName.Caption = argstr
                    Case "[COPY"
                        If fileopen Then
                            Close #2
                            End If
                        winDir = String$(100, 0)
                        Call GetWindowsDirectory(winDir, 99)
                        winDir = Left$(winDir, InStr(winDir, Chr$(0)) - 1)
                        argstr = Replace$(argstr, "<APP>", App.Path)
                        argstr = Replace$(argstr, "<WIN>", winDir)
                        argstr = Replace$(argstr, "<ROOT>", Left$(winDir, 2))
                        spliceWords (argstr)
                        FileCopy word(0), word(1)
                    Case "[DELETE"
                        If fileopen Then
                            Close #2
                            End If
                        winDir = String$(100, 0)
                        Call GetWindowsDirectory(winDir, 99)
                        winDir = Left$(winDir, InStr(winDir, Chr$(0)) - 1)
                        argstr = Replace$(argstr, "<APP>", App.Path)
                        argstr = Replace$(argstr, "<WIN>", winDir)
                        argstr = Replace$(argstr, "<ROOT>", Left$(winDir, 2))
                        spliceWords (argstr)
                        If Not Dir$(word(0)) = "" Then
                            Kill word(0)
                            End If
                    Case "[SET_MSG"
                        'Set the current operation
                        patcherForm.cop.Caption = argstr
                    Case "[WIN_BOX"
                        argstr = Replace(argstr, "<NEWLINE>", CRLF)
                        argstr = Replace(argstr, "<TAB>", Chr$(9))
                        MsgBox argstr, vbInformation + vbOKOnly + vbApplicationModal, "SouthWest AutoPatch - " & patcherForm.setupName.Caption
                    Case "[OPEN_FILE"
                        If fileopen Then
                            Close #2
                            fileopen = False
                            End If
                        If fileExists(argstr) Then
                            Kill argstr
                            End If
                        Open argstr For Output As #2
                        fileopen = True
                        Case "[MAKE_DIR"
                        If Dir$(argstr, 16) = "" Then
                            MkDir argstr
                            End If
                    End Select
                Case ">"
                    Print #2, argstr & CRLF;
                End Select
        End If
    Next count
If fileopen Then
    Close #2
    End If
Close #1
End Function

Public Sub percentage(percent As Integer)
Dim count As Integer
If lastprog = percent And Not lastprog = 0 Then
    Exit Sub
    End If
If lastprog > percent Or lastprog = 0 Then
    patcherForm.Progress_Bar.Cls
    lastprog = 0
    End If
For count = Int((lastprog / 100) * patcherForm.Progress_Bar.Width) To Int(patcherForm.Progress_Bar.Width * (percent / 100))
    DoEvents
    patcherForm.Progress_Bar.Line (count, 0)-(count, patcherForm.Progress_Bar.Height)
    Next count
patcherForm.Label4.Caption = Str$(Int(percent)) & "%"
lastprog = percent
End Sub

Sub changePortBinding(ByRef sock As Socket, newPort As Integer)
Select Case sock.name
    Case "Socket1"
        If Not sock.LocalPort = newPort Then
            mainForm.Caption = "SouthWest - " & system.mainPort
            writeSyslog "Redirecting main port to " & newPort
            End If
    Case "Netlink"
        If Not sock.LocalPort = newPort Then
            writeSyslog "Redirecting Netlink port to " & newPort
            End If
    Case "http"
        If Not sock.LocalPort = newPort Then
            writeSyslog "Redirecting HTTP port to " & newPort
            End If
        End Select
sock.Action = SOCKET_CLOSE
sock.LocalPort = newPort
sock.Listen
End Sub

