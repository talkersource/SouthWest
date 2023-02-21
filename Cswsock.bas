Attribute VB_Name = "Cswsock"
'I have taken out many of these constants because they were not
'used and I wanted to save up on memory. If you need the entire
'set up constants or help on the SocketWrench control, visit
'the Catalyst website and download the SocketWrench control free
'of charge. (Remember, not SocketTools, SocketWrench is what we want)
'I also commented out my additions to the constants supplied with
'the catalyst control.

'
'Telnet Options
'
Public Const TELCMD_IAC = 255
Public Const TELCMD_DONT = 254
Public Const TELCMD_DO = 253
Public Const TELCMD_WONT = 252
Public Const TELCMD_WILL = 251
Public Const TELCMD_SB = 250
Public Const TELCMD_NOP = 241
Public Const TELCMD_SE = 240
'Public Const TELCMD_AYT = 246 'The following telnet commands
'Public Const TELCMD_EC = 247  'were not, for some reason, put
'Public Const TELCMD_EL = 248  'in the SocketWrench constants
'Public Const TELCMD_AO = 245  'module so I added them to make
'Public Const TELCMD_GA = 249  'it more complete. For info see
'Public Const TELCMD_IP = 244  'RFC 854 at the following URL:
'Public Const TELCMD_BRK = 243 'http://andrew2.andrew.cmu.edu/rfc/rfc854.html_who

Public Const TELOPT_TTYPE = 24
Public Const TELQUAL_IS = 0
'
' SocketWrench Control Actions
'
Public Const SOCKET_OPEN = 1
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_LISTEN = 3
Public Const SOCKET_ACCEPT = 4
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_CLOSE = 7
Public Const SOCKET_DISCONNECT = 7
Public Const SOCKET_ABORT = 8

'
' SocketWrench Control States
'
Public Const SOCKET_NONE = 0
Public Const SOCKET_IDLE = 1
Public Const SOCKET_LISTENING = 2
Public Const SOCKET_CONNECTING = 3
Public Const SOCKET_ACCEPTING = 4
Public Const SOCKET_RECEIVING = 5
Public Const SOCKET_SENDING = 6
Public Const SOCKET_CLOSING = 7

'
' Socket Address Families
'
Public Const AF_INET = 2

'
' Socket Types
'
Public Const SOCK_STREAM = 1

'
' Protocol Types
'
Public Const IPPROTO_TCP = 6

'
' SocketWrench Error Response
'
Public Const SOCKET_ERRIGNORE = 0
Public Const SOCKET_ERRDISPLAY = 1
