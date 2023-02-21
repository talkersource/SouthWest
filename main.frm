VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SouthWest"
   ClientHeight    =   4950
   ClientLeft      =   1380
   ClientTop       =   1995
   ClientWidth     =   6855
   ForeColor       =   &H80000008&
   Icon            =   "main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4950
   ScaleWidth      =   6855
   Begin SocketWrenchCtrl.Socket mailsock 
      Left            =   5190
      Top             =   3750
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin SocketWrenchCtrl.Socket Netlink 
      Index           =   0
      Left            =   4350
      Top             =   3750
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   5000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin SocketWrenchCtrl.Socket Socket3 
      Left            =   3510
      Top             =   3750
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   -1  'True
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin SocketWrenchCtrl.Socket Socket2 
      Index           =   0
      Left            =   3120
      Top             =   3750
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   2730
      Top             =   3750
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin SocketWrenchCtrl.Socket http 
      Index           =   0
      Left            =   5610
      Top             =   3750
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin MSComctlLib.ImageList ToolbarImageList 
      Left            =   2190
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":0466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":05C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":0722
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":153E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox action_frm 
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   2070
      ScaleHeight     =   2985
      ScaleWidth      =   4785
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   -4500
      Width           =   4785
      Begin VB.PictureBox g 
         ForeColor       =   &H000000FF&
         Height          =   2175
         Left            =   300
         ScaleHeight     =   2115
         ScaleWidth      =   4395
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   0
         Width           =   4455
      End
      Begin VB.Label graphing 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HTTP Connections"
         Height          =   195
         Left            =   3360
         TabIndex        =   48
         Top             =   2190
         Width           =   1365
      End
      Begin VB.Label lb1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Peak Value:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   37
         ToolTipText     =   "Peak activity on current graphing"
         Top             =   2490
         WhatsThisHelpID =   10042
         Width           =   870
      End
      Begin VB.Label lb4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Longest Idle:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   45
         ToolTipText     =   "Longest idle for the current graphing"
         Top             =   2730
         WhatsThisHelpID =   10043
         Width           =   915
      End
      Begin VB.Label peak 
         AutoSize        =   -1  'True
         Caption         =   "E"
         Height          =   195
         Left            =   1440
         TabIndex        =   19
         ToolTipText     =   "Peak activity on current graphing"
         Top             =   2490
         WhatsThisHelpID =   10045
         Width           =   135
      End
      Begin VB.Label idle 
         AutoSize        =   -1  'True
         Caption         =   "E"
         Height          =   195
         Left            =   1440
         TabIndex        =   46
         ToolTipText     =   "Longest idle for the current graphing"
         Top             =   2730
         WhatsThisHelpID =   10046
         Width           =   135
      End
      Begin VB.Label lb2 
         AutoSize        =   -1  'True
         Caption         =   "Last Idle:"
         Height          =   195
         Left            =   2010
         TabIndex        =   38
         ToolTipText     =   "How long the talker has been idle"
         Top             =   2490
         WhatsThisHelpID =   10060
         Width           =   645
      End
      Begin VB.Label lb5 
         AutoSize        =   -1  'True
         Caption         =   "Heartbeat:"
         Height          =   195
         Left            =   2010
         TabIndex        =   43
         ToolTipText     =   "Should read 60"
         Top             =   2730
         WhatsThisHelpID =   10061
         Width           =   750
      End
      Begin VB.Label lastidle 
         AutoSize        =   -1  'True
         Caption         =   "E"
         Height          =   195
         Left            =   2790
         TabIndex        =   18
         ToolTipText     =   "How long the talker has been idle"
         Top             =   2490
         WhatsThisHelpID =   10062
         Width           =   135
      End
      Begin VB.Label heartbeat 
         AutoSize        =   -1  'True
         Caption         =   "E"
         Height          =   195
         Left            =   2790
         TabIndex        =   44
         ToolTipText     =   "Should read 60"
         Top             =   2730
         WhatsThisHelpID =   10063
         Width           =   135
      End
      Begin VB.Label lb3 
         AutoSize        =   -1  'True
         Caption         =   "Activity:"
         Height          =   195
         Left            =   3420
         TabIndex        =   40
         ToolTipText     =   "Server activity last minute"
         Top             =   2490
         WhatsThisHelpID =   10064
         Width           =   555
      End
      Begin VB.Label lb6 
         AutoSize        =   -1  'True
         Caption         =   "Act %:"
         Height          =   195
         Left            =   3420
         TabIndex        =   42
         ToolTipText     =   "Activity to idle ratio"
         Top             =   2730
         WhatsThisHelpID =   10065
         Width           =   450
      End
      Begin VB.Label actstat 
         AutoSize        =   -1  'True
         Caption         =   "E"
         Height          =   195
         Left            =   4020
         TabIndex        =   39
         ToolTipText     =   "Server activity last minute"
         Top             =   2490
         WhatsThisHelpID =   10066
         Width           =   135
      End
      Begin VB.Label actper 
         AutoSize        =   -1  'True
         Caption         =   "E"
         Height          =   195
         Left            =   4020
         TabIndex        =   41
         ToolTipText     =   "Activity to idle ratio"
         Top             =   2730
         WhatsThisHelpID =   10067
         Width           =   135
      End
      Begin VB.Label timespan 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Chart covers the past 0 hours"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   49
         Top             =   2190
         WhatsThisHelpID =   10041
         Width           =   2535
      End
      Begin VB.Line Liner1 
         BorderColor     =   &H80000014&
         X1              =   480
         X2              =   4710
         Y1              =   2445
         Y2              =   2445
      End
      Begin VB.Line Liner2 
         BorderColor     =   &H80000010&
         X1              =   480
         X2              =   4710
         Y1              =   2430
         Y2              =   2430
      End
      Begin VB.Label side 
         Alignment       =   1  'Right Justify
         Height          =   165
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Top             =   1980
         WhatsThisHelpID =   10047
         Width           =   255
      End
      Begin VB.Label side 
         Alignment       =   1  'Right Justify
         Height          =   165
         Index           =   10
         Left            =   0
         TabIndex        =   16
         Top             =   330
         WhatsThisHelpID =   10048
         Width           =   255
      End
      Begin VB.Label side 
         Alignment       =   1  'Right Justify
         Height          =   165
         Index           =   9
         Left            =   0
         TabIndex        =   15
         Top             =   495
         WhatsThisHelpID =   10049
         Width           =   255
      End
      Begin VB.Label side 
         Alignment       =   1  'Right Justify
         Height          =   165
         Index           =   8
         Left            =   0
         TabIndex        =   14
         Top             =   660
         WhatsThisHelpID =   10050
         Width           =   255
      End
      Begin VB.Label side 
         Alignment       =   1  'Right Justify
         Height          =   165
         Index           =   7
         Left            =   0
         TabIndex        =   13
         Top             =   825
         WhatsThisHelpID =   10051
         Width           =   255
      End
      Begin VB.Label side 
         Alignment       =   1  'Right Justify
         Height          =   165
         Index           =   6
         Left            =   0
         TabIndex        =   12
         Top             =   990
         WhatsThisHelpID =   10052
         Width           =   255
      End
      Begin VB.Label side 
         Alignment       =   1  'Right Justify
         Height          =   165
         Index           =   5
         Left            =   0
         TabIndex        =   11
         Top             =   1155
         WhatsThisHelpID =   10053
         Width           =   255
      End
      Begin VB.Label side 
         Alignment       =   1  'Right Justify
         Height          =   165
         Index           =   4
         Left            =   0
         TabIndex        =   10
         Top             =   1320
         WhatsThisHelpID =   10054
         Width           =   255
      End
      Begin VB.Label side 
         Alignment       =   1  'Right Justify
         Height          =   165
         Index           =   3
         Left            =   0
         TabIndex        =   9
         Top             =   1485
         WhatsThisHelpID =   10055
         Width           =   255
      End
      Begin VB.Label side 
         Alignment       =   1  'Right Justify
         Height          =   165
         Index           =   2
         Left            =   0
         TabIndex        =   8
         Top             =   1650
         WhatsThisHelpID =   10056
         Width           =   255
      End
      Begin VB.Label side 
         Alignment       =   1  'Right Justify
         Height          =   165
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   1815
         WhatsThisHelpID =   10057
         Width           =   255
      End
      Begin VB.Label side 
         Alignment       =   1  'Right Justify
         Height          =   165
         Index           =   12
         Left            =   0
         TabIndex        =   6
         Top             =   0
         WhatsThisHelpID =   10058
         Width           =   255
      End
      Begin VB.Label side 
         Alignment       =   1  'Right Justify
         Height          =   165
         Index           =   11
         Left            =   0
         TabIndex        =   5
         Top             =   165
         WhatsThisHelpID =   10059
         Width           =   255
      End
   End
   Begin VB.Timer light_time 
      Interval        =   15000
      Left            =   6030
      Top             =   3750
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   4680
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   441
            MinWidth        =   441
            Picture         =   "main.frx":221A
            Object.ToolTipText     =   "Recent Activity Indicator"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10365
            MinWidth        =   10365
            Object.ToolTipText     =   "Event Notifcation Bar"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   575
            MinWidth        =   575
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Users Logged On"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   575
            MinWidth        =   575
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Active Netlinks"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Shutdown_Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4770
      Top             =   3750
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   3930
      Top             =   3750
   End
   Begin VB.VScrollBar vscroll 
      Height          =   2737
      LargeChange     =   12
      Left            =   6645
      Max             =   0
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   690
      Width           =   195
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   2100
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3210
      Width           =   4515
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   2100
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3000
      Width           =   4515
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   2100
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2790
      Width           =   4515
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   2100
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2580
      Width           =   4515
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   2100
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2370
      Width           =   4515
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   2100
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2160
      Width           =   4515
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   2100
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1950
      Width           =   4515
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   2100
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1740
      Width           =   4515
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   2100
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1530
      Width           =   4515
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   2100
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1320
      Width           =   4515
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   2100
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1110
      Width           =   4515
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   2100
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   900
      Width           =   4515
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   2100
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   690
      Width           =   4515
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ToolbarImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "TOOLBAR REBOOT"
            Description     =   "SouthWest Tools"
            Object.ToolTipText     =   "Reboot Server"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   11
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "REBOOT SERVER"
                  Text            =   "Reboot Server"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bar291"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Full Reload"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bar292"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "R COMMANDS"
                  Text            =   "Reload Commands"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "R FIGLET FONTS"
                  Text            =   "Reload Figlet Fonts"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "R LOGIN SCREEN"
                  Text            =   "Reload Login Screen"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "R NETLINKS"
                  Text            =   "Reload Netlinks"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "R ROOMS"
                  Text            =   "Reload Rooms"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "R SETTINGS"
                  Text            =   "Reload Settings"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "R SWEARS"
                  Text            =   "Reload Swears"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "SYSTEM TRAY"
            Description     =   "System Tray"
            Object.ToolTipText     =   "System Tray"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "SW TOOLS"
            Description     =   "SouthWest Tools"
            Object.ToolTipText     =   "SouthWest Tools"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SW SCRIPT"
            Description     =   "SouthWest Script"
            Object.ToolTipText     =   "SouthWest Script"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BROADCAST"
            Description     =   "Broadcast"
            Object.ToolTipText     =   "Broadcast"
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin VB.TextBox bar 
         Height          =   285
         Left            =   -5000
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
         Width           =   4275
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6270
         Picture         =   "main.frx":2376
         ScaleHeight     =   210
         ScaleWidth      =   525
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   30
         Width           =   585
      End
   End
   Begin VB.ListBox connectionsList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   5160
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Connections List Box"
      Top             =   3450
      WhatsThisHelpID =   10022
      Width           =   1695
   End
   Begin VB.ListBox Syslog 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      ItemData        =   "main.frx":225A0
      Left            =   0
      List            =   "main.frx":225A2
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Running Systems Logbook"
      Top             =   3450
      WhatsThisHelpID =   10021
      Width           =   5145
   End
   Begin MSComctlLib.TreeView tree 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   5318
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      SingleSel       =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image StatImgHold 
      Height          =   240
      Index           =   1
      Left            =   60
      Picture         =   "main.frx":225A4
      Top             =   4710
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image StatImgHold 
      Height          =   105
      Index           =   0
      Left            =   90
      Top             =   4800
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2100
      TabIndex        =   2
      ToolTipText     =   "SouthWest Viewer"
      Top             =   450
      UseMnemonic     =   0   'False
      Width           =   75
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   8700
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   8160
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu LoggingFile 
         Caption         =   "&Logging"
         Checked         =   -1  'True
      End
      Begin VB.Menu ClearLogFile 
         Caption         =   "&Clear Log"
      End
      Begin VB.Menu bar328 
         Caption         =   "-"
      End
      Begin VB.Menu NotepadConfigFile 
         Caption         =   "&Edit Configuration Script"
         Shortcut        =   ^C
      End
      Begin VB.Menu bar10 
         Caption         =   "-"
      End
      Begin VB.Menu RebootTools 
         Caption         =   "&Reboot Server"
         Begin VB.Menu RebootReboot 
            Caption         =   "Reboot Server"
         End
         Begin VB.Menu bar222 
            Caption         =   "-"
         End
         Begin VB.Menu FullReMenu 
            Caption         =   "Full Reload"
         End
         Begin VB.Menu bar230 
            Caption         =   "-"
         End
         Begin VB.Menu Re 
            Caption         =   "Reload Commands"
            Index           =   0
         End
         Begin VB.Menu Re 
            Caption         =   "Reload Figlet Fonts"
            Index           =   1
         End
         Begin VB.Menu Re 
            Caption         =   "Reload Login Screen"
            Index           =   2
         End
         Begin VB.Menu Re 
            Caption         =   "Reload Netlinks"
            Index           =   3
         End
         Begin VB.Menu Re 
            Caption         =   "Reload Rooms"
            Index           =   4
         End
         Begin VB.Menu Re 
            Caption         =   "Reload Settings"
            Index           =   5
         End
         Begin VB.Menu Re 
            Caption         =   "Reload Swears"
            Index           =   6
         End
      End
      Begin VB.Menu ShutdownTools 
         Caption         =   "&Shutdown Server"
      End
      Begin VB.Menu bar3 
         Caption         =   "-"
      End
      Begin VB.Menu UserfileConversionsFile 
         Caption         =   "&Userfile Conversions"
         Enabled         =   0   'False
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu ExitFile 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu MiscMenu 
      Caption         =   "&Other"
      Begin VB.Menu PluginsMenu 
         Caption         =   "&Plugins"
         Begin VB.Menu PluginPageTools 
            Caption         =   "&Find && Download Plugins..."
         End
         Begin VB.Menu bar2384 
            Caption         =   "-"
         End
         Begin VB.Menu Plugin 
            Caption         =   "(None)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu FigletsOther 
         Caption         =   "&Figlet Fonts"
         Begin VB.Menu FigletList 
            Caption         =   "(None)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu bar287 
         Caption         =   "-"
      End
      Begin VB.Menu SouthWestToolsTools 
         Caption         =   "&SouthWest Tools"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu WindowMenu 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu RefreshWindow 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu barwin1 
         Caption         =   "-"
      End
      Begin VB.Menu SystemTrayWindow 
         Caption         =   "&Minimize to System Tray"
      End
      Begin VB.Menu TopWindow 
         Caption         =   "Always On &Top"
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu ContentsHelp 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu bar203 
         Caption         =   "-"
      End
      Begin VB.Menu HelpSWOnline 
         Caption         =   "SouthWest Homepage"
      End
      Begin VB.Menu HelpSWOnlineHelp 
         Caption         =   "SouthWest Online Help"
      End
      Begin VB.Menu bar236 
         Caption         =   "-"
      End
      Begin VB.Menu LicenceHelp 
         Caption         =   "&License Info..."
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu AboutHelp 
         Caption         =   "&About SouthWest..."
      End
   End
   Begin VB.Menu mnuSystemtray 
      Caption         =   "mnuSystemtray"
      Visible         =   0   'False
      Begin VB.Menu OpenSW_System 
         Caption         =   "Open SouthWest"
      End
      Begin VB.Menu bar9 
         Caption         =   "-"
      End
      Begin VB.Menu Shutdown_System 
         Caption         =   "Shutdown"
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "mnuUser"
      Visible         =   0   'False
      Begin VB.Menu UserPopInfo 
         Caption         =   "Info"
      End
      Begin VB.Menu UserPopKill 
         Caption         =   "Kill"
      End
      Begin VB.Menu UserPopPromote 
         Caption         =   "Promote"
      End
      Begin VB.Menu UserPopDemote 
         Caption         =   "Demote"
      End
      Begin VB.Menu mnuICQActive 
         Caption         =   "ICQ"
      End
      Begin VB.Menu bar8 
         Caption         =   "-"
      End
      Begin VB.Menu UserPopPickle 
         Caption         =   "Pickle"
      End
   End
   Begin VB.Menu mnuTreeUser 
      Caption         =   "mnuTreeUser"
      Visible         =   0   'False
      Begin VB.Menu TUInfo 
         Caption         =   "User Information"
      End
      Begin VB.Menu mnuOfflineICQ 
         Caption         =   "ICQ Message"
      End
      Begin VB.Menu bar238 
         Caption         =   "-"
      End
      Begin VB.Menu TUEraseAccount 
         Caption         =   "Erase Account"
      End
   End
   Begin VB.Menu mnuTreeNetlink 
      Caption         =   "mnuTreeNetlink"
      Visible         =   0   'False
      Begin VB.Menu TNConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu TNrefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim islighton As Boolean, preGrapherInit As Boolean
Dim captured_keys As String
Public dumpCheck As Boolean

Private Sub loadPlugins()
'This loadup routine will load plugins. Plugins can have several types
'of hooks and not all plugins are listed in the menu.
Dim linein As LOAD_OBJECT, FromFile As String, file As String
Dim parse As Boolean, count As Integer, found As Boolean
Dim currentPlug As Integer, usedMenu As Boolean, mnuPos As Integer
If BOOTING Then
    lighter "Loading plugins"
    End If
file = Dir$(App.Path & "\Plugins\*.DPD")
Do While Not file = ""
    For count = LBound(plugs) To UBound(plugs)
        If Not plugs(count).inuse Then
            plugs(count).inuse = True
            currentPlug = count
            Exit For
            End If
        Next count
    Open App.Path & "\Plugins\" & file For Input As #1
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
                Case "name"
                    plugs(currentPlug).name = Trim$(linein.value)
                Case "version"
                    plugs(currentPlug).version = Trim$(linein.value)
                Case "hooks"
                    plugs(currentPlug).hooks = LCase$(linein.value)
                Case "pointer"
                    plugs(currentPlug).exeptr = linein.value
                    End Select
            End If
        Loop
    Close #1
    If InStr(plugs(currentPlug).hooks, "menu") Then
        If usedMenu Then
            Load Plugin(Plugin.UBound + 1)
            mnuPos = Plugin.UBound
            Else
                mnuPos = 0
                usedMenu = True
                End If
        plugs(currentPlug).menuPos = mnuPos
        Plugin(mnuPos).Enabled = True
        Plugin(mnuPos).Caption = plugs(currentPlug).name
        End If
    file = Dir$
    Loop
End Sub

Public Sub updateActiveNetlinks()
Dim count As Integer, maxNetlinks As Integer
For count = LBound(net) To UBound(net)
    If net(count).state = NETLINK_UP Then
        maxNetlinks = maxNetlinks + 1
        End If
    Next count
mainForm.StatusBar.Panels(4).text = Trim$(maxNetlinks)
End Sub

Private Sub autoPatchCheck()
Dim file As String, found As Boolean
file = Dir$(App.Path & "\AutoPatch\*.AP")
Do While Len(file) > 3
    If Not found Then
        writeSyslog "Installing patch/plugin files "
        End If
    autoPatchFile = "\AutoPatch\" & file
    If Not found Then
        found = True
        patcherForm.visible = True
        Me.Enabled = False
        Else
            apEngine
            End If
    Kill App.Path & "\AutoPatch\" & file
    'we reinitialize the dir statement because AP will use it
    'and we delete used patches anyways
    file = Dir$(App.Path & "\AutoPatch\*.AP")
    Loop
If (InStr(Command$, "-systray") = 0) And (system.rebooted = False) Then
    Me.show
    End If
Unload patcherForm
Me.Enabled = True
Me.SetFocus
End Sub

Private Sub SetRadioMenuChecks(Mnu As Menu, ByVal mnuItem As Long)
   Dim hMenu As Long
   Dim mInfo As MENUITEMINFO
  'get the menuitem handle
   hMenu& = GetSubMenu(GetSubMenu(GetMenu(Mnu.Parent.hwnd), 1), 1)
  'copy its attributes to the new Type,
  'changing the checkmark to a radiobutton
   With mInfo
     .cbSize = Len(mInfo)
     .fType = MFT_RADIOCHECK
     .fMask = MIIM_TYPE
     .dwTypeData = Mnu.Caption & Chr$(0)
   End With
  'change the menu check mark
   SetMenuItemInfo hMenu&, mnuItem&, 1, mInfo
End Sub

Sub boldPopup(ByRef Mnu As Menu, move_left As Integer, move_down As Integer)
Mnu.Caption = vbNullString
Mnu.visible = True
Call SetMenuDefaultItem(GetSubMenu(GetMenu(Mnu.Parent.hwnd), move_left), move_down, True)
PopupMenu Mnu
Mnu.visible = False
End Sub

Sub loadFigletMenu()
Dim file As String, count As Integer, found As Boolean
For count = FigletList.LBound To FigletList.UBound
    If count > 0 Then
        Unload FigletList(count)
        End If
    Next count
FigletList(0).Caption = "(None)"
FigletList(0).Checked = False
FigletList(0).Enabled = False
file = Dir$(App.Path & "\Figlets\*.F")
If Len(file) > 2 Then
    file = Left$(file, Len(file) - 2)
    Else
        file = ""
        End If
count = 0
Do While Not file = ""
    If UCase$(file) = UCase$(system.figlet) Then
        found = True
        End If
    If count > 0 Then
        Load FigletList(count)
        End If
    FigletList(count).Enabled = True
    FigletList(count).Caption = file
    FigletList(count).Checked = True
    If UCase$(file) = UCase$(system.figlet) Then
        FigletList(count).Checked = True
        Else
            FigletList(count).Checked = False
            End If
    file = Dir$()
    If Len(file) > 2 Then
        file = Left$(file, Len(file) - 2)
        Else
            file = ""
            End If
    SetRadioMenuChecks FigletList(count), count
    count = count + 1
    Loop
If Not found Then
    system.figlet = ""
    loadFiglets
    End If
End Sub

Sub hideGrapher()
actup = False
mainForm.action_frm.visible = False
'Because the AutoGrapher is shown at startup, we need a special
'check to see if we should compress the branch
If Not tree.SelectedItem.Key = "GRAPHER" Then
    tree.Nodes("GRAPHER").Expanded = False
    End If
End Sub

Sub showGrapher()
actup = True
mainForm.action_frm.visible = True
End Sub

Private Sub ClearLogFile_Click()
Dim count As Integer
For count = LBound(logbook) To UBound(logbook)
    logbook(count) = ""
    Next count
logpos = 0
lighter "System logbook erased"
If mainForm.tree.SelectedItem.Key = "SYSLOG" Then
    mainForm.treeLoad
    End If
End Sub

Private Sub connectionsList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    If Not connectionsList.ListIndex = -1 Then
        setOnlinePopup
        End If
    End If
End Sub

Private Sub HelpSWOnline_Click()
Call ShellExecute(Me.hwnd, vbNullString, "http://talker.com/southwest/", vbNullString, "c:\", 1) 'SW_SHOWNORAL
End Sub

Private Sub HelpSWOnlineHelp_Click()
Call ShellExecute(Me.hwnd, vbNullString, "http://talker.com/southwest/help.htm", vbNullString, "c:\", 1) 'SW_SHOWNORAL
End Sub

Private Sub mnuICQActive_Click()
Dim count As Integer
For count = 1 To UBound(user)
    If user(count).listing = mainForm.connectionsList.ListIndex Then
        If ICQSendMessage(CLng(user(count).ICQ), "") = 1 Then
            lighter "Sending an ICQ message to " & userCap(user(count).name)
            Else
                lighter "Can't send message. ICQ is not available."
                End If
        Exit For
        End If
    Next count
End Sub

Private Sub mnuOfflineICQ_Click()
user(0).name = tree.SelectedItem.text
loadUserData 0
If ICQSendMessage(CLng(user(0).ICQ), "") = 1 Then
    lighter "Sending an ICQ message to " & userCap(user(0).name)
    Else
        lighter "Can't send message. ICQ is not available."
        End If
End Sub

Private Sub OpenSW_System_Click()
Me.WindowState = 0
Me.visible = True
Me.show
Me.SetFocus
Call Shell_NotifyIcon(NIM_DELETE, VBGTray)
End Sub

Private Sub SouthWestToolsTools_Click()
Shell App.Path & "\Plugins\Tools\tools.exe", vbNormalFocus
End Sub

Private Sub FigletList_Click(Index As Integer)
Dim count As Integer
For count = FigletList.LBound To FigletList.UBound
    FigletList(count).Checked = False
    Next count
FigletList(Index).Checked = True
system.figlet = FigletList(Index).Caption
loadFiglets
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
captured_keys = captured_keys & Chr$(KeyAscii)
If Len(captured_keys) > 6 Then
    captured_keys = Right$(captured_keys, 5)
    End If
If UCase$(captured_keys) = "DEBUG" Then
    mainForm.Timer2.Enabled = False
    mainForm.Timer2.Interval = 250
    mainForm.Timer2.Enabled = True
    captured_keys = ""
    loadGrapher
    End If
End Sub

Public Sub loadGrapher()
Dim gf As Double
If whatToGraph = graph.GRAPH_NULL Then
    whatToGraph = graph.GRAPH_ACTIONS
    End If
Select Case whatToGraph
    Case graph.GRAPH_ACTIONS
        ReDim grapher(LBound(acts) To UBound(acts))
        grapher = acts
        graphing.Caption = "User Activity"
    Case graph.GRAPH_USER_LOGINS
        ReDim grapher(LBound(logins_graph) To UBound(logins_graph))
        grapher = logins_graph
        graphing.Caption = "User Logins"
    Case graph.GRAPH_HTTP_CONNECTIONS
        ReDim grapher(LBound(httpConnects) To UBound(httpConnects))
        grapher = httpConnects
        graphing.Caption = "HTTP Connections"
    Case graph.GRAPH_HTTP_REQUESTS
        ReDim grapher(LBound(http_acts) To UBound(http_acts))
        grapher = http_acts
        graphing.Caption = "HTTP Requests"
    End Select
actup = True
preGrapherInit = False
heartbeat.Caption = Str$(Int(mainForm.Timer2.Interval / 1000))
If heartbeat.Caption = " 0" Then
    heartbeat.Caption = " < 1"
    End If
timespan.Caption = "Chart covers the past "
gf = UBound(grapher) / 60 * (mainForm.Timer2.Interval / 60000)
If Int(gf) = gf Then
    timespan.Caption = timespan.Caption & Format$(gf, "###") & " hour"
    Else
        timespan.Caption = timespan.Caption & Format$(gf, "###.##") & " hour"
        End If
If gf > 1 Then
    timespan.Caption = timespan.Caption & "s"
    End If
g_Paint
End Sub

Sub lightbulbOn(switch As Boolean)
If Not islighton = switch Then
    StatImgHold(0).picture = StatusBar.Panels(1).picture
    StatusBar.Panels(1).picture = StatImgHold(1).picture
    StatImgHold(1).picture = StatImgHold(0).picture
    islighton = switch
    End If
End Sub

Sub setOnlinePopup()
'This will configure the popup menu for online users
'Since we dont want to be able to promote people in the login stage,
'or do much of anything else to them for that matter, we must
'block out certain selections
Dim loggedin As Boolean, count As Integer, validICQ As Boolean
loggedin = False
For count = 1 To UBound(user)
    If user(count).listing = mainForm.connectionsList.ListIndex Then
        loggedin = userIsOnline(user(count).name)
        Exit For
        End If
    Next count
validICQ = IIf(Val(user(count).ICQ) > 0, True, False)
validICQ = validICQ And system.icqHook
mnuICQActive.Enabled = loggedin And validICQ
UserPopInfo.Enabled = loggedin
UserPopPromote.Enabled = loggedin And user(count).rank < UBound(ranks)
UserPopDemote.Enabled = loggedin And user(count).rank > LBound(ranks)
UserPopPickle.Enabled = loggedin
boldPopup mnuUser, 4, 6
End Sub

Sub treeLoad()
Dim main As String, Top As String, FromFile As String, temp As String
Dim tmp As Integer, count As Integer, datehold As Date, un As Integer
Dim t As DHMS_OBJECT, dateStr As String, hi As Integer
If Not tree.SelectedItem.FullPath = tree.SelectedItem.text Then
If Not InStr(tree.SelectedItem.FullPath, tree.PathSeparator) = Len(tree.SelectedItem.FullPath) Then
main = Left$(tree.SelectedItem.FullPath, InStr(tree.SelectedItem.FullPath, tree.PathSeparator) - 1)
Top = Right$(tree.SelectedItem.FullPath, Len(tree.SelectedItem.FullPath) - Len(main) - 1)
If Not tree.SelectedItem.FullPath = tree.SelectedItem.text Then
    Select Case UCase$(main)
        Case "AUTOGRAPH"
            Select Case tree.SelectedItem.Key
                Case "USER ACTIVITY"
                    whatToGraph = graph.GRAPH_ACTIONS
                Case "USER LOGINS"
                    whatToGraph = graph.GRAPH_USER_LOGINS
                Case "HTTP CONNECTIONS"
                    whatToGraph = graph.GRAPH_HTTP_CONNECTIONS
                Case "HTTP REQUESTS"
                    whatToGraph = graph.GRAPH_HTTP_REQUESTS
                End Select
            showGrapher
            loadGrapher
        Case "USER DATA"
        If Not Dir(App.Path & "\Users\" & Top & ".D") = "" Then
            Label1.Caption = Top & "'s User Data"
            If Not userIsOnline(Top) Then
                user(0).name = Top
                loadUserData 0
                datehold = num2date(user(0).lastLogin)
temp = "   ~OLDesc:~RS " & user(0).desc & CRLF & _
       "   ~OLRank: ~RS" & ranks(user(0).rank) & CRLF & _
       "   ~OLTotal Login: ~RS" & deriveTimeString(spliceTime(user(0).totalTime), False) & CRLF & _
       "   ~OLLast Login: ~RS" & Format$(datehold, "dddd d") & getOrdinal(Int(Format$(datehold, "d"))) & Format$(datehold, " mmmm yyyy" & " at " & Format$(datehold, "hh:nn")) & CRLF & _
       "   ~OLWhich Was: ~RS" & deriveTimeString(spliceTime(date2num(Now) - user(0).lastLogin)) & CRLF & _
       "   ~OLWas On For: ~RS" & deriveTimeString(spliceTime(user(0).timeon), False) & CRLF & _
       "   ~OLEnter Message: ~RS" & user(0).enterMsg & CRLF & _
       "   ~OLExit Message: ~RS" & user(0).exitMsg & CRLF & _
       "   ~OLLast Site: ~RS" & user(0).site & CRLF & _
       "   ~OLEmail: ~RS" & user(0).email & CRLF & _
       "   ~OLICQ: ~RS" & user(0).ICQ & CRLF & _
       "   ~OLLogins: ~RS" & user(0).logins & CRLF & _
       "   ~OLGender: ~RS" & user(0).gender & CRLF & _
       "   ~OLAge: ~RS" & IIf(Val(user(0).age) > 0, user(0).age, "Unset") & CRLF & _
       "   ~OLArrested: ~RS" & bool2YN(user(0).arrested) & CRLF & _
       "   ~OLVisible: ~RS" & bool2YN(user(0).visible) & CRLF & _
       "   ~OLMuzzled: ~RS" & bool2YN(user(0).muzzled) & CRLF & _
       "   ~OLExpires: ~RS" & bool2YN(user(0).expires) & CRLF & _
       "   ~OLEmail Fwd: ~RS" & bool2YN(user(0).sfRec)
                Else
                    un = getUser(Top)
temp = "   ~OLDesc:~RS " & user(un).desc & CRLF & _
       "   ~OLRank: ~RS" & ranks(user(un).rank) & CRLF & _
       "   ~OLTotal Login: ~RS" & deriveTimeString(spliceTime(user(un).totalTime), False) & CRLF & _
       "   ~OLOn For: ~RS" & deriveTimeString(spliceTime(user(un).timeon), False) & CRLF & _
       "   ~OLIdle For: ~RS" & deriveTimeString(spliceTime(user(un).idle * 60), False) & CRLF & _
       "   ~OLRoom: ~RS" & user(un).room & CRLF
If user(un).atNetlink >= 0 Then
    temp = temp & "   ~OLAt Netlink: ~RS" & net(user(un).atNetlink).name & CRLF
    End If
If user(un).netlinkType Then
    temp = temp & "   ~OLFrom Netlink: ~RS" & net(user(un).netlinkFrom).name & CRLF
    End If
temp = temp & "   ~OLEnter Message: ~RS" & user(un).enterMsg & CRLF & _
       "   ~OLExit Message: ~RS" & user(un).exitMsg & CRLF & _
       "   ~OLSite: ~RS" & user(un).site & CRLF & _
       "   ~OLEmail: ~RS" & user(un).email & CRLF & _
       "   ~OLICQ: ~RS" & user(un).ICQ & CRLF & _
       "   ~OLLogins: ~RS" & user(un).logins & CRLF & _
       "   ~OLGender: ~RS" & user(un).gender & CRLF & _
       "   ~OLAge: ~RS" & IIf(Val(user(0).age) > 0, user(0).age, "Unset") & CRLF & _
       "   ~OLArrested: ~RS" & bool2YN(user(un).arrested) & CRLF & _
       "   ~OLVisible: ~RS" & bool2YN(user(un).visible) & CRLF & _
       "   ~OLMuzzled: ~RS" & bool2YN(user(un).muzzled) & CRLF & _
       "   ~OLExpires: ~RS" & bool2YN(user(un).expires) & CRLF & _
       "   ~OLEmail Fwd: ~RS" & bool2YN(user(un).sfRec)
                    End If
            load_text temp
            End If
        Case "USER HISTORIES"
        If Not Dir$(App.Path & "\Users\" & Top & ".His") = "" Then
            Dim A As String
            Open App.Path & "\Users\" & Top & ".His" For Input As #1
            Do While Not EOF(1)
                Line Input #1, FromFile
                temp = temp & FromFile
                If Not EOF(1) Then
                    temp = temp & CRLF
                    End If
                Loop
            Label1.Caption = Top & "'s History File"
            load_text temp
            Close #1
            End If
        Case "NETLINKS"
            tmp = -1
            For count = LBound(net) To UBound(net)
                If net(count).name = mainForm.tree.SelectedItem.text Then
                    tmp = count
                    Exit For
                    End If
                Next count
            If tmp = -1 Then
                loadViewer
                Exit Sub
                End If
            Label1.Caption = net(tmp).name & " Netlink Information"
            temp = "   ~OLName: ~RS" & net(tmp).name & CRLF & "~OLStatus: ~RS" & _
                netlinkStates(net(tmp).state) & CRLF & "~OLHost: ~RS" & _
                net(tmp).site & CRLF & "~OLPort: ~RS" & net(tmp).port & CRLF & _
                "   ~OLOutgoing Access: ~RS" & BoolYN(net(tmp).allowOut) & _
                CRLF & "   ~OLIncoming Access: ~RS" & BoolYN(net(tmp).allowIn) & _
                CRLF & "   ~OLAutomaticly Connect: ~RS" & BoolYN(net(tmp).autoConnect) & _
                CRLF & "   ~OLBytes In: ~RS" & net(tmp).bytesIn & CRLF & _
                "   ~OLBytes Out: ~RS" & net(tmp).bytesOut
                load_text temp
        End Select
    End If
End If
Else
    temp = ""
    Select Case UCase$(tree.SelectedItem.Key)
        Case "GRAPHER"
            Label1.Caption = "SouthWest AutoGraph"
            temp = "   SouthWest AutoGraph can draw visual representations of some" & CRLF & _
            "   talker data. SouthWest can graph ~FGUser Activity~RS, which is the" & CRLF & _
            "   amount of command and speech actions that are executed by" & CRLF & _
            "   users, ~FGUser Logins~RS, which takes into account both Netlink" & CRLF & _
            "   and directly connected users, ~FGHTTP Connections~RS, which is" & CRLF & _
            "   the number of ~OLunique~RS machine connections, and its sister" & CRLF & _
            "   function, ~FGHTTP Requests~RS, which details the amount of pages." & CRLF & _
            "   that have been served."
        Case "USER DATA"
            Label1.Caption = "User Data Viewer"
            temp = "   To retrieve statistics and information about a user, select a user" & CRLF & _
            "   from the list below. Information similar to the kind displayed" & CRLF & _
            "   when the ~FG.examine~RS command is executed will be shown in the" & CRLF & _
            "   viewer window."
        Case "USER HISTORIES"
            Label1.Caption = "User History Viewer"
            temp = "   Some commands log their usage in a user's profile. The Viewer" & CRLF & _
            "   can be used to review these files. After selecting a user by" & CRLF & _
            "   name in the tree, the user's history will be loaded." & CRLF & CRLF & _
            "   To see at what date and time a particular event occurred, move" & CRLF & _
            "   the mouse cursor over the event for which you want the date" & CRLF & _
            "   and time and it will be displayed in a tip box."
        Case "NETLINKS"
            Label1.Caption = "Netlink Monitor"
            temp = "   To quickly view the status of any registered Netlinks, click" & CRLF & _
            "   on any of the members of the Netlink tree entry. The viewer" & CRLF & _
            "   will give you any available statistics for the selected Netlink. To" & CRLF & _
            "   connect, disconnect, or perform any other operations on the" & CRLF & _
            "   Netlink, right-click on the desired Netlink and use the popup" & CRLF & _
            "   context menu."
        Case "SERVER_INFO"
            'This is done for us in the draw event but since this may
            'take longer to load some info, wipe it first
            For count = 0 To 12
                mainForm.p(count).ToolTipText = ""
                mainForm.p(count).Cls
                Next count
            mainForm.MousePointer = vbHourglass
            mainForm.p(0).Print "Loading..."
            DoEvents
            Label1.Caption = "Server Information"
            temp = "   ~OLName:~RS " & system.talkerName & CRLF & "~OLHost:~RS " & _
            Socket1.LocalName & CRLF & "   ~OLIP Address:~RS " & Socket1.LocalAddress & CRLF & _
            "   ~OLSouthWest Version:~RS " & App.Major & "." & App.Minor & "." & App.Revision & CRLF & _
            "   ~OLMain Port: ~RS" & Socket1.LocalPort & CRLF & "   ~OLNetlink Port: ~RS" & _
            Netlink(0).LocalPort & CRLF & "   ~OLHTTP Port: ~RS" & http(0).LocalPort & CRLF & _
            "   ~OLSMTP Server: ~RS" & IIf(system.smtpServer = vbNullString, "~FRNot Found~RS", system.smtpServer)
            'We need to do this again because it may have changed
            mainForm.MousePointer = vbDefault
            Label1.Caption = "Server Information"
        Case "SYSLOG"
            Label1.Caption = "System's Logbook Reader"
            For count = logpos To UBound(logbook)
                If Not logbook(count) = "" Then
                    temp = temp & logbook(count) & CRLF
                    End If
                Next count
            For count = LBound(logbook) To logpos
                If count < logpos And Not logbook(count) = "" Then
                    temp = temp & logbook(count) & CRLF
                    End If
                Next count
            If Len(temp) > 2 Then
                temp = Left$(temp, Len(temp) - 2)
                Else
                    temp = "   The system logbook is empty."
                    End If
        End Select
    load_text temp
    End If
End Sub

Sub loadSockets()
Dim count As Integer
lighter "Loading sockets"
'To make the algorythms easier we will never use this.
madeSockets(0) = 1
madeNetlinks(0) = True

'Gets your ip address
If system.siteAtStartup Then
    If Socket1.LocalAddress = "" Then   'special consideration that windows 95
    'gives to it's Then
        writeSyslog "~FRWARNING: Unable to retrieve local ip address"
        Else
            writeSyslog "My IP address is " & Socket1.LocalAddress & "~FB [~RS" & Socket1.LocalName & "~FB]"
            End If
    End If
'Hey, dont look at me like that. I am well aware of the
'fact that SocketWrench has the LastError event but there
'arose a little bug in the 32-bit versions of the control
'that caused the LastError not to be fired at all in the
'IDE which really sucks for development and I am sure that
'there would be some other bugs there as well. So we will
'do it this way. :)
On Error Resume Next
Socket1.AddressFamily = AF_INET
Socket1.Protocol = IPPROTO_TCP
Socket1.SocketType = SOCK_STREAM
Socket1.Blocking = False
Socket1.LocalPort = system.mainPort
Socket1.Action = SOCKET_LISTEN
writeSyslog "Main socket initialized and listening on port ~FG" & Trim$(Socket1.LocalPort)
If Not Err = 0 Then
    MsgBox "The main socket is in use by another application or instance of SouthWest. Configure the sockets correctly using SouthWest Tools.", 16, "SouthWest - Error"
    Socket1.Action = SOCKET_ABORT
    End
    Exit Sub
    End If

Netlink(0).AddressFamily = AF_INET
Netlink(0).Protocol = IPPROTO_TCP
Netlink(0).SocketType = SOCK_STREAM
Netlink(0).Blocking = False
Netlink(0).LocalPort = system.netlinkPort
Netlink(0).Action = SOCKET_LISTEN
If Not Err = 0 Then
    MsgBox "The netlink socket is in use by another application or instance of SouthWest. Configure the sockets correctly using SouthWest Tools.", 16, "SouthWest - Error"
    Socket1.Action = SOCKET_ABORT
    End
    Exit Sub
    End If
writeSyslog "Netlink socket initialized and listening on port ~FG" & Trim$(system.netlinkPort)

For count = 1 To 5
    Load http(count)
    Next count
http(0).AddressFamily = AF_INET
http(0).Protocol = IPPROTO_TCP
http(0).SocketType = SOCK_STREAM
http(0).Blocking = False
http(0).LocalPort = system.httpPort
http(0).Binary = True
http(0).Listen
writeSyslog "HTTP server initialized and listening on port ~FG" & http(0).LocalPort

'This will set up the auxillary socket that is used for when things just
'dont go right.
Socket3.AddressFamily = AF_INET
Socket3.Protocol = IPPROTO_TCP
Socket3.SocketType = SOCK_STREAM
Socket3.Binary = True
Socket3.BufferSize = 1024
Socket3.Blocking = False

'To make it easier to stop processes we will put the port in the proccesses
'list box (Pressing Ctrl-Alt-Del in 95/NT/98) and also put it on the caption.
App.Title = "SouthWest - " & Socket1.LocalPort
mainForm.Caption = "SouthWest - " & Socket1.LocalPort
VBGTray.szTip = "SouthWest " & Str$(system.mainPort) & vbNullChar
End Sub

Private Sub loadMotd()
'Lets get the MOTD (Message of the Day) that will be
'displayed when someone first logs on.
lighter "Loading logon screen"
MOTD = ""
If Not Dir$(App.Path & "\Misc\Login Screen.S") = "" Then
    Dim linemotd As String
    Open App.Path & "\Misc\Login Screen.S" For Input As #1
    Do While Not EOF(1)
        DoEvents
        Line Input #1, linemotd
        MOTD = MOTD & linemotd & CRLF
        Loop
    Close #1
    Else
        MOTD = "There is no MOTD." & CRLF & CRLF
        End If
'pre-parse colors because of the buffered send
MOTD = parseColors(MOTD)
End Sub

Public Sub loadSystem()
Dim linein As LOAD_OBJECT, FromFile As String
Dim parse As Boolean, count As Integer, found As Boolean
lighter "Loading system configuration script"
Open App.Path & "\southwest.s" For Input As #1
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
            Case "main_port"
                system.mainPort = Val(linein.value)
                If Not BOOTING Then
                    changePortBinding mainForm.Socket1, system.mainPort
                    End If
            Case "netlink_port"
                system.netlinkPort = Val(linein.value)
                If Not BOOTING Then
                    changePortBinding mainForm.Netlink(0), system.netlinkPort
                    End If
            Case "http_port"
                system.httpPort = Val(linein.value)
                If Not BOOTING Then
                    changePortBinding mainForm.http(0), system.httpPort
                    End If
            Case "talker_name"
                system.talkerName = linein.value
            Case "email_address"
                system.emailAddress = linein.value
            Case "maxUsers"
                maxUsers = Val(linein.value)
            Case "smtp_server"
                system.smtpServer = linein.value
            Case "allow_gatecrashing"
                system.gatecrash = TF(linein.value)
            Case "purge_period"
                system.purgeLength = Int(linein.value)
            Case "idle_user_timeout"
                system.maxIdle = Int(linein.value)
            Case "autoconnect_netlinks"
                system.autoConnect = TF(linein.value)
            Case "gatecrash_level"
                For count = LBound(ranks) To UBound(ranks)
                    If UCase$(ranks(count)) = UCase$(linein.value) Then
                        found = True
                        Exit For
                       End If
                    Next count
                If found Then
                    system.gatecrashLevel = count
                    Else
                        writeSyslog "~FRBad rank given for gatecrash level"
                        system.gatecrashLevel = UBound(ranks)
                        End If
            Case "swear_ban"
                Select Case linein.value
                    Case "MIN"
                        system.swearing = SWEAR_MIN
                    Case "MAX"
                        system.swearing = SWEAR_MAX
                    Case Else
                        system.swearing = SWEAR_OFF
                        End Select
            Case "timeout_max_level"
                found = False
                For count = LBound(ranks) To UBound(ranks)
                    If UCase(ranks(count)) = UCase(linein.value) Then
                        found = True
                        Exit For
                        End If
                    Next count
                If found Then
                    system.timeoutMaxLevel = count
                    Else
                        'Noone is above the timeout if not specified
                        system.timeoutMaxLevel = UBound(ranks) + 1
                        End If
            Case "max_timeout_level"
                For count = LBound(ranks) To UBound(ranks)
                    If UCase$(ranks(count)) = UCase$(linein.value) Then
                        found = True
                        Exit For
                       End If
                    Next count
                If found Then
                    system.timeoutMaxLevel = count
                    Else
                        writeSyslog "~FRBad rank given for max timeout level"
                        system.timeoutMaxLevel = UBound(ranks)
                        End If
            Case "no_splashscreen"
                If Not noSplash Then
                    noSplash = TF(linein.value)
                    End If
            Case "site_at_startup"
                system.siteAtStartup = TF(linein.value)
            Case "figlet"
                system.figlet = linein.value
                End Select
        End If
    Loop
Close #1
End Sub

Public Sub viewerScrollRefresh()
Dim count As Integer, count2 As Integer
If UBound(lines) <= 13 Then
    mainForm.vscroll.Max = 0
    Else
        mainForm.vscroll.Max = UBound(lines) - 13
        End If
count2 = -1
For count = mainForm.vscroll.value To UBound(lines)
    count2 = count2 + 1
    Plot_Viewer lines(count), count2
    If count2 >= 12 Then
        Exit For
        End If
    Next count
End Sub

Private Sub GoSystemTray()
  VBGTray.cbSize = Len(VBGTray)
  VBGTray.hwnd = Me.hwnd
  VBGTray.uId = vbNull
  VBGTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  VBGTray.ucallbackMessage = WM_MOUSEMOVE
  VBGTray.hIcon = Me.Icon
  'tool tip text
  VBGTray.szTip = "SouthWest - " & system.talkerName & vbNullChar
  Call Shell_NotifyIcon(NIM_ADD, VBGTray)
  App.TaskVisible = False   'remove application from taskbar
  Me.Hide
End Sub

Public Sub keyScroll(KeyAscii As Integer)
'38 small up,     33 big up
'40 small down,   34 big down
Select Case KeyAscii
    Case 38
        If vscroll.value - vscroll.SmallChange >= vscroll.Min Then
            vscroll.value = vscroll.value - 1
            End If
    Case 40
        If vscroll.value + vscroll.SmallChange <= vscroll.Max Then
            vscroll.value = vscroll.value + vscroll.SmallChange
            End If
    Case 34
        If vscroll.value + vscroll.LargeChange <= vscroll.Max Then
            vscroll.value = vscroll.value + vscroll.LargeChange
            Else
                vscroll.value = vscroll.Max
                End If
    Case 33
        If vscroll.value - vscroll.LargeChange >= vscroll.Min Then
            vscroll.value = vscroll.value - vscroll.LargeChange
            Else
                vscroll.value = vscroll.Min
                End If
        End Select
End Sub

Private Sub bar_GotFocus()
bar.Left = 1950
bar.Top = 30
End Sub

Private Sub bar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not bar.text = "" Then
        writeRoom "", BELL & "~OL~FW~BR BROADCAST FROM SERVER: " & bar.text & " ~RS" & CRLF
        End If
    bar.visible = False
    lighter "Broadcasting: " & doubleQuote & bar.text & doubleQuote
    bar.text = ""
    Toolbar.Buttons("BROADCAST").value = tbrUnpressed
    Toolbar.Buttons("BROADCAST").Tag = "Closing"
    p(0).SetFocus
    End If
End Sub

Private Sub bar_LostFocus()
bar.visible = False
Toolbar.Buttons("BROADCAST").value = tbrUnpressed
If Not Toolbar.Buttons("BROADCAST").Tag = "Closing" Then
    lighter "Broadcast canceled"
    End If
Toolbar.Buttons("BROADCAST").Tag = ""
End Sub

Private Sub ExitFile_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Static lngMsg As Long
  Static blnFlag As Boolean
  Dim result As Long

lngMsg = x / Screen.TwipsPerPixelX
  If blnFlag = False Then
        blnFlag = True
        Select Case lngMsg
        'doubleclick
        Case WM_LBUTTONDBLCLICK
            Me.WindowState = 0
            Me.visible = True
            Me.show
            Me.SetFocus
            Call Shell_NotifyIcon(NIM_DELETE, VBGTray)
        'right-click
        Case WM_RBUTTONUP
          result = SetForegroundWindow(Me.hwnd)
          'Popup a menu with bold first item
          boldPopup mnuSystemtray, 4, 0
        End Select
        blnFlag = False
 End If
End Sub

Sub parseArguments()
Dim count As Integer
lighter "Parsing command line arguments"
'Parse the arguments sent to the program when it is run
If Len(Command$) = 0 Then
    Exit Sub
    End If
spliceWords (Command$)
count = 0
Do While Not word(count) = "" Or count >= UBound(word) - 1
    Select Case word(count)
        Case "-reboot"
            system.rebooted = True
            noSplash = True
            GoSystemTray
        Case "-systray"
            noSplash = True
            GoSystemTray
        Case "-port"
            system.mainPort = Val(word(count + 1))
            count = count + 1
        Case "-netport"
            system.netlinkPort = Val(word(count + 1))
            count = count + 1
        Case "-httpport"
            system.httpPort = Val(word(count + 1))
            count = count + 1
        Case "-minimize"
            mainForm.WindowState = 1 'minimized
        End Select
    count = count + 1
    Loop
End Sub

Sub loadSplash()
'Shows the 'Splash Screen'
If Not App.TaskVisible Then
    Me.visible = False
    Exit Sub
    End If
If noSplash Then
    mainForm.show
    Else
        mainForm.Hide
        splashForm.visible = -1
        End If
End Sub

Private Sub AboutHelp_Click()
aboutForm.visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
VBGTray.cbSize = Len(VBGTray)
VBGTray.hwnd = Me.hwnd
VBGTray.uId = vbNull
Call Shell_NotifyIcon(NIM_DELETE, VBGTray)

Dim nets As Integer, count As Integer
'Kill the netlinks in a somewhat more polite way
For nets = 1 To UBound(net)
    If madeNetlinks(nets) Then
        If Netlink(nets).Connected Then
            netout "DISCONNECT" & LF, nets
            dropNetlink (nets)
            End If
        For count = 1 To UBound(user)
            If user(count).netlinkType And user(count).netlinkFrom = count Then
                writeRoom "", "~FM" & user(count).name & " drifts back across the ethers" & CRLF
                removeUser (count)
                End If
            Next count
        End If
    Next nets
End Sub

Private Sub Form_Resize()
Line1.X2 = mainForm.Width
Line2.X2 = mainForm.Width
lightbulbOn islighton
End Sub

Private Sub FullReMenu_Click()
Dim count As Integer
For count = Re.LBound To Re.UBound
    Re_Click (count)
    Next count
lighter system.talkerName & " reloaded"
End Sub

Private Sub g_Paint()
Dim count As Integer, graphpos As Integer, lastpos As Integer
Dim largest As Integer, idle_big As Integer, idletime As Integer
Dim idlenow As Integer, Total As Integer, idles As Integer
Dim found As Boolean, csave As Long
Static wasnonelast  As Boolean
Const OFFSET_PAINT = 0.5
If preGrapherInit Then
    Exit Sub
    End If
found = False
g.Cls
If actPos < LBound(grapher) Then
    actPos = LBound(grapher)
    End If
If whatToGraph = graph.GRAPH_NULL Then
    Exit Sub
    End If
g.ScaleWidth = UBound(grapher) + 1
For count = LBound(grapher) To UBound(grapher)
    If grapher(count) > largest Then
        largest = grapher(count)
        End If
    Next count
peak.Caption = Str$(largest)
largest = largest + 1
If largest < 13 Then
    largest = 13
    End If
g.ScaleHeight = largest
For count = 0 To 12
    If Not Trim$(side(count).Caption) = Trim$(Str$(Int((largest / 12) * count))) Then
        side(count).Caption = Trim$(Str$(Int((largest / 12) * count)))
        End If
    Next count

For count = actPos To LBound(grapher) Step -1
    graphpos = graphpos + 1
    If grapher(count) = 0 Then
        idles = idles + 1
        idle_big = idle_big + 1
        idlenow = idlenow + 1
        If idle_big > idletime Then
            idletime = idle_big
            End If
        Total = Total + 1
        Else
            If Not grapher(count) = -1 Then
                found = True
                idlenow = 0
                Total = Total + 1
                End If
            idle_big = 0
            End If
    lastpos = grapher(count)
    Next count

For count = UBound(grapher) To actPos Step -1
    graphpos = graphpos + 1
    If grapher(count) = 0 Then
        idles = idles + 1
        idle_big = idle_big + 1
        idlenow = idlenow + 1
        If idle_big > idletime Then
            idletime = idle_big
            End If
        Total = Total + 1
        Else
            If Not grapher(count) = -1 Then
                found = True
                idlenow = 0
                Total = Total + 1
                End If
            idle_big = 0
            End If
    g.Line (graphpos - OFFSET_PAINT, g.ScaleHeight - lastpos)-(graphpos - OFFSET_PAINT + 1, g.ScaleHeight - grapher(count))
    lastpos = grapher(count)
    Next count
idle.Caption = Str$(idletime)
lastidle.Caption = Str$(idlenow)
If Not actPos = UBound(grapher) Then
    actstat.Caption = Str$(grapher(actPos + 1))
    Else
        If grapher(LBound(grapher)) = -1 Then
            actstat.Caption = "0"
            Else
                actstat.Caption = Str$(grapher(LBound(grapher)))
                End If
        End If
If Total = 0 Or Total - idles = 0 Then
    actper.Caption = "0"
    Else
        actper.Caption = Str$(100 - Int(100 * (idles / Total)))
        End If
If Not found Then
    csave = g.ForeColor
    g.ForeColor = &O0
    g.CurrentX = 1
    g.CurrentY = 0
    g.Print "No Activity"
    g.ForeColor = csave
    wasnonelast = True
    End If
End Sub

Private Sub graphing_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
graphing.ToolTipText = "Graphing " & graphing.Caption
End Sub

Private Sub http_Accept(Index As Integer, SocketId As Integer)
Dim count As Integer, open_http As Integer, rememberaddy As String
For open_http = 1 To 5
    If Not http(open_http).Connected Then
        http(open_http).BufferSize = 2000
        http(open_http).Binary = True
        http(open_http).AddressFamily = AF_INET
        http(open_http).Protocol = IPPROTO_TCP
        http(open_http).SocketType = SOCK_STREAM
        http(open_http).Blocking = False
        http(open_http).Accept = SocketId
        Exit For
        End If
    Next open_http
If open_http > 5 Then
    For count = 1 To 5
        http(count).Abort
        Next count
    Socket3.Abort
    Socket3.Accept = SocketId
    Socket3.Disconnect
    Socket3.Flush
    Exit Sub
    End If

rememberaddy = http(open_http).PeerName
DoEvents
For count = LBound(http_sites) To UBound(http_sites)
    If http_sites(count) = rememberaddy Then
        count = -1
        Exit For
        End If
    Next count
If Not count = -1 Then
    http_connections = http_connections + 1
    If http_connections >= 32500 Then
        http_connections = 0
        End If
    http_sitepos_pointer = http_sitepos_pointer + 1
    http_sites(http_sitepos_pointer) = rememberaddy
    End If
httpActions = httpActions + 1
End Sub

Private Sub http_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)
Response = SOCKET_ERRIGNORE
End Sub

Private Sub http_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
Dim html As String, count As Integer, cin As String, choose As String
If DataLength > 1999 Then
    http(Index).Flush
    End If
http(Index).RecvLen = DataLength
cin = http(Index).RecvData
spliceWords (cin)
If Len(word(1)) > 1 Then
    word(1) = Right$(word(1), Len(word(1)) - 1)
    If Left$(word(1), 1) = "&" Then
        choose = "&"
        If Len(word(1)) > 1 Then
            word(1) = Right(word(1), Len(word(1)) - 1)
            Else
                choose = ""
                End If
        Else
            choose = word(1)
            End If
    Else
        choose = word(1)
        End If
If Len(choose) > 1 And Right$(choose, 1) = "/" Then
    choose = Left$(choose, Len(choose) - 1)
    End If
Select Case LCase$(choose)
    Case "&"
        html = html_ex_user(word(1))
    Case "/"
        html = html_index
    Case "who"
        html = html_who
    Case "examine"
        html = html_examine
    Case Else
        html = html_error
        End Select
http_requests = http_requests + 1
http(Index).Flush
http(Index).SendLen = Len(html)
http(Index).SendData = html
http(Index).Disconnect
End Sub

Private Sub light_time_Timer()
lightbulbOn False
StatusBar.Panels(2).text = ""
light_time.Enabled = False
End Sub

Private Sub LoggingFile_Click()
If LoggingFile.Checked Then
    ClearLogFile_Click
    writeSyslog "The system logbook has been ~FRdisabled"
    LoggingFile.Checked = False
    Else
        LoggingFile.Checked = True
        writeSyslog "The system logbook is now ~FGenabled"
        End If
End Sub

Private Sub mailsock_Connect()
SMTP_STATE = EMAIL_HELO
End Sub

Private Sub mailsock_Disconnect()
SMTP_STATE = EMAIL_NOT_CONNECTED
If mail_out.success Then
    writeSyslog "A mail message sent ~FGsuccessfully"
    Else
        writeSyslog "~FRA mail message was unsuccessfully sent"
        End If
End Sub

Private Sub mailsock_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
Dim UserName As String
If userIsOnline(mail_out.u_to) Then
    user(getUser(mail_out.u_to)).sfVerifyed = False
    user(getUser(mail_out.u_to)).sfRec = False
    user(getUser(mail_out.u_to)).sfVercode = gen_ver_code
    Else
        user(0).name = mail_out.u_to
        loadUserData 0
        user(0).sfVerifyed = False
        user(0).sfRec = False
        user(0).sfVercode = gen_ver_code
        End If
drop_smtp
End Sub

Private Sub mailsock_Read(DataLength As Integer, IsUrgent As Integer)
Dim dataread As String
mailsock.RecvLen = DataLength
dataread = mailsock.RecvData
'Mail States
'EMAIL_NOT_CONNECTED = 0  'EMAIL_CONNECTING = 1
'EMAIL_HELO = 2           'EMAIL_WAIT_GREET = 3
'EMAIL_WAIT_MYID_ACK = 4  'EMAIL_WAIT_SET_RECEIVER_ACK = 5
'EMAIL_WAIT_EDIT = 6      'EMAIL_EDITING = 7
'EMAIL_END_EDIT = 8       'EMAIL_QUIT = 9
Select Case SMTP_STATE
    Case EMAIL_HELO
        smtp_send "HELO " & mailsock.LocalName & CRLF
        SMTP_STATE = EMAIL_WAIT_GREET
    Case EMAIL_WAIT_GREET
        smtp_send "MAIL FROM:<none@none.net>" & CRLF
        SMTP_STATE = EMAIL_WAIT_MYID_ACK
    Case EMAIL_WAIT_MYID_ACK
        If Get_Code(dataread) = 220 Or Get_Code(dataread) = 250 Then
            smtp_send "RCPT TO:<" & mail_out.u_email & ">" & CRLF
            SMTP_STATE = EMAIL_WAIT_SET_RECEIVER_ACK
            Else
                drop_smtp
                End If
    Case EMAIL_WAIT_SET_RECEIVER_ACK
        If Get_Code(dataread) = 250 Or Get_Code(dataread) = 251 Then
            smtp_send "DATA" & CRLF
            SMTP_STATE = EMAIL_WAIT_EDIT
            Else
                drop_smtp
                End If
    Case EMAIL_WAIT_EDIT
        If Get_Code(dataread) = 354 Then
            SMTP_STATE = EMAIL_EDITING
            smtp_send "Date: " & Format$(Now, "dd mmm yyyy hh:nn:ss") & CRLF _
            & "From: " & mail_out.userid & " <" & mail_out.u_from & ">" & CRLF _
            & "Subject: " & system.talkerName & " mail message" & CRLF _
            & "To: " & mail_out.u_to & " <" & mail_out.u_email & ">" & CRLF _
            & CRLF & "This is a forward message from " & system.talkerName & " to " & mail_out.u_to & CRLF _
            & "On " & Format$(mail_out.timestamp, "dddd, mmmm d, yyyy") & " at " _
            & Format$(mail_out.timestamp, "hh:nn:ss") & " " & mail_out.userid _
            & " wrote:" & CRLF _
            & CRLF & mail_out.message & CRLF & CRLF & "." & CRLF
            SMTP_STATE = EMAIL_END_EDIT
            Else
                drop_smtp
                End If
    Case EMAIL_END_EDIT
        SMTP_STATE = EMAIL_QUIT
        smtp_send "QUIT" & CRLF
        mail_out.success = True
'        drop_smtp
    Case Else
        If Not SMTP_STATE = EMAIL_QUIT Then
            drop_smtp
            End If
        End Select
End Sub

Private Sub Netlink_Accept(Index As Integer, SocketId As Integer)
Dim openSocket As Integer, count As Integer
Dim nn As Integer
openSocket = -1
For count = 1 To UBound(net)
    If madeNetlinks(count) = False Then
        openSocket = count
        Exit For
        End If
    Next count
For count = LBound(net) To UBound(net)
    If net(count).line = -1 And s2n(Index) = -1 And net(count).name = "" Then
        net(count).line = openSocket
        Exit For
        End If
    Next count
nn = s2n(openSocket)
'If we could not find an available netlink, we just
'give em the ol DoS(Denial of Service) message and
'kick em right out ;)
If openSocket = -1 Or nn = -1 Then
    Socket3.Accept = SocketId
    Socket3.Flush
    Socket3.SendLen = 17
    Socket3.SendData = "DENIED CONNECT 2" & Chr$(10)
    Socket3.Disconnect
    DoEvents
    Exit Sub
    End If
                  
Load Netlink(openSocket)
madeNetlinks(openSocket) = True
Netlink(openSocket).Accept = SocketId
Netlink(openSocket).Flush
net(nn).inpstr = ""

'Now I never saw the need for the site thing for netlinks.
'The things are passworded so it should be no problem. The
'site check on NUTS based server made the implementation of
'TOPs especially difficult for servers that have dynammic
'IP addresses.
netout "NUTS 3.3.3SouthWest1.0" & LF, openSocket
'Probably an idiot portscanner or web browser if this is triggered
If Not madeNetlinks(openSocket) Then
    Exit Sub
    End If
DoEvents
netout "GRANTED CONNECT" & LF, openSocket
net(nn).state = NETLINK_VERIFYING
updateNetstat
mainForm.tree.refresh
End Sub

Private Sub Netlink_Connect(Index As Integer)
'This will let us know that we have connected and
'change the netlink's state to indicate that we
'have sucessfully established a connection with
'the remote host.
Dim nn As Integer
'I have no idea why but whenever the thing makes an
'accept call, this event is also triggered. I designed
'this event to only be called on a connection out.
nn = s2n(Index)
If madeNetlinks(Index) = True And net(nn).state = NETLINK_CONNECTING Then
    If Not net(nn).wasAutoConnected Then
        writeSyslog "Connection made to ~FG" & net(s2n(Index)).name
        net(nn).wasAutoConnected = False
        End If
    End If
updateNetstat
End Sub

Private Sub Netlink_Disconnect(Index As Integer)
Dim count As Integer
dropNetlink (Index)
End Sub

Private Sub Netlink_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)
'Dim nn As Integer
'MsgBox Netlink(Index).LastError
'Netlink(Index).LastError = 0
'Netlink(Index).Flush
'If Not Index = 0 Then
'    If net(s2n(Index)).state = NETLINK_CONNECTING Then
'        writeSyslog "Could not connect to " & net(s2n(Index)).name
'        net(s2n(Index)).state = NETLINK_DOWN
'        madeNetlinks(Index) = False
'        Unload mainForm.Netlink(Index)
'        End If
'    End If
'updateNetstat
'Response = SOCKET_ERRIGNORE
'nn = s2n(Index)
'If nn >= 0 Then
'    net(nn).wasAutoConnected = False
'    End If
''dropNetlink Index
Response = SOCKET_ERRIGNORE
End Sub

Private Sub Netlink_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
Dim count As Integer, count2 As Integer, lenToSpace As Integer
Dim dataread As String, msg As String, found As Boolean
Dim nn As Integer, usernum As Integer, firstLine As String
'I know this seems odd but if the user reloads the netlink settings
'then the netlink will be lost.
nn = s2n(Index)
If nn < 0 Then
    dropNetlink Index
    Exit Sub
    End If
net(nn).bytesIn = net(nn).bytesIn + DataLength
Netlink(Index).RecvLen = DataLength
On Error Resume Next
dataread = Netlink(Index).RecvData
If Err Then
    dropNetlink Index
    End If

'As specified in the NUTS NetLINK protocol spec, items
'are terminated by line feed. If we do not find
'a line feed, it should not be processed.
net(nn).inpstr = net(nn).inpstr & dataread
If Right$(dataread, 1) = LF Or Right$(dataread, 1) = LF Then
    If Len(net(nn).inpstr) > 1 Then
        net(nn).inpstr = Left$(net(nn).inpstr, Len(net(nn).inpstr) - 1)
        If Len(net(nn).inpstr) > 1 Then
            If Right$(net(nn).inpstr, 1) = LF Or Right$(net(nn).inpstr, 1) = Chr$(13) Then
                net(nn).inpstr = Left$(net(nn).inpstr, Len(net(nn).inpstr) - 1)
                End If
            End If
        Else
            net(nn).inpstr = ""
            End If
    Else
        updateNetstat
        Exit Sub
        End If

'Now if someone is sending us null input, we had better
'be conformists and send ERROR, even though this should
'never happen unless someone is attacking the port or
'something equally stupid.
If Right$(net(nn).inpstr, 1) = LF Then
    netout "ERROR" & LF, Index
    net(nn).inpstr = ""
    updateNetstat
    Exit Sub
    End If
'SouthWest Netlinks like to tag along LFs for some reason
If Right$(net(nn).inpstr, 1) = LF And Len(net(nn).inpstr) > 1 Then
    net(nn).inpstr = Replace$(net(nn).inpstr, LF, "")
    End If
'Now we will be loading the input into words so that
'we can more easily work with arguments and such.
spliceWords (net(nn).inpstr)

Dim x As String
x = word(2)

'I was perplexed when I didnt have this check in
'there and got this funny 'NL' thing when a netlink
'user pressed enter. I found out by scanning through
'the NUTS code that we should convert this to a LF
'when received.
If word(2) = "NL" And word(3) = "" Then
    word(2) = LF
    net(nn).inpstr = word(0) & word(1) & word(2)
    End If

'This is the point as which we are ready to start
'routing the data to different subs. We will first
'route depending on the state of the netlink. Then
'we will route on the first word of the input.
Select Case net(nn).state
    Case NETLINK_CONNECTING
        Select Case word(0)
            Case "KA"
                'Ignore it, it is just a keepalive signal.
            Case "NUTS"
                net(nn).version = word(1)
                If InStr(net(nn).version, "SW") Then
                    net(nn).southwest = True
                    Else
                        net(nn).southwest = False
                        End If
                'Darn Nagle Algrorythm... Since these
                'two are sometimes sent together as a
                'result of the Nagle Algorythm, we will
                'just put a check in for it. This is the
                'only time Ive run into probs with it tho.
                If InStr(net(nn).inpstr, "GRANTED") Then
                    'We arent even going to bother checking
                    'what they are granting, we already know.
                    netout "VERIFICATION " & net(nn).password & "3.3.3SouthWest1.0" & LF, Index
                    End If
            Case "GRANTED"
                If InStr(word(1), "CONNECT") Then
                    'Now we have been told that we can connect
                    'so we send our verification.
                    netout "VERIFICATION " & net(nn).password & " 3.3.3SouthWest1.0" & LF, Index
                    End If
            Case "VERIFY"
                Select Case word(1)
                    Case "BAD"
                        netout "DISCONNECT" & LF, Index
                        dropNetlink (Index)
                    Case "OK"
                        statbarNetlinksUpdate
                        writeRoom "", "~FY~OLSYSTEM: ~RSNew connection to service " & net(nn).name & " in the " & net(nn).room & CRLF
                        net(nn).state = NETLINK_UP
                        updateActiveNetlinks
                        Select Case word(2)
                            'These are the accesses for the
                            'other talker that it passes along
                            'with the VERIFY response.
                            Case "IN"
                                net(nn).access = ACCESS_IN
                            Case "OUT"
                                net(nn).access = ACCESS_OUT
                            Case "ALL"
                                net(nn).access = ACCESS_ALL
                            Case Else
                                writeSyslog "~FRInvalid access sent"
                                netout "DISCONNECT" & LF, Index
                                dropNetlink (Index)
                                End Select
                Case Else
                        netout "DISCONNECT" & LF, Index
                        dropNetlink (Index)
                        End Select
            Case Else
                'During handshaking routines we wont tolerate
                'anything unexpected.
                netout "DISCONNECT" & LF, Index
                dropNetlink (Index)
                End Select
    Case NETLINK_VERIFYING
        'The NUTS server is very strict on only accepting
        'the VERIFICATION code and if anything else comes
        'along, even a keepalive, we disconnect. This is
        'tricky as it means more code for the KA send but
        'is a good security measure.
        If word(0) = "VERIFICATION" Then
            found = False
            For count = LBound(net) To UBound(net)
                If word(1) = net(count).password Then
                    found = True
                    statbarNetlinksUpdate
                    net(count).state = NETLINK_UP
                    net(count).bytesIn = Len(dataread)
                    net(count).bytesOut = 34
                    writeSyslog "Incoming netlink from ~FG" & net(count).name
                    writeRoom "", "~OL~FYSYSTEM: ~RSNew connection to service " & net(count).name & " in the " & net(count).room
                    If s2n(Index) > -1 Then
                        net(s2n(Index)).line = -1
                        End If
                    net(count).line = Index
                    nn = s2n(Index)
                    If net(nn).allowIn And net(nn).allowOut Then
                        netout "VERIFY OK ALL" & LF, Index
                        ElseIf net(nn).allowIn And Not net(nn).allowOut Then
                            netout "VERIFY OK IN" & LF, Index
                        Else
                            netout "VERIFY OK OUT" & LF, Index
                            End If
                    Exit For
                    End If
                Next count
            updateActiveNetlinks
            If Not found Then
                netout "DISCONNECT" & LF, Index
                dropNetlink (Index)
                Exit Sub
                End If
            Else
                netout "DISCONNECT" & LF, Index
                dropNetlink (Index)
                End If
    Case NETLINK_UP
        Select Case word(0)
            Case "KA"
                'Ignore this one, its just the keepalive
            Case "DISCONNECT"
                'The netlink is disconnecting
                netout "DISCONNECT" & LF, Index
                dropNetlink (Index)
            Case "TRANS"
                'Little check to see if it is allowed
                If Not net(nn).allowIn = True Then
                    netout "DENIED " & word(1) & " 4" & LF, Index
                    updateNetstat
                    net(nn).inpstr = ""
                    Exit Sub
                    End If
                'This will be called when a user is to be transfered
                'across the netlink.
                'Find a open spot to store the user.
                usernum = -1
                userResize True
                For count = 1 To UBound(user)
                    If Not user(count).operational Then
                        usernum = count
                        user(usernum).line = count
                        Exit For
                        End If
                    Next count
                'There is already another user logged on with the
                'same name. So we must reject the connection.
                If userIsOnline(word(1)) Then
                    netout "DENIED " & word(1) & " 5" & LF, Index
                    updateNetstat
                    net(nn).inpstr = ""
                    Exit Sub
                    End If
                'Plain old failure to create a session.
                If usernum < 0 Then
                    netout "DENIED " & word(1) & " 6" & LF, Index
                    Else
                        If userExists(word(1)) Then
                            If word(2) = loadUserPassword(word(1)) Then
                                user(usernum).name = word(1)
                                loadUserData (usernum)
                                With user(usernum)
                                    .netlinkType = True
                                    .netlinkFrom = Index
                                    .listening = True
                                    .listing = mainForm.connectionsList.ListCount
                                    .state = STATE_NORMAL
                                    .operational = True
                                    End With
                                netout "GRANTED " & word(1) & LF, Index
                                masterLogin user(usernum)
                                mainForm.connectionsList.AddItem word(1) ' & " (remote)"
                                writeSyslog "User ~FB" & user(usernum).name & "~RS transfered from ~FG" & net(s2n(user(usernum).netlinkFrom)).name
                                writeRoom "", "~OL~FM" & user(usernum).name & " steps in from cyberspace~RS" & CRLF
                                user(usernum).room = rooms(1).name
                                If user(usernum).room = "" Then
                                    user(usernum).room = "Jail"
                                    End If
                                Else
                                    netout "DENIED " & word(1) & " 7" & LF, Index
                                    End If
                            Else
                                user(usernum).name = word(1)
                                createNewAccount user(usernum), usernum
                                user(usernum).password = word(2)
                                user(usernum).rank = 2
                                user(usernum).listening = True
                                user(usernum).netlinkType = True
                                user(usernum).netlinkFrom = Index
                                user(usernum).state = STATE_NORMAL
                                user(usernum).operational = True
                                user(usernum).listing = mainForm.connectionsList.ListCount
                                loadUserData (usernum)
                                netout "GRANTED " & word(1) & LF, Index
                                masterLogin user(usernum)
                                mainForm.connectionsList.AddItem word(1) ' & " (remote)"
                                writeSyslog "User ~FB" & user(usernum).name & " transfered from ~FG" & net(s2n(user(usernum).netlinkFrom - 1)).name
                                writeRoom "", "~OL~FM" & user(usernum).name & " steps in from cyberspace~RS" & CRLF
                                user(usernum).room = rooms(1).name
                                If user(usernum).room = "" Then
                                    user(usernum).room = "Jail"
                                    End If
                                End If
                        End If
            Case "ACT"
                'If wordCount(net(nn).inpstr) < 3 Or NetToUser(word(1)) = 0 Then
                '    netout "ERROR" & LF, Index
                '    Else
                usernum = NetToUser(word(1))
                user(usernum).inpstr = stripOne(stripOne(net(nn).inpstr))
                actions = actions + 1
                processNormal (usernum)
                '        End If
            Case "REL"
                'Good bye! They want to take their user back
                removeUser (getUser(word(1)))
            Case "REMVD"
                usernum = getUser(word(1))
                user(usernum).atNetlink = -1
                returnedFromNetlink usernum
            Case "MSG"
                Dim packetLen As Integer
                msg = net(nn).inpstr
                packetLen = InStr(msg, LF & "EMSG")
                If packetLen = 0 Then
                    net(nn).inpstr = net(nn).inpstr & CRLF
                    Exit Sub
                    End If
                Do
                    firstLine = Left$(msg, packetLen + 4)
                    msg = Right$(msg, Len(msg) - (packetLen + 4))
                    If Left$(msg, 1) = LF And Len(msg) > 1 Then
                        msg = Right$(msg, Len(msg) - 1)
                        End If
                    netPacket firstLine
                    packetLen = InStr(msg, LF & "EMSG")
                    Loop While packetLen > 0
                net(nn).inpstr = vbNullString
            Case "GRANTED"
                usernum = getUser(word(1))
                user(usernum).oldRoom = user(usernum).room
                user(usernum).room = "@" & net(s2n(Index)).name
                user(usernum).netlinkPending = False
                user(usernum).atNetlink = Index
                user(usernum).listening = False
                send "~FMYou have entered " & net(s2n(Index)).name & "~RS" & CRLF, usernum
                writeRoom user(usernum).room, "~FM" & user(usernum).name & " " & user(usernum).exitMsg & " Netlink of " & net(s2n(Index)).name & CRLF
                user(usernum).inpstr = ""
                netout "ACT " & user(usernum).name & " look" & LF, net(count).line
            Case "DENIED"
                usernum = getUser(word(1))
                user(usernum).netlinkPending = False
                Select Case Val(word(2))
                    Case 4
                        send "This link is for incoming users only" & CRLF, usernum
                    Case 5
                        send "There is a user of the same name logged in at the remote site" & CRLF, usernum
                    Case 6
                        send "The remote service was unable to create a session for you" & CRLF, usernum
                    Case 7
                        'this one gives probs at times
                        net(nn).inpstr = vbNullString
                        send "Incorrect password... Use .go <netlink> [password]" & CRLF, usernum
                    Case 8
                        send "Login below minimum login level" & CRLF, usernum
                    Case 9
                        send "You have been banished from that talker" & CRLF, usernum
                        End Select
            Case "RSTAT"
                msg = "~BB*** Remote statistics ***" & CRLF & CRLF & _
                            "SouthWest version    : " & Trim$(Str$(App.Major) & "." & App.Minor & "." & App.Revision) & CRLF & _
                            "Host                 : " & Socket1.LocalName & CRLF & _
                            "Ports (Main/Link)    : " & Trim$(Str$(Socket1.LocalPort)) & ", " & Trim$(Str$(Netlink(0).LocalPort)) & CRLF & _
                            "Number of users      : " & Trim$(usersOnline) & CRLF & _
                            "Remote user maxlevel : " & ranks(UBound(ranks)) & CRLF & _
                            "Remote user deflevel : " & ranks(1) & CRLF
                msgout Index, word(1), msg
            Case Else
                netout "ERROR" & LF, Index
            End Select
    Case Else
        Netlink(Index).Flush
        netout "ERROR" & LF, Index
        End Select
net(nn).inpstr = ""
updateNetstat
End Sub

Private Sub ContentsHelp_Click()
HTMLHelpContents 1, "main"
End Sub

Private Sub Form_Load()
Dim count As Integer, myObj As Object
On Error Resume Next
BOOTING = True: preGrapherInit = True
'Set default menu items
Call SetMenuDefaultItem(GetSubMenu(GetMenu(Me.hwnd), 3), 7, True)
Call SetMenuDefaultItem(GetSubMenu(GetMenu(Me.hwnd), 1), 3, True)
Call SetMenuDefaultItem(GetSubMenu(GetMenu(Me.hwnd), 2), 0, True)
'Register with the ICQ API
system.icqHook = IIf(ICQSetKey("Scott Lloyd", "password", "4C4AD7C13902539C") = 1, True, False)

'Dimension some arrays that autosize
ReDim user(2)
ReDim clones(2)
maxUsers = 100
action_frm.Left = 2070
action_frm.Top = 450
'Disable buttons that could cause screwups
Toolbar.Buttons(1).Enabled = False
Toolbar.Buttons(2).Enabled = False
FullReMenu.Enabled = False
RebootReboot.Enabled = False
For count = Re.LBound To Re.UBound
    Re(count).Enabled = False
    Next count

'Check and see if we need to run Autopatch
If Not fileExists("southwest.s") Then
    Me.Enabled = False
    Me.Hide
    autoPatchFile = "\master.ap"
    patcherForm.visible = True
    If Not fileExists("southwest.s") Then
        writeSyslog "SouthWest was unable to boot properly"
        For Each myObj In Controls
            myObj.Enabled = False
            Next myObj
        For Each myObj In Toolbar.Buttons
            myObj.Enabled = False
            Next myObj
        Me.show
        Me.Enabled = True
        Me.SetFocus
        myObj = Nothing
        End
        Exit Sub
        End If
    Me.show
    Me.Enabled = True
    Unload patcherForm
    Me.SetFocus
    End If

'Load up the system configuration
'Display that spiffy message
writeSyslog "Server is booting"
Picture1.ToolTipText = "SouthWest v" & App.Major & _
    "." & App.Minor & "." & App.Revision

'Load various system components
'In some cases, the order that these tasks are executed
'makes a difference, so do not rearrange the order
userResize
loadRanks
parseArguments
loadSystem
If (InStr(Command$, "-systray") = 0) And (system.rebooted = False) Then
    Me.show
    App.TaskVisible = True
    Else
        Me.Hide
        App.TaskVisible = False
        End If
loadSplash
autoPatchCheck
loadPlugins
loadColors
loadGlobals
showGrapher
loadGrapher
loadFiglets
loadFigletMenu
loadRooms
loadNetlink
loadCommands
loadSpool
loadViewer
tree.Nodes("GRAPHER").Expanded = True
tree.Nodes(2).Selected = True
loadMotd
loadSwears
resizeClones
loadSockets
loadConnectNetlinks

'Super-Plugin checks
Toolbar.Buttons(3).Enabled = fileExists(App.Path & "\Plugins\tools\tools.exe")
SouthWestToolsTools.Enabled = Toolbar.Buttons(3).Enabled
UserfileConversionsFile.Enabled = fileExists(App.Path & "\Plugins\convert\convert.exe")

'Initialize user states
For count = LBound(user) To UBound(user)
    cleanUser user(count)
    Next count
    
'Re-enable everything
Toolbar.Buttons(1).Enabled = True
Toolbar.Buttons(2).Enabled = True
FullReMenu.Enabled = True
RebootReboot.Enabled = True
For count = Re.LBound To Re.UBound
    Re(count).Enabled = True
    Next count
Unload splashForm
If tree.SelectedItem.Key = "SERVER_INFO" Or tree.SelectedItem.Key = "SYSLOG" Then
    treeLoad
    End If
'Finish boot
mainForm.tree.SetFocus
lighter system.talkerName & " is ready"
BOOTING = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim count As Integer, nets As Integer, usrSize As Integer

writeRoom "", "~BR~FKPROGRAM TERMINATING... SHUTTING DOWN NOW!" & CRLF

'Close all connections gracefully
If Socket1.listening Then
    Socket1.Action = SOCKET_CLOSE
    End If

count = 1
usrSize = UBound(user)
Do Until count > usrSize
    If userIsOnline(user(count).name) Then
        killUser (count)
        End If
    usrSize = UBound(user)
    count = count + 1
    Loop
'The following code may seem a little odd but the program
'was pulling some shindings like not totally leaving so I
'put the following lines in and it works fine now since
'it unloads every form.
Dim Form As Form
For Each Form In Forms
    Unload Form
    Set Form = Nothing
    Next Form
    
'Save mail data to a spool
save_spool

'reboot server if we are supposed to do so
If system.shutdownType = SHUTDOWN_REBOOT Then
    App.Title = "Rebooting SouthWest"
    system.shutdownType = SHUTDOWN_NONE
    On Error Resume Next
    Call ShellExecute(Me.hwnd, "Open", ".\rboot.exe", App.EXEName & " -reboot", App.Path, 1)
    End If
End
End Sub

Private Sub LicenceHelp_Click()
splashForm.visible = True
End Sub

Private Sub Logon(linenum As Integer)
'We got one comming in! Better display one of those catchy MOTDs
'(Message of the Day) like all the other fancy talkers.
Dim outMOTD As String, nextcr As Integer
user(linenum).operational = True
user(linenum).visible = True
user(linenum).listening = False
user(linenum).netlinkType = False
user(linenum).name = ""
user(linenum).password = ""
user(linenum).rank = 2
user(linenum).state = 0
user(linenum).pager = 27
user(linenum).room = rooms(1).name
'The MOTD is pre-colorparsed. It can also be VERY large. This means
'that we have to chop it up WHILE respecting line length
outMOTD = MOTD
Do
    If Len(outMOTD) > SEND_CHOP Then
        nextcr = InStr(SEND_CHOP, outMOTD, CRLF)
        If nextcr = 0 Then
            nextcr = SEND_CHOP
            End If
        Else
            nextcr = Len(outMOTD)
            End If
    send Left$(outMOTD, nextcr), linenum
    If Len(outMOTD) > nextcr Then
        outMOTD = Right$(outMOTD, Len(outMOTD) - (nextcr + 1))
        Else
            outMOTD = vbNullString
            End If
    Loop While Len(outMOTD) > 1

If user(linenum).room = "" Then
    user(linenum).room = "Jail"
    End If
user(linenum).timeon = 0
send CRLF & "Enter your name: ", linenum
End Sub

Sub loadRanks()
lighter "Loading ranks"
If Not fileExists(App.Path & "\Misc\Ranks.S") Then
    MsgBox "Cannot find Misc\Ranks.S", vbCritical
    End
    End If
Dim count As Integer, FromFile As String, max_ranks As Integer
Open App.Path & ".\Misc\Ranks.S" For Input As #1
max_ranks = -1
Do While Not EOF(1)
    Line Input #1, FromFile
        If Not Trim$(FromFile) = "" Then
            max_ranks = max_ranks + 1
            End If
    Loop
Close #1
ReDim ranks(max_ranks)
Open App.Path & "\Misc\Ranks.S" For Input As #1
For count = 0 To max_ranks
    Line Input #1, FromFile
    FromFile = Trim$(FromFile)
    If Len(FromFile) > 10 Then
        FromFile = Left$(FromFile, 10)
        End If
    If Not FromFile = "" Then
        ranks(count) = FromFile
        End If
    Next count
Close #1
End Sub

Private Sub NotepadConfigFile_Click()
Shell "NOTEPAD " & App.Path & "\southwest.s", vbNormalFocus
End Sub

Private Sub OpenSouthWest_System_Click()
Me.WindowState = 0
Me.show
lightbulbOn (islighton)
Call Shell_NotifyIcon(NIM_DELETE, VBGTray)
End Sub

Private Sub p_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
keyScroll KeyCode
End Sub

Private Sub Picture1_DblClick()
'EASTER EGGS RULE!!!
If easterEgg = True Then
    eggForm.visible = -1
    eggForm.Left = mainForm.Left
    eggForm.Top = mainForm.Top
    End If
easterEgg = False
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
'This will detect if the user has clicked on our
'secret little easter egg.
Const SHIFT_MASK = 1
Const CTRL_MASK = 2
Const ALT_MASK = 4
Dim ShiftDown, AltDown, CtrlDown As Integer

    ShiftDown = (Shift And SHIFT_MASK) > 0
    AltDown = (Shift And ALT_MASK) > 0
    CtrlDown = (Shift And CTRL_MASK) > 0
If ShiftDown And AltDown And CtrlDown Then
    easterEgg = True
    End If
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
easterEgg = False
End Sub

Private Sub Plugin_Click(Index As Integer)
Dim count As Integer
For count = LBound(plugs) To UBound(plugs)
    If plugs(count).menuPos = Index Then
        Err = 0
        On Error Resume Next
        Call Shell(App.Path & plugs(count).exeptr, vbNormalFocus)
        If Err Then
            lighter "There was an error running the plugin " & doubleQuote & plugs(count).name & doubleQuote
            Err = 0
            Exit Sub
            End If
        Exit For
        End If
    Next count
End Sub

Private Sub PluginPageTools_Click()
Call ShellExecute(Me.hwnd, vbNullString, "http://talker.com/southwest/", vbNullString, "c:\", 1) 'SW_SHOWNORAL
End Sub

Private Sub Re_Click(Index As Integer)
Select Case Index
    Case 0
        loadCommands
        lighter "Commands reloaded"
    Case 1
        loadFigletMenu
        loadFiglets
        lighter "Figlet fonts reloaded"
    Case 2
        loadMotd
        lighter "Login screen reloaded"
    Case 3
        If loadNetlink Then
            lighter "Netlinks reloaded"
            End If
    Case 4
        loadRooms
        lighter "Rooms reloaded"
    Case 5
        loadSystem
        lighter "Settings reloaded"
    Case 6
        loadSwears
        lighter "Swears reloaded"
        End Select
End Sub

Private Sub RebootReboot_Click()
system.shutdownType = SHUTDOWN_REBOOT
writeRoom "", "~BR~FWSYSTEM: Reboot Initiated" & CRLF
Unload mainForm
End Sub

Private Sub RefreshWindow_Click()
treeLoad
End Sub

Private Sub Shutdown_System_Click()
system.shutdownType = SHUTDOWN_SHUTDOWN
writeRoom "", "~BR~FWSYSTEM: Shutdown Initiated" & CRLF
Unload mainForm
End Sub

Private Sub Shutdown_Timer_Timer()
Dim s As Integer, smsg As String
If system.shutdownType = SHUTDOWN_SHUTDOWN Then
    smsg = "Shutting down"
    Else
        smsg = "Rebooting"
        End If
system.shutdownCount = system.shutdownCount - 1
s = system.shutdownCount
If s > 0 And s < 11 Or s = 30 Then
    writeRoom "", "~BR~FWSYSTEM: " & smsg & " in" & Str$(system.shutdownCount) & " seconds" & CRLF
    End If
If s = 60 Then
    writeRoom "", "~BR~FWSYSTEM: " & smsg & " in one minute" & CRLF
    End If
If s = 3600 Then
    writeRoom "", "~BR~FWSYSTEM: " & smsg & " in one hour" & CRLF
    End If
If s <= 0 Then
    writeRoom "", "~BR~FWSYSTEM: " & smsg & " NOW!" & CRLF
    Unload mainForm
    End If
End Sub

Private Sub ShutdownTools_Click()
system.shutdownType = SHUTDOWN_SHUTDOWN
writeRoom "", "~BR~FWSYSTEM: Shutdown Initiated" & CRLF
Unload mainForm
End Sub


Private Sub Socket1_Accept(SocketId As Integer)
Dim openSocket As Integer
Dim count As Integer
'This section will check for an open socket. If the var
'openSocket is not altered, that means that none were found
'and that the talker is full.
openSocket = -1
'Checks for dumps (occurs during attacks on server)
dumpCheck = True
'A very odd function... It will help us manage memory automaticly tho
userResize True

For count = 1 To UBound(user)
    If Not user(count).operational Then
        openSocket = count
        Load Socket2(openSocket)
        madeSockets(openSocket) = True
        Exit For
        End If
    Next count

'If the server is full (Past maxUsers) then this will connect for a brief
'moment on a special socket, tell the person that the server is full, and
'then dump the person off.
If openSocket = -1 Then
    Socket3.Accept = SocketId
    If Socket3.Connected = True Then
        Dim text As String
        text = "Im sorry but this server is full and cannot receive any more connections"
        Socket3.SendLen = Len(text)
        Socket3.SendData = text
        If Socket3.Connected Then
            Socket3.Disconnect
            End If
        Exit Sub
        End If
    End If

'Since one was found, we will now fill the socket in about it's self and then
'it will be passed the socket descriptor and accept it.
Socket2(openSocket).AddressFamily = AF_INET
Socket2(openSocket).Protocol = IPPROTO_TCP
Socket2(openSocket).SocketType = SOCK_STREAM
Socket2(openSocket).Binary = True
Socket2(openSocket).BufferSize = 1024
Socket2(openSocket).Blocking = False
Socket2(openSocket).Accept = SocketId

'Better check since there may be a lot of closes during
'attacks on the server.
DoEvents
If dumpCheck = False Or Not madeSockets(openSocket) Then
    Exit Sub
    End If

'Is it banned?
If isBanned(Socket2(openSocket).PeerName, BAN_SITE) Then
    send "Your site has been blocked" & CRLF, openSocket
    removeUser openSocket
    Exit Sub
    End If

'Lets tell the GUI user that they are here
user(openSocket).listing = connectionsList.ListCount
user(openSocket).line = openSocket
If Socket2(openSocket).PeerName = "" Then
    connectionsList.AddItem Socket2(openSocket).PeerAddress
    ElseIf dumpCheck Then
        connectionsList.AddItem Socket2(openSocket).PeerName
        Else
            Exit Sub
            End If
DoEvents
If Not dumpCheck Then
    Exit Sub
    End If
If Socket2(openSocket).PeerName = "" Then
    user(openSocket).site = Socket2(openSocket).PeerAddress
    Else
        user(openSocket).site = Socket2(openSocket).PeerName
        End If
'Runs the main setup of the userfiles, prints the motd, ect...
Logon openSocket
End Sub

Private Sub Socket1_Disconnect()
dumpCheck = False
End Sub

Private Sub Socket2_Disconnect(Index As Integer)
Dim UserName As String
'Remove the socket from memory, we dont need to have idle sockets lurking
removeUser Index
End Sub

Private Sub Socket2_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)
Response = SOCKET_ERRIGNORE
End Sub

Private Sub Socket2_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
'Hahaha, whats that? You thought that being the server would be easy? Well,
'in a way your right, we are on top. The thing is some clients go through
'these identity crisises and want the server to tell them who they are and
'how they should act. So we will tell them that they are and should act like
'a DEC-VT100 terminal if they are so inclined to ask.
Dim sBuffer As String, sOutput As String, sReply As String
Dim nRead As Integer, nIndex As Integer, nChar As Integer
Dim nCmd As Integer, nOpt As Integer, nQual As Integer
Dim gotCRLF As Boolean, newinpstr As String, firstchar As String

Socket2(Index).RecvLen = DataLength
sBuffer = Socket2(Index).RecvData: nRead = Socket2(Index).RecvLen
nIndex = 1
While nIndex <= nRead
    nChar = Asc(Mid$(sBuffer, nIndex, 1))
    ' If this is the Telnet IAC (Is A Command) character, then
    ' the next byte is the command
    If nChar = TELCMD_IAC Then
        If Len(sBuffer) <= 1 Then
            Exit Sub
            End If
        If nIndex < Len(sBuffer) Then
            nIndex = nIndex + 1: nCmd = Asc(Mid$(sBuffer, nIndex, 1))
            End If
        Select Case nCmd
        ' Two IAC bytes means that this isn't really a command
            Case TELCMD_IAC
                sOutput = sOutput + Chr$(nChar)
            ' The SB (sub-option) command tells us that the the client
            ' wants to know who they are. Well at least thats the only
            ' SB we are going to deal with.
            Case TELCMD_SB
                nIndex = nIndex + 1: nOpt = Asc(Mid$(sBuffer, nIndex, 1))
                If nIndex <= Len(sBuffer) Then
                
                    nIndex = nIndex + 1: nQual = Asc(Mid$(sBuffer, nIndex, 1))
                    If nOpt = TELOPT_TTYPE Then
                        ' Build a sub-option reply string and send it to
                        ' the client telling them to act like a DEC-VT100
                        sReply = Chr$(TELCMD_IAC) + Chr$(TELCMD_SB) + Chr$(nOpt) + Chr$(TELQUAL_IS) + "DEC-VT100" + Chr$(TELCMD_IAC) + Chr$(TELCMD_SE)
                        Socket2(Index).SendLen = Len(sReply): Socket2(Index).SendData = sReply
                        End If
                    End If
            End Select
        End If
    'If it starts with 27, they are prolly pressing arrow keys and all
    'that follows will be gibberish... So we just leave.
    If nChar = 27 Then
        Exit Sub
        End If
    'Check for carrige returns, line feeds, and backspaces
    If nChar > 31 And nChar < 127 Then
        If nChar <> 10 And nChar <> 13 Then
            user(Index).inpstr = user(Index).inpstr & Chr$(nChar)
            newinpstr = newinpstr & Chr$(nChar)
            End If
        ElseIf nChar = 13 Or nChar = 10 Then
            gotCRLF = True
        ElseIf nChar = 8 Or nChar = 127 Then
            If Len(newinpstr) > 1 Then
                newinpstr = Left$(newinpstr, Len(newinpstr) - 1)
                Else
                    newinpstr = ""
                    End If
            If Len(user(Index).inpstr) > 1 Then
                user(Index).inpstr = Left$(user(Index).inpstr, _
                    Len(user(Index).inpstr) - 1)
                Else
                    user(Index).inpstr = ""
                    End If
            End If
    nIndex = nIndex + 1
    Wend

'Well, you may have been wondering why I didnt just bypass inpchar all
'together and this is why: There are two major types of terminals. One is
'the line client, that sends us a line at a time, whenever the user presses
'the enter key. But there is also the wicked charictor echo client, that
'gives us the keys as the user presses them. Here is how we deal with them
'so that the other subs dont have to know the difference. One thing though:
'sometimes a char echo client echos more than one char.

'We dont want a severe buffer overflow problem, now do we?
If Len(user(Index).inpstr) >= MAX_DATA_LEN Then
    user(Index).inpstr = ""
    If user(Index).state <= STATE_LOGIN3 Then
        lighter "Recieved extremly large ammount of data from " & user(Index).site
        Else
            lighter "Recieved extremly large ammount of data from " & user(Index).name
            End If
    End If

'For charictor echo clients with no local echo support
If user(Index).charEchoing Then
    send newinpstr, Index
    End If

user(Index).idle = 0
'Are we there yet? Are we there yet? Well, here is how we find out if they
'have pressed enter. If they did we can report that, and then the server
'activity chart will be much more accurate. :)
If gotCRLF Then
    actions = actions + 1
    
 'Ohhh... You were alive all this time?
If user(Index).afk Then
    user(Index).afk = False
    writeRoom user(Index).room, user(Index).name & " shakes their head and wakes up" & CRLF
    End If
    
'Lets filter out a little bit of junk that gets inserted by some
'clients when we do the no echo thing and all
If Len(user(Index).inpstr) > 0 Then
    If Asc(user(Index).inpstr) = 1 Then
        If Len(user(Index).inpstr) > 1 Then
            user(Index).inpstr = Right$(user(Index).inpstr, Len(user(Index).inpstr) - 1)
            Else
                user(Index).inpstr = ""
                Exit Sub
                End If
        End If
    End If

'If it is a period on a single line, we will just redo the last
'command the user issued.
If user(Index).inpstr = "." And user(Index).state = STATE_NORMAL Then
    user(Index).inpstr = user(Index).oldInpstr
    End If
user(Index).oldInpstr = user(Index).inpstr
'There is a time for everything and trouble is, we have to know what the user
'is ready to handle--in a way. We have to check what state they are in. If
'you need more help on what state is for what, see the SERVER.BAS where they
'are defined as Public Const. There are heavier comments there on them.
    Select Case user(Index).state
        Case STATE_LOGIN1
        'There are two special commands that a user can input from
        'the login prompt: who and quit
        user(Index).inpstr = LCase(user(Index).inpstr)
        If LCase$(user(Index).inpstr) = "who" Then
            who (Index)
            user(Index).inpstr = ""
            End If
        If user(Index).inpstr = "quit" Then
            send "Abandoning login" & CRLF, Index
            user(Index).inpstr = ""
            killUser (Index)
            Exit Sub
            End If
        'And by what name may I have the pleasure of calling you?
        'We certainly can't have them entering a null name. If they didnt do
        'it right, they will be doing it twice.
        If isBanned(user(Index).inpstr, BAN_USER) Then
            send "This account has been frozen" & CRLF, Index
            user(Index).inpstr = ""
            End If
         If Not isNameValid(user(Index).inpstr) Then
            send "Enter your name: ", Index
            user(Index).inpstr = ""
            Exit Sub
            End If
         'Do the nasty stuff to see if they exist already
         user(Index).name = ""
         If userExists(user(Index).inpstr) Then
            'send "Welcome " & user(index).name & "!" & CRLF, index
            user(Index).password = loadUserPassword(user(Index).inpstr)
            user(Index).newUser = False
            Else
                send "New user..." & CRLF, Index
                user(Index).newUser = True
                End If
         user(Index).name = userCap(user(Index).inpstr)
         user(Index).state = STATE_LOGIN2
         send echoOff, Index
         send "Enter your password: ", Index
        Case STATE_LOGIN2
        'Now we will retreive the users's password
        'and compare it with the one on file. If the
        'user is new, we will set his state to LOGIN3
        'so that he can retype his password for verification
        'purposes; if not, the user goes to the state NORMAL.
            If Len(user(Index).inpstr) < 3 Then
                send echoOn, Index
                send "Must enter a valid password" & CRLF & "Enter your name: ", Index
                user(Index).state = STATE_LOGIN1
                user(Index).inpstr = ""
                Exit Sub
                End If
            If Asc(Left(user(Index).inpstr, 1)) = 1 Then
                user(Index).inpstr = Right(user(Index).inpstr, Len(user(Index).inpstr) - 1)
                End If
            If containsCorrupt(user(Index).inpstr) Then
                send "Must enter a valid password... Numbers and letters only" & CRLF, Index
                send echoOn, Index
                send "Enter your name: ", Index
                user(Index).state = STATE_LOGIN1
                user(Index).inpstr = ""
                Exit Sub
                End If
            If user(Index).newUser Then
                If user(Index).state = STATE_LOGIN2 Then
                    user(Index).password = crypt(user(Index).inpstr)
                    user(Index).state = STATE_LOGIN3
                    send echoOff, Index
                    rules Index
                    send "Please retype your password: ", Index
                    user(Index).inpstr = ""
                    Exit Sub
                    End If
                End If
            'Check and make sure they are themself
            If Not crypt(user(Index).inpstr) = user(Index).password Then
                send echoOn, Index
                send "Incorrect Password" & CRLF & "Enter your name: ", Index
                user(Index).state = STATE_LOGIN1
                Else
                    send "~BB~FK~LI" & "Logon success!~RS" & CRLF & CRLF & CRLF, Index
                    writeRoomExcept "", "~OL~FTEntering:~RS " & user(Index).name & " " & user(Index).desc & BELL & CRLF, user(Index).name
                    'Make them hear and let them be normal
                    loadUserData Index
                    postLoadup user(Index)
                    user(Index).listening = True
                    send echoOn, Index
                    user(Index).Index = Index
                    connectionsList.List(user(Index).listing) = user(Index).name
                    user(Index) = alreadyLoggedOn(user(Index))
                    clearScreen Index
                    look Index
                    writeSyslog "~FB" & user(Index).name & "~RS logged on from ~FG" & user(Index).site
                    If user(Index).unread Then
                        send "~LI~OL~FTYOU HAVE UNREAD MAIL" & CRLF, Index
                        End If
                    user(Index).inpstr = ""
                    user(Index).oldInpstr = ""
                    End If
        Case STATE_LOGIN3
            'This is for new users to enter their passwords
            'a second time for verification purposes.
            If crypt(user(Index).inpstr) = user(Index).password Then
                createNewAccount user(Index), Index
                send "~BB~FK~LILogon success!~RS" & CRLF & CRLF & CRLF, Index
                writeRoom "", "~OL~FTEntering:~RS " & user(Index).name & " " & user(Index).desc & CRLF
                'Make them hear and let them be normal
                loadUserData Index
                send echoOn, Index
                connectionsList.List(user(Index).listing) = user(Index).name
                user(Index) = alreadyLoggedOn(user(Index))
                user(Index).rank = 1
                postLoadup user(Index)
                clearScreen Index
                look Index
                user(Index).inpstr = ""
                writeSyslog "New user, ~FB" & user(Index).name & "~RS, logged on from ~FG" & user(Index).site
                Exit Sub
                Else
                    send "Passwords do not match" & CRLF & "Enter your name: ", Index
                    send echoOn, Index
                    user(Index).state = STATE_LOGIN1
                    End If
        Case STATE_NORMAL
            'I had this sub all comformist and local
            'and everything but yeash, I went and ruined
            'it all by making one part a sub. Why did I
            'do such a horendous thing, you ask? Well it
            'is because Netlink users are people too and
            'will be calling that sub also.
            processNormal (Index)
        Case STATE_EDITOR
            lineEditor Index, user(Index).inpstr
        Case STATE_EDPICK
            If Not user(Index).inpstr = "" Then
                editorOptions Index, user(Index).inpstr
                End If
        Case STATE_OPTION
            Select Case user(Index).options
                Case OPTION_SHUTDOWN
                    If UCase$(user(Index).inpstr) = "Y" Or UCase$(user(Index).inpstr) = "YES" Then
                        system.shutdownType = SHUTDOWN_SHUTDOWN
                        If user(Index).options = SHUTDOWN_USERCHOOSING_SHUT Then
                            system.shutdownType = SHUTDOWN_SHUTDOWN
                            Else
                                system.shutdownType = SHUTDOWN_REBOOT
                                End If
                        user(Index).state = STATE_NORMAL
                        user(Index).options = OPTION_NONE
                        If system.shutdownType = SHUTDOWN_SHUTDOWN Then
                            writeRoom "", "~BR~FW*" & user(Index).name & " has initiated a shutdown*" & CRLF
                            Else
                                writeRoom "", "~BR~FW*" & user(Index).name & " has initiated a reboot*" & CRLF
                                End If
                        Shutdown_Timer.Enabled = True
                        Else
                            system.shutdownType = SHUTDOWN_NONE
                            user(Index).state = STATE_NORMAL
                            user(Index).options = OPTION_NONE
                            End If
                Case OPTION_REBOOT
                    If UCase$(user(Index).inpstr) = "Y" Or UCase$(user(Index).inpstr) = "YES" Then
                        system.shutdownType = SHUTDOWN_SHUTDOWN
                        If user(Index).options = SHUTDOWN_USERCHOOSING_SHUT Then
                            system.shutdownType = SHUTDOWN_SHUTDOWN
                            Else
                                system.shutdownType = SHUTDOWN_REBOOT
                                End If
                        user(Index).state = STATE_NORMAL
                        user(Index).options = OPTION_NONE
                        If system.shutdownType = SHUTDOWN_SHUTDOWN Then
                            writeRoom "", "~BR~FW*" & user(Index).name & " has initiated a shutdown*" & CRLF
                            Else
                                writeRoom "", "~BR~FW*" & user(Index).name & " has initiated a reboot*" & CRLF
                                End If
                        Shutdown_Timer.Enabled = True
                        Else
                            system.shutdownType = SHUTDOWN_NONE
                            user(Index).state = STATE_NORMAL
                            user(Index).options = OPTION_NONE
                            End If
                Case OPTION_SUICIDE
                    If UCase$(user(Index).inpstr) = "Y" Or UCase$(user(Index).inpstr) = "YES" Then
                        Dim UserName As String
                        UserName = user(Index).name
                        removeUser Index
                        deleteAccount UserName
                        writeSyslog "~FR" & UserName & " commits suicide"
                        writeRoom "", "~LI~OL~FY" & UserName & " commits suicide" & CRLF
                        send "So long" & CRLF & "Have a nice life" & BELL & CRLF, Index
                        Else
                            send CRLF & "Good choice" & CRLF, Index
                            user(Index).state = STATE_NORMAL
                            user(Index).options = OPTION_NONE
                            End If
                    End Select
        End Select
    user(Index).inpstr = ""
    End If
End Sub

Private Sub Syslog_GotFocus()
mainForm.SetFocus
End Sub

Private Sub Syslog_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim count As Integer
For count = 0 To Syslog.ListCount - 1
    Syslog.Selected(count) = False
    Next count
mainForm.SetFocus
End Sub

Private Sub Syslog_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim count As Integer
For count = 0 To Syslog.ListCount - 1
    Syslog.Selected(count) = False
    Next count
mainForm.SetFocus
End Sub

Private Sub SystemTrayWindow_Click()
GoSystemTray
End Sub

Private Sub Timer2_Timer()
Dim nn As Integer, un As Integer, count As Integer
On Error Resume Next
'Sizes the user array. Since we run into lock problems if this is
'done automaticly, we will call it from here every minute because
'it is safe.
userResize False, True
'Ups the Actions graph data and reset the current actions ticker
acts(actPos) = actions
http_acts(actPos) = httpActions
httpConnects(actPos) = http_connections
logins_graph(actPos) = userLoginsGraph
If actPos <= LBound(acts) Then
    actPos = UBound(acts)
    Else
        actPos = actPos - 1
        End If
actions = 0: http_connections = 0: httpActions = 0
userLoginsGraph = 0
If actup Then
    Select Case whatToGraph
        Case graph.GRAPH_ACTIONS
            grapher = acts
        Case graph.GRAPH_USER_LOGINS
            grapher = logins_graph
        Case graph.GRAPH_HTTP_CONNECTIONS
            grapher = httpConnects
        Case graph.GRAPH_HTTP_REQUESTS
            grapher = http_acts
        End Select
    g_Paint
    End If
'This will increment all users's time
un = UBound(user): count = 0
Do While count < un
    count = count + 1
    If user(count).operational And user(count).state >= STATE_NORMAL Then
        If user(count).idle = system.maxIdle And user(count).rank <= system.timeoutMaxLevel And system.timeoutMaxLevel > 0 Then
            send BELL & "~FYWake up before you are expunged~RS" & CRLF, count
            ElseIf user(count).idle > system.maxIdle And user(count).rank <= system.timeoutMaxLevel And system.timeoutMaxLevel > 0 Then
                send "~FRTo the sea, from which ye came" & CRLF, count
                killUser count
                un = UBound(user)
                End If
        user(count).timeon = user(count).timeon + 60
        user(count).totalTime = user(count).totalTime + 60
        End If
    If user(count).operational Then
        user(count).idle = user(count).idle + 1
        End If
    If user(count).state < STATE_NORMAL And user(count).idle > 2 Then
        send "Login timeout" & CRLF, count
        removeUser count
        End If
    Loop

'Send keepalive signals through netlinks so they dont hang. Also
'unjam stuck netlinks
For count = 1 To UBound(net)
    If madeNetlinks(count) = True And s2n(count) > -1 Then
        nn = s2n(count)
        If net(nn).state >= NETLINK_UP Then
            netout "KA" & LF, count
            End If
        End If
    Next count

'Mail System
For count = 0 To MAX_MAIL_SLOTS
    If mail(count).inuse Then
        If Not mainForm.mailsock.Connected Then
            mail_out.success = False
            lighter "Mail daemon executing queued message to " & mail(count).u_email
            SMTP_STATE = EMAIL_CONNECTING
            mail_out = mail(count)
            mail(count).inuse = False
            smtp_out
            Exit For
            Else
                mainForm.mailsock.Disconnect
                End If
        End If
    Next count
If tree.SelectedItem.Key Like "UD *" Then
    treeLoad
    End If
End Sub

Private Sub TNConnect_Click()
Dim count As Integer, openSocket As Integer, netnum As Integer, count2 As Integer
mainForm.MousePointer = vbHourglass
netnum = -1
For count = LBound(net) To UBound(net)
    If net(count).name = tree.SelectedItem.text Then
        netnum = count
        Exit For
        End If
    Next count
If netnum = -1 Then
    lighter "That Netlink was not found"
    mainForm.MousePointer = vbDefault
    Exit Sub
    End If
Select Case net(netnum).state
    Case NETLINK_DOWN
        Call connectNetlink(netnum)
    Case NETLINK_CONNECTING
        dropNetlink n2s(netnum)
    Case Is >= NETLINK_VERIFYING
        If n2s(netnum) < 0 Then
            Exit Sub
            End If
        dropNetlink n2s(netnum)
        End Select
mainForm.tree.refresh
mainForm.MousePointer = vbDefault
End Sub

Private Sub TNrefresh_Click()
treeLoad
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "TOOLBAR REBOOT"
        FullReMenu_Click
    Case "SW TOOLS"
        If Not Dir$(App.Path & "Plugins\\tools\tools.exe") = "" Then
            Shell App.Path & "\Plugins\Tools\Tools.exe", vbMaximizedFocus
            Else
                Button.Enabled = False
                End If
    Case "SYSTEM TRAY"
        GoSystemTray
    Case "SW SCRIPT"
        NotepadConfigFile_Click
    Case "BROADCAST"
        If Button.Tag = "" Then
            lighter "Enter a message to broadcast"
            bar.visible = True
            bar.SetFocus
            Button.value = tbrPressed
            Button.Tag = "Pressed"
            Else
                bar.visible = False
                bar.text = ""
                lighter "Broadcast canceled"
                Button.Tag = ""
                End If
        End Select
End Sub

Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.Index > 4 Then
    Re_Click ButtonMenu.Index - 5
    ElseIf ButtonMenu.Index = 1 Then
        RebootReboot_Click
        Else
            FullReMenu_Click
            End If
End Sub

Private Sub TopWindow_Click()
If TopWindow.Checked = True Then
    Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    TopWindow.Checked = False
    Else
        Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
        TopWindow.Checked = True
        End If
End Sub

Private Sub tree_Expand(ByVal node As MSComctlLib.node)
If Not BOOTING Then
    If node.Key = "GRAPHER" Then
        tree.Nodes("USER DATA").Expanded = False
        tree.Nodes("USER HISTORIES").Expanded = False
        If UBound(net) > 0 Then
            tree.Nodes("NETLINKS").Expanded = False
            End If
        tree.Nodes("SERVER_INFO").Expanded = False
        Else
            tree.Nodes("GRAPHER").Expanded = False
            End If
    End If
End Sub

Private Sub tree_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim parental_node As String, count As Integer, un As Integer
Const RIGHT_BUTTON = 2
Me.tree.refresh
If Button = RIGHT_BUTTON Then
'    Select Case tree.SelectedItem.Parent
    If tree.HitTest(x, y) Is Nothing Then
        Exit Sub
        End If
    tree.SelectedItem = tree.HitTest(x, y)
    If tree.SelectedItem.Parent Is Nothing Then
        Exit Sub
        End If
    treeLoad
    Select Case UCase$(tree.SelectedItem.Parent)
        Case "NETLINKS"
            For count = LBound(net) To UBound(net)
                If net(count).name = tree.SelectedItem.text Then
                    Select Case net(count).state
                        Case NETSTATES.NETLINK_DOWN
                            TNConnect.Caption = "Connect"
                        Case NETLINK_UP Or net(count).state = NETLINK_VERIFYING
                            TNConnect.Caption = "Disconnect"
                        Case Else
                            TNConnect.Caption = "Abort"
                            End Select
                    End If
                Next count
            boldPopup mnuTreeNetlink, 4, 0
        Case "USER DATA", "USER HISTORIES"
            un = getUser(tree.SelectedItem)
            If un > 0 Then
                connectionsList.ListIndex = user(un).listing
                If connectionsList.ListIndex > -1 Then
                    setOnlinePopup
                    End If
                Else
                    TUInfo.Enabled = IIf(UCase$(tree.SelectedItem.Parent) = "USER DATA", False, True)
                    user(0).name = tree.SelectedItem.text
                    loadUserData 0
                    mnuOfflineICQ.Enabled = IIf(Val(user(0).ICQ) > 0, True, False)
                    PopupMenu mnuTreeUser
                    End If
        Case "SYSTEMS LOGBOOK"
        End Select
    End If
End Sub

Public Sub tree_NodeClick(ByVal node As MSComctlLib.node)
If node.Parent Is Nothing Then
    hideGrapher
    Else
        If Not node.Parent.Key = "GRAPHER" Then
            hideGrapher
            Else
                showGrapher
                End If
        End If
treeLoad
End Sub

Private Sub TUEraseAccount_Click()
deleteAccount (tree.SelectedItem.text)
End Sub

Private Sub TUInfo_Click()
Dim noder As String
noder = "UD " & UCase$(tree.SelectedItem.text)
tree.SetFocus
tree.Nodes("USER DATA").Expanded = True
tree.Nodes(noder).Selected = True
mainForm.tree_NodeClick tree.Nodes(noder)
End Sub

Private Sub UserfileConversionsFile_Click()
Call Shell(App.Path & "\Tools\Convert\convert.exe", vbNormalFocus)
End Sub

Private Sub UserPopDemote_Click()
user(0).name = "A server administrator"
user(0).rank = 999
Dim count As Integer
For count = 1 To UBound(user)
    If user(count).listing = mainForm.connectionsList.ListIndex Then
    word(1) = user(count).name
        demote 0, user(count).name
        Exit For
        End If
    Next count
End Sub

Private Sub UserPopInfo_Click()
'GUI user has clicked info on the context menu so we will automaticly
'take them to this user's info file.
Dim count As Integer
For count = 1 To UBound(user)
    If user(count).listing = mainForm.connectionsList.ListIndex Then
        If user(count).state > STATE_LOGIN3 Then
            tree.Nodes("USER DATA").Expanded = True
            tree.Nodes("UD " & UCase$(user(count).name)).Selected = True
            mainForm.tree_NodeClick tree.Nodes("UD " & UCase$(user(count).name))
            tree.Nodes("UD " & UCase$(user(count).name)).Selected = True
            End If
        End If
    Next count
End Sub

Private Sub UserPopKill_Click()
'GUI user has clicked kill on the context menu
Dim count As Integer
For count = 1 To UBound(user)
    If user(count).listing = mainForm.connectionsList.ListIndex Then
        If Not user(count).netlinkType Then
            If madeSockets(count) = True Then
                send CRLF & "~LI~BR~FWYou have been removed by the server administrators.~RS" & CRLF, count
                If user(count).state <= STATE_LOGIN3 Then
                    writeSyslog "An administrator removed a login ~FB[~RS" & mainForm.Socket2(count).PeerName & "~FB]"
                    Else
                        writeSyslog "A server administrator has killed ~FB" & user(count).name
                        writeRoom user(count).room, user(count).name & " has been removed by a server administrator" & CRLF
                        End If
                killUser count
                Exit For
                End If
            Else
                send CRLF & "~LI~BR~FWYou have been removed by the server administrators.~RS" & CRLF, count
                killUser count
                Exit For
                End If
        End If
    Next count
End Sub

Private Sub UserPopPickle_Click()
'This is just a little gag. I'm not sure what posessed me to put it in
'here in the first place. Maybe just to add a touch of fun to the whole
'this, which is what talkers are about anyways.
Dim count As Integer
For count = 1 To UBound(user)
    If user(count).listing = mainForm.connectionsList.ListIndex Then
        lighter "Sending a pickle to " & user(count).name
        send "~OL~FGA server administrator has sent you a pickle" & CRLF, count
        Exit For
        End If
    Next count
End Sub

Private Sub UserPopPromote_Click()
'We promote as the GUI by faking the machine :)
user(0).name = "A server administrator"
user(0).rank = 999
Dim count As Integer
For count = 1 To UBound(user)
    If user(count).listing = mainForm.connectionsList.ListIndex Then
    user(0).inpstr = "promote " & user(count).name
    word(1) = user(count).name
        promote 0, user(count).name
        Exit For
        End If
    Next count
End Sub

Private Sub vscroll_Change()
viewerScrollRefresh
End Sub

Private Sub vscroll_GotFocus()
p(0).SetFocus
End Sub

Private Sub vscroll_Scroll()
viewerScrollRefresh
End Sub
