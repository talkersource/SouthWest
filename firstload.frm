VERSION 5.00
Begin VB.Form patcherForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SouthWest Autopatch"
   ClientHeight    =   1425
   ClientLeft      =   495
   ClientTop       =   735
   ClientWidth     =   5475
   Icon            =   "firstload.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3540
      Top             =   30
   End
   Begin VB.PictureBox Progress_Bar 
      AutoRedraw      =   -1  'True
      FillColor       =   &H8000000F&
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1860
      ScaleHeight     =   195
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   1020
      WhatsThisHelpID =   10165
      Width           =   3495
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   1680
         TabIndex        =   5
         Top             =   0
         Width           =   75
      End
   End
   Begin VB.Label setupName 
      Height          =   195
      Left            =   1860
      TabIndex        =   6
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label cop 
      Caption         =   "Initializing"
      Height          =   195
      Left            =   2550
      TabIndex        =   4
      Top             =   750
      WhatsThisHelpID =   10167
      Width           =   2805
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Process:"
      Height          =   195
      Index           =   1
      Left            =   1860
      TabIndex        =   3
      Top             =   750
      WhatsThisHelpID =   10166
      Width           =   615
   End
   Begin VB.Label Label 
      Caption         =   "SouthWest File Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   0
      WhatsThisHelpID =   10164
      Width           =   5295
   End
   Begin VB.Label Label 
      Caption         =   "SouthWest has detected a new patch file and will now install it and update the server."
      Height          =   795
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      WhatsThisHelpID =   10163
      Width           =   1575
   End
End
Attribute VB_Name = "patcherForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Form_Load()
Dim hSysMenu As Long, nCnt As Long
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
hSysMenu = GetSystemMenu(Me.hwnd, False)
If hSysMenu Then
    nCnt = GetMenuItemCount(hSysMenu)
    If nCnt Then
        RemoveMenu hSysMenu, nCnt - 1, MF_BYPOSITION Or MF_REMOVE
        RemoveMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_REMOVE
        End If
    End If
Me.show
apEngine
End Sub
