VERSION 5.00
Begin VB.Form splashForm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Scott Lloyd's Windows Talker Server"
   ClientHeight    =   3660
   ClientLeft      =   -15
   ClientTop       =   1110
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "license.frx":0000
   LinkTopic       =   "Form3"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3660
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   $"license.frx":000C
      ForeColor       =   &H000000FF&
      Height          =   2385
      Index           =   2
      Left            =   2790
      TabIndex        =   2
      Top             =   870
      WhatsThisHelpID =   10037
      Width           =   4515
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "©1999 Scott Lloyd ""Scott Lloyd"""
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   5040
      TabIndex        =   3
      Top             =   3360
      WhatsThisHelpID =   10038
      Width           =   2295
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "The only true Windows talker server"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   2820
      TabIndex        =   1
      Top             =   540
      WhatsThisHelpID =   10036
      Width           =   4365
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SouthWest Talker Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   0
      Left            =   2820
      TabIndex        =   0
      Top             =   60
      WhatsThisHelpID =   10035
      Width           =   4485
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   3570
      Left            =   30
      Picture         =   "license.frx":0239
      Top             =   60
      WhatsThisHelpID =   10034
      Width           =   2730
   End
End
Attribute VB_Name = "splashForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub leave()
mainForm.visible = True
If mainForm.Enabled Then
    mainForm.SetFocus
    End If
Unload Me
End Sub

Private Sub Form_Click()
leave
End Sub

Private Sub Form_Load()
If mainForm.TopWindow.Checked Then
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    End If
End Sub

Private Sub Form_LostFocus()
leave
End Sub

Private Sub Image1_Click()
leave
End Sub

Private Sub Label_Click(Index As Integer)
leave
End Sub

