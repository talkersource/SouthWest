VERSION 5.00
Begin VB.Form aboutForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About - Scott Lloyd's Windows Talker Server"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5310
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "about.frx":000C
   ScaleHeight     =   2850
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   4860
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   1155
      Left            =   30
      Picture         =   "about.frx":0316
      ScaleHeight     =   1095
      ScaleWidth      =   765
      TabIndex        =   3
      Top             =   60
      WhatsThisHelpID =   10155
      Width           =   825
   End
   Begin VB.PictureBox CreditsForm 
      Height          =   1905
      Left            =   2070
      ScaleHeight     =   1845
      ScaleWidth      =   3135
      TabIndex        =   2
      Top             =   900
      WhatsThisHelpID =   10154
      Width           =   3195
   End
   Begin VB.Timer ScrollCredits 
      Interval        =   10
      Left            =   4440
      Top             =   0
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "southwest@talker.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   60
      MouseIcon       =   "about.frx":20540
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "mailto:southwest@talker.com"
      Top             =   1650
      Width           =   1620
   End
   Begin VB.Label Label2 
      Caption         =   $"about.frx":2084A
      Height          =   1185
      Left            =   60
      TabIndex        =   8
      Top             =   1260
      Width           =   1965
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SouthWest Programmer"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   2610
      WhatsThisHelpID =   10159
      Width           =   1680
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scott Lloyd"
      Height          =   225
      Left            =   360
      TabIndex        =   5
      Top             =   2430
      WhatsThisHelpID =   10158
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "Well, here it is, the talker server that they've all been talking about. I hope you enjoy"
      Height          =   1185
      Left            =   900
      TabIndex        =   4
      Top             =   60
      WhatsThisHelpID =   10156
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The only true Windows talker server"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   630
      WhatsThisHelpID =   10153
      Width           =   3060
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SouthWest"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   2070
      TabIndex        =   0
      Top             =   30
      WhatsThisHelpID =   10151
      Width           =   3105
   End
End
Attribute VB_Name = "aboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FlashWindow& Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long)
Dim creditsScroll As Integer

Private Sub c()
CreditsForm.CurrentX = 150
End Sub

Private Sub Form_Deactivate()
aboutForm.visible = False
Unload aboutForm
End Sub

Private Sub Form_Load()
creditsScroll = CreditsForm.ScaleHeight - TextHeight("SouthWest")
If App.Revision = 0 Then
    Me.Caption = "About - SouthWest v" & App.Major & "." & App.Minor
    Else
        Me.Caption = "About - SouthWest v" & App.Major & "." & App.Minor & "." & App.Revision
        End If
If mainForm.TopWindow.Checked Then
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    End If
End Sub

Private Sub Label7_Click()
Label7.ForeColor = &HC000C0
Call ShellExecute(Me.hwnd, vbNullString, "mailto:southwest@talker.com", vbNullString, "c:\", 1)
End Sub

Private Sub ScrollCredits_Timer()
With CreditsForm
    .AutoRedraw = True
    .FontSize = 8
    creditsScroll = creditsScroll - 8
    .Cls
    .CurrentY = creditsScroll
    .CurrentX = CreditsForm.ScaleWidth / 2 - TextWidth("SouthWest Credits") / 2
    .FontBold = True
    CreditsForm.Print "SouthWest Credits"
    .FontBold = False
    CreditsForm.Print "Programming": c
    CreditsForm.Print "Scott Lloyd"
    CreditsForm.Print "Beta Testers": c
    CreditsForm.Print "Rick Szajkowski": c
    CreditsForm.Print "Krysia Kwiatkowski"
    CreditsForm.Print "Special Thanks": c
    CreditsForm.Print "The fine people at Talker.com": c
    CreditsForm.Print "Joan Stark": c
    CreditsForm.Print "Neil Robertson": c
    CreditsForm.Print "Andrew Collington": c
    CreditsForm.Print "David Gatwood": c
    CreditsForm.Print
    .FontBold = True
    .ForeColor = &HFF00FF
    .CurrentX = .Width / 2 - .TextWidth("To Krysia") / 2
    CreditsForm.Print "To Krysia": c
    .CurrentX = .Width / 2 - .TextWidth("Whom shall be forever in my heart") / 2
    CreditsForm.Print "Whom shall be forever in my heart"
    .FontBold = False
    .ForeColor = &H80000012
    End With
    If creditsScroll <= (0 - CreditsForm.ScaleHeight) - 1100 Then
        creditsScroll = CreditsForm.ScaleHeight
        End If
    If creditsScroll < -2400 Then
        ScrollCredits.Enabled = False
        End If
End Sub

Private Sub Timer1_Timer()
FlashWindow Me.hwnd, -1
End Sub
