VERSION 5.00
Begin VB.Form eggForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   615
   ClientLeft      =   2115
   ClientTop       =   1440
   ClientWidth     =   1950
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
   Icon            =   "egg.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   615
   ScaleWidth      =   1950
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Scott"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   570
      TabIndex        =   0
      Top             =   60
      WhatsThisHelpID =   10032
      Width           =   1245
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   60
      Picture         =   "egg.frx":000C
      Top             =   60
      WhatsThisHelpID =   10031
      Width           =   480
   End
End
Attribute VB_Name = "eggForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub egg()
If eggForm.Left = mainForm.Left And eggForm.Top = mainForm.Top Then
    eggForm.Left = mainForm.Left + mainForm.Width - eggForm.Width
    done% = 1
    End If
If eggForm.Left = mainForm.Left + mainForm.Width - eggForm.Width And eggForm.Top = mainForm.Top And done% = 0 Then
    eggForm.Top = mainForm.Top + mainForm.Height - eggForm.Height
    done% = 1
    End If
If eggForm.Top = mainForm.Top + mainForm.Height - eggForm.Height And eggForm.Left = mainForm.Left + mainForm.Width - eggForm.Width And done% = 0 Then
    eggForm.Left = mainForm.Left
    done% = 1
    End If
If eggForm.Left = mainForm.Left And eggForm.Top = mainForm.Top + mainForm.Height - eggForm.Height And done% = 0 Then
    eggForm.Top = mainForm.Top
    End If
done% = 0
End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
If mainForm.TopWindow.Checked Then
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    End If
End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
egg
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
egg
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
egg
End Sub

