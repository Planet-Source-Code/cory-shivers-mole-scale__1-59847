VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000008&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  About"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   1200
      Top             =   2400
   End
   Begin VB.Timer tmrScroll 
      Interval        =   1
      Left            =   600
      Top             =   2400
   End
   Begin VB.Frame FrameAbout 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   -120
      TabIndex        =   0
      Top             =   -120
      Width           =   4815
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Mole Scale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.TextBox txtAbout 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   5655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3000
      Width           =   3255
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
tmrScroll.Enabled = False
txtAbout.Top = 3000
frmAbout.Visible = False
End Sub

Private Sub Form_Load()
txtAbout.Text = " .:: Mole Scale ::. " & vbCrLf & _
 ".:: Version: " & App.Major & "." & App.Minor & " ::." & vbCrLf & _
".:: Build: " & App.Revision & " ::." & vbCrLf & vbCrLf & _
".:: Made by Nate & Cory ::." & vbCrLf & vbCrLf & _
".:: We'd Like to give a shout out to ::." & vbCrLf & vbCrLf & _
".:: Mr. Smith without him giving us ::." & vbCrLf & vbCrLf & _
".:: Our knowledge this wouldnt be possible! ::."
txtAbout.SelStart = 1000
tmrScroll.Enabled = False
tmrScroll.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
tmrScroll.Enabled = False
txtAbout.Top = 3000
frmAbout.Visible = False
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub

Private Sub tmrScroll_Timer()
On Error Resume Next

If txtAbout.Top <= -5000 Then
    txtAbout.Top = 3000
    Exit Sub
End If

txtAbout.Top = txtAbout.Top - 4

End Sub
