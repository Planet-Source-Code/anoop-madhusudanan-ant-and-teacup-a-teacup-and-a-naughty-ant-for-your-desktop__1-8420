VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tea Cup"
   ClientHeight    =   3315
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMain 
      Caption         =   "Don't show this at startup"
      Height          =   525
      Left            =   60
      TabIndex        =   7
      Top             =   2730
      Width           =   2145
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   465
      Left            =   2910
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "&Done"
      Height          =   465
      Left            =   4260
      TabIndex        =   3
      Top             =   2760
      Width           =   1245
   End
   Begin VB.PictureBox picMain 
      Height          =   2565
      Left            =   60
      ScaleHeight     =   2505
      ScaleWidth      =   5385
      TabIndex        =   0
      Top             =   60
      Width           =   5445
      Begin VB.PictureBox picSub 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   0
         ScaleHeight     =   2625
         ScaleWidth      =   1305
         TabIndex        =   1
         Top             =   -30
         Width           =   1305
         Begin VB.PictureBox pic 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   525
            Index           =   0
            Left            =   300
            Picture         =   "frmMenu.frx":0000
            ScaleHeight     =   525
            ScaleWidth      =   495
            TabIndex        =   2
            Top             =   210
            Width           =   495
         End
      End
      Begin VB.Label Label2 
         Caption         =   $"frmMenu.frx":030A
         Height          =   915
         Left            =   1410
         TabIndex        =   5
         Top             =   990
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Tea Cup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1410
         TabIndex        =   4
         Top             =   540
         Width           =   1095
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuUp 
         Caption         =   "&Upright Cup"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear Teamarks"
      End
      Begin VB.Menu mnuFB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Tea Cup .."
      End
      Begin VB.Menu mnuAuthor 
         Caption         =   "&About Author"
      End
      Begin VB.Menu mnuFB2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================================================================
'
' Developed by Anoop. M
' anoopj12 @ yahoo.com
'
' Anoop M, Govindanikethan, Nedumkunnam P.O, Kottayam,
' Kerala, India - 686 542
'
' Hey sir, Kindly rate this code, if you like it.
'
' READ THIS BEFORE USING THE CODE:
'
' You can study and view the source code for creating your
' own apps, but do not reproduce/release Icon Hunter fully
' or partially for any commercial and/or personal purposes. All
' rights of this product is related to it's author. Any violation
' of above conditions will be treated seriously and is punishable.
'
' I recently inveted a technology for streaming audio, and is
' now looking promoters/investors to invest in a web-phone network
' project.
'
' VISIT MY WEBSITE : http://www.geocities.com/streamingaudio for details
'=============================================================================================================================

Private Sub chkMain_Click()
SaveSetting "Time Tech", "TeaCup", "IfStart", chkMain.Value
End Sub

Private Sub cmdHelp_Click()
frmHelp.Show vbModal
End Sub

Private Sub cmdMain_Click()
Unload Me
End Sub

Private Sub Form_Load()
X = GetSetting("Time Tech", "TeaCup", "IfStart", "0")
chkMain.Value = X
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdHelp.Visible = False

End Sub

Private Sub mnuAbout_Click()

frmMenu.Show

End Sub

Private Sub mnuAuthor_Click()
frmAbout.Show
End Sub

Private Sub mnuClear_Click()
Dim frmL As Form

For Each frmL In Forms
If frmL.Name = "frmTea" Then
Unload frmL
End If
Next frmL

End Sub

Private Sub mnuExit_Click()
Unload frmMain
End
End Sub

Private Sub mnuHelp_Click()
frmHelp.Show
End Sub

Private Sub mnuUp_Click()
mnuClear_Click
frmMain.IfDown = False
frmMain.timMain.Enabled = True
frmDust.Show
End Sub

Private Sub pic_Click(Index As Integer)
Dim msg As String
msg = "Tea Cup:" + vbCrLf + "A simple program for nothing..." + vbCrLf + vbCrLf + "Created by Anoop M, Software Engineer, Time Technologies" + vbCrLf + "Contact anoopj12@angelfire.com"
MsgBox msg, vbOKOnly + vbInformation, "Author"

End Sub
