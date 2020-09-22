VERSION 5.00
Begin VB.Form frmDust 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1800
   ClientLeft      =   1245
   ClientTop       =   -1650
   ClientWidth     =   1410
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   1410
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   2
      Left            =   690
      Picture         =   "frmDust.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   1230
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   1
      Left            =   660
      Picture         =   "frmDust.frx":0290
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   660
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   0
      Left            =   660
      Picture         =   "frmDust.frx":0520
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   90
      Width           =   555
   End
   Begin VB.Timer timMain 
      Interval        =   100
      Left            =   -270
      Top             =   330
   End
   Begin TC.TransParentCtl tp 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   1085
      MaskColor       =   16711680
   End
End
Attribute VB_Name = "frmDust"
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

'Current position of counter, Maximum counter value
Dim CurPos As Integer, MaxNo As Integer
Dim IfDown As Boolean


Private Sub Form_Load()
IfDown = False

'This will set the form on top of other forms
SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left / 15, _
                        Me.Top / 15, Me.Width / 15, _
                        Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW

'Counter initialization

CurPos = 0

MaxNo = pic.Count

MbMov = False

'Default Picture
Set tp.MaskPicture = pic(CurPos)
Me.Left = frmMain.Left

End Sub


Private Sub Form_Paint()

SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left / 15, _
                        Me.Top / 15, Me.Width / 15, _
                        Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW

End Sub

Private Sub timMain_Timer()
frmDust.Top = frmDust.Top + 60


If frmDust.Left > (frmMain.Left - frmDust.Width \ 2) And frmDust.Left < (frmMain.Left + frmMain.Width) Then
If frmDust.Top > frmMain.Top - (frmDust.Height \ 2) And frmDust.Top < frmMain.Top + (frmMain.Height + 100) Then
timMain.Enabled = False
frmDust.Hide
Unload frmDust
frmMain.AntIn
Exit Sub
End If
End If

If frmDust.Top > Screen.Height + frmDust.Height Then
frmDust.Visible = False
frmDust.Left = frmMain.Left
frmDust.Top = -(frmDust.Height)
frmDust.Visible = True
End If

CurPos = CurPos + 1
If CurPos = MaxNo Then CurPos = 0
Set tp.MaskPicture = pic(CurPos)
End Sub

