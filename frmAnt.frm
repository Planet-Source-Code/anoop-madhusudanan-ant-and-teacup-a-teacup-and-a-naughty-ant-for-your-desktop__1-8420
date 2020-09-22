VERSION 5.00
Begin VB.Form frmAnt 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   885
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   3
      Left            =   1410
      Picture         =   "frmAnt.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   1980
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   2
      Left            =   1410
      Picture         =   "frmAnt.frx":0290
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   1350
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   1
      Left            =   1410
      Picture         =   "frmAnt.frx":0520
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   720
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   0
      Left            =   1410
      Picture         =   "frmAnt.frx":07B0
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   150
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   4
      Left            =   2040
      Picture         =   "frmAnt.frx":0A40
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   150
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   5
      Left            =   2070
      Picture         =   "frmAnt.frx":0CD0
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   690
      Width           =   555
   End
   Begin VB.Timer timMain 
      Interval        =   100
      Left            =   0
      Top             =   330
   End
   Begin CF.TransParentCtl tp 
      Height          =   615
      Left            =   300
      TabIndex        =   6
      Top             =   0
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   1085
      MaskColor       =   16711680
   End
End
Attribute VB_Name = "frmAnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================================================
'frmMain.frm
'
'The 'Fly Form'.
'
'Created by Anoop. M, Software Manager, Time Technologies
'
'In case of doubts, contact anoopm@vsnl.com
'==========================================================================================

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
frmAnt.Top = frmAnt.Top + 60


If frmAnt.Left > (frmMain.Left - frmAnt.Width \ 2) And frmAnt.Left < (frmMain.Left + frmMain.Width) Then
If frmAnt.Top > frmMain.Top - (frmAnt.Height \ 2) And frmAnt.Top < frmMain.Top + (frmMain.Height + 100) Then
timMain.Enabled = False
frmAnt.Hide
Unload frmAnt
frmMain.AntIn
Exit Sub
End If
End If

If frmAnt.Top > Screen.Height + frmAnt.Height Then
frmAnt.Hide
frmAnt.Left = frmMain.Left
frmAnt.Top = -(frmAnt.Height)
frmAnt.Show
End If

CurPos = CurPos + 1
If CurPos = MaxNo Then CurPos = 0
Set tp.MaskPicture = pic(CurPos)
End Sub



