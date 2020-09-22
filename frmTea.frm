VERSION 5.00
Begin VB.Form frmTea 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   90
   ClientLeft      =   1920
   ClientTop       =   -1650
   ClientWidth     =   90
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   2
      Left            =   1110
      Picture         =   "frmTea.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   1350
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   1
      Left            =   1110
      Picture         =   "frmTea.frx":0282
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   720
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   0
      Left            =   1110
      Picture         =   "frmTea.frx":0504
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   150
      Width           =   555
   End
   Begin VB.Timer timMain 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   360
      Top             =   390
   End
   Begin TC.TransParentCtl tp 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   661
      MaskColor       =   16711680
   End
End
Attribute VB_Name = "frmTea"
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
Randomize (2)
CurPos = (Rnd * 2) / 1
MaxNo = pic.Count

MbMov = False

'Default Picture
Set tp.MaskPicture = pic(CurPos)


End Sub


Private Sub Form_Paint()

SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left / 15, _
                        Me.Top / 15, Me.Width / 15, _
                        Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW

End Sub

Private Sub timMain_Timer()
frmDust.Show

'Incrementing the counter and drawing the cup
CurPos = CurPos + 1
If CurPos = MaxNo Then CurPos = 0
Set tp.MaskPicture = pic(CurPos)
End Sub

Private Sub tp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
frmMenu.mnuUp.Enabled = IfDown
PopupMenu frmMenu.mnuFile
Exit Sub
End If
End Sub

