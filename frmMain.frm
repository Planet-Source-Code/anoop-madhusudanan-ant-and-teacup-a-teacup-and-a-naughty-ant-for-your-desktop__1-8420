VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   585
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   735
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   585
   ScaleWidth      =   735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDown 
      Height          =   435
      Left            =   1890
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   375
      ScaleWidth      =   405
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picM2 
      Height          =   435
      Left            =   510
      Picture         =   "frmMain.frx":059A
      ScaleHeight     =   375
      ScaleWidth      =   405
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picM1 
      Height          =   435
      Left            =   540
      Picture         =   "frmMain.frx":08A4
      ScaleHeight     =   375
      ScaleWidth      =   405
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   5
      Left            =   1740
      Picture         =   "frmMain.frx":0BAE
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   720
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   4
      Left            =   1740
      Picture         =   "frmMain.frx":0E3E
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   120
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   3
      Left            =   1110
      Picture         =   "frmMain.frx":10CE
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   1980
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   2
      Left            =   1110
      Picture         =   "frmMain.frx":135E
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   1350
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   1
      Left            =   1110
      Picture         =   "frmMain.frx":15EE
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   720
      Width           =   555
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Index           =   0
      Left            =   1110
      Picture         =   "frmMain.frx":187E
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   150
      Width           =   555
   End
   Begin VB.Timer timMain 
      Interval        =   500
      Left            =   360
      Top             =   390
   End
   Begin TC.TransParentCtl tp 
      Height          =   765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1349
      MaskColor       =   16711680
      MouseIcon       =   "frmMain.frx":1B0E
      MousePointer    =   99
   End
End
Attribute VB_Name = "frmMain"
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
Public IfDown As Boolean

'For moving the form when user drags the cup
Dim MbMov As Boolean
Dim PrevX As Long, PrevY As Long
Dim TeaTop, TeaLeft




Private Sub Form_Load()
Me.Show

IfDown = False

Load frmDust
frmDust.Top = -(frmDust.Height)

frmDust.Show



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


End Sub


Private Sub Form_Paint()

SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left / 15, _
                        Me.Top / 15, Me.Width / 15, _
                        Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW

End Sub

Private Sub timMain_Timer()

'Incrementing the counter and drawing the cup
CurPos = CurPos + 1
If CurPos = MaxNo Then CurPos = 0
Set tp.MaskPicture = pic(CurPos)
End Sub

Private Sub tp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'User starts to drag the cup
If Button = 2 Then
frmMenu.mnuUp.Enabled = IfDown
PopupMenu frmMenu.mnuFile
Exit Sub
End If

MbMov = True
PrevX = X
PrevY = Y
Set tp.MouseIcon = picM2.Picture

End Sub

Private Sub tp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'User is moving the cup

If MbMov Then Me.Move (Me.Left + X - PrevX), (Me.Top + Y - PrevY)
End Sub

Private Sub tp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'User stopped dragging the cup

MbMov = False
Set tp.MouseIcon = picM1.Picture

If IfDown Then
frmTea.Visible = False
Set frmT = New frmTea

frmT.Visible = False

frmT.Top = TeaTop
frmT.Left = TeaLeft

frmT.Visible = True


TeaTop = Me.Top
TeaLeft = Me.Left
End If

End Sub

Public Sub AntIn()
Set tp.MaskPicture = picDown.Picture
timMain.Enabled = False
IfDown = True
TeaTop = Me.Top
TeaLeft = Me.Left
End Sub
