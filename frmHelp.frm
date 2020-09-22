VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Teacup Help"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMain 
      Height          =   2565
      Left            =   30
      ScaleHeight     =   2505
      ScaleWidth      =   5385
      TabIndex        =   1
      Top             =   60
      Width           =   5445
      Begin VB.PictureBox picSub 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   0
         ScaleHeight     =   2625
         ScaleWidth      =   1305
         TabIndex        =   2
         Top             =   -30
         Width           =   1305
         Begin VB.PictureBox pic 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   525
            Index           =   0
            Left            =   300
            Picture         =   "frmHelp.frx":0000
            ScaleHeight     =   525
            ScaleWidth      =   495
            TabIndex        =   3
            Top             =   210
            Width           =   495
         End
      End
      Begin VB.Label Label1 
         Caption         =   $"frmHelp.frx":030A
         Height          =   1275
         Left            =   1440
         TabIndex        =   5
         Top             =   1170
         Width           =   3885
      End
      Begin VB.Label Label2 
         Caption         =   $"frmHelp.frx":0451
         Height          =   855
         Left            =   1470
         TabIndex        =   4
         Top             =   120
         Width           =   3795
      End
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "&Close"
      Height          =   465
      Left            =   4230
      TabIndex        =   0
      Top             =   2730
      Width           =   1245
   End
End
Attribute VB_Name = "frmHelp"
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

Private Sub cmdMain_Click()
Unload Me
End Sub
