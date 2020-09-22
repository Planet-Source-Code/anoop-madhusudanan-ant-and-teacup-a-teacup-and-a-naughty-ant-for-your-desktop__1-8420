VERSION 5.00
Begin VB.UserControl TransParentCtl 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "TransParentCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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


'Event Declarations:
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
'Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
'Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Get MaskPicture() As Picture
'Get the mask picture

    Set MaskPicture = UserControl.MaskPicture
End Property

Public Property Set MaskPicture(ByVal picNew As Picture)
'Set the mask picture

Set UserControl.MaskPicture = picNew
'Using Refresh() code before the Set Picture may have good results
Me.Refresh
Set UserControl.Picture = picNew

'Raising the event
PropertyChanged "MaskPicture"
End Property



Public Property Get MaskColor() As OLE_COLOR

    MaskColor = UserControl.MaskColor
End Property



Public Property Let MaskColor(ByVal clrMaskColor As OLE_COLOR)

    UserControl.MaskColor = clrMaskColor
    Me.Refresh
    PropertyChanged "MaskColor"
End Property

'Refresh() to changed the container region with usercontrol's


Public Sub Refresh()

    'On Local Error Resume Next
    Dim hRgnNormal As Long


    With UserControl


        If .MaskPicture = 0 Then
            hRgnNormal = CreateRectRgn(0, 0, .ScaleX(.Width), .ScaleY(.Height))
            SetWindowRgn .Extender.Container.hWnd, hRgnNormal, True
        Else
            .Size .ScaleX(.MaskPicture.Width), .ScaleY(.MaskPicture.Height)
            .Extender.Container.Width = .Width
            .Extender.Container.Height = .Height
            .Extender.Move 0, 0
            
            'Give system the time to finish the special regions created

            DoEvents
                'Set New Regions
                SetWindowRgn .Extender.Container.hWnd, Me.hRgn, True


                If Err Then
                    MsgBox "The Container not support the mothods"
                End If

        End If

    End With

End Sub



Public Property Get hRgn() As OLE_HANDLE

    hRgn = CreateRectRgn(0, 0, 1, 1)
    GetWindowRgn UserControl.hWnd, hRgn
End Property


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'Persist the control's property

    Me.MaskColor = PropBag.ReadProperty("MaskColor", &H8000000F)
    Set Me.MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "MaskColor", Me.MaskColor, &H8000000F
    PropBag.WriteProperty "MaskPicture", Me.MaskPicture, Nothing
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
'
'Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseMove(Button, Shift, X, Y)
'End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

