VERSION 5.00
Begin VB.UserControl XPGUI 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1320
   ControlContainer=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   88
   Begin VB.Timer tmrWindow 
      Interval        =   10
      Left            =   450
      Top             =   375
   End
   Begin VB.Image imgActiveNewTitlebar 
      Height          =   315
      Index           =   2
      Left            =   405
      Top             =   975
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgActiveNewTitlebar 
      Height          =   315
      Index           =   1
      Left            =   330
      Stretch         =   -1  'True
      Top             =   975
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image imgActiveNewTitlebar 
      Height          =   315
      Index           =   0
      Left            =   0
      Top             =   975
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgDisNewTitlebar 
      Height          =   315
      Index           =   2
      Left            =   405
      Picture         =   "XPGUI.ctx":0000
      Top             =   1350
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDisNewTitlebar 
      Height          =   315
      Index           =   1
      Left            =   330
      Picture         =   "XPGUI.ctx":04D9
      Stretch         =   -1  'True
      Top             =   1350
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image imgDisNewTitlebar 
      Height          =   315
      Index           =   0
      Left            =   0
      Picture         =   "XPGUI.ctx":066B
      Top             =   1350
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgNewLeft 
      Height          =   1620
      Index           =   1
      Left            =   1275
      Picture         =   "XPGUI.ctx":0AE0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgNewRight 
      Height          =   1620
      Index           =   0
      Left            =   975
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgNewBottom 
      Height          =   45
      Index           =   0
      Left            =   -75
      Stretch         =   -1  'True
      Top             =   1725
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image imgNewRight 
      Height          =   1620
      Index           =   1
      Left            =   1050
      Picture         =   "XPGUI.ctx":0B79
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgNewBottom 
      Height          =   45
      Index           =   1
      Left            =   -75
      Picture         =   "XPGUI.ctx":0C0E
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image imgNewLeft 
      Height          =   1620
      Index           =   0
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgExit 
      Height          =   195
      Left            =   630
      Picture         =   "XPGUI.ctx":0CA3
      Top             =   75
      Width           =   195
   End
   Begin VB.Image imgNewX 
      Height          =   195
      Index           =   0
      Left            =   45
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNewX 
      Height          =   195
      Index           =   2
      Left            =   270
      Picture         =   "XPGUI.ctx":0EED
      Top             =   375
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNewX 
      Height          =   195
      Index           =   1
      Left            =   45
      Picture         =   "XPGUI.ctx":1137
      Top             =   375
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNewX 
      Height          =   195
      Index           =   3
      Left            =   270
      Picture         =   "XPGUI.ctx":1381
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgRight 
      Height          =   570
      Left            =   0
      Picture         =   "XPGUI.ctx":15CB
      Stretch         =   -1  'True
      Top             =   315
      Width           =   45
   End
   Begin VB.Image imgLeft 
      Height          =   570
      Left            =   855
      Picture         =   "XPGUI.ctx":1687
      Stretch         =   -1  'True
      Top             =   315
      Width           =   45
   End
   Begin VB.Image imgBottom 
      Height          =   45
      Left            =   0
      Picture         =   "XPGUI.ctx":1721
      Stretch         =   -1  'True
      Top             =   885
      Width           =   900
   End
   Begin VB.Label lblTitleBar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XPGUI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   525
   End
   Begin VB.Image imgTopLeft 
      Height          =   315
      Left            =   0
      Picture         =   "XPGUI.ctx":17CF
      Top             =   0
      Width           =   330
   End
   Begin VB.Image imgTopMiddle 
      Height          =   315
      Left            =   330
      Picture         =   "XPGUI.ctx":19E3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   75
   End
   Begin VB.Image imgTopRight 
      Height          =   315
      Left            =   410
      Picture         =   "XPGUI.ctx":1B75
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "XPGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Original Versions 1.0.0 [04/1/03]
'Brought to you and written by Mike Canejo
'-----------------------------------------
'AOL/AIM: Mikey3dd
'Email: MikeCanejo@hotmail.com
'-----------------------------------------
'Comments:
'-----------------------------------------
'ActiveXP is a activex control which immatates
'the XP gui with borderstyle 4 & 5. By design it
'will not resize, i'm sure it wouldn't be hard
'to impliment, the best way would be by "blocks"
'to prevent flickering. Example is Winamp 2.0's
'playlist.

'The code is very straight forward, slightly
'commented because of it. The control + graphics
'were based on my Pricewatch Navigator program
'which is also included in this download.

'Oh, if you find anything useful in my code or think it
'would be worth your time, please goto:
'http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=44416&lngWId=1
'and vote, not much to ask ;)
'-----------------------------------------
'DISCLAIMER: You are free to distribute/modify the
'source code but please leave the comments located
'in all the modules declaration sections intact.
'That's all I ask :)
'-----------------------------------------

Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Dim intTwipX        As Integer
Dim intTwipY        As Integer
Dim strCaption      As String
Dim blnDown         As Boolean
Dim blnOver         As Boolean
Dim blnActive       As Boolean
Dim blnOnce(1)      As Boolean
Dim blnAutoSize     As Boolean
Dim blnControlBox   As Boolean

Event Click()
Event DblClick()
Event Resize()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event ExitButton(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Get ControlBox() As Boolean

    ControlBox = blnControlBox
    
End Property

Public Property Let ControlBox(ByVal blnBoxChange As Boolean)

    blnControlBox = blnBoxChange
    PropertyChanged "ControlBox"
    
End Property

Public Property Get Caption() As String
    
    Caption = strCaption
    
End Property

Public Property Let Caption(ByVal strNewCap As String)

    strCaption = strNewCap
    lblTitleBar.Caption = strNewCap
    PropertyChanged "Caption"
    
End Property

Public Property Get AutoSize() As Boolean

    AutoSize = blnAutoSize
    
End Property

Public Property Let AutoSize(blnSizeChange As Boolean)

    'Auto sizes to a form's dimensions.
    'Ingenius.
    
    blnAutoSize = blnSizeChange
    UserControl_Resize
    PropertyChanged "AutoSize"
    
End Property

Public Property Get ActiveWindow() As Boolean

    ActiveWindow = blnActive
    
End Property

Public Property Let ActiveWindow(blnActiveChange As Boolean)
    
    blnActive = blnActiveChange
    
    If blnActive Then
        tmrWindow.Interval = 10
    Else
        tmrWindow.Interval = 0
    End If
    
    PropertyChanged "ActiveWindow"
    
End Property

Private Sub FormDrag(TheForm As Form)

    ActiveNow
    
    DoEvents 'Added doevents to let graphics get a chance to update
    
    ReleaseCapture
    SendMessage TheForm.hwnd, &HA1, 2, 0&
    
End Sub

Private Sub imgBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        imgExit = imgNewX(1)
        blnDown = True
    End If

End Sub

Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Gets mouse move coordinates to make sure the
    'arrow is over the exit button when clicked.
    
    If Button = 1 Then
        
        If X / intTwipX > 0 And X / intTwipX < imgExit.Width _
        And Y / intTwipY > 0 And Y / intTwipY < imgExit.Height Then
            
            If imgExit.Tag = "" _
            Or imgExit.Tag = "2" Then
                imgExit.Tag = "1"
                imgExit = imgNewX(1)
            End If
        
        Else
            
            If imgExit.Tag = "" _
            Or imgExit.Tag = "1" Then
                imgExit.Tag = "2"
                imgExit = imgNewX(0)
            End If
        
        End If
    
    Else
        
        If blnOver = False Then
            blnOver = True
            imgExit = imgNewX(2)
        End If
        
    End If
    
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
    
        blnDown = False
        
        If X / intTwipX >= 0 And X / intTwipX < imgExit.Width _
        And Y / intTwipY >= 0 And Y / intTwipY < imgExit.Height Then
            RaiseEvent ExitButton(Button, Shift, X, Y)
        End If
    
    End If

End Sub

Private Sub imgLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub imgRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub imgTopLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub imgTopMiddle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub imgTopRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub lblTitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub UserControl_Click()

    RaiseEvent Click
    
End Sub

Private Sub UserControl_DblClick()

    RaiseEvent DblClick
    
End Sub

Private Sub UserControl_Initialize()

    intTwipX = Screen.TwipsPerPixelX
    intTwipY = Screen.TwipsPerPixelY
    SetupImages
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)
    
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)
    
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    
    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    On Error Resume Next
    'Used to read settings saved.
    AutoSize = PropBag.ReadProperty("AutoSize", False)
    ActiveWindow = PropBag.ReadProperty("ActiveWindow", True)
    
    ControlBox = PropBag.ReadProperty("ControlBox", True)
    Caption = PropBag.ReadProperty("Caption", "ActiveXP by Mikey3dd")

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    On Error Resume Next
    'Used to write settings saved.
    PropBag.WriteProperty "AutoSize", blnAutoSize, False
    PropBag.WriteProperty "ActiveWindow", blnActive, True
    PropBag.WriteProperty "ControlBox", blnControlBox, True
    PropBag.WriteProperty "Caption", strCaption, "ActiveXP by Mikey3dd"

End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    'Tells the control how to conform on the
    'parent form when it's loaded/resized.
       
    Dim lngPos As Long

    RaiseEvent Resize
    
    If blnAutoSize Then
        Width = UserControl.Parent.Width
        Height = UserControl.Parent.Height
    End If
    
    imgTopRight.Top = 0
    imgTopLeft.Left = 0
    imgTopLeft.Top = 0
    imgTopMiddle.Left = 22
    imgTopMiddle.Top = 0
    lngPos = ScaleWidth - imgTopMiddle.Left - imgTopRight.Width
    imgTopRight.Left = ScaleWidth - imgTopRight.Width
    imgTopMiddle.Width = lngPos
    
    imgLeft.Top = 21
    imgLeft.Left = 0
    imgLeft.Height = ScaleHeight - imgBottom.Height - 20
    imgRight.Top = 21
    imgRight.Left = ScaleWidth - imgRight.Width
    imgRight.Height = ScaleHeight - imgBottom.Height - 20
    imgBottom.Left = 0
    imgBottom.Top = ScaleHeight - imgBottom.Height
    imgBottom.Width = ScaleWidth
    
    imgExit.Top = 5
    imgExit.Left = ScaleWidth - 18
    
End Sub

Private Sub UserControl_Show()

    'Start timer when control is loaded at runtime
    
    If (GetActiveWindow = UserControl.Parent.hwnd) _
    And blnActive = True Then
        
        tmrWindow.Enabled = True
        
        If blnAutoSize Then
            Width = UserControl.Parent.Width
            Height = UserControl.Parent.Height
        End If
    
    Else
    
        tmrWindow.Enabled = False
        
    End If

End Sub

Private Sub UserControl_Terminate()

    Erase blnOnce()
    
End Sub

Private Sub ActiveNow()

    'Since i had to call this same piece of
    'code more than once, i made a sub for it.
    'It updates the form with the "active window"
    'state graphics.
    
    If blnDown = False Then
        imgExit = imgNewX(0)
    End If
    
    imgLeft = imgNewLeft(0)
    imgRight = imgNewRight(0)
    imgBottom = imgNewBottom(0)
    imgTopLeft = imgActiveNewTitlebar(0)
    imgTopMiddle = imgActiveNewTitlebar(1)
    imgTopRight = imgActiveNewTitlebar(2)
    lblTitleBar.ForeColor = &HFFFFFF

End Sub

Private Sub tmrWindow_Timer()

    'Gets active window status and if the parent
    'form is active it will skin it appropriately
    'and vice versa. It only updates the graphic
    'once per event to prevent flickering and increase
    'performance. This method is present throughout
    'the project, it's the best way really.
    
    If (GetActiveWindow = UserControl.Parent.hwnd) Then
        
        blnOnce(1) = False
        
        If blnOnce(0) = False Then
            blnOnce(0) = True
            ActiveNow
            Exit Sub
        End If
    
    Else
        
        blnOnce(0) = False
        
        If blnOnce(1) = False Then
            blnOnce(1) = True
            imgExit = imgNewX(3)
            imgLeft = imgNewLeft(1)
            imgRight = imgNewRight(1)
            imgBottom = imgNewBottom(1)
            imgTopLeft = imgDisNewTitlebar(0)
            imgTopMiddle = imgDisNewTitlebar(1)
            imgTopRight = imgDisNewTitlebar(2)
            lblTitleBar.ForeColor = &HE0E0E0
            Exit Sub
        End If
    
    End If

End Sub

Private Sub imgTopLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag UserControl.Parent
End Sub

Private Sub imgTopMiddle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag UserControl.Parent
End Sub

Private Sub imgTopRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag UserControl.Parent
End Sub

Private Sub lblTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag UserControl.Parent
End Sub

Private Sub OverX()

    'The exit button animation here.
    'Since this sub is called alot during
    'mouse movements, it only skins once and
    'updates only if the bOnce boolean changes.
    
    'This method is used through out the project
    'to prevent flickering.
    
    If blnOver = True Then
    
        blnOver = False
        
        If blnOnce(1) Then
            imgExit = imgNewX(3)
        Else
            imgExit = imgNewX(0)
        End If
    
    End If

End Sub

Private Sub SetupImages()

    'Copy images in memory at runtime
    'to save file space, logical.
    
    imgNewX(0) = imgExit
    imgNewLeft(0) = imgLeft
    imgNewRight(0) = imgRight
    imgNewBottom(0) = imgBottom
    imgActiveNewTitlebar(0) = imgTopLeft
    imgActiveNewTitlebar(1) = imgTopMiddle
    imgActiveNewTitlebar(2) = imgTopRight
    
End Sub


