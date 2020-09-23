VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3180
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   3180
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00800080&
      BackStyle       =   0  'Transparent
      Caption         =   "09/08/02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1950
      TabIndex        =   5
      Top             =   2895
      Width           =   825
   End
   Begin VB.Label lblBuild 
      AutoSize        =   -1  'True
      BackColor       =   &H00800080&
      BackStyle       =   0  'Transparent
      Caption         =   "build: 1.0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   2895
      Width           =   975
   End
   Begin VB.Label lblWeb 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   0
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   0
      Width           =   4740
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MikeCanejo@hotmail.com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      MouseIcon       =   "frmAbout.frx":286D
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   375
      Width           =   4725
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Created ßy: Mike Canejo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      MouseIcon       =   "frmAbout.frx":2B77
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   105
      Width           =   4725
   End
   Begin VB.Label lblClose 
      BackColor       =   &H00800080&
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4155
      MousePointer    =   10  'Up Arrow
      TabIndex        =   0
      Top             =   2895
      Width           =   465
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Original Versions 1.0.0 - 1.2.7 [04/1/03]
'Brought to you and written by Mike Canejo
'-----------------------------------------
'AOL/AIM: Mikey3dd
'Email: MikeCanejo@hotmail.com
'-----------------------------------------
'Comments:
'-----------------------------------------
'Graphic by me, created in Bryce3D
'The revision + date is updated by global
'variables set on frmMain declarations.
'-----------------------------------------
Option Explicit

Dim blnOver(2) As Boolean

Private Sub Form_Load()
    Me.Move frmMain.Left - Width / 2 + frmMain.Width / 2, _
    frmMain.Top - Height / 2 + frmMain.Height / 2, 4755, 3210
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase blnOver()
    Set frmAbout = Nothing
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnOver(0) = False
    blnOver(2) = False
    If blnOver(1) = False Then
        blnOver(1) = True
        lblClose.ForeColor = &HFFFFFF
        lblClose.FontUnderline = False
        lblEmail.ForeColor = &HFFFFFF
        lblAbout.ForeColor = &HFFFFFF
        lblAbout.FontUnderline = False
        lblEmail.FontUnderline = False
    End If
End Sub

Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnOver(1) = False
    If blnOver(0) = False Then
        blnOver(0) = True
        lblAbout.ForeColor = &HC0C0FF
        lblEmail.ForeColor = &HC0C0FF
        lblAbout.FontUnderline = True
        lblEmail.FontUnderline = True
    End If
End Sub

Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
    frmMain.Show
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnOver(1) = False
    If blnOver(2) = False Then
        blnOver(2) = True
        lblClose.ForeColor = &HC0C0FF
        lblClose.FontUnderline = True
    End If
End Sub

Private Sub lblEmail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GotoURL "mailto:mikecanejo@hotmail.com"
End Sub

Private Sub lblAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEmail_MouseDown 1, 1, 1, 1
End Sub

Private Sub lblWeb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEmail_MouseDown 1, 1, 1, 1
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAbout_MouseMove 1, 1, 1, 1
End Sub


