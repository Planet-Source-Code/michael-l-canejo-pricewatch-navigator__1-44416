VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2970
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   198
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   StartUpPosition =   3  'Windows Default
   Begin ActiveXP.XPGUI XPGUI1 
      Height          =   2970
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   5239
      AutoSize        =   -1  'True
      Caption         =   "Mikes ActiveXP Control"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub XPGUI1_ExitButton(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Example terminate code
    'Clear arrays first here...
    'Unload forms here after...
    'etc
    
    '
    End
End Sub
