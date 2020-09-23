VERSION 5.00
Begin VB.Form frmDummy 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   30
   ClientLeft      =   -10035
   ClientTop       =   -120
   ClientWidth     =   1530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDummy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   1530
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Menu mnuMain 
      Caption         =   "Menu1"
      Begin VB.Menu mnuSaveList 
         Caption         =   "Save List"
      End
   End
   Begin VB.Menu mnuMain2 
      Caption         =   "Menu2"
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmDummy"
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
'I used a dummy form for my popup menus because
'having menus on the main form gave the border a
'thin black line simular to BorderSyle = 1
'-----------------------------------------
Option Explicit

Private Sub mnuRefresh_Click()
    frmMain.RefreshLists
End Sub

Private Sub mnuSaveList_Click()
    frmMain.SaveList
End Sub
