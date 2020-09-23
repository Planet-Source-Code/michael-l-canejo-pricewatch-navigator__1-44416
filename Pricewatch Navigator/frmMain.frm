VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Pricewatch Navigator"
   ClientHeight    =   6180
   ClientLeft      =   -10005
   ClientTop       =   -10005
   ClientWidth     =   4545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   4545
   Visible         =   0   'False
   Begin MSComctlLib.ListView lstVCatalog 
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   570
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Default"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Timer tmrWindow 
      Interval        =   1
      Left            =   2600
      Top             =   4515
   End
   Begin VB.ListBox lstCategory 
      Height          =   1815
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   570
      Width           =   2040
   End
   Begin MSWinsockLib.Winsock sckProduct 
      Left            =   3495
      Top             =   4515
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckCatalog 
      Left            =   3045
      Top             =   4515
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstVProduct 
      Height          =   2340
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2640
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   4128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Price"
         Object.Width           =   1535
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product"
         Object.Width           =   5239
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3450
      Top             =   5175
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "multimedia"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":097E
            Key             =   "cpu"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A35
            Key             =   "new"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AEE
            Key             =   "for notebooks"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B96
            Key             =   "error"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C7B
            Key             =   "i/o"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D36
            Key             =   "input"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DFF
            Key             =   "memory"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EBB
            Key             =   "networking"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F82
            Key             =   "output"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1042
            Key             =   "storage"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":110F
            Key             =   "other"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11B2
            Key             =   "systems"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgNewTitleBar 
      Height          =   315
      Index           =   0
      Left            =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.Image imgNewX 
      Height          =   195
      Index           =   3
      Left            =   4350
      Picture         =   "frmMain.frx":1274
      Top             =   5100
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNewTitleBar 
      Height          =   315
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":14BE
      Top             =   5775
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.Image imgNewLeft 
      Height          =   4770
      Index           =   0
      Left            =   4350
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgNewBottom 
      Height          =   45
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":19DB
      Top             =   5250
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.Image imgNewRight 
      Height          =   4770
      Index           =   1
      Left            =   4200
      Picture         =   "frmMain.frx":1A70
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgNewBottom 
      Height          =   45
      Index           =   0
      Left            =   0
      Top             =   5175
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.Image imgNewRight 
      Height          =   4770
      Index           =   0
      Left            =   4125
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgNewLeft 
      Height          =   4770
      Index           =   1
      Left            =   4425
      Picture         =   "frmMain.frx":1B05
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgNewX 
      Height          =   195
      Index           =   1
      Left            =   4350
      Picture         =   "frmMain.frx":1B9E
      Top             =   4875
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNewX 
      Height          =   195
      Index           =   2
      Left            =   4125
      Picture         =   "frmMain.frx":1DE8
      Top             =   4875
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgNewX 
      Height          =   195
      Index           =   0
      Left            =   4125
      Top             =   5100
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgExit 
      Height          =   195
      Left            =   3810
      Picture         =   "frmMain.frx":2032
      Top             =   75
      Width           =   195
   End
   Begin VB.Label lblTitleBar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pricewatch Navigator"
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
      TabIndex        =   7
      Top             =   60
      Width           =   1800
   End
   Begin VB.Image imgBottom 
      Height          =   45
      Left            =   0
      Picture         =   "frmMain.frx":227C
      Top             =   5050
      Width           =   4080
   End
   Begin VB.Image imgRight 
      Height          =   4770
      Left            =   4040
      Picture         =   "frmMain.frx":232A
      Top             =   300
      Width           =   45
   End
   Begin VB.Image imgLeft 
      Height          =   4770
      Left            =   0
      Picture         =   "frmMain.frx":23E6
      Top             =   300
      Width           =   45
   End
   Begin VB.Image imgTitleBar 
      Height          =   315
      Left            =   0
      Picture         =   "frmMain.frx":2480
      Top             =   0
      Width           =   4080
   End
   Begin VB.Label lblCatalog 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   345
      Width           =   45
   End
   Begin VB.Label lblCategory 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Top             =   345
      Width           =   45
   End
   Begin VB.Label lblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double click the first box to begin."
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   345
      Width           =   2400
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   105
      TabIndex        =   3
      Top             =   2400
      Width           =   45
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Created ßy: Mike Canejo (mikecanejo@hotmail.com)"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   3840
   End
   Begin VB.Label lblProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   2400
      Width           =   45
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   3510
      TabIndex        =   0
      Top             =   345
      Width           =   435
   End
End
Attribute VB_Name = "frmMain"
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
'The purpose of this project was to find an easier way
'to just download prices on the newest devices in a
'smaller more convenient sleek compact style. My main
'use for this program was for my old website which
'included the latest prices for video cards and CPUs
'via the save output to file product feature. It outputs
'a very organized list of products and prices based on
'the contents of the treeview.

'This project is commented moderately, but the code is very
'clean, optimized and straight forward for the typical
'intermediate vb programmer. You will see nice sorting
'functions i wrote as well as how to use a tree view with
'icons via the imagelist control. Arrays are present and
'a complete XP gui is present. The xp gui is identical to
'the real thing. Soon after this project i made an xp gui
'control ocx, but due to it's slight larger size, i decided
'to use the "raw" approach. But as a bonus i included my
'ActiveXP control in this download :) This project was set
'to be under 125kb and by doing so it made me cautious of
'my coding and graphics i used. I spent countless hours
'optimizing the graphics to be visually acceptable as well
'size. Some icons like the "Other" symbol with a ? was made
'by me in MSpaint. I'm not a artist so it wasn't easy ;)

'This project is error and bug free. It's coded for
'performance and clean organized code. Arrays are being
'erased, forms are being unloaded and images are being
'stored in memory to save file size. Graphics manipulation
'is only being done when nessacary to greatly increase
'performance and prevent the dreaded "flickering" curse.
'Form cannot be resized by design, i preferred it to be
'stationary and static in dimensions. Try to compare the
'xp gui to a real one and you'll see it's identical. Most
'xp gui's are cheesy and a ripoff, you can spot it right
'away most of the time. Feel free to email me for any
'comments/etc at the email above.

'Enjoy and hopefully learn something new :)

'Oh, if you find anything useful in my code or think it
'would be worth your time, please goto:
'http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=44416&lngWId=1
'and vote, not much to ask ;)
'-----------------------------------------
'DISCLAIMER: You are free to distribute/modify the source
'code but please leave the comments located in all the
'modules declaration sections intact.
'-----------------------------------------

Option Explicit

Private Const strTitle  As String = "Pricewatch Navigator "
Private Const strBuild  As String = "1.2.7 "

Private lngPaks         As Long
Private strHost         As String
Private strRoot         As String
Private strWeb()        As String
Private strBuffer       As String
Private strCatalog      As String
Private strCategory     As String
Private blnError        As Boolean
Private blnOver         As Boolean
Private blnDown         As Boolean
Private blnDone         As Boolean
Private blnClick(1)     As Boolean
Private blnOnce(1)      As Boolean

'HTTP header for pricewatch.com
Private Const strHeader As String = _
      "GET root% HTTP/1.0" & vbNewLine _
    & "Accept: */*" & vbNewLine _
    & "Accept -Language: en -us" & vbNewLine _
    & "User-Agent: Mozilla/4.0 " _
    & "Host: web%" & vbNewLine & vbNewLine
            

Private Sub Form_Load()
    
    Setup 'copy images to decrease filesize
    lblTitleBar.Caption = strTitle & strBuild
    frmMain.Move Screen.Width / 2 - frmMain.Width / 2, _
    Screen.Height / 2 - frmMain.Height / 2, 4080, 5100
    Show
    RefreshLists
    FirstTime

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    sckCatalog.Close
    sckProduct.Close
    Clipboard.Clear
    Terminate

End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        imgExit = imgNewX(1)
        blnDown = True
    End If

End Sub

Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        
        If X > 0 And X < imgExit.Width _
        And Y > 0 And Y < imgExit.Height Then
            
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
        
        If X >= 0 And X < imgExit.Width _
        And Y >= 0 And Y < imgExit.Height Then
            Form_Unload 0
        End If
    
    End If

End Sub

Private Sub lblAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    OverX
    
    With frmAbout
        .lblBuild.Caption = "build: " & strBuild
        .lblDate.Caption = Date
    End With
    
    frmAbout.Show 1, Me

End Sub

Private Sub lstVCatalog_DblClick()

    Dim lngFind         As Long
    Dim lngIndex        As Long
    Dim strFind         As String
    Dim strIndex        As String
    Dim strProduct      As String
    Dim intPos          As Integer
        
    If blnError Then Exit Sub
    
    If lstVCatalog.ListItems.Count = 0 Then Exit Sub
    
    lstCategory.Clear
    
    lblCategory.Caption = "Select a product:"
    
    strFind = lstVCatalog.ListItems.Item _
    (lstVCatalog.SelectedItem.Index)
    
    lblAbout.Visible = True
    lblCatalog.Caption = strFind
    lstCategory.Tag = strFind
    
    lngIndex = InStr(1, strCatalog, ":" & strFind & ";")
    
    If lngIndex = 0 Then
        lstCategory.AddItem "Not found, please refresh."
        Exit Sub
    End If
    
    strIndex = "|" & Mid(strCatalog, lngIndex + Len(strFind) + 2)
    strIndex = Left(strIndex, InStr(1, strIndex, ":") - 1)
    
    ReDim strWeb(0)
    
    Do
    
        lngFind = InStr(lngFind + 1, strIndex, "|")
        If lngFind = 0 Then Exit Do
        
        strProduct = Mid(strIndex, lngFind + 1)
        If strProduct = "" Then Exit Do
        
        strProduct = Left(strProduct, InStr(1, strProduct, "|") - 1)
        
        ReDim Preserve strWeb(intPos)
        strWeb(intPos) = strWeb(intPos) & strProduct
        strProduct = Left(strProduct, InStr(1, strProduct, "~") - 1)
        lstCategory.AddItem Trim(strProduct)
        
        intPos = UBound(strWeb) + 1
        
    Loop
    
    SortByLength_List lstCategory
    lblCatalog.Caption = lblCatalog.Caption _
    & " - [" & lstVCatalog.ListItems.Count & "]"
    
End Sub

Private Sub lstVCatalog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        If lstVCatalog.ListItems.Count = 0 Then Exit Sub
        PopupMenu frmDummy.mnuMain2, 1
    End If

End Sub

Private Sub lstCategory_DblClick()

    Dim intX            As Integer
    Dim strFind(1)      As String
    
    If blnError Then Exit Sub
    If lstCategory.ListCount = 0 Then Exit Sub
    
    blnDone = False
    strBuffer = vbNullString
    
    strFind(0) = lstCategory.List(lstCategory.ListIndex)
    
    lblCategory.Caption = strFind(0) _
    & " - [" & lstCategory.ListCount & "]"
    
    lblProduct.Caption = ""
    
    sckProduct.Tag = strFind(0)
    lblInfo.Caption = "Waiting for products..."
    lblAbout.Visible = False
    
    For intX = 0 To UBound(strWeb)
        
        strFind(1) = Left(strWeb(intX), _
        InStr(1, strWeb(intX), "~") - 1)
        
        If strFind(0) = strFind(1) Then
            
            strRoot = "/menus/" & Mid(strWeb(intX), _
            InStr(1, strWeb(intX), "~") + 1)
            
            sckProduct.Close
            sckProduct.Connect strHost, 80
            
            Exit For
            
        End If
        
    Next
        
End Sub

Private Sub lstVProduct_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Select Case (ColumnHeader.Index)
        
        Case 1
            
            blnClick(1) = False
            
            If blnClick(0) = False Then
                blnClick(0) = True
                SortByPrice lstVProduct, True
            Else
                blnClick(0) = False
                SortByPrice lstVProduct, False
            End If
        
        Case 2
            
            blnClick(0) = False
            lstVProduct.SortKey = ColumnHeader.Index - 1
            
            If blnClick(1) = False Then
                blnClick(1) = True
                lstVProduct.SortOrder = lvwAscending
                lstVProduct.Sorted = True
            Else
                blnClick(1) = False
                lstVProduct.SortOrder = lvwDescending
                lstVProduct.Sorted = True
            End If
            
    End Select
    
End Sub

Private Sub lstVProduct_DblClick()

    Dim lngX            As Long
    Dim strWeb          As String
    Dim strProduct      As String

    If blnError Then Exit Sub
    If lstVProduct.ListItems.Count = 0 Then Exit Sub
    
    lngX = lstVProduct.SelectedItem.Index
    
    strProduct = lstVProduct.ListItems(lngX).ListSubItems(1)
    strProduct = Replace(strProduct, Chr(32), "+")
    
    strWeb = "http://castle.pricewatch.com/search/search.idq?qc=22" _
    & UCase(strProduct) & "22*+AND+40totalcost3E0&cr=" & LCase(strProduct)
    
    'Goto pricewatch.com's site to get
    'cheapest prices on selected product
    GotoURL strWeb
    
End Sub

Private Sub lstVProduct_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        If lstVProduct.ListItems.Count = 0 Then Exit Sub
        PopupMenu frmDummy.mnuMain, 1
    End If
    
End Sub

Private Sub sckCatalog_Close()

    Dim lngFind(1)      As Long
    Dim strTopic(1)     As String
    Dim strFind(2)      As String
    Dim strTerm         As String
    Dim strLink         As String
    Dim strTitle        As String
    Dim strIcon         As String
    Dim intX            As Integer

    blnDone = True
    blnError = False
    
    sckCatalog.Close
    
    lstVCatalog.ListItems.Clear
    lstVProduct.Sorted = False
    
    If strBuffer = "" Then Exit Sub
    
    strBuffer = LCase(strBuffer)
    
    strFind(0) = "<td><b>"
    strFind(1) = "<br><b>"
    strFind(2) = "<a href="
    
    If InStr(1, strBuffer, strFind(0)) = 0 Then Exit Sub
    If InStr(1, strBuffer, strFind(1)) = 0 Then Exit Sub

    strCatalog = ""
    
    strBuffer = Replace(strBuffer, " target=" & Chr(34) _
    & "lf" & Chr(34) & " id=" & Chr(34), "")
    
    For intX = 0 To 1

        Do
        
            lngFind(0) = InStr(lngFind(0) + Len(strFind _
            (intX)), strBuffer, strFind(intX))
            
            If lngFind(0) = 0 Then Exit Do
            
            strTopic(0) = Mid(strBuffer, _
            lngFind(0) + Len(strFind(intX)))
            
            strTopic(0) = Left(strTopic(0), _
            InStr(1, strTopic(0), "<") - 1)
            
            strIcon = Trim(strTopic(0))
            
            If strIcon <> "cpu" _
            And strIcon <> "networking" _
            And strIcon <> "for notebooks" _
            And strIcon <> "i/o" _
            And strIcon <> "storage" _
            And strIcon <> "multimedia" _
            And strIcon <> "other" _
            And strIcon <> "memory" _
            And strIcon <> "input" _
            And strIcon <> "systems" _
            And strIcon <> "output" Then
                strIcon = "new"
                'New menu added to pricewatch's homepage
                'this would mean my program don't support
                'the newer device categories ;)
            End If
 
            lstVCatalog.ListItems.Add , , strIcon, , strIcon
            strCatalog = strCatalog & ":" & strIcon & ";"
            lngFind(1) = lngFind(0)
            
            Do
            
                lngFind(1) = InStr(lngFind(1) + _
                Len(strFind(2)), strBuffer, strFind(2))
                
                If lngFind(1) = 0 Then Exit Do
                
                strTopic(1) = Mid(strBuffer, _
                lngFind(1) + Len(strFind(2)))
                
                strTerm = Mid(strTopic(1), _
                InStr(1, strTopic(1), "</a>") + 4)

                strLink = Mid(strTopic(1), 2, _
                InStr(2, strTopic(1), Chr(34)) - 2)
                
                strTitle = Mid(strTopic(1), _
                InStr(1, strTopic(1), ">") + 1)
                
                strTitle = Left(strTitle, _
                InStr(1, strTitle, "<") - 1)
                
                strCatalog = strCatalog & _
                strTitle & "~" & strLink & "|"
                
                If Left(strTerm, 11) = "<br><br><b>" Then
                    Exit Do
                ElseIf Left(strTerm, 12) = "</td><td><b>" Then
                    Exit Do
                End If
                
            Loop
            
        Loop
            
    Next
    
    lngFind(0) = 0
        
    strCatalog = strCatalog & ":"
    SortByLength_Tree lstVCatalog
    
    blnClick(1) = False
    blnClick(0) = True
    
End Sub

Private Sub sckCatalog_Connect()

    Dim strNew As String
    
    strNew = Replace(strHeader, "web%", strHost)
    strNew = Replace(strNew, "root%", strRoot)
    sckCatalog.SendData strNew
    
End Sub

Private Sub sckCatalog_DataArrival(ByVal bytesTotal As Long)

    Dim strData As String
    
    sckCatalog.GetData strData
    strBuffer = strBuffer & strData
    
    lngPaks = lngPaks + 1
    
    lblTitleBar.Caption = strTitle & strBuild _
    & " - [" & AddCommas(lngPaks) & " packets]"
    
    If InStr(1, LCase(strData), "</html>") Then
        sckCatalog_Close
    End If
    
End Sub

Private Sub sckCatalog_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    blnError = True
    sckProduct.Close
    
    MsgBox "An error has occured while trying " _
    & "to connect to Pricewatch.com, please try " _
    & "again later.", vbCritical, "Error:"
    
    lstVCatalog.ListItems.Clear
    lstVCatalog.ListItems.Add , , "TCP/IP Error!", , "error"
    
End Sub

Private Sub sckProduct_Close()

    Dim lngFind         As Long
    Dim strFind         As String
    Dim strIcon         As String
    Dim strPrice        As String
    Dim strProduct      As String
    Dim intX            As Integer
    
    blnDone = True
    sckProduct.Close
    
    lstVProduct.Sorted = False
    lstVProduct.ListItems.Clear
    lstVProduct.ColumnHeaders(2).Width = 2970
    
    If strBuffer = "" Then Exit Sub
    
    strFind = "<tr><td>"
    
    strBuffer = Replace(strBuffer, "</td><td align" _
    & "=center>&nbsp;-&nbsp;</td><td>", Chr(32))
    
    strBuffer = Replace(strBuffer, "</td><td align" _
    & "=center><img src=" & Chr(34) & "/dn.gif" _
    & Chr(34) & "></td><td>", Chr(32))
    
    strBuffer = Replace(strBuffer, "</td><td align" _
    & "=center><img src=" & Chr(34) & "/up.gif" _
    & Chr(34) & "></td><td>", Chr(32))

    Do
    
        lngFind = InStr(lngFind _
        + Len(strFind), strBuffer, strFind)
        
        If lngFind = 0 Then Exit Do
        
        strProduct = Mid(strBuffer, _
        lngFind + Len(strFind))
        
        strProduct = Left(strProduct, _
        InStr(1, strProduct, "</A>") - 1)
        
        strProduct = RemHTML(strProduct)
        strProduct = Trim(strProduct)
        strProduct = RemSpaces(strProduct)
        
        If InStr(1, strProduct, "All in Category") = 0 Then
        
            lstVProduct.ListItems.Add , , Left(strProduct, _
            InStr(1, strProduct, Chr(32)) - 1), , lstCategory.Tag
            
            lstVProduct.ListItems(lstVProduct. _
            ListItems.Count).ListSubItems.Add , , _
            Mid(strProduct, InStr(1, strProduct, Chr(32)) + 1)
            
            If lstVProduct.ListItems.Count > 8 Then _
            lstVProduct.ColumnHeaders(2).Width = 2684
                
            lblInfo.Caption = "Waiting for products..." _
            & "[" & lstVProduct.ListItems.Count & "]"
        
        End If
    
    Loop
    
    blnClick(1) = False
    blnClick(0) = True
    
    SortByPrice lstVProduct
    lblInfo.Caption = "Products for:"
    
    lblProduct.Caption = sckProduct.Tag _
    & " - [" & lstVProduct.ListItems.Count & "]"
    
End Sub

Private Sub sckProduct_Connect()

    Dim strNew As String
    
    'Format header by replacing strings
    strNew = Replace(strHeader, "web%", strHost)
    strNew = Replace(strNew, "root%", strRoot)
    sckProduct.SendData strNew
    
End Sub

Private Sub sckProduct_DataArrival(ByVal bytesTotal As Long)

    Dim strData As String

    sckProduct.GetData strData
    
    'Add total data arrived to buffer to prevent
    'lost packets. Not all data comes in one packet.
    strBuffer = strBuffer & strData
    
    lngPaks = lngPaks + 1
    
    lblTitleBar.Caption = strTitle & strBuild _
    & " - [" & AddCommas(lngPaks) & " packets]"
    
    If InStr(1, LCase(strData), "</html>") Then
        sckProduct_Close
    End If
    
End Sub

Private Sub sckProduct_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    blnError = True
    sckProduct.Close
    
    MsgBox "An error has occured while trying " _
    & "to connect to Pricewatch.com, please try " _
    & "again later.", vbCritical, "Error:"
    
    lstVProduct.ListItems.Clear
    lstVProduct.ListItems.Add , , "NOTE:", , "error"
    lstVProduct.ListItems(1).ListSubItems.Add , , "TCP/IP Failure!"

End Sub

Public Sub ActiveWindow()

    If blnDown = False Then
        imgExit = imgNewX(0)
    End If
    
    imgLeft = imgNewLeft(0)
    imgRight = imgNewRight(0)
    imgBottom = imgNewBottom(0)
    imgTitleBar = imgNewTitleBar(0)
    lblTitleBar.ForeColor = &HFFFFFF
    
End Sub

Private Sub tmrWindow_Timer()

    'Very optimized for performance. This wont keep
    'skinning the gui, it does it once and only once.
    'When the value of getactivewindow changes, it
    'switches to inactive mode, rinse and repeat.

    If (GetActiveWindow = hwnd) Then
        
        blnOnce(1) = False
        
        If blnOnce(0) = False Then
        
            blnOnce(0) = True
            ActiveWindow
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
            imgTitleBar = imgNewTitleBar(1)
            lblTitleBar.ForeColor = &HE0E0E0
            Exit Sub
            
        End If
    
    End If

End Sub

Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If blnOver Then
        
        blnOver = False
        
        If blnOnce(1) Then
            imgExit = imgNewX(3)
        Else
            imgExit = imgNewX(0)
        End If
    
    End If
    
    With lblAbout
        
        If .Tag = "" Then
            .Tag = "Over"
            .ForeColor = &HFF&
            .FontBold = True
        End If
    
    End With

End Sub

Private Sub Terminate()

    'Delete Arrays! Unload forms properly
    
    Set frmMain = Nothing
    Set frmDummy = Nothing
    Set frmAbout = Nothing
    
    Erase blnClick()
    Erase blnOnce()
    Erase strWeb()
    
    Unload frmAbout
    Unload frmDummy
    Unload Me
    
    End
    
End Sub

Private Function AddCommas(ByVal strNum As String)
 
    Dim intCycle As Integer
    
    For intCycle = Len(strNum) - 3 To 1 Step -3
        strNum = Left(strNum, intCycle) _
        & "," & Mid(strNum, intCycle + 1)
    Next
    
    AddCommas = strNum
    
End Function

Private Function RemHTML(strStr As String) As String

    Dim lngFind(1) As Long
    
    lngFind(0) = InStr(1, strStr, "<")
    lngFind(1) = InStr(1, strStr, ">")
    RemHTML = Replace(strStr, Mid(strStr, _
    lngFind(0), lngFind(1) - lngFind(0) + 1), "")
    
End Function

Private Function RemSpaces(strStr As String) As String

    Dim intX As Integer
    
    For intX = 10 To 2 Step -1
        strStr = Replace(strStr, _
        String(intX, Chr(32)), Chr(32))
    Next
    
    RemSpaces = strStr
    
End Function

Private Sub Setup()

    'Instead of using identical images which takes
    'up filesize, i simply copied them in memory after
    'runtime via image copy.
    
    imgNewX(0) = imgExit
    imgNewLeft(0) = imgLeft
    imgNewRight(0) = imgRight
    imgNewBottom(0) = imgBottom
    imgNewTitleBar(0) = imgTitleBar
    
End Sub

Public Sub SaveList()

    'Straight forward, output
    'treeview in organized list.
    
    dFileName = vbNullString
    
    Dialog "Text Files (*.txt)" + Chr(0) + "*.txt" _
    + Chr(0) + "All Files (*.*)" + Chr(0) + "*.*" _
    + Chr(0), "Save Products", Me, ".txt", App.Path, False
    
    If dFileName = "" Then Exit Sub
    
    SaveListBox lstVProduct, dFileName, False
    
End Sub

Private Sub OverX()
    'Optimize mouseover X button and about label,
    'skins once then waits for change to skin again.
    
    If blnOver Then
    
        blnOver = False
        If blnOnce(1) Then
            imgExit = imgNewX(3)
        Else
        
            imgExit = imgNewX(0)
        End If
        
    End If
    
    With lblAbout
    
        If .Tag = "Over" Then
            .Tag = ""
            .ForeColor = &HC0&
            .FontBold = False
        End If
        
    End With
    
End Sub

Private Sub FirstTime()

    'Shows splash for the first time then
    'disables from showing again on startup.
    'Clicking "About" label brings it up.
    
    If GetOpt("Splash", "No") = "No" Then
        SaveOpt "Splash", "Yes"
        frmAbout.Show
    End If
    
End Sub

Public Sub RefreshLists()

    blnDone = False
    strBuffer = vbNullString
    
    strRoot = "/newhomepage.htm"
    strHost = "www.pricewatch.com"
    
    lblHelp(0).Visible = False
    lblHelp(1).Visible = False
    
    lstCategory.Clear
    lstVCatalog.ListItems.Clear
    lstVProduct.ListItems.Clear
    
    lblProduct.Caption = ""
    lblCatalog.Caption = "Select a catalog:"
    lblCategory.Caption = "Select a category:"
    lblInfo.Caption = "No categories selected."
    
    sckCatalog.Close
    sckCatalog.Connect strHost, 80
    
End Sub

Public Sub SaveListBox(lstView As ListView, _
    strPath As String, Optional bAppend As Boolean)

    'Outputs the product tree view in a nice organized
    'style which is copy and paste ready for a website :)
    'All i ask if to include a credit to my name so people
    'will email me with curiousity. Thanks :)
    

    Dim lngX        As Long
    Dim strCreator  As String
    Dim strLine1    As String
    Dim strLine2    As String
    Dim intFree     As Integer
    

    strCreator = _
        " ==================================" & vbNewLine & _
        "       *Pricewatch Navigator*      " & vbNewLine & _
        "          By: Mike Canejo          " & vbNewLine & _
        "      AIM: Mike3dd or Mikey3dd     " & vbNewLine & _
        "   E-mail: mikecanejo@hotmail.com  " & vbNewLine & _
        " ==================================" & vbNewLine & _
        vbNewLine & vbNewLine & _
        " Products for: cat%" & vbNewLine & _
        " =================================="
        
    On Error GoTo No_File
    
    intFree = FreeFile
    
    If bAppend Then
        Open strPath For Append As intFree
    Else
        Open strPath For Output As intFree
    End If
    
    Print intFree, Replace(strCreator, _
    "cat%", lblProduct.Caption)
    
    For lngX = 1 To lstView.ListItems.Count
    
        strLine1 = lstView.ListItems.Item(lngX)
        strLine2 = lstView.ListItems(lngX).ListSubItems.Item(1)
        
        If strLine1 <> vbNullString _
        And strLine2 <> vbNullString Then
        
            Print intFree, Chr(32) _
            & strLine1 & vbTab & strLine2
            
        End If
        
    Next
    
    Print intFree, " ==================================";
    
No_File:

    Close intFree

End Sub

Public Sub SortByLength_Tree(lstView As ListView, _
                        Optional blnLowOrHigh As Boolean = True, _
                        Optional intDefMaxLen = 1000)
                        
    'Function created to make a organized display of
    'the treeview items from sorting by length.
    
    Dim intLen          As Integer
    Dim intLowLen       As Integer
    Dim intHighLen      As Integer
    Dim intPos          As Integer
    Dim strTmp(1)       As String
    Dim strLens         As String
    Dim lngY            As Long
    Dim lngX            As Long
    
    
    If lstView.ListItems.Count = 0 Then Exit Sub
    
    intLowLen = intDefMaxLen
    strLens = ":"
    intLowLen = intDefMaxLen
    
    For lngY = 1 To lstView.ListItems.Count
    
        intLen = Len(lstView.ListItems.Item(lngY))
        
        If InStr(1, strLens, ":" & intLen & ":") = 0 Then _
        strLens = strLens & intLen & ":"

        If intLen > intHighLen Then
            intHighLen = intLen
        End If
        
        If intLen < intLowLen Then
            intLowLen = intLen
        End If
        
    Next

    If blnLowOrHigh Then
    
        intLen = intLowLen
        intLowLen = intHighLen
        intHighLen = intLen
        intLen = 1
        
    Else
        intLen = -1
        
    End If


    For lngY = intHighLen To intLowLen Step intLen
    
        If InStr(1, strLens, ":" & lngY & ":") <> 0 Then
            
            For lngX = 1 To lstView.ListItems.Count
                
                If Len(lstView.ListItems.Item(lngX)) = lngY Then
                
                    intPos = intPos + 1
                    strTmp(0) = lstView.ListItems.Item(intPos)
                    strTmp(1) = lstView.ListItems.Item(intPos).SmallIcon

                    lstView.ListItems.Item(intPos) = _
                    lstView.ListItems.Item(lngX)
                    lstView.ListItems.Item(intPos).SmallIcon = _
                    lstView.ListItems.Item(lngX).SmallIcon

                    lstView.ListItems.Item(lngX) = strTmp(0)
                    lstView.ListItems.Item(lngX).SmallIcon = strTmp(1)
                
                End If
            Next
            
        End If
        
    Next
    
End Sub

Public Sub SortByLength_List(lstBox As ListBox, _
                         Optional blnLowOrHigh As Boolean = True)
                         
    'Decided to write another one because i opted to
    'use a listbox for the sub-category display. I also
    'kept in consideration that a person might find these
    'sorting functions useful so listbox & treeview are
    'better than one.

    Dim intY            As Integer
    Dim intX            As Integer
    Dim intLen          As Integer
    Dim intLowLen       As Integer
    Dim intHighLen      As Integer

    If lstBox.ListCount = 0 Then Exit Sub
    
    intLowLen = 1000

    For intY = 0 To lstBox.ListCount - 1
    
        intLen = Len(lstBox.List(intY))

        If intLen > intHighLen Then
            intHighLen = intLen
        End If
        
        If intLen < intLowLen Then
            intLowLen = intLen
        End If
        
    Next

    If blnLowOrHigh Then
    
        intLen = -1
        
    Else
    
        intLen = intLowLen
        intLowLen = intHighLen
        intHighLen = intLen
        intLen = 1
        
    End If
    

    For intY = intHighLen To intLowLen Step intLen

        For intX = 0 To lstBox.ListCount - 1
        
            If Len(lstBox.List(intX)) = intY Then
                
                lstBox.AddItem lstBox.List(intX), 0
                lstBox.RemoveItem intX + 1
                
            End If
            
        Next
        
    Next
    
End Sub

Public Sub SortByPrice(lstView As ListView, _
                        Optional bLowOrHigh As Boolean = True)

    Dim lngX            As Long
    Dim lngY            As Long
    Dim lngPos          As Long
    Dim lngPrice        As Long
    Dim lngLowPrice     As Long
    Dim lngHighPrice    As Long
    Dim strTmp          As String
    Dim strPrices       As String
    
    'Most convenient sort function. Sorts by cheapest
    'to most expensive or vice versa. I had to make my
    'own sort function for this because the sorting
    'algorithm built into the tree view control doesn't
    'sort numbers correctly to my likes (ie. 10 1 23 2 3).
    

    If lstView.ListItems.Count = 0 Then Exit Sub
    
    lngLowPrice = 1000000
    strPrices = ":"

    For lngX = 1 To lstView.ListItems.Count
    
        lngPrice = CInt(Mid(lstView.ListItems.Item(lngX), 2))
        
        If InStr(1, strPrices, ":" & lngPrice & ":") = 0 Then _
        strPrices = strPrices & lngPrice & ":"

        If lngPrice > lngHighPrice Then
            lngHighPrice = lngPrice
        End If
        
        If lngPrice < lngLowPrice Then
            lngLowPrice = lngPrice
        End If
        
    Next

    If bLowOrHigh Then
    
        lngPrice = lngLowPrice
        lngLowPrice = lngHighPrice
        lngHighPrice = lngPrice
        lngPrice = 1
        
    Else
    
        lngPrice = -1
        
    End If

    For lngY = lngHighPrice To lngLowPrice Step lngPrice
    
        If InStr(1, strPrices, ":" & lngY & ":") <> 0 Then

            For lngX = 1 To lstView.ListItems.Count
            
                lngPrice = CInt(Mid(lstView.ListItems.Item(lngX), 2))
                
                If lngPrice = lngY Then

                    lngPos = lngPos + 1

                    strTmp = lstView.ListItems. _
                    Item(lngPos)
                    
                    lstView.ListItems.Item(lngPos) _
                    = lstView.ListItems.Item(lngX)
                    
                    lstView.ListItems.Item(lngX) = strTmp

                    strTmp = lstView.ListItems( _
                    lngPos).ListSubItems.Item(1)
                    
                    lstView.ListItems(lngPos).ListSubItems.Item(1) _
                    = lstView.ListItems(lngX).ListSubItems.Item(1)
                    
                    lstView.ListItems(lngX).ListSubItems.Item(1) = strTmp

                End If
                
            Next
            
        End If
        
    Next
    
End Sub




'Mouse movement procedures
'-------------------------

Private Sub imgTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub lblTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub imgRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub imgTitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub lblHelp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub lblTitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub lstVCatalog_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub lstCategory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub

Private Sub lstVProduct_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverX
End Sub
