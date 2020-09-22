VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmBRW 
   AutoRedraw      =   -1  'True
   Caption         =   "Browser"
   ClientHeight    =   5940
   ClientLeft      =   3060
   ClientTop       =   3315
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBRW.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '³Ì¤j¤Æ
   Begin SHDocVwCtl.WebBrowser BRW 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   7080
      ExtentX         =   12488
      ExtentY         =   7011
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      BorderStyle     =   0  '¨S¦³®Ø½u
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8310
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   8310
      Begin VB.TextBox cboAddress 
         Height          =   315
         Left            =   3960
         TabIndex        =   2
         Top             =   60
         Width           =   3735
      End
      Begin ProFile.CB btnBrw 
         Height          =   435
         Index           =   0
         Left            =   0
         TabIndex        =   4
         ToolTipText     =   "Back"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":000C
         PICN            =   "frmBRW.frx":0028
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnBrw 
         Height          =   435
         Index           =   1
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "Forward"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":073A
         PICN            =   "frmBRW.frx":0756
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnBrw 
         Height          =   435
         Index           =   2
         Left            =   840
         TabIndex        =   6
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":0E68
         PICN            =   "frmBRW.frx":0E84
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnBrw 
         Height          =   435
         Index           =   3
         Left            =   1200
         TabIndex        =   7
         ToolTipText     =   "Stop"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":1596
         PICN            =   "frmBRW.frx":15B2
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnBrw 
         Height          =   435
         Index           =   4
         Left            =   1560
         TabIndex        =   8
         ToolTipText     =   "Home"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":1CC4
         PICN            =   "frmBRW.frx":1CE0
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnBrw 
         CausesValidation=   0   'False
         Height          =   435
         Index           =   6
         Left            =   2280
         TabIndex        =   9
         ToolTipText     =   "Favorites"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":23F2
         PICN            =   "frmBRW.frx":240E
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnGo 
         CausesValidation=   0   'False
         Height          =   435
         Left            =   7800
         TabIndex        =   10
         ToolTipText     =   "Go"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":2B20
         PICN            =   "frmBRW.frx":2B3C
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnBrw 
         Height          =   435
         Index           =   5
         Left            =   1920
         TabIndex        =   11
         ToolTipText     =   "Zoom"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":324E
         PICN            =   "frmBRW.frx":326A
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblAddress 
         Caption         =   "A&ddress:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3165
         TabIndex        =   1
         Tag             =   "&Address:"
         Top             =   90
         Width           =   3075
      End
   End
End
Attribute VB_Name = "frmBRW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EventRunning As Boolean
Dim LastLogMsg As String
Public CurrentlyOpenFile As String
Dim YouCantChangeMyAddyBarTextNow As Boolean 'tag


Dim BrwGen As HTMLGenericElement
Dim BrwHref As HTMLAnchorElement
Dim BrwEvent As IHTMLEventObj
Dim WithEvents BrwDoc As HTMLDocument
Attribute BrwDoc.VB_VarHelpID = -1
Dim old_element As HTMLGenericElement


Private Declare Sub SHAutoComplete Lib "shlwapi.dll" (ByVal hwndEdit As Long, ByVal dwFlags As Long) 'hinted by Juanito Dado Jr

Private Sub BRW_BeforeNavigate2(ByVal pDisp As Object, url As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    On Error Resume Next
    Form_Resize
    SetAddy BRW.LocationURL
    If GetSet("BRW_Log", "1") = "1" Then
        If BRW.LocationURL <> LastLogMsg Then
            cboAddress.BackColor = IIf(Left$(BRW.LocationURL, 5) = "https", RGB(204, 255, 204), RGB(255, 255, 255))
            Log BRW.LocationURL, True
            LastLogMsg = BRW.LocationURL
        End If
    End If
    
End Sub



Private Sub BRW_DocumentComplete(ByVal pDisp As Object, url As Variant)
    On Error Resume Next
    EventRunning = True
    SetAddy BRW.LocationURL
    EventRunning = False
    CCaption BRW.LocationName, Me
    
    CurrentlyOpenFile = CStr(url)
    
    If GetSet("BRW_Log", "1") = "1" Then
        If BRW.LocationURL <> LastLogMsg Then
            cboAddress.BackColor = IIf(Left$(BRW.LocationURL, 5) = "https", RGB(204, 255, 204), RGB(255, 255, 255))
            Log BRW.LocationURL, True
            LastLogMsg = BRW.LocationURL
        End If
    End If
    
    Set BrwDoc = BRW.Document
    
End Sub

Private Sub BRW_FileDownload(Cancel As Boolean)
    On Error Resume Next
    If GetSet("BRW_Download", "1") = "0" Then Cancel = True
End Sub

Private Sub BRW_NewWindow2(ppDisp As Object, Cancel As Boolean)
    On Error Resume Next
    If GetSet("Browser_AllowNewWindow", "1") = "1" Then
        Dim F As New frmBRW
        Set ppDisp = F.BRW.object
        F.Show
    Else
        Cancel = True
    End If
End Sub

Private Sub BRW_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    SProgress Progress, 0, ProgressMax
End Sub

Private Sub BRW_StatusTextChange(ByVal Text As String)
    SStatus Text, vbInformation
End Sub

Private Sub BRW_TitleChange(ByVal Text As String)
    On Error Resume Next
    CCaption Text, Me
End Sub

Private Sub BRW_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
    On Error Resume Next
    Unload Me
End Sub

Public Sub btnBrw_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            BRW.GoBack
        Case 1
            BRW.GoForward
        Case 2
            BRW.Refresh
        Case 3
            BRW.Stop
            CCaption BRW.LocationName, Me
        Case 4
            BRW.GoHome
        Case 5
            'BRW.GoSearch
            PopupMenu frmMain.titBrowserZoom, , btnBrw(5).Left, btnBrw(5).Top + btnBrw(5).Height
        Case 6
            PopupMenu frmMain.titBrowserP, , btnBrw(6).Left, btnBrw(6).Top + btnBrw(6).Height, frmMain.titBrowserPBMThis
            'frmMain.titBrowserFavorites_Click
    End Select
End Sub

Private Sub cboAddress_Change()
    On Error Resume Next
    cboAddress.BackColor = IIf(Left$(BRW.LocationURL, 5) = "https", RGB(204, 255, 204), RGB(255, 255, 255))
End Sub

Private Sub cboAddress_DblClick()
    On Error Resume Next
    With cboAddress
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cboAddress_GotFocus()
    On Error Resume Next
    YouCantChangeMyAddyBarTextNow = True
    cboAddress_DblClick
End Sub

Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Shift = 4 Then cboAddress_DblClick
    If KeyCode = vbKeyReturn Then
        Select Case Shift
            Case 0 'none
                If EventRunning Then Exit Sub
                LoadFile ParsedAddy(cboAddress.Text)
                BRW.SetFocus
            Case 2 'ctrl
                LoadFile "http://www." & cboAddress.Text & ".com"
                BRW.SetFocus
        End Select
    End If
End Sub

Private Sub cboAddress_LostFocus()
    On Error Resume Next
    YouCantChangeMyAddyBarTextNow = False 'reset the tag
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    InitCommonControls
    frmMain.titBrowser.Visible = True
    BRW.Silent = frmMain.titBrowserSilent.Checked
    If GetSet("BRW_AutoFavsBarSwitch", "1") = "1" Then frmMain.GoToPath FavsPath, False
    frmMain.COF = Me.CurrentlyOpenFile
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    With frmMain
        .titBrowser.Visible = False
        If GetSet("BRW_AutoFavsBarSwitch", "1") = "1" Then .GoToPath GetSet("Recent_Path"), False
    End With
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.BRW.Silent = True
    BRW.GoHome
    Me.Icon = frmMain.Icon
    SkinForm Me
    SkinFormEx Me

    SHAutoComplete cboAddress.hWnd, &H0

    Dim I As Integer
    For I = 0 To 5 Step 1
        btnBrw(I + 1).Left = btnBrw(I).Left + btnBrw(I).Width + 15
    Next
    Form_Resize
    
    DSA 17
    EventSound "WinOpen"
End Sub

Private Sub BRW_DownloadComplete()
    On Error Resume Next
    EventRunning = True
    SetAddy BRW.LocationURL
    EventRunning = False
    CCaption BRW.LocationName, Me
End Sub

Public Function FavAddy(WhichFile As String, Optional SignalFromFavForm As Boolean) As String
    On Error Resume Next

        Dim A As String
        A = ReadINI("DEFAULT", "BASEURL", WhichFile)
        If Len(A) > 0 Then
            FavAddy = A
        Else
            A = ReadINI("InternetShortcut", "URL", WhichFile)
            If Len(A) > 0 Then
                FavAddy = A
            End If
        End If
End Function

Private Sub BRW_NavigateComplete2(ByVal pDisp As Object, url As Variant)
    On Error Resume Next
    Dim I As Integer
    Dim bFound As Boolean
    CCaption BRW.LocationName, Me
    EventRunning = True
    
    Form_Resize
    SetAddy BRW.LocationURL
    If GetSet("BRW_Log", "1") = "1" Then
        If BRW.LocationURL <> LastLogMsg Then
            Log BRW.LocationURL, True
            LastLogMsg = BRW.LocationURL
        End If
    End If
    
    EventRunning = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If GetSet("Browser_RestoreSingleSession", "1") = "1" Then
        SaveSet "Browser_LastURL", BRW.LocationURL
    Else
        SaveSet "Browser_LastURL", "" 'security, leave nothing there
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    btnGo.Move Me.ScaleWidth - btnGo.Width
    cboAddress.Width = Me.ScaleWidth - cboAddress.Left - btnGo.Width
    With BRW
        .Move 0, 480, Me.ScaleWidth, Me.ScaleHeight - 480 '480 being the top
    End With
End Sub

Public Function LoadFile(TheFN As String) As Long
    On Error Resume Next
    BRW.Navigate2 AddRecentItem(TheFN)
    CurrentlyOpenFile = TheFN
    CCaption FileNameOnly(TheFN), Me 'TrimFileNameLOL(TheFN), Me
    Me.Tag = TheFN
    Me.Show
    Form_Resize
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Form_Deactivate
    
    EventSound "WinClose"
    
End Sub

Function SetAddy(WhatText As String)
    On Error Resume Next
    If YouCantChangeMyAddyBarTextNow = False Then
        cboAddress.Text = WhatText
    End If
End Function
