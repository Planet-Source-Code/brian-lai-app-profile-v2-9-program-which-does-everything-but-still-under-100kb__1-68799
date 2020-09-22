VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmWMP 
   AutoRedraw      =   -1  'True
   Caption         =   "Media"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWMP.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   7560
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.PictureBox picSet 
      Appearance      =   0  '¥­­±
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   1200
      ScaleHeight     =   2505
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   5025
      Begin VB.CommandButton btnExec 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton btnExec 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   7
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtSpd 
         Alignment       =   1  '¾a¥k¹ï»ô
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Text            =   "1"
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Playback speed: (0.5 ~ 16)"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2565
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "This tool changes the way this media is played."
         Height          =   450
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   600
         Width           =   3990
         WordWrap        =   -1  'True
      End
      Begin VB.Image IMG 
         Height          =   480
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Media Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00AB7013&
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   72
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1455
         Left            =   3720
         TabIndex        =   4
         Top             =   -120
         Width           =   1440
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   999
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8070
      _cy             =   4683
   End
End
Attribute VB_Name = "frmWMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CurrentlyOpenFile As String

Private Sub btnExec_Click(Index As Integer)
    On Error Resume Next
    If Index = 0 Then WMP.settings.Rate = Val(txtSpd.Text)
    picSet.Visible = False
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    InitCommonControls
    frmMain.titMedia.Visible = True
    frmMain.COF = Me.CurrentlyOpenFile
    SStatus WMP.url, vbInformation
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    frmMain.titMedia.Visible = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
'    Mod32BitIcon.SetIcon Me.hwnd, "AAA"
    'settings
    Me.Icon = frmMain.Icon
    SkinForm Me
    SkinFormEx Me

    WMP.stretchToFit = GetSet("Media_Stretch", "1")
    WMP.uiMode = IIf(GetSet("Media_Controls", "1") = "1", "full", "none")
    
    EventSound "WinOpen"
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    WMP.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    picSet.Move (Me.ScaleWidth - picSet.Width) / 2, (Me.ScaleHeight - picSet.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Form_Deactivate

    EventSound "WinClose"

End Sub

Public Function LoadFile(TheFN As String) As Long
    On Error Resume Next
    Dim K As String, L As String
    WMP.url = AddRecentItem(TheFN)
    CCaption FileNameOnly(TheFN), Me 'TrimFileNameLOL(TheFN), Me
    
    PSMThis WMP.currentMedia.Name
    
    SaveSet "Media_Last", TheFN
    
    CurrentlyOpenFile = TheFN
    Me.Tag = TheFN
    Me.Show
    SStatus Me.Name & " opened " & TheFN, vbInformation
End Function

Private Sub WMP_MediaChange(ByVal Item As Object)
    On Error Resume Next
    PSMThis Item.Name
    CCaption Item.Name, Me
End Sub

Private Sub WMP_MediaError(ByVal pMediaObject As Object)
    On Error Resume Next
    SStatus "Error when trying to play " & WMP.url, vbCritical
End Sub

Public Sub PSMThis(What As String)
    On Error Resume Next
    Dim K As String
    If GetSet("Sync_PSM", "0") = "1" Then 'sync MSN PSM
        K = GetSet("PSM", "%n (%l)")
        Select Case GetSet("PSM_ShowCredit", "1")
            Case "0"
                'whatever
            Case "1"
                K = App.ProductName & " - " & K
            Case "2"
                K = App.CompanyName & " - " & K
            Case Else
                'whatever
        End Select

        K = Replace(K, "%n", WMP.currentMedia.Name)
        K = Replace(K, "%t", Now())
        K = Replace(K, "%f", WMP.url)
        K = Replace(K, "%s", Round(Val(FileLen(WMP.url)) / 1024 / 1024, 2) & " MB")
        K = Replace(K, "%l", WMP.currentMedia.durationString)
        
        Call CMD6("psm " & K)
    End If
End Sub
