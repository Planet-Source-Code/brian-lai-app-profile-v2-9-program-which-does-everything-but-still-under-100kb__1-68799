VERSION 5.00
Begin VB.Form frmOptimize 
   BorderStyle     =   4  '³æ½u©T©w¤u¨ãµøµ¡
   Caption         =   "Optimizing Preferences"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.Frame Frame1 
      Caption         =   "Optimizing for appearance"
      Height          =   3615
      Index           =   1
      Left            =   4440
      TabIndex        =   2
      Top             =   840
      Width           =   4215
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   3255
         Index           =   1
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   3975
         TabIndex        =   5
         Top             =   240
         Width           =   3975
         Begin VB.CommandButton btnAdjust 
            Caption         =   "Optimize for appearance"
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   6
            Top             =   2880
            Width           =   3975
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   2775
            Index           =   1
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   3975
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Optimizing for speed"
      Height          =   3615
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4215
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   3255
         Index           =   0
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   3975
         TabIndex        =   3
         Top             =   240
         Width           =   3975
         Begin VB.CommandButton btnAdjust 
            Caption         =   "Optimize for program speed"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   2880
            Width           =   3975
         End
         Begin ProFile.F F1 
            Left            =   0
            Top             =   480
            _ExtentX        =   979
            _ExtentY        =   450
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   2775
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   3975
         End
      End
   End
   Begin VB.Image IMGbkg 
      Height          =   4575
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Takes effect after you start the program again."
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   3435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This tool will optimize your settings when you choose one of the two buttons below."
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmOptimize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAdjust_Click(Index As Integer)
    On Error Resume Next
    Dim K As String
    K = Str$(Index)
    If MsgBox("Are you sure you want to do that?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    SaveSet "Fade", K 'fading / transparency
    SaveSet "DropShadow", K 'shadows
    If Index = 0 Then _
        SaveSet "CTL_Flatten", K 'flat controls (disable only)
    SaveSet "OpenOnIdiot", K 'beginner dlg
    SaveSet "Splash", K 'splash
    SaveSet "SND_Toggle", K 'sounds
    SaveSet "BRW_AutoFavsBarSwitch", K 'favs bar
    SaveSet "Sync_PSM", K 'msn
    If Index = 0 Then
        SaveSet "Lang", "" 'langs (disable only)
        SaveSet "SkinFile", "" 'skin (disable only)
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    SkinForm Me
    SkinFormEx Me
    
    Label2(0).Caption = "Disables:" & vbCrLf & vbCrLf & _
                                  "- Fading, transparency and shadows" & vbCrLf & _
                                  "- Flat controls" & vbCrLf & _
                                  "- Beginner dialogs" & vbCrLf & _
                                  "- Splash window" & vbCrLf & _
                                  "- Sounds" & vbCrLf & _
                                  "- Automatic favorites bar" & vbCrLf & _
                                  "- Messenger song notifications" & vbCrLf & _
                                  "- Language packs" & vbCrLf & _
                                  "- Skins"
                                  
    Label2(1).Caption = "Enables:" & vbCrLf & vbCrLf & _
                                  "- Fading, transparency and shadows" & vbCrLf & _
                                  "- Beginner dialogs" & vbCrLf & _
                                  "- Splash window" & vbCrLf & _
                                  "- Sounds" & vbCrLf & _
                                  "- Automatic favorites bar" & vbCrLf & _
                                  "- Messenger song notifications"
End Sub

