VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  '¥­­±
   BackColor       =   &H00404040&
   BorderStyle     =   0  '¨S¦³®Ø½u
   Caption         =   "frmSplash"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.Timer T1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   2520
   End
   Begin ProFile.F F1 
      Left            =   720
      Top             =   2640
      _ExtentX        =   979
      _ExtentY        =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Thinc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   1  '¾a¥k¹ï»ô
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "2.41"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4080
      TabIndex        =   1
      Top             =   2280
      Width           =   315
   End
   Begin VB.Label lblProgName 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "ProFile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   915
   End
   Begin VB.Shape S1 
      BorderColor     =   &H00808080&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image imgSplash 
      Height          =   2895
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   48
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   975
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ForceClick As Boolean

Private Sub F1_FadeInReady()
    If ForceClick = False Then T1.Enabled = True
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    InitCommonControls
    If GetSet("Fade", "1") = "1" Then
        F1.FadeIn
    Else
        Sleep 750
        F1.FadeIn
    End If
End Sub

Private Sub Form_Click()
    imgSplash_Click
End Sub

Private Sub Form_Deactivate()
    F1.FadeOut
End Sub

Private Sub Form_Load()
    On Error Resume Next
    OnTop Me.hWnd, True
    
    S1.Move 0, 0, ScaleWidth, ScaleHeight
    imgSplash.Move 0, 0, ScaleWidth, ScaleHeight
    
    SkinForm Me
    SkinFormEx Me
    
    Randomize
    
    Me.BackColor = RGB(Int(Rnd * 30), Int(Rnd * 30), Int(Rnd * 30))
    
    Label1.Caption = MyVer

    F1.PrepareFade

End Sub

Private Sub imgSplash_Click()
    Unload Me 'so many unload me commands.... not bothered to make them an array.
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub

Private Sub lblCaption_Click()
    Unload Me
End Sub

Private Sub lblProgName_Click()
    Unload Me
End Sub

Private Sub T1_Timer()
    On Error Resume Next
    Unload Me
End Sub
