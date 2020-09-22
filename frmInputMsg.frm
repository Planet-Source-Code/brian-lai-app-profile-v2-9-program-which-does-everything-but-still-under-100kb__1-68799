VERSION 5.00
Begin VB.Form frmInputMsg 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   " "
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInputMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.CommandButton btnYN 
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton btnYN 
      Caption         =   "&Yes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton BTN 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CheckBox CHK 
      Caption         =   "Always use this answer"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
   End
   Begin ProFile.F F1 
      Left            =   120
      Top             =   1440
      _ExtentX        =   979
      _ExtentY        =   450
   End
   Begin VB.Image IMG 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Label LBL 
      BackStyle       =   0  '³z©ú
      Height          =   1530
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Image IMGbkg 
      Height          =   2175
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5055
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
      Left            =   3840
      TabIndex        =   5
      Top             =   -240
      Width           =   1440
   End
End
Attribute VB_Name = "frmInputMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ORLY As Integer

Public Function MyMsgBoxEx(Prompt As String, SaveNum As Integer, Optional MSGStyle As VbMsgBoxStyle = vbOKOnly, Optional titLE As String, Optional HideCheckBox As Boolean) As VbMsgBoxResult
    On Error Resume Next
    Dim I As Integer
    I = Val(GetSet("DSA" & SaveNum, "0"))
    If I <> 0 Then 'if this message is set not to show again
        MyMsgBoxEx = ValToResult(I)
        'SStatus "DSA: " & SaveNum
        
        EventSound "MSGSkip"
        
        Exit Function 'then exit
    End If
    ShowButtonType IIf(MSGStyle = vbOKOnly, 0, 1) 'changes buttons
    LBL.Caption = Prompt
    CHK.Visible = Not HideCheckBox 'shows and hides the checkbox
    CCaption titLE, Me
    Me.Tag = SaveNum
    Me.Show 1
    MyMsgBoxEx = ValToResult(ORLY)
    'Debug.Print "ORLY 3:" & ORLY
End Function

Private Sub BTN_Click()
    On Error Resume Next
    ORLY = 1 'say 2 is YES and 3 is NO
    SaveSet "DSA" & Me.Tag, IIf(CHK.Value = 0, 0, ORLY)
    Unload Me
End Sub

Private Sub btnYN_Click(Index As Integer)
    On Error Resume Next
    ORLY = Index + 2 'say 2 is YES and 3 is NO
    SaveSet "DSA" & Me.Tag, IIf(CHK.Value = 0, 0, ORLY)
    Unload Me
End Sub

Private Sub Form_Activate()
    InitCommonControls
    F1.FadeIn
End Sub

Private Sub Form_Deactivate()
    F1.FadeOut
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    SkinForm Me
    SkinFormEx Me

    EventSound "WinOpen"
    
    F1.PrepareFade

End Sub

Sub ShowButtonType(Optional Which As Integer = 0)
    On Error Resume Next
    BTN.Visible = (Which = 0)
    btnYN(0).Visible = (Which <> 0)
    btnYN(1).Visible = (Which <> 0)
End Sub

Function ValToResult(Valu As Integer) As Integer
    On Error Resume Next
        Select Case Valu
        Case 1
            ValToResult = Val(vbOK)
        Case 2
            ValToResult = Val(vbYes)
        Case 3
            ValToResult = Val(vbNo)
    End Select
End Function

Private Sub Form_Unload(Cancel As Integer)

    EventSound "WinClose"

End Sub

