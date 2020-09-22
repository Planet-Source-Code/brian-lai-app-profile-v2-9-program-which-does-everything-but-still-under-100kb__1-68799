VERSION 5.00
Begin VB.Form frmIdiot 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Welcome"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdiot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4095
      Index           =   0
      Left            =   1920
      ScaleHeight     =   4095
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Click ""Next"" to continue."
         Height          =   210
         Index           =   22
         Left            =   120
         TabIndex        =   58
         Top             =   3240
         Width           =   4770
         WordWrap        =   -1  'True
      End
      Begin VB.Image IMG 
         Height          =   480
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "And hopefully this will only appear once."
         Height          =   210
         Index           =   15
         Left            =   120
         TabIndex        =   38
         Top             =   2400
         Width           =   4770
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "This will take only 30 seconds."
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   4770
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "ProFile would like to know something about your IQ before you using it."
         Height          =   420
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   4800
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         BackStyle       =   0  '³z©ú
         Caption         =   "What is this?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label lblNotification 
         BackStyle       =   0  '³z©ú
         Caption         =   "Welcome to the wizard!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000001&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4815
      ScaleWidth      =   1815
      TabIndex        =   50
      Top             =   0
      Width           =   1815
      Begin VB.Label lblSteps 
         BackStyle       =   0  '³z©ú
         Caption         =   "Welcome"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   57
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblSteps 
         BackStyle       =   0  '³z©ú
         Caption         =   "Question 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblSteps 
         BackStyle       =   0  '³z©ú
         Caption         =   "Question 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblSteps 
         BackStyle       =   0  '³z©ú
         Caption         =   "Question 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   54
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblSteps 
         BackStyle       =   0  '³z©ú
         Caption         =   "Ready"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   53
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblSteps 
         BackStyle       =   0  '³z©ú
         Caption         =   "Results"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   52
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblSteps 
         BackStyle       =   0  '³z©ú
         Caption         =   "Results"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   51
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.CommandButton btnTabProc 
      Caption         =   "Back"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   6
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton btnTabProc 
      Caption         =   "Next"
      Default         =   -1  'True
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4095
      Index           =   5
      Left            =   1920
      ScaleHeight     =   4095
      ScaleWidth      =   4935
      TabIndex        =   31
      Top             =   120
      Width           =   4935
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   $"frmIdiot.frx":000C
         Height          =   630
         Index           =   19
         Left            =   120
         TabIndex        =   44
         Top             =   1560
         Width           =   4665
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "Retarded"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   37
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "What we believe in is that you do have some talents in life that may enhance your mental skills. "
         Height          =   420
         Index           =   9
         Left            =   135
         TabIndex        =   36
         Top             =   960
         Width           =   4665
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         BackStyle       =   0  '³z©ú
         Caption         =   "Go get a life and try again."
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   35
         Top             =   2400
         Width           =   4455
      End
      Begin VB.Label lblNotification 
         BackStyle       =   0  '³z©ú
         Caption         =   "You are"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4095
      Index           =   2
      Left            =   1920
      ScaleHeight     =   4095
      ScaleWidth      =   4935
      TabIndex        =   15
      Top             =   120
      Width           =   4935
      Begin VB.Frame FrameAns 
         Height          =   3615
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   4695
         Begin VB.PictureBox picFrameAns 
            BorderStyle     =   0  '¨S¦³®Ø½u
            Height          =   3255
            Index           =   1
            Left            =   120
            ScaleHeight     =   3255
            ScaleWidth      =   4455
            TabIndex        =   17
            Top             =   240
            Width           =   4455
            Begin VB.OptionButton OptAns 
               Caption         =   "You can sit on Windows(R)"
               Height          =   375
               Index           =   7
               Left            =   120
               TabIndex        =   21
               Top             =   120
               Width           =   4215
            End
            Begin VB.OptionButton OptAns 
               Caption         =   "It is like a hole punch"
               Height          =   375
               Index           =   6
               Left            =   120
               TabIndex        =   20
               Top             =   600
               Width           =   4215
            End
            Begin VB.OptionButton OptAns 
               Caption         =   "Pens are stored in a Windows(R)"
               Height          =   375
               Index           =   5
               Left            =   120
               TabIndex        =   19
               Top             =   1080
               Width           =   4215
            End
            Begin VB.OptionButton OptAns 
               Caption         =   "I use Windows(R) for this computer"
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   18
               Top             =   1560
               Width           =   4215
            End
         End
      End
      Begin VB.Label lblNotification 
         BackStyle       =   0  '³z©ú
         Caption         =   "Describe Windows(R):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4095
      Index           =   1
      Left            =   1920
      ScaleHeight     =   4095
      ScaleWidth      =   4935
      TabIndex        =   7
      Top             =   120
      Width           =   4935
      Begin VB.Frame FrameAns 
         Height          =   3615
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4695
         Begin VB.PictureBox picFrameAns 
            BorderStyle     =   0  '¨S¦³®Ø½u
            Height          =   3255
            Index           =   0
            Left            =   120
            ScaleHeight     =   3255
            ScaleWidth      =   4455
            TabIndex        =   10
            Top             =   240
            Width           =   4455
            Begin VB.OptionButton OptAns 
               Caption         =   "A machine with a screen"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   14
               Top             =   1560
               Width           =   4215
            End
            Begin VB.OptionButton OptAns 
               Caption         =   "What you see when you are swimming"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   13
               Top             =   1080
               Width           =   4215
            End
            Begin VB.OptionButton OptAns 
               Caption         =   "What you drink"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   12
               Top             =   600
               Width           =   4215
            End
            Begin VB.OptionButton OptAns 
               Caption         =   "Close to a cat"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   11
               Top             =   120
               Width           =   4215
            End
         End
      End
      Begin VB.Label lblNotification 
         BackStyle       =   0  '³z©ú
         Caption         =   "What is a computer?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4095
      Index           =   3
      Left            =   1920
      ScaleHeight     =   4095
      ScaleWidth      =   4935
      TabIndex        =   23
      Top             =   120
      Width           =   4935
      Begin VB.Frame FrameAns 
         Height          =   3615
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   4695
         Begin VB.PictureBox picFrameAns 
            BorderStyle     =   0  '¨S¦³®Ø½u
            Height          =   3255
            Index           =   2
            Left            =   120
            ScaleHeight     =   3255
            ScaleWidth      =   4455
            TabIndex        =   25
            Top             =   240
            Width           =   4455
            Begin VB.OptionButton OptAns 
               Caption         =   "None of the above"
               Height          =   375
               Index           =   11
               Left            =   120
               TabIndex        =   29
               Top             =   1560
               Width           =   4215
            End
            Begin VB.OptionButton OptAns 
               Caption         =   "Windows ShortHorn"
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   28
               Top             =   1080
               Width           =   4215
            End
            Begin VB.OptionButton OptAns 
               Caption         =   "Version 1"
               Height          =   375
               Index           =   9
               Left            =   120
               TabIndex        =   27
               Top             =   600
               Width           =   4215
            End
            Begin VB.OptionButton OptAns 
               Caption         =   "I don't know."
               Height          =   375
               Index           =   8
               Left            =   120
               TabIndex        =   26
               Top             =   120
               Width           =   4215
            End
         End
      End
      Begin VB.Label lblNotification 
         BackStyle       =   0  '³z©ú
         Caption         =   "What is this version of Windows?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4095
      Index           =   4
      Left            =   1920
      ScaleHeight     =   4095
      ScaleWidth      =   4935
      TabIndex        =   39
      Top             =   120
      Width           =   4935
      Begin VB.Label lblNotification 
         BackStyle       =   0  '³z©ú
         Caption         =   "ProFile is ready"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   4935
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "We have collected enough information about you and whether or not you will be able to use this program."
         Height          =   420
         Index           =   18
         Left            =   120
         TabIndex        =   42
         Top             =   1200
         Width           =   4770
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Click next to continue."
         Height          =   210
         Index           =   17
         Left            =   120
         TabIndex        =   41
         Top             =   1920
         Width           =   4770
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "If you want to change any answers, click back."
         Height          =   210
         Index           =   16
         Left            =   120
         TabIndex        =   40
         Top             =   2400
         Width           =   4770
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4095
      Index           =   6
      Left            =   1920
      ScaleHeight     =   4095
      ScaleWidth      =   4935
      TabIndex        =   33
      Top             =   120
      Width           =   4935
      Begin VB.CheckBox chkWhenStart 
         Caption         =   "Are you a noob?"
         Height          =   255
         Left            =   0
         TabIndex        =   48
         Top             =   3360
         Width           =   4575
      End
      Begin VB.Label lblNotification 
         BackStyle       =   0  '³z©ú
         Caption         =   "You can think about this later."
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   21
         Left            =   240
         TabIndex        =   49
         Top             =   3600
         Width           =   4455
      End
      Begin VB.Label lblNotification 
         BackStyle       =   0  '³z©ú
         Caption         =   "Click finish to close this wizard."
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   47
         Top             =   1800
         Width           =   4455
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   $"frmIdiot.frx":0096
         Height          =   630
         Index           =   12
         Left            =   135
         TabIndex        =   46
         Top             =   960
         Width           =   4665
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "Smart Enough"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   240
         TabIndex        =   45
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label lblNotification 
         BackStyle       =   0  '³z©ú
         Caption         =   "You are"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   4935
      End
   End
   Begin ProFile.F F1 
      Left            =   2160
      Top             =   4440
      _ExtentX        =   979
      _ExtentY        =   450
   End
End
Attribute VB_Name = "frmIdiot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TabOnTop As Integer, TabOnTop2 As Integer

Private Sub btnTabProc_Click(Index As Integer)
    On Error Resume Next
    Dim I As Integer
    
        If Index = 1 And TabOnTop < 1 Then
            btnTabProc(1).Enabled = False
            Exit Sub
        Else
            btnTabProc(1).Enabled = True
        End If
        
    If btnTabProc(3).Caption = "Finish" Then
        If Index = 3 Then Unload Me
        If TabOnTop2 = 5 Then
            'Debug.Print "lol"
            End
        Else
            SaveSet "FirstRun", "Not" 'Prevents reappearance
        End If
    Else

        TabOnTop = TabOnTop + (Index - 2)
        
        For I = 0 To lblSteps.UBound Step 1
            lblSteps(I).FontBold = (TabOnTop = I)
        Next
        If TabOnTop > 4 Then 'if the questions are over
            If OptAns(3).Value And OptAns(4).Value And OptAns(11).Value Then 'if the pro scored them all right
                picPage(6).ZOrder 0
                TabOnTop2 = 6
            Else 'idiot
                picPage(5).ZOrder 0
                TabOnTop2 = 5
            End If
            btnTabProc(3).Caption = "Finish"
            btnTabProc(1).Enabled = False
        Else
            picPage(TabOnTop).ZOrder 0
        End If
    End If
End Sub

Private Sub chkWhenStart_Click()
    SaveSet "OpenOnIdiot", Str(chkWhenStart.Value)
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
    
    F1.PrepareFade
    
    Me.Icon = frmMain.Icon
    SkinForm Me
    SkinFormEx Me
    chkWhenStart.Value = Val(GetSet("OpenOnIdiot", "1"))
    
    EventSound "WinOpen"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    EventSound "WinClose"

End Sub
