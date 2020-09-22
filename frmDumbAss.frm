VERSION 5.00
Begin VB.Form frmDumbAss 
   Caption         =   "Drag something here"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   Icon            =   "frmDumbAss.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   OLEDropMode     =   1  '¤â°Ê
   ScaleHeight     =   5430
   ScaleWidth      =   6765
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.PictureBox Picture1 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   6765
      TabIndex        =   1
      Top             =   5175
      Width           =   6765
      Begin VB.CheckBox chkWhenStart 
         Caption         =   "Show this when ProFile starts"
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
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   4575
      End
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  '¥­­±
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   240
      OLEDropMode     =   1  '¤â°Ê
      ScaleHeight     =   4185
      ScaleWidth      =   6225
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.Label Label1 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "Close this window to gain full access to ProFile."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         OLEDropMode     =   1  '¤â°Ê
         TabIndex        =   5
         Top             =   3240
         Width           =   6255
      End
      Begin VB.Label Label1 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "Drag a file TO here."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         OLEDropMode     =   1  '¤â°Ê
         TabIndex        =   4
         Top             =   2880
         Width           =   6255
      End
      Begin VB.Label Label1 
         Alignment       =   2  '¸m¤¤¹ï»ô
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   72
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1590
         Index           =   0
         Left            =   2280
         OLEDropMode     =   1  '¤â°Ê
         TabIndex        =   3
         ToolTipText     =   "Drag file here"
         Top             =   1320
         Width           =   1965
      End
   End
End
Attribute VB_Name = "frmDumbAss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkWhenStart_Click()
    On Error Resume Next
    SaveSet "OpenOnIdiot", Str(chkWhenStart.Value)
End Sub

Private Sub Form_Activate()
    InitCommonControls
    frmMain.picBrw.Visible = False
End Sub

Private Sub Form_Deactivate()
    frmMain.picBrw.Visible = True
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Icon = frmMain.Icon
    SkinForm Me
    SkinFormEx Me
    
    chkWhenStart.Value = Val(GetSet("OpenOnIdiot", "1"))
    
    EventSound "WinOpen"
'    frmMain.picBrw.Width = frmMain.Dragger(1).Width 'added ... so noobs dont mess around with the files.
'    frmMain.picBrw.Visible = False

    Form_Activate
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
        frmMain.MDIForm_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    P1.Move (Me.ScaleWidth - P1.Width) / 2, (Me.ScaleHeight - P1.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    EventSound "WinClose"
    Form_Deactivate
    
End Sub

Private Sub Label1_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    P1_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub P1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
        frmMain.MDIForm_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
