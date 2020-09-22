VERSION 5.00
Begin VB.Form frmFile 
   Caption         =   "File Browser"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFile.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   8325
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.PictureBox pic 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   315
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   3045
      TabIndex        =   11
      Top             =   0
      Width           =   3045
      Begin VB.TextBox txtSearch 
         Height          =   315
         Index           =   0
         Left            =   1470
         TabIndex        =   12
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '³z©ú
         Caption         =   "Search:"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   30
         Width           =   975
      End
   End
   Begin VB.FileListBox Fil 
      Height          =   2610
      Hidden          =   -1  'True
      Left            =   3840
      System          =   -1  'True
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox Lst 
      Height          =   4980
      Index           =   4
      IntegralHeight  =   0   'False
      ItemData        =   "frmFile.frx":000C
      Left            =   6720
      List            =   "frmFile.frx":000E
      TabIndex        =   9
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox Lst 
      Height          =   4980
      Index           =   3
      IntegralHeight  =   0   'False
      ItemData        =   "frmFile.frx":0010
      Left            =   5160
      List            =   "frmFile.frx":0012
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox Lst 
      Height          =   4980
      Index           =   2
      IntegralHeight  =   0   'False
      ItemData        =   "frmFile.frx":0014
      Left            =   3600
      List            =   "frmFile.frx":0016
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox Lst 
      Height          =   4980
      Index           =   1
      IntegralHeight  =   0   'False
      ItemData        =   "frmFile.frx":0018
      Left            =   2040
      List            =   "frmFile.frx":001A
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox Lst 
      Height          =   4980
      Index           =   0
      IntegralHeight  =   0   'False
      ItemData        =   "frmFile.frx":001C
      Left            =   0
      List            =   "frmFile.frx":001E
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblCat 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Caption         =   "Modified"
      Height          =   255
      Index           =   4
      Left            =   6720
      TabIndex        =   8
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblCat 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Caption         =   "Accessed"
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblCat 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Caption         =   "Created"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblCat 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Caption         =   "Size"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblCat 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SearchStr = "Search..."

Private Sub Form_Activate()
    On Error Resume Next
    InitCommonControls
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Icon = frmMain.Icon
    SkinForm Me
    SkinFormEx Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim I As Integer
    
    lblCat(0).Width = Me.ScaleWidth * 0.4
    For I = 1 To lblCat.UBound Step 1
        lblCat(I).Move lblCat(I - 1).Left + lblCat(I - 1).Width, lblCat(0).Top, Me.ScaleWidth * 0.15
    Next
    
    Lst(0).Move 0, lblCat(0).Top + lblCat(0).Height, lblCat(0).Width, Me.ScaleHeight - lblCat(0).Top - lblCat(0).Height
    For I = 1 To 4 Step 1
        Lst(I).Move Lst(I - 1).Left + Lst(I - 1).Width, lblCat(I).Top + lblCat(I).Height, lblCat(I).Width, Me.ScaleHeight - lblCat(I).Top - lblCat(I).Height
    Next
    For I = Lst.LBound To Lst.UBound Step 1
        If I <> Lst.UBound Then ShowScrollBar Lst(I).hWnd, SB_VERT, False
    Next
    pic.Left = Me.ScaleWidth - pic.Width
End Sub

Private Sub Lst_Click(Index As Integer)
    On Error Resume Next
    Dim I As Integer
    For I = Lst.LBound To Lst.UBound Step 1
        Lst(I).TopIndex = Lst(Index).TopIndex
        Lst(I).ListIndex = Lst(Index).ListIndex
        If I <> Lst.UBound Then ShowScrollBar Lst(I).hWnd, SB_VERT, False
    Next
End Sub

Private Sub Lst_DblClick(Index As Integer)
    On Error Resume Next
    frmMain.DecideOnType FindPath(Fil.Path, Lst(0).List(Lst(0).ListIndex))
End Sub

Private Sub Lst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lst_Click Index
    
    On Error Resume Next
    Dim Ix As Long, I As Integer
    Dim Mx As Long, My As Long
    Dim K As Double
    
    Mx = CLng(X / Screen.TwipsPerPixelX)
    My = CLng(Y / Screen.TwipsPerPixelY)
    Ix = SendMessage(Lst(Index).hWnd, LB_ITEMFROMPOINT, 0, ByVal ((My * 65536) + Mx))
    If Button = 0 Then
        K = Round(Val(FileLen(FindPath(Fil.Path, Fil.List(Ix)))) / 1024 / 1024, 2)
        For I = Lst.LBound To Lst.UBound Step 1
            Lst(I).ToolTipText = Lst(0).List(Ix) & " (" & K & " MB)"
        Next
    End If
End Sub

Private Sub Lst_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
End Sub

Private Sub Lst_Scroll(Index As Integer)
    Lst_Click Index
End Sub

Public Function LoadFile(TheFN As String) As Long
    On Error Resume Next
    LoadFile = LoadPath(PathOnly(TheFN))
End Function

Public Function LoadPath(ThePath As String) As Long
    On Error Resume Next
    Dim I As Long, L As Long
        
    Fil.Path = ThePath
    
    For I = Lst.LBound To Lst.UBound Step 1
        Lst(I).Clear
        Lst(I).Visible = False
    Next
    
    For I = 0 To Fil.ListCount - 1 Step 1
        For L = Lst.LBound To Lst.UBound Step 1
            Lst(L).AddItem " " 'buffer
        Next
        Lst(0).List(I) = Fil.List(I)
        Lst(1).List(I) = Round(FileLen(FindPath(Fil.Path, Fil.List(I))) / 1024, 0) & " KB"
        Lst(2).List(I) = GetFileDate((FindPath(Fil.Path, Fil.List(I))), Created)
        Lst(3).List(I) = GetFileDate((FindPath(Fil.Path, Fil.List(I))), Accessed)
        Lst(4).List(I) = GetFileDate((FindPath(Fil.Path, Fil.List(I))), Modified)
    Next
    
    For I = Lst.LBound To Lst.UBound Step 1
        Lst(I).Visible = True
        If I <> Lst.UBound Then ShowScrollBar Lst(I).hWnd, SB_VERT, False
    Next
    Me.Show
End Function

Private Sub txtSearch_Change(Index As Integer)
    On Error Resume Next
    With txtSearch(Index)
        If .Text = SearchStr Or .Text = "" Then
            Fil.Pattern = "*"
        Else
            Fil.Pattern = "*" & .Text & "*"
        End If
    End With
    LoadPath Fil.Path
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    TBFocus txtSearch(0), True, SearchStr
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
    TBFocus txtSearch(0), False, SearchStr
End Sub
