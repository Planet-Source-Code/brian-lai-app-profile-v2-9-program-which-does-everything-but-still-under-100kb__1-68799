VERSION 5.00
Begin VB.Form frmIMG 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Image"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIMG.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   4695
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.Image IMG 
      Appearance      =   0  '¥­­±
      BorderStyle     =   1  '³æ½u©T©w
      Height          =   1095
      Left            =   0
      OLEDropMode     =   1  '¤â°Ê
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image imgBG 
      Height          =   600
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "frmIMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldX As Long, OldY As Long

Public CurrentlyOpenFile As String

Dim BGTempRemove As Boolean

Private Sub Form_Activate()
    On Error Resume Next
        InitCommonControls
    frmMain.titImage.Visible = True
    'DoStretch frmMain.titImageStretch.Checked
    Form_Resize 'so nothing is outdated
    Me.imgBG.Visible = frmMain.titImageCheckers.Checked
    IMG.BorderStyle = frmMain.titImageBorder.Checked
    frmMain.COF = Me.CurrentlyOpenFile
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    frmMain.titImage.Visible = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
'    Mod32BitIcon.SetIcon Me.hwnd, "AAA"
    Form_Resize
    
    Me.Icon = frmMain.Icon
    SkinForm Me
    SkinFormEx Me

    EventSound "WinOpen"

End Sub

Public Sub Form_Resize()
    On Error Resume Next
    DoStretch Val(GetSet("Image_Stretch", "2"))
    imgBG.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Form_Deactivate
    
    EventSound "WinClose"
    
End Sub

Public Function LoadFile(TheFN As String) As Long
    On Error Resume Next
        IMG.Picture = LoadPicture(AddRecentItem(TheFN))
        
    CurrentlyOpenFile = TheFN
    Me.Tag = TheFN
    CCaption FileNameOnly(TheFN), Me 'TrimFileNameLOL(TheFN), Me
    Me.Show
    Form_Resize
End Function

Public Function DoStretch(TehMode As Integer)
    On Error Resume Next
    Dim W As Long, H As Long
    Dim Rwh As Double
    With IMG
        Select Case TehMode
            Case 0 'none
                .Stretch = False
            Case 1 'full
                .Stretch = True
                .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            Case 2 'natural
                .Stretch = False: .Visible = False
                W = .Width: H = .Height
                If W / H >= 1 Then
AdjWidth:
                    .Width = Me.ScaleWidth
                    Rwh = Me.ScaleWidth / W
                    .Height = .Height * Rwh
                    If .Height > Me.ScaleHeight Then
                        W = .Width: H = .Height
                        GoTo Adjheight
                    End If
                ElseIf W / H < 1 Then
Adjheight:
                    .Height = Me.ScaleHeight
                    Rwh = Me.ScaleHeight / H
                    .Width = .Width * Rwh
                    If .Width > Me.ScaleWidth Then
                        W = .Width: H = .Height
                        GoTo AdjWidth
                    End If
                End If
                .Stretch = True: .Visible = True
        End Select
        CenterPic
    End With
End Function

Public Function CenterPic()
    On Error Resume Next
    With IMG
        .Move (Me.ScaleWidth - .Width) / 2, (Me.ScaleHeight - .Height) / 2
    End With
End Function

Private Sub IMG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    OldX = X
    OldY = Y
    If GetSet("IMG_DragChecker", "0") = "0" And imgBG.Visible = True Then
        imgBG.Visible = False
        BGTempRemove = True
    End If
    If Button = 2 Then
        PopupMenu frmMain.titImage
    End If
End Sub

Private Sub IMG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        IMG.Move IMG.Left - OldX + X, IMG.Top - OldY + Y
    End If
End Sub

Private Sub IMG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If BGTempRemove Then imgBG.Visible = True
End Sub

Private Sub IMG_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    LoadFile Data.Files.Item(1) 'just load it
End Sub
