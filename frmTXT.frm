VERSION 5.00
Begin VB.Form frmTXT 
   AutoRedraw      =   -1  'True
   Caption         =   "Text"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00AB7013&
   Icon            =   "frmTXT.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   5910
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.PictureBox picROT 
      Appearance      =   0  '¥­­±
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   480
      ScaleHeight     =   2865
      ScaleWidth      =   5025
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame Frame1 
         Caption         =   "Options"
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   4815
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  '¨S¦³®Ø½u
            Height          =   855
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   4575
            TabIndex        =   6
            Top             =   240
            Width           =   4575
            Begin VB.TextBox txtVal 
               Alignment       =   1  '¾a¥k¹ï»ô
               Height          =   315
               Left            =   1560
               TabIndex        =   9
               Text            =   "4"
               Top             =   480
               Width           =   3015
            End
            Begin VB.OptionButton opt 
               Caption         =   "Decrypt"
               Height          =   255
               Index           =   1
               Left            =   2280
               TabIndex        =   8
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton opt 
               Caption         =   "Encrypt"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   7
               Top             =   0
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Encryption table:"
               Height          =   255
               Left            =   0
               TabIndex        =   10
               Top             =   480
               Width           =   1695
            End
         End
      End
      Begin VB.CommandButton btnExec 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   3
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton btnExec 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   2
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Text Encryption"
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
         TabIndex        =   11
         Top             =   120
         Width           =   1665
      End
      Begin VB.Image IMG 
         Height          =   480
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '³z©ú
         Caption         =   "This tool can encrypt and decrypt text which many other Thinc programs use."
         Height          =   495
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   480
         Width           =   4215
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
         TabIndex        =   12
         Top             =   -240
         Width           =   1440
      End
   End
   Begin VB.TextBox txtBox 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      OLEDropMode     =   1  '¤â°Ê
      ScrollBars      =   2  '««ª½±²¶b
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmTXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CurrentlyOpenFile As String

Private Sub btnExec_Click(Index As Integer)
    'On Error Resume Next
    If Int(txtVal.Text) <= 0 Or Int(txtVal.Text) >= 10 Then
        MsgBox "You must enter an integer value higher than 0 and lower than 10."
        Exit Sub
    End If
    
    If Index > 0 Then txtBox.Text = Encrypt(txtBox.Text, opt(0).Value, Int(txtVal.Text))
    
    picROT.Visible = False
    
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    InitCommonControls
    frmMain.titText.Visible = True
    frmMain.COF = Me.CurrentlyOpenFile
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    frmMain.titText.Visible = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
'    Mod32BitIcon.SetIcon Me.hwnd, "AAA"
    Me.Icon = frmMain.Icon
    SkinForm Me
    SkinFormEx Me

    EventSound "WinOpen"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If Right$(Me.Caption, 1) = "*" Then 'unsaved eh
        Dim A As VbMsgBoxResult
        A = MsgBox("Do you want to save this file first?", vbYesNoCancel + vbQuestion)
        Select Case A
            Case vbYes
                frmMain.titTextFileSave_Click
            Case vbNo
                'do nothing?
            Case vbCancel
                Cancel = 1
        End Select
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtBox.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    picROT.Move (Me.ScaleWidth - picROT.Width) / 2, (Me.ScaleHeight - picROT.Height) / 2
End Sub

Public Function LoadFile(TheFN As String) As Long
    On Error Resume Next
    If FileLen(TheFN) > 64000 Then 'use for big files
        txtBox.Text = FileText(AddRecentItem(TheFN))
    Else
        Dim F As Integer
        Dim tmp As String, K As String
        F = FreeFile
        Open TheFN For Input As #F
            Do
                Line Input #F, tmp
                K = K & tmp & vbCrLf
            Loop Until EOF(F)
        Close #F
        txtBox.Text = K
    End If
    
    CurrentlyOpenFile = TheFN
    
    Me.Tag = TheFN
    CCaption FileNameOnly(TheFN), Me 'TrimFileNameLOL(TheFN), Me
    Me.Show
    SStatus Me.Name & " opened " & TheFN, vbInformation
End Function

Public Sub ChangeFont()
    On Error Resume Next
    'Dim Response As VbMsgBoxResult
    With txtBox
        SelectFont.mFontName = txtBox.FontName
        SelectFont.mFontSize = txtBox.FontSize
        SelectFont.mBold = txtBox.FontBold
        SelectFont.mFontColor = txtBox.ForeColor
        SelectFont.mItalic = txtBox.FontItalic
        SelectFont.mStrikethru = txtBox.FontStrikethru
        SelectFont.mUnderline = txtBox.FontUnderline
        
        ShowFont
        .FontName = SelectFont.mFontName
        .FontSize = SelectFont.mFontSize
        .FontBold = SelectFont.mBold
        .FontItalic = SelectFont.mItalic
        .FontStrikethru = SelectFont.mStrikethru
        .FontUnderline = SelectFont.mUnderline
        .ForeColor = SelectFont.mFontColor
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Form_Deactivate

    EventSound "WinClose"

End Sub

Private Sub txtBox_Change()
    On Error Resume Next
    If Right$(Me.Caption, 1) <> "*" Then CCaption Me.Caption & "*", Me 'state of change
    
    EventSound "Type"
    
    SStatus Len(txtBox.Text) & " characters [" & CurrentlyOpenFile & " ]", vbInformation
    
End Sub

Private Sub txtBox_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next 'shortcuts
    Select Case Shift
        Case 2 'Ctrl
            Select Case KeyCode
                Case vbKeyA
                    frmMain.titTextEditSelectAll_Click
                Case vbKeyC
                    frmMain.titTextEditCopy_Click
                Case vbKeyD
                    frmMain.titTextViewFont_Click
                Case vbKeyO
                    frmMain.titTextFileOpen_Click
                Case vbKeyP
                    frmMain.titTextEditPaste_Click
                Case vbKeyS
                    frmMain.titTextFileSave_Click
                Case vbKeyU
                    frmMain.titTextFileOpenURL_Click
                Case vbKeyX
                    frmMain.titTextEditCut_Click
            End Select
    End Select
End Sub

Private Sub txtBox_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    LoadFile Data.Files.Item(1) 'just load it
End Sub
