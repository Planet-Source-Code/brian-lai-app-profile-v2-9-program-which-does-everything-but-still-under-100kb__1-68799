VERSION 5.00
Begin VB.UserControl F 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   630
   ForwardFocus    =   -1  'True
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   330
   ScaleWidth      =   630
   Windowless      =   -1  'True
   Begin VB.Label lbName 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   1  '³æ½u©T©w
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   555
   End
End
Attribute VB_Name = "F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This code is based om a submission to PSC by Ed Preston
'And then <i>ULLI the guy made it a control
'Now this is a concise version of what ProFile needs

Private ParhWnd As Long
Public MyTransNow As Integer
Public Event FadeInReady()
Public Event FadeOutReady()

Public Sub PrepareFade()
    If GetSet("Fade", "1") = "0" Then Exit Sub
    MakeTransparent ParhWnd, 1
End Sub

Public Sub FadeIn()
    Dim I As Integer, D As Integer
    Dim L As Long, K As Long
    Dim O As Double

    If GetSet("Fade", "1") = "0" Then
        RaiseEvent FadeInReady
        Exit Sub
    End If

    I = Val(GetSet("Opacity", 100))
    If I > 0 And MyTransNow <> I Then
        L = GetTickCount()

        For D = MyTransNow To I Step Speed 'So its dependent on the size of the form
            'note how MyTransNow changes
            MakeTransparent ParhWnd, D
            DoEvents
            Sleep 1
        Next

        K = GetTickCount()

        O = Val(GetSet("Opacity_Speed", "1")) * ((K - L) / 1000)
        Debug.Print "O2: " & O

        If O < 1 Then O = 1
        If O > 10 Then O = 10
        SaveSet "Opacity_Speed", Str(O)
        MakeTransparent ParhWnd, I 'End up value of I
        MyTransNow = I
    Else
        MakeOpaque ParhWnd
        MyTransNow = 100
    End If
    RaiseEvent FadeInReady
End Sub

Public Sub FadeOut()
    Dim I As Integer, D As Integer
    If GetSet("Fade", "1") = "0" Then Exit Sub
    I = Val(GetSet("Opacity", 100))
        If I > TotalOut And MyTransNow <> 0 Then
            For D = I To TotalOut Step -Speed / 2 'So its dependent on the size of the form
                MakeTransparent ParhWnd, D
                DoEvents
                Sleep 1
            Next
                MakeTransparent ParhWnd, TotalOut
                MyTransNow = TotalOut
        Else
            MakeOpaque ParhWnd
            MyTransNow = 100
        End If
    RaiseEvent FadeOutReady
End Sub

Private Sub UserControl_Paint()
    lbName = Ambient.DisplayName
    UserControl_Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ParhWnd = Parent.hWnd
End Sub

Private Sub UserControl_Resize()
    Size lbName.Width, lbName.Height
End Sub

Private Function Speed() As Double
    On Error GoTo Errr
    Dim K As Double
    K = Val(GetSet("Opacity_Speed", "1"))
    'Debug.Print "Raw K: " & K
    Speed = (Parent.Width \ 1000 - 1) * K
    If Speed < 1 Then Speed = 1
    Exit Function
Errr:
    Speed = 5
End Function

Private Function TotalOut() As Integer
    Dim K As Integer
    K = Val(GetSet("Opacity_Min", "0"))
    If K <= 0 Or K >= 20 Then K = 1 'just to prevent crashing
    TotalOut = Val(GetSet("Opacity_Out", "80"))
    If TotalOut < K Then TotalOut = K
End Function
