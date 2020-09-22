VERSION 5.00
Begin VB.Form frmPrefs 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Preferences"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrefs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.ListBox LstTab 
      Height          =   2220
      IntegralHeight  =   0   'False
      ItemData        =   "frmPrefs.frx":000C
      Left            =   4080
      List            =   "frmPrefs.frx":002E
      TabIndex        =   155
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   5415
      TabIndex        =   151
      Top             =   0
      Width           =   5415
      Begin ProFile.CB btnTab 
         Height          =   975
         Index           =   0
         Left            =   0
         TabIndex        =   152
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         BTYPE           =   8
         TX              =   "General"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16777215
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPrefs.frx":0090
         PICPOS          =   2
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnTab 
         Height          =   975
         Index           =   1
         Left            =   1080
         TabIndex        =   153
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         BTYPE           =   8
         TX              =   "Advanced"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16777215
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPrefs.frx":00AC
         PICPOS          =   2
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnTab 
         Height          =   975
         Index           =   2
         Left            =   2160
         TabIndex        =   154
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         BTYPE           =   8
         TX              =   "About"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16777215
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPrefs.frx":00C8
         PICPOS          =   2
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.CommandButton btnUnloadMe 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   6120
      Width           =   975
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4935
      Index           =   8
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5175
      TabIndex        =   51
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   27
         Left            =   0
         TabIndex        =   58
         ToolTipText     =   "Opacity_Speed,1"
         Top             =   3000
         Width           =   5175
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   26
         Left            =   0
         TabIndex        =   56
         ToolTipText     =   "Opacity,100"
         Top             =   960
         Width           =   5175
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   29
         Left            =   0
         TabIndex        =   53
         ToolTipText     =   "Opacity_Min,0"
         Top             =   2280
         Width           =   5175
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   28
         Left            =   0
         TabIndex        =   52
         ToolTipText     =   "Opacity_Out,80"
         Top             =   1560
         Width           =   5175
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "This is the advanced fading preferences page."
         Height          =   210
         Index           =   15
         Left            =   0
         TabIndex        =   81
         Top             =   0
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Speed: (1 is normal, >1 is faster)"
         Height          =   210
         Index           =   17
         Left            =   0
         TabIndex        =   59
         Top             =   2760
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "When form is active, the opacity is (%)"
         Height          =   210
         Index           =   6
         Left            =   0
         TabIndex        =   57
         Top             =   720
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Minimum opacity for forms when faded out (%):"
         Height          =   210
         Index           =   22
         Left            =   0
         TabIndex        =   55
         Top             =   2040
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "When form is inactive, the opacity is (%)"
         Height          =   210
         Index           =   21
         Left            =   0
         TabIndex        =   54
         Top             =   1320
         Width           =   4620
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Index           =   0
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5175
      TabIndex        =   71
      Top             =   1080
      Width           =   5175
      Begin VB.Frame Frame1 
         Caption         =   "Display"
         Height          =   2775
         Index           =   0
         Left            =   0
         TabIndex        =   129
         Top             =   0
         Width           =   5175
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  '¨S¦³®Ø½u
            Height          =   2415
            Index           =   0
            Left            =   120
            ScaleHeight     =   2415
            ScaleWidth      =   4935
            TabIndex        =   130
            Top             =   240
            Width           =   4935
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   31
               Left            =   3840
               TabIndex        =   156
               ToolTipText     =   "CTL_FontSize,"
               Top             =   2040
               Width           =   495
            End
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   5
               Left            =   0
               TabIndex        =   145
               ToolTipText     =   "SkinFile,"
               Top             =   720
               Width           =   4335
            End
            Begin VB.CommandButton btnBrowse 
               Caption         =   "..."
               Height          =   375
               Index           =   5
               Left            =   4440
               TabIndex        =   144
               Top             =   690
               Width           =   495
            End
            Begin VB.CommandButton btnSelectFont 
               Caption         =   "..."
               Height          =   375
               Left            =   4440
               TabIndex        =   143
               Top             =   2010
               Width           =   495
            End
            Begin VB.TextBox txtData 
               Height          =   315
               Index           =   24
               Left            =   1200
               TabIndex        =   142
               ToolTipText     =   "Font,MS Shell Dlg"
               Top             =   2040
               Width           =   2535
            End
            Begin VB.CommandButton btnGoTab 
               Caption         =   "Language..."
               Height          =   375
               Index           =   7
               Left            =   0
               TabIndex        =   141
               Top             =   1200
               Width           =   1335
            End
            Begin VB.CommandButton btnGoTab 
               Caption         =   "Settings"
               Height          =   375
               Index           =   8
               Left            =   3840
               TabIndex        =   132
               Top             =   0
               Width           =   1095
            End
            Begin VB.CheckBox chkOpt 
               Caption         =   "Fade windows"
               Height          =   255
               Index           =   27
               Left            =   0
               TabIndex        =   131
               ToolTipText     =   "Fade,1"
               Top             =   60
               Width           =   2415
            End
            Begin VB.Label lblSkin 
               AutoSize        =   -1  'True
               Caption         =   "N / A"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   0
               Left            =   1440
               TabIndex        =   148
               Top             =   1200
               Width           =   3450
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblNotification 
               AutoSize        =   -1  'True
               Caption         =   "Choose a skin."
               Height          =   210
               Index           =   8
               Left            =   0
               TabIndex        =   147
               Top             =   480
               Width           =   4620
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblNotification 
               AutoSize        =   -1  'True
               Caption         =   "Default font:"
               Height          =   210
               Index           =   13
               Left            =   0
               TabIndex        =   146
               Top             =   2085
               Width           =   1050
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Favorites"
         Height          =   735
         Index           =   2
         Left            =   0
         TabIndex        =   137
         Top             =   4200
         Width           =   5175
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  '¨S¦³®Ø½u
            Height          =   375
            Index           =   2
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   4935
            TabIndex        =   138
            Top             =   240
            Width           =   4935
            Begin VB.CommandButton btnGoTab 
               Caption         =   "Settings"
               Height          =   375
               Index           =   6
               Left            =   3840
               TabIndex        =   140
               Top             =   0
               Width           =   1095
            End
            Begin VB.CheckBox chkOpt 
               Caption         =   "Enable the favorites manager"
               Height          =   255
               Index           =   21
               Left            =   0
               TabIndex        =   139
               ToolTipText     =   "FAV_Enable,1"
               Top             =   60
               Width           =   5055
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Multimedia"
         Height          =   1215
         Index           =   1
         Left            =   0
         TabIndex        =   133
         Top             =   2880
         Width           =   5175
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  '¨S¦³®Ø½u
            Height          =   855
            Index           =   1
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   4935
            TabIndex        =   134
            Top             =   240
            Width           =   4935
            Begin VB.CommandButton btnGoTab 
               Caption         =   "Settings"
               Height          =   375
               Index           =   4
               Left            =   3840
               TabIndex        =   149
               Top             =   480
               Width           =   1095
            End
            Begin VB.CheckBox chkOpt 
               Caption         =   "Sync song name with messenger"
               Height          =   255
               Index           =   20
               Left            =   0
               TabIndex        =   150
               ToolTipText     =   "Sync_PSM,0"
               Top             =   540
               Width           =   5055
            End
            Begin VB.CheckBox chkOpt 
               Caption         =   "Sound feedback"
               Height          =   255
               Index           =   18
               Left            =   0
               TabIndex        =   136
               ToolTipText     =   "SND_Toggle,0"
               Top             =   60
               Width           =   2415
            End
            Begin VB.CommandButton btnGoTab 
               Caption         =   "Settings"
               Height          =   375
               Index           =   5
               Left            =   3840
               TabIndex        =   135
               Top             =   0
               Width           =   1095
            End
         End
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4935
      Index           =   7
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5175
      TabIndex        =   48
      Top             =   1080
      Width           =   5175
      Begin VB.CommandButton btnBrowse 
         Caption         =   "..."
         Height          =   375
         Index           =   9
         Left            =   4680
         TabIndex        =   68
         Top             =   2130
         Width           =   495
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   9
         Left            =   0
         TabIndex        =   67
         ToolTipText     =   "Lang,"
         Top             =   2160
         Width           =   4575
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Allow skin to override my sounds and other settings"
         Height          =   255
         Index           =   25
         Left            =   0
         TabIndex        =   49
         ToolTipText     =   "SkinSet,1"
         Top             =   0
         Width           =   5055
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Slower if you use one."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   49
         Left            =   0
         TabIndex        =   70
         Top             =   2520
         Width           =   3780
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Use a language (select a pack)"
         Height          =   210
         Index           =   47
         Left            =   0
         TabIndex        =   69
         Top             =   1920
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   $"frmPrefs.frx":00E4
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   11
         Left            =   360
         TabIndex        =   50
         Top             =   240
         Width           =   4620
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4935
      Index           =   6
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5175
      TabIndex        =   22
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   19
         Left            =   0
         TabIndex        =   38
         ToolTipText     =   "FAV_Bookmarks,"
         Top             =   600
         Width           =   4095
      End
      Begin VB.CommandButton btnFolderBrowse 
         Caption         =   "Browse"
         Height          =   375
         Index           =   19
         Left            =   4200
         TabIndex        =   37
         Top             =   570
         Width           =   975
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   20
         Left            =   0
         TabIndex        =   36
         ToolTipText     =   "FAV_Media,"
         Top             =   1320
         Width           =   4095
      End
      Begin VB.CommandButton btnFolderBrowse 
         Caption         =   "Browse"
         Height          =   375
         Index           =   20
         Left            =   4200
         TabIndex        =   35
         Top             =   1290
         Width           =   975
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   21
         Left            =   0
         TabIndex        =   34
         ToolTipText     =   "FAV_Pictures,"
         Top             =   2040
         Width           =   4095
      End
      Begin VB.CommandButton btnFolderBrowse 
         Caption         =   "Browse"
         Height          =   375
         Index           =   21
         Left            =   4200
         TabIndex        =   33
         Top             =   2010
         Width           =   975
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   22
         Left            =   0
         TabIndex        =   32
         ToolTipText     =   "FAV_Programs,"
         Top             =   2760
         Width           =   4095
      End
      Begin VB.CommandButton btnFolderBrowse 
         Caption         =   "Browse"
         Height          =   375
         Index           =   22
         Left            =   4200
         TabIndex        =   31
         Top             =   2730
         Width           =   975
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   23
         Left            =   0
         TabIndex        =   30
         ToolTipText     =   "FAV_Text,"
         Top             =   3480
         Width           =   4095
      End
      Begin VB.CommandButton btnFolderBrowse 
         Caption         =   "Browse"
         Height          =   375
         Index           =   23
         Left            =   4200
         TabIndex        =   29
         Top             =   3450
         Width           =   975
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Bookmarks:"
         Height          =   210
         Index           =   20
         Left            =   0
         TabIndex        =   44
         Top             =   360
         Width           =   4020
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "You can change where ProFile finds those favorites."
         Height          =   210
         Index           =   39
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Media:"
         Height          =   210
         Index           =   40
         Left            =   0
         TabIndex        =   42
         Top             =   1080
         Width           =   4020
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Pictures:"
         Height          =   210
         Index           =   41
         Left            =   0
         TabIndex        =   41
         Top             =   1800
         Width           =   4020
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Programs:"
         Height          =   210
         Index           =   42
         Left            =   0
         TabIndex        =   40
         Top             =   2520
         Width           =   4020
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Text:"
         Height          =   210
         Index           =   43
         Left            =   0
         TabIndex        =   39
         Top             =   3240
         Width           =   4020
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4935
      Index           =   5
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5175
      TabIndex        =   21
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   10
         Left            =   1320
         TabIndex        =   99
         ToolTipText     =   "SND_Start,(none)"
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   11
         Left            =   1320
         TabIndex        =   98
         ToolTipText     =   "SND_Close,(none)"
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   12
         Left            =   1320
         TabIndex        =   97
         ToolTipText     =   "SND_TBSize,(none)"
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   13
         Left            =   1320
         TabIndex        =   96
         ToolTipText     =   "SND_WinTile,(none)"
         Top             =   2085
         Width           =   3255
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   14
         Left            =   1320
         TabIndex        =   95
         ToolTipText     =   "SND_WinOpen,(none)"
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   15
         Left            =   1320
         TabIndex        =   94
         ToolTipText     =   "SND_WinClose,(none)"
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   16
         Left            =   1320
         TabIndex        =   93
         ToolTipText     =   "SND_MSGSkip,(none)"
         Top             =   3600
         Width           =   3255
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   17
         Left            =   1320
         TabIndex        =   92
         ToolTipText     =   "SND_Type,(none)"
         Top             =   4080
         Width           =   3255
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   18
         Left            =   1320
         TabIndex        =   91
         ToolTipText     =   "SND_CMD,(none)"
         Top             =   4560
         Width           =   3255
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   4680
         TabIndex        =   90
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   4680
         TabIndex        =   89
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   2
         Left            =   4680
         TabIndex        =   88
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   3
         Left            =   4680
         TabIndex        =   87
         Top             =   2085
         Width           =   495
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   4
         Left            =   4680
         TabIndex        =   86
         Top             =   2640
         Width           =   495
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   5
         Left            =   4680
         TabIndex        =   85
         Top             =   3120
         Width           =   495
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   6
         Left            =   4680
         TabIndex        =   84
         Top             =   3600
         Width           =   495
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   7
         Left            =   4680
         TabIndex        =   83
         Top             =   4080
         Width           =   495
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   8
         Left            =   4680
         TabIndex        =   82
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Choose some audio files here which ProFile will play when events take place."
         Height          =   420
         Index           =   48
         Left            =   0
         TabIndex        =   109
         Top             =   0
         Width           =   5100
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "App start"
         Height          =   210
         Index           =   27
         Left            =   0
         TabIndex        =   108
         Top             =   645
         UseMnemonic     =   0   'False
         Width           =   765
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "App close"
         Height          =   210
         Index           =   29
         Left            =   0
         TabIndex        =   107
         Top             =   1125
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Toolbar"
         Height          =   210
         Index           =   30
         Left            =   0
         TabIndex        =   106
         Top             =   1605
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Window tiling"
         Height          =   210
         Index           =   31
         Left            =   0
         TabIndex        =   105
         Top             =   2085
         UseMnemonic     =   0   'False
         Width           =   1110
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Window opens"
         Height          =   210
         Index           =   32
         Left            =   0
         TabIndex        =   104
         Top             =   2685
         UseMnemonic     =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Window closes"
         Height          =   210
         Index           =   33
         Left            =   0
         TabIndex        =   103
         Top             =   3165
         UseMnemonic     =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Decision made"
         Height          =   210
         Index           =   34
         Left            =   0
         TabIndex        =   102
         Top             =   3645
         UseMnemonic     =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Typing"
         Height          =   210
         Index           =   35
         Left            =   0
         TabIndex        =   101
         Top             =   4125
         UseMnemonic     =   0   'False
         Width           =   555
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Commands"
         Height          =   210
         Index           =   36
         Left            =   0
         TabIndex        =   100
         Top             =   4605
         UseMnemonic     =   0   'False
         Width           =   885
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4935
      Index           =   3
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5175
      TabIndex        =   11
      Top             =   1080
      Width           =   5175
      Begin VB.CheckBox chkOpt 
         Caption         =   "Also allow commands to be typed here"
         Height          =   255
         Index           =   24
         Left            =   0
         TabIndex        =   46
         ToolTipText     =   "SearchCommand,1"
         Top             =   360
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Show the search bar"
         Height          =   255
         Index           =   23
         Left            =   0
         TabIndex        =   45
         ToolTipText     =   "SearchBar,1"
         Top             =   0
         Width           =   5055
      End
      Begin VB.ListBox lstSearchURL 
         Height          =   1740
         ItemData        =   "frmPrefs.frx":0185
         Left            =   1920
         List            =   "frmPrefs.frx":01A4
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ListBox lstSearchName 
         Height          =   2370
         ItemData        =   "frmPrefs.frx":03C4
         Left            =   0
         List            =   "frmPrefs.frx":03E3
         TabIndex        =   18
         Top             =   1080
         Width           =   5175
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   2
         Left            =   0
         TabIndex        =   13
         ToolTipText     =   "Search_Provider_Name,Google"
         Top             =   3840
         Width           =   5175
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   3
         Left            =   0
         TabIndex        =   12
         ToolTipText     =   "Search_Provider_URL,http://www.google.com/search?hl=en&q=%s&btnG=Google+Search"
         Top             =   4560
         Width           =   5175
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Select a search engine you would like to use."
         Height          =   210
         Index           =   9
         Left            =   0
         TabIndex        =   16
         Top             =   840
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Search Provider Name:"
         Height          =   210
         Index           =   4
         Left            =   0
         TabIndex        =   15
         Top             =   3600
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Search Provider Search String:"
         Height          =   210
         Index           =   5
         Left            =   0
         TabIndex        =   14
         Top             =   4320
         Width           =   4620
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4935
      Index           =   4
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5175
      TabIndex        =   20
      Top             =   1080
      Width           =   5175
      Begin VB.ComboBox cboOpt 
         Height          =   330
         Index           =   2
         ItemData        =   "frmPrefs.frx":045D
         Left            =   0
         List            =   "frmPrefs.frx":046D
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   160
         ToolTipText     =   "PSM_ShowCredit,1"
         Top             =   4560
         Width           =   5175
      End
      Begin VB.CommandButton btnBrowse 
         Caption         =   "..."
         Height          =   375
         Index           =   7
         Left            =   4680
         TabIndex        =   26
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   7
         Left            =   0
         TabIndex        =   25
         ToolTipText     =   "PSMLoc,{app}\PSMChanger.exe"
         Top             =   750
         Width           =   4575
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   8
         Left            =   0
         TabIndex        =   24
         ToolTipText     =   "PSM,%n (%l)"
         Top             =   2040
         Width           =   4575
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Replace other YE company names with pow!!"
         Height          =   255
         Index           =   19
         Left            =   0
         TabIndex        =   23
         ToolTipText     =   "YE_Elim,1"
         Top             =   0
         Width           =   5055
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "How should song names look on MSN?"
         Height          =   210
         Index           =   12
         Left            =   0
         TabIndex        =   161
         Top             =   4320
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "%l  - length of the song"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   28
         Left            =   0
         TabIndex        =   66
         Top             =   2880
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "%s  - the size of the file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   46
         Left            =   0
         TabIndex        =   65
         Top             =   3360
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "%t  - the time the message is changed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   45
         Left            =   0
         TabIndex        =   64
         Top             =   3600
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "%f  - name of the file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   44
         Left            =   0
         TabIndex        =   63
         Top             =   2640
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "%n  - name of the song"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   38
         Left            =   0
         TabIndex        =   62
         Top             =   3120
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Syntax:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   25
         Left            =   0
         TabIndex        =   61
         Top             =   2400
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblURL 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "What's this?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   28
         Left            =   0
         TabIndex        =   60
         ToolTipText     =   "http://thinc.no-ip.info/projs/profile"
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Your plugin:"
         Height          =   210
         Index           =   23
         Left            =   0
         TabIndex        =   28
         Top             =   465
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Show your personal message like this:"
         Height          =   210
         Index           =   24
         Left            =   0
         TabIndex        =   27
         Top             =   1800
         Width           =   4620
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4935
      Index           =   1
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5175
      TabIndex        =   5
      Top             =   1080
      Width           =   5175
      Begin VB.ComboBox cboOpt 
         Height          =   330
         Index           =   0
         ItemData        =   "frmPrefs.frx":04CF
         Left            =   2640
         List            =   "frmPrefs.frx":04DF
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   157
         ToolTipText     =   "Multiple_Instance,0"
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Ask when closing many windows"
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   123
         ToolTipText     =   "MDIForm_MDIWarning,1"
         Top             =   960
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "ProFile cannot be closed"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   122
         ToolTipText     =   "MDIForm_DisableUnload,0"
         Top             =   600
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Log what web sites I go to"
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   121
         ToolTipText     =   "BRW_Log,1"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Are you a noob?"
         Height          =   255
         Index           =   14
         Left            =   0
         TabIndex        =   120
         ToolTipText     =   "OpenOnIdiot,1"
         Top             =   0
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Sandbox mode"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   119
         ToolTipText     =   "Sandbox,"
         Top             =   1680
         Width           =   5055
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  '¼È¤î
         Index           =   25
         Left            =   0
         PasswordChar    =   "n"
         TabIndex        =   118
         ToolTipText     =   "Password,"
         Top             =   3960
         Width           =   5175
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   117
         ToolTipText     =   "Filtre_String,Filter"
         Top             =   4560
         Width           =   5175
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "The browser can download things to this computer"
         Height          =   255
         Index           =   26
         Left            =   0
         TabIndex        =   116
         ToolTipText     =   "BRW_Download,1"
         Top             =   2160
         Width           =   5055
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   30
         Left            =   2640
         TabIndex        =   115
         ToolTipText     =   "SleepFor,0"
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Multiple instances of program:"
         Height          =   210
         Index           =   2
         Left            =   0
         TabIndex        =   158
         Top             =   2820
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Allows easy access when ProFile starts."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   26
         Left            =   240
         TabIndex        =   128
         Top             =   240
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Sandbox prevents changes in the settings."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   127
         Top             =   1920
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Password protect profile (PPP):"
         Height          =   210
         Index           =   1
         Left            =   0
         TabIndex        =   126
         Top             =   3645
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Show this when search box is inactive:"
         Height          =   210
         Index           =   3
         Left            =   0
         TabIndex        =   125
         Top             =   4320
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Read setting delay (ms):"
         Height          =   210
         Index           =   37
         Left            =   0
         TabIndex        =   124
         Top             =   3285
         Width           =   4620
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4935
      Index           =   9
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   1080
      Width           =   5175
      Begin VB.ListBox LstCredits 
         Height          =   2295
         IntegralHeight  =   0   'False
         ItemData        =   "frmPrefs.frx":051A
         Left            =   2400
         List            =   "frmPrefs.frx":0548
         TabIndex        =   6
         Top             =   2640
         Width           =   2775
      End
      Begin VB.CommandButton btnWriteXPVS 
         Caption         =   "&Make Manifest"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton btnShellINI 
         Caption         =   "&Edit INI..."
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3720
         Width           =   1935
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Make on startup"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "MakeManifest,0"
         Top             =   4680
         Width           =   5175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "List of beta testers: (thank you!)"
         Height          =   210
         Left            =   2400
         TabIndex        =   7
         Top             =   2400
         Width           =   2715
      End
      Begin VB.Image imgLogo 
         Height          =   645
         Left            =   0
         Picture         =   "frmPrefs.frx":0602
         Top             =   480
         Width           =   2190
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   5160
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label lblProdVer 
         BackStyle       =   0  '³z©ú
         Caption         =   "Version "
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
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label lblProdDes 
         BackStyle       =   0  '³z©ú
         Caption         =   "Description"
         Height          =   1095
         Left            =   0
         TabIndex        =   4
         Top             =   1200
         Width           =   5175
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
         TabIndex        =   47
         Top             =   360
         Width           =   1440
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4935
      Index           =   2
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   5175
      TabIndex        =   19
      Top             =   1080
      Width           =   5175
      Begin VB.ComboBox cboOpt 
         Height          =   330
         Index           =   1
         ItemData        =   "frmPrefs.frx":0E32
         Left            =   1920
         List            =   "frmPrefs.frx":0E4B
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   159
         ToolTipText     =   "OpenOnStart,"
         Top             =   3720
         Width           =   3255
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Use favorites toolbar"
         Height          =   255
         Index           =   17
         Left            =   0
         TabIndex        =   114
         ToolTipText     =   "BRW_AutoFavsBarSwitch,1"
         Top             =   1800
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Show full paths in menus"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   113
         ToolTipText     =   "ShowFullPaths,"
         Top             =   2520
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Open files after parsing"
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   112
         ToolTipText     =   "OpenOnParse,1"
         Top             =   1440
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Clear commands after they are run"
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   111
         ToolTipText     =   "MDIForm_DeleteCMD,1"
         Top             =   2160
         Width           =   5055
      End
      Begin VB.CommandButton btnSplash 
         Caption         =   "Show now"
         Height          =   375
         Left            =   3840
         TabIndex        =   72
         Top             =   660
         Width           =   1335
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Make text boxes look flat"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   78
         ToolTipText     =   "CTL_Flatten,0"
         Top             =   0
         Width           =   5055
      End
      Begin VB.CommandButton btnBrowse 
         Caption         =   "..."
         Height          =   375
         Index           =   6
         Left            =   4680
         TabIndex        =   77
         Top             =   4485
         Width           =   495
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   6
         Left            =   0
         TabIndex        =   76
         ToolTipText     =   "WebEditor,notepad"
         Top             =   4515
         Width           =   4575
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Correct when ProFile flies out of screen"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   75
         ToolTipText     =   "MDIForm_AutoCenter,1"
         Top             =   1080
         Width           =   4935
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Show shadows under windows"
         Height          =   255
         Index           =   15
         Left            =   0
         TabIndex        =   74
         ToolTipText     =   "DropShadow,1"
         Top             =   360
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Show splash window"
         Height          =   255
         Index           =   22
         Left            =   0
         TabIndex        =   73
         ToolTipText     =   "Splash,1"
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Args: [exe] %f"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   18
         Left            =   1320
         TabIndex        =   110
         Top             =   4200
         Width           =   3780
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Let this program edit my web pages:"
         Height          =   210
         Index           =   16
         Left            =   0
         TabIndex        =   80
         Top             =   4185
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Show on startup:"
         Height          =   210
         Index           =   7
         Left            =   0
         TabIndex        =   79
         Top             =   3780
         Width           =   4620
         WordWrap        =   -1  'True
      End
   End
   Begin ProFile.F F1 
      Left            =   120
      Top             =   6120
      _ExtentX        =   979
      _ExtentY        =   450
   End
   Begin VB.Image IMGbkg 
      Height          =   6615
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5415
   End
   Begin VB.Menu titSounds 
      Caption         =   "Sounds"
      Visible         =   0   'False
      Begin VB.Menu titSoundsPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu titS14 
         Caption         =   "-"
      End
      Begin VB.Menu titSoundsBrowseFile 
         Caption         =   "Browse for a file"
      End
      Begin VB.Menu titSoundsRemoveFile 
         Caption         =   "Do not play this file"
      End
   End
   Begin VB.Menu titFiles 
      Caption         =   "Files"
      Visible         =   0   'False
      Begin VB.Menu titFilesBrowse 
         Caption         =   "Browse..."
      End
      Begin VB.Menu titFilesGoto 
         Caption         =   "Goto path"
      End
      Begin VB.Menu titFilesClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu titPages 
      Caption         =   "Tab1Popup"
      Visible         =   0   'False
      Begin VB.Menu titPR 
         Caption         =   "Filler"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu titPR 
         Caption         =   "Other general prefs"
         Index           =   1
      End
      Begin VB.Menu titPR 
         Caption         =   "Appearance"
         Index           =   2
      End
      Begin VB.Menu titPR 
         Caption         =   "Search options"
         Index           =   3
      End
      Begin VB.Menu titPR 
         Caption         =   "Messenger options"
         Index           =   4
      End
      Begin VB.Menu titPR 
         Caption         =   "Sound prefs"
         Index           =   5
      End
      Begin VB.Menu titPR 
         Caption         =   "Favorites options"
         Index           =   6
      End
      Begin VB.Menu titPR 
         Caption         =   "Skin prefs"
         Index           =   7
      End
      Begin VB.Menu titPR 
         Caption         =   "Transparency options"
         Index           =   8
      End
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InaSub As Boolean
Dim TheFilter As String
Dim InTheTab As Integer

Private Sub btnBrowse_Click(Index As Integer)
    On Error Resume Next

    titFiles.Tag = Index
    
    If Index = 5 Then
        TheFilter = "Configuration file (*.ini)|*.ini"
        PopupMenu titFiles, , picTabSwitch(InTheTab).Left + btnBrowse(Index).Left + Picture1(0).Left + Frame1(0).Left, picTabSwitch(5).Top + btnBrowse(Index).Top + btnBrowse(Index).Height + Picture1(0).Top + Frame1(0).Top, titFilesBrowse
    Else
        TheFilter = "All files (*.*)|*.*"
        PopupMenu titFiles, , picTabSwitch(InTheTab).Left + btnBrowse(Index).Left, picTabSwitch(5).Top + btnBrowse(Index).Top + btnBrowse(Index).Height, titFilesBrowse
    End If
    
End Sub

Private Sub btnFolderBrowse_Click(Index As Integer)
    On Error Resume Next
    K = BrowseForFolder(Me.hWnd)
    If Len(K) = 0 Then Exit Sub
    
    txtData(Index).Text = K
End Sub

Private Sub btnGoTo_Click(Index As Integer)
    On Error Resume Next
    Shell "explorer " & PathOnly(txtData(Index).Text), vbNormalFocus
End Sub

Private Sub btnGoTab_Click(Index As Integer)
    On Error Resume Next
    GoToTab Index
End Sub

Private Sub btnOK_Click()
    On Error Resume Next
        
    Dim I As Integer
    For I = 0 To chkOpt.UBound Step 1 'Save Settings
        If Len(chkOpt(I).Tag) > 0 Then
            WriteINI UserName, GetString(chkOpt(I).ToolTipText), Str$(chkOpt(I).Value), SettingsFile
        End If
    Next
    For I = 0 To txtData.UBound Step 1 'Save Settings
        If Len(txtData(I).Tag) > 0 Then
            WriteINI UserName, GetString(txtData(I).ToolTipText), txtData(I).Text, SettingsFile
        End If
    Next
    For I = 0 To cboOpt.UBound Step 1 'Save Settings
        Debug.Print I
        If Len(cboOpt(I).Tag) > 0 Then
            WriteINI UserName, GetString(cboOpt(I).ToolTipText), cboOpt(I).ListIndex - 1, SettingsFile
            Debug.Print "comboBox edit"
        End If
    Next
    If Len(txtData(24).Tag) > 0 Or Len(txtData(31).Tag) > 0 Then 'if display is changed
        SkinForm frmMain
        SkinFormEx frmMain 'for the heck of it
    End If
    
    Unload Me
End Sub


Private Sub btnSelectFont_Click()
    If Len(txtData(31).Text) > 0 Then SelectFont.mFontSize = Val(txtData(31).Text)
    SelectFont.mFontName = txtData(24).Text
    ShowFont
    If Len(SelectFont.mFontName) > 0 Then txtData(24).Text = SelectFont.mFontName
    If Len(SelectFont.mFontSize) > 0 Then txtData(31).Text = SelectFont.mFontSize
End Sub

Private Sub btnShellINI_Click()
    On Error Resume Next
    Shell "notepad " & SettingsFile, vbNormalFocus
End Sub

Private Sub btnSndPlaySound_Click(Index As Integer)
    On Error Resume Next
    titFiles.Tag = Index + 10
    PopupMenu titSounds, , picTabSwitch(5).Left + btnSndPlaySound(Index).Left, picTabSwitch(5).Top + btnSndPlaySound(Index).Top + btnSndPlaySound(Index).Height, titSoundsPlay
End Sub

Private Sub btnSplash_Click()
    On Error Resume Next
    With frmSplash
        .ForceClick = True
        .T1.Enabled = False
        .Show 1
    End With
End Sub

Private Sub btnTab_Click(Index As Integer)
    Select Case Index
        Case 0
            GoToTab 0
        Case 1
            PopupMenu titPages, , btnTab(Index).Left, btnTab(Index).Height
        Case 2
            GoToTab 9
    End Select
End Sub

Private Sub btnUnloadMe_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub btnWriteXPVS_Click()
    On Error Resume Next
    If MsgBox("You do not normally require this tool, but this will enable XP styles. Continue?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    XPVB
    MsgBox "Manifest Written. Please restart " & App.ProductName & " to see effect.", vbInformation
End Sub

Private Sub cboOpt_Change(Index As Integer)
    On Error Resume Next 'for real-time stuff, no undo
    If InaSub Then Exit Sub
    cboOpt(Index).Tag = "EDITED"
End Sub

Private Sub cboOpt_Click(Index As Integer)
    cboOpt_Change Index 'stub
End Sub

Private Sub chkOpt_Click(Index As Integer)
    On Error Resume Next 'for real-time stuff, no undo
    If InaSub Then Exit Sub
    With chkOpt(Index)
        Select Case Index
            Case 3
                If .Value = 1 Then DSA 9
        End Select
    End With
    chkOpt(Index).Tag = "EDITED"
End Sub

Private Sub Form_Activate()
    'InitCommonControls
    If CheckPW = False Then Unload Me 'protection
    F1.FadeIn
End Sub

Private Sub Form_Deactivate()
    F1.FadeOut
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim K As String
    F1.PrepareFade
    InaSub = True
    
    lblProdVer.Caption = App.ProductName & " " & App.Major & "." & App.Minor & "." & App.Revision
    lblProdDes.Caption = App.ProductName & " " & MyVer & ", some rights reserved by Thinc." & vbCrLf & _
    "Made by Brian Lai" & vbCrLf & SoftwareHomePage
        
    SkinForm Me
    SkinFormEx Me

    SStatus "Loading checkboxes", vbExclamation
    For I = 0 To chkOpt.UBound Step 1 'Load Settings
        K = GetString(chkOpt(I).ToolTipText, 1)
        K = ReplaceDynamicPaths(K)
        chkOpt(I).Value = GetSet(GetString(chkOpt(I).ToolTipText), K, , , True)
        SProgress CLng(I), , chkOpt.UBound - 1
        DoEvents
    Next
    
    SStatus "Loading combo boxes", vbExclamation
    For I = 0 To cboOpt.UBound Step 1 'Load Settings
        K = GetString(cboOpt(I).ToolTipText, 1)
        K = ReplaceDynamicPaths(K)
        Dim L As String
        L = GetSet(GetString(cboOpt(I).ToolTipText), K, , , True)
        If Len(L) > 0 Then
            cboOpt(I).ListIndex = Val(L) + 1 '+1 is to make up for "nothing"=0
        Else
            cboOpt(I).ListIndex = 0
        End If
        SProgress CLng(I), , chkOpt.UBound - 1
        DoEvents
    Next
    
    SStatus "Loading textboxes", vbExclamation
    For I = 0 To txtData.UBound Step 1 'Load Settings
        K = GetString(txtData(I).ToolTipText, 1)
        K = ReplaceDynamicPaths(K)
        txtData(I).Text = GetSet(GetString(txtData(I).ToolTipText), K, , , True)
        SProgress CLng(I), , txtData.UBound - 1
        DoEvents
    Next
    
    InaSub = False
    
    lblSkin(0).Caption = "Skin Info: " & SkinInfo("Info")
    EventSound "WinOpen"
    SStatus
    
End Sub

Public Function SkinInfo(KeyName As String) As String
    On Error Resume Next
    SkinInfo = ReadINI("Skin", KeyName, GetSet("SkinFile"))
    If Len(SkinInfo) = 0 Then SkinInfo = "n/a"
End Function

'Private Sub Form_Resize()
'    On Error Resume Next
'    'This is a fix for a random WinPos bug
'    Me.Height = btnUnloadMe.Top + btnUnloadMe.Height + 120 + (Me.Height - Me.ScaleHeight)
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    EventSound "WinClose"
End Sub

Private Sub lblURL_Click(Index As Integer)
    On Error Resume Next
    Shell "explorer " & lblURL(Index).ToolTipText, vbNormalFocus
End Sub

Private Sub lstSearchName_Click()
    On Error Resume Next
    txtData(2).Text = lstSearchName.List(lstSearchName.ListIndex)
    txtData(3).Text = lstSearchURL.List(lstSearchName.ListIndex)
End Sub

Private Sub LstTab_Click()
    On Error Resume Next
    'picTabSwitch(LstTab.ListIndex).ZOrder 0
    GoToTab LstTab.ListIndex
End Sub

Public Function GoToTab(Index As Integer)
    On Error Resume Next
    Dim I As Integer
    InTheTab = Index
    
    For I = 0 To picTabSwitch.UBound
        picTabSwitch(I).Visible = False 'for the sake of making things accessible
    Next
    picTabSwitch(Index).Visible = True
End Function

Private Sub titFilesBrowse_Click()
    On Error Resume Next
    With cmndlg
        .filefilter = TheFilter 'load the one
        If Len(.filefilter) = 0 Then .filefilter = "any file (*.*)|*.*" 'if there isnt one
        OpenFile
        If Len(.FileName) = 0 Then Exit Sub
        txtData(Val(titFiles.Tag)).Text = .FileName
    End With
End Sub

Private Sub titFilesClear_Click()
    txtData(Val(titFiles.Tag)).Text = ""
End Sub

Private Sub titFilesGoto_Click()
    Shell "explorer " & PathOnly(txtData(Val(titFiles.Tag)).Text), vbNormalFocus
End Sub

Private Sub titPR_Click(Index As Integer)
    GoToTab Index
End Sub

Private Sub titSoundsBrowseFile_Click()
    TheFilter = "wave files (*.wav)|*.wav"
    titFilesBrowse_Click
End Sub

Private Sub titSoundsPlay_Click()
    On Error Resume Next
    sndPlaySound txtData(Val(titFiles.Tag)).Text, 1
End Sub

Private Sub titSoundsRemoveFile_Click()
    On Error Resume Next
    txtData(Val(titFiles.Tag)).Text = "(None)"
End Sub

Private Sub txtData_Change(Index As Integer)
    If InaSub Then Exit Sub
    If Index <= 18 And Index >= 10 Then 'this is for the sake of having no music file
        If txtData(Index).Text = "" Then txtData(Index).Text = "(None)"
    End If
    txtData(Index).Tag = "EDITED"
End Sub
