VERSION 5.00
Begin VB.MDIForm frmMain 
   Appearance      =   0  '¥­­±
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "ProFile"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   -345
   ClientWidth     =   11040
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  '¤â°Ê
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.PictureBox picBrw 
      Align           =   3  '¹ï»ôªí³æ¥ª¤è
      BorderStyle     =   0  '¨S¦³®Ø½u
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7080
      Left            =   0
      ScaleHeight     =   7080
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   0
      Width           =   2500
      Begin VB.ListBox lstDeleteHelper 
         Height          =   240
         ItemData        =   "frmMain.frx":00D2
         Left            =   1560
         List            =   "frmMain.frx":00D4
         TabIndex        =   16
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin ProFile.CB btnFilerefresh 
         Height          =   375
         Left            =   0
         TabIndex        =   15
         ToolTipText     =   "Refresh file list"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   "q"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMain.frx":00D6
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox TrayIcon 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   255
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Dragger 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   7350
         Index           =   1
         Left            =   2280
         MousePointer    =   9  'ªF-¦è¦V
         ScaleHeight     =   7350
         ScaleWidth      =   135
         TabIndex        =   5
         Top             =   0
         Width           =   135
         Begin VB.Image imgDrag 
            Height          =   435
            Index           =   1
            Left            =   0
            MousePointer    =   9  'ªF-¦è¦V
            Picture         =   "frmMain.frx":00F2
            Top             =   3000
            Width           =   105
         End
      End
      Begin VB.DriveListBox Drive 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   10
         Top             =   0
         Width           =   1935
      End
      Begin VB.PictureBox picFileTB 
         BorderStyle     =   0  '¨S¦³®Ø½u
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         ScaleHeight     =   285.714
         ScaleMode       =   0  '¨Ï¥ÎªÌ¦Û­q
         ScaleWidth      =   2295
         TabIndex        =   4
         Top             =   6600
         Width           =   2295
         Begin VB.TextBox txtQuickFilter 
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
            Height          =   285
            Left            =   0
            TabIndex        =   12
            Text            =   "Filter..."
            ToolTipText     =   "Enter a file filter here"
            Top             =   0
            Width           =   1335
         End
         Begin VB.CommandButton btnSelType 
            Caption         =   "&Filter..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            ToolTipText     =   "Left click to see menu, Right click to cancel"
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox Dragger 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   120
         Index           =   0
         Left            =   0
         MousePointer    =   7  '¥_-«n¦V
         ScaleHeight     =   120
         ScaleWidth      =   2295
         TabIndex        =   3
         Top             =   2640
         Width           =   2295
         Begin VB.Image imgDrag 
            Height          =   105
            Index           =   0
            Left            =   840
            MousePointer    =   7  '¥_-«n¦V
            Picture         =   "frmMain.frx":013C
            Top             =   0
            Width           =   435
         End
      End
      Begin VB.FileListBox File 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3600
         Hidden          =   -1  'True
         Left            =   0
         MultiSelect     =   2  '¶i¶¥¦h­«¿ï¨ú
         OLEDragMode     =   1  '¦Û°Ê
         OLEDropMode     =   1  '¤â°Ê
         System          =   -1  'True
         TabIndex        =   2
         Top             =   2760
         Width           =   2295
      End
      Begin VB.DirListBox Dir 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   0
         OLEDropMode     =   1  '¤â°Ê
         TabIndex        =   1
         ToolTipText     =   "Open a folder here"
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11040
      TabIndex        =   6
      Top             =   7080
      Width           =   11040
      Begin VB.PictureBox picProgress 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
         Begin VB.Image iomgRandomImage 
            Height          =   225
            Left            =   0
            Top             =   0
            Width           =   225
         End
         Begin VB.Image ImgProgressVal 
            Height          =   225
            Left            =   240
            Picture         =   "frmMain.frx":018C
            Stretch         =   -1  'True
            Top             =   0
            Width           =   375
         End
         Begin VB.Image ImgProgress 
            Height          =   225
            Left            =   0
            Picture         =   "frmMain.frx":01D2
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "Sea&rch"
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
         Left            =   9480
         TabIndex        =   9
         Top             =   0
         Width           =   855
      End
      Begin VB.TextBox txtSearch 
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
         Height          =   270
         Left            =   7560
         TabIndex        =   8
         Text            =   "Google..."
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Ready"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   30
         TabIndex        =   7
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   1425
      End
      Begin VB.Image imgStatusBack 
         Height          =   255
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin ProFile.F F1 
      Left            =   2640
      Top             =   120
      _extentx        =   979
      _extenty        =   450
   End
   Begin VB.Menu titLE 
      Caption         =   "&PF"
      Begin VB.Menu titLENew 
         Caption         =   "New"
         Visible         =   0   'False
      End
      Begin VB.Menu titLEOpen 
         Caption         =   "Open"
         Begin VB.Menu titPFOpenFile 
            Caption         =   "Open File..."
            Shortcut        =   ^O
         End
         Begin VB.Menu titPFOpenURL 
            Caption         =   "Open URL..."
            Shortcut        =   ^U
         End
         Begin VB.Menu titLEOpenPath 
            Caption         =   "Open Path..."
         End
         Begin VB.Menu titLEOpenRandomFile 
            Caption         =   "Open Random File"
            Shortcut        =   ^M
         End
         Begin VB.Menu titS06 
            Caption         =   "-"
         End
         Begin VB.Menu titLENewTextViewer 
            Caption         =   "Text Editor"
         End
         Begin VB.Menu titLEMediaPlayerStub 
            Caption         =   "Media player"
            Begin VB.Menu titLENewMediaPlayer 
               Caption         =   "Open Media Player"
               Shortcut        =   +{F1}
            End
            Begin VB.Menu titLEOpenSelFile 
               Caption         =   "Quick play selected files"
            End
            Begin VB.Menu titLEOpenMediaPlayerPlayAllShown 
               Caption         =   "Quick play shown files"
               Enabled         =   0   'False
            End
         End
         Begin VB.Menu titLENewWebBrowser 
            Caption         =   "Web Browser"
            Shortcut        =   ^T
         End
         Begin VB.Menu titLENewImageViewer 
            Caption         =   "Image Viewer"
         End
         Begin VB.Menu titS17 
            Caption         =   "-"
         End
         Begin VB.Menu titBrowserFavorites 
            Caption         =   "Favorites"
            Shortcut        =   ^{F2}
         End
      End
      Begin VB.Menu titLERecentFiles 
         Caption         =   "Recent Files"
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   9
         End
         Begin VB.Menu titS12 
            Caption         =   "-"
         End
         Begin VB.Menu titLERecentFilesClear 
            Caption         =   "Clear"
         End
      End
      Begin VB.Menu titLERecentFolders 
         Caption         =   "Recent Folders"
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   9
         End
         Begin VB.Menu titS13 
            Caption         =   "-"
         End
         Begin VB.Menu titLERecentFoldersFavs 
            Caption         =   "Internet Favorites"
         End
         Begin VB.Menu titLERecentFoldersClear 
            Caption         =   "Clear"
         End
      End
      Begin VB.Menu titS01 
         Caption         =   "-"
      End
      Begin VB.Menu titClose 
         Caption         =   "Close"
         Begin VB.Menu titCloseThisWindow 
            Caption         =   "Current Window"
            Shortcut        =   ^W
         End
         Begin VB.Menu titCloseAllWindows 
            Caption         =   "All Windows"
         End
      End
      Begin VB.Menu titLEExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu FilterMenuPopup 
      Caption         =   "FilterMenuPopup"
      Begin VB.Menu titTrayE 
         Caption         =   "TrayIcon"
         Visible         =   0   'False
         Begin VB.Menu titTrayERes 
            Caption         =   "Restore"
         End
         Begin VB.Menu titTrayEEx 
            Caption         =   "Exit"
         End
      End
      Begin VB.Menu FilterMenuPopupAll 
         Caption         =   "Show all"
      End
      Begin VB.Menu titS11 
         Caption         =   "-"
      End
      Begin VB.Menu FilterMenuPopupMusic 
         Caption         =   "Music only"
      End
      Begin VB.Menu FilterMenuPopupVideo 
         Caption         =   "Video only"
      End
      Begin VB.Menu FilterMenuPopupImage 
         Caption         =   "Images only"
      End
      Begin VB.Menu titS10 
         Caption         =   "-"
      End
      Begin VB.Menu FilterMenuPopupCustom 
         Caption         =   "Custom"
         Begin VB.Menu FilterMenuPopupCustomArchived 
            Caption         =   "Archived"
            Checked         =   -1  'True
         End
         Begin VB.Menu FilterMenuPopupCustomHidden 
            Caption         =   "Hidden"
            Checked         =   -1  'True
         End
         Begin VB.Menu FilterMenuPopupCustomNormal 
            Caption         =   "Normal"
            Checked         =   -1  'True
         End
         Begin VB.Menu FilterMenuPopupCustomReadOnly 
            Caption         =   "Read Only"
            Checked         =   -1  'True
         End
         Begin VB.Menu FilterMenuPopupCustomSystem 
            Caption         =   "System"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu titFile 
      Caption         =   "&File"
      Begin VB.Menu titFileOpenProc 
         Caption         =   "Open"
         Begin VB.Menu titFileOpen 
            Caption         =   "Open"
         End
         Begin VB.Menu titFileOpenAs 
            Caption         =   "Open as..."
         End
         Begin VB.Menu titFileOpenOpenThisFolder 
            Caption         =   "Open this folder"
         End
         Begin VB.Menu titFileOpenThisFolder 
            Caption         =   "Explore"
         End
         Begin VB.Menu titS22 
            Caption         =   "-"
         End
         Begin VB.Menu titFileToolsQuickPlay 
            Caption         =   "Quick play as media"
         End
      End
      Begin VB.Menu titFileSave 
         Caption         =   "Save"
         Begin VB.Menu titFileToolsExportFileList 
            Caption         =   "File list..."
         End
         Begin VB.Menu titS16 
            Caption         =   "-"
         End
         Begin VB.Menu titFileToolsExportm3u 
            Caption         =   "M3U..."
         End
         Begin VB.Menu titFileToolsExportwpl 
            Caption         =   "WPL..."
         End
      End
      Begin VB.Menu titFileSelect 
         Caption         =   "Select"
         Begin VB.Menu titFileSelectAll 
            Caption         =   "Select All"
         End
         Begin VB.Menu titFileSelectInvert 
            Caption         =   "Invert selection"
         End
      End
      Begin VB.Menu titFileShell 
         Caption         =   "Shell"
         Begin VB.Menu titFileShellOpen 
            Caption         =   "Open"
         End
         Begin VB.Menu titFileShellEdit 
            Caption         =   "Edit"
         End
         Begin VB.Menu titS23 
            Caption         =   "-"
         End
         Begin VB.Menu titFileCopyTo 
            Caption         =   "Copy..."
         End
         Begin VB.Menu titFileMoveTo 
            Caption         =   "Move..."
         End
      End
      Begin VB.Menu titFileRenameProc 
         Caption         =   "Rename"
         Begin VB.Menu titFileRename 
            Caption         =   "Rename..."
         End
         Begin VB.Menu titFileToolsAllRename 
            Caption         =   "Batch rename..."
         End
      End
      Begin VB.Menu titFileDeleteProc 
         Caption         =   "Delete"
         Begin VB.Menu titFileDelete 
            Caption         =   "Selected"
         End
         Begin VB.Menu titFileToolsAllDelete 
            Caption         =   "All"
         End
         Begin VB.Menu titFileMoveToRecycleBin 
            Caption         =   "Move to recycle bin"
         End
      End
      Begin VB.Menu titS27 
         Caption         =   "-"
      End
      Begin VB.Menu titFileInfo 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu titText 
      Caption         =   "&Text"
      Begin VB.Menu titTextFile 
         Caption         =   "File"
         Begin VB.Menu titTextFileOpen 
            Caption         =   "Open File..."
         End
         Begin VB.Menu titTextFileOpenURL 
            Caption         =   "Open URL..."
         End
         Begin VB.Menu titTextFileSave 
            Caption         =   "Save"
         End
         Begin VB.Menu titTextFileSaveAs 
            Caption         =   "Save as..."
         End
         Begin VB.Menu titS24 
            Caption         =   "-"
         End
         Begin VB.Menu titTextFileFavoriteThis 
            Caption         =   "Favorite this"
         End
      End
      Begin VB.Menu titTextEdit 
         Caption         =   "Edit"
         Begin VB.Menu titTextEditCut 
            Caption         =   "Cut"
         End
         Begin VB.Menu titTextEditCopy 
            Caption         =   "Copy"
         End
         Begin VB.Menu titTextEditPaste 
            Caption         =   "Paste"
         End
         Begin VB.Menu titS02 
            Caption         =   "-"
         End
         Begin VB.Menu titTextEditSelectAll 
            Caption         =   "Select All"
         End
         Begin VB.Menu titS19 
            Caption         =   "-"
         End
         Begin VB.Menu titTextEditReplace 
            Caption         =   "Replace..."
         End
         Begin VB.Menu titTextEditEncrypt 
            Caption         =   "Encrypt..."
         End
      End
      Begin VB.Menu titTextInsert 
         Caption         =   "Insert"
         Begin VB.Menu titTextInsertTimeStamp 
            Caption         =   "Time Stamp"
         End
      End
      Begin VB.Menu titTextView 
         Caption         =   "View"
         Begin VB.Menu titTextViewFont 
            Caption         =   "Font..."
         End
         Begin VB.Menu titTextEditSelText 
            Caption         =   "Selected text..."
            Begin VB.Menu titTextViewSelTextOpen 
               Caption         =   "Open"
            End
            Begin VB.Menu titTextViewSelTextOpenAsWeb 
               Caption         =   "Open as web page"
               Shortcut        =   ^{F3}
            End
            Begin VB.Menu titTextViewSelTextOpenAsImage 
               Caption         =   "Open as image"
            End
            Begin VB.Menu titTextViewSelTextOpenAsMedia 
               Caption         =   "Open as Media"
            End
         End
         Begin VB.Menu titS05 
            Caption         =   "-"
         End
         Begin VB.Menu titTextViewRunThisCode 
            Caption         =   "Run this code"
            Shortcut        =   {F5}
         End
         Begin VB.Menu titTextViewRunSelection 
            Caption         =   "Run Selection"
         End
         Begin VB.Menu titTextViewOpenScriptingEngine 
            Caption         =   "Open Scripting Engine"
         End
      End
   End
   Begin VB.Menu titMedia 
      Caption         =   "&Media"
      Begin VB.Menu titMediaFile 
         Caption         =   "File"
         Begin VB.Menu titMediaFileOpen 
            Caption         =   "Open File..."
         End
         Begin VB.Menu titMediaFileOpenURL 
            Caption         =   "Open from URL..."
         End
         Begin VB.Menu titMediaFileOpenClipboardFileName 
            Caption         =   "Open clipboard file name"
         End
         Begin VB.Menu titS25 
            Caption         =   "-"
         End
         Begin VB.Menu titMediaFileFavoriteThis 
            Caption         =   "Favorite this"
         End
      End
      Begin VB.Menu titMediaSettings 
         Caption         =   "Settings"
         Begin VB.Menu titMediaSettingsSpeed 
            Caption         =   "Speed..."
         End
         Begin VB.Menu titMediaControls 
            Caption         =   "Controls"
            Checked         =   -1  'True
         End
         Begin VB.Menu titMediaStretchVideo 
            Caption         =   "Stretch Video"
            Checked         =   -1  'True
         End
         Begin VB.Menu titMediaSyncPSM 
            Caption         =   "Sync with MSN Messenger"
            Checked         =   -1  'True
         End
         Begin VB.Menu titMediaPlaySongOnStartup 
            Caption         =   "Play song on start up"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu titMediaView 
         Caption         =   "View"
         Begin VB.Menu titMediaViewSearchSong 
            Caption         =   "Search for this song (Google)"
            Index           =   0
         End
         Begin VB.Menu titMediaViewSearchSong 
            Caption         =   "Search for lyrics (Google)"
            Index           =   1
         End
      End
   End
   Begin VB.Menu titImage 
      Caption         =   "&Image"
      Begin VB.Menu titImageFile 
         Caption         =   "File"
         Begin VB.Menu titImageFileOpen 
            Caption         =   "Open File..."
         End
         Begin VB.Menu titImageFileOpenURL 
            Caption         =   "Open URL..."
         End
      End
      Begin VB.Menu titImageBorder 
         Caption         =   "Border"
         Checked         =   -1  'True
      End
      Begin VB.Menu titImageStretch 
         Caption         =   "Stretch"
         Checked         =   -1  'True
      End
      Begin VB.Menu titImageCheckerOptions 
         Caption         =   "Checker Options"
         Begin VB.Menu titImageCheckers 
            Caption         =   "Checkers on background"
            Checked         =   -1  'True
         End
         Begin VB.Menu titImageCheckerDisableOnDrag 
            Caption         =   "Disable when dragging"
         End
      End
      Begin VB.Menu titImageTools 
         Caption         =   "Tools"
         Begin VB.Menu titImageToolsPaint 
            Caption         =   "Open with Paint"
         End
         Begin VB.Menu titLEToolsSlw 
            Caption         =   "Slideshow Maker"
         End
      End
   End
   Begin VB.Menu titBrowser 
      Caption         =   "&Browser"
      Begin VB.Menu titBrowserP 
         Caption         =   "Popup"
         Visible         =   0   'False
         Begin VB.Menu titBrowserPBMThis 
            Caption         =   "Add bookmark"
         End
         Begin VB.Menu titBrowserPOpenBMs 
            Caption         =   "Open bookmarks"
         End
      End
      Begin VB.Menu titBrowserFile 
         Caption         =   "File"
         Begin VB.Menu titBrowserFileOpen 
            Caption         =   "Open File..."
         End
         Begin VB.Menu titBrowserFileOpenURL 
            Caption         =   "Open URL..."
         End
         Begin VB.Menu titBrowserFileSavePage 
            Caption         =   "Save Page..."
         End
         Begin VB.Menu titS08 
            Caption         =   "-"
         End
         Begin VB.Menu titBrowserFilePrint 
            Caption         =   "Print Page"
         End
         Begin VB.Menu titBrowserFilePrintPreview 
            Caption         =   "Print Preview"
         End
         Begin VB.Menu titBrowserFilePrintSetup 
            Caption         =   "Print Setup..."
         End
         Begin VB.Menu titS07 
            Caption         =   "-"
         End
         Begin VB.Menu titBrowserFileProperties 
            Caption         =   "Properties..."
         End
      End
      Begin VB.Menu titBrowserView 
         Caption         =   "View"
         Begin VB.Menu titBrowserViewHistory 
            Caption         =   "View History"
         End
         Begin VB.Menu titBrowserSource 
            Caption         =   "View Source"
         End
         Begin VB.Menu titBrowserZoom 
            Caption         =   "Zoom"
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "10%"
               Index           =   0
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "20%"
               Index           =   1
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "30%"
               Index           =   2
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "40%"
               Index           =   3
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "50%"
               Index           =   4
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "60%"
               Index           =   5
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "70%"
               Index           =   6
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "80%"
               Index           =   7
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "90%"
               Index           =   8
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "100%"
               Index           =   9
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "110%"
               Index           =   10
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "120%"
               Index           =   11
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "130%"
               Index           =   12
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "140%"
               Index           =   13
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "150%"
               Index           =   14
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "160%"
               Index           =   15
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "170%"
               Index           =   16
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "180%"
               Index           =   17
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "190%"
               Index           =   18
            End
            Begin VB.Menu titBrowserZoomArray 
               Caption         =   "200%"
               Index           =   19
            End
            Begin VB.Menu titS04 
               Caption         =   "-"
            End
            Begin VB.Menu titBrowserZoomFull 
               Caption         =   "Full Screen"
            End
         End
      End
      Begin VB.Menu titBrowserEdit 
         Caption         =   "Edit"
         Begin VB.Menu titBrowserEditOpenEditor 
            Caption         =   "Open editor"
         End
         Begin VB.Menu titBrowserEditEditMode 
            Caption         =   "Edit mode"
         End
         Begin VB.Menu titBrowserEditDragDrop 
            Caption         =   "Toggle Drag and Drop"
         End
      End
      Begin VB.Menu titBrowserSilent 
         Caption         =   "Silent"
         Checked         =   -1  'True
      End
      Begin VB.Menu titBrowserAllowNewWindow 
         Caption         =   "Allow New Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu titBrowserRestoreSingleSession 
         Caption         =   "Restore Single Session"
         Checked         =   -1  'True
      End
      Begin VB.Menu titS09 
         Caption         =   "-"
      End
      Begin VB.Menu titBrowserInternetOptions 
         Caption         =   "Internet Options..."
      End
   End
   Begin VB.Menu titLEView 
      Caption         =   "&View"
      Begin VB.Menu titLEViewUseSkin 
         Caption         =   "Use skin..."
      End
      Begin VB.Menu titLEViewOpacity 
         Caption         =   "Opacity"
         Begin VB.Menu titLEViewOpacityArray 
            Caption         =   "10%"
            Index           =   0
         End
         Begin VB.Menu titLEViewOpacityArray 
            Caption         =   "20%"
            Index           =   1
         End
         Begin VB.Menu titLEViewOpacityArray 
            Caption         =   "30%"
            Index           =   2
         End
         Begin VB.Menu titLEViewOpacityArray 
            Caption         =   "40%"
            Index           =   3
         End
         Begin VB.Menu titLEViewOpacityArray 
            Caption         =   "50%"
            Index           =   4
         End
         Begin VB.Menu titLEViewOpacityArray 
            Caption         =   "60%"
            Index           =   5
         End
         Begin VB.Menu titLEViewOpacityArray 
            Caption         =   "70%"
            Index           =   6
         End
         Begin VB.Menu titLEViewOpacityArray 
            Caption         =   "80%"
            Index           =   7
         End
         Begin VB.Menu titLEViewOpacityArray 
            Caption         =   "90%"
            Index           =   8
         End
         Begin VB.Menu titLEViewOpacityArray 
            Caption         =   "100%"
            Index           =   9
         End
      End
      Begin VB.Menu titLEViewRefreshFileList 
         Caption         =   "Refresh file list"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu titLEStatusBar 
         Caption         =   "Status Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu titLEViewTogSidebar 
         Caption         =   "Toggle Sidebar"
         Shortcut        =   ^Q
      End
      Begin VB.Menu titLEViewAlwaysOnTop 
         Caption         =   "Always on top"
         Checked         =   -1  'True
      End
      Begin VB.Menu titLEViewMinToTray 
         Caption         =   "Minimize to tray"
      End
   End
   Begin VB.Menu titLETools 
      Caption         =   "&Tools"
      Begin VB.Menu titLEToolsExportfilelist 
         Caption         =   "Export file list"
      End
      Begin VB.Menu titS21 
         Caption         =   "-"
      End
      Begin VB.Menu titLEToolsCalc 
         Caption         =   "EQ Calculator"
      End
      Begin VB.Menu titLEToolsFTP 
         Caption         =   "FTP Text Client"
      End
      Begin VB.Menu titLEToolsRSS 
         Caption         =   "RSS Reader"
      End
      Begin VB.Menu titS03 
         Caption         =   "-"
      End
      Begin VB.Menu titLESandbox 
         Caption         =   "Sandbox Mode"
         Checked         =   -1  'True
      End
      Begin VB.Menu titLEOptions 
         Caption         =   "Options..."
         Shortcut        =   ^R
      End
      Begin VB.Menu titViewOptimize 
         Caption         =   "Optimize ProFile..."
      End
   End
   Begin VB.Menu titWindows 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
      Begin VB.Menu titWindowsControl 
         Caption         =   "Window Toggle"
         Shortcut        =   {F1}
      End
      Begin VB.Menu titWindowsMaxAll 
         Caption         =   "Maximize"
      End
      Begin VB.Menu titWindowsMin 
         Caption         =   "Restore"
      End
      Begin VB.Menu titWindowsTile 
         Caption         =   "Tile"
         Index           =   0
      End
      Begin VB.Menu titWindowsTile 
         Caption         =   "Tile Horizontally"
         Index           =   1
      End
      Begin VB.Menu titWindowsTile 
         Caption         =   "Tile Vertically"
         Index           =   2
      End
   End
   Begin VB.Menu titLEHelp 
      Caption         =   "&Help"
      Begin VB.Menu titHelpSet 
         Caption         =   "Where do I find the settings?"
      End
      Begin VB.Menu titS28 
         Caption         =   "-"
      End
      Begin VB.Menu titLENewNoob 
         Caption         =   "I am really, really noob"
      End
      Begin VB.Menu titLEHelpProdInfo 
         Caption         =   "Product Information..."
      End
      Begin VB.Menu titLEAbout 
         Caption         =   "About ProFile..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldX As Long, OldY As Long
Dim MyPreviousWindowState As Integer
Public COF As String 'THe variable that tells you the active form's file name

Const HarryIsDumb As Boolean = True
Const BarWidth As Long = 2500


Private Sub btnFilerefresh_Click()
    On Error Resume Next
    titLEViewRefreshFileList_Click
End Sub

Private Sub btnSearch_Click()
    On Error Resume Next
    Dim A As String
    A = txtSearch.Text
    If Left$(A, 1) = "/" Then 'commands
        If GetSet("SearchCommand", "1") = "1" Then  'only if user permits so
            Call CMD6(Mid$(A, 2))
        End If
        If GetSet("MDIForm_DeleteCMD", "1") = "1" Then txtSearch.Text = ""
    Else
        If A = GetSet("Search_Provider_Name", DefaultSearchAgent) & "..." Then Exit Sub
        A = Replace(GetSet("Search_Provider_URL", DefaultSearchURL), "%s", A)
        If Len(A) = 0 Then Exit Sub
        Dim B As New frmBRW
        B.BRW.Navigate A
        B.Show
    End If
End Sub

Private Sub btnSearch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then
        txtSearch.Text = ""
        txtSearch_LostFocus
        btnSearch.SetFocus
    End If
End Sub

Private Sub btnSelType_Click()
    On Error Resume Next
    PopupMenu FilterMenuPopup, , btnSelType.Left + btnSelType.Width, picFileTB.Top + btnSelType.Top
End Sub

Private Sub btnSelType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then
        txtQuickFilter.Text = ""
        txtQuickFilter_Change
        txtQuickFilter_LostFocus
        File.SetFocus
    End If
End Sub

Private Sub Dir_Change()
    On Error Resume Next
    GoToPath Dir.Path
End Sub

Public Function GoToPath(Where As String, Optional RecordNew As Boolean = True)
    On Error Resume Next
    If Len(Where) = 0 Then Exit Function
    Drive.Drive = Left$(Where, 1) 'C or D or E, etc
    Dir.Path = Where
    File.Path = Where
    If RecordNew Then SaveSet "Recent_Path", Where
End Function

Private Sub Dir_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    File_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub Dragger_DblClick(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
        If picBrw.Width = BarWidth Then
            picBrw.Width = Dragger(1).Width
        Else
            picBrw.Width = BarWidth
        End If
        picBrw_Resize
        If Index = 1 Then SaveSet "picBrw_Width", picBrw.Width
    End If
    EventSound "TBSize" 'play sound
End Sub

Private Sub Dragger_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    OldX = X: OldY = Y
    Dragger(Index).BackColor = RGB(255, 0, 0)
End Sub

Private Sub Dragger_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim A As Long
    Select Case Button
        Case 1
            Select Case Index
                Case 0
                    A = Dragger(0).Top - OldY + Y
                    A = A - A Mod 225 + 60  'to disable integral height thingy
                    If A > Me.ScaleHeight - 300 Then A = Me.ScaleHeight - Dragger(0).Height
                    If A < Drive.Height Then A = Drive.Height
                    Dragger(0).Move 0, A
                Case 1
                    A = picBrw.Width - OldX + X
                    If A < Dragger(1).Width + 300 Then A = Dragger(1).Width   'Redraw
                    If A > Me.Width - 300 Then A = Me.Width - Dragger(1).Width
                    If A < BarWidth + 300 And A > BarWidth - 300 Then A = BarWidth
                    If A < Me.Width / 2 + 300 And A > Me.Width / 2 - 300 Then A = Me.Width / 2
                    btnFilerefresh.Visible = (A > 300)
                    picBrw.Width = A
                    X = OldX
                    A = Null
            End Select
    End Select
End Sub

Private Sub Dragger_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next 'handles refresh
    picBrw_Resize
    Select Case Index
        Case 0
            SaveSet "Dragger0_Top", Dragger(0).Top
        Case 1
            SaveSet "picBrw_Width", picBrw.Width
    End Select
    Dragger(Index).BackColor = picBrw.BackColor '&H8000000F
    'Debug.Print Dir.Height
End Sub

Private Sub Dragger_Resize(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            imgDrag(Index).Left = (Dragger(Index).Width - imgDrag(Index).Width) / 2
        Case 1
            imgDrag(Index).Top = (Dragger(Index).Height - imgDrag(Index).Height) / 2
    End Select
End Sub

Private Sub Drive_Change()
    On Error Resume Next
    Dir.Path = Me.Drive
End Sub

Private Sub File_Click()
    On Error Resume Next
    Dim I As Integer, J As Integer
    Dim K As Double
    
    For I = 0 To File.ListCount - 1 Step 1
        If File.Selected(I) = True Then
            J = J + 1
            K = K + Round(Val(FileLen(FindPath(File.Path, File.List(I)))) / 1024 / 1024, 2)
        End If
    Next
    
    SStatus J & " file(s) selected - " & K & " MB", vbInformation
    
End Sub

Private Sub File_DblClick()
    On Error Resume Next
    Dim I As Integer
    
    AddRecentFolder File.Path
    
    For I = 0 To File.ListCount - 1 Step 1
        If File.Selected(I) = True Then
            DecideOnType FindPath(File.Path, File.List(I))
        End If
        DoEvents
    Next
End Sub

Public Function DecideOnType(eFilePathPlusName As String, Optional IgnoreRMBFileExtFlag As Boolean)
    On Error Resume Next
    Dim A As String, G As String
    Dim eFileName As String
    Dim E As Long
    Debug.Print "Path+Name:" & eFilePathPlusName
    eFileName = FileNameOnly(eFilePathPlusName)
    Debug.Print "Path+Name2:" & eFilePathPlusName
    Debug.Print "Name:" & eFileName
                            'TrimFileNameLOL(eFilePathPlusName, , , IIf(InStr(1, eFilePathPlusName, "/") > 0, "/", "\")) 'auto-generates file name
    G = Right$(eFileName, Len(eFileName) - InStrRev(eFileName, "."))
    E = OpenFileDlg.AsType(G, , IgnoreRMBFileExtFlag)
    Select Case E
        Case 0 'text
            Dim B As New frmTXT
            B.LoadFile eFilePathPlusName
        Case 1 'media
            Dim C As New frmWMP
            Dim WMPOLD As Form
            For Each WMPOLD In Me
                WMPOLD.WMP.Controls.Pause
            Next
            C.LoadFile eFilePathPlusName
        Case 2 'image
            Dim D As New frmIMG
            D.LoadFile eFilePathPlusName
        Case 3 'web
            Dim F As New frmBRW
            F.LoadFile eFilePathPlusName
        Case 4 'bookmark
            F.LoadFile F.FavAddy(eFilePathPlusName)
        Case 5 'default
            Call ShellExecute(Me.hWnd, "open", eFilePathPlusName, "", File.Path, 1)
        Case 6 'the "I HAVE NO IDEA" option
            F.LoadFile "filext.com/detaillist.php?extdetail=" & G
        Case 99 'fails
            Exit Function
    End Select
    If E < 3 Then SStatus App.ProductName & " opened " & eFileName, vbInformation
End Function

Private Sub File_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Debug.Print Shift, Chr$(KeyCode)
    If Shift = 2 Then 'Ctrl
        If KeyCode = Asc("a") Then
            titFileSelectAll_Click
        End If
    End If
    
    
    If Shift = 0 And KeyCode >= Asc("a") And KeyCode <= Asc("z") Then
        With txtQuickFilter
            .SetFocus
            .Text = Chr$(KeyCode)
            .SelStart = Len(.Text)
            .SelLength = 0
        End With
    End If
End Sub

Private Sub File_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then File_DblClick 'enables ENTER opening
End Sub

Private Sub File_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim I As Long
    If Shift = 2 And UCase(Chr(KeyCode)) = "A" Then
        For I = 0 To File.ListCount - 1
            File.Selected(I) = True
        Next
    End If
End Sub

Private Sub File_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then
        File_MouseMove Button, Shift, X, Y
    End If
End Sub

Private Sub File_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim Ix As Long, I As Integer
    Dim Mx As Long, My As Long
    Dim K As Double
    
    Mx = CLng(X / Screen.TwipsPerPixelX)
    My = CLng(Y / Screen.TwipsPerPixelY)
    Ix = SendMessage(File.hWnd, LB_ITEMFROMPOINT, 0, ByVal ((My * 65536) + Mx))
    
    If Button = 0 Then
        K = Round(Val(FileLen(FindPath(File.Path, File.List(Ix)))) / 1024 / 1024, 2)
        File.ToolTipText = File.List(Ix) & " (" & K & " MB)"
    ElseIf Button = 2 Then
        If Ix < File.ListCount Then
            File.Selected(Ix) = True
            PopupMenu titFile, , Mx * Screen.TwipsPerPixelX + File.Left, My * Screen.TwipsPerPixelY + File.Top ', titFileOpen
        End If
    End If
End Sub

Private Sub File_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    'go to the folder containing the dragged file
    GoToPath PathOnly(Data.Files.Item(1))
End Sub

Private Sub FilterMenuPopupAll_Click()
    On Error Resume Next
    File.Pattern = "*.*"
End Sub

Private Sub FilterMenuPopupCustomArchived_Click()
    On Error Resume Next
    
    File.Archive = Not File.Archive
    File.Refresh
    LoadFileCheckboxes
    
End Sub

Private Sub FilterMenuPopupCustomHidden_Click()
    On Error Resume Next
    
    File.Hidden = Not File.Hidden
    File.Refresh
    LoadFileCheckboxes
    
End Sub

Private Sub FilterMenuPopupCustomNormal_Click()
    On Error Resume Next
    
    File.Normal = Not File.Normal
    File.Refresh
    LoadFileCheckboxes
    
End Sub

Private Sub FilterMenuPopupCustomReadOnly_Click()
    On Error Resume Next
    
    File.ReadOnly = Not File.ReadOnly
    File.Refresh
    LoadFileCheckboxes
    
End Sub

Private Sub FilterMenuPopupCustomSystem_Click()
    On Error Resume Next
    
    File.System = Not File.System
    File.Refresh
    LoadFileCheckboxes
    
End Sub

Private Sub FilterMenuPopupImage_Click()
    On Error Resume Next
    File.Pattern = "*.jpg;*.bmp;*.gif;*.png;*.tga;*.psd;*.mdi;*.tiff"
End Sub

Private Sub FilterMenuPopupMusic_Click()
    On Error Resume Next
    File.Pattern = "*.mp3;*.wma;*.wav;*.mp4;*.ogg;*.aac;*.flac"
End Sub

Private Sub FilterMenuPopupVideo_Click()
    On Error Resume Next
    File.Pattern = "*.mpg;*.mp2;*.mp4;*.wmv;*.mpeg;*.ogm;*.avi;*.asx;*.rm;*.rmvb"
End Sub

Private Sub imgDrag_DblClick(Index As Integer)
    Dragger_DblClick Index
End Sub

Private Sub imgDrag_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragger_MouseDown Index, Button, Shift, X, Y
End Sub

Private Sub imgDrag_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragger_MouseMove Index, Button, Shift, X, Y
End Sub

Private Sub imgDrag_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragger_MouseUp Index, Button, Shift, X, Y
End Sub

Private Sub MDIForm_Activate()
    F1.FadeIn
End Sub

Private Sub MDIForm_Deactivate()
    F1.FadeOut
End Sub

Private Sub MDIForm_Initialize()
    On Error Resume Next
    InitCommonControls
    EventSound "Start" 'play sound
    If App.PrevInstance Then
        Select Case Val(GetSet("Multiple_Instance", "1"))
            Case 0
                'do nothing
            Case 1
                If MyMsgBox("You already have another instance of " & App.ProductName & _
                            " running." & vbCrLf & vbCrLf & "Are you sure you want to open another one?", 15, vbYesNo) = vbNo Then End
            Case 2
                End
        End Select
    End If
End Sub

Private Sub MDIForm_Load()
    On Error Resume Next
    Dim IAmHiding As Boolean
    
    F1.PrepareFade
    
    If GetSet("FirstRun") = "" Then frmIdiot.Show 1
    
    If GetSet("Splash", "1") = "1" Then frmSplash.Show 1
    
    OnTop Me.hWnd, (GetSet("MDIForm_OnTop", "0") = "1")
    Me.WindowState = Val(GetSet("MDIForm_WinMode"))
    Me.Width = Val(GetSet("MDIForm_Width", Str(Me.Width)))
    Me.Height = Val(GetSet("MDIForm_Height", Str(Me.Height)))
    Dragger(0).Top = Val(GetSet("Dragger0_Top", Str(Dragger(0).Top)))
    TrayIcon.Picture = Me.Icon
    picBrw.Width = Val(GetSet("picBrw_Width", Str(picBrw.Width)))
    txtSearch.Text = GetSet("Search_Provider_Name", DefaultSearchAgent) & "..."
    txtQuickFilter.Text = DisplayFilter
    txtQuickFilter_Change
    'File.Pattern = "*" & GetSet("File_Pattern") & "*"
    GoToPath GetSet("Recent_Path", Dir.Path)
    picStatus.Visible = (GetSet("Status_Bar", "1") = "1")
    txtSearch.Visible = (GetSet("SearchBar", "1") = "1")
    btnSearch.Visible = (GetSet("SearchBar", "1") = "1")
    'Mod32BitIcon.SetIcon Me.hwnd, "AAA"

    Call LoadRecents
    Call LoadRecentFolders
    Call OptionalMenus(False) 'Yes I do this on purpose - allows design-time menu access
    Call LoadCheckBoxes


    SkinForm Me
    SkinFormEx Me

    DSA 2
    
    picBrw_Resize
    Call LoadSearchProvider
    'startup loader
    Select Case GetSet("OpenOnStart")
        Case "1"
            titLENewTextViewer_Click
        Case "2"
            titLENewImageViewer_Click
        Case "3"
            titLENewMediaPlayer_Click
            If GetSet("Media_StartPlay", "1") = "1" Then
                AF.LoadFile GetSet("Media_Last")
            End If
        Case "4"
            titLENewWebBrowser_Click
        Case "5"
            IAmHiding = True
    End Select
    
    If GetSet("Browser_RestoreSingleSession", "1") = "1" Then 'the restore page patch
        If Len(GetSet("Browser_LastURL")) > 0 Then
            titLENewWebBrowser_Click 'new window
            AF.LoadFile GetSet("Browser_LastURL")
        End If
    End If
    
    If GetSet("OpenOnIdiot", "1") = "1" Then
        frmDumbAss.Show
    End If
    
    If IAmHiding Then
        CMD6 "tray"
    Else
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
    
End Sub

Private Sub LoadCheckBoxes()
    On Error Resume Next
    titMediaControls.Checked = (GetSet("Media_Controls", "1") = "1")
    titMediaPlaySongOnStartup.Checked = (GetSet("Media_StartPlay", "0") = "1")
    titMediaStretchVideo.Checked = (GetSet("Media_Stretch", "1") = "1")
    titMediaSyncPSM.Checked = (GetSet("Sync_PSM", "1") = "1")
    titImageBorder.Checked = (GetSet("Image_Border", "1") = "1")
    titImageCheckerDisableOnDrag.Checked = (GetSet("IMG_DragChecker", "1") = "1")
    titImageCheckers.Checked = (GetSet("Image_Checkers", "1") = "1")
    titImageStretch.Checked = (Val(GetSet("Image_Stretch", "2")) >= 1)
    titBrowserAllowNewWindow.Checked = (GetSet("Browser_AllowNewWindow", "1") = "1")
    titBrowserRestoreSingleSession.Checked = (GetSet("Browser_RestoreSingleSession", "1") = "1")
    titBrowserSilent.Checked = (GetSet("Browser_Silent", "1") = "1")
    titLESandbox.Checked = (GetSet("Sandbox") = "1")
    titLEStatusBar.Checked = (GetSet("Status_Bar") = "1")
    titLEViewAlwaysOnTop.Checked = (GetSet("MDIForm_OnTop", "0") = "1")
    
    LoadFileCheckboxes
    
End Sub

Public Sub LoadFileCheckboxes()
    On Error Resume Next
    FilterMenuPopupCustomArchived.Checked = frmMain.File.Archive
    FilterMenuPopupCustomHidden.Checked = frmMain.File.Hidden
    FilterMenuPopupCustomNormal.Checked = frmMain.File.Normal
    FilterMenuPopupCustomReadOnly.Checked = frmMain.File.ReadOnly
    FilterMenuPopupCustomSystem.Checked = frmMain.File.System
End Sub

Private Sub OptionalMenus(Optional TrueOrFalse As Boolean = True)
    On Error Resume Next
    FilterMenuPopup.Visible = TrueOrFalse
    titFile.Visible = TrueOrFalse
    titText.Visible = TrueOrFalse
    titMedia.Visible = TrueOrFalse
    titImage.Visible = TrueOrFalse
    titBrowser.Visible = TrueOrFalse
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then PopupMenu titLE
End Sub

Public Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim I As Integer, B As Integer
    For I = 1 To Data.Files.Count Step 1
        AddRecentFolder PathOnly(Data.Files.Item(I)) 'hmm just to be fair?
        DecideOnType Data.Files.Item(I)
        B = B + 1
        If B > Val(GetSet("ForNextThreshold", "100")) Then Exit For
    Next
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
'   LoadLibrary "comctl32.dll"
    Dim K As VbMsgBoxResult
    If Forms.Count > 2 And GetSet("MDIForm_MDIWarning", "1") = "1" Then 'forms.count-1 is excluding the MDIForm
        K = MyMsgBox("There are more than one windows running in " & App.ProductName & ". Are you sure you want to close " & App.ProductName & "?", 14, vbYesNo)
        If K = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    If GetSet("MDIForm_DisableUnload", "0") = "1" Then
        DSA 12
        Cancel = 1
    End If
    
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    If MyPreviousWindowState <> Me.WindowState And GetSet("MDIForm_AutoCenter", "1") = "1" Then
        If Me.Width > Screen.Width Then Me.Width = Screen.Width
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 'autocorrect onscreen position
    End If
    SaveSet "MDIForm_Width", Me.Width
    SaveSet "MDIForm_Height", Me.Height
    MyPreviousWindowState = Me.WindowState
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
    EventSound "Close" 'play sound
    DeleteIcon TrayIcon
    File.Pattern = Replace(File.Pattern, "**", "*")
    File.Pattern = Replace(File.Pattern, ".*.", ".")
    SaveSet "File_Pattern", File.Pattern
    SaveSet "MDIForm_WinMode", Me.WindowState
    
    Dim A As Control
    
    For Each A In Me
        Unload A 'prevents leakage i suppose... I can be wrong...
    Next
    End
End Sub

Private Sub picBrw_Resize()
    On Error Resume Next
    Dragger(0).Width = picBrw.Width - Dragger(1).Width
    Dragger(1).Move picBrw.Width - Dragger(1).Width, 0, Dragger(1).Width, picBrw.Height
    Dir.Move 0, Drive.Height, picBrw.Width - Dragger(1).Width, Dragger(0).Top - Drive.Height
    Dim A As Long: A = Dragger(0).Top + Dragger(0).Height
    File.Move 0, A, picBrw.Width - Dragger(1).Width, picBrw.Height - A - picFileTB.Height - 30
    picFileTB.Move 0, File.Top + File.Height + 15, picBrw.Width - Dragger(1).Width
    btnFilerefresh.Height = Drive.Height
    Drive.Move btnFilerefresh.Width, Drive.Top, picFileTB.Width - Drive.Left ', btnFilerefresh.Height
    btnFilerefresh.Visible = (picBrw.Width > 300)
End Sub

Private Sub picFileTB_Resize()
    On Error Resume Next
    btnSelType.Move picFileTB.Width - btnSelType.Width, 0, btnSelType.Width, picFileTB.Height - 15
    txtQuickFilter.Move 0, txtQuickFilter.Top, picFileTB.Width - btnSelType.Width - 30
End Sub

Private Sub picStatus_Resize()
    On Error Resume Next
    btnSearch.Move picStatus.Width - btnSearch.Width, 15, btnSearch.Width, picStatus.Height - 15
    txtSearch.Move btnSearch.Left - txtSearch.Width - 15, 15, txtSearch.Width, picStatus.Height - 30
    imgStatusBack.Move 0, 0, picStatus.Width, picStatus.Height
    'lblStatus.Move 0, 15, txtSearch.Left, 225 '225 as in, to prevent showing of second line
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Sld1_Change(MyVal As Long, myMaxVal As Long)
    Debug.Print MyVal
End Sub

Private Sub titBrowserAllowNewWindow_Click()
    On Error Resume Next
    With titBrowserAllowNewWindow
        .Checked = Not .Checked
        SaveSet "Browser_AllowNewWindow", IIf(.Checked, "1", "0")
    End With
End Sub

Private Sub titBrowserEditDragDrop_Click()
    On Error Resume Next
    AF.BRW.RegisterAsDropTarget = Not AF.BRW.RegisterAsDropTarget
End Sub

Private Sub titBrowserEditEditMode_Click()
    On Error Resume Next
    If titBrowserEditEditMode.Caption = "Edit mode" Then
        AF.BRW.Document.designMode = "On"
        titBrowserEditEditMode.Caption = "Exit edit mode"
    Else
        AF.BRW.Document.designMode = "Off"
        titBrowserEditEditMode.Caption = "Edit mode"
    End If
End Sub

Private Sub titBrowserEditOpenEditor_Click()
    On Error Resume Next
    Shell GetSet("WebEditor", "notepad") & " " & AF.BRW.LocationURL
End Sub

Public Sub titBrowserFavorites_Click()
    Dim A As String
'    A = AF.OpenFavorites() 'allows oppurtunity to detect len
'    If Len(A) > 0 Then AF.LoadFile A
    frmFav2.Show 1
End Sub

Private Sub titBrowserFileOpen_Click()
    titTextFileOpen_Click
End Sub

Private Sub titBrowserFileOpenURL_Click()
    On Error Resume Next
    Dim A As String
    A = InputBox("Enter URL or Location:", , Clipboard.GetText)
    If Len(A) > 0 Then AF.LoadFile A
End Sub

Private Sub titBrowserFilePrint_Click()
    On Error Resume Next
    AF.BRW.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub titBrowserFilePrintPreview_Click()
    On Error Resume Next
    AF.BRW.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub titBrowserFilePrintSetup_Click()
    On Error Resume Next
    AF.BRW.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub titBrowserFileProperties_Click()
    On Error Resume Next
    AF.BRW.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub titBrowserFileSavePage_Click()
    On Error Resume Next
    AF.BRW.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub titBrowserInternetOptions_Click()
    On Error Resume Next
    Shell "rundll32.exe shell32.dll,Control_RunDLL Inetcpl.cpl", vbNormalFocus
End Sub

Private Sub titBrowserPBMThis_Click()
    On Error Resume Next
    Dim A As String
    A = InputBox("Enter Name of your shortcut:", , Replace(AF.BRW.LocationName, "/", "_"))
    If Len(A) > 0 Then
        WriteINI "InternetShortcut", "URL", AF.BRW.LocationURL, FindPath(GetSet("FAV_Bookmarks", FavsPath), A & ".url")
    End If
End Sub

Private Sub titBrowserPOpenBMs_Click()
    titBrowserFavorites_Click
End Sub

Private Sub titBrowserRestoreSingleSession_Click()
    On Error Resume Next
    With titBrowserRestoreSingleSession
        .Checked = Not .Checked
        SaveSet "Browser_RestoreSingleSession", IIf(.Checked, "1", "0")
    End With
End Sub

Private Sub titBrowserSilent_Click()
    On Error Resume Next
    With titBrowserSilent
        .Checked = Not .Checked
        SaveSet "Browser_Silent", IIf(.Checked, "1", "0")
        AF.BRW.Silent = .Checked
    End With
End Sub

Private Sub titBrowserSource_Click()
    On Error Resume Next
    Dim FF As Integer
    Dim K As String
    FF = FreeFile
    K = FindPath(GetTempDir, "source.tmp")
    Open K For Output As #FF
        Print #FF, AF.BRW.Document.body.innerHTML
    Close #FF
    Shell "notepad " & K, vbNormalFocus
End Sub

Private Sub titBrowserViewHistory_Click()
    On Error Resume Next
    Dim B As New frmTXT
            B.LoadFile FindPath(App.Path, App.ProductName & ".brw.log")
End Sub

Private Sub titBrowserZoomArray_Click(Index As Integer)
    On Error Resume Next
    AF.BRW.Document.body.Style.Zoom = titBrowserZoomArray(Index).Caption
End Sub

Private Sub titBrowserZoomFull_Click()
    On Error Resume Next
    DSA 10
    Shell FindPath(App.Path, "TEExt.exe ") & AF.BRW.LocationURL, vbNormalFocus
End Sub

Private Sub titCloseAllWindows_Click()
    On Error Resume Next
    Do While Not ActiveForm Is Nothing
        Unload ActiveForm
    Loop
End Sub

Private Sub titCloseThisWindow_Click()
    On Error Resume Next
    Unload AF
End Sub

Private Sub titFileCopyTo_Click()
    On Error Resume Next
    File.Tag = "Copy To..."
    JustDoIt1
    File.Refresh
End Sub

Private Function JustDoIt1() As Boolean
    On Error Resume Next
    Dim A As String
    Dim I As Integer
    A = BrowseForFolder(Me.hWnd, File.Tag)
    If Len(A) > 0 Then
        For I = 0 To File.ListCount - 1 Step 1
            If File.Selected(I) Then
                FileCopy FindPath(File.Path, File.List(I)), FindPath(A, File.List(I))
            End If
            DoEvents
        Next
        JustDoIt1 = True
    End If
End Function

Private Sub titFileDelete_Click()
    On Error Resume Next
    Dim I As Integer
    Dim K As String
    If MsgBox("Are you sure you want to delete selected files permanently?", vbYesNo + vbQuestion) = vbYes Then
    
        lstDeleteHelper.Clear 'empty the current array
    
        For I = 0 To File.ListCount - 1 Step 1
            If File.Selected(I) = True Then
                lstDeleteHelper.AddItem FindPath(File.Path, File.List(I)) 'add to stack
                DoEvents
            End If
        Next
        For I = 0 To lstDeleteHelper.ListCount - 1 Step 1
            Kill lstDeleteHelper.List(I)
            DoEvents
        Next
        
        lstDeleteHelper.Clear 'empty the current array
        
        File.Refresh
    End If
End Sub

Private Sub titFileInfo_Click()
    On Error Resume Next
    Dim A As String, B As String
    A = FindPath(File.Path, File.FileName)
    If GetAttrib(A, vbArchive) = True Then B = B & "Archived"
    If GetAttrib(A, vbCompressed) = True Then B = B & ", Compressed"
    If GetAttrib(A, vbDirectory) = True Then B = B & ", Directory"
    If GetAttrib(A, vbHidden) = True Then B = B & ", Hidden"
    If GetAttrib(A, vbNormal) = True Then B = B & ", Normal"
    If GetAttrib(A, vbReadOnly) = True Then B = B & ", Read Only"
    If GetAttrib(A, vbTemporary) = True Then B = B & ", Temporary"
    If GetAttrib(A, vbVolume) = True Then B = B & ", Volume"
    If Left$(B, 2) = ", " Then B = Mid$(B, 3)
    MyMsgBox A & vbCrLf & vbCrLf & _
                    B & vbCrLf & vbCrLf & _
                    Round(Val(FileLen(A)) / 1024 / 1024, 2) & " MB", 8, , "File Info", True
End Sub

Private Sub titFileMoveTo_Click()
    On Error Resume Next
    Dim I As Integer
    File.Tag = "Move To..."
    If JustDoIt1 = True Then
    
        lstDeleteHelper.Clear 'empty the current array
    
        For I = 0 To File.ListCount - 1 Step 1
            If File.Selected(I) = True Then
                lstDeleteHelper.AddItem FindPath(File.Path, File.List(I)) 'add to stack
                DoEvents
            End If
        Next
        For I = 0 To lstDeleteHelper.ListCount - 1 Step 1
            Kill lstDeleteHelper.List(I)
            DoEvents
        Next
        
        lstDeleteHelper.Clear 'empty the current array
        
        File.Refresh
        
    End If
End Sub

Private Sub titFileMoveToRecycleBin_Click()
    On Error Resume Next
    Dim typOperation As SHFILEOPSTRUCT
    With typOperation
            .wFunc = &H3
            .pFrom = FindPath(File.Path, File.FileName)
            .fFlags = &H40
        End With
        SHFileOperation typOperation
    File.Refresh
End Sub

Private Sub titFileOpen_Click()
    On Error Resume Next
    File_DblClick
End Sub

Private Sub titFileOpenAs_Click()
    On Error Resume Next
    AddRecentFolder File.Path
    DecideOnType FindPath(File.Path, File.FileName), True
End Sub

Private Sub titFileOpenOpenThisFolder_Click()
    On Error Resume Next
    ViewFolderDetailed File.Path
End Sub

Private Sub titFileOpenThisFolder_Click()
    On Error Resume Next
    Shell "explorer " & File.Path, vbNormalFocus
End Sub

Private Sub titFileRename_Click()
    On Error Resume Next
    Dim A As String, B As String
    With File
        B = .FileName
        A = InputBox("Enter new file name:", "Rename", B)
        If Len(A) > 0 Then
            Name FindPath(.Path, B) As FindPath(.Path, A) 'rename kewl!!1@
            .Refresh
        End If
    End With
End Sub

Private Sub titFileSelectAll_Click()
    On Error Resume Next
    Dim I As Integer
    For I = 0 To File.ListCount - 1 Step 1
        File.Selected(I) = True
    Next
End Sub

Private Sub titFileSelectInvert_Click()
    On Error Resume Next
    Dim I As Integer
    For I = 0 To File.ListCount - 1 Step 1
        File.Selected(I) = Not File.Selected(I)
    Next
End Sub

Private Sub titFileShellEdit_Click()
    On Error Resume Next
    Call ShellExecute(Me.hWnd, "edit", FindPath(File.Path, File.FileName), "", File.Path, 1)
End Sub

Private Sub titFileShellOpen_Click()
    On Error Resume Next
    Call ShellExecute(Me.hWnd, "open", FindPath(File.Path, File.FileName), "", File.Path, 1)
End Sub

Private Sub titFileToolsAllDelete_Click()
    On Error Resume Next
    Dim I As Integer
    
    If MsgBox("Are you sure you want to delete everything on the file panel?", vbQuestion + vbYesNo) = vbYes Then
        Kill FindPath(File.Path, "*.*")
    End If
End Sub

Private Sub titFileToolsAllRename_Click()
    On Error Resume Next
    Dim K As String, J As String
    Dim I As Integer
    
    K = InputBox("This tool allows you to rename selected files to a custom file name." & vbCrLf & _
                     "the symbol * is the character for the file number. You must have that in your input below." & vbCrLf & vbCrLf & _
                     "Enter desired sequence here:", , FileNameOnly(File.Path) & "(*).ProFile.dat") 'TrimFileNameLOL(File.Path) & "(*).ProFile.dat")
                     
    If Len(K) = 0 Then Exit Sub 'no patience
                     
    If InStr(1, K, "*") = 0 Then
        MsgBox "The symbol * did not appear on your input. Please try again."
        Exit Sub
    End If
    
    For I = 0 To File.ListCount - 1 Step 1
        J = Replace(K, "*", Trim$(Str(I)))
            If File.Selected(I) = True Then
                Name FindPath(File.Path, File.List(I)) As FindPath(File.Path, J)
            End If
        DoEvents
    Next
    
    File.Refresh
    
End Sub

Private Sub titFileToolsExportFileList_Click()
    On Error Resume Next
    Dim K As String, L As String
    Dim I As Long
    For I = 0 To File.ListCount - 1 Step 1
        K = K & File.List(I) & vbCrLf
    Next
    K = K & vbCrLf & "(" & I - 1 & " items)"
    L = FindPath(App.Path, "files.txt")
    TXTFileSave K, L
    If GetSet("OpenOnParse", "1") = "1" Then
        Dim B As New frmTXT
        B.LoadFile L
    Else
        DSA 16
    End If
End Sub

Private Sub titFileToolsExportm3u_Click()
    On Error Resume Next
    Dim K As String, L As String
    Dim I As Long
    K = SaveDlg("M3U playlist|*.m3u")
    If Len(K) = 0 Then Exit Sub
    If LCase$(Right$(K, 4)) <> ".m3u" Then K = K & ".m3u" 'add extension
    
    L = "#EXTM3U" & vbCrLf
    For I = 0 To File.ListCount - 1 Step 1
        L = L & "#EXTINF:0," & File.List(I) & vbCrLf & FindPath(File.Path, File.List(I)) & vbCrLf & vbCrLf
    Next
    
    TXTFileSave L, K
    If GetSet("OpenOnParse", "1") = "1" Then
        Dim B As New frmWMP
        B.LoadFile K
    Else
        DSA 16
    End If
End Sub

Private Sub titFileToolsExportwpl_Click()
    On Error Resume Next
    Dim K As String
    Dim I As Long
    K = SaveDlg("WPL playlist|*.wpl")
    If Len(K) = 0 Then Exit Sub
    If LCase$(Right$(K, 4)) <> ".wpl" Then K = K & ".wpl" 'add extension
    
    SaveWPL K
    
    If GetSet("OpenOnParse", "1") = "1" Then
        Dim B As New frmWMP
        B.LoadFile K
    Else
        DSA 16
    End If
End Sub

Private Sub SaveWPL(WhereFile As String)
    Dim L As String
    Dim I As Integer
    L = "<?wpl version=""1.0""?>" & vbCrLf & "<smil>" & vbCrLf & "<head>" & vbCrLf & _
            "<meta name=""Generator"" content=""Microsoft Windows Media Player -- 10.0.0.3802""/>" & vbCrLf & _
            "<title>ProFile playlist</title>" & vbCrLf & "</head>" & vbCrLf & "<body>" & vbCrLf & "<seq>" & vbCrLf
            
    For I = 0 To File.ListCount - 1 Step 1
        If File.Selected(I) Then 'wow! big difference here...
            L = L & "<media src=""" & FindPath(File.Path, File.List(I)) & """/>" & vbCrLf
        End If
    Next
    
    L = L & "</seq>" & vbCrLf & "</body>" & vbCrLf & "</smil>"
    
    TXTFileSave L, WhereFile
End Sub

Private Sub titFileToolsQuickPlay_Click()
    On Error Resume Next
    Dim K As String
    K = FindPath(GetTempDir, "ProfileQuickPlay.wpl")
    SaveWPL K
    Dim B As New frmWMP
    B.LoadFile K
End Sub

Private Sub titHelpSet_Click()
    On Error Resume Next
    SStatus "Here.", vbExclamation
    titLEOptions_Click
End Sub

Private Sub titImageBorder_Click()
    With titImageBorder
        .Checked = Not .Checked
        SaveSet "Image_Border", IIf(.Checked, "1", "0")
        AF.IMG.BorderStyle = IIf(.Checked, 1, 0)
    End With
End Sub

Private Sub titImageCheckerDisableOnDrag_Click()
    On Error Resume Next
    With titImageCheckerDisableOnDrag
        .Checked = Not .Checked
        SaveSet "IMG_DragChecker", IIf(.Checked, "1", "0")
    End With
End Sub

Private Sub titImageCheckers_Click()
    On Error Resume Next
    With titImageCheckers
        .Checked = Not .Checked
        SaveSet "Image_Checkers", IIf(.Checked, "1", "0")
        AF.imgBG.Visible = .Checked
        AF.Form_Resize
    End With
End Sub

Private Sub titImageFileOpen_Click()
    titTextFileOpen_Click
End Sub

Private Sub titImageFileOpenURL_Click()
    On Error Resume Next
    Dim A As String
    DSA 4
    A = InputBox("Enter URL or Location:", , Clipboard.GetText)
    If Len(A) > 0 Then AF.LoadFile DownloadFile(A)
End Sub

Private Sub titImageStretch_Click()
    On Error Resume Next
    Dim K As Integer
    With titImageStretch
        K = IIf(.Checked, "2", "0")
        .Checked = Not .Checked
        SaveSet "Image_Stretch", Str(K)
        AF.DoStretch K 'launch 1 if you want...
        AF.Form_Resize
    End With
End Sub

Private Sub titImageToolsPaint_Click()
    On Error Resume Next
    Shell "mspaint """ & AF.CurrentlyOpenFile & """", vbNormalFocus
End Sub

Private Sub titLEAbout_Click()
    On Error Resume Next
    CMD6 "about"
End Sub

Private Sub titLEExit_Click()
    On Error Resume Next
    Unload Me
    'End
End Sub

Private Sub titLEHelpProdInfo_Click()
    On Error Resume Next
    Dim F As New frmBRW
    F.LoadFile "http://thinc.no-ip.info/projs/profile/docs.htm"
End Sub

Private Sub titLENewNoob_Click()
    On Error Resume Next
    frmDumbAss.Show
End Sub

Private Sub titLEOpenMediaPlayerPlayAllShown_Click()
    On Error Resume Next
    titFileToolsQuickPlay_Click
End Sub

Private Sub titLEOpenPath_Click()
    On Error Resume Next
    Dim K As String
    K = InputBox("Enter path:", , File.Path)
    GoToPath K
End Sub

Private Sub titLEOpenRandomFile_Click()
    On Error Resume Next
    Dim I As Integer
    Randomize
    
    For I = 0 To File.ListCount - 1 Step 1
        File.Selected(I) = False 'as to clear the stuff so only one file opens
    Next
    
    File.Selected(Round(Rnd * (File.ListCount - 1), 0)) = True
    File_DblClick
End Sub

Private Sub titLEOpenSelFile_Click()
    On Error Resume Next
    titFileToolsQuickPlay_Click
End Sub

Private Sub titLERecentFilesArray_Click(Index As Integer)
    On Error Resume Next
    DecideOnType titLERecentFilesArray(Index).Tag
End Sub

Private Sub titLENewImageViewer_Click()
    On Error Resume Next
    Dim B As New frmIMG
    B.Show
End Sub

Private Sub titLENewMediaPlayer_Click()
    On Error Resume Next
    Dim B As New frmWMP
    B.Show
End Sub

Private Sub titLENewTextViewer_Click()
    On Error Resume Next
    Dim B As New frmTXT
    B.Show
End Sub

Private Sub titLENewWebBrowser_Click()
    On Error Resume Next
    Dim B As New frmBRW
    B.Show
End Sub

Private Sub titLEOptions_Click()
    On Error Resume Next
    frmPrefs.GoToTab 0
    frmPrefs.Show 1
End Sub

Private Sub titLERecentFilesClear_Click()
    On Error Resume Next
    Dim I As Integer
    If MsgBox("Are you sure you want to clear the recent files list?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    For I = 0 To 9 Step 1
        SaveSet "Recent" & I, ""
    Next
    LoadRecents
End Sub

Private Sub titLERecentFoldersArray_Click(Index As Integer)
    On Error Resume Next
    GoToPath titLERecentFoldersArray(Index).Tag
End Sub

Private Sub titLERecentFoldersClear_Click()
    On Error Resume Next
    Dim I As Integer
    If MsgBox("Are you sure you want to clear the recent folders list?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    For I = 0 To 9 Step 1
        SaveSet "RecentF" & I, ""
    Next
    LoadRecentFolders
    MsgBox "Please restart " & App.ProductName & " to take effect.", vbInformation
End Sub

Private Sub titLERecentFoldersFavs_Click()
    On Error Resume Next
    GoToPath FavsPath
End Sub

Private Sub titLESandbox_Click()
    On Error Resume Next
    DSA 5
    With titLESandbox
        .Checked = Not .Checked
        WriteINI UserName, "Sandbox", IIf(.Checked, "1", "0"), SettingsFile
    End With
End Sub

Private Sub titLEStatusBar_Click()
    On Error Resume Next
    With titLEStatusBar
        .Checked = Not .Checked
        SaveSet "Status_Bar", IIf(.Checked, "1", "0")
        picStatus.Visible = .Checked
    End With
End Sub

Private Sub titLEToolsCalc_Click()
    On Error Resume Next
    DSA 10
    Shell FindPath(App.Path, "TEExt.exe tetracal"), vbNormalFocus
End Sub

Private Sub titLEToolsExportfilelist_Click()
    On Error Resume Next
    titFileToolsExportFileList_Click
End Sub

Private Sub titLEToolsFTP_Click()
    On Error Resume Next
    DSA 10
    Shell FindPath(App.Path, "TEExt.exe TetraFTP"), vbNormalFocus
End Sub

Private Sub titLEToolsRSS_Click()
    On Error Resume Next
    DSA 10
    Shell FindPath(App.Path, "TEExt.exe TetraRSS"), vbNormalFocus
End Sub

Private Sub titLEToolsSlw_Click()
    On Error Resume Next
    DSA 10
    Shell FindPath(App.Path, "TEExt.exe tetraslw ") & File.Path, vbNormalFocus
End Sub

Private Sub titLEViewAlwaysOnTop_Click()
    On Error Resume Next
    With titLEViewAlwaysOnTop
        .Checked = Not .Checked
        SaveSet "MDIForm_OnTop", IIf(.Checked, "1", "0")
        OnTop Me.hWnd, .Checked
    End With
End Sub

Private Sub titLEViewMinToTray_Click()
    On Error Resume Next
    CMD6 "tray"
End Sub

Private Sub titLEViewOpacityArray_Click(Index As Integer)
    On Error Resume Next
    CMD6 "trans " & Index + 1
End Sub

Private Sub titLEViewRefreshFileList_Click()
    On Error Resume Next
    Dim A As String, B As String
    A = Left$(App.Path, 3) 'like, C:\ or D:\
    B = File.Path
    GoToPath A
    GoToPath B
End Sub

Private Sub titLEViewTogSidebar_Click()
    On Error Resume Next
    Dragger_DblClick 1
End Sub

Private Sub titLEViewUseSkin_Click()
    On Error Resume Next
    With cmndlg
        .filefilter = "skin file|*.ini|all files|*.*"
        OpenFile
        If Len(.FileName) = 0 Then Exit Sub
        SaveSet "SkinFile", .FileName
        MsgBox "You might have to restart " & App.ProductName & " to see the effect.", vbExclamation
    End With
End Sub

Private Sub titMediaControls_Click()
    On Error Resume Next
    With titMediaControls
        .Checked = Not .Checked
        SaveSet "Media_Controls", IIf(.Checked, "1", "0")
        Me.ActiveForm.WMP.uiMode = IIf(GetSet("Media_Controls", "1") = "1", "full", "none")
    End With
End Sub

Private Sub titMediaFileFavoriteThis_Click()
    titTextFileFavoriteThis_Click
End Sub

Private Sub titMediaFileOpen_Click()
    On Error Resume Next
    titTextFileOpen_Click
End Sub

Private Sub titMediaFileOpenClipboardFileName_Click()
    On Error Resume Next
    AF.LoadFile Clipboard.GetText
End Sub

Private Sub titMediaFileOpenURL_Click()
    On Error Resume Next
    Dim A As String
    A = InputBox("Enter URL or location here:", , AF.WMP.url)
    If Len(A) > 0 Then
        AF.LoadFile A
    End If
End Sub

Private Sub titMediaPlaySongOnStartup_Click()
    On Error Resume Next
    With titMediaPlaySongOnStartup
        .Checked = Not .Checked
        If .Checked Then
            DSA 6
            SaveSet "OpenOnStart", "3" 'so a media pops up
        End If
        SaveSet "Media_StartPlay", IIf(.Checked, "1", "0")
    End With
End Sub

Private Sub titMediaSettingsSpeed_Click()
    On Error Resume Next
    AF.picSet.Visible = True
End Sub

Private Sub titMediaStretchVideo_Click()
    On Error Resume Next
    With titMediaStretchVideo
        .Checked = Not .Checked
        SaveSet "Media_Stretch", IIf(.Checked, "1", "0")
        AF.WMP.stretchToFit = .Checked
    End With
End Sub

Private Sub titMediaSyncPSM_Click()
    On Error Resume Next
    With titMediaSyncPSM
        .Checked = Not .Checked
        If .Checked Then
            DSA 7
        End If
        SaveSet "Sync_PSM", IIf(.Checked, "1", "0")
    End With
End Sub

Private Sub titMediaViewSearchSong_Click(Index As Integer)
    On Error Resume Next
    Dim F As New frmBRW
    Dim K As String
    K = AF.WMP.currentMedia.Name
    If Index = 1 Then K = K & " lyrics"
    F.LoadFile "http://www.google.com/search?hl=en&q=" & K & "&btnG=Google+Search"
End Sub

Public Sub titPFOpenFile_Click()
    On Error Resume Next
    Dim myFiles() As String
    Dim MyPath As String
    Dim I As Long
    
    'http://vbcity.com/forums/faq.asp?fid=6&cat=Common+Dialog
    'well, partially
    
    With cmndlg
        
        .flags = &H200 + &H80000 + &H200000
        .filefilter = "All Files|*.*|" & _
                        "Text Files|*.txt;*.dat;*.ini;*.sys;*.htm;*.html;*.xml|" & _
                        "Media Files|*.wav;*.mp1;*.mp2;*.mp3;*.mpg;*.mpeg;*.m4a;*.wma;*.wmv;*.mid;*.aiff;*.dat|" & _
                        "Image Files|*.jpg;*.jpeg;*.jpe;*.gif;*.bmp"
        OpenFile
        
        If Len(.FileName) = 0 Then Exit Sub
        
        Debug.Print "FileName: " & .FileName
        
        myFiles = Split(.FileName, vbNullChar)
        
        Select Case UBound(myFiles)
            Case 0 'if only one was selected we are done
                DecideOnType myFiles(0)
            Case Is > 0 'if more than one, we need to loop through it and append the root directory
                For I = 1 To UBound(myFiles)
                    MyPath = myFiles(0) & IIf(Right(myFiles(0), 1) <> "\", "\", "") & myFiles(I)
                    DecideOnType MyPath
                Next I
        End Select
        
    End With
End Sub

Private Sub titPFOpenURL_Click()
    On Error Resume Next
    Dim K As String, H As String
    H = AF.CurrentlyOpenFile
    K = InputBox("Enter URL:" & vbCrLf & vbCrLf & "This can be a file, a web page, or file on a web site...", , H)
    If Len(K) = 0 Then Exit Sub
    DecideOnType K
End Sub

Public Sub titTextEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText AF.txtBox.SelText
End Sub

Public Sub titTextEditCut_Click()
    On Error Resume Next
        With AF.txtBox
            Clipboard.SetText .SelText
            .SelText = ""
        End With
End Sub

Private Sub titTextEditEncrypt_Click()
    On Error Resume Next
    AF.picROT.Visible = True
    AF.picROT.ZOrder 0
End Sub

Public Sub titTextEditPaste_Click()
    On Error Resume Next
    AF.txtBox.SelText = Clipboard.GetText
End Sub

Private Sub titTextEditReplace_Click()
    'On Error Resume Next
    Dim K As String, L As String
    K = InputBox("Find this string:" & vbCrLf & vbCrLf & "^p=vbCrLf" & vbCrLf & "^c=vbCr" & vbCrLf & "^l=vbLf")
    L = InputBox("Replace with this string:" & vbCrLf & vbCrLf & "^p=vbCrLf" & vbCrLf & "^c=vbCr" & vbCrLf & "^l=vbLf")
    
    If Len(K) = 0 Then Exit Sub
    If Len(L) = 0 Then Exit Sub
    
    K = GetRepd(K)
    L = GetRepd(L)
    AF.txtBox.Text = Replace(AF.txtBox.Text, K, L)
End Sub

Public Function GetRepd(Where As String) As String
    On Error Resume Next
    Where = Replace(Where, "^p", vbCrLf)
    Where = Replace(Where, "^c", vbCr)
    GetRepd = Replace(Where, "^l", vbLf)
End Function

Public Sub titTextEditSelectAll_Click()
    On Error Resume Next
    With AF.txtBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub titTextFileFavoriteThis_Click()
    On Error Resume Next
    Dim K As String, J As String
    Dim I As Long
    
    K = InputBox("Type some names here as tags (will help you find this file later):", , FileNameOnly(AF.Tag)) 'TrimFileNameLOL(AF.Tag))
    
    For I = 30 To 38 Step 1
        K = Replace(K, Chr$(I), " ")
    Next
    
    If Len(K) > 0 Then
        J = FindPath(GetSet("FAV_Text", FindPath(App.Path, "Favs\Text")), K & ".url")
        'Debug.Print J
        WriteINI "InternetShortcut", "URL", AF.Tag, J
    End If
End Sub

Public Sub titTextFileOpen_Click()
    On Error Resume Next
    Dim Response As VbMsgBoxResult
    With cmndlg
        '.flags = &H200 + &H80000 + &H200000
        .filefilter = "any file (*.*)|*.*"
        OpenFile
        If Len(.FileName) = 0 Then Exit Sub
        AF.LoadFile .FileName
    End With
End Sub

Public Sub titTextFileOpenURL_Click()
    On Error Resume Next
    titImageFileOpenURL_Click
End Sub

Public Sub titTextFileSave_Click()
    On Error Resume Next
    If Len(AF.Tag) = 0 Then
        titTextFileSaveAs_Click 'if theres no tag... save as!
        Exit Sub
    End If
    TXTFileSave AF.txtBox.Text, AF.Tag
    If Right$(AF.Caption, 1) = "*" Then AF.Caption = Left$(AF.Caption, Len(AF.Caption) - 1)
End Sub

Private Sub titTextFileSaveAs_Click()
    With cmndlg
        .filefilter = "All files (*.*)|*.*"
        .flags = 5 Or 2
        SaveFile
        If Len(.FileName) = 0 Then Exit Sub
        AF.Tag = .FileName
    End With
    titTextFileSave_Click
    AF.Caption = FileNameOnly(AF.Tag) 'TrimFileNameLOL(AF.Tag)
End Sub

Private Sub titTextInsertTimeStamp_Click()
    On Error Resume Next
    AF.txtBox.SelText = Now()
End Sub

Public Sub titTextViewFont_Click()
    On Error Resume Next
    AF.ChangeFont
End Sub

Private Sub titTextViewOpenScriptingEngine_Click()
    On Error Resume Next
    Shell FindPath(App.Path, "ESE.exe")
End Sub

Private Sub titTextViewRunSelection_Click()
    On Error Resume Next
    RunCode AF.txtBox.SelText
End Sub

Private Sub titTextViewRunThisCode_Click()
    On Error Resume Next
    RunCode AF.txtBox.Text
End Sub

Public Sub RunCode(WhatCode As String)
    On Error Resume Next
    Dim A As String
    DSA 3
    A = FindPath(GetTempDir, "ProFile.vbs")
    TXTFileSave WhatCode, A
    Shell FindPath(App.Path, "ESE.exe ") & A
End Sub

Private Sub titTextViewSelTextOpen_Click()
    On Error Resume Next
    DecideOnType AF.txtBox.SelText
End Sub

Private Sub titTextViewSelTextOpenAsImage_Click()
    On Error Resume Next
    Dim A As New frmIMG
    A.LoadFile AF.txtBox.SelText
End Sub

Private Sub titTextViewSelTextOpenAsMedia_Click()
    On Error Resume Next
    Dim A As New frmWMP
    A.LoadFile AF.txtBox.SelText
End Sub

Private Sub titTextViewSelTextOpenAsWeb_Click()
    On Error Resume Next
    Dim A As New frmBRW
    A.LoadFile AF.txtBox.SelText
End Sub

Private Sub titTrayEEx_Click()
    On Error Resume Next
    End
End Sub

Private Sub titTrayERes_Click()
    On Error Resume Next
    NoSysIcon True
End Sub

Private Sub titViewOptimize_Click()
    On Error Resume Next
    frmOptimize.Show 1
End Sub

Private Sub titWindowsControl_Click()
    On Error Resume Next
    If AF.WindowState = 2 Then
        titWindowsTile_Click 1
    Else
        AF.WindowState = 2
    End If
End Sub

Private Sub titWindowsMaxAll_Click()
    On Error Resume Next
    AF.WindowState = 2
End Sub

Private Sub titWindowsMin_Click()
    On Error Resume Next
    AF.WindowState = 0
End Sub

Private Sub titWindowsTile_Click(Index As Integer)
    On Error Resume Next
    Me.Arrange Index
    EventSound "WinTile"
End Sub

Private Sub Trayicon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Me.Visible = False Then
        Dim Msg As Long
        Msg = (X And &HFF) * &H100
        SetForegroundWindow Me.hWnd
        Select Case Msg
            Case 0 'mouse moves
            Case &HF00  'left mouse button down
                PopupMenu titTrayE
            Case &H1E00 'left mouse button up
            Case &H3C00  'right mouse button down
                PopupMenu titTrayE
            Case &H2D00 'left mouse button double click
                NoSysIcon True    'Show App on double clicking Mouse's Left Button
            Case &H4B00 'right mouse button up
            Case &H5A00 'right mouse button double click
        End Select
    End If
End Sub

Public Sub txtQuickFilter_Change()
    On Error Resume Next
    With txtQuickFilter
        If Trim$(.Text) = GetSet("Filtre_String", DefaultFilterString) & "..." Or .Text = "" Then
            File.Pattern = "*.*"
        Else
            File.Pattern = "*" & .Text & "*"
        End If
        SaveSet "File_Pattern", .Text
    End With
End Sub

Private Sub txtQuickFilter_GotFocus()
    TBFocus txtQuickFilter, True, GetSet("Filtre_String", DefaultFilterString) & "..."
End Sub

Private Sub txtQuickFilter_LostFocus()
    TBFocus txtQuickFilter, False, GetSet("Filtre_String", DefaultFilterString) & "..."
End Sub

Private Sub txtSearch_GotFocus()
    TBFocus txtSearch, True, GetSet("Search_Provider_Name", DefaultSearchAgent) & "..."
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then btnSearch_Click
End Sub

Private Sub txtSearch_LostFocus()
    TBFocus txtSearch, False, GetSet("Search_Provider_Name", DefaultSearchAgent) & "..."
End Sub

Public Function ViewFolderDetailed(Optional WhatPath As String)
    On Error Resume Next
    Dim K As New frmFile
    If Len(WhatPath) = 0 Then WhatPath = App.Path
    K.LoadPath WhatPath
End Function
