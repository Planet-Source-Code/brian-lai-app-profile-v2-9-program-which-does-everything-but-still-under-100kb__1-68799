Attribute VB_Name = "ModINI"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal sSectionName As String, ByVal sReturnedString As String, ByVal lSize As Long, ByVal sFileName As String) As Long
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Const PublicUserName = "Public"

Function ReadINI(Section As String, KeyName As String, FileName As String) As String
On Error Resume Next
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), FileName))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
On Error Resume Next
    Dim r
    r = writeprivateprofilestring(sSection, sKeyName, sNewString, sFileName)
End Function

Function GetSet(Key As String, Optional Default As String, Optional ForUser As String, Optional ReplaceAppWithPath As Boolean = True, Optional OnlySettingsFromSetFile As Boolean) As String
    On Error Resume Next
    Dim Buffer As String, Buffer2 As String
    Dim G As Integer
    
    If Len(ForUser) = 0 Then ForUser = UserName 'if there is no username then fill in user name for me
    
    Sleep Val(ReadINI(ForUser, "SleepFor", SettingsFile)) 'delay
    
    If OnlySettingsFromSetFile = False Then
        Buffer2 = ReadINI(ForUser, "SkinSet", SettingsFile) 'get if user wants interaction
        If Buffer2 <> "0" Then 'if the user didn't disable this
            Buffer2 = ReadINI(ForUser, "SkinFile", SettingsFile) 'get where the skin file is
            If Len(Buffer2) > 0 Then 'if theres a skin
                Buffer = ReadINI("Settings", Key, Buffer2) 'try again with the public section
                If Len(Buffer) > 0 Then
                    G = 1
                    GetSet = Buffer 'if theres something then so be it
                    GoTo ExitFunction
                End If
            End If
        End If
    End If
        
    Buffer = ReadINI(ForUser, Key, SettingsFile) 'read the entry from my username section
    If Len(Buffer) > 0 Then 'if there's an entry in my user name
        G = 2
        GetSet = Buffer 'then so be it
        GoTo ExitFunction
    End If
    
    Buffer = ReadINI(PublicUserName, Key, SettingsFile) 'try again with the public section
    If Len(Buffer) > 0 Then 'if there's an entry in the public user name
        G = 3
        GetSet = Buffer 'then so be it
        GoTo ExitFunction
    End If
    
    If Len(Default) > 0 Then 'if nothing else is present but I have a defined preset...
        G = 4
        GetSet = Default 'then so be it
        GoTo ExitFunction
    End If

Exit Function

ExitFunction:
    If ReplaceAppWithPath Then GetSet = ReplaceDynamicPaths(GetSet)
'    If InStr(1, GetSet, "ini") > 0 Then Debug.Print G & ":" & GetSet
End Function

Function GetRes(WhichType As String, ControlName As String) As String
    On Error Resume Next
    GetRes = ReadINI(WhichType, ControlName, FindPath(App.Path, "theme.ini"))
End Function

Function SaveSet(Key As String, Value As String, Optional ForUser As String) As String
On Error Resume Next
    If Len(ForUser) = 0 Then ForUser = UserName
    If ReadINI(UserName, "Sandbox", SettingsFile) <> "1" Then 'this stops all ini writes via SaveSet if sandbox is on
        WriteINI ForUser, Key, Value, SettingsFile
    End If
SaveSet = Key
End Function

Public Function SettingsFile() As String
    On Error Resume Next
    SettingsFile = FindPath(App.Path, App.ProductName & ".ini")
End Function

