Attribute VB_Name = "ModSkin"
Option Explicit
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal sSectionName As String, ByVal sReturnedString As String, ByVal lSize As Long, ByVal sFileName As String) As Long
'No need for other INI readers

Function SkinForm(WhichForm As Form, Optional FromINIFile As String)
    On Error Resume Next
    Dim MyKeys As String * 60000 'This number is the limit of the length of the read string
    Dim EachElement As Variant, EachKey As Variant
    Dim CtlName As String, CtlProp As String, CtlPropVal As String
    Dim ItemIdx As Integer
        
    If Len(FromINIFile) = 0 Then FromINIFile = GetSet("SkinFile", DefaultSkinFile) 'FindPath(App.Path, "skin.ini"))
    If Dir(FromINIFile) = "" Then Exit Function 'If there's no such file, who cares?
    
    'SkinForm uses GetPrivateProfileSection
    GetPrivateProfileSection WhichForm.Name, MyKeys, 60000, FromINIFile
    EachKey = Split(MyKeys, Chr(0))
    For Each EachElement In EachKey
        If EachElement = "" Then Exit For 'not bothered
                
        'EachElement is in Label1 BackColor=255 form right now, so split with GetPrivString
        CtlName = GetPrivString(GetPrivString(EachElement, 0, "="), 0, " ") 'This gets the control name
        ItemIdx = Val(Mid$(CtlName, InStr(1, CtlName, "(") + 1, (InStrRev(CtlName, ")") - (InStr(1, CtlName, "(") + 1))))
        'ItemIdx is calculated by a crazy length of code
        CtlProp = GetPrivString(GetPrivString(EachElement, 0, "="), 1, " ") 'This gets the property, e.g. BackColor
        CtlPropVal = GetPrivString(EachElement, 1, "=") 'This gets the value of that property
        CtlPropVal = ReplaceDynamicPaths(CtlPropVal)
                    
        Select Case UCase$(CtlProp)
            Case "BACKCOLOR", "BC"
'                If UCase$(CtlProp) = "BC" Then CtlProp = "backcolor" 'restore shortened var
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).Item(ItemIdx).BackColor = Val(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).BackColor = Val(CtlPropVal)
                End If
            Case "BACKOVER", "BO"
'                If UCase$(CtlProp) = "BO" Then CtlProp = "backover" 'restore shortened var
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).Item(ItemIdx).BackOver = Val(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).BackOver = Val(CtlPropVal)
                End If
            Case "FORECOLOR", "FC"
'                If UCase$(CtlProp) = "FC" Then CtlProp = "forecolor" 'restore shortened var
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).Item(ItemIdx).ForeColor = Val(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).ForeColor = Val(CtlPropVal)
                End If
            Case "PICTURE", "PIC"
'                If UCase$(CtlProp) = "PIC" Then CtlProp = "picture" 'restore shortened var
                If InStr(1, CtlPropVal, "/") > 0 Then 'if this is not a local thing
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    CtlPropVal = DownloadFile(CtlPropVal) 'downloads the file from the net and returns path
                End If
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).Item(ItemIdx).Picture = LoadPicture(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).Picture = LoadPicture(CtlPropVal)
                End If
            Case "PICTURENORMAL", "PN"
'                If UCase$(CtlProp) = "PN" Then CtlProp = "picturenormal" 'restore shortened var
                If InStr(1, CtlPropVal, "/") > 0 Then 'if this is not a local thing
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    CtlPropVal = DownloadFile(CtlPropVal) 'downloads the file from the net and returns path
                End If
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).Item(ItemIdx).PictureNormal = LoadPicture(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).PictureNormal = LoadPicture(CtlPropVal)
                End If
            Case "PICTUREOVER", "PO"
'                If UCase$(CtlProp) = "PO" Then CtlProp = "pictureover" 'restore shortened var
                If InStr(1, CtlPropVal, "/") > 0 Then 'if this is not a local thing
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    CtlPropVal = DownloadFile(CtlPropVal) 'downloads the file from the net and returns path
                End If
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).Item(ItemIdx).PictureOver = LoadPicture(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).PictureOver = LoadPicture(CtlPropVal)
                End If
            Case "CAPTION", "CPN"
'                If UCase$(CtlProp) = "CPN" Then CtlProp = "caption" 'restore shortened var
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).Item(ItemIdx).Caption = CtlPropVal
                Else
                    WhichForm.Controls(CtlName).Caption = CtlPropVal
                End If
            Case "LEFT", "L"
'                If UCase$(CtlProp) = "L" Then CtlProp = "left" 'restore shortened var
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).Item(ItemIdx).Left = Val(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).Left = Val(CtlPropVal)
                End If
            Case "TOP", "T"
'                If UCase$(CtlProp) = "T" Then CtlProp = "top" 'restore shortened var
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).Item(ItemIdx).Top = Val(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).Top = Val(CtlPropVal)
                End If
            Case "WIDTH", "W"
'                If UCase$(CtlProp) = "W" Then CtlProp = "width" 'restore shortened var
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).Item(ItemIdx).Width = Val(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).Width = Val(CtlPropVal)
                End If
            Case "HEIGHT", "H"
'                If UCase$(CtlProp) = "H" Then CtlProp = "height" 'restore shortened var
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).Item(ItemIdx).Height = Val(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).Height = Val(CtlPropVal)
                End If
            Case "VISIBLE", "VS"
'                If UCase$(CtlProp) = "VS" Then CtlProp = "visible" 'restore shortened var
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).Item(ItemIdx).Visible = (Val(CtlPropVal) = 1)
                Else
                    WhichForm.Controls(CtlName).Visible = (Val(CtlPropVal) = 1)
                End If
        End Select
    Next
End Function

Private Function GetPrivString(Which As Variant, Optional SectionNo As Long = 0, Optional Delimiter As String = ",") As String
    'This GetPrivString is only for use in SkinForm because the source is a Variant there...
    On Error Resume Next
    Dim Arr() As String
    Arr = Split(Which, Delimiter)
    GetPrivString = Arr(SectionNo)
End Function

Public Function ReplaceDynamicPaths(FromWhat As String) As String
    On Error Resume Next
    'Converts to local paths
    FromWhat = Replace(FromWhat, "{a}", App.Path)
    FromWhat = Replace(FromWhat, "{app}", App.Path)
    FromWhat = Replace(FromWhat, "{s}", SkinPath)
    FromWhat = Replace(FromWhat, "{skin}", SkinPath)
    ReplaceDynamicPaths = FromWhat
End Function

Public Function SkinPath() As String
    On Error Resume Next
    Dim K As String
    
    K = ReadINI(UserName, "SkinFile", SettingsFile)
    If Len(K) = 0 Then K = FindPath(App.Path, "skin.ini") 'low level programming - because too many things depend on this
    
    SkinPath = Left$(K, InStrRev(K, "\") - 1)
End Function
