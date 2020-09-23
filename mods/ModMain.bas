Attribute VB_Name = "ModMain"
'--------------------------------------------------'
' DM Ebook Designer Beta 1                         '
' Written and designed by Ben Jones                '
' Email1 dreamvb@yahoo.com                         '
' Email2 vbdream2k@yahoo.com                       '
' Web-site http://dmeasyhttp.2ya.com               '
' Last updated 28-10-03                            '
' Freeware Open Source eBook Designer for Windows  '
'--------------------------------------------------'

' If you would like to make some chnages to this project
' Then your are free to do so I whould also like to see if some
' Someone has upadted it.

' If you like to post the code onto your website then please do.
' But please remmber were it came from. Please just don't says it all your own work.
' This Project is FREEWARE that means you may not use it to gain any sort of profit

'Thank you.
'Ben Jones

Public BookPath As String


Type MyEbook
    eBookTitle As String
    eBookAuthor As String
    eBookHomePage As String
    eBookkExeName As String
    eBookCompDate As Date
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public TEBook As MyEbook

' Default consts for the ebook
Public Const m_def_ebookTitle = "My eBook"
Public Const m_def_eBookAuthor = "Ben Jones"
Public Const m_def_eBookEXEName = "ebook.exe"

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000
Private Const GWL_EXSTYLE = (-20)
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private Const RT_STRING = 6& ' String Resource const

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
' API calls used for chnageing and updateing a programs resource files
Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As String, ByVal lpName As Integer, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long

Public Function OpenFile(lzFilename As String) As String
Dim TFile As Long
Dim StrBuffer As String

    TFile = FreeFile
    Open lzFilename For Binary As #TFile
        StrBuffer = Space(LOF(TFile))
        Get #TFile, , StrBuffer
    Close #TFile
    OpenFile = StrBuffer
    
    StrBuffer = ""
End Function

Public Function FindFile(lzFile As String) As Boolean
    ' This function will retun a result of a file of exsitence file found will return with a true value
    If Dir$(lzFile) = "" Then FindFile = False Else FindFile = True
End Function

Public Function FixPath(lzPath As String) As String
    ' Fixes a path by adding a back slash if required
    If Right$(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function GetFileExt(lzFile As String) As String
Dim I As Long, iPart As Long, StrA As String
   For I = Len(lzFile) To 1 Step -1
        StrA = Mid(lzFile, I, 1)
        If StrA = "." Then
            iPart = I
            Exit For
        End If
   Next
   
   If iPart = 0 Then
        GetFileExt = ""
    Else
        GetFileExt = UCase$(Mid$(lzFile, iPart + 1, Len(lzFile)))
   End If
   iPart = 0: I = 0
   StrA = ""
   
End Function

Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String) As String
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim OffSet As Integer

    bInf.hOwner = hWndOwner
    bInf.lpszTitle = sTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
        OffSet = InStr(RetPath, Chr$(0))
        GetFolder = Left$(RetPath, OffSet - 1)
    End If

End Function

Private Function FlatBorder(ByVal hwnd As Long, MakeControlFlat As Boolean)
Dim TFlat As Long
    TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
    If MakeControlFlat Then
        TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    Else
        TFlat = TFlat And Not WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
    End If
    SetWindowLong hwnd, GWL_EXSTYLE, TFlat
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
  
End Function

Public Function MakeFlatControls(frm As Form)
Dim Icnt As Long
    ' Returns long 32bit hangle of each control found for the flatborder function
    For Icnt = 0 To frm.Controls.Count - 1
        Select Case TypeName(frm.Controls(Icnt))
            Case "ListBox", "TextBox"
                FlatBorder frm.Controls(Icnt).hwnd, True ' applys flatborder to each control found
        End Select
    Next Icnt
    Icnt = 0
    
End Function

Public Function AddInfoRes(mResFile As String, mInfo As String) As Long
Dim iRet As Long
Dim hUpdate As Long

    hUpdate = BeginUpdateResource(mResFile, False)
    
    If hUpdate = 0 Then AddInfoRes = 0: Exit Function
    
    iRet = UpdateResource(hUpdate, "CUSTOM", 101, 1033, ByVal mInfo, Len(mInfo))
    If iRet = 0 Then AddInfoRes = 0: Exit Function
    
    iRet = EndUpdateResource(hUpdate, False)
    
    If iRet = 0 Then AddInfoRes = 0: Exit Function
    AddInfoRes = 1
    
End Function


