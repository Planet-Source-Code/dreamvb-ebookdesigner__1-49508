Attribute VB_Name = "ModHelp"
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

Private mHelp() As String

Public Sub LoadHelpFile(lzFile As String)
Dim hlpFile As Long, CntLn As Long, iPart As Long, StrLn As String
On Error Resume Next

    hlpFile = FreeFile
    Open lzFile For Input As #hlpFile
        Do While Not EOF(hlpFile)
            Input #hlpFile, StrLn
            iPart = InStr(1, StrLn, "=", vbTextCompare)
            If iPart > 0 Then CntLn = CntLn + 1
            ReDim Preserve mHelp(1 To CntLn)
            mHelp(CntLn) = Mid(StrLn, iPart + 1, Len(StrLn) - iPart)
            DoEvents
        Loop
    Close #hlpFile
    CntLn = 0
    iPart = 0
    StrLn = ""
    
End Sub

Public Sub Showhelp(HelpID As Integer)
On Error Resume Next
    frmhelp.lblhelp.Caption = mHelp(HelpID)
    frmhelp.Show vbModal
End Sub
