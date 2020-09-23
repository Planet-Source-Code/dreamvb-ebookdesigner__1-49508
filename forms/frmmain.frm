VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM Ebook Designer Beta 1"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   350
      Left            =   4125
      TabIndex        =   19
      Top             =   6195
      Width           =   1215
   End
   Begin VB.CommandButton cmdabout 
      Caption         =   "&About"
      Height          =   350
      Left            =   2655
      TabIndex        =   18
      Top             =   6195
      Width           =   1215
   End
   Begin VB.CommandButton cmdcompile 
      Caption         =   "&Compile"
      Height          =   350
      Left            =   1170
      TabIndex        =   17
      Top             =   6195
      Width           =   1215
   End
   Begin VB.CommandButton cmdoutput 
      Caption         =   "...."
      Height          =   330
      Left            =   4530
      TabIndex        =   16
      Top             =   4867
      Width           =   510
   End
   Begin VB.TextBox txtoutfolder 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4890
      Width           =   2880
   End
   Begin VB.CommandButton cmdselall 
      Caption         =   "&Select All"
      Height          =   330
      Left            =   3885
      TabIndex        =   12
      Top             =   3270
      Width           =   1230
   End
   Begin VB.TextBox txthomepage 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   10
      Top             =   2430
      Width           =   2325
   End
   Begin VB.ListBox lstFiles 
      Height          =   1185
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   3300
      Width           =   3660
   End
   Begin VB.TextBox txtoutfile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1530
      TabIndex        =   7
      Top             =   5340
      Width           =   3375
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "...."
      Height          =   345
      Left            =   4560
      TabIndex        =   6
      Top             =   1665
      Width           =   435
   End
   Begin VB.TextBox txtpath 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   5
      Top             =   1680
      Width           =   4350
   End
   Begin VB.TextBox txtAuthor 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1620
      TabIndex        =   3
      Top             =   780
      Width           =   3450
   End
   Begin VB.TextBox txttitle 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1620
      TabIndex        =   1
      Top             =   315
      Width           =   3450
   End
   Begin VB.Image imghelp 
      Height          =   210
      Index           =   6
      Left            =   5145
      Tag             =   "7"
      Top             =   5400
      Width           =   240
   End
   Begin VB.Image imghelp 
      Height          =   210
      Index           =   5
      Left            =   5145
      Tag             =   "6"
      Top             =   4920
      Width           =   240
   End
   Begin VB.Image imghelp 
      Height          =   210
      Index           =   4
      Left            =   5145
      Tag             =   "5"
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image imghelp 
      Height          =   210
      Index           =   3
      Left            =   2580
      Tag             =   "4"
      Top             =   2490
      Width           =   240
   End
   Begin VB.Image imghelp 
      Height          =   210
      Index           =   2
      Left            =   5145
      Tag             =   "3"
      Top             =   1732
      Width           =   240
   End
   Begin VB.Image imghelp 
      Height          =   210
      Index           =   0
      Left            =   5145
      Tag             =   "1"
      Top             =   352
      Width           =   240
   End
   Begin VB.Image imghelp 
      Height          =   210
      Index           =   1
      Left            =   5145
      Tag             =   "2"
      Top             =   840
      Width           =   240
   End
   Begin VB.Image img1 
      Height          =   210
      Left            =   165
      Picture         =   "frmmain.frx":08CA
      Top             =   7935
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   5310
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   15
      X2              =   5325
      Y1              =   5895
      Y2              =   5895
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output File:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   15
      Top             =   5400
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output Folder:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   13
      Top             =   4935
      Width           =   1230
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   5370
      Y1              =   4725
      Y2              =   4725
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   75
      X2              =   5385
      Y1              =   4740
      Y2              =   4740
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Please check the files you want to include:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   11
      Top             =   2940
      Width           =   3645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Index Page: Include the name of your index page:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   9
      Top             =   2175
      Width           =   4365
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   135
      X2              =   5445
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5430
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Files:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   1425
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Author Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ebook Title:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   345
      Width           =   1020
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As Integer, ByVal lpName As String, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long

Sub LbSelect()
Dim Icount As Long
' I used this to select or deselect all items in a listbox
    For Icount = 0 To lstFiles.ListCount - 1
        lstFiles.Selected(Icount) = True
    Next
    
    lstFiles.Refresh
    Icount = 0
    
End Sub

Private Sub cmdabout_Click()
Dim Msg As String
    
    Msg = Msg & "DM Ebook Designer Beta 1" _
    & vbCrLf & vbCrLf & "THIS PROGRAM IS FREEWARE." & vbCrLf & vbCrLf & "Simply and easily create your own Ebook for Windows." _
    & vbCrLf & vbCrLf & "Program written and designed by Ben Jones." _
    & vbCrLf & vbCrLf & "Please send any comments or questions to: " _
    & vbCrLf & "VBdream2k@yahoo.com"
    
    MsgBox Msg, vbInformation, "About... " & frmmain.Caption ' Shows the programs about box.
    
End Sub

Private Sub cmdcompile_Click()
Dim sHead As String, mBookSfx As String, StrData As String, iResult As Long, iRet As Long, hUpdate As Long
Dim lzFullPath As String, I As Long, OutputDir As String, OutPutFile As String

    ' The code below just does some simple validation checks
    If Len(BookPath) = 0 Then
        MsgBox "No project files were found please select your project folder.", vbInformation, frmmain.Caption
        Exit Sub
    ElseIf lstFiles.ListCount = 0 Then
        MsgBox "No project files were found please select your project folder.", vbInformation, frmmain.Caption
        Exit Sub
    ElseIf Len(Trim(txthomepage.Text)) = 0 Then
        MsgBox "You need to include the name of your default index page.", vbInformation, frmmain.Caption
        Exit Sub
    ElseIf FindFile(BookPath & txthomepage.Text) = False Then
        MsgBox "The index page name you entered was not found." & vbCrLf & vbCrLf & BookPath & txthomepage, vbCritical, frmmain.Caption
        Exit Sub
    ElseIf Len(Trim(txtoutfile.Text)) = 0 Then
        MsgBox "You have not entered in a name for your ebook. A default name will be provided for you.", vbInformation, frmmain.Caption
    ElseIf Len(txtoutfile.Text) > 0 And Not UCase(Right(txtoutfile.Text, 4)) = ".EXE" Then
        txtoutfile.Text = txtoutfile.Text & ".exe"
    End If
    
    ' Update the Ebook type Struc
    TEBook.eBookTitle = txttitle.Text       ' Add ebook title
    TEBook.eBookAuthor = txtAuthor.Text     ' Add the Author of the ebook
    TEBook.eBookHomePage = txthomepage.Text     ' Add the default index page will be the start page
    TEBook.eBookkExeName = txtoutfile.Text   ' This will become executable name of ebook
    TEBook.eBookCompDate = Format(Date, "DD/MMM/YY") ' Add in the compile date the ebook was made
    
    If Trim(Len(txttitle.Text)) = 0 Then ' If no title was found set a default one
        TEBook.eBookTitle = m_def_ebookTitle
    End If
    
    If Trim(Len(txtAuthor.Text)) = 0 Then
        TEBook.eBookAuthor = m_def_eBookAuthor ' If no author was set a default one
    End If
    
    If Trim(Len(txtoutfile.Text)) = 0 Then
        TEBook.eBookkExeName = m_def_eBookEXEName ' If no ebook exe name was found set the default one
    End If

    OutputDir = txtoutfolder.Text ' Path of the ebook
    OutPutFile = OutputDir & TEBook.eBookkExeName ' Full path and Filename of the ebook

    ' Ok the code below will copy the ebook sfx file to the new file for the ebook.
    ' You can find the sfx file in the dmsfx folder of this project.
    ' You need to compile first as you already no that that :)
    
    mBookSfx = FixPath(App.Path) & "dmsfx\dmsfx.dat" ' Path to the ebook executable file

    If FindFile(mBookSfx) = False Then ' Check to see if the sfx file exsists
        MsgBox "There was an error finding the file:" & vbCrLf & vbCrLf _
        & "Please check that all files are installed correctly.", vbCritical, frmmain.Caption
        Exit Sub
    Else
        ' Below we use the code to copy the sfx file to
        ' were the ebook is to be created as we do not want to chnage the original sfx file
        FileCopy mBookSfx, OutPutFile
    End If
    
    sHead = TEBook.eBookTitle & ":" & TEBook.eBookAuthor & ":" & TEBook.eBookHomePage _
    & ":" & TEBook.eBookCompDate ' Information for the ebook

    iResult = AddInfoRes(OutPutFile, sHead)
    sHead = "" ' Clear info buffer
    
    If Not iResult > 0 Then
        MsgBox "There was an error while compileing your ebook", vbCritical, "Compile Error"
        Kill OutPutFile ' Kill the output file we no need for this now
        Exit Sub ' we can also Stop here
    Else
        For I = 0 To lstFiles.ListCount - 1 ' Loop though the listbox
            If lstFiles.Selected(I) = True Then ' See if the item is selected
                lzFullPath = BookPath & lstFiles.List(I) ' Get the full path of the file in the listbox
                StrData = OpenFile(lzFullPath) ' Get the data from the file
                hUpdate = BeginUpdateResource(OutPutFile, False) ' Get the hangle of the file
                iRet = UpdateResource(hUpdate, 2110, UCase(lstFiles.List(I)), 1033, ByVal StrData, Len(StrData)) ' Update files resource with our ebook data
                iRet = EndUpdateResource(hUpdate, False) ' Save the data to the file
            End If
        Next
    End If
    
    ' Clear vars
    StrData = ""
    lzFullPath = ""
    iRet = 0
    MsgBox "Your ebook has now been compiled to" & vbCrLf & vbCrLf & OutPutFile, vbInformation, "Compile Finished."
    
End Sub

Private Sub cmdExit_Click()
    Unload frmmain ' Unload the form
End Sub

Private Sub cmdopen_Click()
Dim FolName As String
Dim X As String

    FolName = FixPath(GetFolder(frmmain.hwnd, "Choose the folder were all your files are in below:"))
    If Len(FolName) <= 1 Then
        Exit Sub
        ' Exit out sub if the folder length is lower of equal to 1
    Else
        txtpath.Text = FolName ' Assign textbox with folder path
        BookPath = FolName ' Assign Bookpath with folder path
        txtoutfolder.Text = FolName
        lstFiles.Clear   ' Clear the list box
    
        X = Dir(FolName)
        ' Code below will loop though all the files in the folder.
        ' Note I not added code for sub folders.
        ' If you like to add this in please do so.
        Do While X <> ""
            lstFiles.AddItem X
            X = Dir
            DoEvents
        Loop
    End If

End Sub

Private Sub cmdoutput_Click()
Dim FolName As String
    FolName = FixPath(GetFolder(frmmain.hwnd, "Choose the folder were all your files are in below:"))
    If Len(FolName) <= 1 Then Exit Sub
    ' Exit out sub if the folder length is lower of equal to 1
    txtoutfolder.Text = FolName
    
End Sub

Private Sub cmdselall_Click()
    LbSelect ' Select all items in the listbox
End Sub

Private Sub Form_Load()
    MakeFlatControls frmmain ' Function that turns all controls flat see ModMain
    LoadHelpFile FixPath(App.Path) & "help.txt"
    For I = 0 To imghelp.Count - 1
        imghelp(I).Picture = img1.Picture
    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing ' Release the form from memory
End Sub

Private Sub imghelp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imghelp(Index).BorderStyle = 1
End Sub

Private Sub imghelp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imghelp(Index).BorderStyle = 0
    Showhelp imghelp(Index).Tag
End Sub
