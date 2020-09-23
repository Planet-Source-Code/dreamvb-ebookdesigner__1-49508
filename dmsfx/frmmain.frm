VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmmain 
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebView 
      Height          =   690
      Left            =   15
      TabIndex        =   2
      Top             =   870
      Width           =   1530
      ExtentX         =   2699
      ExtentY         =   1217
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6015
      Top             =   3405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":124C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2550
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2ED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3854
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":41D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4B58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar1 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   1323
      ButtonWidth     =   1244
      ButtonHeight    =   1270
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "M_BK"
            Object.ToolTipText     =   "Go Back"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "M_FW"
            Object.ToolTipText     =   "Go Forward"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "M_STP"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            Key             =   "M_HOME"
            Object.ToolTipText     =   "Go Home"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Key             =   "M_FIND"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "M_PRT"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            Key             =   "M_IFO"
            Object.ToolTipText     =   "Info"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "M_EX"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Stb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3990
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11139
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   0
      X2              =   795
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   795
      Y1              =   810
      Y2              =   810
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ebook Viewer
' Coded and Designed by Ben Jones
' Created on 28/10/03
' Email dreamvb@yahoo.com or vbdream2k@yahoo.com

Dim HomePage As String ' Holds the ebooks default home page

Private Function FixPath(lzPath As String) As String
    ' Fixes a path by adding a back slash if required
    If Right$(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Private Sub Form_Load()
On Error Resume Next
Dim StrB As String, StrV As Variant

    StrB = StrConv(LoadResData(101, "CUSTOM"), vbUnicode) ' Losd in the ebooks info
    StrV = Split(StrB, ":")
   
    frmmain.Caption = StrV(0)
    frmAbout.lbltitle.Caption = StrV(0) ' Set the title of the ebook
    frmAbout.lblauhor.Caption = StrV(1) ' Set the auhors name of the ebook
    frmAbout.compdate.Caption = StrV(3) ' Set the compile ebook date
    
    If Len(StrV(2)) = 0 Then
        MsgBox "Unable to locate the main main index page.", vbCritical, "Error Loading Home Page"
        Exit Sub ' There was an error finding the home page to stop here
    End If
    
    HomePage = "RES://" & FixPath(App.Path) & App.EXEName & ".exe/" & StrV(2) ' Setup up home page location
    WebView.Navigate HomePage ' Move to the default home page
    
    StrB = "" ' Clear var
    Erase StrV ' Erase StrV array
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    ' Code below used to resize the controls on the form
    Line1(0).X2 = frmmain.ScaleWidth
    Line1(1).X2 = frmmain.ScaleWidth
    WebView.Width = frmmain.ScaleWidth - WebView.Left
    WebView.Height = (frmmain.ScaleHeight - WebView.Top) - Stb.Height
    If Err Then Err.Clear
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing ' Release form from memory
    End ' End the program
End Sub

Private Sub tBar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

    Select Case Button.Key
        Case "M_BK"
            WebView.GoBack ' Made the broswer go back a page
        Case "M_FW"
            WebView.GoForward ' Made the broswer go forward a page
        Case "M_STP"
            WebView.Stop ' Stop current loading of a page
        Case "M_HOME"
            WebView.Navigate HomePage ' Move to home page
        Case "M_FIND"
            WebView.SetFocus ' Set focus on the web control
            SendKeys "^f" ' Send key action to show find dialog
        Case "M_PRT"
            WebView.SetFocus ' Set focus on the web control
            SendKeys "^p"   ' Send key action to show pring dialog
        Case "M_IFO"
            frmAbout.Show vbModal, frmmain ' Show the ebooks about box
        Case "M_EX"
            HomePage = "" ' Clear home page buffer
            Unload frmmain ' Unload the form
            End ' End the program
    End Select
    
End Sub

