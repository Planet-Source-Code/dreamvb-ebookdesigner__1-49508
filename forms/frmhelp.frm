VERSION 5.00
Begin VB.Form frmhelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pop up Help"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picpophelp 
      Align           =   1  'Align Top
      BackColor       =   &H80000018&
      Height          =   2145
      Left            =   0
      ScaleHeight     =   2085
      ScaleWidth      =   4005
      TabIndex        =   0
      Top             =   0
      Width           =   4065
      Begin VB.CommandButton Command1 
         Caption         =   "Close Help"
         Height          =   300
         Left            =   2775
         TabIndex        =   3
         Top             =   1695
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         X1              =   570
         X2              =   3570
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DM EBook Designer Beta 1"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   570
         TabIndex        =   2
         Top             =   150
         Width           =   3315
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   60
         Picture         =   "frmhelp.frx":0000
         Top             =   60
         Width           =   480
      End
      Begin VB.Label lblhelp 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   585
         TabIndex        =   1
         Top             =   555
         Width           =   3360
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmhelp"
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

Private Sub Command1_Click()
    Unload frmhelp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lblhelp.Caption = ""
    Set frmhelp = Nothing
End Sub

