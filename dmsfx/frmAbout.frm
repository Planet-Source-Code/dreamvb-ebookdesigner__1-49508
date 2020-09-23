VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "eBook Info"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   350
      Left            =   1425
      TabIndex        =   3
      Top             =   3240
      Width           =   885
   End
   Begin VB.PictureBox picbase 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   2670
      Left            =   0
      ScaleHeight     =   2610
      ScaleWidth      =   3660
      TabIndex        =   0
      Top             =   0
      Width           =   3720
      Begin VB.Label compdate 
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
         Height          =   240
         Left            =   1695
         TabIndex        =   6
         Top             =   1905
         Width           =   1920
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Compiled on:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   405
         TabIndex        =   5
         Top             =   1905
         Width           =   1215
      End
      Begin VB.Label lblauhor 
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   405
         TabIndex        =   2
         Top             =   1515
         Width           =   2940
      End
      Begin VB.Label lbltitle 
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   405
         TabIndex        =   1
         Top             =   1110
         Width           =   2940
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   285
         X2              =   3495
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Image Image2 
         Height          =   795
         Left            =   795
         Picture         =   "frmAbout.frx":0000
         Top             =   90
         Width           =   2310
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   90
         Picture         =   "frmAbout.frx":0F76
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "eBook created with DM ebook Designer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   2865
      Width           =   3390
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
    Unload frmAbout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbout = Nothing
End Sub
