VERSION 5.00
Begin VB.Form About 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FCFCFC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Ludo Game"
   ClientHeight    =   5205
   ClientLeft      =   2865
   ClientTop       =   3225
   ClientWidth     =   7095
   ClipControls    =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Proudly Nigerian"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5550
      TabIndex        =   11
      Top             =   720
      Width           =   1890
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   6150
      Picture         =   "About.frx":076A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "It All Began With A Thought..."
      Height          =   255
      Left            =   4275
      TabIndex        =   10
      Top             =   1680
      Width           =   2265
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "This Project Would Have Been Impossible Without God's Mercy; All Glory to HIM."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   75
      TabIndex        =   9
      Top             =   4920
      Width           =   6990
   End
   Begin VB.Label Label9 
      BackColor       =   &H00F5F5F5&
      Caption         =   "Both Are Currently Students Of YABA COLLEGE OF TECHNOLOGY In The Department Of Computer Science And In Their 2nd Year. "
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   375
      TabIndex        =   8
      Top             =   4440
      Width           =   6390
   End
   Begin VB.Label Label8 
      BackColor       =   &H00F5F5F5&
      Caption         =   $"About.frx":0E1F
      ForeColor       =   &H00808080&
      Height          =   1095
      Left            =   3675
      TabIndex        =   7
      Top             =   3360
      Width           =   3090
   End
   Begin VB.Label Label7 
      BackColor       =   &H00F5F5F5&
      Caption         =   $"About.frx":0EF3
      ForeColor       =   &H00808080&
      Height          =   1095
      Left            =   375
      TabIndex        =   6
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ABOUT THE AUTHORS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2550
      TabIndex        =   5
      Top             =   3120
      Width           =   2040
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "We Finally Present :"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3075
      TabIndex        =   4
      Top             =   360
      Width           =   2640
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "From STORM GAMES"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1725
      TabIndex        =   3
      Top             =   120
      Width           =   3465
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":0FAA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      TabIndex        =   2
      Top             =   1920
      Width           =   5640
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(c) April - May 2008 STORM GAMES"
      Height          =   255
      Left            =   3300
      TabIndex        =   1
      Top             =   1200
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   75
      Picture         =   "About.frx":1128
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The Ludo Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1725
      TabIndex        =   0
      Top             =   720
      Width           =   3915
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'??????????????????????????????????????????????????????????
'**********************************************************
'Program Title: The Ludo Game  version 1.0.0
'Author:        Samuel Johnson A.c
'Date:          April to May 2008
'copyright:     © 2008
'Email:         stormsamany@yahoo.co.uk
'Phone:         +234 808 574 6108

'Last Updated:  2nd of June 2008 by the Author
'if you have any problem going through this code
'you can contact the Author through the email address



' All Right Reserved

'It considered a criminal offence if this Program
'in whole or in part is published under a diffrent
'name, title or any alteration which may affect the
'functionality of the game for better or worse
'of which such changes were made, with the aim of
'publishing or reproducing such altered copies
'without the prior permission of the Author
'in a stored audio format (e.g a casette or audio disc)

'??????????????????????????????????????????????????????????
'**********************************************************

Private Sub Form_DblClick()
Me.Hide
End Sub

Private Sub Form_Load()
'Me.Left = Screen.Height / 2
'Me.Top = Screen.Height / 4

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
Me.Hide
End Sub

Private Sub Image1_DblClick()
Me.Hide
End Sub

Private Sub Image2_DblClick()
Me.Hide
End Sub

Private Sub Label1_DblClick()
Me.Hide
End Sub

Private Sub Label10_DblClick()
Me.Hide
End Sub

Private Sub Label3_DblClick()
Me.Hide
End Sub

Private Sub Label4_DblClick()
Me.Hide
End Sub

Private Sub Label5_DblClick()
Me.Hide
End Sub
'??????????????????????????????????????????????????????????
'**********************************************************
'Program Title: The Ludo Game  version 1.0.0
'Author:        Samuel Johnson A.c
'Date:          April to May 2008
'copyright:     © 2008
'Email:         stormsamany@yahoo.co.uk
'Phone:         +234 808 574 6108

'Last Updated:  2nd of June 2008 by the Author
'if you have any problem going through this code
'you can contact the Author through the email address



' All Right Reserved

'It considered a criminal offence if this Program
'in whole or in part is published under a diffrent
'name, title or any alteration which may affect the
'functionality of the game for better or worse
'of which such changes were made, with the aim of
'publishing or reproducing such altered copies
'without the prior permission of the Author
'in a stored audio format (e.g a casette or audio disc)

'??????????????????????????????????????????????????????????
'**********************************************************

