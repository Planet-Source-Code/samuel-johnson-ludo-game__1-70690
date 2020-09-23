VERSION 5.00
Begin VB.Form Splash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5535
   ClientLeft      =   3570
   ClientTop       =   2730
   ClientWidth     =   7455
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Splash.frx":2CCA
   ScaleHeight     =   5535
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1200
      Top             =   3240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H001E1E1E&
      Caption         =   "   ...Creativity speaks for itself"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004588CB&
      Height          =   255
      Left            =   4080
      TabIndex        =   0
      Top             =   5280
      Width           =   3375
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Counter As Integer

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
'Counter = 200
End Sub

Private Sub Form_Load()
'Me.Left = Screen.Height / 2
'Me.Top = Screen.Height / 4
 Dirxtry = App.Path & "\"

End Sub

Private Sub Timer1_Timer()
Static p As Integer
Counter = Counter + 5
'ProgressBar1.Value = Counter
p = (p Mod 8) + 1
Select Case p
Case 1
Label1.Caption = "Creativity Speaks For Itself."
Case 3
Label1.Caption = "Creativity Speaks For Itself.."
Case 5
Label1.Caption = "Creativity Speaks For Itself..."
Case 7
Label1.Caption = "Creativity Speaks For Itself...."
End Select
If Counter = 50 Then
Load Setn
ElseIf Counter = 100 Then
Load Board
ElseIf Counter = 150 Then
Setn.SoundClick.URL = Dirxtry & "2wav.wav"
Load Wins
ElseIf Counter >= 200 Then
Load About
Load LudoHelp
Unload Me
Setn.Show
Setn.SoundClick.URL = Dirxtry & "3wav.wav"
Timer1.Enabled = False
End If
End Sub
