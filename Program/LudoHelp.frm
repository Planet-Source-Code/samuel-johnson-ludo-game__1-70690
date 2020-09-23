VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form LudoHelp 
   Caption         =   "Ludo Help"
   ClientHeight    =   8550
   ClientLeft      =   3015
   ClientTop       =   2025
   ClientWidth     =   9495
   Icon            =   "LudoHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtbHelpFile 
      Height          =   3240
      Left            =   3450
      TabIndex        =   0
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5715
      _Version        =   393217
      BackColor       =   16777215
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      FileName        =   "D:\VB98\sam's vb project\My vb projects\Ludo Game\Program\ludo.txt"
      TextRTF         =   $"LudoHelp.frx":2052
      MouseIcon       =   "LudoHelp.frx":33B6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "LudoHelp"
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

Private Sub Form_Load()
Me.Left = Screen.Height / 2.5
Me.Top = Screen.Height / 8

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
Me.Hide
End Sub


Private Sub Form_Resize()
rtbHelpFile.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
