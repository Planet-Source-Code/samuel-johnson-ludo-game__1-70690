VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Wins 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6990
   ClientLeft      =   2220
   ClientTop       =   2565
   ClientWidth     =   12090
   ControlBox      =   0   'False
   Icon            =   "Won.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Won.frx":076A
   ScaleHeight     =   6990
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFlower 
      Interval        =   2
      Left            =   4680
      Top             =   5040
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   360
      Picture         =   "Won.frx":8A2CC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Quit"
      Top             =   6240
      Width           =   495
   End
   Begin VB.CommandButton Nex 
      Height          =   495
      Left            =   5040
      Picture         =   "Won.frx":8A80C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Continue Game"
      Top             =   6240
      Width           =   495
   End
   Begin VB.CommandButton cmdOk 
      Height          =   735
      Left            =   6720
      Picture         =   "Won.frx":8AD49
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "OK"
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpHailWinner 
      Height          =   495
      Left            =   5700
      TabIndex        =   12
      Top             =   120
      Width           =   390
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   688
      _cy             =   873
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TOSINprof"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   38.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   960
      Left            =   6405
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   5505
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "      THE  CHAMPION              IS"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3585
      Left            =   6960
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   5115
   End
   Begin VB.Image Flowers 
      Height          =   6360
      Left            =   6375
      Picture         =   "Won.frx":8B278
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   5490
   End
   Begin VB.Image Flower 
      Height          =   1050
      Index           =   7
      Left            =   6480
      Picture         =   "Won.frx":ED37A
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   930
   End
   Begin VB.Image Flower 
      Height          =   1050
      Index           =   6
      Left            =   7320
      Picture         =   "Won.frx":13F1A4
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   930
   End
   Begin VB.Image Flower 
      Height          =   1050
      Index           =   5
      Left            =   6480
      Picture         =   "Won.frx":1A12A6
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   930
   End
   Begin VB.Image Flower 
      Height          =   1050
      Index           =   4
      Left            =   7320
      Picture         =   "Won.frx":2033A8
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   930
   End
   Begin VB.Image Flower 
      Height          =   1050
      Index           =   3
      Left            =   6600
      Picture         =   "Won.frx":2654AA
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   930
   End
   Begin VB.Image Flower 
      Height          =   1050
      Index           =   2
      Left            =   7320
      Picture         =   "Won.frx":2E2B2C
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   930
   End
   Begin VB.Image Flower 
      Height          =   1050
      Index           =   1
      Left            =   7320
      Picture         =   "Won.frx":3601AE
      Stretch         =   -1  'True
      Top             =   360
      Width           =   930
   End
   Begin VB.Image Flower 
      Height          =   1050
      Index           =   0
      Left            =   6480
      Picture         =   "Won.frx":3B1FD8
      Stretch         =   -1  'True
      Top             =   360
      Width           =   930
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   375
      Picture         =   "Won.frx":403E02
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   360
      Picture         =   "Won.frx":404826
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   360
      Picture         =   "Won.frx":40524A
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "Won.frx":405C31
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3rd       KAZMA"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   675
      Index           =   3
      Left            =   900
      TabIndex        =   6
      ToolTipText     =   "Lost Out"
      Top             =   4800
      Width           =   3600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " 4th      TOSINprof"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   675
      Index           =   4
      Left            =   465
      TabIndex        =   5
      ToolTipText     =   "Lost Out"
      Top             =   5520
      Width           =   4515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2nd        LOVETH"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   675
      Index           =   1
      Left            =   990
      TabIndex        =   4
      ToolTipText     =   "Potential Champion"
      Top             =   2040
      Width           =   4020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3rd        KAZMA"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   675
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Potencial Loozer"
      Top             =   3120
      Width           =   3750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1st        STORM"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   675
      Index           =   0
      Left            =   1215
      TabIndex        =   2
      ToolTipText     =   "Potential Champion"
      Top             =   1320
      Width           =   3630
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "...Losers"
      BeginProperty Font 
         Name            =   "Script"
         Size            =   60
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1350
      Index           =   1
      Left            =   1425
      TabIndex        =   1
      Top             =   3600
      Width           =   2685
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Winners..."
      BeginProperty Font 
         Name            =   "Script"
         Size            =   60
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1350
      Index           =   0
      Left            =   1170
      TabIndex        =   0
      Top             =   240
      Width           =   3315
   End
End
Attribute VB_Name = "Wins"
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

Dim TrePly As Boolean
Private Sub cmdOk_Click()
 With Board
    For k = 0 To 71
        .BG(k).Picture = LoadPicture("")
        .BG(k).Tag = ""
        bgNum(k) = 0
    Next
    For F = 0 To 3   'they are all in initially
        .Line1(F).Visible = False
        .Ply1(F).Tag = 0
        .Ply2(F).Tag = 0
        .Ply3(F).Tag = 0
        .Ply4(F).Tag = 0
        .Ply1(F).Visible = True
        .Ply2(F).Visible = True
        .Ply3(F).Visible = True
        .Ply4(F).Visible = True
    Next
        .Ply1Home.Picture = LoadPicture("")
        .Ply2Home.Picture = LoadPicture("")
        .Ply3Home.Picture = LoadPicture("")
        .Ply4Home.Picture = LoadPicture("")
        .wmpTalkDie.Close
    
    
    .Initialise 'initialise all form level variables
    
    
 End With
 
 Ptyp1 = 0
 Ptyp2 = 0
 Ptyp3 = 0
 Ptyp4 = 0

 Unload Board
 Unload Wins
 Unload Setn
 Load Setn
 Setn.Show
End Sub
Private Sub Command2_Click()
F = MsgBox("The OverAll Champion Has Not Yet And Must Be Decided." & vbCrLf & _
         "If You Quit, It Will Be Awarded To The First Winner;" + vbCrLf _
         & vbCrLf & "Do You Want To Quit This Game?", vbYesNo + vbInformation)
  If F = vbYes Then
     ShowWinner
     Label4.Caption = Trim(Mid(Label2(0).Caption, 4, Len(Label2(0).Caption)))
  End If
End Sub

Public Sub ShowWinner()
  For J = 0 To 1
  Label1(J).Visible = False
  Next
  For J = 0 To 4
  Label2(J).Visible = False
  Next
 Nex.Visible = False
  Command2.Visible = False
  Flowers.Visible = True
cmdOk.Visible = True
Label3.Visible = True
Label4.Visible = True

End Sub

Private Sub Form_Load()
tmrFlower.Enabled = True
'Flowers.Visible = False
Flowers.Left = 240
'cmdOk.Visible = False
cmdOk.Left = 360
'Label3.Visible = False
Label3.Left = 720
'Label4.Visible = False
Label4.Left = 240

Me.Width = 5970
'Me.Left = Screen.Height / 1.8
'Me.Top = Screen.Height / 5

TrePly = False
wmpHailWinner.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(2).ForeColor = Label2(0).ForeColor
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(2).ForeColor = Label2(0).ForeColor

End Sub

Private Sub Label2_Click(Index As Integer)
If Index <> 2 Then Beep: Exit Sub
If Image1(1).Visible = False Then
Image1(1).Visible = True
Label2(2).ToolTipText = "Potential Loozer"
TrePly = False
Else
Image1(1).Visible = False
Label2(2).ToolTipText = "Potential Champion"
TrePly = True
End If
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 2 Then
Label2(Index).ForeColor = 199
Exit Sub
End If
Label2(2).ForeColor = Label2(0).ForeColor
End Sub

Private Sub Nex_Click()
With Board
.Initialise

For k = 0 To 71
    .BG(k).Picture = LoadPicture("")
    .BG(k).Tag = ""
    bgNum(k) = 0
Next
If Pnum = 4 Then
   .Label4(0).Caption = Pnam1
   .Label4(1).Caption = Pnam2
   .Label4(2).Caption = Pnam3
   .Label4(3).Caption = Pnam4
ElseIf Pnum = 3 Then
   .Label4(0).Caption = Pnam1
   .Label4(1).Caption = Pnam2
   .Label4(2).Caption = Pnam3
   .Label4(3).Caption = ""
End If


'For J = 0 To 3   'removes the first 3 xters 1.e their position
'.Label4(J).Caption = Trim(Mid(.Label4(J).Caption, 4, Len(.Label4(J).Caption)))
'Next
For F = 0 To 3   'they are all in initially
    .Line1(F).Visible = False
    .Ply1(F).Tag = 0
    .Ply2(F).Tag = 0
    .Ply3(F).Tag = 0
    .Ply4(F).Tag = 0
    .Ply1(F).Visible = True
    .Ply2(F).Visible = True
    .Ply3(F).Visible = True
    .Ply4(F).Visible = True
Next
If Pnum = 4 Then
    If Pnam1 = ArrayOfWinners(3) Then
        For F = 0 To 3
            .Ply1(F).Visible = False
            .Ply1(F).Tag = 1
            .Label4(0).Caption = "4th " & Pnam1
        Next
    ElseIf Pnam2 = ArrayOfWinners(3) Then
        For F = 0 To 3
            .Ply2(F).Visible = False
            .Ply2(F).Tag = 1
            .Label4(1).Caption = "4th " & Pnam2
        Next
    ElseIf Pnam3 = ArrayOfWinners(3) Then
        For F = 0 To 3
            .Ply3(F).Visible = False
            .Ply3(F).Tag = 1
            .Label4(2).Caption = "4th " & Pnam3
        Next
    ElseIf Pnam4 = ArrayOfWinners(3) Then
        For F = 0 To 3
            .Ply4(F).Visible = False
            .Ply4(F).Tag = 1
            .Label4(3).Caption = "4th " & Pnam4
        Next
    End If
End If


'find the 3rd and remove it seeds from the board
'now if there are 4 players  and one lost out this will not be performed
'bcos pnum will be = 3
'now after the 3 players have finished playing one goes out and pnum=2
'find the 3rd and remove all it's seed from the board
If PnumX = 2 Then
   If Pnam1 = ArrayOfWinners(2) Then
      For F = 0 To 3
        .Ply1(F).Visible = False
        .Ply1(F).Tag = 1
        .Label4(0).Caption = "3rd " & Pnam1
      Next
   ElseIf Pnam2 = ArrayOfWinners(2) Then
      For F = 0 To 3
        .Ply2(F).Visible = False
        .Ply2(F).Tag = 1
        .Label4(1).Caption = "3rd " & Pnam1
      Next
   ElseIf Pnam3 = ArrayOfWinners(2) Then
      For F = 0 To 3
        .Ply3(F).Visible = False
        .Ply3(F).Tag = 1
        .Label4(2).Caption = "3rd " & Pnam1
      Next
   ElseIf Pnam4 = ArrayOfWinners(2) Then
      For F = 0 To 3
        .Ply4(F).Visible = False
        .Ply4(F).Tag = 1
        .Label4(3).Caption = "3rd " & Pnam1
      Next
   End If
End If
'If P1out = 0 Then
'For f = 0 To 3
'.Ply1(f).Visible = False
'.Ply1(f).Tag = 1
'.Label4(0).Caption = "4th " & Pnam1
'Next
''P1out = 5
'End If
'
'If P2out = 0 Then
'For f = 0 To 3
'.Ply2(f).Visible = False
'.Ply2(f).Tag = 1
'.Label4(1).Caption = "4th " & Pnam2
'Next
''P2out = 5
'End If
'If P3out = 0 Then
'For f = 0 To 3
'.Ply3(f).Visible = False
'.Ply3(f).Tag = 1
'.Label4(2).Caption = "4th " & Pnam3
'Next
''P3out = 5
'End If
'If P4out = 0 Or P4out = 999 Then
'For f = 0 To 3
'.Ply4(f).Visible = False
'.Ply4(f).Tag = 1
'.Label4(3).Caption = "4th " & Pnam4
'Next
''P4out = 5
'End If
.Ply1Home.Picture = LoadPicture("")
.Ply2Home.Picture = LoadPicture("")
.Ply3Home.Picture = LoadPicture("")
.Ply4Home.Picture = LoadPicture("")
Ply1Hom = 0
Ply2Hom = 0
Ply3Hom = 0
Ply4Hom = 0

.TransSlate.Enabled = True
      
If Pnum = 4 And PnumX <> 2 Then 'to determine if 3rd winner will be included or not
    If TrePly = False Then PnumX = 2 Else PnumX = 3

    If Pnam1 = ArrayOfWinners(2) Then  '3rd position
        If TrePly = True Then  'user included the 3rd winner
            For F = 0 To 3
                .Ply1(F).Visible = True
                .Ply1(F).Tag = 0
            Next
            .Label4(0).Caption = Pnam1
        Else 'user don't want 3rd winner to continue
            For F = 0 To 3
                .Ply1(F).Visible = False
                .Ply1(F).Tag = 1
            Next
            .Label4(0).Caption = "3rd " & Pnam1
        End If
    
    ElseIf Pnam2 = ArrayOfWinners(2) Then
        If TrePly = True Then
            For F = 0 To 3
                .Ply2(F).Visible = True
                .Ply2(F).Tag = 0
            Next
            .Label4(1).Caption = Pnam2
        Else
            For F = 0 To 3
                .Ply2(F).Visible = False
                .Ply2(F).Tag = 1
            Next
            .Label4(1).Caption = "3rd " & Pnam2
        End If
    
    ElseIf Pnam3 = ArrayOfWinners(2) Then
    
        If TrePly = True Then
            For F = 0 To 3
               .Ply3(F).Visible = True
               .Ply3(F).Tag = 0
            Next
            .Label4(2).Caption = Pnam3
        Else
            For F = 0 To 3
               .Ply3(F).Visible = False
               .Ply3(F).Tag = 1
            Next
            .Label4(2).Caption = "3rd " & Pnam3
        End If
    
    ElseIf Pnam4 = ArrayOfWinners(2) Then
    
        If TrePly = True Then
            For F = 0 To 3
                .Ply4(F).Visible = True
                .Ply4(F).Tag = 0
            Next
           .Label4(3).Caption = Pnam4

        Else
            For F = 0 To 3
                .Ply4(F).Visible = False
                .Ply4(F).Tag = 1
            Next
           .Label4(3).Caption = "3rd " & Pnam4
       End If
    End If
    
End If
 
 
 If PnumX = 2 Then  'this not repitition,incase we have 3 players and one is out
    If Pnam1 = ArrayOfWinners(2) Then
          .Label4(0).Caption = "3rd " & Pnam1
     ElseIf Pnam2 = ArrayOfWinners(2) Then
          .Label4(1).Caption = "3rd " & Pnam2
     ElseIf Pnam3 = ArrayOfWinners(2) Then
          .Label4(2).Caption = "3rd " & Pnam3
     ElseIf Pnam4 = ArrayOfWinners(2) Then
          .Label4(3).Caption = "3rd " & Pnam4
     End If
 End If

If Pnum = 3 Then 'only 3 players started from the onset that means player four was never used remove all it's seed from board
            For F = 0 To 3
                .Ply4(F).Visible = False
                .Ply4(F).Tag = 1
            Next
           .Label4(3).Caption = ""
End If
End With
'also got to initialise some variables here
'this project has taken my time more than i bargained for -epileptic power supply
P1out = 0
P2out = 0
P3out = 0
P4out = 0
Unload Me





End Sub

Private Sub tmrFlower_Timer()
Static k As Integer
Randomize
k = 7 * Rnd
Flowers.Picture = Flower(k).Picture
tmrFlower.Enabled = False
End Sub
