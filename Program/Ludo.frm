VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Board 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   9795
   ClientLeft      =   2460
   ClientTop       =   1740
   ClientWidth     =   12720
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   Icon            =   "Ludo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Ludo.frx":2CCA
   Moveable        =   0   'False
   Picture         =   "Ludo.frx":5994
   ScaleHeight     =   9795
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrVibrate 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3750
      Top             =   480
   End
   Begin VB.CommandButton PauseStart 
      Height          =   1215
      Left            =   9975
      Picture         =   "Ludo.frx":5E9C5
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Pause Game (Pause or Enter)"
      Top             =   840
      Width           =   1065
   End
   Begin VB.CommandButton MustTrans 
      DisabledPicture =   "Ludo.frx":5F5B9
      Height          =   1215
      Left            =   9975
      Picture         =   "Ludo.frx":600E4
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Transfer Slate (Arrow Key)"
      Top             =   2760
      Width           =   1065
   End
   Begin VB.CommandButton Help 
      Height          =   1215
      Left            =   9975
      Picture         =   "Ludo.frx":60D92
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Help (F1)"
      Top             =   4680
      Width           =   1065
   End
   Begin VB.CommandButton cmdAbout 
      Height          =   1215
      Left            =   9975
      Picture         =   "Ludo.frx":61AD1
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "About (F2)"
      Top             =   6720
      Width           =   1065
   End
   Begin VB.CommandButton Quit 
      Height          =   615
      Left            =   10200
      Picture         =   "Ludo.frx":62780
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Quit (Esc)"
      Top             =   8760
      Width           =   585
   End
   Begin VB.Timer tmrToolSlider 
      Interval        =   200
      Left            =   3360
      Top             =   480
   End
   Begin VB.Timer tmrSpeakOut 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   480
   End
   Begin VB.Timer moveSeedAI 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   480
   End
   Begin VB.Timer throwdiceAI 
      Left            =   2280
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1920
      Top             =   480
   End
   Begin VB.PictureBox Motion 
      Height          =   375
      Left            =   16290
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   22
      Top             =   240
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   1560
      Top             =   480
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   14550
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2280
      Width           =   795
   End
   Begin VB.Timer TransSlate 
      Enabled         =   0   'False
      Interval        =   2600
      Left            =   1200
      Top             =   480
   End
   Begin VB.PictureBox Dice 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   6
      Left            =   15075
      Picture         =   "Ludo.frx":62DB8
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   19
      Top             =   2640
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Index           =   1
      Left            =   15225
      Picture         =   "Ludo.frx":63097
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   18
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Index           =   0
      Left            =   15375
      Picture         =   "Ludo.frx":63342
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   2640
      Width           =   375
   End
   Begin VB.Timer tmrDice 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   480
   End
   Begin VB.PictureBox Dice 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   16500
      Picture         =   "Ludo.frx":635DC
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox Dice 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   15300
      Picture         =   "Ludo.frx":6395E
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Dice 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   15540
      Picture         =   "Ludo.frx":63CD8
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Dice 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   15930
      Picture         =   "Ludo.frx":64048
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Dice 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   15300
      Picture         =   "Ludo.frx":643C9
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Dice 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   15540
      Picture         =   "Ludo.frx":64745
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   3270
      TabIndex        =   0
      Top             =   8400
      Width           =   3495
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         DownPicture     =   "Ludo.frx":64ABA
         Height          =   615
         Left            =   0
         Picture         =   "Ludo.frx":64F34
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Die 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   1
         Left            =   2520
         Picture         =   "Ludo.frx":651DF
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   3
         Top             =   120
         Width           =   495
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Index           =   1
            X1              =   480
            X2              =   0
            Y1              =   480
            Y2              =   0
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Index           =   0
            X1              =   0
            X2              =   480
            Y1              =   480
            Y2              =   0
         End
      End
      Begin VB.PictureBox Die 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   3000
         Picture         =   "Ludo.frx":65559
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   120
         Width           =   495
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Index           =   3
            X1              =   480
            X2              =   0
            Y1              =   480
            Y2              =   0
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Index           =   2
            X1              =   0
            X2              =   480
            Y1              =   480
            Y2              =   0
         End
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00020002&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004588CB&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Your Turn To Roll The Dice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004588CB&
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Timer tmrMoved 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   450
      Top             =   480
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpFreedSound 
      Height          =   495
      Left            =   12150
      TabIndex        =   32
      Top             =   360
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
   Begin WMPLibCtl.WindowsMediaPlayer wmpPopup 
      Height          =   495
      Left            =   11775
      TabIndex        =   31
      Top             =   360
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Don't Delete this label it's  for preventn the user from interacting with the game when paused"
      Height          =   615
      Left            =   3075
      TabIndex        =   30
      Top             =   4800
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      BorderWidth     =   5
      Index           =   0
      X1              =   11325
      X2              =   11325
      Y1              =   0
      Y2              =   9840
   End
   Begin VB.Image Pause 
      Height          =   1170
      Index           =   1
      Left            =   13125
      Picture         =   "Ludo.frx":658CE
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1170
   End
   Begin VB.Image Pause 
      Height          =   1170
      Index           =   0
      Left            =   13500
      Picture         =   "Ludo.frx":6645B
      Top             =   1680
      Width           =   1170
   End
   Begin VB.Image Starts 
      Height          =   1170
      Index           =   0
      Left            =   13950
      Picture         =   "Ludo.frx":6AC85
      Top             =   960
      Width           =   1170
   End
   Begin VB.Image Starts 
      Height          =   1170
      Index           =   1
      Left            =   13950
      Picture         =   "Ludo.frx":6B8C8
      Top             =   840
      Width           =   1170
   End
   Begin VB.Image Image3 
      Height          =   1170
      Left            =   13950
      Picture         =   "Ludo.frx":6C461
      Top             =   840
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   9855
      Left            =   9750
      Picture         =   "Ludo.frx":6D1BA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   9735
      Left            =   9240
      TabIndex        =   24
      Top             =   0
      Width           =   615
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpTalkDie 
      Height          =   495
      Left            =   15465
      TabIndex        =   23
      Top             =   600
      Width           =   615
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
      _cx             =   1085
      _cy             =   873
   End
   Begin VB.Image Ply3 
      Height          =   600
      Index           =   3
      Left            =   2400
      Picture         =   "Ludo.frx":F6D1C
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   600
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   735
      Left            =   14850
      TabIndex        =   20
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H004588CB&
      Caption         =   "Label5"
      ForeColor       =   &H004588CB&
      Height          =   615
      Left            =   15705
      TabIndex        =   16
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player4"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Index           =   3
      Left            =   5760
      TabIndex        =   15
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player3"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Index           =   2
      Left            =   720
      TabIndex        =   14
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   13
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Index           =   0
      Left            =   5880
      TabIndex        =   12
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Image Ply4Home 
      Height          =   495
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   495
   End
   Begin VB.Image Ply3Home 
      Height          =   495
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   495
   End
   Begin VB.Image Ply2Home 
      Height          =   495
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image Store 
      Height          =   375
      Index           =   1
      Left            =   15585
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Store 
      Height          =   375
      Index           =   0
      Left            =   15585
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Ply1Home 
      Height          =   495
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   495
   End
   Begin VB.Image Ply2_1 
      Height          =   375
      Left            =   15840
      Picture         =   "Ludo.frx":F7A5C
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image Ply1_1 
      Height          =   375
      Left            =   15840
      Picture         =   "Ludo.frx":FF80E
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image Ply3_1 
      Height          =   375
      Left            =   15840
      Picture         =   "Ludo.frx":109D38
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image Ply4_1 
      Height          =   375
      Left            =   15840
      Picture         =   "Ludo.frx":112E5B
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image Ply4_4 
      Height          =   375
      Left            =   16920
      OLEDropMode     =   1  'Manual
      Picture         =   "Ludo.frx":11CD35
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image Ply4_3 
      Height          =   375
      Left            =   16560
      Picture         =   "Ludo.frx":12E748
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image Ply4_2 
      Height          =   375
      Left            =   16200
      Picture         =   "Ludo.frx":13CA5B
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image Ply3_4 
      Height          =   375
      Left            =   16920
      Picture         =   "Ludo.frx":148A20
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image Ply3_3 
      Height          =   375
      Left            =   16560
      Picture         =   "Ludo.frx":1568EB
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image Ply3_2 
      Height          =   375
      Left            =   16200
      Picture         =   "Ludo.frx":1646B9
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image Ply2_4 
      Height          =   375
      Left            =   16920
      Picture         =   "Ludo.frx":16F46E
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image Ply2_3 
      Height          =   375
      Left            =   16560
      Picture         =   "Ludo.frx":17D29F
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image Ply2_2 
      Height          =   375
      Left            =   16200
      Picture         =   "Ludo.frx":188E26
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image Ply1_4 
      Height          =   375
      Left            =   16920
      Picture         =   "Ludo.frx":19274F
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image Ply1_3 
      Height          =   375
      Left            =   16560
      Picture         =   "Ludo.frx":1A3F0A
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image Ply1_2 
      Height          =   375
      Left            =   16200
      Picture         =   "Ludo.frx":1B5FC9
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image Ply4 
      Height          =   585
      Index           =   3
      Left            =   7560
      Picture         =   "Ludo.frx":1C297D
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   645
   End
   Begin VB.Image Ply4 
      Height          =   705
      Index           =   2
      Left            =   7560
      Picture         =   "Ludo.frx":1C34A0
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   645
   End
   Begin VB.Image Ply4 
      Height          =   705
      Index           =   1
      Left            =   6720
      Picture         =   "Ludo.frx":1C42D1
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image Ply4 
      Height          =   705
      Index           =   0
      Left            =   6720
      Picture         =   "Ludo.frx":1C5063
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   645
   End
   Begin VB.Image Ply3 
      Height          =   600
      Index           =   2
      Left            =   1560
      Picture         =   "Ludo.frx":1C5ED9
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   600
   End
   Begin VB.Image Ply3 
      Height          =   600
      Index           =   1
      Left            =   2520
      Picture         =   "Ludo.frx":1C6B52
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   600
   End
   Begin VB.Image Ply3 
      Height          =   600
      Index           =   0
      Left            =   1560
      Picture         =   "Ludo.frx":1C77E5
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   600
   End
   Begin VB.Image Ply2 
      Height          =   585
      Index           =   3
      Left            =   1680
      Picture         =   "Ludo.frx":1C83FE
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   525
   End
   Begin VB.Image Ply2 
      Height          =   585
      Index           =   2
      Left            =   1680
      Picture         =   "Ludo.frx":1C8DE0
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   525
   End
   Begin VB.Image Ply2 
      Height          =   585
      Index           =   1
      Left            =   2520
      Picture         =   "Ludo.frx":1C97C2
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   525
   End
   Begin VB.Image Ply2 
      Height          =   585
      Index           =   0
      Left            =   2520
      Picture         =   "Ludo.frx":1CA1A4
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   525
   End
   Begin VB.Image Ply1 
      Height          =   585
      Index           =   3
      Left            =   6720
      Picture         =   "Ludo.frx":1CAB86
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   525
   End
   Begin VB.Image Ply1 
      Height          =   585
      Index           =   2
      Left            =   7680
      Picture         =   "Ludo.frx":1CB716
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   525
   End
   Begin VB.Image Ply1 
      Height          =   585
      Index           =   1
      Left            =   7560
      Picture         =   "Ludo.frx":1CC325
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   525
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   71
      Left            =   4680
      Picture         =   "Ludo.frx":1CCEF1
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   70
      Left            =   4680
      Picture         =   "Ludo.frx":1CD51C
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   69
      Left            =   4680
      Picture         =   "Ludo.frx":1CDB47
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   68
      Left            =   4680
      Picture         =   "Ludo.frx":1CE172
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   67
      Left            =   4680
      Picture         =   "Ludo.frx":1CE79D
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   66
      Left            =   3555
      Picture         =   "Ludo.frx":1CEDC8
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   65
      Left            =   3000
      Picture         =   "Ludo.frx":1CF3F3
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   64
      Left            =   2400
      Picture         =   "Ludo.frx":1CFA1E
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   63
      Left            =   1800
      Picture         =   "Ludo.frx":1D0049
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   62
      Left            =   1320
      Picture         =   "Ludo.frx":1D0674
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   61
      Left            =   4680
      Picture         =   "Ludo.frx":1D0C9F
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   60
      Left            =   4680
      Picture         =   "Ludo.frx":1D12CA
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   59
      Left            =   4680
      Picture         =   "Ludo.frx":1D18F5
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   58
      Left            =   4680
      Picture         =   "Ludo.frx":1D1F20
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   57
      Left            =   4680
      Picture         =   "Ludo.frx":1D254B
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   56
      Left            =   5880
      Picture         =   "Ludo.frx":1D2B76
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   55
      Left            =   6360
      Picture         =   "Ludo.frx":1D31A1
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   54
      Left            =   6960
      Picture         =   "Ludo.frx":1D37CC
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   53
      Left            =   7560
      Picture         =   "Ludo.frx":1D3DF7
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   52
      Left            =   8040
      Picture         =   "Ludo.frx":1D4422
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   51
      Left            =   8760
      Picture         =   "Ludo.frx":1D4A4D
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   50
      Left            =   8760
      Picture         =   "Ludo.frx":1D5078
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   49
      Left            =   8760
      Picture         =   "Ludo.frx":1D56A3
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   48
      Left            =   8160
      Picture         =   "Ludo.frx":1D5CCE
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   47
      Left            =   7560
      Picture         =   "Ludo.frx":1D62F9
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   46
      Left            =   6960
      Picture         =   "Ludo.frx":1D6924
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   45
      Left            =   6360
      Picture         =   "Ludo.frx":1D6F4F
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   44
      Left            =   5760
      Picture         =   "Ludo.frx":1D757A
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   43
      Left            =   5280
      Picture         =   "Ludo.frx":1D7BA5
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   42
      Left            =   5280
      Picture         =   "Ludo.frx":1D81D0
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   41
      Left            =   5280
      Picture         =   "Ludo.frx":1D87FB
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   40
      Left            =   5280
      Picture         =   "Ludo.frx":1D8E26
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   39
      Left            =   5280
      Picture         =   "Ludo.frx":1D9451
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   38
      Left            =   5280
      Picture         =   "Ludo.frx":1D9A7C
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   37
      Left            =   4680
      Picture         =   "Ludo.frx":1DA0A7
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   30
      Left            =   3600
      Picture         =   "Ludo.frx":1DA6D2
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   29
      Left            =   3000
      Picture         =   "Ludo.frx":1DACFD
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   28
      Left            =   2400
      Picture         =   "Ludo.frx":1DB328
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   27
      Left            =   1800
      Picture         =   "Ludo.frx":1DB953
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   26
      Left            =   1320
      Picture         =   "Ludo.frx":1DBF7E
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   25
      Left            =   720
      Picture         =   "Ludo.frx":1DC5A9
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   24
      Left            =   720
      Picture         =   "Ludo.frx":1DCBD4
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   23
      Left            =   720
      Picture         =   "Ludo.frx":1DD1FF
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   22
      Left            =   1320
      Picture         =   "Ludo.frx":1DD82A
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   21
      Left            =   1800
      Picture         =   "Ludo.frx":1DDE55
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   20
      Left            =   2400
      Picture         =   "Ludo.frx":1DE480
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   19
      Left            =   3000
      Picture         =   "Ludo.frx":1DEAAB
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   18
      Left            =   3600
      Picture         =   "Ludo.frx":1DF0D6
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   17
      Left            =   4080
      Picture         =   "Ludo.frx":1DF701
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   16
      Left            =   4080
      Picture         =   "Ludo.frx":1DFD2C
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   15
      Left            =   4080
      Picture         =   "Ludo.frx":1E0357
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   14
      Left            =   4080
      Picture         =   "Ludo.frx":1E0982
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   13
      Left            =   4080
      Picture         =   "Ludo.frx":1E0FAD
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   12
      Left            =   4080
      Picture         =   "Ludo.frx":1E15D8
      Stretch         =   -1  'True
      Top             =   8760
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   11
      Left            =   4680
      Picture         =   "Ludo.frx":1E1C03
      Stretch         =   -1  'True
      Top             =   8760
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   10
      Left            =   5280
      Picture         =   "Ludo.frx":1E222E
      Stretch         =   -1  'True
      Top             =   8760
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   9
      Left            =   5280
      Picture         =   "Ludo.frx":1E2859
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   8
      Left            =   5280
      Picture         =   "Ludo.frx":1E2E84
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   7
      Left            =   5280
      Picture         =   "Ludo.frx":1E34AF
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   6
      Left            =   5280
      Picture         =   "Ludo.frx":1E3ADA
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   5
      Left            =   5280
      Picture         =   "Ludo.frx":1E4105
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   4
      Left            =   5760
      Picture         =   "Ludo.frx":1E4730
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   3
      Left            =   6360
      Picture         =   "Ludo.frx":1E4D5B
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   2
      Left            =   6960
      Picture         =   "Ludo.frx":1E5386
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   1
      Left            =   7560
      Picture         =   "Ludo.frx":1E59B1
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   0
      Left            =   8160
      Picture         =   "Ludo.frx":1E5FDC
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image Ply1 
      Height          =   585
      Index           =   0
      Left            =   6720
      Picture         =   "Ludo.frx":1E6607
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   525
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   36
      Left            =   4080
      Picture         =   "Ludo.frx":1E7251
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   35
      Left            =   4080
      Picture         =   "Ludo.frx":1E787C
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   34
      Left            =   4080
      Picture         =   "Ludo.frx":1E7EA7
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   33
      Left            =   4080
      Picture         =   "Ludo.frx":1E84D2
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   32
      Left            =   4080
      Picture         =   "Ludo.frx":1E8AFD
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image BG 
      Height          =   375
      Index           =   31
      Left            =   4080
      Picture         =   "Ludo.frx":1E9128
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
End
Attribute VB_Name = "Board"
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








Private Rand As Integer, RandNx As Integer, Rand1 As Integer, Rand2 As Integer, DiceRowed As Boolean, Killed As Boolean
Private Indexx As Integer, Last As Integer, Running As Boolean, Limit As Integer, NxLimit As Integer, LtLimit As Integer
Private Seen As Boolean, overLap As Boolean, Identity As String, Nam As String, Exceed As Boolean
Private Clik As Integer, Turn As String, Nclik As Integer, NumOut As Integer, One As Boolean, Entering As Integer, YesNo As Boolean
Private Temp As Integer, HomeNum As Integer, Finished As Integer, FstDie As Boolean, SndDie As Boolean
Private Out As Boolean, LastStep As Boolean, FreeThrow As Boolean, Start As Integer, Counter As Integer, ComOne As Boolean
Private CheckedFrontDoor As Boolean, CheckedDoor As Boolean, OneDie As Boolean, BothDie As Boolean, BringOut As Boolean
Private KillWith2die As Boolean, StillPlaying As Boolean, Seed1 As Integer, Seed2 As Integer, Slider As Boolean
Private killWith1die As Boolean, OnlyOneCanKill As Boolean, MoveNum As Integer, From6Die As Boolean, SixDie As Boolean
Private MustTransfer As Boolean, FirstToFinish As String, Stored As Boolean, NumMove As Integer
Private DiceRolling As Boolean, ClickIsFromComputer As Boolean, Msgbx As Integer
Private Hunter As String, Hunted As String
Private Vibrating As Boolean

Public Pone As Boolean, Ptwo As Boolean, Pthree As Boolean, Pfour As Boolean
Public Sub Initialise()
Rand = 0
RandNx = 0
Rand1 = 0
Rand2 = 0
DiceRowed = False
Killed = False
Indexx = 0  'it has the same value as Index in tmrmoved
Last = 0 'target position of a moving seed
Running = False
Limit = 0
NxLimit = 0
LtLimit = 0
Seen = False
overLap = False
Identity = ""
Nam = ""
Exceed = False
Clik = 0
Turn = ""
nclick = 0
NumOut = 0
One = False
Entering = 0
YesNo = False
Temp = 0
HomeNum = 0
Finished = 0
FstDie = False
SndDie = False
Out = False
LastStep = False
FreeThrow = False
Start = 0 'the door of the present player,in this case player1 but anyway this value is changed once it's a specific player's turn
Counter = 0
ComOne = False
CheckedFrontDoor = False
CheckedDoor = False
OneDie = False
BothDie = False
BringOut = False
KillWith2die = False
StillPlaying = False
Seed1 = 0
Seed2 = 0
Slider = False
killWith1die = False
OnlyOneCanKill = False
MoveNum = 0
From6Die = False
SixDie = False
MustTransfer = False
FirstToFinish = ""
Stored = False
NumMove = 0
DiceRolling = False
ClickIsFromComputer = False
Hunter = ""
Hunted = ""
Pone = False
Ptwo = False
Pthree = False
Pfour = False
Loozer = ""
loozernam = ""
Vibrating = False
Ply1Hom = 0
Ply2Hom = 0
Ply3Hom = 0
Ply4Hom = 0


End Sub



Private Sub cmdAbout_Click()
About.Show

End Sub

Private Sub cmdAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider = True
End Sub

'bg(index).Tag contains player's identity i.e p1,p2,p3 or p4
'bgnum(index) contains number of seed on a location

Private Sub BG_Click(Index As Integer)

tmrToolSlider.Interval = 1: Slider = False
If Index = 111 Then MsgBox "Bg_click has an indexx of 111,comone=" & ComOne: Exit Sub
Timer3.Enabled = True
'bg_click takes care of d initial positn,Checkers takes care of the target positn after the simulation(movement)
If DiceRowed = False Then Label2.Caption = "Please Row The Dice": Exit Sub
If bgNum(Index) = 0 Then Label2.Caption = "Click On A Seed": Exit Sub 'user clicked on an empty space,he has to click on a player to move
 If Running = True Then Label2.Caption = "Please Wait...": Exit Sub
If Turn <> BG(Index).Tag Then Label2.Caption = "Not Turn Yet": Exit Sub


 If PlayerType = 11 Then
     If Not ClickIsFromComputer Then
        Label2.Caption = "That's Computer's Seed, Forbear!"
        Exit Sub
     End If
 End If

    '*********************************************
         'if we have say 3 seeds remaining for a player and only one seed is on the field _
         if the player has a score of say 6,3 and this only seed out can kill with a 6 _
         but what happens to the 3, the program will be stuck, to avoid this if the above _
         conditions are met then move a 3 first b4 moving the 6 _
         i suppose the player to bring out a seed but they might act otherwise ,don't want to leave a chance
         Dim Kanter As Integer, Kant As Integer, Temp As Integer
         For Kant = 0 To 71
             If BG(Kant).Tag = Turn Then Kanter = Kanter + 1
         Next
          If Kanter = 1 And Nclik = 0 And One = False Then 'if one is true rand must have been sumed if the sum is 6 there is a problem,this problem lingered for a very long time
             If Rand = 6 Then
                Rand = RandNx
                RandNx = 6
             End If
          End If
    '*****************************************************
         
         
         StillPlaying = True 'will be made false when the 2nd seed has finished  moving

       Indexx = Index
        Last = Rand + Indexx   'target position of the moving seed
        Identity = Trim(BG(Index).Tag)

Select Case Turn  'this will be used in the timer event.the reason y we can't use bg().tag is that indexx is not stable in the said event
       Case "P1": Limit = 51: NxLimit = 52: LtLimit = 56: Nam = "P1"
       Case "P2": Limit = 12: NxLimit = 57: LtLimit = 61: Nam = "P2"
       Case "P3": Limit = 25: NxLimit = 62: LtLimit = 66: Nam = "P3"
       Case "P4": Limit = 38: NxLimit = 67: LtLimit = 71: Nam = "P4"
End Select
  'If Last = Limit Then Last = NxLimit
      'CheckPlayers
      Distance Index
       If Exceed = True Then Label2.Caption = "Cannot Move This Seed": Exit Sub
       Running = True
       Nclik = Nclik + 1
  
If One = True Or Nclik = 2 Then
  For H = 0 To 3
    Line1(H).Visible = True
  Next
End If
  If Rand = Rand2 Then
          Line1(2).Visible = True
         Line1(3).Visible = True
 Else
         Line1(0).Visible = True
         Line1(1).Visible = True
End If


Select Case bgNum(Index)  'we have to know how many seeds are on it, at least one player should be on it
  Case 1 'move this seed
       Store(1).Picture = BG(Index).Picture   'stores the moving pic seed
        BG(Index).Tag = ""             'now it does not have any player's seed on it
        bgNum(Index) = 0      'the last num of seed on it was one now it 's zero  -last positn
       overLap = False       'no seed remains
Case 2  ' it contains 2 seed of d  same player,move one and retain one
   Select Case BG(Index).Tag   'we have to know which player has this seeds
          Case "P1"      'the moving seed has to be only one
               Store(1).Picture = Ply1_1.Picture
               BG(Index).Picture = Ply1_1.Picture
          Case "P2"
               Store(1).Picture = Ply2_1.Picture
               BG(Index).Picture = Ply2_1.Picture
          Case "P3"
               Store(1).Picture = Ply3_1.Picture
               BG(Index).Picture = Ply3_1.Picture
          Case "P4"
               Store(1).Picture = Ply4_1.Picture
               BG(Index).Picture = Ply4_1.Picture
  End Select
    
      
    
  overLap = True
  bgNum(Index) = 1      'the last num of seed on it was two now it's one -last positn

Case 3 ' it contains 3 seed,move one and retain two
   Select Case BG(Index).Tag   'we have to know which player has this seeds
          Case "P1"      'the moving seed has to be only one
               Store(1).Picture = Ply1_1.Picture
               BG(Index).Picture = Ply1_2.Picture
          Case "P2"
               Store(1).Picture = Ply2_1.Picture
               BG(Index).Picture = Ply2_2.Picture
          Case "P3"
               Store(1).Picture = Ply3_1.Picture
               BG(Index).Picture = Ply3_2.Picture
          Case "P4"
               Store(1).Picture = Ply4_1.Picture
               BG(Index).Picture = Ply4_2.Picture
  End Select
    
  bgNum(Index) = 2      'the last num of seed on it was three now it's two -last positn
  overLap = True
Case 4 ' it contains 4 seed,move one and retain three
   Select Case BG(Index).Tag   'we have to know which player has this seeds
          Case "P1"      'the moving seed has to be only one
               Store(1).Picture = Ply1_1.Picture
               BG(Index).Picture = Ply1_3.Picture
          Case "P2"
               Store(1).Picture = Ply2_1.Picture
               BG(Index).Picture = Ply2_3.Picture
          Case "P3"
               Store(1).Picture = Ply3_1.Picture
               BG(Index).Picture = Ply3_3.Picture
          Case "P4"
               Store(1).Picture = Ply4_1.Picture
                BG(Index).Picture = Ply4_3.Picture
          Case Else: Beep
  End Select
    
  bgNum(Index) = 3      'the last num of seed on it was four now it's three -last positn
  overLap = True
End Select
  tmrMoved.Enabled = True
      


End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tmrToolSlider.Interval = 1: Slider = False
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command7_Click(Index As Integer)

End Sub


Private Sub Form_Activate()
wmpPopup.URL = Dirxtry & "pop.ogg"
End Sub

Private Sub Form_Initialize()
'Load Wins
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then PauseStart_Click
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 112 'F1
    Help.Value = True
Case 113 'F2
     cmdAbout.Value = True
 Case 19  'pause/break key
     PauseStart.Value = True
 Case 35, 27 'end ,esc key
     Quit.Value = True
 Case 37, 39
 If MustTrans.Enabled Then
     MustTrans.Value = True
 End If
End Select
End Sub

Private Sub Help_Click()
LudoHelp.Show
End Sub

Private Sub Help_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider = True

End Sub

Private Sub MustTrans_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider = True
End Sub

Private Sub Quit_Click()
'the .EXE file i created was behaving strange if a msgbox comes up the program continues although modal,
'i expect all other computation to stop( like all timer event),don't know what caused it as it does not behave so if i run the source code
'since this is crucial to the smooth running of the program decided to manually pause the program
Dim blnInitialState As Boolean 'store current state of the game if it is either paused or unpaused
blnInitialState = Paused
If Not Paused Then 'if the game is not already paused pause it, this is essential,cos when d form loads it is paused
PauseStart.Value = True 'pause game
End If

If DiceRolling Then   'if the dice is rolling and the user returns(quit) back to the setting when a new game is loaded complicated errors occur which i'm not ready to fix the same error occur when the user clicks on the transfer slate button which i had to control by either enabling or disabling it as required, or should i just disable it(the quit button) under this condition too? hmmm...

    If MsgBox("Cannot Quit Now,The Dice Is Rolling. " + vbCrLf + _
        "If You Must Quit, The Program Will End Abruptly!" + vbCrLf + _
        vbCrLf + "Sure To Quit?", vbCritical + vbOKCancel + vbDefaultButton2) = vbOK Then
   
        End
    End If
           'user clicked cancel
If Paused Then 'unpause---if there is any word like that
PauseStart.Value = True 'resume
End If
PauseStart.Value = blnInitialState
    Exit Sub
End If

If ArrayOfWinners(0) <> "" Then
If MsgBox("The Current Game Is Yet To Be Finalised, If You Quit Now, " + vbCrLf _
  + ArrayOfWinners(0) + " Will Be " + "Awarded As The Champion." _
  + vbCrLf + vbCrLf _
  + "Sure To Quit?", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
     'DiceRowed = False
     tmrDice.Enabled = False
     Paused = True
     Wins.ShowWinner
     Wins.Label4.Caption = ArrayOfWinners(0)
     Wins.Show vbModal
  End If
  
Else   'no player has finished first
  If MsgBox("The Current Game Is Yet To Be Finalised, If You Quit Now, No Player Will Be " + vbCrLf _
          + "Awarded As The Champion And The Present Game May Not Be Recovered. " _
    + vbCrLf + vbCrLf _
    + "Sure To Quit?", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
     'DiceRowed = False
     tmrDice = False
     Paused = True
     Initialise  'initialise all module variables,despite the fact that board will be unloaded it remembers the valu of all module level variables
     Unload Me
     Load Setn
     Setn.Show
  End If
End If
If Paused Then 'unpause---if there is any word like that
PauseStart.Value = True 'resume
End If
PauseStart.Value = blnInitialState 'return to original state either paused or unpaused

End Sub





Private Sub Command3_Click()
If Timer2.Enabled = True Then Timer2.Enabled = False Else Timer2.Enabled = True
End Sub

Private Sub Command9_Click()

End Sub

Private Sub Command6_Click()

End Sub

Private Sub MustTrans_Click()
'the .EXE file i created was behaving strange if a msgbox comes up the program continues although modal,
'i except all other computation to stop( like all timer event),don't know what caused it as it does not behave so if i run the source code
'since this is crucial to the smooth running of the program decided to manually pause the program
Dim blnInitialState As Boolean  ' return to it pause state after this guy is clicked
blnInitialState = Paused
If Not Paused Then 'if the game is not already paused pause it, this is essential,cos when d form loads it is paused
PauseStart.Value = True 'pause game
End If
If MsgBox("The Current Player May Loose Turn To Play As" + vbCrLf + _
          "Dice Slate Will Be Transfered To The Next Player." + vbCrLf + vbCrLf + _
          "Sure To Transfer?", vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
MustTransfer = True
TransSlate.Enabled = True
End If
If Paused Then 'unpause---if there is any word like that
PauseStart.Value = True 'resume
End If
PauseStart.Value = blnInitialState

End Sub

Private Sub Die_Click(Index As Integer)
'Exit Sub
If Not DiceRowed Then Exit Sub 'the dice has not been rolled
If One Then Exit Sub 'it must have sumed up the dice for only one player
If Nclik = 1 Then Exit Sub 'user has already clicked once,so s/he has no choice than to use the remaining one
    
wmpPopup.URL = Dirxtry & "start.wav"
If Index = 1 Then
   Rand = Rand1
   RandNx = Rand2
Else
   Rand = Rand2
   RandNx = Rand1
End If

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tmrToolSlider.Interval = 1: Slider = False
If Paused Then
PauseStart.Picture = Starts(0).Picture
Else
PauseStart.Picture = Pause(0).Picture
End If
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tmrToolSlider.Interval = 1: Slider = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 Slider = True  'if mouse is on it always keep it open

If Paused Then
PauseStart.Picture = Starts(0).Picture
Else
PauseStart.Picture = Pause(0).Picture
End If

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tmrToolSlider.Interval = 1: Slider = False
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tmrToolSlider.Interval = 1: Slider = False
End Sub

Private Sub Label4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
tmrToolSlider.Interval = 1: Slider = False
End Sub

Private Sub Label6_Click()
TransSlate.Enabled = True
Beep
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.Caption = One & "  xcd" & Exceed & "  trn" & Turn
End Sub


Private Sub Label7_Click()
Beep
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tmrToolSlider.Interval = 1
Slider = False
End Sub

Private Sub PauseStart_Click()

Paused = Not Paused
If Paused Then
    PauseStart.Picture = Starts(0).Picture
    PauseStart.ToolTipText = "Resume Game"
    Label3.Caption = "Game Paused"
    Label2.Caption = "Press Enter To Resume"
    Label3.ForeColor = vbRed
    Label2.ForeColor = vbRed
    Frame1.Enabled = False
    Label7.Visible = True
    MustTrans.Enabled = True
Else
    PauseStart.Picture = Pause(0).Picture
    PauseStart.ToolTipText = "Pause Game"
    Label3.Caption = "Game Resumed"
    Label3.ForeColor = Label5.BackColor
    Label2.ForeColor = Label5.BackColor
    Label2.Caption = "Continue..."
    Frame1.Enabled = True
    Label7.Visible = False
End If

End Sub
Sub P_S()
End Sub

Private Sub PauseStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Slider = True  'if mouse is on it always keep it open

If Paused Then
PauseStart.Picture = Starts(1).Picture
Else
PauseStart.Picture = Pause(1).Picture
End If

End Sub

Private Sub Ply2_Click(Index As Integer)
tmrToolSlider.Interval = 1: Slider = False
If DiceRowed = True Then 'dice must have been rowed
      If Turn <> "P2" Then Label2.Caption = "Not Turn Yet, Click On the Right Seed": Exit Sub
If PlayerType = 11 Then
     If Not ClickIsFromComputer Then
        Label2.Caption = "That's Computer's Seed, Forbear!"
        Exit Sub
     End If
 End If
 If PlayerType <> Ptyp2 Then Beep: Exit Sub 'the same as the above if then
 
        If Rand2 = 6 And Nclik <> 1 Then  'so that it will automatically use the die with a 6 without the user specifying ,this should happen only at first clik otherwise if the user cliks d 2nd time he'll get another 6
        Rand = 6
        RandNx = Rand1
        End If
   
  If One = True And Rand = 6 Then Label2.Caption = "Can't Move This Seed With A Throw of" & Str(Rand1) & Str(Rand2): Exit Sub
     If Ply2(Index).Tag = 1 Then 'already out
          Beep 'bg_click will take care of this
     Else    'not yet out

        If Rand = 6 Then   'can only come out with a throw of six
         wmpFreedSound.URL = Dirxtry & "button1.ogg"
        StillPlaying = True 'will be made false when the 2nd seed has finished has moving
         Nclik = Nclik + 1
             If Rand1 = 6 Then
           Line1(0).Visible = True
           Line1(1).Visible = True
           ElseIf Rand2 = 6 Then
           Line1(2).Visible = True
           Line1(3).Visible = True
           End If
     
                 If Nclik = 2 Then
                    For H = 0 To 3
                        Line1(H).Visible = True
                    Next
                 End If
                 
             If FreeThrow Then GoTo 9 'for two 6, it should not automatically move the seed
      
           
          Out = False
         For J = 0 To 71  'let's see if any seed is out at all,if a player gets a throw of say a 6 and a 4 and another player's seed is at it 'door step' it kills with a 6 but what happens with the 4 ? this code solves it,
         If BG(J).Tag = "P2" Then Out = True  ' (contn) by not allowing it to settle at it's step if this is the case
         Next
         If Out = False Then 'no seed is out
         YesNo = False
         HomeNum = 13
           If bgNum(HomeNum) <> 0 Then 'let'see if there is any seed there, store it b4 d simulation
              Motion.Picture = BG(HomeNum).Picture
              Motion.Tag = BG(HomeNum).Tag
              Temp = bgNum(HomeNum)
              YesNo = True
           End If
           
          BG(HomeNum).Picture = Ply2_1.Picture
        Ply2(Index).Tag = 1 'Now it is out\
          Ply2(Index).Visible = False   'remove from prison
        BG(HomeNum).Tag = "P2"  'put player1's seed on it
        bgNum(HomeNum) = 1 'one player is on it but on motion
        Nclik = 2
        Rand = RandNx  'already out with 6 use the next die
        BG_Click HomeNum  'move it
        DiceRowed = False
        Exit Sub
        End If
                         Hunter = Pnam2
        
9         Ply2(Index).Tag = 1 'Now it is out
         
            Select Case BG(13).Tag  'to see if a player's seed has already occupied this positn
           
               Case "P1"   'already occupied by player1
                         Hunted = Pnam1
                         tmrVibrate.Enabled = True
                         Select Case bgNum(13)   'how many of player1 seed are on it
                         
                         Case 1 ' only one, Captured  by player2
                    
                             BG(13).Picture = LoadPicture("") 'no seed will appear  on it
                             bgNum(13) = 0                     'num  of seeds is zero
                             BG(13).Tag = ""   'no player's seed is on it
                         
                         Case 2 ' only two, one is Captured  by player1
                         
                         BG(13).Picture = Ply1_1.Picture ' one has been killed it remains one
                         bgNum(13) = 1
                         'bg(13).tag is stil  occupied by "p1"
            
                          Case 3 ' only three, one is  Captured  by player2
                         
                         BG(13).Picture = Ply1_2.Picture ' one has been killed it remains two
                         bgNum(13) = 2
                          'bg(13).tag is stil  occupied by "p1"
                         
                          Case 4 ' only four, one is  Captured  by player1
                         
                         BG(13).Picture = Ply1_3.Picture ' one has been killed it remains three
                         bgNum(13) = 3
                          'bg(13).tag is stil  occupied by "p1"
                   End Select
                   
                       
                         For J = 0 To 3 'to return player1 back to it's prison
                         If Ply1(J).Visible = False Then Exit For 'to find an empty imagebox to put it in
                         Next
                        ' Ply1(j).Picture = Ply1_1.Picture
                         Ply1(J).Visible = True
                         Ply1(J).Tag = 0  'not yet out
                        
                         Ply2(Index).Visible = False
                         Ply2(Index).Tag = 3 'it goes home for capturing an enemy
                         'Home2
                    'f = MsgBox(Pnam2 & " Has Captured " & Pnam1 & "'s seed", vbInformation)
              Case "P2"  'already occupied by player2
                         
                    Ply2(Index).Visible = False   'remove seed from prison
                Select Case bgNum(13) 'check how many of player2 seed is on it
                  Case 1      'increment the num of players & also the pic
                  bgNum(13) = 2
                  BG(13).Picture = Ply2_2.Picture
                  Case 2
                  bgNum(13) = 3
                  BG(13).Picture = Ply2_3.Picture
                  Case 3
                  bgNum(13) = 4
                  BG(13).Picture = Ply2_4.Picture
                  Case Else: MsgBox "impossible error occured" 'impossible or error
                  End Select

                         
                  
                   
            Case "P3"  'already occupied by player3
                         Hunted = Pnam3
                         tmrVibrate.Enabled = True
                     Select Case bgNum(13)   'how many of player3 seed are on it
                         
                         Case 1 ' only one, Captured  by player1
                       
                         BG(13).Picture = LoadPicture("") 'no seed will appear  on it
                         bgNum(13) = 0                     'num  of seeds is zero
                         BG(13).Tag = ""   'no player's seed is on it
                        
                         Case 2 ' only two, one is  Captured  by player1
                        
                         BG(13).Picture = Ply3_1.Picture ' one has been killed it remains one
                         bgNum(13) = 1
                
                          Case 3 ' only three, one is  Captured  by player1
                         
                         BG(13).Picture = Ply3_2.Picture ' one has been killed it remains two
                         bgNum(13) = 2
                        
                          Case 4 ' only four, one is  Captured  by player1
                         
                         BG(13).Picture = Ply3_3.Picture ' one has been killed it remains three
                         bgNum(13) = 3
                        
                 End Select
                
                         
                         For J = 0 To 3 'to return player3 back to it's prison
                         If Ply3(J).Visible = False Then Exit For 'to find an empty imagebox to put it in
                         Next
                         'Ply3(j).Picture = Ply3_1.Picture
                         Ply3(J).Visible = True
                         Ply3(J).Tag = 0  'not yet out
                         
                         Ply2(Index).Visible = False
                         Ply2(Index).Tag = 3 'it goes home for capturing an enemy
                         'Home2
                          'f = MsgBox(Pnam2 & " Has Captured " & Pnam3 & "'s seed", vbInformation)
            Case "P4"        'already occupied by " &  pnam4  & "
                  Hunted = Pnam4
                  tmrVibrate.Enabled = True
                         
                     Select Case bgNum(13)   'how many of player4 seed are on it
                         
                         Case 1 ' only one, Captured  by player2
                       
                         BG(13).Picture = LoadPicture("") 'no seed will appear  on it
                         bgNum(13) = 0                     'num  of seeds is zero
                         BG(13).Tag = ""   'no player's seed is on it
                        
                         Case 2 ' only two, one is  Captured  by player2
                        
                         BG(13).Picture = Ply4_1.Picture ' one has been killed it remains one
                         bgNum(13) = 1
                
                          Case 3 ' only three, one is  Captured  by player2
                         
                         BG(13).Picture = Ply4_2.Picture ' one has been killed it remains two
                         bgNum(13) = 2
                        
                          Case 4 ' ony four, one is  Captured  by player2
                         
                         BG(13).Picture = Ply4_3.Picture ' one has been killed it remains three
                         bgNum(13) = 3
                        
                 End Select
                   
                         For J = 0 To 3 'to return player4 back to it's prison
                         If Ply4(J).Visible = False Then Exit For 'to find an empty imagebox to put it in
                         Next
                       '  Ply4(j).Picture = Ply4_1.Picture
                         Ply4(J).Visible = True
                         Ply4(J).Tag = 0  'not yet out
                         
                         Ply2(Index).Visible = False
                         Ply2(Index).Tag = 3 'it goes home for capturing an enemy
                         'Home2
              F = MsgBox(Pnam2 & " Has Captured " & Pnam4 & "'s seed", vbInformation)
            Case Else   'no player's seed is on it
            
            BG(13).Tag = "P2"  'put player2's seed on it
            BG(13).Picture = Ply2_1.Picture
            bgNum(13) = 1 'one player is on it
            Ply2(Index).Visible = False   'remove from prison
            
            End Select
          Else 'rand <> 6
         Label2.Caption = "Can't Move This Seed With A Throw of" & Str(Rand)
         Exit Sub
        End If
      End If
    If Nclik = 1 Then
        Rand = RandNx
        DiceRowed = True
    Else
        StillPlaying = False
        TransSlate.Enabled = True 'transfer dice slate to the next player
        DiceRowed = False
        Nclik = 0
        Rand1 = 0
        Rand2 = 0
  End If

Else   'user has not rowed d dice
Label2.Caption = "Please Row The Dice"
End If
End Sub

Private Sub Ply3_Click(Index As Integer)
tmrToolSlider.Interval = 1: Slider = False
If DiceRowed = True Then 'dice must have been rowed
    If Turn <> "P3" Then Label2.Caption = "Not Turn Yet, Click On the Right Seed": Exit Sub
          If PlayerType = 11 Then
     If Not ClickIsFromComputer Then
        Label2.Caption = "That's Computer's Seed, Forbear!"
        Exit Sub
     End If
 End If
    
              If PlayerType <> Ptyp3 Then Beep: Exit Sub

        If Rand2 = 6 And Nclik <> 1 Then  'so that it will automatically use the die with a 6 without the user specifying ,this should happen only at first clik otherwise if the user cliks d 2nd time he'll get another 6
        Rand = 6
        RandNx = Rand1
        End If
  If One = True And Rand = 6 Then Label2.Caption = "Can't Move This Seed With A Throw of" & Str(Rand1) & Str(Rand2): Exit Sub

     If Ply3(Index).Tag = 1 Then 'already out
          Beep 'bg_click will take care of this
          MsgBox "the chosen seed is already out"
     Else    'not yet out
        
        If Rand = 6 Then   'can only come out with a throw of six
         wmpFreedSound.URL = Dirxtry & "button1.ogg"
            StillPlaying = True 'will be made false when the 2nd seed has finished has moving

         Nclik = Nclik + 1
                If Rand1 = 6 Then
           Line1(0).Visible = True
           Line1(1).Visible = True
           ElseIf Rand2 = 6 Then
           Line1(2).Visible = True
           Line1(3).Visible = True
           End If

         If Nclik = 2 Then
            For H = 0 To 3
                Line1(H).Visible = True
            Next
        End If
             If FreeThrow Then GoTo 9 'for two 6, it should not automatically move the seed
          
 Out = False
         For J = 0 To 71  'let's see if any seed is out at all,if a player gets a throw of say a 6 and a 4 and another player's seed is at it 'door step' it kills with a 6 but what happens with the 4 ? this code solves it,
         If BG(J).Tag = "P3" Then Out = True  ' (contn) by not allowing it to settle at it's step if this is the case
         Next
         If Out = False Then 'no seed is out
         YesNo = False
         HomeNum = 26
           If bgNum(HomeNum) <> 0 Then 'let'see if there is any seed there, store it b4 d simulation
              Motion.Picture = BG(HomeNum).Picture
              Motion.Tag = BG(HomeNum).Tag
              Temp = bgNum(HomeNum)
              YesNo = True
           End If
           
          BG(HomeNum).Picture = Ply3_1.Picture
        Ply3(Index).Tag = 1 'Now it is out\
          Ply3(Index).Visible = False   'remove from prison
        BG(HomeNum).Tag = "P3"  'put player1's seed on it
        bgNum(HomeNum) = 1 'one player is on it but on motion
        Nclik = 2
        Rand = RandNx  'already out with 6 use the next die
        BG_Click HomeNum  ' clik on it & move it
        DiceRowed = False
        Exit Sub
        End If
          Hunter = Pnam3
9         Ply3(Index).Tag = 1 'Now it is out
         
            Select Case BG(26).Tag  'to see if a player's seed has already occupied this positn
           
               Case "P1"   'already occupied by player1
                         Hunted = Pnam1
                         tmrVibrate.Enabled = True
                         Select Case bgNum(26)   'how many of player1 seed are on it
                         
                         Case 1 ' only one, Captured  by player3
                    
                             BG(26).Picture = LoadPicture("") 'no seed will appear  on it
                             bgNum(26) = 0                     'num  of seeds is zero
                             BG(26).Tag = ""   'no player's seed is on it
                         
                         Case 2 ' only two, one is  Captured  by player3
                         
                         BG(26).Picture = Ply1_1.Picture ' one has been killed it remains one
                         bgNum(26) = 1
                         'bg(26).tag is stil  occupied by "p1"
            
                          Case 3 ' only three, one is  Captured  by player3
                         
                         BG(26).Picture = Ply1_2.Picture ' one has been killed it remains two
                         bgNum(26) = 2
                          'bg(26).tag is stil  occupied by "p1"
                         
                          Case 4 ' only four, one is  Captured  by player3
                         
                         BG(26).Picture = Ply1_3.Picture ' one has been killed it remains three
                         bgNum(26) = 3
                          'bg(26).tag is stil  occupied by "p1"
                   End Select
                   
                       
                         For J = 0 To 3 'to return player1 back to it's prison
                         If Ply1(J).Visible = False Then Exit For 'to find an empty imagebox to put it in
                         Next
                         'Ply1(j).Picture = Ply1_1.Picture
                         Ply1(J).Visible = True
                         Ply1(J).Tag = 0  'not yet out
                        
                         Ply3(Index).Visible = False
                         Ply3(Index).Tag = 3 'it goes home for capturing an enemy
                         'Home3
                    'f = MsgBox(Pnam3 & " Has Captured " & Pnam1 & "'s seed", vbInformation)
          Case "P2"  'already occupied by player2
                     Hunted = Pnam2
                     tmrVibrate.Enabled = True
                    Select Case bgNum(26)   'how many of player2 seed are on it
                         
                         Case 1 ' only one, Captured  by player1
                       
                         BG(26).Picture = LoadPicture("") 'no seed will appear  on it
                         bgNum(26) = 0                     'num  of seeds is zero
                         BG(26).Tag = ""   'no player's seed is on it
                        
                         Case 2 ' only two, one is  Captured  by player3
                        
                         BG(26).Picture = Ply2_1.Picture ' one has been killed it remains one
                         bgNum(26) = 1
                
                          Case 3 ' only three, one is  Captured  by player3
                         
                         BG(26).Picture = Ply2_2.Picture ' one has been killed it remains two
                         bgNum(26) = 2
                        
                          Case 4 ' only four, one is  Captured  by player3
                         
                         BG(26).Picture = Ply2_3.Picture ' one has been killed it remains three
                         bgNum(26) = 3
                        
                 End Select
                
                         
                         For J = 0 To 3 'to return player3 back to it's prison
                         If Ply2(J).Visible = False Then Exit For 'to find an empty imagebox to put it in
                         Next
                         'Ply2(j).Picture = Ply2_1.Picture
                         Ply2(J).Visible = True
                         Ply2(J).Tag = 0  'not yet out
                         
                         Ply3(Index).Visible = False
                         Ply3(Index).Tag = 3 'it goes home for capturing an enemy
                         'Home3
                       ' f = MsgBox(Pnam3 & " Has Captured " & Pnam2 & "'s seed", vbInformation)

                   
            Case "P3"  'already occupied by player3
                        
                        
                   Ply3(Index).Visible = False   'remove seed from prison
                Select Case bgNum(26) 'check how many of player3 seed is on it
                  Case 1      'increment the num of players & also the pic
                  bgNum(26) = 2
                  BG(26).Picture = Ply3_2.Picture
                  Case 2
                  bgNum(26) = 3
                  BG(26).Picture = Ply3_3.Picture
                  Case 3
                  bgNum(26) = 4
                  BG(26).Picture = Ply3_4.Picture
                  Case Else: MsgBox "impossible error occured" 'impossible or error
                  End Select

             
             Case "P4"        'already occupied by player4
                  
                     Hunted = Pnam4
                     tmrVibrate.Enabled = True
                     Select Case bgNum(26)   'how many of player4 seed are on it
                         
                         Case 1 ' only one, Captured  by player3
                       
                         BG(26).Picture = LoadPicture("") 'no seed will appear  on it
                         bgNum(26) = 0                     'num  of seeds is zero
                         BG(26).Tag = ""   'no player's seed is on it
                        
                         Case 2 ' only two, one is  Captured  by player3
                        
                         BG(26).Picture = Ply4_1.Picture ' one has been killed it remains one
                         bgNum(26) = 1
                
                          Case 3 ' only three, one is  Captured  by player3
                         
                         BG(26).Picture = Ply4_2.Picture ' one has been killed it remains two
                         bgNum(26) = 2
                        
                          Case 4 ' only four, one is  Captured  by player3
                         
                         BG(26).Picture = Ply4_3.Picture ' one has been killed it remains three
                         bgNum(26) = 3
                        
                 End Select
                   
                         For J = 0 To 3 'to return player4 back to it's prison
                         If Ply4(J).Visible = False Then Exit For 'to find an empty imagebox to put it in
                         Next
                         'Ply4(j).Picture = Ply4_1.Picture
                         Ply4(J).Visible = True
                         Ply4(J).Tag = 0  'not yet out
                         
                         Ply3(Index).Visible = False
                         Ply3(Index).Tag = 3 'it goes home for capturing an enemy
                         'Home3
            ' f = MsgBox(Pnam3 & " Has Captured " & Pnam4 & "'s seed", vbInformation)
            Case Else   'no player's seed is on it
            
            BG(26).Tag = "P3"  'put player3's seed on it
            BG(26).Picture = Ply3_1.Picture
            bgNum(26) = 1 'one player is on it
            Ply3(Index).Visible = False   'remove from prison
            
            End Select
          'Label2.Caption = BG(26).Tag: Beep
          Else 'rand <> 6
         Label2.Caption = "Can't Move This Seed With A Throw of" & Str(Rand)
         Exit Sub
        End If
      End If
    If Nclik = 1 Then
        Rand = RandNx
        DiceRowed = True
    Else
        StillPlaying = False
        TransSlate.Enabled = True 'transfer dice slate to the next player
        DiceRowed = False
        Nclik = 0
        Rand1 = 0
        Rand2 = 0
  End If

Else   'user has not rowed d dice
Label2.Caption = "Please Row The Dice"
End If

End Sub

Private Sub Ply4_Click(Index As Integer)
tmrToolSlider.Interval = 1: Slider = False
If DiceRowed = True Then 'dice must have been rowed
    If Turn <> "P4" Then Label2.Caption = "Not Turn Yet, Click On the Right Seed": Exit Sub
If PlayerType = 11 Then
     If Not ClickIsFromComputer Then
        Label2.Caption = "That's Computer's Seed, Forbear!"
        Exit Sub
     End If
 End If
              
              If PlayerType <> Ptyp4 Then Beep: Exit Sub

        If Rand2 = 6 And Nclik <> 1 Then  'so that it will automatically use the die with a 6 without the user specifying ,this should happen only at first clik otherwise if the user cliks d 2nd time he'll get another 6
        Rand = 6
        RandNx = Rand1
        End If
  If One = True And Rand = 6 Then Label2.Caption = "Can't Move This Seed With A Throw of" & Str(Rand1) & Str(Rand2): Exit Sub

     If Ply4(Index).Tag = 1 Then 'already out
          Beep 'bg_click will take care of this
     Else    'not yet out
        
        If Rand = 6 Then   'can only come out with a throw of six
         wmpFreedSound.URL = Dirxtry & "button1.ogg"

         StillPlaying = True 'will be made false when the 2nd seed has finished has moving

         Nclik = Nclik + 1
         
         If Rand1 = 6 Then
           Line1(0).Visible = True
           Line1(1).Visible = True
           ElseIf Rand2 = 6 Then
           Line1(2).Visible = True
           Line1(3).Visible = True
           End If

         If Nclik = 2 Then
            For H = 0 To 3
                Line1(H).Visible = True
            Next
        End If
                      If FreeThrow Then GoTo 9 'for two 6, it should not automatically move the seed
           
 Out = False
         For J = 0 To 71  'let's see if any seed is out at all,if a player gets a throw of say a 6 and a 4 and another player's seed is at it 'door step' it kills with a 6 but what happens with the 4 ? this code solves it,
         If BG(J).Tag = "P4" Then Out = True  ' (contn) by not allowing it to settle at it's step if this is the case
         Next
         If Out = False Then 'no seed is out
         YesNo = False
         HomeNum = 39
           If bgNum(HomeNum) <> 0 Then 'let'see if there is any seed there, store it b4 d simulation
              Motion.Picture = BG(HomeNum).Picture
              Motion.Tag = BG(HomeNum).Tag
              Temp = bgNum(HomeNum)
              YesNo = True
           End If
           
          BG(HomeNum).Picture = Ply4_1.Picture
        Ply4(Index).Tag = 1 'Now it is out\
          Ply4(Index).Visible = False   'remove from prison
        BG(HomeNum).Tag = "P4"  'put player1's seed on it
        bgNum(HomeNum) = 1 'one player is on it but on motion
        Nclik = 2
        Rand = RandNx  'already out with 6 use the next die
        BG_Click HomeNum  'move it
        DiceRowed = False
        Exit Sub
        End If
           Hunter = Pnam4
9         Ply4(Index).Tag = 1 'Now it is out
         
            Select Case BG(39).Tag  'to see if a player's seed has already occupied this positn
           
               Case "P1"   'already occupied by player1
                         
                     Hunted = Pnam1
                     tmrVibrate.Enabled = True
                         Select Case bgNum(39)   'how many of player1 seed are on it
                         
                         Case 1 ' only one, Captured  by player4
                    
                             BG(39).Picture = LoadPicture("") 'no seed will appear  on it
                             bgNum(39) = 0                     'num  of seeds is zero
                             BG(39).Tag = ""   'no player's seed is on it
                         
                         Case 2 ' only two, one is  Captured  by player4
                         
                         BG(39).Picture = Ply1_1.Picture ' one has been killed it remains one
                         bgNum(39) = 1
                         'bg(39).tag is stil  occupied by "p1"
            
                          Case 3 ' only three, one is  Captured  by player4
                         
                         BG(39).Picture = Ply1_2.Picture ' one has been killed it remains two
                         bgNum(39) = 2
                          'bg(39).tag is stil  occupied by "p1"
                         
                          Case 4 ' only four, one is  Captured  by player4
                         
                         BG(39).Picture = Ply1_3.Picture ' one has been killed it remains three
                         bgNum(39) = 3
                          'bg(39).tag is stil  occupied by "p1"
                   End Select
                   
                       
                         For J = 0 To 3 'to return player1 back to it's prison
                         If Ply1(J).Visible = False Then Exit For 'to find an empty imagebox to put it in
                         Next
                        ' Ply1(j).Picture = Ply1_1.Picture
                         Ply1(J).Visible = True
                         Ply1(J).Tag = 0  'not yet out
                        
                         Ply4(Index).Visible = False
                         Ply4(Index).Tag = 3 'it goes home for capturing an enemy
                         'Home4
                   'f = MsgBox(Pnam4 & " Has Captured Player1's seed", vbInformation)
              Case "P2"  'already occupied by player2
                         
                     Hunted = Pnam2
                     tmrVibrate.Enabled = True
                    Select Case bgNum(39)   'how many of player2 seed are on it
                         
                         Case 1 ' only one, Captured  by player4
                       
                         BG(39).Picture = LoadPicture("") 'no seed will appear  on it
                         bgNum(39) = 0                     'num  of seeds is zero
                         BG(39).Tag = ""   'no player's seed is on it
                        
                         Case 2 ' only two, one is  Captured  by player4
                        
                         BG(39).Picture = Ply4_1.Picture ' one has been killed it remains one
                         bgNum(39) = 1
                
                          Case 3 ' only three, one is  Captured  by player4
                         
                         BG(39).Picture = Ply4_2.Picture ' one has been killed it remains two
                         bgNum(39) = 2
                        
                          Case 4 ' only four, one is  Captured  by player4
                         
                         BG(39).Picture = Ply4_3.Picture ' one has been killed it remains three
                         bgNum(39) = 3
                        
                 End Select
                
                         
                         For J = 0 To 3 'to return player2 back to it's prison
                         If Ply2(J).Visible = False Then Exit For 'to find an empty imagebox to put it in
                         Next
                         'Ply2(j).Picture = Ply2_1.Picture
                         Ply2(J).Visible = True
                         Ply2(J).Tag = 0  'not yet out
                         
                         Ply4(Index).Visible = False
                         Ply4(Index).Tag = 3 'it goes home for capturing an enemy
                         'Home4
                        ' f = MsgBox(Pnam4 & " Has Captured " & Pnam2 & "'s seed", vbInformation)

                   
            Case "P3"  'already occupied by player3
                     Hunted = Pnam3
                     tmrVibrate.Enabled = True
                     Select Case bgNum(39)   'how many of player3 seed are on it
                         
                         Case 1 ' only one, Captured  by player4
                       
                         BG(39).Picture = LoadPicture("") 'no seed will appear  on it
                         bgNum(39) = 0                     'num  of seeds is zero
                         BG(39).Tag = ""   'no player's seed is on it
                        
                         Case 2 ' only two, one is  Captured  by player4
                        
                         BG(39).Picture = Ply3_1.Picture ' one has been killed it remains one
                         bgNum(39) = 1
                
                          Case 3 ' only three, one is  Captured  by player4
                         
                         BG(39).Picture = Ply3_2.Picture ' one has been killed it remains two
                         bgNum(39) = 2
                        
                          Case 4 ' only four, one is  Captured  by player4
                         
                         BG(39).Picture = Ply3_3.Picture ' one has been killed it remains three
                         bgNum(39) = 3
                        
                 End Select
                   
                         For J = 0 To 3 'to return player4 back to it's prison
                         If Ply3(J).Visible = False Then Exit For 'to find an empty imagebox to put it in
                         Next
                         'Ply3(j).Picture = Ply3_1.Picture
                         Ply3(J).Visible = True
                         Ply3(J).Tag = 0  'not yet out
                         
                         Ply4(Index).Visible = False
                         Ply4(Index).Tag = 3 'it goes home for capturing an enemy
                         'Home4
              'f = MsgBox(Pnam4 & " Has Captured " & Pnam3 & "'s seed", vbInformation)


             
             Case "P4"        'already occupied by player4
                  
                     Ply4(Index).Visible = False   'remove seed from prison
                Select Case bgNum(39) 'check how many of player4 seed is on it
                  Case 1      'increment the num of players & also the pic
                  bgNum(39) = 2
                  BG(39).Picture = Ply4_2.Picture
                  Case 2
                  bgNum(39) = 3
                  BG(39).Picture = Ply4_3.Picture
                  Case 3
                  bgNum(39) = 4
                  BG(39).Picture = Ply4_4.Picture
                  Case Else: MsgBox "An impossible error occured" 'impossible or error
                End Select

        Case Else   'no player's seed is on it
            
            BG(39).Tag = "P4"  'put player3's seed on it
            BG(39).Picture = Ply4_1.Picture
            bgNum(39) = 1 'one player is on it
            Ply4(Index).Visible = False   'remove from prison
            
            End Select
          'Label2.Caption = BG(39).Tag: Beep
          Else 'rand <> 6
         Label2.Caption = "Can't Move This Seed With A Throw of" & Str(Rand)
         Exit Sub
        End If
      End If
    If Nclik = 1 Then
        Rand = RandNx
        DiceRowed = True
    Else
        StillPlaying = False
        TransSlate.Enabled = True 'transfer dice slate to the next player
        DiceRowed = False
        Nclik = 0
        Rand1 = 0
        Rand2 = 0
  End If

Else   'user has not rowed d dice
Label2.Caption = "Please Row The Dice"
End If

End Sub
Sub CheckFrontDoor()
Dim K1 As Boolean, K2 As Boolean, Seed As Integer
      CheckedFrontDoor = False
      If Counter = 0 Then Exit Sub 'no seed is inside

      If Rand1 <> 6 And Rand2 <> 6 Then Exit Sub 'must have a 6

     'instance 1
     'probably another ply seed is some steps in front of the door,kill it
     K1 = False
     K2 = False
     'remember rand1 or rand2 contains 6
     If Rand2 = 6 Then
     If BG(Start + Rand1).Tag <> "" And BG(Start + Rand1).Tag <> Turn Then K1 = True
     Else
     If BG(Start + Rand2).Tag <> "" And BG(Start + Rand2).Tag <> Turn Then K2 = True 'probably another ply seed is some steps in front of the door,kill it
     End If
     
       If K1 = True Or K2 = True Then
3      Seed = Int(Rnd * 4)
        Select Case Turn
          Case "P1"
          If Ply1(Seed).Tag = 1 Then GoTo 3 'the seed randomly choosed is already out take another
          Ply1_Click Seed
          Case "P2"
          If Ply2(Seed).Tag = 1 Then GoTo 3 'the seed randomly choosed is already out take another
          Ply2_Click Seed
          Case "P3"
          If Ply3(Seed).Tag = 1 Then GoTo 3 'the seed randomly choosed is already out take another
          Ply3_Click Seed
          Case "P4"
          If Ply4(Seed).Tag = 1 Then GoTo 3 'the seed randomly choosed is already out take another
          Ply4_Click Seed
          Case Else
          MsgBox "Impossible Error Occured At moveSeedAI"
        End Select
        BG_Click Start  'move the seed to kill it
      CheckedFrontDoor = True
      End If

End Sub
Sub CheckDoor()
Dim Seed As Integer
 'instance 2
      'another player is on it's door kill it
      'it must be able to move another seed with thye other die
            CheckedDoor = False
           If Counter = 0 Then Exit Sub  'no seed is inside exit
      If Rand1 <> 6 And Rand2 <> 6 Then Exit Sub  'must have a 6

      If BG(Start).Tag <> "" And BG(Start).Tag <> Turn Then
              ComOne = True
              

2            Seed = Int(Rnd * 4)
        Select Case Turn
          Case "P1"
          If Ply1(Seed).Tag = 1 Then GoTo 2 'the seed randomly choosed is already out take another
          Ply1_Click Seed
          Case "P2"
          If Ply2(Seed).Tag = 1 Then GoTo 2 'the seed randomly choosed is already out take another
          Ply2_Click Seed
          Case "P3"
          If Ply3(Seed).Tag = 1 Then GoTo 2 'the seed randomly choosed is already out take another
          Ply3_Click Seed
          Case "P4"
          If Ply4(Seed).Tag = 1 Then GoTo 2 'the seed randomly choosed is already out take another
          Ply4_Click Seed
          Case Else
          MsgBox "Impossible Error Occured At checkfrontdoor;Turn Had An unknown VAlue--" & Turn: Exit Sub
        End Select
        
                 MoveToKill  ' not so simple;  if it does not find another seed to move the other die then exit sub
     
        CheckedDoor = True
       End If
  ComOne = False
End Sub
Sub CheckCanKillWithOneDie()
 Dim J As Integer, Temp As Integer, Seed As Integer
'instance 3
 OneDie = False
 If Counter = 0 Then Exit Sub 'no seed inside
 If Rand1 <> 6 And Rand2 <> 6 Then Exit Sub
       If Rand1 = 6 Then
       Temp = Rand2
       Else
       Temp = Rand1
       End If
       Dim TurningPoint As Integer
       'temp contains the die that don't have a 6 except when we have a 6,6
       For J = 0 To 71  'or 0 To 51     'probably this player's seed already out, can kill a seed with the other die
         If BG(J).Tag = Turn Then
              If J + Temp > 71 Then GoTo 111
              TurningPoint = J
              
              If J + Temp > 51 Then
                 For H = 1 To Temp
                    TurningPoint = TurningPoint + 1
                    If TurningPoint = 52 Then TurningPoint = 0
                 Next
              Else
                 TurningPoint = TurningPoint + Temp
              End If
              If BG(TurningPoint).Tag <> "" And BG(TurningPoint).Tag <> Turn Then 'can it kill if this seed go with this die
                 Seed = 0
                 Select Case Turn 'let's first move out the seed inside b4 moving to kill so that the die with a 6 will be used
                        Case "P1"
1                            If Ply1(Seed).Tag = 1 Then
                                Seed = Seed + 1
                                If Seed = 4 Then MsgBox ("Error:4 seeds are inside and checkcankillwithonedie was called upon"): Exit Sub
                                GoTo 1
                            End If
                         
                            Ply1_Click Seed
                        Case "P2"
2                             If Ply2(Seed).Tag = 1 Then
                                Seed = Seed + 1
                                If Seed = 4 Then MsgBox ("Error:4 seeds are inside and checkcankillwithonedie was called upon"): Exit Sub
                                GoTo 2
                            End If
                         
                            Ply2_Click Seed
                        Case "P3"
3                            If Ply3(Seed).Tag = 1 Then
                                Seed = Seed + 1
                                If Seed = 4 Then MsgBox ("Error:4 seeds are inside and checkcankillwithonedie was called upon"): Exit Sub
                                GoTo 3
                            End If
                         
                            Ply3_Click Seed
                        Case "P4"
4                             If Ply4(Seed).Tag = 1 Then
                                Seed = Seed + 1
                                If Seed = 4 Then MsgBox ("Error:4 seeds are inside and checkcankillwithonedie was called upon"): Exit Sub
                                GoTo 4
                            End If
                         
                            Ply4_Click Seed
                       
                  End Select
                  
                 BG_Click J  'move seed to kill
                 OneDie = True
                 Exit Sub
            End If
          End If
111      Next

End Sub
Sub CheckCanKillWith6Die()
 'well, if we had, say 6,3 checkwithonedie can't kill with a 3,but what if there are more than
                 'one seed on the field and a die could kill with a 6 and the other move with d 3
                 'if we don't include this checkcanbringoutseed will bring out a seed and thus we'll
                 'loose this chance of killn
                 
                 'now we must have more than 1 seed on the field for this to work
                 From6Die = False
                 SixDie = False
                 Dim Kount As Integer, Indekx(4) As Integer, Foo As Integer, J As Integer
                
                   For J = 0 To 71
                      If BG(J).Tag = Turn Then
                        Foo = Foo + bgNum(J)
                         Kount = Kount + 1
                         Indekx(Kount) = J
                         End If
                   Next
                   If Foo < 2 Then Exit Sub 'we need at least 2 seeds on the field
                 ' one of the die must at least be a 6
                 Dim target As Integer
                   For J = 1 To Kount 'now let's see if any of those that are out can kill with a 6 and move another player with the other die
                       If Indekx(J) + 6 > 71 Then GoTo 31 'to wave off errors
                       target = Indekx(J)
                       If (target + 6) > 51 Then  '--u turn problem., turning point
                           For H = 1 To 6
                               target = target + 1
                               If target = 52 Then target = 0
                            Next
                        Else
                            target = target + 6
                        End If
                        
                       If BG(target).Tag <> "" And BG(target).Tag <> Turn Then
                        'make sure the active die is the one with a 6
                          If Rand <> 6 Then
                            Foo = Rand
                            Rand = RandNx
                            RandNx = Foo
                          End If
                          From6Die = True  'after the simulation of this, the next die is moved at MoveToKill
                          BG_Click Indekx(J)
                          SixDie = True
                          Exit Sub
                       End If
31                   Next
                       
                     
                       

End Sub
Sub CheckCanKillWithBothDie()
  Dim k As Integer
  'this only done incase we have a 6,3 and a die can kill with the sum of both,
  'instead of bringing out a seed kill
  BothDie = False
  If Rand1 <> 6 And Rand2 <> 6 Then Exit Sub 'or if (rand1 and rand2)<> 6 then... bit level
 ' If Not (Rand1 Xor Rand2) Then Exit Sub   'simply put, if rand1=rand2=6 then ...   we dont need a 6,6 if only one player is out and some are still inside, we may not have another player to move the freethrow,xcept when we have one and only one seed remaining
  Dim X As Integer
  X = 0
  For k = 0 To 71
   If X = 52 Then X = 0
   If BG(X).Tag = Turn Then
      If X + Rand1 + Rand2 > 71 Then Exit Sub
      If BG(X + Rand1 + Rand2).Tag <> "" And BG(X + Rand1 + Rand2).Tag <> Turn Then
         Rand = Rand1 + Rand2
         One = True
         Nclik = 2
         BG_Click X
         BothDie = True
         Exit Sub
     End If
   End If
   X = X + 1
  Next
End Sub
Sub CheckCanBringOutSeed()
Dim Seed As Integer
'A seed must be inside and must have a 6 to bring out a seed
BringOut = False
ComOne = False
If Counter = 0 Then Exit Sub 'no sed is inside
If Rand1 <> 6 And Rand2 <> 6 Then Exit Sub
Seed = -1
1 Seed = Seed + 1
If Seed > 3 Then MsgBox ("Error: Seed had a value of 4 and counter says seeds are inside"): Exit Sub

    Select Case Turn
           Case "P1"
           
                If Ply1(Seed).Tag = 1 Then GoTo 1
             
                Ply1_Click Seed
                
           Case "P2"
                If Ply2(Seed).Tag = 1 Then GoTo 1
                  
                Ply2_Click Seed
            Case "P3"
                If Ply3(Seed).Tag = 1 Then GoTo 1
              
                Ply3_Click Seed
                
           Case "P4"
                If Ply4(Seed).Tag = 1 Then GoTo 1
                  
                Ply4_Click Seed
          Case Else
                MsgBox "variable 'Turn' Assumed An Unknown Value,Turn=" & Turn: Exit Sub
         End Select
 ComOne = True
MoveToKill
BringOut = True
End Sub
Sub MoveToKill()
Dim Fdie As Boolean, Sdie As Boolean, TempRand1 As Integer, TempRand2 As Integer
'first check if any seed can kill with the first die
'check if another seed can kill with the other die
'check if any seed can kill with the sum of both die
'else strategic move ,i know what to do

'if comone=true then the computer needs to move only one seed

  Dim Indx As Integer, J As Integer

  If killWith1die = True Then GoTo 2
  If OnlyOneCanKill Then GoTo 3
  Fdie = False
  Sdie = False
  Seed1 = 111 'reset; don't want zero
  Seed2 = 111
  TempRand1 = Rand
  TempRand2 = RandNx
  KillWith2die = False
  
  If ComOne = True Then  'was called from routine that only needs to move a seed with d die that does not contain a 6,well xcept both are 6,6
     If Rand = 6 Then 'either of then must be a 6
        TempRand1 = RandNx  'making both have the same valu is just to ensure that only one die is used although there're some repititions,or else we'll have to write it own code,taking space
        TempRand2 = RandNx
    Else 'rand2 has the 6
        TempRand1 = Rand
        TempRand2 = Rand
    End If
 End If
 Dim Kount As Integer, Tarket As Integer
 Indx = 0
 
  For Indx = 0 To 52 'suppose to be 72 but waste unnecessary time cause any seed above this index cannot kill it is on it private route
        If BG(Indx).Tag = Turn Then
           If Indx + TempRand1 > 71 Then GoTo 1 'to avoid overflow
           Tarket = Indx + TempRand1
           If Indx + TempRand1 > 51 Then 'u turn exceded
               Tarket = Indx
              For Kount = 1 To TempRand1
                  Tarket = Tarket + 1
                  If Tarket = 52 Then Tarket = 0
              Next
            End If
            
           If BG(Tarket).Tag <> "" And BG(Tarket).Tag <> Turn Then 'can it kill with d first die
                Fdie = True
                Seed1 = Indx
           ElseIf Indx + TempRand2 < 72 Then
            Tarket = Indx + TempRand2
               If Indx + TempRand2 > 51 Then 'u turn exceded
                  Tarket = Indx
                  For Kount = 1 To TempRand2
                      Tarket = Tarket + 1
                      If Tarket = 52 Then Tarket = 0
                  Next
               End If

               If BG(Tarket).Tag <> "" And BG(Tarket).Tag <> Turn Then 'can it kill with d 2nd die
                Sdie = True
                Seed2 = Indx
               End If
          End If
           If ComOne = True Then GoTo 1 'it should not sum up when we are seeking to move a seed with a single die
           If From6Die Then GoTo 1
           If ((Indx + TempRand1 + TempRand2) < 72) Then
            Tarket = Indx + TempRand1 + TempRand2
               If Indx + TempRand1 + TempRand2 > 51 Then 'u turn exceded
                  Tarket = Indx
                  For Kount = 1 To TempRand1 + TempRand2
                      Tarket = Tarket + 1
                      If Tarket = 52 Then Tarket = 0
                  Next
               End If

           If BG(Tarket).Tag <> "" And BG(Tarket).Tag <> Turn Then 'can it kill with both dice,d first condition is to avoid error
                Seed1 = Indx
                KillWith2die = True 'we've found a seed that can kill with the sum of the two dice
            End If
           End If
          
           If Fdie = True And Sdie = True Then Exit For 'we can kill two diffrent enemy's seeds with 2 diffrent seeds of the present player
       End If
1
  Next
           'it's not essential thet at the end of the loop fdie and sdie must be true,theysimply tell if a we can kill with d value of the dice
           If Fdie And Sdie Then
              killWith1die = True
              BG_Click Seed1 ' this position holds a seed a that can kill (seed1 and seed2)
              Exit Sub  'exit so that d computer will allow this seed to finish moving b4 clikn on the next one;since it is a timer that does the movn, and will not be performed until this sub is completed
2             killWith1die = False 'set to false so that this not get repeated since this is the 2nd move
              BG_Click Seed2
              Exit Sub
           End If
           
           If ComOne Or From6Die Then
              'Rand = TempRand1 'so that it moves with the value that may not be a 6
              If Fdie = True Then 'well for comone only fdie could be true even if temprand1 and temprand2 are the same cause of the if else
                  BG_Click Seed1  'this seed can kill
              Else 'no seed can kill simply move a seed
                 'find a best move to make
               MoveNum = 1 'move only one seed
               StrategicMovement
              End If
              ComOne = False
              From6Die = False
              Exit Sub
           End If
           
           
           If (Fdie = True And Sdie = False) Or (Fdie = False And Sdie = True) Then 'only one of the d seed can kill simply move the another seed
               OnlyOneCanKill = True
               If Fdie Then
                  BG_Click Seed1
               Else 'first die is false 2nd die is true
               tem = Rand 'swap ,use the 2nd die
               Rand = RandNx
               RandNx = tem
                  BG_Click Seed2
               End If
               Exit Sub 'allow the seed clicked to move b4 clicking the 2nd time
3              'make the best move
               MoveNum = 1 'we've already moved one seed move another
               StrategicMovement
               OnlyOneCanKill = False
              Exit Sub
          End If
          
           If KillWith2die Then  'a singe seed can kill with the sum of both dice
              If Seed1 = 111 Then GoTo 34
               One = True
               Nclik = 2
               Rand = Rand1 + Rand2
               BG_Click Seed1
               KillWith2die = False
               Exit Sub
            End If
          
          If Fdie = False And Sdie = False Then
            'make the best move since no seed can presently kill;remember you have to move two seed or even one but make use of the two dice
34           MoveNum = 2
            StrategicMovement 'does not involve any kills since none is possible at this point,simply strategically place seed
          End If
         
         
                

End Sub
Sub StrategicMovement()
'the best mode of defence is to attack
'so attack will comes b4 defence
'in a measure to attack don't leave yourself open
'get the position of all seeds that can pose a threat to the opponent _
 i.e can move so close(12 steps in range from the valu of presentPositn+Rand) _
 to an opponent seed with d value of the present dice _
get the position of those that in so doing will cross over the seed of an enemy _
thereby leaving itself open
'if any seed is at the door or within the corridor of an opponent move it away -defence
'if by moving a seed it target position will be the door of an opponent consider another option if any

 Dim Position(4) As Integer, Bestmove(4) As Integer, Counter As Integer, Counter1 As Integer, Counter2 As Integer, Count As Integer
 Dim CrossOver As Boolean, Counter3 As Integer, Counter4 As Integer, HRecommended(4) As Integer, Closer(4) As Integer
 Dim Seed As Integer, J As Integer, Stert(4) As Integer, Ranges As Integer
 Dim k As Integer, H As Integer, LastLine As Integer


 Stert(1) = 0
 Stert(2) = 13
 Stert(3) = 26
 Stert(4) = 39
 
 Seed = 111
 
 
  For Counter = 1 To 4  'initialise;didn't want to use zero bcos there exist bg(0)
      Position(Counter) = 111
      Bestmove(Counter) = 111
      Closer(Counter) = 111
      HRecommended(Counter) = 111
  Next
  Count = 0
  Dim C As Integer
  C = 0
  ' let's get the no of seeds on the field
  For Counter = 0 To 71
      If C = 52 Then C = 0 'u turn
      If BG(Counter).Tag = Turn Then
         Distance Counter     'ensure that the seed can move
         If Exceed = False Then
            If C + Rand < 72 Then
               If BG(C + Rand).Tag <> "" And BG(C + Rand).Tag <> Turn Then
                  BG_Click Counter  'let's see probably movetokill left it off
                  Exit Sub
               End If
           End If
           Count = Count + 1
           Position(Count) = Counter
         End If
      End If
      C = C + 1
 Next
 
 If Count = 0 Then 'found out that this is usually true if only one seed is on d field and it is on motion i.e d timer has not finished simulation
   ' MsgBox "Error:No Seed Was Found On The Field or the one out cannot move,Yet Control Was Transfered To ArtificialIntelligence": Exit Sub
 End If
  Dim Posti As Integer
  
 For Counter = 1 To 4
     CrossOver = False
  
     If Position(Counter) = 111 Then Exit For 'no more seeds are out
     Posti = Position(Counter)
     
         'decided to add this for..next loop so that it will b smarter
       'if it is going to land on opponents door this seed is not the best option
          For H = 1 To 4
           If Position(Counter) + Rand = Stert(H) Then CrossOver = True
          Next
            If Position(Counter) + Rand = 52 Then CrossOver = True
     For Counter1 = 1 To Rand
       Posti = Posti + 1
       If Posti > 71 Then Exit For
       If Posti = 52 Then Posti = 0  'so that it will crossover -continiuos
       If BG(Posti).Tag <> "" And BG(Posti).Tag <> Turn Then
          'if this seed is moved it will cross over an opponents seed opening itself to attack
          CrossOver = True
        End If
     Next
     
     If CrossOver = False Then
           Bestmove(Counter) = Position(Counter) 'this seed does not have an opponents seed on it path if moved
     End If
 Next
   Counter2 = 0
 For Counter = 1 To 4  'let's see if we have any best move
     If Bestmove(Counter) = 111 Then Counter2 = Counter2 + 1
 Next
 
 If Counter2 = 4 Then
   'no seed has it path clear
   'work with position() for defence
   'move a seed that is at d door of an opponent
       For k = 1 To 4
             For H = 1 To 4
                 If Position(k) = Stert(H) Then 'at the door of an opponent move it
                   BG_Click Stert(H)
                   Exit Sub
                 End If
              Next
         Next
     
   'move a seed that has opponent seed at it behind
    For k = 1 To 4
              If Position(k) = 111 Then Exit For
                 LastLine = Position(k)
              For H = 1 To 12
                  LastLine = LastLine - 1
                  If LastLine = -1 Then LastLine = 51
                  If BG(LastLine).Tag <> "" And BG(LastLine).Tag <> Turn Then
                     BG_Click Position(k)
                     Exit Sub
                  End If
              Next
         Next
         
   'move a seed that is within the premises of an opponent
            For k = 1 To 4
             For H = 1 To 4
                 Ranges = Stert(H) + 6
                 If Position(k) > Start And Position(k) <= Ranges Then
                    BG_Click Position(k)
                    Exit Sub
                 End If
              Next
         Next

   'if no other measure of defence is applicable in this scenario simply move a seed
    For k = 0 To 71
             For H = 1 To 4
                 If Position(H) = k Then
                    BG_Click k
                    Exit Sub
                 End If
             Next
         Next
   
 Else
   'which of them will pose a threat to an opponent
   'by getting so close
   'we have to check out d distance from the target position to an opponents seed
   'and we move the seed with the shortest distance
    Counter4 = 0
    Dim Sum As Integer
    For Counter3 = 1 To 4
        If Bestmove(Counter3) = 111 Then GoTo 22
        Sum = Bestmove(Counter3) + Rand
        For Counter1 = 1 To 12  'let's see if it can get close by max of 12 steps
            Sum = Sum + 1
            If Sum = 52 Then Sum = 0
            If Sum > 71 Then Exit For
            'If Bestmove(Counter3) + Rand + Counter1 > 71 Then Exit For
            'If BG(Bestmove(Counter3) + Rand + Counter1).Tag <> "" And BG(Bestmove(Counter3) + Rand + Counter1).Tag <> Turn Then
            If BG(Sum).Tag <> "" And BG(Sum).Tag <> Turn Then
               Counter4 = Counter4 + 1    'this seed will have some enemy's seed somesteps in front to attack
               If Counter4 > 4 Then Exit For
               HRecommended(Counter4) = Bestmove(Counter3)
               Closer(Counter4) = Counter1  'this is used to determine the distance from this seed target position to the position of the opponent then will choose the shortest distance
            End If
        Next
22   Next
      If Counter4 = 0 Then 'all the seeds that have their track clear cannot pose a threat to an opponent
         'go on to defence
         'work with bestmove()
         'move a seed that is at d door of an opponent or even at it own door
         For k = 1 To 4
             For H = 1 To 4
                 If Bestmove(k) = Stert(H) Then 'at the door of an opponent move it
                   BG_Click Stert(H)
                   Exit Sub
                 End If
              Next
         Next
         
                               
         'move a seed that has opponent seed at it behind by 12 step
          For k = 1 To 4
              If Bestmove(k) = 111 Then Exit For
                 LastLine = Bestmove(k)
              For H = 1 To 12
                  LastLine = LastLine - 1
                  If LastLine = -1 Then LastLine = 51
                  If BG(LastLine).Tag <> "" And BG(LastLine).Tag <> Turn Then
                     If ComOne = True Then 'move with the other die that is not a 6
                        If Rand = 6 Then
                        T = Rand
                        Rand = RandNx
                        RandNx = T
                        End If
                     End If
                     BG_Click Bestmove(k)
                     Exit Sub
                  End If
              Next
         Next
                 
         'move a seed that is within the premises of an opponent, from door to 6 steps forward
        
         For k = 1 To 4
             If Bestmove(k) = 111 Then Exit For
             For H = 1 To 4
                 Ranges = Stert(H) + 6
                 If Bestmove(k) > Stert(H) And Bestmove(k) <= Ranges Then
                    BG_Click Bestmove(k)
                    Exit Sub
                 End If
              Next
         Next
         
         
         
         'if no other measure of defence simply move a seed
         
         For k = 71 To 0 Step -1
             For H = 1 To 4
                 If Bestmove(H) = k Then
                    BG_Click k
                    Exit Sub
                 End If
             Next
         Next
         
         
         
     Else  'move the seed that will get closer to opponent seed
          'simply find which has the least value from 1 to 12 i.e from closer()
               Seed = 1
          For Counter = 1 To 4
              If Closer(Counter) = 111 Then Exit For 'this should not happen in the first pass
         
              For J = 1 To 4
                  If Closer(J) = 111 Then GoTo 23
                  If Counter = J Then GoTo 23
                  If Closer(Counter) < Closer(J) Then
                     Seed = Counter
                     
                  Else
                     Seed = J
                  End If
23              Next
          Next
     'wow by now we have a seed that has passed all this rigorous test
     'simply click onit,this is the highly recommended seed
     If Seed = 111 Then Exit Sub: Beep 'this should never be performed
         BG_Click HRecommended(Seed)
     End If
  End If
End Sub

Private Sub moveSeedAI_Timer()

If Paused Then Exit Sub
ClickIsFromComputer = True
Dim Numb As Integer, Seed As Integer, Locatn(4) As Integer, Bestmove(4) As Integer, WrongMove(4) As Integer
Dim p As Integer, M As Integer
Randomize

For M = 1 To 4  'reset; didn't want to use zero because it is an index n bg
Locatn(M) = 99
Bestmove(M) = 99
Next

For k = 0 To 71
  If BG(k).Tag = Turn Then
  Numb = Numb + bgNum(k)
  p = p + 1
  Locatn(p) = k
  End If
Next
Counter = 0
     Dim Ins As Integer
     Select Case Turn
            Case "P1"
            
                For Ins = 0 To 3
                    If Ply1(Ins).Tag = 0 Then Counter = Counter + 1
                Next
            Case "P2"
            
                For Ins = 0 To 3
                    If Ply2(Ins).Tag = 0 Then Counter = Counter + 1
                Next
            Case "P3"
            
                For Ins = 0 To 3
                    If Ply3(Ins).Tag = 0 Then Counter = Counter + 1
                Next
            Case "P4"
            
                For Ins = 0 To 3
                    If Ply4(Ins).Tag = 0 Then Counter = Counter + 1
                Next
            Case Else
              MsgBox "Error:Turn Variable Contains An Unknown Data--" & Turn
      End Select

If Numb = 0 And Counter <> 0 Then   'ensures that a seed is brought out
   If Rand1 = 6 Or Rand2 = 6 Then
10 Seed = -1
11    Seed = Seed + 1
      If Seed = 4 Then MoveToKill: Exit Sub
       Select Case Turn
          Case "P1"
          If Ply1(Seed).Tag = 1 Then GoTo 11
          Ply1_Click Seed
          Case "P2"
          If Ply2(Seed).Tag = 1 Then GoTo 11
          Ply2_Click Seed
          Case "P3"
          If Ply3(Seed).Tag = 1 Then GoTo 11
          Ply3_Click Seed
          Case "P4"
          If Ply4(Seed).Tag = 1 Then GoTo 11
          Ply4_Click Seed
          Case Else
          MsgBox "Impossible Error Occured At moveSeedAI"
      End Select
      If Rand1 = 6 And Rand2 = 6 Then 'if this is true the program will not automatically move the other die
      'simply bring out another seed
      'we should not test if it can kill since if we kill we may not have another seed to move for the free throw
      'especially if no 6 shows up
      GoTo 10 'at the second click rand1 and rand2 are initialise to 0
      End If
    Else
   '           MsgBox "Error: No seed is out and the two die is not 6,yet Ai was called on fstdie and snddie is true:" & "  " & FstDie & SndDie
   End If
Else 'If Numb = 1 Then
      'only one player is one the field, d only reason we may have, not to click on
      'this seed is if one d die is a 6 and we have seed(s) inside.
      'if we can kill with d sum of both die
      'if only one seed and we have a 6 bring out another seed
      
                   'hint:i can Create diffrent procedures for each situation
                  '       and call d required ones in order of 'precedence'
     
      If Rand1 <> 6 And Rand2 <> 6 Then GoTo 21
      
      CheckFrontDoor  'see if an enemy's seed is at a place such as 6,5 and the die matchs this
      If CheckedFrontDoor = True Then GoTo Endn
      
      CheckDoor  'see if there is an enemy's seed at the door then kill it,move the other die with the other seed,this seed must be able to move
      If CheckedDoor = True Then GoTo Endn
      
      CheckCanKillWith6Die ' if the die is say 6,4 and a seed can move  & kill with a 6 and another can move with a 4
      If SixDie = True Then GoTo Endn
      
      
      CheckCanKillWithOneDie 'if the die is say 6,4 and the seed out can kill with a 4 then move out a 6
      If OneDie = True Then GoTo Endn
      
      CheckCanKillWithBothDie  'see if the seed out can kill with the sum of both die,this will be necessary if any of the die is a 6,since i've already made the program in such a way that it moves a seed(only one on the field)with the sum of both dice
      If BothDie = True Then GoTo Endn
      
      
      CheckCanBringOutSeed
      If BringOut = True Then GoTo Endn
      
21      MoveToKill
      GoTo Endn
      
                 
                  
End If
'End If

Endn:
moveSeedAI.Enabled = False
          
          
          


End Sub

Private Sub Quit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider = True

End Sub

Private Sub throwdiceAI_Timer()
If Vibrating Then Exit Sub
If Paused Then Exit Sub

Static Numklik As Integer
Dim hhh As Integer
Numklik = Numklik + 1
Command1_Click         'clik d button twice
If Numklik >= 2 Then
Numklik = 0
throwdiceAI.Enabled = False
If FstDie = True And SndDie = True Then moveSeedAI.Enabled = True   'some seeds can move
Exit Sub
End If
throwdiceAI.Interval = 1500 'allow human to see d dice rolling


End Sub
Sub ArtificialIntelligence()
Command1.Enabled = False 'human player should not interfer
throwdiceAI.Interval = 500   'first of all row the dice and see the outcome
throwdiceAI.Enabled = True
End Sub




Private Sub tmrToolSlider_Timer()
If Vibrating Then Exit Sub

Static Kount As Long
    tmrToolSlider.Interval = 1
    
    If Paused Then
        Label3.Caption = "Game Paused"
        Label2.Caption = "Press Enter To Resume"
        Label3.ForeColor = vbRed
        Label2.ForeColor = vbRed
    Else
       Label2.ForeColor = Label3.ForeColor 'to turn back to brown
    End If
    
    If Slider = True Then
        Kount = Kount + 1
        If Kount = 1 Then
            wmpFreedSound.URL = Dirxtry & "gemvanishes.ogg"
        End If
        If Board.Width >= 11310 Then Slider = False: tmrToolSlider.Interval = 0: Exit Sub
        Board.Width = Board.Width + 30
    Else
        If Board.Width <= 9750 Then Board.Width = 9750: Kount = 0: Exit Sub
        Board.Width = Board.Width - 5
        wmpFreedSound.URL = ""
    End If
End Sub


Private Sub tmrVibrate_Timer()
 Static Kount As Integer, Flik As Integer, Leftt As Integer, Topp As Integer
    tmrVibrate.Interval = 20
    Kount = Kount + 1
    Vibrating = True
    If Kount = 1 Then 'to avoid overkill---no sound will be heard without this if then
        wmpFreedSound.URL = Dirxtry & "ballsdestroyed4.ogg" 'time lapse: 1 sec
        Leftt = Me.Left  'store the current position of the form b4 vibrating
        Topp = Me.Top
    End If
    
    Flik = (Flik Mod 2) + 1
    If Flik = 1 Then
        Board.Move Board.Left + 50, Board.Top + 50
    Else '=2
        Board.Move Board.Left - 50, Board.Top - 50
    End If
    
    If Kount = 10 Then 'stop vibration
        Kount = 0
        Me.Left = Leftt 'original position
        Me.Top = Top
        tmrVibrate.Enabled = False
        If Not Paused Then 'if the game is not already paused pause it, this is essential,cos when d form loads it is paused
           PauseStart.Value = True 'pause game
        End If
    
        Msgbx = MsgBox(Hunter & " Has Captured " & Hunted & "'s Seed", vbInformation + vbOKOnly)
        
       If Msgbx = vbOK Then
       Vibrating = False
           If Paused Then 'unpause---if there is any word like that
              PauseStart.Value = True 'resume
           End If
        
            Select Case Identity
                   Case "P1": Home1
                   Case "P2": Home2
                   Case "P3": Home3
                   Case "P4": Home4
             End Select
       End If
    End If
End Sub

Private Sub TransSlate_Timer()
If Pnum = 0 Then MsgBox ("Invalid Entry, Program Cannot Begin From Here"): End
If MustTransfer Then MustTransfer = False: GoTo 1
If Paused Then Exit Sub
If Vibrating Then Exit Sub

Dim Inn As Integer, Numo As Integer, Numm As Integer
'If StillPlaying = True Then GoTo 11

1 Command1.Enabled = True

ClickIsFromComputer = False 'will be made true once control is transfered to computer player

Inn = 99 'the '99' is not used just as a default valu
LastStep = False

If FreeThrow Then GoTo 2  'the last valu gotn was 6,6 dont transfer dice slate to the next payer
       
       Rotate = (Rotate Mod Pnum) + 1 'rotate turn between num of players
2

 PlayerType = 1 'human
Select Case Rotate
Case 1
           LtLimit = 56
            NxLimit = 52
            Start = 0
           'we want to see if a player hav finished his/her course
            For J = 0 To 71
              If BG(J).Tag = "P1" Then '
                 Numo = Numo + bgNum(J)  'didn't want to use 'one' bcos we could have more then one player on a position
               End If
             Next
             Inn = 0
                For k = 0 To 3  'let's see if any is inside
                 If Ply1(k).Tag = 0 Then Inn = Inn + 1
                  Next
                  If Numo = 0 Then 'no player is out
                      If Inn = 0 Then 'no player is inside
                                      FreeThrow = False 'if a player should go home with a 6,6 we'll enter an end less loop

                    GoTo 1 'transfer dice slate to next player
                 End If
              End If
              
             Frame1.Move P1Left, P12Top
             Turn = "P1"
           
             If Ptyp1 = 11 Then PlayerType = 11: ArtificialIntelligence
    
Case 2
               LtLimit = 61
                NxLimit = 57
                Start = 13
               For J = 0 To 71
              If BG(J).Tag = "P2" Then '
                 Numo = Numo + bgNum(J)  'didn't want to use 'one' bcos we could have more then one player on a position
               End If
             Next
               Inn = 0
                For k = 0 To 3  'let's see if any is inside
                 If Ply2(k).Tag = 0 Then Inn = Inn + 1
                  Next
                  If Numo = 0 Then 'no player is out
          
                If Inn = 0 Then 'no player is inside 1.e they've all finished their course
                                   FreeThrow = False 'if a player should go home with a 6,6 we'll enter an end less loop
 
                    GoTo 1 'transfer dice slate to next player
                 End If
               End If
             Frame1.Move P2Left, P12Top  'move to the next player
             Turn = "P2"
             If Ptyp2 = 11 Then PlayerType = 11: ArtificialIntelligence
Case 3
                  LtLimit = 66
                   NxLimit = 62
                   Start = 26
                  For J = 0 To 71
              If BG(J).Tag = "P3" Then '
                 Numo = Numo + bgNum(J)  'didn't want to use 'one' bcos we could have more then one player on a position
               End If
             Next
            Inn = 0
                For k = 0 To 3  'let's see if any is inside
                 If Ply3(k).Tag = 0 Then Inn = Inn + 1
                  Next
                  If Numo = 0 Then 'no player is out
             
                If Inn = 0 Then 'no player is inside
                FreeThrow = False 'if a player should go home with a 6,6 we'll enter an end less loop
                    GoTo 1 'transfer dice slate to next player
                 End If
          End If
             Frame1.Move P3Left, P34Top  'move to the next player
             Turn = "P3"
             If Ptyp3 = 11 Then PlayerType = 11: ArtificialIntelligence
             
Case 4
            LtLimit = 71
             NxLimit = 67
             Start = 39
            For J = 0 To 71
              If BG(J).Tag = "P4" Then '
                 Numo = Numo + bgNum(J)  'didn't want to use 'one' bcos we could have more then one player on a position
               End If
             Next
             Inn = 0
                For k = 0 To 3  'let's see if any is inside
                 If Ply4(k).Tag = 0 Then Inn = Inn + 1
                  Next
             If Numo = 0 Then 'no player is out
             
                If Inn = 0 Then 'no player is inside
                                FreeThrow = False 'if a player should go home with a 6,6 and that's the last seed ,we'll enter an end less loop

                    GoTo 1 'transfer dice slate to next player
                 End If
                End If
             Frame1.Move P4left, P34Top  'move to the next player
             Turn = "P4"
             If Ptyp4 = 11 Then PlayerType = 11: ArtificialIntelligence
End Select
    TransSlate.Enabled = False
    For H = 0 To 3
    Line1(H).Visible = False
    Next
   
DiceRowed = False
Label3.Caption = "Your Turn To Roll The Dice"
If FreeThrow Then Label3.Caption = "You Have A Free Throw"
Label2.Caption = ""
Die(0).Picture = Dice(6).Picture
Die(1).Picture = Dice(6).Picture
Nclik = 0

Verify Turn   'ensures that there are only and only 4 seeds per player

       For k = NxLimit To LtLimit
       If BG(k).Tag <> "" Then
        Numm = Numm + bgNum(k)     'total no of player in private route
        End If
       Next
         
       'if total num of player remaining is = to the num of player in the route then use only one die
     
    If Numm = Numo And Inn = 0 Then  'all the seed for this player has entered it's private route to home use only one die
        LastStep = True
       Die(0).Visible = False
       Else
        LastStep = False
        Die(0).Visible = True
      End If
11
End Sub

Private Sub Timer1_Timer()

End Sub


Private Sub Timer2_Timer()

If Running Or DiceRolling Then
   MustTrans.Enabled = False
   MustTransfer = False
Else
   MustTrans.Enabled = True
End If

End Sub

Private Sub Timer3_Timer()
If Vibrating Then Exit Sub
If Paused Then Exit Sub
Timer3.Enabled = False
If LastStep = True Then
Line1(2).Visible = False
Line1(3).Visible = False
Die(0).Visible = False
'LastStep = False
Else
Die(0).Visible = True
'Beep
End If
Die(1).Visible = True
'Timer3.Enabled = False
End Sub

Private Sub tmrDice_Timer()
If Vibrating Then Exit Sub

If Paused Then Exit Sub
Randomize
Rand1 = Int((Rnd * 6) + 1)
Rand2 = Int((Rnd * 6) + 1)
'Rand1 = 1
'Rand2 = 6
Die(1).Picture = Dice(Rand1 - 1).Picture 'index problem
Die(0).Picture = Dice(Rand2 - 1).Picture

End Sub

Private Sub tmrMoved_Timer()
If Vibrating Then Exit Sub
If Paused Then Exit Sub
'On Error Resume Next
Dim PrevIdx As Integer

    wmpPopup.URL = Dirxtry & "start.wav"
    NumMove = NumMove + 1  'no of simulated moves
    Indexx = Indexx + 1
    PrevIdx = Indexx - 1


       If Indexx = 52 Then  'only player1 never gets to this place
          Indexx = 0             'simply continue from 0 rather than 52 and recalculate the target position
          Last = 0 + (Rand - NumMove)
       End If
       
       If Indexx = Limit Then 'entry point for this seed
          PrevIdx = Limit - 1    'so that it will clear the last simulated move,only for player1 is this not needed
          Indexx = NxLimit
          Last = NxLimit + (Rand - NumMove)   'recalculate the target position,essential bcos it was last cal based on the value of indexx @ bg_click
       End If
       
   
        If Stored = True Then   'this positn contained a seed b4 it was swapped
            'Store(1).Picture = BG(Indexx).Picture  'moving seed
            BG(PrevIdx).Picture = Store(0).Picture  'retain it
            Stored = False
            Seen = True
        Else: Seen = False
        End If

        If Indexx <> Last Then   'make sure it's not the target, since it just a simulation,checkers takes care of the target i.e it should kill or top up the target position
            If bgNum(Indexx) <> 0 Then  'it already contains a player's seed, store the pic b4 swapping
                Store(0).Picture = BG(Indexx).Picture  'store the pic
                Stored = True    'set flag on so as to replace after d simulation
            Else
                Stored = False
            End If
        End If


        If overLap = True Then GoTo 1   'position contains more than 1 seed

        If Seen = False Then          'only clear it when it does not contain a seed at rest
            BG(PrevIdx).Picture = LoadPicture("")
        End If

1       overLap = False



    If Indexx - 1 = LtLimit Then
                                    Select Case Nam
                                            Case "P1"
                                                  Home1
                                            Case "P2"
                                                  Home2
                                            Case "P3"
                                                  Home3
                                            Case "P4"
                                                  Home4
                                      End Select
                                      GoTo 5
       End If
       
       
        BG(Indexx).Picture = Store(1).Picture  ' d main thing
        
        If NumMove = 1 Then 'this is for replacing a seed at the entrance point of another seed,it replaces it immiediately the moving seed gets to the next position
            If YesNo = True Then 'a seed was there b4, return it
                BG(HomeNum).Picture = Motion.Picture
                BG(HomeNum).Tag = Motion.Tag
                bgNum(HomeNum) = Temp
            End If
        End If

        If NumMove >= Rand Then  'simulation completed,the greater than ought not be there but it's safer
             Checkers    'checks target position,if it landed on an opponents seed
5            Running = False
             NumMove = 0
             tmrMoved.Enabled = False
             'DiceRowed = False
             Label2.Caption = ""
             YesNo = False
             If One = True Or LastStep = True Then Nclik = 2: One = False 'only one player is out and rand1 and rand2 has already been summed into rand in checkers
             If Nclik = 1 Then  'movement of seed has completed using d val of d first die
                Rand = RandNx    'let the current die be the val of d second die
                DiceRowed = True
                ' Print Nclik
                StillPlaying = True
            Else ' Nclik = 2'movement of seed has completed using d val of d 2nd die
                StillPlaying = False
                TransSlate.Enabled = True 'transfer dice slate to the next player
                DiceRowed = False 'reset
                Nclik = 0
                Exit Sub
            End If
            If PlayerType = 1 Then Exit Sub 'the remaining lines of code is to allow the computer play;it's position here is necessry bcos the computer has to wait for the movement of the seed b4 clicking the 2nd time
            If MoveNum = 2 Then StrategicMovement: MoveNum = 0 'the computer needs to move twice
           ' If ComTwo = 123 Then MoveToKill  'it remains one more seed to move for computer
            If OnlyOneCanKill Or killWith1die Or From6Die Then MoveToKill  'the computer can kill two diffrent seeds with 2 diffrent seeds moving with the valu of each die
             
        End If
End Sub

Public Sub Distance(Index)
Dim NumPlyNow As Integer
Select Case Turn   'this will be used in the timer event.the reason y we can't use bg().tag is that indexx is not stable in the said event
       Case "P1": Limit = 51: NxLimit = 52: LtLimit = 56: Entering = 44: Nam = "P1": NumPlyNow = P1out
       Case "P2": Limit = 12: NxLimit = 57: LtLimit = 61: Entering = 5: Nam = "P2": NumPlyNow = P2out
       Case "P3": Limit = 25: NxLimit = 62: LtLimit = 66: Entering = 18: Nam = "P3": NumPlyNow = P3out
       Case "P4": Limit = 38: NxLimit = 67: LtLimit = 71: Entering = 31: Nam = "P4": NumPlyNow = P4out
       Case Else: MsgBox ("Impossible Error")
End Select
Exceed = False
 'If One = True Then
    If Index >= Entering And Index < Limit Then
    inde = Index
       For k = 1 To Rand
       inde = inde + 1
       If inde = Limit Then inde = NxLimit
       If inde > LtLimit + 1 Then GoTo 3
       Next
       Exit Sub
   End If
'End If

If Index >= NxLimit Then
    If Rand + Index > LtLimit + 1 Then
3    valu = LtLimit - Index + 1
       Label3.Caption = "Sorry,You'll Have To Throw A " & valu & " To Go Home"
       Exceed = True
      ' If NumPlyNow = 1 Then TransSlate.Enabled = True ' only one seed for the present player is out and cannot move transfer dice slate
     End If
End If
       
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

Private Sub Command1_Click()
On Error Resume Next  'for that  .setfocus; sometimes gives error '340
If Running = True Then Label2.Caption = "Please Wait...": Exit Sub
If tmrDice.Enabled = True Then   'the dice simulation has already began end it
DiceRolling = False
tmrDice.Enabled = False
Command1.Picture = Picture1(1).Picture
DiceRowed = True
Label3.ForeColor = Label5.BackColor
Clik = 2
Nclik = 0
tmrSpeakOut.Interval = 1
tmrSpeakOut = True
If Rand1 = 6 Or Rand2 = 6 Then Label3.Caption = "You Can Move Out A Seed"
CheckPlayers
Else    'start dice simulation
DiceRolling = True
DiceRowed = False
tmrDice.Enabled = True
Command1.Picture = Picture1(0).Picture
Label3.Caption = "Click Again"
Label3.ForeColor = vbWhite
Clik = 1
End If
Me.SetFocus
If Clik = 2 Then
Command1.Enabled = False
End If
End Sub
Private Static Sub tmrSpeakOut_Timer()
If Vibrating Then Exit Sub
If Paused Then Exit Sub

Dim Numtimes As Integer, X As Integer
'just to speakout the score from the dice,read the greater value first
    Numtimes = Numtimes + 1
    If Numtimes = 1 Then 'goes through this if then once
        If Rand1 > Rand2 Then
           X = Rand1 'speak first the greater value
        Else
           X = Rand2
         End If
         
        tmrSpeakOut.Interval = 1300
    Else  'this is the 2nd pass thru the code
        If Rand1 < Rand2 Then
           X = Rand1
        Else
           X = Rand2
         End If
    
        Numtimes = 0
        tmrSpeakOut = False
    End If
    
    
    Select Case X
       Case 1
            wmpTalkDie.URL = Dirxtry & "1.wav"
       Case 2
            wmpTalkDie.URL = Dirxtry & "2.wav"
       Case 3
            wmpTalkDie.URL = Dirxtry & "3.wav"
       Case 4
            wmpTalkDie.URL = Dirxtry & "4.wav"
       Case 5
            wmpTalkDie.URL = Dirxtry & "5.wav"
       Case 6
            wmpTalkDie.URL = Dirxtry & "6.wav"
    End Select

End Sub
Public Sub CheckPlayers()
Dim Inn As Boolean, CannotMove As Integer
Dim JJ As Integer, H As Integer
'we have to know who is the current player
'if he has any seed out and how many,can he move with this throw?
'we also have to know the total num of players ...pnum

'this could have been a very long code but i compacted it--could have been 4times as lenti
NumOut = 0
CannotMove = 0
One = False
FstDie = False
SndDie = False

Select Case Rotate
   Case 1   'player1's turn to play
        Turn = "P1"
        JJ = 0     'let's see if any seed is inside
        For H = 0 To 3
        If Board.Ply1(H).Tag = 0 Then JJ = 1
        Next
   Case 2   'player2's turn to play
        Turn = "P2"
        JJ = 0     'let's see if any seed is inside
        For H = 0 To 3
        If Ply2(H).Tag = 0 Then JJ = 1
        Next
    Case 3   'player3's turn to play
        Turn = "P3"
        JJ = 0     'let's see if any seed is inside
        For H = 0 To 3
        If Ply3(H).Tag = 0 Then JJ = 1
        Next
     Case 4   'player4's turn to play
        Turn = "P4"
        JJ = 0     'let's see if any seed is inside
        For H = 0 To 3
        If Ply4(H).Tag = 0 Then JJ = 1
        Next
   End Select
                   
             For J = 0 To 71
              If BG(J).Tag = Turn Then '
                 NumOut = NumOut + bgNum(J)  'didn't want to use 'one' bcos we could have more then one player on a position
               End If
             Next
           FreeThrow = False
        If Rand1 = 6 And Rand2 = 6 Then 'in the Nigerian variation, which is being used to write this program,if a player get's a 6 in the 2 dice he is given a bonus to throw the dice again,if he get's the same value in the bonus throw(i.e 6,6),the bonus throw continues-as long as he keeps on geting a 6,6
            FreeThrow = True
         End If
                                                        
    If NumOut = 0 Then   'no seed for this player is on the field
              ' If JJ = 0 Then 'no player is out no player is in .'.all seeds have completed there course for this player,transfer slate to the next player
               '   TransSlate.Enabled = True
                '  Exit Sub
                'End If
         If Rand1 = 6 Then 'let the present rand be = to the die that has a 6
            Rand = Rand1
            RandNx = Rand2
        ElseIf Rand2 = 6 Then
            Rand = Rand2
            RandNx = Rand1 'randnx will be the next value of rand
        Else    'the 2 dice does not have a 6 and therefore no seed can come out
            Label3.Caption = "You Need A Six To Move Out A Seed"
            StillPlaying = False
           TransSlate.Enabled = True
             Exit Sub
       End If
                   FstDie = True  'necessary for AI
                   SndDie = True
     Else 'one or more player is out let's see if they can actually move
                   
       If LastStep = True Then '  'only one die is used when all the seeds are in private route else hooked at d last step
           Rand2 = 0
           If Rand1 = 6 Then FreeThrow = True  'only one die,if the player get's a 6 then a bonus since it cannot move with it as long t is within it private route
        End If
      If NumOut = 1 Then 'only one player is out,to avoid killing a players seed with just one die,having no other player to move the 2nd die,causing the slate not to move thereby the program "hangs up", add the two die's value
               
                
           If Rand1 = 6 Or Rand2 = 6 Then  'to give the player a chance of bringing out a seed
                   If JJ = 1 Then  ' seed(s) is/are inside, see the select case structure above
                   FstDie = True
                   SndDie = True
                   GoTo 2
                   End If
               End If
                   Rand = Rand1 + Rand2
                   One = True    'if one is true nclik will be assigned a val of 2 in tmrmove
      End If
           FstDie = False
           SndDie = False
           
          
          For J = 0 To 71
              If BG(J).Tag = Turn Then '
              Rand = Rand1 + Rand2
              Distance J  'check if this seed can move with the sum of both dice
                    If Exceed = True Then
                    If One = True Or LastStep = True Then GoTo 2 'only one seed is out and it can't be moved by the dice,transfer slate
                     Rand = Rand1
                     Distance J  'check if this seed can move with d first die
                        If Exceed = True Then  'this seed can't move with either the sum or just d first die
                            Rand = Rand2
                            Distance J  'check if this seed can move with the 2nd die
                            If Exceed = True Then 'this seed can't be moved by this throw
                                 CannotMove = CannotMove + bgNum(J)  '---not used anyway
                             Else  'this seed can move with the 2nd die,find another seed that can move with the 1st die
                                  SndDie = True 'found a seed that can move with d 2nd die
                            End If
                         Else  'this seed can move with d first die,we have to find another seed that can move with the 2nd die
                         FstDie = True
                         If bgNum(J) > 1 And Rand1 = Rand2 Then GoTo 1  'there are 2 or more seeds on a position and the dice have d same value,it tested for the first seed and it can move then the other seeds can also move as far d 2 dice have the same value
                         End If
                   Else 'this seed can move with the sum of the two
1                   FstDie = True
                   SndDie = True
                   GoTo 2 ' to save computing time
                 End If
              End If
          Next 'at the end of all this fstdie and snddie must be true for the seed(s) to be able to move

2          If (FstDie = False) Or (SndDie = False) Then
          Label3.Caption = "None Of Your Seed Can Move With This Throw"
          TransSlate.Enabled = True
          Else   'seed(s) can move
             Label3.Caption = "Make A Move"
        If One = True Then Rand = Rand1 + Rand2: Exit Sub 'if only one player then it has to be the sum of the two die
          Rand = Rand1   'reassign the values to work with
          RandNx = Rand2
          End If
 End If
 
End Sub

Private Sub Form_Load()
On Error Resume Next 'for    .setfocus
'Leftt = Screen.Height / 2.5  'this two variables are used for the vibration
'Topp = Screen.Height / 12

'Me.Left = Leftt
'Me.Top = Topp

Initialise

FirstToFinish = ""
throwdiceAI.Enabled = False
tmrVibrate.Enabled = False
wmpTalkDie.Visible = False
Label7.Visible = False
Label7.Caption = ""
Label7.Move 450, 480, 8865, 8775 'covers the board to avoid user interactn with the object when pausdd
Label7.BackStyle = 0 'transparent
Me.Width = 9800
'Me.Height = 10160
For k = 0 To 71
    BG(k).Picture = LoadPicture("")
    BG(k).Tag = ""
    bgNum(k) = 0
Next
For F = 0 To 3   'they are all in initially (seeds)
    Line1(F).Visible = False
    Ply1(F).Tag = 0
    Ply2(F).Tag = 0
    Ply3(F).Tag = 0
    Ply4(F).Tag = 0
Next
Ply1Hom = 0
Ply2Hom = 0
Ply3Hom = 0
Ply4Hom = 0
tmrMoved.Interval = 400
'DiceRowed = False
tmrDice.Enabled = False
Die(0).Picture = Dice(6).Picture
Die(1).Picture = Dice(6).Picture
Frame1.Move P1Left, P12Top
Nclik = 0
P1out = 0
P2out = 0
P3out = 0
P4out = 0
moveSeedAI.Enabled = False
Out = False
PauseStart.SetFocus
DiceRowed = False
One = False
StillPlaying = False
Running = False
ClickIsFromComputer = False
End Sub
Sub Checkers()
 Dim Msgbx As Integer
 Select Case BG(Last).Tag   'let's see what is on the target positn
       Case ""    'no seed on the target position
                 Select Case Identity  'this not essential but sometimes clickin on a moving seed can cause the picture on it to disappear but the seed is still on the board
                        Case "P1"
                        BG(Last).Picture = Ply1_1.Picture
                        Case "P2"
                        BG(Last).Picture = Ply2_1.Picture
                        Case "P3"
                        BG(Last).Picture = Ply3_1.Picture
                        Case "P4"
                        BG(Last).Picture = Ply4_1.Picture
                  End Select
                 BG(Last).Tag = Trim(Identity)   'exchange identity
                 bgNum(Last) = 1       'this now has one seed on it                        -next position
       Case "P1" 'target position already has player1 seed on it
                    Select Case Identity   'see which players seed is movin i.e @ d initial position
                           Case "P1" 'well, they're d same, no kills,simply top up
                                 Select Case bgNum(Last)  'check how many of player1 seed is already there at the target position
                                        Case 1
                                         bgNum(Last) = 2  'now make it two
                                        BG(Last).Picture = Ply1_2.Picture 'update pic
                                        Case 2
                                        bgNum(Last) = 3  'now make it three
                                        BG(Last).Picture = Ply1_3.Picture 'update pic
                                        Case 3
                                        bgNum(Last) = 4  'now make it four
                                        BG(Last).Picture = Ply1_4.Picture 'update pic
                                       ' Case Else: Beep
                                End Select

                          Case "P2", "P3", "P4" 'Player1 seed has been Captured  by either player2,player3 or player4,this is the moving seed
                                     
                                     Select Case Identity
                                            Case "P2"
                                                  Hunter = Pnam2
                                                  'msgbx = MsgBox(Pnam2 & " Has Captured " & Pnam1 & "'s seed", vbInformation)
                                                  'Home2
                                            Case "P3"
                                                  Hunter = Pnam3
                                                 ' msgbx = MsgBox(Pnam3 & " Has Captured " & Pnam1 & "'s seed", vbInformation)
                                                  'home3
                                            Case "P4"
                                                 Hunter = Pnam4
                                                 'msgbx = MsgBox(Pnam4 & " Has Captured " & Pnam1 & "'s seed", vbInformation)
                                                  'home4
                                      End Select
                                      Hunted = Pnam1   'vibrate and show a msgbox after vibration
                                      tmrVibrate.Enabled = True 'the valu of hunter is needed
                                             
                                 Select Case bgNum(Last)  'let's see how many of player1 seed is on it
                                       Case 1
                                             BG(Last).Picture = LoadPicture("")
                                             bgNum(Last) = 0
                                             BG(Last).Tag = ""
                                       Case 2
                                            BG(Last).Picture = Ply1_1.Picture
                                            bgNum(Last) = 1
                                       Case 3
                                           BG(Last).Picture = Ply1_2.Picture
                                           bgNum(Last) = 2
                                       Case 4
                                          BG(Last).Picture = Ply1_3.Picture
                                          bgNum(Last) = 3
                                 End Select
                                                                       
                                          For J = 0 To 3 'to return player1 back to it's prison
                                             If Ply1(J).Tag = 1 Then Exit For  'to find an empty imagebox to put it in
                                          Next
                                             'Ply1(j).Picture = Ply1_1.Picture
                                             Ply1(J).Visible = True
                                             Ply1(J).Tag = 0  'not yet out
                                          
                        End Select
                        
      Case "P2" 'target position already has player2 seed on it
  
                    Select Case Identity   'see which players seed is movin i.e @ d initial position
                           Case "P2" 'well, they're d same, no kills,simply top up
                           Select Case bgNum(Last)  'check how many of player2 seed is already there at the target position
                                        Case 1
                                         bgNum(Last) = 2  'now make it two
                                        BG(Last).Picture = Ply2_2.Picture 'update pic
                                        Case 2
                                        bgNum(Last) = 3  'now make it three
                                        BG(Last).Picture = Ply2_3.Picture 'update pic
                                        Case 3
                                        bgNum(Last) = 4  'now make it four
                                        BG(Last).Picture = Ply2_4.Picture 'update pic
                                        'Case Else: Beep
                                End Select

                          Case "P1", "P3", "P4" 'Player2 seed has been Captured  by either player1,player3 or player4,these are the moving seed
                                     Select Case Identity
                                            Case "P1"
                                                  Hunter = Pnam1
                                                  'msgbx = MsgBox(Pnam1 & " Has Captured " & Pnam2 & "'s seed", vbInformation)
                                                  'Home1
                                            Case "P3"
                                                  Hunter = Pnam3
                                                  'msgbx = MsgBox(Pnam3 & " Has Captured " & Pnam2 & "'s seed", vbInformation)
                                                  'Home3
                                            Case "P4"
                                                  'msgbx = MsgBox(Pnam4 & " Has Captured " & Pnam2 & "'s seed", vbInformation)
                                                  Hunter = Pnam4
                                                  'Home4
                                      End Select
                                          
                                          Hunted = Pnam2
                                          tmrVibrate.Enabled = True
                                             
                                 Select Case bgNum(Last)  'let's see how many of player2 seed is on it
                                       Case 1
                                             BG(Last).Picture = LoadPicture("")
                                             bgNum(Last) = 0
                                             BG(Last).Tag = ""
                                       Case 2
                                            BG(Last).Picture = Ply2_1.Picture
                                            bgNum(Last) = 1
                                       Case 3
                                           BG(Last).Picture = Ply2_2.Picture
                                           bgNum(Last) = 2
                                       Case 4
                                          BG(Last).Picture = Ply2_3.Picture
                                          bgNum(Last) = 3
                                 End Select
                                  
                                          For J = 0 To 3 'to return player2 back to it's prison
                                             If Ply2(J).Tag = 1 Then Exit For  'to find an empty imagebox to put it in
                                          Next
                                             'Ply2(j).Picture = Ply2_1.Picture
                                             Ply2(J).Visible = True
                                             Ply2(J).Tag = 0  'not yet out
                        
                              End Select
                                                  
          Case "P3" 'target position already has player3 seed on it
  
                    Select Case Identity   'see which players seed is movin i.e @ d initial position
                           Case "P3" 'well, they're d same, no kills,simply top up
                           Select Case bgNum(Last)  'check how many of player3 seed is already there at the target position
                                        Case 1
                                         bgNum(Last) = 2  'now make it two
                                        BG(Last).Picture = Ply3_2.Picture 'update pic
                                        Case 2
                                        bgNum(Last) = 3  'now make it three
                                        BG(Last).Picture = Ply3_3.Picture 'update pic
                                        Case 3
                                        bgNum(Last) = 4  'now make it four
                                        BG(Last).Picture = Ply3_4.Picture 'update pic
                                        'Case Else: Beep
                                End Select

                          Case "P1", "P2", "P4" 'Player3 seed has been Captured  by either player1,player2 or player4,these are the moving seed
                                     Select Case Identity
                                            Case "P1"
                                                  'msgbx = MsgBox(Pnam1 & " Has Captured " & Pnam3 & "'s seed", vbInformation)
                                                  Hunter = Pnam1
                                                  'Home1
                                            Case "P2"
                                                  'msgbx = MsgBox(Pnam2 & " Has Captured " & Pnam3 & "'s seed", vbInformation)
                                                  Hunter = Pnam2
                                                  'Home2
                                            Case "P4"
                                                  'msgbx = MsgBox(Pnam4 & " Has Captured " & Pnam3 & "'s seed", vbInformation)
                                                  Hunter = Pnam4
                                                  'Home4
                                      End Select
                                           
                                           Hunted = Pnam3
                                           tmrVibrate.Enabled = True
                                             
                                 Select Case bgNum(Last)  'let's see how many of player3 seed is on it
                                       Case 1
                                             BG(Last).Picture = LoadPicture("")
                                             bgNum(Last) = 0
                                             BG(Last).Tag = ""
                                       Case 2
                                            BG(Last).Picture = Ply3_1.Picture
                                            bgNum(Last) = 1
                                       Case 3
                                           BG(Last).Picture = Ply3_2.Picture
                                           bgNum(Last) = 2
                                       Case 4
                                          BG(Last).Picture = Ply3_3.Picture
                                          bgNum(Last) = 3
                                 End Select
                                 
                                          For J = 0 To 3 'to return player3 back to it's prison
                                             If Ply3(J).Tag = 1 Then Exit For  'to find an empty imagebox to put it in
                                            Next
                                            ' Ply3(j).Picture = Ply3_1.Picture
                                             Ply3(J).Visible = True
                                             Ply3(J).Tag = 0  'not yet out
                                  
                        
                              End Select
 
                

                  
     Case "P4" 'target position already has player4 seed on it
  
                    Select Case Identity   'see which players seed is movin i.e @ d initial position
                           Case "P4" 'well, they're d same, no kills,simply top up
                           Select Case bgNum(Last)  'check how many of player4 seed is already there at the target position
                                        Case 1
                                         bgNum(Last) = 2  'now make it two
                                        BG(Last).Picture = Ply4_2.Picture 'update pic
                                        Case 2
                                        bgNum(Last) = 3  'now make it three
                                        BG(Last).Picture = Ply4_3.Picture 'update pic
                                        Case 3
                                        bgNum(Last) = 4  'now make it four
                                        BG(Last).Picture = Ply4_4.Picture 'update pic
                                        'Case Else: Beep
                                End Select

                          Case "P1", "P3", "P2" 'Player4 seed has been Captured  by either player1,player3 or player2,these are the moving seed
                                     Select Case Identity
                                            Case "P1"
                                                  'msgbx = MsgBox(Pnam1 & " Has Captured " & Pnam4 & "'s seed", vbInformation)
                                                  Hunter = Pnam1
                                                  'Home1
                                            Case "P3"
                                                  'msgbx = MsgBox(Pnam3 & " Has Captured " & Pnam4 & "'s seed", vbInformation)
                                                  Hunter = Pnam3
                                                  'Home3
                                            Case "P2"
                                                  'msgbx = MsgBox(Pnam2 & " Has Captured " & Pnam4 & "'s seed", vbInformation)
                                                  Hunter = Pnam2
                                                  'Home2
                                      End Select
                                          
                                          Hunted = Pnam4
                                         tmrVibrate.Enabled = True
                                             
                                 Select Case bgNum(Last)  'let's see how many of player2 seed is on it
                                       Case 1
                                             BG(Last).Picture = LoadPicture("")
                                             bgNum(Last) = 0
                                             BG(Last).Tag = ""
                                       Case 2
                                            BG(Last).Picture = Ply4_1.Picture
                                            bgNum(Last) = 1
                                       Case 3
                                           BG(Last).Picture = Ply4_2.Picture
                                           bgNum(Last) = 2
                                       Case 4
                                          BG(Last).Picture = Ply4_3.Picture
                                          bgNum(Last) = 3
                                 End Select
                                      
                                      
                                 
                                          For J = 0 To 3 'to return player2 back to it's prison
                                             If Ply4(J).Tag = 1 Then Exit For  'to find an empty imagebox to put it in
                                            Next
                                            ' Ply4(j).Picture = Ply4_1.Picture
                                             Ply4(J).Visible = True
                                             Ply4(J).Tag = 0  'not yet out
                                  
                        
                              End Select
                 
     End Select
End Sub




Private Sub Ply1_Click(Index As Integer)
tmrToolSlider.Interval = 1: Slider = False
If DiceRowed = True Then 'dice must have been rowed
       If Turn <> "P1" Then Label2.Caption = "Not Turn Yet, Click On the Right Seed": Exit Sub
      If PlayerType = 11 Then
     If Not ClickIsFromComputer Then
        Label2.Caption = "That's Computer's Seed, Forbear!"
        Exit Sub
     End If
 End If

      If PlayerType <> Ptyp1 Then Beep: Exit Sub

        If Rand2 = 6 And Nclik <> 1 Then  'so that it will automatically use the die with a 6 without the user specifying ,this should happen only at first clik otherwise if the user cliks d 2nd time he'll get another 6
        Rand = 6
        RandNx = Rand1
        End If
  
  If One = True And Rand = 6 Then Label2.Caption = "Can't Move This Seed With A Throw of" & Str(Rand1) & Str(Rand2): Exit Sub

     If Ply1(Index).Tag = 1 Then 'already out
     Beep
       MsgBox ("Error: ply1().tag=1,supposed to be zero")   'bg_click will take care of this
     Else    'not yet out

        If Rand = 6 Then   'can only come out with a throw of six
         wmpFreedSound.URL = Dirxtry & "button1.ogg"
        StillPlaying = True 'will be made false when the 2nd seed has finished has moving

         Nclik = Nclik + 1
         
              If Rand1 = 6 Then
           Line1(0).Visible = True
           Line1(1).Visible = True
           ElseIf Rand2 = 6 Then
           Line1(2).Visible = True
           Line1(3).Visible = True
           End If
           
     
     If Nclik = 2 Then
            For H = 0 To 3
                Line1(H).Visible = True
            Next
        End If
        
           If FreeThrow Then GoTo 9 'for two 6, it should not automatically move the seed

         Out = False
         For J = 0 To 71  'let's see if any seed is out at all,if a player gets a throw of say a 6 and a 4 and another player's seed is at it 'door step' it kills with a 6 but what happens with the 4 ? this code solves it,
         If BG(J).Tag = "P1" Then Out = True  ' (contn) by not allowing it to settle at it's step if this is the case
         Next
         If Out = False Then 'no seed is out
         YesNo = False
         HomeNum = 0
           If bgNum(HomeNum) <> 0 Then 'let'see if there is any seed there, store it b4 d simulation
              Motion.Picture = BG(HomeNum).Picture
              Motion.Tag = BG(HomeNum).Tag
              Temp = bgNum(HomeNum)
              YesNo = True
           End If
           
          BG(HomeNum).Picture = Ply1_1.Picture
        Ply1(Index).Tag = 1 'Now it is out\
          Ply1(Index).Visible = False   'remove from prison
        BG(HomeNum).Tag = "P1"  'put player1's seed on it
        bgNum(HomeNum) = 1 'one player is on it but on motion
        Nclik = 2
        Rand = RandNx  'already out with 6 use the next die
        BG_Click HomeNum  'move it
        DiceRowed = False
        Exit Sub
        End If
           Hunter = Pnam1
9         Ply1(Index).Tag = 1 'Now it is out\
          
         
            Select Case BG(0).Tag  'to see if a player's seed has already occupied this positn
            
            Case "P1"   'already occupied by player1
             Ply1(Index).Visible = False   'remove seed from prison
                 Select Case bgNum(0) 'check how many of player1 seed is on it
                  Case 1      'increment the num of players & also the pic
                  bgNum(0) = 2
                  BG(0).Picture = Ply1_2.Picture
                  Case 2
                  bgNum(0) = 3
                  BG(0).Picture = Ply1_3.Picture
                  Case 3
                  bgNum(0) = 4
                  BG(0).Picture = Ply1_4.Picture
                  Case 4: MsgBox "impossible error occured" 'impossible or error
                  End Select
                  
            Case "P2"  'already occupied by player2
            
                     Hunted = Pnam2
                     tmrVibrate.Enabled = True
                         For J = 0 To 3 'to return player2 back to it's prison
                         If Ply2(J).Visible = False Then Exit For 'to find an empty imagebox to put it in
                         Next
                        ' Ply2(j).Picture = Ply2_1.Picture
                         Ply2(J).Visible = True
                         Ply2(J).Tag = 0  'not yet out
                         
                         Ply1(Index).Visible = False
                         Ply1(Index).Tag = 3 'it goes home for capturing an enemy
                         'Home1
                         
                  Select Case bgNum(0)   'how many of player2 seed are on it
                         
                         Case 1 ' only one, Captured  by player1
                    
                             BG(0).Picture = LoadPicture("") 'no seed will appear  on it
                             bgNum(0) = 0                     'num  of seeds is zero
                             BG(0).Tag = ""   'no player's seed is on it
                         
                         Case 2 ' only two, one is  Captured  by player1
                         
                         BG(0).Picture = Ply2_1.Picture ' one has been killed it remains one
                         bgNum(0) = 1
                         'bg(0).tag is stil  occupied by "p2"
            
                          Case 3 ' only three, one is  Captured  by player1
                         
                         BG(0).Picture = Ply2_2.Picture ' one has been killed it remains two
                         bgNum(0) = 2
                          'bg(0).tag is stil  occupied by "p2"
                         
                          Case 4 ' only four, one is  Captured  by player1
                         
                         BG(0).Picture = Ply2_3.Picture ' one has been killed it remains three
                         bgNum(0) = 3
                          'bg(0).tag is stil  occupied by "p2"
                   End Select
'                   f = MsgBox(Pnam1 & " Has Captured " & Pnam2 & "'s seed", vbInformation)
            Case "P3"  'already occupied by player3
                        
                     Hunted = Pnam3
                     tmrVibrate.Enabled = True
                         For J = 0 To 3 'to return player3 back to it's prison
                         If Ply3(J).Visible = False Then Exit For 'to find an empty imagebox to put it in
                         Next
                         'Ply3(j).Picture = Ply3_1.Picture
                         Ply3(J).Visible = True
                         Ply3(J).Tag = 0  'not yet out
                         
                         Ply1(Index).Visible = False
                         Ply1(Index).Tag = 3 'it goes home for capturing an enemy
                         'Home1
                         
                     Select Case bgNum(0)   'how many of player3 seed are on it
                         
                         Case 1 ' only one, Captured  by player1
                       
                         BG(0).Picture = LoadPicture("") 'no seed will appear  on it
                         bgNum(0) = 0                     'num  of seeds is zero
                         BG(0).Tag = ""   'no player's seed is on it
                        
                         Case 2 ' only two, one is  Captured  by player1
                        
                         BG(0).Picture = Ply3_1.Picture ' one has been killed it remains one
                         bgNum(0) = 1
                
                          Case 3 ' only three, one is  Captured  by player1
                         
                         BG(0).Picture = Ply3_2.Picture ' one has been killed it remains two
                         bgNum(0) = 2
                        
                          Case 4 ' only four, one is  Captured  by player1
                         
                         BG(0).Picture = Ply3_3.Picture ' one has been killed it remains three
                         bgNum(0) = 3
                        
                 End Select
               '  f = MsgBox(Pnam1 & " Has Captured " & Pnam3 & "'s seed", vbInformation)
            
            Case "P4"        'already occupied by player4
                     Hunted = Pnam4
                     tmrVibrate.Enabled = True
                         For J = 0 To 3 'to return player3 back to it's prison
                         If Ply4(J).Visible = False Then Exit For 'to find an empty imagebox to put it in
                         Next
                         'Ply4(j).Picture = Ply4_1.Picture
                         Ply4(J).Visible = True
                         Ply4(J).Tag = 0  'not yet out
                         
                         Ply1(Index).Visible = False
                         Ply1(Index).Tag = 3 'it goes home for capturing an enemy
                         'Home1
                         
                     Select Case bgNum(0)   'how many of player3 seed are on it
                         
                         Case 1 ' only one, Captured  by player1
                       
                         BG(0).Picture = LoadPicture("") 'no seed will appear  on it
                         bgNum(0) = 0                     'num  of seeds is zero
                         BG(0).Tag = ""   'no player's seed is on it
                        
                         Case 2 ' only two, one is  Captured  by player1
                        
                         BG(0).Picture = Ply4_1.Picture ' one has been Captured it remains one
                         bgNum(0) = 1
                
                          Case 3 ' only three, one is  Captured  by player1
                         
                         BG(0).Picture = Ply4_2.Picture ' one has been Captured it remains two
                         bgNum(0) = 2
                        
                          Case 4 ' only four, one is  Captured  by player1
                         
                         BG(0).Picture = Ply4_3.Picture ' one has been killed it remains three
                         bgNum(0) = 3
                        
                 End Select
             'f = MsgBox(Pnam1 & " Has Captured " & Pnam4 & "'s seed", vbInformation)
            Case Else   'no player's seed is on it
            
            BG(0).Tag = "P1"  'put player1's seed on it
            BG(0).Picture = Ply1_1.Picture
            bgNum(0) = 1 'one player is on it
            Ply1(Index).Visible = False   'remove from prison
            
            End Select
         Else 'rand <> 6
         Label2.Caption = "Can't Move This Seed With A Throw of" & Str(Rand)
         Exit Sub
        End If
      End If
    If Nclik = 1 Then
        Rand = RandNx
        DiceRowed = True
    Else
        StillPlaying = False
        TransSlate.Enabled = True 'transfer dice slate to the next player
        DiceRowed = False
        Nclik = 0
        Rand1 = 0
        Rand2 = 0
  End If

Else   'user has not rowed d dice
Label2.Caption = "Please Row The Dice"
End If

End Sub
Public Sub Home1()
 Ply1Hom = Ply1Hom + 1
                   Nam = Pnam1
                         Select Case Ply1Hom
                                Case 1
                                    Ply1Home.Picture = Ply1_1.Picture
                                Case 2
                                    Ply1Home.Picture = Ply1_2.Picture
                                Case 3
                                    Ply1Home.Picture = Ply1_3.Picture
                                Case 4
                                    If Not Paused Then 'if the game is not already paused pause it, this is essential,cos when d form loads it is paused
                                        PauseStart.Value = True 'pause game
                                    End If
    
                                    F = MsgBox(Pnam1 & " Has completed it's course", vbInformation) ' when this msgbox is being displayed the game computer can still play,in the .exe this usually occur don't know what caused it,i expect all computation to halt while a msgbox is being displayed.
                                    
                                    If Paused Then  'if the game is not already paused pause it, this is essential,cos when d form loads it is paused
                                        PauseStart.Value = True 'resume game
                                    End If
                                    
                                    Ply1Home.Picture = Ply1_4.Picture
                                    Finished = Finished + 1
                                    Finish
                                
                         End Select
End Sub

Public Sub Home2()
 Ply2Hom = Ply2Hom + 1
  Nam = Pnam2
                         Select Case Ply2Hom
                                Case 1
                                Ply2Home.Picture = Ply2_1.Picture
                                Case 2
                                Ply2Home.Picture = Ply2_2.Picture
                                Case 3
                                Ply2Home.Picture = Ply2_3.Picture
                                Case 4
                                    If Not Paused Then 'if the game is not already paused pause it, this is essential,cos when d form loads it is paused
                                        PauseStart.Value = True 'pause game
                                    End If
                                
                                F = MsgBox(Pnam2 & " Has completed it's course", vbInformation)
                                    If Paused Then  'if the game is not already paused pause it, this is essential,cos when d form loads it is paused
                                        PauseStart.Value = True 'resume game
                                    End If
                                
                                Ply2Home.Picture = Ply2_4.Picture
                                Finished = Finished + 1
                                Finish
                         End Select
End Sub

Public Sub Home3()
 Ply3Hom = Ply3Hom + 1
  Nam = Pnam3
  
                         Select Case Ply3Hom
                                Case 1
                                Ply3Home.Picture = Ply3_1.Picture
                                Case 2
                                Ply3Home.Picture = Ply3_2.Picture
                                Case 3
                                Ply3Home.Picture = Ply3_3.Picture
                                Case 4
                                    If Not Paused Then 'if the game is not already paused pause it, this is essential,cos when d form loads it is paused
                                        PauseStart.Value = True 'pause game
                                    End If
                                
                               F = MsgBox(Pnam3 & " Has completed it's course", vbInformation)
                                    If Paused Then  'if the game is not already paused pause it, this is essential,cos when d form loads it is paused
                                        PauseStart.Value = True 'resume game
                                    End If
                                
                                Ply3Home.Picture = Ply3_4.Picture
                                Finished = Finished + 1
                                Finish
                         End Select
End Sub


Public Sub Home4()
 Ply4Hom = Ply4Hom + 1
 Nam = Pnam4
                         Select Case Ply4Hom
                                Case 1
                                Ply4Home.Picture = Ply4_1.Picture
                                Case 2
                                Ply4Home.Picture = Ply4_2.Picture
                                Case 3
                                Ply4Home.Picture = Ply4_3.Picture
                                Case 4
                                    If Not Paused Then 'if the game is not already paused pause it, this is essential,cos when d form loads it is paused
                                        PauseStart.Value = True 'pause game
                                    End If
                                
                             F = MsgBox(Pnam4 & " Has completed it's course", vbInformation)
                                    
                                    If Paused Then  'if the game is not already paused pause it, this is essential,cos when d form loads it is paused
                                        PauseStart.Value = True 'resume game
                                    End If
                             
                                Ply4Home.Picture = Ply4_4.Picture
                                Finished = Finished + 1
                                Finish
                         End Select
End Sub



Sub Finish()
Pone = False 'i'm not not taking chance
Ptwo = False
Pthree = False
Pfour = False
Dim Indkx As Integer
If Pnum = 3 Then P4out = 999
Select Case PnumX     ' pnum is the num of players
       Case 2           'two players
       
            Select Case Finished
                        Case 1  '--select case struct not necessary but safer
                            Finished = 0
                            If Not Paused Then 'if the game is not already paused pause it, this is essential,cos when d form loads it is paused
                                PauseStart.Value = True 'pause game
                            End If
                            'the timer still acts despite the showing of a modal form,manually pause it.
                             F = MsgBox(Nam & " Has Won This Game", vbInformation)
                             Wins.Label4.Caption = Nam
                            Wins.wmpHailWinner.URL = Dirxtry & "FM_crowd_trick.wav"
                            Wins.ShowWinner
                            Wins.Show vbModal
                            wmpTalkDie.Close
                            
            End Select
            
       Case 3   'we'll need only two winners to contest for the crown
          
      Select Case Finished
                       Case 1
                            ArrayOfWinners(0) = Nam
                            Select Case Nam
                               Case Pnam1   'Player1 is the first winner
                               Label4(0).Caption = "1st " & Pnam1
                               P1out = 1
                               Case Pnam2
                                Label4(1).Caption = "1st " & Pnam2
                                P2out = 1
                              Case Pnam3
                                Label4(2).Caption = "1st " & Pnam3
                                P3out = 1
                               Case Pnam4
                                 Label4(3).Caption = "1st " & Pnam4
                                 P4out = 1
                              End Select
                              FirstToFinish = Nam
                              Wins.Label2(0).Caption = "1ST        " & Nam
                        Case 2
                           ArrayOfWinners(1) = Nam
                            Select Case Nam
                               Case Pnam1   'Player1 is the 2nd winner
                               Label4(0).Caption = "2nd " & Pnam1
                               P1out = 2
                               Case Pnam2
                                Label4(1).Caption = "2nd " & Pnam2
                                P2out = 2
                              Case Pnam3
                                Label4(2).Caption = "2nd " & Pnam3
                                P3out = 2
                               Case Pnam4
                                 Label4(3).Caption = "2nd " & Pnam4
                                P4out = 2
                            End Select
                               
                             
                             With Wins
                                    .Label2(1).Caption = "2ND       " & Nam
                                    .Image1(0).Visible = False
                                    .Image1(1).Visible = False
                                    .Label2(2).Visible = False
                                    .Label2(4).Visible = False
                                    For Indkx = 0 To 1
                                        Select Case ArrayOfWinners(Indkx)
                                               Case Pnam1: Pone = True
                                               Case Pnam2: Ptwo = True
                                               Case Pnam3: Pthree = True
                                               Case Pnam4: Pfour = True
                                        End Select
                                    Next  'two will remain false
                                    
                                    If Pnum = 4 Then
                                                                           
                                        If Pone = False And Pnam1 <> ArrayOfWinners(3) Then
                                            ArrayOfWinners(2) = Pnam1
                                        ElseIf Ptwo = False And Pnam2 <> ArrayOfWinners(3) Then
                                            ArrayOfWinners(2) = Pnam2
                                        ElseIf Pthree = False And Pnam3 <> ArrayOfWinners(3) Then
                                            ArrayOfWinners(2) = Pnam3
                                        ElseIf Pfour = False And Pnam4 <> ArrayOfWinners(3) Then
                                            ArrayOfWinners(2) = Pnam4
                                        End If
                                        Nam = ArrayOfWinners(2)
                                        
                                   ElseIf Pnum = 3 Then
                                    
                                        If Pone = False Then
                                            ArrayOfWinners(2) = Pnam1
                                        ElseIf Ptwo = False Then
                                            ArrayOfWinners(2) = Pnam2
                                        ElseIf Pthree = False Then
                                            ArrayOfWinners(2) = Pnam3
                                        End If
                                        Nam = ArrayOfWinners(2)
                                  Else
                                    Msgbx = MsgBox("Pnum Has An Illegal Value in sub Finish,Case 3 value= " & Pnum, vbCritical): Exit Sub
                                  End If
                                    
                                    
'    If P1out = 0 And Pnam1 <> Loozer Then
'    Nam = Pnam1
'    ElseIf P2out = 0 And Pnam2 <> Loozer Then
'    Nam = Pnam2
'    ElseIf P3out = 0 And Pnam3 <> Loozer Then
'    Nam = Pnam3
'    ElseIf P4out = 0 And Pnam4 <> Loozer Then
'    Nam = Pnam4
'    Else
'    'MsgBox ("Impossible Error @ Winning,occured")
'    End If
                                    If Pnum = 4 Then
                                    .Label2(4).Caption = "4th        " & ArrayOfWinners(3)  'loosername
                                    .Label2(4).Visible = True
                                    End If
                                    .Label2(3).Caption = "3RD        " & ArrayOfWinners(2)
                              End With
                              Finished = 0
                              PnumX = 2
                              'Load Wins
                              PauseStart.Value = True
                              Wins.Show vbModal
                              PauseStart.Value = True
                 End Select
                              
       Case 4   'we'll need only three winners to contest of which one will dropout leaving two to contest for the crown
                Select Case Finished
                       Case 1
                                ArrayOfWinners(0) = Nam
                           Select Case Nam
                               Case Pnam1   'Player1 is the first winner
                               Label4(0).Caption = "1st " & Pnam1
                               P1out = 1
                               Case Pnam2
                                Label4(1).Caption = "1st " & Pnam2
                                P2out = 1
                              Case Pnam3
                                Label4(2).Caption = "1st " & Pnam3
                                P3out = 1
                               Case Pnam4
                                 Label4(3).Caption = "1st " & Pnam4
                                 P4out = 1
                              End Select
                              FirstToFinish = Nam
                              Wins.Label2(0).Caption = "1ST        " & Nam
                        Case 2
                            ArrayOfWinners(1) = Nam
                            Select Case Nam
                               Case Pnam1   'Player1 is the 2nd winner
                               Label4(0).Caption = "2nd " & Pnam1
                               P1out = 2
                               Case Pnam2
                                Label4(1).Caption = "2nd " & Pnam2
                                P2out = 2
                              Case Pnam3
                                Label4(2).Caption = "2nd " & Pnam3
                                P3out = 2
                               Case Pnam4
                                 Label4(3).Caption = "2nd " & Pnam4
                                 P4out = 2
                             End Select
                             Wins.Label2(1).Caption = "2ND        " & Nam
           
                    Case 3
                            ArrayOfWinners(2) = Nam
                            Select Case Nam
                               Case Pnam1   'Player1 is the 2nd winner
                               Label4(0).Caption = "3rd " & Pnam1
                               P1out = 3
                               Case Pnam2
                                Label4(1).Caption = "3rd " & Pnam2
                                P2out = 3
                              Case Pnam3
                                Label4(2).Caption = "3rd " & Pnam3
                                P3out = 3
                               Case Pnam4
                                 Label4(3).Caption = "3rd " & Pnam4
                                 P4out = 3
                             End Select
                             Wins.Label2(2).Caption = "3RD        " & Nam
                                 'check whose name did not appear in arrayofwinners
                                   For Indkx = 0 To 2
                                       Select Case ArrayOfWinners(Indkx)
                                              Case Pnam1: Pone = True
                                              Case Pnam2: Ptwo = True
                                              Case Pnam3: Pthree = True
                                              Case Pnam4: Pfour = True
                                      End Select
                                   Next
                                    'only 3 of them will have a true value
                                    If Pone = False Then 'anyone that is false is the loozer
                                       ArrayOfWinners(3) = Pnam1
                                    ElseIf Ptwo = False Then
                                       ArrayOfWinners(3) = Pnam2
                                    ElseIf Pthree = False Then
                                       ArrayOfWinners(3) = Pnam3
                                    ElseIf Pfour = False Then
                                       ArrayOfWinners(3) = Pnam4
                                    End If
                                    Nam = ArrayOfWinners(3)
' If P1out = 0 Then   'WHO'S D LOOZER
' Nam = Pnam1
' ElseIf P2out = 0 Then
' Nam = Pnam2
' ElseIf P3out = 0 Then
' Nam = Pnam3
' ElseIf P4out = 0 Then
' Nam = Pnam4
' Else
' MsgBox ("Impossible Error @ Winning,occured")
' End If
                                    Loozer = Nam
                                    LooserNam = "4TH        " & Nam
                                   Wins.Label2(3).Caption = "4TH        " & Nam
                                  Wins.Label2(4).Visible = False
                                  Finished = 0
                                  PnumX = 3
                                  PauseStart.Value = True
                                  Wins.Show vbModal
                                  PauseStart.Value = True
                                  
       
                    End Select
                              

 End Select
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








