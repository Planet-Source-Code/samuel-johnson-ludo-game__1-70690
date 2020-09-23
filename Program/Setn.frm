VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Setn 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   7020
   ClientLeft      =   2670
   ClientTop       =   2580
   ClientWidth     =   11175
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Setn.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Setn.frx":076A
   ScaleHeight     =   7020
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   360
      TabIndex        =   89
      Top             =   360
      Width           =   615
      Begin VB.Image Image2 
         Height          =   375
         Index           =   1
         Left            =   0
         Picture         =   "Setn.frx":1B6C6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   0
         Left            =   0
         Picture         =   "Setn.frx":1CBBD
         Stretch         =   -1  'True
         ToolTipText     =   "Next"
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   3
      Left            =   6120
      TabIndex        =   74
      Top             =   5160
      Width           =   4935
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   15
         Left            =   720
         Picture         =   "Setn.frx":1E0CE
         ScaleHeight     =   375
         ScaleWidth      =   1815
         TabIndex        =   77
         Top             =   960
         Width           =   1815
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "HUMAN "
            BeginProperty Font 
               Name            =   "Franklin Gothic Medium"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   15
            Left            =   120
            TabIndex        =   78
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   13
         Left            =   720
         Picture         =   "Setn.frx":1FA9A
         ScaleHeight     =   375
         ScaleWidth      =   1815
         TabIndex        =   75
         Top             =   240
         Width           =   1815
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "COMPUTER"
            BeginProperty Font 
               Name            =   "Franklin Gothic Medium"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   13
            Left            =   120
            TabIndex        =   76
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture5 
         AutoRedraw      =   -1  'True
         Height          =   1455
         Index           =   3
         Left            =   0
         Picture         =   "Setn.frx":21466
         ScaleHeight     =   1395
         ScaleWidth      =   4995
         TabIndex        =   79
         Top             =   0
         Width           =   5055
         Begin VB.PictureBox Picture6 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   14
            Left            =   720
            Picture         =   "Setn.frx":50D62
            ScaleHeight     =   375
            ScaleWidth      =   1815
            TabIndex        =   83
            Top             =   960
            Width           =   1815
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "HUMAN"
               BeginProperty Font 
                  Name            =   "Franklin Gothic Medium"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Index           =   14
               Left            =   120
               TabIndex        =   84
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H004588CB&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   375
            Index           =   3
            Left            =   3120
            TabIndex        =   80
            Top             =   960
            Width           =   1575
         End
         Begin VB.PictureBox Picture6 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   12
            Left            =   720
            Picture         =   "Setn.frx":52886
            ScaleHeight     =   375
            ScaleWidth      =   1815
            TabIndex        =   81
            Top             =   240
            Width           =   1815
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "COMPUTER"
               BeginProperty Font 
                  Name            =   "Franklin Gothic Medium"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Index           =   12
               Left            =   120
               TabIndex        =   82
               Top             =   0
               Width           =   1815
            End
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "LOVETH"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   375
            Index           =   3
            Left            =   3120
            TabIndex        =   87
            Top             =   240
            Width           =   1455
         End
         Begin VB.Image Image1 
            Height          =   615
            Index           =   15
            Left            =   0
            Top             =   720
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   615
            Index           =   13
            Left            =   0
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "ENTER NAME"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   375
            Index           =   3
            Left            =   3000
            TabIndex        =   86
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "PLAYER 4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   85
            Top             =   0
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   615
            Index           =   12
            Left            =   0
            Top             =   0
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   615
            Index           =   14
            Left            =   0
            Top             =   720
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   2
      Left            =   6120
      TabIndex        =   60
      Top             =   3720
      Width           =   4935
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   11
         Left            =   720
         Picture         =   "Setn.frx":543AA
         ScaleHeight     =   375
         ScaleWidth      =   1815
         TabIndex        =   63
         Top             =   960
         Width           =   1815
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "HUMAN "
            BeginProperty Font 
               Name            =   "Franklin Gothic Medium"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   64
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   720
         Picture         =   "Setn.frx":55D76
         ScaleHeight     =   375
         ScaleWidth      =   1815
         TabIndex        =   61
         Top             =   240
         Width           =   1815
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "COMPUTER"
            BeginProperty Font 
               Name            =   "Franklin Gothic Medium"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   62
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture5 
         AutoRedraw      =   -1  'True
         Height          =   1455
         Index           =   2
         Left            =   0
         Picture         =   "Setn.frx":57742
         ScaleHeight     =   1395
         ScaleWidth      =   4995
         TabIndex        =   65
         Top             =   0
         Width           =   5055
         Begin VB.PictureBox Picture6 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   10
            Left            =   720
            Picture         =   "Setn.frx":8703E
            ScaleHeight     =   375
            ScaleWidth      =   1815
            TabIndex        =   69
            Top             =   960
            Width           =   1815
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "HUMAN"
               BeginProperty Font 
                  Name            =   "Franklin Gothic Medium"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   70
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox Picture6 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   8
            Left            =   720
            Picture         =   "Setn.frx":88B62
            ScaleHeight     =   375
            ScaleWidth      =   1815
            TabIndex        =   67
            Top             =   240
            Width           =   1815
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "COMPUTER"
               BeginProperty Font 
                  Name            =   "Franklin Gothic Medium"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Index           =   8
               Left            =   120
               TabIndex        =   68
               Top             =   0
               Width           =   1815
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H004588CB&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   375
            Index           =   2
            Left            =   3120
            TabIndex        =   66
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TOSINprof"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   375
            Index           =   2
            Left            =   3120
            TabIndex        =   73
            Top             =   240
            Width           =   1695
         End
         Begin VB.Image Image1 
            Height          =   615
            Index           =   11
            Left            =   0
            Top             =   720
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   615
            Index           =   9
            Left            =   0
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "ENTER NAME"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   375
            Index           =   2
            Left            =   3000
            TabIndex        =   72
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "PLAYER 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   71
            Top             =   0
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   615
            Index           =   8
            Left            =   0
            Top             =   0
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   615
            Index           =   10
            Left            =   0
            Top             =   720
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   1
      Left            =   6120
      TabIndex        =   45
      Top             =   2280
      Width           =   4935
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   720
         Picture         =   "Setn.frx":8A686
         ScaleHeight     =   375
         ScaleWidth      =   1815
         TabIndex        =   56
         Top             =   240
         Width           =   1815
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "COMPUTER"
            BeginProperty Font 
               Name            =   "Franklin Gothic Medium"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   57
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   720
         Picture         =   "Setn.frx":8C052
         ScaleHeight     =   375
         ScaleWidth      =   1815
         TabIndex        =   54
         Top             =   960
         Width           =   1815
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "HUMAN "
            BeginProperty Font 
               Name            =   "Franklin Gothic Medium"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   55
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture5 
         AutoRedraw      =   -1  'True
         Height          =   1455
         Index           =   1
         Left            =   0
         Picture         =   "Setn.frx":8DA1E
         ScaleHeight     =   1395
         ScaleWidth      =   4995
         TabIndex        =   46
         Top             =   0
         Width           =   5055
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H004588CB&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   375
            Index           =   1
            Left            =   3120
            TabIndex        =   51
            Top             =   960
            Width           =   1575
         End
         Begin VB.PictureBox Picture6 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   4
            Left            =   720
            Picture         =   "Setn.frx":BD31A
            ScaleHeight     =   375
            ScaleWidth      =   1815
            TabIndex        =   49
            Top             =   240
            Width           =   1815
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "COMPUTER"
               BeginProperty Font 
                  Name            =   "Franklin Gothic Medium"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   50
               Top             =   0
               Width           =   1815
            End
         End
         Begin VB.PictureBox Picture6 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   6
            Left            =   720
            Picture         =   "Setn.frx":BEE3E
            ScaleHeight     =   375
            ScaleWidth      =   1815
            TabIndex        =   47
            Top             =   960
            Width           =   1815
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "HUMAN"
               BeginProperty Font 
                  Name            =   "Franklin Gothic Medium"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Index           =   6
               Left            =   120
               TabIndex        =   48
               Top             =   0
               Width           =   1215
            End
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "PLAYER 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   59
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "ENTER NAME"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   375
            Index           =   1
            Left            =   3000
            TabIndex        =   53
            Top             =   480
            Width           =   1935
         End
         Begin VB.Image Image1 
            Height          =   615
            Index           =   5
            Left            =   0
            Top             =   0
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   615
            Index           =   7
            Left            =   0
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "STORM"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   375
            Index           =   1
            Left            =   3240
            TabIndex        =   52
            Top             =   240
            Width           =   1455
         End
         Begin VB.Image Image1 
            Height          =   615
            Index           =   4
            Left            =   0
            Top             =   0
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   615
            Index           =   6
            Left            =   0
            Top             =   720
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   0
      Left            =   6120
      TabIndex        =   32
      Top             =   840
      Width           =   4935
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   720
         Picture         =   "Setn.frx":C0962
         ScaleHeight     =   375
         ScaleWidth      =   1815
         TabIndex        =   42
         Top             =   960
         Width           =   1815
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "HUMAN "
            BeginProperty Font 
               Name            =   "Franklin Gothic Medium"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   43
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   720
         Picture         =   "Setn.frx":C232E
         ScaleHeight     =   375
         ScaleWidth      =   1815
         TabIndex        =   38
         Top             =   240
         Width           =   1815
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "COMPUTER"
            BeginProperty Font 
               Name            =   "Franklin Gothic Medium Cond"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture5 
         AutoRedraw      =   -1  'True
         Height          =   1455
         Index           =   0
         Left            =   0
         Picture         =   "Setn.frx":C3CFA
         ScaleHeight     =   1395
         ScaleWidth      =   4995
         TabIndex        =   33
         Top             =   0
         Width           =   5055
         Begin VB.PictureBox Picture6 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   720
            Picture         =   "Setn.frx":F35F6
            ScaleHeight     =   375
            ScaleWidth      =   1815
            TabIndex        =   40
            Top             =   960
            Width           =   1815
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "HUMAN"
               BeginProperty Font 
                  Name            =   "Franklin Gothic Medium"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   41
               Top             =   0
               Width           =   1215
            End
         End
         Begin VB.PictureBox Picture6 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   720
            Picture         =   "Setn.frx":F511A
            ScaleHeight     =   375
            ScaleWidth      =   1815
            TabIndex        =   36
            Top             =   240
            Width           =   1815
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "COMPUTER"
               BeginProperty Font 
                  Name            =   "Franklin Gothic Medium"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   37
               Top             =   0
               Width           =   1815
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H004588CB&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   375
            Index           =   0
            Left            =   3120
            TabIndex        =   34
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "PLAYER 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   58
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "KAZMA"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   375
            Index           =   0
            Left            =   3120
            TabIndex        =   44
            Top             =   240
            Width           =   1455
         End
         Begin VB.Image Image1 
            Height          =   645
            Index           =   3
            Left            =   0
            Top             =   720
            Width           =   645
         End
         Begin VB.Image Image1 
            Height          =   615
            Index           =   2
            Left            =   0
            Top             =   720
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   675
            Index           =   1
            Left            =   0
            Picture         =   "Setn.frx":F6C3E
            Top             =   0
            Width           =   675
         End
         Begin VB.Image Image1 
            Height          =   675
            Index           =   0
            Left            =   0
            Picture         =   "Setn.frx":F716D
            Top             =   0
            Width           =   675
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "ENTER NAME"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   375
            Index           =   0
            Left            =   3000
            TabIndex        =   35
            Top             =   480
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   975
      Left            =   480
      TabIndex        =   25
      Top             =   5280
      Width           =   975
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1050
         Index           =   0
         Left            =   0
         Picture         =   "Setn.frx":F771C
         ScaleHeight     =   1050
         ScaleWidth      =   1050
         TabIndex        =   27
         Top             =   0
         Width           =   1050
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1050
         Index           =   1
         Left            =   0
         Picture         =   "Setn.frx":F8C13
         ScaleHeight     =   1050
         ScaleWidth      =   1050
         TabIndex        =   26
         ToolTipText     =   "Next"
         Top             =   0
         Width           =   1050
      End
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   7
      Left            =   480
      Picture         =   "Setn.frx":FA124
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   24
      Top             =   4320
      Width           =   660
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   6
      Left            =   480
      Picture         =   "Setn.frx":FAB0B
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   23
      Top             =   4320
      Width           =   660
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   5
      Left            =   480
      Picture         =   "Setn.frx":FB52F
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   22
      Top             =   3000
      Width           =   660
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   4
      Left            =   480
      Picture         =   "Setn.frx":FBF16
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   21
      Top             =   3000
      Width           =   660
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   3
      Left            =   480
      Picture         =   "Setn.frx":FC93A
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   20
      Top             =   1680
      Width           =   660
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   2
      Left            =   480
      Picture         =   "Setn.frx":FD321
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   19
      Top             =   1680
      Width           =   660
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Index           =   1
      Left            =   4440
      Picture         =   "Setn.frx":FDD45
      ScaleHeight     =   1050
      ScaleWidth      =   1050
      TabIndex        =   18
      Top             =   5280
      Width           =   1050
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Index           =   0
      Left            =   4440
      Picture         =   "Setn.frx":FF552
      ScaleHeight     =   1050
      ScaleWidth      =   1050
      TabIndex        =   17
      ToolTipText     =   "Quit"
      Top             =   5280
      Width           =   1050
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   7
      Left            =   1680
      Picture         =   "Setn.frx":100DB9
      ScaleHeight     =   585
      ScaleWidth      =   2355
      TabIndex        =   13
      Top             =   4320
      Width           =   2355
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Four Players"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Index           =   7
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   6
      Left            =   1680
      Picture         =   "Setn.frx":102785
      ScaleHeight     =   585
      ScaleWidth      =   2355
      TabIndex        =   11
      Top             =   4320
      Width           =   2355
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Four Players"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   5
      Left            =   1680
      Picture         =   "Setn.frx":1042A9
      ScaleHeight     =   585
      ScaleWidth      =   2355
      TabIndex        =   9
      Top             =   3000
      Width           =   2355
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Three Players"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Index           =   5
         Left            =   195
         TabIndex        =   10
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   4
      Left            =   1680
      Picture         =   "Setn.frx":105C75
      ScaleHeight     =   585
      ScaleWidth      =   2355
      TabIndex        =   7
      Top             =   3000
      Width           =   2355
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Three Players"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   4
         Left            =   75
         TabIndex        =   8
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   3
      Left            =   1680
      Picture         =   "Setn.frx":107799
      ScaleHeight     =   585
      ScaleWidth      =   2355
      TabIndex        =   5
      Top             =   1680
      Width           =   2355
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Two  Players"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Index           =   3
         Left            =   315
         TabIndex        =   6
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   2
      Left            =   1680
      Picture         =   "Setn.frx":109165
      ScaleHeight     =   585
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   1680
      Width           =   2355
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Two Players"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   1
      Left            =   5520
      Picture         =   "Setn.frx":10AC89
      ScaleHeight     =   585
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   8280
      Width           =   2355
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "One Player"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   0
      Left            =   5520
      Picture         =   "Setn.frx":10C655
      ScaleHeight     =   585
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   8280
      Width           =   2355
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "One Player"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   10440
      Top             =   1440
   End
   Begin WMPLibCtl.WindowsMediaPlayer SoundClick 
      Height          =   495
      Left            =   8100
      TabIndex        =   91
      Top             =   120
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
   Begin WMPLibCtl.WindowsMediaPlayer SoundMove 
      Height          =   495
      Left            =   7275
      TabIndex        =   90
      Top             =   0
      Width           =   690
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
      _cx             =   1217
      _cy             =   873
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT PLAYERS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   1080
      TabIndex        =   88
      Top             =   360
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   3
      Left            =   4920
      Picture         =   "Setn.frx":10E179
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   2
      Left            =   4920
      Picture         =   "Setn.frx":10F5BB
      Stretch         =   -1  'True
      ToolTipText     =   "Back"
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LUDO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   345
      Index           =   3
      Left            =   360
      TabIndex        =   31
      Top             =   6360
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LUDO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   345
      Index           =   2
      Left            =   4725
      TabIndex        =   30
      Top             =   6360
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LUDO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   345
      Index           =   1
      Left            =   4800
      TabIndex        =   29
      Top             =   360
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LUDO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   345
      Index           =   0
      Left            =   360
      TabIndex        =   28
      Top             =   360
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HOW MANY PLAYERS?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5895
   End
End
Attribute VB_Name = "Setn"
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

Dim Loaded As Boolean
Private PlaySound As Boolean
Private Sub Form_Load()
Me.Width = 5970
'Me.Left = Screen.Height / 1.7
'Me.Top = Screen.Height / 5

Frame1.Visible = False
For k = 0 To 3
Label8(k).Visible = False
Label6(k).Visible = False
Text1(k).Visible = False
Frame2(k).Visible = False
Frame2(k).Left = 480
Image2(k).Visible = False
Next
Label4.Visible = False
Frame3.Visible = False
For k = 2 To 14 Step 2
Image1(k).Picture = Image1(0).Picture
Image1(k + 1).Picture = Image1(1).Picture
Next
Back
'Load Board
Pnam1 = ""
Pnam2 = ""
Pnam3 = ""
Pnam4 = ""
 'Dirxtry = "D:\VB98\sam's vb project\My vb projects\Ludo Game\Ludo Sound\"
 
 End Sub
Sub Back()
Loaded = False
For J = 0 To 3
Label3(J).Visible = True
Frame2(J).Visible = False
Image2(J).Visible = False
Next
For H = 2 To 7
Picture1(H).Visible = True
Picture3(H).Visible = True
Next
For H = 0 To 1
Picture2(H).Visible = True
Next
Label4.Visible = False
Label1.Visible = True
Frame3.Visible = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PlaySound = False Then SoundMove.Close
PlaySound = True

If Loaded = False Then
For k = 1 To 7 Step 2
Picture1(k).Visible = True
Next
Picture2(1).Visible = True
Picture4(0).Visible = True
Else
Image2(1).Visible = True
Image2(3).Visible = True
For J = 0 To 3
Label7(J).ForeColor = vbBlack
Next

End If
End Sub

Private Sub Image2_Click(Index As Integer)
  Dim H As Integer

Select Case Index
Case 0    'forward arrow
  Select Case Ptyp1
        Case 11 'computer
             Pnam1 = "KaZma"
        Case 1  'human
                 Pnam1 = Trim(UCase(Text1(0).Text))
                If Pnam1 = "" Then H = MsgBox("Please Do Enter A Name For Player 1", vbInformation): Exit Sub
                 Case Else
           H = MsgBox("Please Specify Player Type For Player 1", vbInformation): Exit Sub

End Select
  
  Select Case Ptyp2
           Case 11
           Pnam2 = "STORM"
           Case 1
                 Pnam2 = Trim(UCase(Text1(1).Text))
                If Pnam2 = "" Then H = MsgBox("Please Do Enter A Name For Player 2", vbInformation): Exit Sub
 
                Case Else
           H = MsgBox("Please Specify Player Type For Player 2", vbInformation): Exit Sub

 End Select
 Board.Label4(0).Caption = Pnam1
  Board.Label4(1).Caption = Pnam2
Board.Label4(2).Visible = False
Board.Label4(3).Visible = False

  Select Case Pnum  'how many players?
   Case 3   '3 players
     Select Case Ptyp3
           Case 11  'computer
               Pnam3 = "TosinProf"
           Case 1   'human
                Pnam3 = Trim(UCase(Text1(2).Text))
                If Pnam3 = "" Then H = MsgBox("Please Do Enter A Name For Player 3", vbInformation): Exit Sub
           Case Else
           H = MsgBox("Please Specify Player Type For Player 3", vbInformation): Exit Sub
  
     End Select
     Board.Label4(2).Caption = Pnam3
 Board.Label4(2).Visible = True

   Case 4 '4 players
   
        Select Case Ptyp3
           Case 11  'computer
               Pnam3 = "TosinProf"
           Case 1   'human
                Pnam3 = Trim(UCase(Text1(2).Text))
                If Pnam3 = "" Then H = MsgBox("Please Do Enter A Name For Player 3", vbInformation): Exit Sub
               Case Else
           H = MsgBox("Please Specify Player Type For Player 3", vbInformation): Exit Sub
    End Select

     Select Case Ptyp4
            Case 11
            Pnam4 = "Loveth"
            Case 1
                Pnam4 = Trim(UCase(Text1(3).Text))
                If Pnam4 = "" Then H = MsgBox("Please Do Enter A Name For Player 4", vbInformation): Exit Sub
            Case Else
           H = MsgBox("Please Specify Player Type For Player 4", vbInformation): Exit Sub
   End Select
 Board.Label4(2).Caption = Pnam3
Board.Label4(3).Caption = Pnam4
Board.Label4(2).Visible = True
Board.Label4(3).Visible = True

End Select

'If UCase(Pnam1) = UCase(Pnam2) Or UCase(Pnam2) = UCase(Pnam3) Or UCase(Pnam3) = UCase(Pnam4) Or UCase(Pnam1) = UCase(Pnam4) Then
Dim XX As String, XXX As String, U As Integer, T As Integer, F As Integer
'to verify that no two player has been asigned the same name
For U = 0 To 3
XX = UCase(Board.Label4(U).Caption)
For T = 0 To 3
If U = T Then GoTo 5
XXX = UCase(Board.Label4(T).Caption)
If XX = XXX Then
F = MsgBox("No Two Player Must Bear The Same Name," + vbCrLf + "Please Distinguish.", vbExclamation)
Exit Sub
End If
5 Next T
Next

Unload Setn
'Unload Board
Load Board
Board.Show
If Pnum = 0 Then Beep: Exit Sub
Board.Frame1.Move P1Left, P12Top
Rotate = 1
PlayerType = 1
Paused = True

For T = 0 To 3
   ArrayOfWinners(T) = ""
Next

If Ptyp1 = 11 Then PlayerType = 11: Board.ArtificialIntelligence

Case 2   'barkward arrow
SoundClick.URL = Dirxtry & "lighttrail2.ogg"
Back
End Select
End Sub

Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index Mod 2 = 0 Then Exit Sub
Image2(Index).Visible = False

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PlaySound = False Then SoundMove.Close
PlaySound = True

If Loaded = False Then
For k = 1 To 7 Step 2
Picture1(k).Visible = True
Next
Picture2(1).Visible = True
Picture4(0).Visible = True
Else
Image2(1).Visible = True
Image2(3).Visible = True
For J = 0 To 3
Label7(J).ForeColor = vbBlack
Next

End If

End Sub

Private Sub Label2_Click(Index As Integer)
SoundClick.URL = Dirxtry & "3wav.wav"
Frame1.Visible = True  'the effects
For J = 3 To 7 Step 2
Picture3(J).Visible = True
Next
Picture3(Index).Visible = True
Picture3(Index + 1).Visible = False
Select Case Index    'to know how many players the user choosed
Case 2
Pnum = 2
For k = 0 To 3
Board.Ply3(k).Visible = False
Board.Ply4(k).Visible = False
Next
Case 4
Pnum = 3
For k = 0 To 3
Board.Ply3(k).Visible = True
Board.Ply4(k).Visible = False
Next
Case 6
Pnum = 4
For k = 0 To 3
Board.Ply3(k).Visible = True
Board.Ply4(k).Visible = True
Next

End Select
PnumX = Pnum
'ReDim ArrayOfWinners(PnumX)
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If PlaySound = True Then SoundMove.URL = Dirxtry & "2wav.wav"
PlaySound = False

Select Case Index
Case 1, 3, 5, 7
Case Else
Exit Sub
End Select
Picture1(Index).Visible = False
Picture1(Index - 1).Visible = True


End Sub

Private Sub Label5_Click(Index As Integer)
SoundClick.URL = Dirxtry & "3wav.wav"
  
  Image1(Index + 1).Visible = False
Select Case Index
Case 0, 4, 8, 12
Image1(Index + 3).Visible = True
Case Else
Image1(Index - 1).Visible = True
End Select
For J = 0 To 1
Image2(J).Visible = True
Next
Frame3.Visible = True
Select Case Index
Case 0
Ptyp1 = 11
Label8(0).Visible = True
Label6(0).Visible = False
Text1(0).Visible = False
Case 2
Ptyp1 = 1
Label8(0).Visible = False
Label6(0).Visible = True
Text1(0).Visible = True
Text1(0).SetFocus
Case 4
Ptyp2 = 11
Label8(1).Visible = True
Label6(1).Visible = False
Text1(1).Visible = False
Case 6
Ptyp2 = 1
Label8(1).Visible = False
Label6(1).Visible = True
Text1(1).Visible = True
Text1(1).SetFocus
Case 8
Ptyp3 = 11
Label8(2).Visible = True
Label6(2).Visible = False
Text1(2).Visible = False
Case 10
Ptyp3 = 1
Label8(2).Visible = False
Label6(2).Visible = True
Text1(2).Visible = True
Text1(2).SetFocus

Case 12
Ptyp4 = 11
Label8(3).Visible = True
Label6(3).Visible = False
Text1(3).Visible = False
Case 14
Ptyp4 = 1
Label8(3).Visible = False
Label6(3).Visible = True
Text1(3).Visible = True
Text1(3).SetFocus

End Select



End Sub

Private Sub Label5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MousMov Index
End Sub
Sub MousMov(Index)
If PlaySound = True Then SoundMove.URL = Dirxtry & "start.wav"
PlaySound = False
If Index Mod 2 = 0 Then Exit Sub
Picture6(Index).Visible = False

End Sub

Private Sub Picture1_Click(Index As Integer)
Label2_Click (Index)
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If PlaySound = True Then SoundMove.URL = Dirxtry & "2wav.wav"
PlaySound = False

If Index Mod 2 = 0 Then Exit Sub
Picture1(Index).Visible = False
End Sub

Private Sub Picture2_Click(Index As Integer)
F = MsgBox("Sure To Exit?", vbYesNo + vbExclamation + vbDefaultButton2)
If F = vbYes Then End
End Sub

Private Sub Picture2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If PlaySound = True Then SoundMove.URL = Dirxtry & "2wav.wav"
PlaySound = False
Picture2(1).Visible = False
End Sub

Private Sub Picture4_Click(Index As Integer)

SoundClick.URL = Dirxtry & "lighttrail2.ogg"

Loaded = True    'clear up
For J = 0 To 3
Label3(J).Visible = False
Image2(J).Visible = True
Next

For H = 0 To 1
Frame2(H).Visible = True
Picture2(H).Visible = False
Next

Select Case Pnum   'to dtermin how many options to display
Case 3
Frame2(2).Visible = True
Case 4
Frame2(2).Visible = True
Frame2(3).Visible = True
End Select

For H = 2 To 7
Picture1(H).Visible = False
Picture3(H).Visible = False
Next

Frame1.Visible = False
Label4.Visible = True
Label1.Visible = False
Ptyp1 = 0
Ptyp2 = 0
Ptyp3 = 0
Ptyp4 = 0

End Sub

Private Sub Picture4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If PlaySound = True Then SoundMove.URL = Dirxtry & "2wav.wav"
PlaySound = False
Picture4(0).Visible = False
End Sub

Private Sub Picture5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
PlaySound = True
For k = 1 To 15 Step 2
Picture6(k).Visible = True
Next

For J = 0 To 3
Label7(J).ForeColor = vbBlack
Next
Label7(Index).ForeColor = vbRed
End Sub

Private Sub Picture6_Click(Index As Integer)
Label5_Click (Index)
End Sub

Private Sub Picture6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MousMov Index
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case KeyAscii
        Case Asc("A") To Asc("Z"), Asc("a") To Asc("z"), Asc(vbBack)
        Case Else
        KeyAscii = 0
End Select
        
End Sub

Private Sub Timer1_Timer()
'Exit Sub
'Timer1.Enabled = False
Static k As Integer, Max As Boolean, b As Integer
If Max = False Then
k = k + 1
Label1.Font.Size = k
Label4.Font.Size = k ' - b
If k > 20 Then Max = True
Else
k = k - 1
Label1.Font.Size = k
Label4.Font.Size = k '- b

If k = 15 Then Max = False: b = 5
End If

End Sub
