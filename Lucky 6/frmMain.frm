VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "LETS PLAY LOTTO"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Play Digit Games (In Exact Order)"
      Height          =   3495
      Left            =   4440
      TabIndex        =   36
      Top             =   120
      Width           =   4695
      Begin VB.Frame Frame10 
         Caption         =   "4 Digit Game"
         Height          =   1455
         Left            =   120
         TabIndex        =   49
         Top             =   1920
         Width           =   4455
         Begin VB.Frame Frame12 
            Height          =   495
            Left            =   240
            TabIndex        =   53
            Top             =   840
            Width           =   2295
            Begin VB.Label lbl4Digit 
               Alignment       =   2  'Center
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   3
               Left            =   1680
               TabIndex        =   57
               Top             =   120
               Width           =   300
            End
            Begin VB.Label lbl4Digit 
               Alignment       =   2  'Center
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   2
               Left            =   1200
               TabIndex        =   56
               Top             =   120
               Width           =   300
            End
            Begin VB.Label lbl4Digit 
               Alignment       =   2  'Center
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   1
               Left            =   720
               TabIndex        =   55
               Top             =   120
               Width           =   300
            End
            Begin VB.Label lbl4Digit 
               Alignment       =   2  'Center
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   0
               Left            =   240
               TabIndex        =   54
               Top             =   120
               Width           =   300
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Enter Combination Here"
            Height          =   615
            Left            =   240
            TabIndex        =   52
            Top             =   240
            Width           =   2295
            Begin VB.TextBox txt4Digit 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   1680
               TabIndex        =   17
               Top             =   240
               Width           =   300
            End
            Begin VB.TextBox txt4Digit 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   1200
               TabIndex        =   16
               Top             =   240
               Width           =   300
            End
            Begin VB.TextBox txt4Digit 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   720
               TabIndex        =   15
               Top             =   240
               Width           =   300
            End
            Begin VB.TextBox txt4Digit 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   240
               TabIndex        =   14
               Top             =   240
               Width           =   300
            End
         End
         Begin VB.CommandButton cmdDraw 
            Caption         =   "Draw"
            Height          =   375
            Index           =   2
            Left            =   2760
            TabIndex        =   50
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox txtBet 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   3360
            TabIndex        =   18
            Top             =   600
            Width           =   975
         End
         Begin VB.Timer tmr4Digit 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   2040
            Top             =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Place Bet Here"
            Height          =   195
            Index           =   1
            Left            =   3000
            TabIndex        =   58
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Php"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   2760
            TabIndex        =   51
            Top             =   600
            Width           =   360
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "6 Digit Game"
         Height          =   1575
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   4455
         Begin VB.Frame Frame7 
            Caption         =   "Enter Combination Here"
            Height          =   735
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   2415
            Begin VB.TextBox txt6Digit 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   0
               Left            =   120
               MaxLength       =   1
               TabIndex        =   7
               Top             =   240
               Width           =   300
            End
            Begin VB.TextBox txt6Digit 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   1
               Left            =   480
               MaxLength       =   1
               TabIndex        =   8
               Top             =   240
               Width           =   300
            End
            Begin VB.TextBox txt6Digit 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   2
               Left            =   840
               MaxLength       =   1
               TabIndex        =   9
               Top             =   240
               Width           =   300
            End
            Begin VB.TextBox txt6Digit 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   3
               Left            =   1200
               MaxLength       =   1
               TabIndex        =   10
               Top             =   240
               Width           =   300
            End
            Begin VB.TextBox txt6Digit 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   4
               Left            =   1560
               MaxLength       =   1
               TabIndex        =   11
               Top             =   240
               Width           =   300
            End
            Begin VB.TextBox txt6Digit 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   5
               Left            =   1920
               MaxLength       =   1
               TabIndex        =   12
               Top             =   240
               Width           =   300
            End
         End
         Begin VB.Frame Frame8 
            Height          =   495
            Left            =   120
            TabIndex        =   39
            Top             =   960
            Width           =   2415
            Begin VB.Label lbl6Digit 
               Alignment       =   2  'Center
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   0
               Left            =   120
               TabIndex        =   45
               Top             =   120
               Width           =   300
            End
            Begin VB.Label lbl6Digit 
               Alignment       =   2  'Center
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   1
               Left            =   480
               TabIndex        =   44
               Top             =   120
               Width           =   300
            End
            Begin VB.Label lbl6Digit 
               Alignment       =   2  'Center
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   2
               Left            =   840
               TabIndex        =   43
               Top             =   120
               Width           =   300
            End
            Begin VB.Label lbl6Digit 
               Alignment       =   2  'Center
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   3
               Left            =   1200
               TabIndex        =   42
               Top             =   120
               Width           =   300
            End
            Begin VB.Label lbl6Digit 
               Alignment       =   2  'Center
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   4
               Left            =   1560
               TabIndex        =   41
               Top             =   120
               Width           =   300
            End
            Begin VB.Label lbl6Digit 
               Alignment       =   2  'Center
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   5
               Left            =   1920
               TabIndex        =   40
               Top             =   120
               Width           =   300
            End
         End
         Begin VB.CommandButton cmdDraw 
            Caption         =   "Draw"
            Height          =   375
            Index           =   1
            Left            =   2760
            TabIndex        =   38
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtBet 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3360
            TabIndex        =   13
            Top             =   720
            Width           =   975
         End
         Begin VB.Timer tmr6Digit 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   1920
            Top             =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Place Bet Here"
            Height          =   195
            Index           =   0
            Left            =   2880
            TabIndex        =   48
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Php"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   2760
            TabIndex        =   47
            Top             =   720
            Width           =   360
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Play Lotto (In Any Order)"
      Height          =   3495
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   4215
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   3735
         Begin VB.Label Numbers 
            Alignment       =   2  'Center
            Caption         =   "49"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   500
         End
         Begin VB.Label Numbers 
            Alignment       =   2  'Center
            Caption         =   "49"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   1
            Left            =   720
            TabIndex        =   34
            Top             =   240
            Width           =   500
         End
         Begin VB.Label Numbers 
            Alignment       =   2  'Center
            Caption         =   "49"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   2
            Left            =   1320
            TabIndex        =   33
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Numbers 
            Alignment       =   2  'Center
            Caption         =   "49"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   3
            Left            =   1920
            TabIndex        =   32
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Numbers 
            Alignment       =   2  'Center
            Caption         =   "49"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   4
            Left            =   2520
            TabIndex        =   31
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Numbers 
            Alignment       =   2  'Center
            Caption         =   "49"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   5
            Left            =   3120
            TabIndex        =   30
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pick Type of Game"
         Height          =   1215
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   1815
         Begin VB.OptionButton optGame 
            Caption         =   "Lotto 6 / 42"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optGame 
            Caption         =   "Mega Lotto"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   520
            Width           =   1215
         End
         Begin VB.OptionButton optGame 
            Caption         =   "Super Lotto"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   26
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Enter Combination Here"
         Height          =   855
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   3735
         Begin VB.TextBox txtNum 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            MaxLength       =   2
            TabIndex        =   0
            Top             =   360
            Width           =   500
         End
         Begin VB.TextBox txtNum 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   720
            MaxLength       =   2
            TabIndex        =   1
            Top             =   360
            Width           =   500
         End
         Begin VB.TextBox txtNum 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   2
            Top             =   360
            Width           =   500
         End
         Begin VB.TextBox txtNum 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   3
            Top             =   360
            Width           =   500
         End
         Begin VB.TextBox txtNum 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   2520
            MaxLength       =   2
            TabIndex        =   4
            Top             =   360
            Width           =   500
         End
         Begin VB.TextBox txtNum 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   3120
            MaxLength       =   2
            TabIndex        =   5
            Top             =   360
            Width           =   500
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Place Bet Here"
         Height          =   1215
         Left            =   2160
         TabIndex        =   21
         Top             =   2040
         Width           =   1815
         Begin VB.CommandButton cmdDraw 
            Caption         =   "Draw"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtBet 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Php"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   360
         End
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   960
         Top             =   480
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   480
         Top             =   480
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7320
      TabIndex        =   19
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      Caption         =   "1000.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   60
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "This is your money Php"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   59
      Top             =   3720
      Width           =   2820
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ctr As Integer
Dim Balls As Integer
Dim BallsToGo As Integer
Dim BallsDrawn As Integer
Dim Mixing As Integer
Dim TypeofGame As Integer
Dim WinCombi As Integer
Dim Ctr4Digit As Integer
Dim Ctr6Digit As Integer
Dim PrizeLotto4 As Integer
Dim PrizeLotto5 As Integer
Dim PrizeLotto6 As Integer


Private Sub cmdDraw_Click(Index As Integer)
Select Case Index

Case Is = 0
If EntryOk = False Then Exit Sub
Timer2.Enabled = True
BallsToGo = 0
Mixing = 0
cmdDraw(0).Enabled = False

Case Is = 1
tmr6Digit.Enabled = True
Ctr6Digit = 0

Case Is = 2
tmr4Digit.Enabled = True
Ctr4Digit = 0

End Select
End Sub

Private Function DrawNumbers()

For ctr = BallsToGo To 5
    Randomize
    Balls = Int((TypeofGame * Rnd) + 1)
    Numbers(ctr).Caption = Balls
Next
    
End Function


Private Sub cmdVerify_Click()
OpenLottoFile
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
GamePick
End Sub

Private Sub Label5_Click()

End Sub

Private Sub optGame_Click(Index As Integer)
GamePick
End Sub

Private Sub Timer1_Timer()
    DrawNumbers
End Sub

Private Function DrawBalls()
Timer1.Enabled = True
Mixing = Mixing + 1
If Mixing = 5 Then
    Mixing = 0
    If IsTheSame = True Then Exit Function
    Numbers(BallsToGo).Caption = Balls
    BallsToGo = BallsToGo + 1
    If BallsToGo > 5 Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        cmdDraw(0).Enabled = True
        CountWinNum
    End If
End If
End Function

Private Sub Timer2_Timer()
Timer1.Enabled = True
DrawBalls
End Sub

Private Function IsTheSame() As Boolean
IsTheSame = False
    If BallsToGo > 0 Then
        For ctr = 1 To BallsToGo
            If Balls = Val(Numbers(ctr - 1).Caption) Then
                IsTheSame = True
                Exit Function
            End If
        Next
    End If
End Function

Private Sub ChangeCaption(tmpString As String)
For ctr = 0 To 5
    Numbers(ctr).Caption = tmpString
Next
End Sub

Private Sub GamePick()
If optGame(0).Value = True Then
    ChangeCaption "42"
    TypeofGame = 42
    PrizeLotto4 = 500
    PrizeLotto5 = 20000
ElseIf optGame(1).Value = True Then
    ChangeCaption "45"
    TypeofGame = 45
    PrizeLotto4 = 600
    PrizeLotto5 = 22000
ElseIf optGame(2).Value = True Then
    ChangeCaption "49"
    TypeofGame = 49
    PrizeLotto4 = 700
    PrizeLotto5 = 25000
End If
End Sub

Private Sub CountWinNum()
Dim mybet As Double
WinCombi = 0
Dim x, y
For x = 0 To 5
    For y = 0 To 5
        If Val(txtNum(x).Text) = Val(Numbers(y).Caption) Then
            WinCombi = WinCombi + 1
        End If
    Next
Next

Select Case WinCombi

Case Is = 0, 1, 2
    MsgBox "Sorry you lost", vbOKOnly + vbInformation, "Loser"
    lblMoney.Caption = Format(Val(lblMoney.Caption) - Val(txtBet(0).Text), "###.00")
    If Val(lblMoney.Caption) <= 0 Then
        MsgBox "Game Over", vbOKOnly + vbInformation, "LOSER"
        End
    End If
Case Is = 3
    MsgBox "You got 3 combination. You won back your bet", vbOKOnly + vbInformation, "You Win"
Case Is = 4
    MsgBox "You got 4 combination. You won Php " & PrizeLotto4 & " for every Php 10 of your bet", vbOKOnly + vbInformation, "You Win"
    mybet = (Val(txtBet(0).Text) / 10) * PrizeLotto4
    lblMoney.Caption = Format(Val(lblMoney.Caption) + mybet, "###.00")
Case Is = 5
    MsgBox "You got 5 combination. You won Php " & PrizeLotto5 & " for evry Php 10 of your bet", vbOKOnly + vbInformation, "You Win"
    mybet = (Val(txtBet(0).Text) / 10) * PrizeLotto5
    lblMoney.Caption = Format(Val(lblMoney.Caption) + Val(txtBet(0).Text), "###.00")
Case Is = 6
    MsgBox "You hit the jackpot!!!!", vbOKOnly + vbInformation, "Jackpot!!!"
    lblMoney.Caption = Format(Val(lblMoney.Caption) + 3000000, "###.00")
End Select
End Sub

Private Function EntryOk() As Boolean
Dim x, y, z
EntryOk = True
For x = 0 To 5
    If txtNum(x).Text = "" Then
        MsgBox "Missing Combination", vbOKOnly + vbInformation, "Abort Draw"
        txtNum(x).SetFocus
        EntryOk = False
        Exit Function
    End If
    
    If Val(txtNum(x).Text) > TypeofGame Then
        MsgBox "Invalid Number in the game you choose", vbOKOnly + vbInformation, "Abort Draw"
        txtNum(x).SetFocus
        SendKeys "{HOME}+{END}"
        EntryOk = False
        Exit Function
    End If
    
    If Val(txtNum(x).Text) <= 0 Then
        MsgBox "Invalid number.  Please enter a number from 1 to " & TypeofGame & ".", vbOKOnly + vbInformation, "Abort Draw"
        txtNum(x).SetFocus
        SendKeys "{HOME}+{END}"
        EntryOk = False
        Exit Function
    End If
    
    If Not IsNumeric(txtNum(x).Text) Then
        MsgBox "Invalid input.  Please enter a number from 1 to " & TypeofGame & ".", vbOKOnly + vbInformation, "Abort Draw"
        txtNum(x).SetFocus
        SendKeys "{HOME}+{END}"
        EntryOk = False
        Exit Function
    End If

    z = txtNum(x).Text
    If x > 0 Then
        For y = 0 To x - 1
        If z = txtNum(y).Text Then
            MsgBox "Duplicate number. You can only input a number once.", vbOKOnly + vbInformation, "Abort Draw"
            txtNum(y).SetFocus
            SendKeys "{HOME}+{END}"
            EntryOk = False
            Exit Function
        End If
        Next
    End If
Next
If txtBet(0).Text = "" Then
    txtBet(0).Text = "0"
End If
If Not IsNumeric(txtBet(0).Text) Then
    MsgBox "Invalid Bet", vbOKOnly + vbInformation, "Abort Draw"
    txtBet(0).SetFocus
    SendKeys "{HOME}+{END}"
    EntryOk = False
    Exit Function
End If

If Val(txtBet(0).Text) < 10 Then
    MsgBox "The minimum bet allowed is Php 10.00", vbOKOnly + vbInformation, "Abort Draw"
    txtBet(0).SetFocus
    SendKeys "{HOME}+{END}"
    EntryOk = False
    Exit Function
End If
End Function

Private Function Draw6Digit()
For ctr = 0 To 5
    Randomize
    lbl6Digit(ctr).Caption = Int((10 * Rnd) + 1)
    If lbl6Digit(ctr).Caption = "10" Then lbl6Digit(ctr).Caption = "0"
    
    Ctr6Digit = Ctr6Digit + 1
    If Ctr6Digit = 500 Then
        tmr6Digit.Enabled = False
        Prize6Digit
        Exit Function
    End If
Next
End Function

Private Function Draw4Digit()
For ctr = 0 To 3
    Randomize
    lbl4Digit(ctr).Caption = Int((10 * Rnd) + 1)
    If lbl4Digit(ctr).Caption = "10" Then lbl4Digit(ctr).Caption = "0"
    
    Ctr4Digit = Ctr4Digit + 1
    If Ctr4Digit = 500 Then
        tmr4Digit.Enabled = False
        Prize4Digit
        Exit Function
    End If
Next
End Function

Private Sub tmr4Digit_Timer()
Draw4Digit
End Sub

Private Sub tmr6Digit_Timer()
Draw6Digit
End Sub

Private Sub Prize6Digit()
Dim mybet  As Double
Dim numbet As String
Dim lottobet As String

numbet = txt6Digit(0).Text & txt6Digit(1).Text & txt6Digit(2).Text & txt6Digit(3).Text & txt6Digit(4).Text & txt6Digit(5).Text
lottobet = lbl6Digit(0).Caption & lbl6Digit(1).Caption & lbl6Digit(2).Caption & lbl6Digit(3).Caption & lbl6Digit(4).Caption & lbl6Digit(5).Caption

If Left(numbet, 3) = Left(lottobet, 3) Or Right(numbet, 3) = Right(lottobet, 3) Then
    MsgBox "You won back your bet", vbOKOnly + vbInformation, "Winner"
    Exit Sub
End If

If Left(numbet, 4) = Left(lottobet, 4) Or Right(numbet, 4) = Right(lottobet, 4) Then
    MsgBox "You won Php 500 for every Php 10 of your bet", vbOKOnly + vbInformation, "Winner"
    mybet = (Val(txtBet(1).Text) / 10) * 500
    lblMoney.Caption = Val(lblMoney.Caption) + mybet
    Exit Sub
End If

If Left(numbet, 5) = Left(lottobet, 5) Or Right(numbet, 5) = Right(lottobet, 5) Then
    MsgBox "You won Php 700 for every Php 10 of your bet", vbOKOnly + vbInformation, "Winner"
    mybet = (Val(txtBet(1).Text) / 10) * 700
    lblMoney.Caption = Val(lblMoney.Caption) + mybet
    Exit Sub
End If

If lottobet = numbet Then
    MsgBox "You hit the Jackpot!!!", vbOKOnly + vbInformation, "JACKPOT!!!"
    lblMoney.Caption = Val(lblMoney.Caption) + 20000
    Exit Sub
    End If
MsgBox "You Lost", vbOKOnly + vbInformation, "Loser"
lblMoney.Caption = Val(lblMoney.Caption) - Val(txtBet(1).Text)
End Sub

Private Sub Prize4Digit()
Dim mybet  As Double
Dim numbet As String
Dim lottobet As String

numbet = txt4Digit(0).Text & txt4Digit(1).Text & txt4Digit(2).Text & txt4Digit(3).Text
lottobet = lbl4Digit(0).Caption & lbl4Digit(1).Caption & lbl4Digit(2).Caption & lbl4Digit(3).Caption
If Left(numbet, 2) = Left(lottobet, 2) Or Right(numbet, 2) = Right(lottobet, 2) Then
    MsgBox "You won back your bet", vbOKOnly + vbInformation, "Winner"
    Exit Sub
End If

If Left(numbet, 3) = Left(lottobet, 3) Or Right(numbet, 3) = Right(lottobet, 3) Then
    MsgBox "You won Php 500 for every Php 10 of your bet", vbOKOnly + vbInformation, "Winner"
    mybet = (Val(txtBet(2).Text) / 10) * 500
    lblMoney.Caption = Val(lblMoney.Caption) + mybet
    Exit Sub
End If


If lottobet = numbet Then
    MsgBox "You hit the Jackpot!!!", vbOKOnly + vbInformation, "JACKPOT!!!"
    lblMoney.Caption = Val(lblMoney.Caption) + 10000
    Exit Sub
End If

MsgBox "You Lost", vbOKOnly + vbInformation, "Loser"
lblMoney.Caption = Val(lblMoney.Caption) - Val(txtBet(2).Text)
If Val(lblMoney.Caption) <= 0 Then
    MsgBox "Game Over", vbOKOnly + vbInformation, "LOSER"
    End
End If
If Val(lblMoney.Caption) <= 0 Then
    MsgBox "Game Over", vbOKOnly + vbInformation, "LOSER"
    End
End If
End Sub

