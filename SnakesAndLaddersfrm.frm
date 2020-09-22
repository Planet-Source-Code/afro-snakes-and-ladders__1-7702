VERSION 5.00
Begin VB.Form SnakesAndLadders 
   Caption         =   "Snakes and Ladders"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8850
   Icon            =   "SnakesAndLaddersfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Interval        =   250
      Left            =   8160
      Top             =   360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8160
      Top             =   720
   End
   Begin VB.Timer tmrAni 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8160
      Top             =   1080
   End
   Begin VB.Timer tmrMove 
      Interval        =   1
      Left            =   8160
      Top             =   2160
   End
   Begin VB.PictureBox MyIcon5 
      Height          =   375
      Left            =   7800
      Picture         =   "SnakesAndLaddersfrm.frx":0442
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   90
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox MyIcon4 
      Height          =   375
      Left            =   7800
      Picture         =   "SnakesAndLaddersfrm.frx":0884
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   89
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox MyIcon3 
      Height          =   375
      Left            =   7800
      Picture         =   "SnakesAndLaddersfrm.frx":0CC6
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   88
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox MyIcon2 
      Height          =   375
      Left            =   7800
      Picture         =   "SnakesAndLaddersfrm.frx":1108
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   87
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox MyIcon1 
      Height          =   375
      Left            =   7800
      Picture         =   "SnakesAndLaddersfrm.frx":154A
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   86
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   7800
      Top             =   4560
   End
   Begin VB.Timer tmrSeconds 
      Interval        =   1000
      Left            =   8160
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8160
      Top             =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Roll"
      Height          =   375
      Left            =   7680
      TabIndex        =   80
      Top             =   3960
      Width           =   975
   End
   Begin VB.PictureBox Square39 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   1680
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   78
      Top             =   480
      Width           =   1455
      Begin VB.Shape Shape39 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "39"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   79
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square38 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   76
      Top             =   480
      Width           =   1455
      Begin VB.Shape Shape38 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "38"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   77
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square37 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   74
      Top             =   480
      Width           =   1455
      Begin VB.Shape Shape37 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "37"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   75
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square36 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   6000
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   72
      Top             =   480
      Width           =   1455
      Begin VB.Shape Shape36 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "36"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   73
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square40 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   70
      Top             =   480
      Width           =   1455
      Begin VB.Shape Shape40 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   71
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square35 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   68
      Top             =   1560
      Width           =   1455
      Begin VB.Shape Shape35 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "35"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   69
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square31 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   6000
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   66
      Top             =   1560
      Width           =   1455
      Begin VB.Shape Shape31 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   67
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square32 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   64
      Top             =   1560
      Width           =   1455
      Begin VB.Shape Shape32 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   65
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square33 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   62
      Top             =   1560
      Width           =   1455
      Begin VB.Shape Shape33 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "33"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   63
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square34 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   1680
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   60
      Top             =   1560
      Width           =   1455
      Begin VB.Shape Shape34 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "34"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   61
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Square29 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   1680
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   29
      Top             =   2640
      Width           =   1455
      Begin VB.Shape Shape29 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   58
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square28 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   28
      Top             =   2640
      Width           =   1455
      Begin VB.Shape Shape28 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   57
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square27 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   27
      Top             =   2640
      Width           =   1455
      Begin VB.Shape Shape27 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   56
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square26 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   6000
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   26
      Top             =   2640
      Width           =   1455
      Begin VB.Shape Shape26 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   55
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square30 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   25
      Top             =   2640
      Width           =   1455
      Begin VB.Shape Shape30 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   59
         Top             =   0
         Width           =   1185
      End
   End
   Begin VB.PictureBox Square24 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   1680
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   24
      Top             =   3720
      Width           =   1455
      Begin VB.Shape Shape24 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   53
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Square23 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   23
      Top             =   3720
      Width           =   1455
      Begin VB.Shape Shape23 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   52
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square22 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   22
      Top             =   3720
      Width           =   1455
      Begin VB.Shape Shape22 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   51
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square21 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   6000
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   21
      Top             =   3720
      Width           =   1455
      Begin VB.Shape Shape21 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   50
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square25 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   20
      Top             =   3720
      Width           =   1455
      Begin VB.Shape Shape25 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   54
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square19 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   1680
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   19
      Top             =   4800
      Width           =   1455
      Begin VB.Shape Shape19 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   48
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square18 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   18
      Top             =   4800
      Width           =   1455
      Begin VB.Shape Shape18 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   47
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square17 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   17
      Top             =   4800
      Width           =   1455
      Begin VB.Shape Shape17 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   46
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square16 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   6000
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   16
      Top             =   4800
      Width           =   1455
      Begin VB.Shape Shape16 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   45
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square20 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   15
      Top             =   4800
      Width           =   1455
      Begin VB.Shape Shape20 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   49
         Top             =   0
         Width           =   1185
      End
   End
   Begin VB.PictureBox Square14 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   1680
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   14
      Top             =   5880
      Width           =   1455
      Begin VB.Shape Shape14 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   43
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Square13 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   13
      Top             =   5880
      Width           =   1455
      Begin VB.Shape Shape13 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "GhoulyCaps"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1050
         Left            =   240
         TabIndex        =   42
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.PictureBox Square12 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   12
      Top             =   5880
      Width           =   1455
      Begin VB.Shape Shape12 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   41
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square11 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   6000
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   11
      Top             =   5880
      Width           =   1455
      Begin VB.Shape Shape11 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   40
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square15 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   10
      Top             =   5880
      Width           =   1455
      Begin VB.Shape Shape15 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   44
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox Square9 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   1680
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   6960
      Width           =   1455
      Begin VB.Shape Shape9 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   360
         TabIndex        =   38
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Square8 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   8
      Top             =   6960
      Width           =   1455
      Begin VB.Shape Shape8 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   360
         TabIndex        =   37
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Square7 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   6960
      Width           =   1455
      Begin VB.Shape Shape7 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   360
         TabIndex        =   36
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Square6 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   6000
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   6960
      Width           =   1455
      Begin VB.Shape Shape6 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   360
         TabIndex        =   35
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Square10 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   5
      Top             =   6960
      Width           =   1455
      Begin VB.Shape Shape10 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   39
         Top             =   0
         Width           =   1185
      End
   End
   Begin VB.PictureBox Square4 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   1680
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   8040
      Width           =   1455
      Begin VB.Shape Shape4 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   360
         TabIndex        =   33
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox Square3 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   8040
      Width           =   1455
      Begin VB.Shape Shape3 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   360
         TabIndex        =   32
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Square2 
      BackColor       =   &H000000FF&
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   8040
      Width           =   1455
      Begin VB.Shape Shape2 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   360
         TabIndex        =   31
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Square5 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   8040
      Width           =   1455
      Begin VB.Shape Shape5 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   360
         TabIndex        =   34
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Square1 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   6000
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   8040
      Width           =   1455
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H0000FF00&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Left            =   480
         Shape           =   3  'Circle
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   360
         TabIndex        =   30
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.Label Seconds 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   7920
      TabIndex        =   85
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Roll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   7680
      TabIndex        =   84
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label MyPos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3960
      TabIndex        =   83
      Top             =   120
      Width           =   120
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Position:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   82
      Top             =   120
      Width           =   1425
   End
   Begin VB.Shape Dice2 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   7860
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Dice5 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8220
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Dice3 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   7860
      Shape           =   3  'Circle
      Top             =   3285
      Width           =   135
   End
   Begin VB.Shape Dice6 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8220
      Shape           =   3  'Circle
      Top             =   3285
      Width           =   135
   End
   Begin VB.Shape Dice4 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   7860
      Shape           =   3  'Circle
      Top             =   3435
      Width           =   135
   End
   Begin VB.Shape Dice7 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8220
      Shape           =   3  'Circle
      Top             =   3435
      Width           =   135
   End
   Begin VB.Shape Dice1 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8025
      Shape           =   3  'Circle
      Top             =   3285
      Width           =   135
   End
   Begin VB.Label lblDice 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8010
      TabIndex        =   81
      Top             =   2760
      Width           =   195
   End
   Begin VB.Shape Shape48 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   7800
      Top             =   3045
      Width           =   615
   End
End
Attribute VB_Name = "SnakesAndLadders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DiceRoll()
Dim Dice
Dice = Int((6 * Rnd) + 1)

Roll = Roll + 1
lblDice = Dice


If Dice = 1 Then
    Dice1.Visible = True
    Dice2.Visible = False
    Dice3.Visible = False
    Dice4.Visible = False
    Dice5.Visible = False
    Dice6.Visible = False
    Dice7.Visible = False
ElseIf Dice = 2 Then
    Dice1.Visible = False
    Dice2.Visible = False
    Dice3.Visible = False
    Dice4.Visible = True
    Dice5.Visible = True
    Dice6.Visible = False
    Dice7.Visible = False
ElseIf Dice = 3 Then
    Dice1.Visible = True
    Dice2.Visible = False
    Dice3.Visible = False
    Dice4.Visible = True
    Dice5.Visible = True
    Dice6.Visible = False
    Dice7.Visible = False
ElseIf Dice = 4 Then
    Dice1.Visible = False
    Dice2.Visible = True
    Dice3.Visible = False
    Dice4.Visible = True
    Dice5.Visible = True
    Dice6.Visible = False
    Dice7.Visible = True
ElseIf Dice = 5 Then
    Dice1.Visible = True
    Dice2.Visible = True
    Dice3.Visible = False
    Dice4.Visible = True
    Dice5.Visible = True
    Dice6.Visible = False
    Dice7.Visible = True
ElseIf Dice = 6 Then
    Dice1.Visible = False
    Dice2.Visible = True
    Dice3.Visible = True
    Dice4.Visible = True
    Dice5.Visible = True
    Dice6.Visible = True
    Dice7.Visible = True
End If


MyPos = MyPos + Dice

If MyPos = 41 Then
    MyPos = 1
ElseIf MyPos = 42 Then
    MyPos = 2
ElseIf MyPos = 43 Then
    MyPos = 3
ElseIf MyPos = 44 Then
    MyPos = 4
ElseIf MyPos = 45 Then
    MyPos = 5
End If
End Sub


Private Sub Command1_Click()
Timer1.Enabled = True
tmrMove.Enabled = True
tmrAni.Enabled = True
Timer3.Enabled = True



End Sub

Private Sub Form_Load()
Randomize Timer
MsgBox "You must try and get to square number 40 to complete the game"
Me.Icon = MyIcon1
MyPos = 0

Dice1.Visible = False
Dice2.Visible = False
Dice3.Visible = False
Dice4.Visible = False
Dice5.Visible = False
Dice6.Visible = False
Dice7.Visible = False

End Sub


Private Sub Timer1_Timer()

Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False
Shape8.Visible = False
Shape9.Visible = False
Shape10.Visible = False
Shape11.Visible = False
Shape12.Visible = False
Shape13.Visible = False
Shape14.Visible = False
Shape15.Visible = False
Shape16.Visible = False
Shape17.Visible = False
Shape18.Visible = False
Shape19.Visible = False
Shape20.Visible = False
Shape21.Visible = False
Shape22.Visible = False
Shape23.Visible = False
Shape24.Visible = False
Shape25.Visible = False
Shape26.Visible = False
Shape27.Visible = False
Shape28.Visible = False
Shape29.Visible = False
Shape30.Visible = False
Shape31.Visible = False
Shape32.Visible = False
Shape33.Visible = False
Shape34.Visible = False
Shape35.Visible = False
Shape36.Visible = False
Shape37.Visible = False
Shape38.Visible = False
Shape39.Visible = False
Shape40.Visible = False

If MyPos = 1 Then
    Shape1.Visible = True
ElseIf MyPos = 2 Then
    Shape2.Visible = True
ElseIf MyPos = 3 Then
    Shape3.Visible = True
ElseIf MyPos = 4 Then
    Shape4.Visible = True
ElseIf MyPos = 5 Then
    Shape5.Visible = True
ElseIf MyPos = 6 Then
    Shape6.Visible = True
ElseIf MyPos = 7 Then
    Shape7.Visible = True
ElseIf MyPos = 8 Then
    Shape8.Visible = True
ElseIf MyPos = 9 Then
    Shape9.Visible = True
ElseIf MyPos = 10 Then
    Shape10.Visible = True
ElseIf MyPos = 11 Then
    Shape11.Visible = True
ElseIf MyPos = 12 Then
    Shape12.Visible = True
ElseIf MyPos = 13 Then
    Shape13.Visible = True
ElseIf MyPos = 14 Then
    Shape14.Visible = True
ElseIf MyPos = 15 Then
    Shape15.Visible = True
ElseIf MyPos = 16 Then
    Shape16.Visible = True
ElseIf MyPos = 17 Then
    Shape17.Visible = True
ElseIf MyPos = 18 Then
    Shape18.Visible = True
ElseIf MyPos = 19 Then
    Shape19.Visible = True
ElseIf MyPos = 20 Then
    Shape20.Visible = True
ElseIf MyPos = 21 Then
    Shape21.Visible = True
ElseIf MyPos = 22 Then
    Shape22.Visible = True
ElseIf MyPos = 23 Then
    Shape23.Visible = True
ElseIf MyPos = 24 Then
    Shape24.Visible = True
ElseIf MyPos = 25 Then
    Shape25.Visible = True
ElseIf MyPos = 26 Then
    Shape26.Visible = True
ElseIf MyPos = 27 Then
    Shape27.Visible = True
ElseIf MyPos = 28 Then
    Shape28.Visible = True
ElseIf MyPos = 29 Then
    Shape29.Visible = True
ElseIf MyPos = 30 Then
    Shape30.Visible = True
ElseIf MyPos = 31 Then
    Shape31.Visible = True
ElseIf MyPos = 32 Then
    Shape32.Visible = True
ElseIf MyPos = 33 Then
    Shape33.Visible = True
ElseIf MyPos = 34 Then
    Shape34.Visible = True
ElseIf MyPos = 35 Then
    Shape35.Visible = True
ElseIf MyPos = 36 Then
    Shape36.Visible = True
ElseIf MyPos = 37 Then
    Shape37.Visible = True
ElseIf MyPos = 38 Then
    Shape38.Visible = True
ElseIf MyPos = 39 Then
    Shape39.Visible = True
ElseIf MyPos = 40 Then
    Shape40.Visible = True
End If
End Sub


Private Sub Timer2_Timer()
If Me.Icon = MyIcon1 Then
    Me.Icon = MyIcon2
    Exit Sub
ElseIf Me.Icon = MyIcon2 Then
    Me.Icon = MyIcon3
    Exit Sub
ElseIf Me.Icon = MyIcon3 Then
    Me.Icon = MyIcon4
    Exit Sub
ElseIf Me.Icon = MyIcon4 Then
    Me.Icon = MyIcon5
    Exit Sub
ElseIf Me.Icon = MyIcon5 Then
    Me.Icon = MyIcon1
    Exit Sub
End If
End Sub

Private Sub Timer3_Timer()
tmrAni.Enabled = False
DiceRoll
tmrSeconds.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
If Square1.BackColor = vbRed Then
    Square1.BackColor = vbBlue
    Square2.BackColor = vbRed
    Square3.BackColor = vbBlue
    Square4.BackColor = vbRed
    Square5.BackColor = vbBlue
    Square6.BackColor = vbRed
    Square7.BackColor = vbBlue
    Square8.BackColor = vbRed
    Square9.BackColor = vbBlue
    Square10.BackColor = vbRed
    Square11.BackColor = vbBlue
    Square12.BackColor = vbRed
    Square13.BackColor = vbBlue
    Square14.BackColor = vbRed
    Square15.BackColor = vbBlue
    Square16.BackColor = vbRed
    Square17.BackColor = vbBlue
    Square18.BackColor = vbRed
    Square19.BackColor = vbBlue
    Square20.BackColor = vbRed
    Square21.BackColor = vbBlue
    Square22.BackColor = vbRed
    Square23.BackColor = vbBlue
    Square24.BackColor = vbRed
    Square25.BackColor = vbBlue
    Square26.BackColor = vbRed
    Square27.BackColor = vbBlue
    Square28.BackColor = vbRed
    Square29.BackColor = vbBlue
    Square30.BackColor = vbRed
    Square31.BackColor = vbBlue
    Square32.BackColor = vbRed
    Square33.BackColor = vbBlue
    Square34.BackColor = vbRed
    Square35.BackColor = vbBlue
    Square36.BackColor = vbRed
    Square37.BackColor = vbBlue
    Square38.BackColor = vbRed
    Square39.BackColor = vbBlue
    Square40.BackColor = vbRed
    
    Label1.ForeColor = vbRed
    Label2.ForeColor = vbBlue
    Label3.ForeColor = vbRed
    Label4.ForeColor = vbBlue
    Label5.ForeColor = vbRed
    Label6.ForeColor = vbBlue
    Label7.ForeColor = vbRed
    Label8.ForeColor = vbBlue
    Label9.ForeColor = vbRed
    Label10.ForeColor = vbBlue
    Label11.ForeColor = vbRed
    Label12.ForeColor = vbBlue
    Label13.ForeColor = vbRed
    Label14.ForeColor = vbBlue
    Label15.ForeColor = vbRed
    Label16.ForeColor = vbBlue
    Label17.ForeColor = vbRed
    Label18.ForeColor = vbBlue
    Label19.ForeColor = vbRed
    Label20.ForeColor = vbBlue
    Label21.ForeColor = vbRed
    Label22.ForeColor = vbBlue
    Label23.ForeColor = vbRed
    Label24.ForeColor = vbBlue
    Label25.ForeColor = vbRed
    Label26.ForeColor = vbBlue
    Label27.ForeColor = vbRed
    Label28.ForeColor = vbBlue
    Label29.ForeColor = vbRed
    Label30.ForeColor = vbBlue
    Label31.ForeColor = vbRed
    Label32.ForeColor = vbBlue
    Label33.ForeColor = vbRed
    Label34.ForeColor = vbBlue
    Label35.ForeColor = vbRed
    Label36.ForeColor = vbBlue
    Label37.ForeColor = vbRed
    Label38.ForeColor = vbBlue
    Label39.ForeColor = vbRed
    Label40.ForeColor = vbBlue
    
Else
    Square1.BackColor = vbRed
    Square2.BackColor = vbBlue
    Square3.BackColor = vbRed
    Square4.BackColor = vbBlue
    Square5.BackColor = vbRed
    Square6.BackColor = vbBlue
    Square7.BackColor = vbRed
    Square8.BackColor = vbBlue
    Square9.BackColor = vbRed
    Square10.BackColor = vbBlue
    Square11.BackColor = vbRed
    Square12.BackColor = vbBlue
    Square13.BackColor = vbRed
    Square14.BackColor = vbBlue
    Square15.BackColor = vbRed
    Square16.BackColor = vbBlue
    Square17.BackColor = vbRed
    Square18.BackColor = vbBlue
    Square19.BackColor = vbRed
    Square20.BackColor = vbBlue
    Square21.BackColor = vbRed
    Square22.BackColor = vbBlue
    Square23.BackColor = vbRed
    Square24.BackColor = vbBlue
    Square25.BackColor = vbRed
    Square26.BackColor = vbBlue
    Square27.BackColor = vbRed
    Square28.BackColor = vbBlue
    Square29.BackColor = vbRed
    Square30.BackColor = vbBlue
    Square31.BackColor = vbRed
    Square32.BackColor = vbBlue
    Square33.BackColor = vbRed
    Square34.BackColor = vbBlue
    Square35.BackColor = vbRed
    Square36.BackColor = vbBlue
    Square37.BackColor = vbRed
    Square38.BackColor = vbBlue
    Square39.BackColor = vbRed
    Square40.BackColor = vbBlue
    
    Label1.ForeColor = vbBlue
    Label2.ForeColor = vbRed
    Label3.ForeColor = vbBlue
    Label4.ForeColor = vbRed
    Label5.ForeColor = vbBlue
    Label6.ForeColor = vbRed
    Label7.ForeColor = vbBlue
    Label8.ForeColor = vbRed
    Label9.ForeColor = vbBlue
    Label10.ForeColor = vbRed
    Label11.ForeColor = vbBlue
    Label12.ForeColor = vbRed
    Label13.ForeColor = vbBlue
    Label14.ForeColor = vbRed
    Label15.ForeColor = vbBlue
    Label16.ForeColor = vbRed
    Label17.ForeColor = vbBlue
    Label18.ForeColor = vbRed
    Label19.ForeColor = vbBlue
    Label20.ForeColor = vbRed
    Label21.ForeColor = vbBlue
    Label22.ForeColor = vbRed
    Label23.ForeColor = vbBlue
    Label24.ForeColor = vbRed
    Label25.ForeColor = vbBlue
    Label26.ForeColor = vbRed
    Label27.ForeColor = vbBlue
    Label28.ForeColor = vbRed
    Label29.ForeColor = vbBlue
    Label30.ForeColor = vbRed
    Label31.ForeColor = vbBlue
    Label32.ForeColor = vbRed
    Label33.ForeColor = vbBlue
    Label34.ForeColor = vbRed
    Label35.ForeColor = vbBlue
    Label36.ForeColor = vbRed
    Label37.ForeColor = vbBlue
    Label38.ForeColor = vbRed
    Label39.ForeColor = vbBlue
    Label40.ForeColor = vbRed
End If
End Sub


Private Sub tmrAni_Timer()

tmrSeconds.Enabled = False

Dice1.Visible = False
Dice2.Visible = False
Dice3.Visible = False
Dice4.Visible = False
Dice5.Visible = False
Dice6.Visible = False
Dice7.Visible = False

Dim Ani

Ani = Int((7 * Rnd) + 1)

If Ani = 1 Then
    Dice1.Visible = True
    Dice2.Visible = True
ElseIf Ani = 2 Then
    Dice7.Visible = True
ElseIf Ani = 3 Then
    Dice7.Visible = True
    Dice4.Visible = True
ElseIf Ani = 4 Then
    Dice6.Visible = True
    Dice3.Visible = True
ElseIf Ani = 5 Then
    Dice1.Visible = True
    Dice2.Visible = True
    Dice3.Visible = True
    Dice4.Visible = True
    Dice5.Visible = True
    Dice6.Visible = True
    Dice7.Visible = True
ElseIf Ani = 6 Then
    Dice1.Visible = True
    Dice5.Visible = True
    Dice6.Visible = True
    Dice7.Visible = True
ElseIf Ani = 7 Then
    Dice3.Visible = True
    Dice4.Visible = True
    Dice5.Visible = True
    Dice6.Visible = True
    Dice7.Visible = True
End If
End Sub

Private Sub tmrMove_Timer()
If MyPos = 13 Then
    Shape13.Visible = True
    MsgBox "Go back to the start!!!"
    MyPos = 1
ElseIf MyPos = 39 Then
    Shape39.Visible = True
    MsgBox "Go back 12 spaces"
    MyPos = MyPos - 12
ElseIf MyPos = 6 Then
    Shape6.Visible = True
    MsgBox "Jump ahead to square 35"
    MyPos = 35
ElseIf MyPos = 30 Then
    Shape30.Visible = True
    MsgBox "Go ahead 2 spaces"
    MyPos = 32
ElseIf MyPos = 40 Then
    Shape40.Visible = True
    MsgBox "It took you " & Seconds & " seconds and " & Roll & " rolls to reach square number 40"
    Unload Me
    End
End If
End Sub

Private Sub tmrSeconds_Timer()
Seconds = Seconds + 1
End Sub


