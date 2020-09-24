VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00CCCCCC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Cheats Codes Center Online"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox CheatsText 
      Appearance      =   0  'Flat
      Height          =   6975
      Left            =   4920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   1440
      Width           =   6135
   End
   Begin VB.PictureBox BtnEnbl 
      Height          =   495
      Index           =   1
      Left            =   720
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   75
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BtnEnbl 
      Height          =   495
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":077A
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   74
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BtnEnbl 
      Height          =   495
      Index           =   2
      Left            =   720
      Picture         =   "Form1.frx":0EF4
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   73
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox BtnEnbl 
      Height          =   495
      Index           =   3
      Left            =   240
      Picture         =   "Form1.frx":22CE
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   72
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PlatBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   3
      Left            =   240
      Picture         =   "Form1.frx":36A8
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   61
      Top             =   3120
      Width           =   1215
      Begin VB.Label PlatformsLBL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   62
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.PictureBox PlatBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   240
      Picture         =   "Form1.frx":4A82
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   63
      Top             =   2760
      Width           =   1215
      Begin VB.Label PlatformsLBL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PS2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   64
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.PictureBox PlatBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   240
      Picture         =   "Form1.frx":5E5C
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   65
      Top             =   2400
      Width           =   1215
      Begin VB.Label PlatformsLBL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PS3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   66
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   26
      Left            =   2760
      Picture         =   "Form1.frx":7236
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   3480
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   26
         Left            =   0
         TabIndex        =   8
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   25
      Left            =   2280
      Picture         =   "Form1.frx":79B0
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   3480
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   25
         Left            =   0
         TabIndex        =   10
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   24
      Left            =   1800
      Picture         =   "Form1.frx":812A
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   3480
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   24
         Left            =   0
         TabIndex        =   12
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   23
      Left            =   4200
      Picture         =   "Form1.frx":88A4
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   3120
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   23
         Left            =   0
         TabIndex        =   14
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   22
      Left            =   3720
      Picture         =   "Form1.frx":901E
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   3120
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   22
         Left            =   0
         TabIndex        =   16
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   21
      Left            =   3240
      Picture         =   "Form1.frx":9798
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   17
      Top             =   3120
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   21
         Left            =   0
         TabIndex        =   18
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   20
      Left            =   2760
      Picture         =   "Form1.frx":9F12
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   19
      Top             =   3120
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   20
         Left            =   0
         TabIndex        =   20
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   19
      Left            =   2280
      Picture         =   "Form1.frx":A68C
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   21
      Top             =   3120
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   19
         Left            =   0
         TabIndex        =   22
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   18
      Left            =   1800
      Picture         =   "Form1.frx":AE06
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   23
      Top             =   3120
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   18
         Left            =   0
         TabIndex        =   24
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   17
      Left            =   4200
      Picture         =   "Form1.frx":B580
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   25
      Top             =   2760
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   17
         Left            =   0
         TabIndex        =   26
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   16
      Left            =   3720
      Picture         =   "Form1.frx":BCFA
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   27
      Top             =   2760
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   16
         Left            =   0
         TabIndex        =   28
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   15
      Left            =   3240
      Picture         =   "Form1.frx":C474
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   29
      Top             =   2760
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   15
         Left            =   0
         TabIndex        =   30
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   14
      Left            =   2760
      Picture         =   "Form1.frx":CBEE
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   31
      Top             =   2760
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   14
         Left            =   0
         TabIndex        =   32
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   13
      Left            =   2280
      Picture         =   "Form1.frx":D368
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   33
      Top             =   2760
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   13
         Left            =   0
         TabIndex        =   34
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   12
      Left            =   1800
      Picture         =   "Form1.frx":DAE2
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   35
      Top             =   2760
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   36
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   11
      Left            =   4200
      Picture         =   "Form1.frx":E25C
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   37
      Top             =   2400
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   38
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   10
      Left            =   3720
      Picture         =   "Form1.frx":E9D6
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   39
      Top             =   2400
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   40
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   9
      Left            =   3240
      Picture         =   "Form1.frx":F150
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   41
      Top             =   2400
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   42
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   8
      Left            =   2760
      Picture         =   "Form1.frx":F8CA
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   43
      Top             =   2400
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   44
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   7
      Left            =   2280
      Picture         =   "Form1.frx":10044
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   45
      Top             =   2400
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   46
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   6
      Left            =   1800
      Picture         =   "Form1.frx":107BE
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   47
      Top             =   2400
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   48
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   5
      Left            =   4200
      Picture         =   "Form1.frx":10F38
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   49
      Top             =   2040
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   50
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   4
      Left            =   3720
      Picture         =   "Form1.frx":116B2
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   51
      Top             =   2040
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   52
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   3
      Left            =   3240
      Picture         =   "Form1.frx":11E2C
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   53
      Top             =   2040
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   54
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   2760
      Picture         =   "Form1.frx":125A6
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   55
      Top             =   2040
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   56
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   2280
      Picture         =   "Form1.frx":12D20
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   57
      Top             =   2040
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   58
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.ComboBox PlatformsDt 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "Form1.frx":1349A
      Left            =   360
      List            =   "Form1.frx":134AA
      Style           =   2  'Dropdown List
      TabIndex        =   70
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox PlatformsDt 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      ItemData        =   "Form1.frx":13532
      Left            =   840
      List            =   "Form1.frx":13542
      Style           =   2  'Dropdown List
      TabIndex        =   69
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PlatBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":135BA
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   67
      Top             =   2040
      Width           =   1215
      Begin VB.Label PlatformsLBL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PSP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   68
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.PictureBox LetterBox 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   1800
      Picture         =   "Form1.frx":14994
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   59
      Top             =   2040
      Width           =   495
      Begin VB.Label Letter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0-9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   60
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.TextBox HeaderOfGames 
      Height          =   375
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":1510E
      Top             =   4800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ListBox GamesList 
      Height          =   1620
      Left            =   1560
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00858585&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   4935
      TabIndex        =   71
      Top             =   1500
      Width           =   4935
      Begin VB.Label GamesCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "All games : 0"
         Height          =   195
         Left            =   3720
         TabIndex        =   79
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Platforms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   78
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First letter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   1800
         TabIndex        =   77
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.ListBox GameNList 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4125
      Left            =   120
      TabIndex        =   76
      Top             =   4320
      Width           =   4695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   5
      Left            =   9600
      TabIndex        =   82
      Top             =   120
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Line Line1 
      X1              =   4920
      X2              =   11040
      Y1              =   8415
      Y2              =   8415
   End
   Begin VB.Label Stat 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   5
      Top             =   8520
      Width           =   4935
   End
   Begin VB.Label Stat 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   4
      Top             =   8520
      Width           =   4215
   End
   Begin VB.Label Stat 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ready"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cheats Codes Games Database"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   1
      Left            =   5520
      TabIndex        =   81
      Top             =   645
      Width           =   5355
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find your game via internet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   0
      Left            =   1920
      TabIndex        =   80
      Top             =   300
      Width           =   4725
   End
   Begin VB.Image Image5 
      Height          =   90
      Left            =   60
      Picture         =   "Form1.frx":15134
      Top             =   4185
      Width           =   4815
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   0
      Picture         =   "Form1.frx":1680E
      Top             =   1200
      Width           =   4200
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   4920
      Picture         =   "Form1.frx":1B080
      Top             =   1200
      Width           =   3090
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   0
      Picture         =   "Form1.frx":1D516
      Top             =   0
      Width           =   11100
   End
   Begin VB.Label Stat 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   8520
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   330
      Index           =   0
      Left            =   4200
      Picture         =   "Form1.frx":4826C
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' this program can read Cheats Codes from [ Cheatcc.com ]
' Written by , Zaid Markabi
' Website : http://yazanmarkabi.webs.com/
' Email : ZaidMarkabi@yahoo.com

Function GetFile(URL As String, SaveTo As String) As String
On Error GoTo Err:
Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")

 Dim d() As Byte

 WinHttpReq.Open "GET", URL, False
 WinHttpReq.Send

 GetFile = WinHttpReq.StatusText

Open SaveTo For Output As #1
Close #1

 Open SaveTo For Binary As #1
 d() = WinHttpReq.ResponseBody
 Put #1, 1, d()
 Close #1
 
 Set WinHttpReq = Nothing
Exit Function
Err:
MsgBox "Error , Try again..."
Err.Clear
Resume Next
End Function

Private Sub Form_Load()
PlatformsLBL_Click (0)
Unload frmSplash
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub GameNList_DblClick()
Stat(1).Caption = " Connecting ..."
Stat(3).Caption = " " + PlatformsDt(1).List(PlatformsDt(1).ListIndex) + GamesList.List(GameNList.ListIndex)
If Len(Stat(3).Caption) > 60 Then
Stat(3).Caption = Left(Stat(3).Caption, 54) + "~.html"
End If
DoEvents

GetFile PlatformsDt(1).List(PlatformsDt(1).ListIndex) + GamesList.List(GameNList.ListIndex), App.Path + "\Temp.txt"

Stat(2).Caption = " " + GameNList.List(GameNList.ListIndex)
Stat(1).Caption = " Reading ..."
DoEvents

Dim ByteTemp As Byte
Dim TextTemp As String
Dim i As Long
Open App.Path + "\Temp.txt" For Binary As #1
Do While EOF(1) = False
i = i + 1
Get #1, i, ByteTemp
TextTemp = TextTemp + Chr(ByteTemp)
Loop
Close #1
CheatsText.Text = TextTemp

i = InStr(1, TextTemp, "<dt><li>")
If i = 0 Then Exit Sub
TextTemp = Right(TextTemp, Len(TextTemp) - i - 7)

CheatsText.Text = TextTemp

i = InStr(1, TextTemp, "</div>")
If i = 0 Then Exit Sub
TextTemp = Left(TextTemp, i - 1)

TextTemp = Trim(TextTemp)

Open App.Path + "\Temp.txt" For Output As #1
Close #1

TextTemp = "• • " + TextTemp
TextTemp = Replace(TextTemp, "<b>", "")
TextTemp = Replace(TextTemp, "</b>", "")
TextTemp = Replace(TextTemp, "<B>", "")
TextTemp = Replace(TextTemp, "</B>", "")
TextTemp = Replace(TextTemp, "<p>", "")
TextTemp = Replace(TextTemp, "</p>", "")
TextTemp = Replace(TextTemp, "<P>", "")
TextTemp = Replace(TextTemp, "</P>", "")
TextTemp = Replace(TextTemp, "<br>", "")
TextTemp = Replace(TextTemp, "</br>", "")
TextTemp = Replace(TextTemp, "<i>", "")
TextTemp = Replace(TextTemp, "</i>", "")
TextTemp = Replace(TextTemp, "<ul>", "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
TextTemp = Replace(TextTemp, "</ul>", "")
TextTemp = Replace(TextTemp, "<dt>", "• ")
TextTemp = Replace(TextTemp, "</dt>", "")
TextTemp = Replace(TextTemp, "<li>", "• ")
TextTemp = Replace(TextTemp, "</li>", "")
TextTemp = Replace(TextTemp, "<hr>", "")
TextTemp = Replace(TextTemp, "</hr>", "")
TextTemp = Replace(TextTemp, "<center> ", "• • ")
TextTemp = Replace(TextTemp, "<center>", "")
TextTemp = Replace(TextTemp, "</center>", "")
TextTemp = Replace(TextTemp, "<table border=5>", "")
TextTemp = Replace(TextTemp, "</table>", "")
TextTemp = Replace(TextTemp, "<TR><TD>", "• ")
TextTemp = Replace(TextTemp, "<tr><td>", "• ")
TextTemp = Replace(TextTemp, "<TR>", "")
TextTemp = Replace(TextTemp, "</TR>", "")
TextTemp = Replace(TextTemp, "<tr>", "")
TextTemp = Replace(TextTemp, "</tr>", "")
TextTemp = Replace(TextTemp, "<TD>", "")
TextTemp = Replace(TextTemp, "</TD>", "")
TextTemp = Replace(TextTemp, "<td>", "")
TextTemp = Replace(TextTemp, "</td>", "")
TextTemp = Replace(TextTemp, "<table border=1 cellpadding=5>", "")
TextTemp = Replace(TextTemp, "<font color=", "")
TextTemp = Replace(TextTemp, "#FF0000", "")
TextTemp = Replace(TextTemp, "</font>", "")

TextTemp = Left(TextTemp, Len(TextTemp) - 8)

DoEvents

Open App.Path + "\Temp.txt" For Output As #1
Close #1

Open App.Path + "\Temp.txt" For Binary As #1
Put #1, 1, TextTemp
Close #1

CheatsText.Text = TextTemp

Stat(1).Caption = " Ready"
End Sub

Private Sub Letter_Click(Index As Integer)
Dim i As Long

For i = 0 To LetterBox.Count - 1
LetterBox(i).Picture = BtnEnbl(1).Picture
Next
LetterBox(Index).Picture = BtnEnbl(0).Picture

Dim x As String
Dim ReadTemp As String

x = Letter(Index).Caption
If x = "0-9" Then x = "0"

Stat(1).Caption = " Connecting ..."
Stat(3).Caption = " " + PlatformsDt(0).List(PlatformsDt(0).ListIndex) + x + ".html"
DoEvents

GetFile PlatformsDt(0).List(PlatformsDt(0).ListIndex) + x + ".html", App.Path + "\Temp.txt"

Stat(1).Caption = " Reading ..."
DoEvents

Dim ByteTemp As Byte
Dim TextTemp As String
Open App.Path + "\Temp.txt" For Binary As #1
Do While EOF(1) = False
i = i + 1
Get #1, i, ByteTemp
TextTemp = TextTemp + Chr(ByteTemp)
Loop
Close #1
ReadTemp = TextTemp

i = InStr(1, ReadTemp, HeaderOfGames.Text)
If i = 0 Then Exit Sub
ReadTemp = Right(ReadTemp, Len(ReadTemp) - i - Len(HeaderOfGames.Text) + 1)

i = InStr(1, ReadTemp, "</div>")
If i = 0 Then Exit Sub
ReadTemp = Left(ReadTemp, i - 1)

ReadTemp = Trim(ReadTemp) + Space(6)

' Exit Sub

Dim GetLink As String
Dim FindLink As Boolean
Dim GetName As String
Dim FindName As Boolean

GamesList.Clear
GameNList.Clear

For i = 1 To Len(ReadTemp) - 5
x = Mid(ReadTemp, i, 5)

If FindLink = True Then
GetLink = GetLink + Mid(ReadTemp, i, 1)
End If

If FindName = True Then
GetName = GetName + Mid(ReadTemp, i, 1)
End If

If x = "href=" Then FindLink = True

If x = "</a><" Then
FindName = False
GameNList.AddItem Left(GetName, Len(GetName) - 1)
GetName = ""
End If

If Left(x, 1) = ">" And FindLink = True Then
FindLink = False
FindName = True
GetLink = Right(GetLink, Len(GetLink) - 5)
GetLink = Left(GetLink, Len(GetLink) - 2)
GamesList.AddItem GetLink
GetLink = ""
End If

Next

Stat(1).Caption = " Ready"
DoEvents

GamesCount.Caption = "All games : " + Format(GamesList.ListCount)
End Sub

Private Sub PlatformsLBL_Click(Index As Integer)
Dim i As Integer

For i = 0 To PlatBox.Count - 1
PlatBox(i).Picture = BtnEnbl(2).Picture
Next
PlatBox(Index).Picture = BtnEnbl(3).Picture

PlatformsDt(0).ListIndex = Index
PlatformsDt(1).ListIndex = Index

Letter_Click (0)
End Sub

Private Sub Stat_Change(Index As Integer)
If Stat(1).Caption = " Connecting ..." Then Label2(5).Visible = True
If Stat(1).Caption = " Ready" Then Label2(5).Visible = False
End Sub
