VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Intro"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4440
      Top             =   120
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
      Left            =   5280
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.yazanmarkabi.webs.com/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   360
      Index           =   4
      Left            =   720
      TabIndex        =   4
      Top             =   2640
      Width           =   5460
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "zaidmarkabi@yahoo.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   360
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   3810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Powered by  Zaid Markabi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   3840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Based on CheatCC.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   360
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find your game via internet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   4140
   End
   Begin VB.Image Image1 
      Height          =   1185
      Index           =   6
      Left            =   480
      Picture         =   "frmSplash.frx":0000
      Top             =   360
      Width           =   6255
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   240
      Picture         =   "frmSplash.frx":1829E
      Top             =   3240
      Width           =   1530
   End
   Begin VB.Image Image1 
      Height          =   1305
      Index           =   5
      Left            =   3360
      Picture         =   "frmSplash.frx":1ACFC
      Top             =   2640
      Width           =   3450
   End
   Begin VB.Image Image1 
      Height          =   1305
      Index           =   4
      Left            =   0
      Picture         =   "frmSplash.frx":2986A
      Top             =   2640
      Width           =   3450
   End
   Begin VB.Image Image1 
      Height          =   1305
      Index           =   3
      Left            =   3360
      Picture         =   "frmSplash.frx":383D8
      Top             =   1320
      Width           =   3450
   End
   Begin VB.Image Image1 
      Height          =   1305
      Index           =   2
      Left            =   3360
      Picture         =   "frmSplash.frx":46F46
      Top             =   0
      Width           =   3450
   End
   Begin VB.Image Image1 
      Height          =   1305
      Index           =   1
      Left            =   0
      Picture         =   "frmSplash.frx":55AB4
      Top             =   1320
      Width           =   3450
   End
   Begin VB.Image Image1 
      Height          =   1305
      Index           =   0
      Left            =   0
      Picture         =   "frmSplash.frx":64622
      Top             =   0
      Width           =   3450
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Label2(5).Visible = True
frmMain.Show
Timer1.Enabled = False
End Sub
