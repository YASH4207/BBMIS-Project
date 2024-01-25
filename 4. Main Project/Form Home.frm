VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   10335
   ClientLeft      =   2925
   ClientTop       =   1680
   ClientWidth     =   18045
   FillColor       =   &H008080FF&
   ForeColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   18045
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   3015
      Left            =   3360
      TabIndex        =   13
      Top             =   6240
      Width           =   4095
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "Prof. V. J. Deshmukh"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   1920
         Left            =   1560
         TabIndex        =   17
         Top             =   720
         Width           =   3555
      End
      Begin VB.Label Label3 
         BackColor       =   &H008080FF&
         Caption         =   "Guide:"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   3855
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   4095
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H008080FF&
      Caption         =   "Frame6"
      Height          =   7695
      Left            =   7320
      TabIndex        =   12
      Top             =   1560
      Width           =   10935
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   16
         Text            =   "Be the reason for someone's heartbeat."
         Top             =   0
         Width           =   10815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   $"Form Home.frx":0000
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3435
         Left            =   720
         TabIndex        =   15
         Top             =   1440
         Width           =   9285
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "Frame3"
      Height          =   4695
      Left            =   3360
      TabIndex        =   10
      Top             =   1560
      Width           =   3975
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   5175
         Left            =   10200
         TabIndex        =   11
         Top             =   120
         Width           =   15
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   3960
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Image Image2 
         Height          =   4680
         Left            =   0
         Picture         =   "Form Home.frx":0131
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3960
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000011&
      Caption         =   "Frame4"
      Height          =   1335
      Left            =   0
      TabIndex        =   8
      Top             =   9240
      Width           =   18135
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Chalkduster"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   1815
         Left            =   0
         TabIndex        =   9
         Text            =   "DESIGNED AND DEVELOPED BY YASH DARNE"
         Top             =   0
         Width           =   18135
      End
   End
   Begin VB.CommandButton ADMLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Admin login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   1200
      TabIndex        =   7
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Search 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   1200
      TabIndex        =   6
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton ABTUs 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "About Us"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1200
      TabIndex        =   5
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Contact 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000080&
      Caption         =   "Button"
      Height          =   9495
      Left            =   14880
      TabIndex        =   2
      Top             =   10080
      Width           =   3375
      Begin VB.CommandButton Home 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Cancel          =   -1  'True
         Caption         =   "Home"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   1200
         TabIndex        =   3
         Top             =   1920
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Caption         =   "Name"
      Height          =   1815
      Left            =   1320
      TabIndex        =   0
      Top             =   -240
      Width           =   16815
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         X1              =   2040
         X2              =   2040
         Y1              =   360
         Y2              =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "Blood Bank Information System"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1500
         Left            =   1620
         TabIndex        =   1
         Top             =   240
         Width           =   15060
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Home_Click(Index As Integer)
Form1.Show
End Sub

Private Sub ADMLogin_Click(Index As Integer)
Form2.Show
End Sub

Private Sub Label5_Click()

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Search_Click(Index As Integer)
Form3.Show
End Sub

Private Sub ABTUs_Click(Index As Integer)
Form4.Show
End Sub

Private Sub Contact_Click(Index As Integer)
Form5.Show
End Sub

