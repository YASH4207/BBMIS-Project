VERSION 5.00
Begin VB.Form Frmadminmodule 
   BorderStyle     =   0  'None
   Caption         =   "Admin Module"
   ClientHeight    =   12960
   ClientLeft      =   0
   ClientTop       =   -315
   ClientWidth     =   22395
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   12960
   ScaleWidth      =   22395
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton BtnShowAll 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "SHOW ALL INFO"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   20160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   11760
      Width           =   2655
   End
   Begin VB.CommandButton BtnDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9000
      Width           =   2655
   End
   Begin VB.CommandButton BtnUpdate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8160
      UseMaskColor    =   -1  'True
      Width           =   2655
   End
   Begin VB.CommandButton ContactUs 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "CONTACT US"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10680
      Width           =   2655
   End
   Begin VB.CommandButton btnABTUs 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "ABOUT US"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9840
      Width           =   2655
   End
   Begin VB.CommandButton btnREPORT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOG OUT"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   11520
      Width           =   2655
   End
   Begin VB.CommandButton btnSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Caption         =   "Frame1"
      ForeColor       =   &H008080FF&
      Height          =   19815
      Left            =   -240
      TabIndex        =   0
      Top             =   -120
      Width           =   23415
      Begin VB.CommandButton BtnPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7440
         Width           =   2655
      End
      Begin VB.CommandButton btnADD 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5760
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "HOME "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   0
         Left            =   8955
         TabIndex        =   16
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DESIGNED AND DEVOLEPED BY YASH DARNE"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   360
         TabIndex        =   15
         Top             =   12480
         Width           =   8835
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BLOOD BANK INFORMATION SYSTEM"
         BeginProperty Font 
            Name            =   "Yu Gothic UI"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1290
         Index           =   1
         Left            =   3360
         TabIndex        =   14
         Top             =   0
         Width           =   16695
      End
      Begin VB.Image ImgLogo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   360
         Picture         =   "Frm Home.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "YASH DARNE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Top             =   5160
         Width           =   3375
      End
      Begin VB.Image ImageMe 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3495
         Left            =   960
         Picture         =   "Frm Home.frx":18DB6
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2970
      End
      Begin VB.Label LabelDecription 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm Home.frx":3F6373
         BeginProperty Font 
            Name            =   "Navigo"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4005
         Left            =   4920
         TabIndex        =   11
         Top             =   4680
         Width           =   8595
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "GUIDED BY: PROF. V. J. DESHMUKH"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   4920
         TabIndex        =   9
         Top             =   9120
         Width           =   4815
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   12990
         Left            =   240
         Picture         =   "Frm Home.frx":3F64AB
         Stretch         =   -1  'True
         Top             =   120
         Width           =   23040
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Be the reason for someone's heartbeat."
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   $"Frm Home.frx":42A509
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6240
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
End
Attribute VB_Name = "Frmadminmodule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnABTUs_Click()
Frmadminmodule.Hide
FrmAbout.Show
End Sub
Private Sub btnADD_Click()
Frmadminmodule.Hide
FrmAdminHome.Show
End Sub

Private Sub BtnDelete_Click()
Frmadminmodule.Hide
FrmDelete.Show

End Sub

Private Sub BtnPrint_Click()
Frmadminmodule.Show
FrmPrint.Show
End Sub

Private Sub btnREPORT_Click()
Frmadminmodule.Hide
FrmAdminLogin.Show
End Sub

Private Sub BtnShowAll_Click()
Frmadminmodule.Hide
FrmShowAll.Show
End Sub

Private Sub BtnUpdate_Click()
Frmadminmodule.Hide
FrmUpdate.Show

End Sub

Private Sub ContactUs_Click()
Frmadminmodule.Hide
FrmContactUs.Show
End Sub

Private Sub btnSearch_Click()
Frmadminmodule.Hide
FrmSearch.Show
End Sub
