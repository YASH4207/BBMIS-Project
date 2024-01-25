VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About Us"
   ClientHeight    =   12960
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   22425
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12960
   ScaleWidth      =   22425
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
      Left            =   19680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   11400
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   13455
      Left            =   0
      TabIndex        =   0
      Top             =   -240
      Width           =   23175
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
         TabIndex        =   13
         Top             =   8520
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   9360
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
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   10200
         Width           =   2655
      End
      Begin VB.CommandButton btnUpdate 
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7680
         Width           =   2655
      End
      Begin VB.CommandButton btnLogOut 
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
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   11040
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
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6840
         UseMaskColor    =   -1  'True
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
         TabIndex        =   6
         Top             =   6120
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ABOUT US"
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
         Left            =   8910
         TabIndex        =   11
         Top             =   1680
         Width           =   2475
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
         TabIndex        =   5
         Top             =   5520
         Width           =   3375
      End
      Begin VB.Image ImageMe 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3855
         Left            =   720
         Picture         =   "Frm About.frx":0000
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   3330
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
         Left            =   120
         TabIndex        =   4
         Top             =   12360
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
         Left            =   3180
         TabIndex        =   3
         Top             =   0
         Width           =   16695
      End
      Begin VB.Image BtnHOME3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   120
         Picture         =   "Frm About.frx":3DD5BD
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm About.frx":3F6373
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   6495
         Left            =   5160
         TabIndex        =   1
         Top             =   3120
         Width           =   10215
      End
      Begin VB.Image Image3 
         Height          =   12990
         Left            =   0
         Picture         =   "Frm About.frx":3F66D0
         Stretch         =   -1  'True
         Top             =   240
         Width           =   23040
      End
   End
   Begin VB.CommandButton BtnHOME1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   12480
      TabIndex        =   2
      Top             =   8520
      Width           =   1095
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnADD_Click()
FrmAbout.Hide
FrmAdminHome.Show
End Sub

Private Sub BtnDelete_Click()
FrmAbout.Hide
FrmDelete.Show
End Sub

Private Sub btnLogOut_Click()
FrmAbout.Hide
FrmAdminLogin.Show
End Sub

Private Sub BtnPrint_Click()
FrmAbout.Hide
FrmPrint.Show
End Sub

Private Sub BtnShowAll_Click()
FrmAbout.Hide
FrmShowAll.Show
End Sub

Private Sub BtnUpdate_Click()
FrmAbout.Hide
FrmUpdate.Show
End Sub

Private Sub ContactUs_Click()
FrmAbout.Hide
FrmContactUs.Show
End Sub

Private Sub btnSearch_Click()
FrmAbout.Hide
FrmSearch.Show
End Sub

Private Sub BtnHOME3_Click()
FrmAbout.Hide
Frmadminmodule.Show
End Sub
