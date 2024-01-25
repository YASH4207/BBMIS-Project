VERSION 5.00
Begin VB.Form FrmContactUs 
   BorderStyle     =   0  'None
   Caption         =   "Contact Us"
   ClientHeight    =   12960
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   22395
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
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
      TabIndex        =   16
      Top             =   11760
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contact Us"
      Height          =   13095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   22935
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
         TabIndex        =   15
         Top             =   9480
         Width           =   2655
      End
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
         TabIndex        =   14
         Top             =   8640
         Width           =   2655
      End
      Begin VB.CommandButton Update 
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
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7800
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   10320
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
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   11160
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
         TabIndex        =   9
         Top             =   6960
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
         TabIndex        =   8
         Top             =   6120
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT US"
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
         Left            =   8940
         TabIndex        =   13
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label6 
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
         TabIndex        =   7
         Top             =   5520
         Width           =   3375
      End
      Begin VB.Image ImageMe 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3855
         Left            =   720
         Picture         =   "FormContact.frx":0000
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   3330
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ydarne8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   7755
         TabIndex        =   6
         Top             =   5280
         Width           =   1185
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   6720
         Picture         =   "FormContact.frx":3DD5BD
         Stretch         =   -1  'True
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "https://goo.gl/maps/9EZBBe3UfKPRM52UA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   7725
         TabIndex        =   5
         Top             =   4560
         Width           =   4485
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   6720
         Picture         =   "FormContact.frx":3E229A
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ydarne8@gmail.com"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   7680
         TabIndex        =   4
         Top             =   3840
         Width           =   3015
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   6720
         Picture         =   "FormContact.frx":3E85E8
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "7620214815"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   7680
         TabIndex        =   3
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   6720
         Picture         =   "FormContact.frx":3ED28D
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   735
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
         Left            =   3120
         TabIndex        =   2
         Top             =   0
         Width           =   16695
      End
      Begin VB.Image BtnHOME4 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   120
         Picture         =   "FormContact.frx":3F2575
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1335
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
         TabIndex        =   1
         Top             =   12360
         Width           =   8835
      End
      Begin VB.Image Image1 
         Height          =   12990
         Left            =   0
         Picture         =   "FormContact.frx":40B32B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   23055
      End
   End
End
Attribute VB_Name = "FrmContactUs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnABTUs_Click()
FrmContactUs.Hide
FrmAbout.Show
End Sub

Private Sub btnADD_Click()
FrmContactUs.Hide
FrmAdminHome.Show
End Sub

Private Sub BtnDelete_Click()
FrmContactUs.Hide
FrmDelete.Show
End Sub

Private Sub BtnHOME4_Click()
FrmContactUs.Hide
Frmadminmodule.Show
End Sub

Private Sub BtnPrint_Click()
FrmContactUs.Hide
FrmPrint.Show
End Sub

Private Sub btnREPORT_Click()
FrmContactUs.Hide
FrmAdminLogin.Show
End Sub

Private Sub btnSearch_Click()
FrmContactUs.Hide
FrmSearch.Show
End Sub

Private Sub BtnShowAll_Click()
FrmContactUs.Hide
FrmShowAll.Show

End Sub

Private Sub Update_Click()
FrmContactUs.Hide
FrmUpdate.Show
End Sub
