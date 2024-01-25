VERSION 5.00
Begin VB.Form FrmSearch 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Search"
   ClientHeight    =   12120
   ClientLeft      =   0
   ClientTop       =   -165
   ClientWidth     =   22395
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   12120
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
      TabIndex        =   30
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8280
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
      TabIndex        =   28
      Top             =   7440
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
      TabIndex        =   15
      Top             =   9960
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
      TabIndex        =   13
      Top             =   10800
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
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6600
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
      TabIndex        =   11
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Frame1"
      Height          =   12735
      Left            =   0
      TabIndex        =   0
      Top             =   -240
      Width           =   22935
      Begin VB.ComboBox TextSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5640
         TabIndex        =   1
         Text            =   "Enter Name to Search"
         Top             =   2520
         Width           =   7215
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
         TabIndex        =   14
         Top             =   9360
         Width           =   2655
      End
      Begin VB.CommandButton BtnSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8280
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "GENDER:"
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
         Height          =   735
         Left            =   4920
         TabIndex        =   27
         Top             =   10920
         Width           =   3855
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WEIGHT:"
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
         Height          =   735
         Left            =   4920
         TabIndex        =   26
         Top             =   9960
         Width           =   3855
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BLOOD GROUP:"
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
         Height          =   735
         Left            =   4920
         TabIndex        =   25
         Top             =   9120
         Width           =   3855
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ADDERESS:"
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
         Height          =   735
         Left            =   4920
         TabIndex        =   24
         Top             =   7920
         Width           =   3855
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ADHAR NUMBER:"
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
         Height          =   735
         Left            =   4920
         TabIndex        =   23
         Top             =   6960
         Width           =   3855
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DONER'S AGE:"
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
         Height          =   735
         Left            =   4920
         TabIndex        =   22
         Top             =   6000
         Width           =   3855
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PHONE NUMBER:"
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
         Height          =   735
         Left            =   4920
         TabIndex        =   21
         Top             =   5160
         Width           =   3855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DONER'S NAME:"
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
         Height          =   465
         Left            =   4920
         TabIndex        =   20
         Top             =   4320
         Width           =   3045
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH"
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
         Left            =   7455
         TabIndex        =   19
         Top             =   1680
         Width           =   2025
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
         TabIndex        =   18
         Top             =   5520
         Width           =   3375
      End
      Begin VB.Image ImageMe 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3855
         Left            =   720
         Picture         =   "Frm Search.frx":0000
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   3330
      End
      Begin VB.Label LblDonerGender 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   9000
         TabIndex        =   10
         Top             =   10800
         Width           =   5535
      End
      Begin VB.Label LblDonerWeight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   9000
         TabIndex        =   9
         Top             =   9840
         Width           =   5535
      End
      Begin VB.Label LblDonerBloodGroup 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   9000
         TabIndex        =   8
         Top             =   9000
         Width           =   5535
      End
      Begin VB.Label LblDonerAdderess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   9000
         TabIndex        =   7
         Top             =   7680
         Width           =   5535
      End
      Begin VB.Label LblAdharNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   9000
         TabIndex        =   6
         Top             =   6840
         Width           =   5535
      End
      Begin VB.Label LblDonerAge 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   9000
         TabIndex        =   5
         Top             =   6000
         Width           =   5535
      End
      Begin VB.Label LblPhoneNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   9000
         TabIndex        =   4
         Top             =   5160
         Width           =   5535
      End
      Begin VB.Label LblDonerName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   9000
         TabIndex        =   3
         Top             =   4320
         Width           =   5535
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
         TabIndex        =   17
         Top             =   12120
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
         Left            =   2400
         TabIndex        =   16
         Top             =   240
         Width           =   16695
      End
      Begin VB.Image BtnHOME 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   120
         Picture         =   "Frm Search.frx":3DD5BD
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   12750
         Left            =   -600
         Picture         =   "Frm Search.frx":3F6373
         Stretch         =   -1  'True
         Top             =   0
         Width           =   23520
      End
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim ors As New ADODB.Recordset
Dim sqlQuery, vID, vDonerName, vPhoneNumber, vDonerAge, vAdharNumber, vDonerAddress, vDonerBloodGroup, vDonerWeight, vDonerGender As String

Private Sub btnABTUs_Click()
FrmSearch.Hide
FrmAbout.Show

 LblDonerName.Caption = " "
    LblPhoneNumber.Caption = " "
    LblDonerAge.Caption = " "
    LblAdharNumber.Caption = " "
    LblDonerAdderess.Caption = " "
    LblDonerBloodGroup.Caption = " "
    LblDonerWeight.Caption = " "
    LblDonerGender.Caption = " "

End Sub

Private Sub btnADD_Click()
FrmSearch.Hide
FrmAdminHome.Show

 LblDonerName.Caption = " "
    LblPhoneNumber.Caption = " "
    LblDonerAge.Caption = " "
    LblAdharNumber.Caption = " "
    LblDonerAdderess.Caption = " "
    LblDonerBloodGroup.Caption = " "
    LblDonerWeight.Caption = " "
    LblDonerGender.Caption = " "

End Sub

Private Sub BtnDelete_Click()
FrmSearch.Hide
FrmDelete.Show
 LblDonerName.Caption = " "
    LblPhoneNumber.Caption = " "
    LblDonerAge.Caption = " "
    LblAdharNumber.Caption = " "
    LblDonerAdderess.Caption = " "
    LblDonerBloodGroup.Caption = " "
    LblDonerWeight.Caption = " "
    LblDonerGender.Caption = " "

End Sub

Private Sub BtnPrint_Click()
FrmSearch.Hide
FrmPrint.Show
 LblDonerName.Caption = " "
    LblPhoneNumber.Caption = " "
    LblDonerAge.Caption = " "
    LblAdharNumber.Caption = " "
    LblDonerAdderess.Caption = " "
    LblDonerBloodGroup.Caption = " "
    LblDonerWeight.Caption = " "
    LblDonerGender.Caption = " "

End Sub

Private Sub btnREPORT_Click()
FrmSearch.Hide
FrmAdminLogin.Show

 LblDonerName.Caption = " "
    LblPhoneNumber.Caption = " "
    LblDonerAge.Caption = " "
    LblAdharNumber.Caption = " "
    LblDonerAdderess.Caption = " "
    LblDonerBloodGroup.Caption = " "
    LblDonerWeight.Caption = " "
    LblDonerGender.Caption = " "

End Sub

Private Sub btnSearch_Click()
Dim vDonerName As String

If Len(TextSearch.Text) = 0 Then

    LblDonerName.Caption = " "
    LblPhoneNumber.Caption = " "
    LblDonerAge.Caption = " "
    LblAdharNumber.Caption = " "
    LblDonerAdderess.Caption = " "
    LblDonerBloodGroup.Caption = " "
    LblDonerWeight.Caption = " "
    LblDonerGender.Caption = " "

MsgBox "Search Box should not be empty!", , "Blood Bank Information System"
TextSearch.SetFocus
Exit Sub
End If


vDonerName = TextSearch
sqlQuery = "SELECT * FROM tblname WHERE DonerName= '" & vDonerName & "';"

conn.Open

ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly

If ors.EOF = True Then
   
    LblDonerName.Caption = " "
    LblPhoneNumber.Caption = " "
    LblDonerAge.Caption = " "
    LblAdharNumber.Caption = " "
    LblDonerAdderess.Caption = " "
    LblDonerBloodGroup.Caption = " "
    LblDonerWeight.Caption = " "
    LblDonerGender.Caption = " "
   MsgBox "Record Not Found!", , "Blood Bank Information System"

Else
    
    vDonerName = ors.Fields("DonerName").Value
    vPhoneNumber = ors.Fields("PhoneNumber").Value
    vDonerAge = ors.Fields("DonerAge").Value
    vAdharNumber = ors.Fields("AdharNumber").Value
    vDonerAddress = ors.Fields("DonerAddress").Value
    vDonerBloodGroup = ors.Fields("DonerBloodGroup").Value
    vDonerWeight = ors.Fields("DonerWeight").Value
    vDonerGender = ors.Fields("DonerGender").Value
   
    LblDonerName.Caption = vDonerName
    LblPhoneNumber.Caption = vPhoneNumber
    LblDonerAge.Caption = vDonerAge
    LblAdharNumber.Caption = vAdharNumber
    LblDonerAdderess.Caption = vDonerAddress
    LblDonerBloodGroup.Caption = vDonerBloodGroup
    LblDonerWeight.Caption = vDonerWeight
    LblDonerGender.Caption = vDonerGender
    
End If
conn.Close
End Sub

Private Sub BtnShowAll_Click()
FrmSearch.Hide
FrmShowAll.Show

    LblDonerName.Caption = " "
    LblPhoneNumber.Caption = " "
    LblDonerAge.Caption = " "
    LblAdharNumber.Caption = " "
    LblDonerAdderess.Caption = " "
    LblDonerBloodGroup.Caption = " "
    LblDonerWeight.Caption = " "
    LblDonerGender.Caption = " "

End Sub

Private Sub BtnUpdate_Click()
FrmSearch.Hide
FrmUpdate.Show

    LblDonerName.Caption = " "
    LblPhoneNumber.Caption = " "
    LblDonerAge.Caption = " "
    LblAdharNumber.Caption = " "
    LblDonerAdderess.Caption = " "
    LblDonerBloodGroup.Caption = " "
    LblDonerWeight.Caption = " "
    LblDonerGender.Caption = " "

End Sub

Private Sub ContactUs_Click()
FrmSearch.Hide
FrmContactUs.Show

    LblDonerName.Caption = " "
    LblPhoneNumber.Caption = " "
    LblDonerAge.Caption = " "
    LblAdharNumber.Caption = " "
    LblDonerAdderess.Caption = " "
    LblDonerBloodGroup.Caption = " "
    LblDonerWeight.Caption = " "
    LblDonerGender.Caption = " "

End Sub

Private Sub Form_Load()
conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "\database\rd.mdb;Persist Security Info=False"

sqlQuery = "Select * FROM tblname;"
conn.Open
    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
    Do Until ors.EOF
    vDonerName = ors.Fields("DonerName").Value
    
    TextSearch.AddItem vDonerName
    ors.MoveNext
    Loop
    
    conn.Close
        
End Sub

Private Sub BtnHOME_Click()
FrmSearch.Hide
Frmadminmodule.Show vbModal

    LblDonerName.Caption = " "
    LblPhoneNumber.Caption = " "
    LblDonerAge.Caption = " "
    LblAdharNumber.Caption = " "
    LblDonerAdderess.Caption = " "
    LblDonerBloodGroup.Caption = " "
    LblDonerWeight.Caption = " "
    LblDonerGender.Caption = " "

End Sub

