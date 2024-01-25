VERSION 5.00
Begin VB.Form FrmShowAll 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12120
   ClientLeft      =   -345
   ClientTop       =   0
   ClientWidth     =   22395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   12120
   ScaleWidth      =   22395
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   20520
      ScaleHeight     =   9705
      ScaleWidth      =   1305
      TabIndex        =   21
      Top             =   1560
      Width           =   1335
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   19080
      ScaleHeight     =   9705
      ScaleWidth      =   1305
      TabIndex        =   20
      Top             =   1560
      Width           =   1335
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   17160
      ScaleHeight     =   9705
      ScaleWidth      =   1785
      TabIndex        =   19
      Top             =   1560
      Width           =   1815
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   13440
      ScaleHeight     =   9705
      ScaleWidth      =   3585
      TabIndex        =   18
      Top             =   1560
      Width           =   3615
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   11760
      ScaleHeight     =   9705
      ScaleWidth      =   1545
      TabIndex        =   17
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   10560
      ScaleHeight     =   9705
      ScaleWidth      =   1065
      TabIndex        =   16
      Top             =   1560
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   8880
      ScaleHeight     =   9705
      ScaleWidth      =   1545
      TabIndex        =   15
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   5280
      ScaleHeight     =   9705
      ScaleWidth      =   3465
      TabIndex        =   14
      Top             =   1560
      Width           =   3495
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
      Top             =   5400
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
      TabIndex        =   10
      Top             =   6240
      UseMaskColor    =   -1  'True
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
      TabIndex        =   9
      Top             =   7080
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
      TabIndex        =   8
      Top             =   7920
      UseMaskColor    =   -1  'True
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
      TabIndex        =   7
      Top             =   8760
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
      TabIndex        =   6
      Top             =   9600
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
      TabIndex        =   5
      Top             =   10440
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
      TabIndex        =   4
      Top             =   11280
      Width           =   2655
   End
   Begin VB.PictureBox PictureId 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   4800
      ScaleHeight     =   9705
      ScaleWidth      =   345
      TabIndex        =   3
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton BtnShowAll 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "SHOW ALL"
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
      Left            =   19440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   11400
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "SHOW ALL INFO"
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
      Left            =   9255
      TabIndex        =   13
      Top             =   960
      Width           =   3915
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
      TabIndex        =   12
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Image ImageMe 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3495
      Left            =   960
      Picture         =   "FrmShowAll.frx":0000
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2970
   End
   Begin VB.Label Label2 
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
      TabIndex        =   2
      Top             =   12360
      Width           =   8835
   End
   Begin VB.Image BtnHOME 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   120
      Picture         =   "FrmShowAll.frx":3DD5BD
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1335
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
      Left            =   2940
      TabIndex        =   0
      Top             =   -120
      Width           =   16695
   End
   Begin VB.Image Image1 
      Height          =   12855
      Left            =   0
      Picture         =   "FrmShowAll.frx":3F6373
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22935
   End
End
Attribute VB_Name = "FrmShowAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim ors As New ADODB.Recordset

Dim sqlQuery, vID, vDonerName, vPhoneNumber, vDonerAge, vAdharNumber, vDonerAddress, vDonerBloodGroup, vDonerWeight, vDonerGender As String
Dim vimagedata() As Byte


Private Sub btnABTUs_Click()
FrmShowAll.Hide
FrmAbout.Show
End Sub

Private Sub btnADD_Click()
FrmShowAll.Hide
FrmAdminHome.Show
End Sub

Private Sub BtnDelete_Click()
FrmShowAll.Hide
FrmDelete.Show
End Sub

Private Sub BtnHOME_Click()
FrmShowAll.Hide
Frmadminmodule.Show
End Sub

Private Sub btnLogOut_Click()
FrmShowAll.Hide
FrmAdminLogin.Show
End Sub

Private Sub BtnPrint_Click()
FrmShowAll.Hide
FrmPrint.Show
End Sub

Private Sub btnSearch_Click()
FrmShowAll.Hide
FrmSearch.Show
End Sub

Private Sub BtnShowAll_Click()
sqlQuery = "SELECT * FROM tblname;"

conn.Open
ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly

Dim vDataShow, vDataShow1, vDataShow2, vDataShow3, vDataShow4, vDataShow5, vDataShow6, vDataShow7, vDataShow8 As String
vDataShow = "Id" & vbCrLf
vDataShow1 = "Doner's Name" & vbCrLf
vDataShow2 = "PhoneNumber" & vbCrLf
vDataShow3 = "DonerAge" & vbCrLf
vDataShow4 = "AdharNumber" & vbCrLf
vDataShow5 = "DonerAddress" & vbCrLf
vDataShow6 = "DonerBloodGroup" & vbCrLf
vDataShow7 = "DonerWeight" & vbCrLf
vDataShow8 = "DonerGender" & vbCrLf

Do Until ors.EOF
    vID = ors.Fields("id").Value
    vDonerName = ors.Fields("DonerName").Value
    vPhoneNumber = ors.Fields("PhoneNumber").Value
    vDonerAge = ors.Fields("DonerAge").Value
    vAdharNumber = ors.Fields("AdharNumber").Value
    vDonerAddress = ors.Fields("DonerAddress").Value
    vDonerBloodGroup = ors.Fields("DonerBloodGroup").Value
    vDonerWeight = ors.Fields("DonerWeight").Value
    vDonerGender = ors.Fields("DonerGender").Value
    
       
    vDataShow = vDataShow & vID & vbCrLf
    vDataShow1 = vDataShow1 & vDonerName & vbCrLf
    vDataShow2 = vDataShow2 & vPhoneNumber & vbCrLf
    vDataShow3 = vDataShow3 & vDonerAge & vbCrLf
    vDataShow4 = vDataShow4 & vAdharNumber & vbCrLf
    vDataShow5 = vDataShow5 & vDonerAddress & vbCrLf
    vDataShow6 = vDataShow6 & vDonerBloodGroup & vbCrLf
    vDataShow7 = vDataShow7 & vDonerWeight & vbCrLf
    vDataShow8 = vDataShow8 & vDonerGender & vbCrLf
    ors.MoveNext
Loop

PictureId.Print vDataShow
Picture1.Print vDataShow1
Picture2.Print vDataShow2
Picture3.Print vDataShow3
Picture4.Print vDataShow4
Picture5.Print vDataShow5
Picture6.Print vDataShow6
Picture7.Print vDataShow7
Picture8.Print vDataShow8

conn.Close
End Sub

Private Sub BtnUpdate_Click()
FrmShowAll.Hide
FrmUpdate.Show
End Sub

Private Sub ContactUs_Click()
FrmShowAll.Hide
FrmContactUs.Show
End Sub

Private Sub Form_Load()
conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "\database\rd.mdb;Persist Security Info=False"
End Sub
