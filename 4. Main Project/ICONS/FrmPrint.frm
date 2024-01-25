VERSION 5.00
Begin VB.Form FrmPrint 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   12960
   ClientLeft      =   210
   ClientTop       =   0
   ClientWidth     =   22395
   ControlBox      =   0   'False
   LinkTopic       =   "Form 2"
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
      TabIndex        =   30
      Top             =   11760
      Width           =   2655
   End
   Begin VB.CommandButton PrintReport 
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10440
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
      Left            =   13560
      MaskColor       =   &H00404040&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
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
      Top             =   2400
      Width           =   7215
   End
   Begin VB.CommandButton CommandSearch 
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
      TabIndex        =   20
      Top             =   8520
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
      TabIndex        =   19
      Top             =   7680
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
      TabIndex        =   18
      Top             =   10200
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
      TabIndex        =   17
      Top             =   9360
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
      TabIndex        =   16
      Top             =   11040
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
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "PRINT REPORT"
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
      Left            =   7965
      TabIndex        =   29
      Top             =   1440
      Width           =   3615
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
      Left            =   9600
      TabIndex        =   11
      Top             =   8520
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
      Left            =   9600
      TabIndex        =   9
      Top             =   7800
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
      Left            =   9600
      TabIndex        =   8
      Top             =   7080
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
      Left            =   9600
      TabIndex        =   7
      Top             =   6000
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
      Left            =   9600
      TabIndex        =   6
      Top             =   5280
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
      Left            =   9600
      TabIndex        =   5
      Top             =   4560
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
      Left            =   9600
      TabIndex        =   4
      Top             =   3840
      Width           =   5535
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
      Left            =   5640
      TabIndex        =   28
      Top             =   8520
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
      Left            =   5640
      TabIndex        =   27
      Top             =   7800
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
      Left            =   5640
      TabIndex        =   26
      Top             =   7080
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
      Left            =   5640
      TabIndex        =   25
      Top             =   6120
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
      Left            =   5640
      TabIndex        =   24
      Top             =   5400
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
      Left            =   5640
      TabIndex        =   23
      Top             =   4680
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
      Left            =   5640
      TabIndex        =   22
      Top             =   3960
      Width           =   3855
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
      Left            =   9600
      TabIndex        =   3
      Top             =   3120
      Width           =   5535
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
      Left            =   5640
      TabIndex        =   21
      Top             =   3240
      Width           =   3045
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
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Image ImageMe 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3855
      Left            =   720
      Picture         =   "FrmPrint.frx":0000
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
      TabIndex        =   12
      Top             =   12360
      Width           =   8835
   End
   Begin VB.Image BtnHOME 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   120
      Picture         =   "FrmPrint.frx":3DD5BD
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
      Left            =   2820
      TabIndex        =   0
      Top             =   0
      Width           =   16695
   End
   Begin VB.Image Image1 
      Height          =   12855
      Left            =   0
      Picture         =   "FrmPrint.frx":3F6373
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22935
   End
End
Attribute VB_Name = "FrmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim ors As New ADODB.Recordset
Dim sqlQuery, vID, vDonerName, vPhoneNumber, vDonerAge, vAdharNumber, vDonerAddress, vDonerBloodGroup, vDonerWeight, vDonerGender As String
Private Sub btnABTUs_Click()
FrmPrint.Hide
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
FrmPrint.Hide
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
FrmPrint.Hide
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
Private Sub BtnHOME_Click()
FrmPrint.Hide
Frmadminmodule.Show

 LblDonerName.Caption = " "
    LblPhoneNumber.Caption = " "
    LblDonerAge.Caption = " "
    LblAdharNumber.Caption = " "
    LblDonerAdderess.Caption = " "
    LblDonerBloodGroup.Caption = " "
    LblDonerWeight.Caption = " "
    LblDonerGender.Caption = " "

End Sub
Private Sub btnLogOut_Click()
FrmPrint.Hide
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
Private Sub BtnPrint_Click()

Dim vTxtID As String
vTxtID = TextSearch.Text
sqlQuery = "SELECT * FROM tblname WHERE AdharNumber='" & vTxtID & "';"

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

    MsgBox "Search Box should not be empty!", , "Blood Bank Information System"
     
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

Private Sub btnSearch_Click()
Dim vTxtID As String

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


vTxtID = TextSearch
sqlQuery = "SELECT * FROM tblname WHERE DonerName= '" & vTxtID & "';"

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
FrmPrint.Hide
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

Private Sub BtnUpdate_Click()
FrmPrint.Hide
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

Private Sub CommandSearch_Click()
FrmPrint.Hide
FrmSearch.Show

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
FrmPrint.Hide
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

Private Sub PrintReport_Click()
PrintForm
End Sub
