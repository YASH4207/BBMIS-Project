VERSION 5.00
Begin VB.Form FrmUpdate 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   12960
   ScaleWidth      =   22920
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
      TabIndex        =   28
      Top             =   8640
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
      TabIndex        =   27
      Top             =   7800
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
      TabIndex        =   26
      Top             =   6960
      UseMaskColor    =   -1  'True
      Width           =   2655
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
      Top             =   3240
      Width           =   7215
   End
   Begin VB.CommandButton CmdSearch 
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
      Left            =   13320
      MaskColor       =   &H00404040&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
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
      TabIndex        =   13
      Top             =   9480
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
      TabIndex        =   14
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
      TabIndex        =   15
      Top             =   11160
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
      TabIndex        =   12
      Top             =   6120
      Width           =   2655
   End
   Begin VB.CommandButton cmdUpdate 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   9240
      MaskColor       =   &H00000000&
      TabIndex        =   11
      Top             =   10920
      Width           =   2175
   End
   Begin VB.TextBox TxtDonerGender 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   10
      Top             =   9840
      Width           =   5895
   End
   Begin VB.TextBox TxtDonerWeight 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      MaxLength       =   3
      TabIndex        =   9
      Top             =   9000
      Width           =   5895
   End
   Begin VB.TextBox TxtDonerBloodGroup 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   8
      Top             =   8160
      Width           =   5895
   End
   Begin VB.TextBox TxtDonerAddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   7
      Top             =   7320
      Width           =   5895
   End
   Begin VB.TextBox TxtDonerAdharNumber 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      MaxLength       =   12
      TabIndex        =   6
      Top             =   6480
      Width           =   5895
   End
   Begin VB.TextBox TxtDonerAge 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      MaxLength       =   2
      TabIndex        =   5
      Top             =   5640
      Width           =   5895
   End
   Begin VB.TextBox TxtDonerPhoneNumber 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      MaxLength       =   10
      TabIndex        =   4
      Top             =   4800
      Width           =   5895
   End
   Begin VB.TextBox TxtDonerName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   3
      Top             =   3960
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE DATA"
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
      Left            =   8040
      TabIndex        =   29
      Top             =   1440
      Width           =   3465
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
      TabIndex        =   25
      Top             =   12360
      Width           =   8835
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
      Left            =   5520
      TabIndex        =   24
      Top             =   9840
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
      Height          =   465
      Left            =   5520
      TabIndex        =   23
      Top             =   9000
      Width           =   4290
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
      Left            =   5520
      TabIndex        =   22
      Top             =   8160
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
      Left            =   5520
      TabIndex        =   21
      Top             =   7320
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
      Left            =   5520
      TabIndex        =   20
      Top             =   6480
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
      Left            =   5520
      TabIndex        =   19
      Top             =   5640
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
      Left            =   5520
      TabIndex        =   18
      Top             =   4800
      Width           =   3855
   End
   Begin VB.Image BtnHOME 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   120
      Picture         =   "FrmUpdate.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
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
      Left            =   5520
      TabIndex        =   17
      Top             =   3960
      Width           =   3045
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
      Top             =   0
      Width           =   16695
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
      TabIndex        =   0
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Image ImageMe 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3855
      Left            =   720
      Picture         =   "FrmUpdate.frx":18DB6
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3330
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   13110
      Left            =   -120
      Picture         =   "FrmUpdate.frx":3F6373
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   23040
   End
End
Attribute VB_Name = "FrmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim ors As New ADODB.Recordset
Dim sqlQuery, vID, vDonerName, vPhoneNumber, vDonerAge, vAdharNumber, vDonerAddress, vDonerBloodGroup, vDonerWeight, vDonerGender As String

Private Sub BtnDelete_Click()
FrmUpdate.Hide
FrmDelete.Show

TxtDonerName.Text = " "
    TxtDonerPhoneNumber.Text = " "
    TxtDonerAge.Text = " "
    TxtDonerAdharNumber.Text = " "
    TxtDonerAddress.Text = " "
    TxtDonerBloodGroup.Text = " "
    TxtDonerWeight.Text = " "
    TxtDonerGender.Text = " "
End Sub

Private Sub BtnPrint_Click()
FrmUpdate.Hide
FrmPrint.Show

TxtDonerName.Text = " "
    TxtDonerPhoneNumber.Text = " "
    TxtDonerAge.Text = " "
    TxtDonerAdharNumber.Text = " "
    TxtDonerAddress.Text = " "
    TxtDonerBloodGroup.Text = " "
    TxtDonerWeight.Text = " "
    TxtDonerGender.Text = " "
End Sub

Private Sub BtnShowAll_Click()
FrmUpdate.Hide
FrmShowAll.Show

TxtDonerName.Text = " "
    TxtDonerPhoneNumber.Text = " "
    TxtDonerAge.Text = " "
    TxtDonerAdharNumber.Text = " "
    TxtDonerAddress.Text = " "
    TxtDonerBloodGroup.Text = " "
    TxtDonerWeight.Text = " "
    TxtDonerGender.Text = " "
End Sub

Private Sub CmdSearch_Click()

Dim vDonerName As String

If Len(TextSearch.Text) = 0 Then

    TxtDonerName.Text = " "
    TxtDonerPhoneNumber.Text = " "
    TxtDonerAge.Text = " "
    TxtDonerAdharNumber.Text = " "
    TxtDonerAddress.Text = " "
    TxtDonerBloodGroup.Text = " "
    TxtDonerWeight.Text = " "
    TxtDonerGender.Text = " "
    
MsgBox "Search Box should not be empty!", , "Blood Bank Information System"
TextSearch.SetFocus
Exit Sub
End If


vDonerName = TextSearch
sqlQuery = "SELECT * FROM tblname WHERE DonerName= '" & vDonerName & "';"

conn.Open

ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly

If ors.EOF = True Then
   
    TxtDonerName.Text = " "
    TxtDonerPhoneNumber.Text = " "
    TxtDonerAge.Text = " "
    TxtDonerAdharNumber.Text = " "
    TxtDonerAddress.Text = " "
    TxtDonerBloodGroup.Text = " "
    TxtDonerWeight.Text = " "
    TxtDonerGender.Text = " "
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
   
    TxtDonerName.Text = vDonerName
    TxtDonerPhoneNumber.Text = vPhoneNumber
    TxtDonerAge.Text = vDonerAge
    TxtDonerAdharNumber.Text = vAdharNumber
    TxtDonerAddress.Text = vDonerAddress
    TxtDonerBloodGroup.Text = vDonerBloodGroup
    TxtDonerWeight.Text = vDonerWeight
    TxtDonerGender.Text = vDonerGender
    
End If
conn.Close
End Sub

Private Sub cmdUpdate_Click()

Dim vDonerName As String
'Dim cmbID As ComboBox
vDonerName = TextSearch.Text
sqlQuery = "SELECT * FROM tblname WHERE DonerName='" & vDonerName & "';"
conn.Open
    ors.Open sqlQuery, conn, adOpenDynamic, adLockOptimistic
    ors.Update "DonerName", TxtDonerName.Text
    ors.Update "PhoneNumber", TxtDonerPhoneNumber.Text
    ors.Update "DonerAge", TxtDonerAge.Text
    ors.Update "AdharNumber", TxtDonerAdharNumber.Text
    ors.Update "DonerAddress", TxtDonerAddress.Text
    ors.Update "DonerBloodGroup", TxtDonerBloodGroup.Text
    ors.Update "DonerWeight", TxtDonerWeight.Text
    ors.Update "DonerGender", TxtDonerGender.Text
    With ors

    End With
    
    MsgBox "Record Updated Successfully", , "Blood Bank Information System"
    
    TxtDonerName.Text = " "
    TxtDonerPhoneNumber.Text = " "
    TxtDonerAge.Text = " "
    TxtDonerAdharNumber.Text = " "
    TxtDonerAddress.Text = " "
    TxtDonerBloodGroup.Text = " "
    TxtDonerWeight.Text = " "
    TxtDonerGender.Text = " "
    
conn.Close

End Sub

Private Sub btnABTUs_Click()
FrmUpdate.Hide
FrmAbout.Show

TxtDonerName.Text = " "
    TxtDonerPhoneNumber.Text = " "
    TxtDonerAge.Text = " "
    TxtDonerAdharNumber.Text = " "
    TxtDonerAddress.Text = " "
    TxtDonerBloodGroup.Text = " "
    TxtDonerWeight.Text = " "
    TxtDonerGender.Text = " "
End Sub

Private Sub btnADD_Click()
FrmUpdate.Hide
FrmAdminHome.Show

TxtDonerName.Text = " "
    TxtDonerPhoneNumber.Text = " "
    TxtDonerAge.Text = " "
    TxtDonerAdharNumber.Text = " "
    TxtDonerAddress.Text = " "
    TxtDonerBloodGroup.Text = " "
    TxtDonerWeight.Text = " "
    TxtDonerGender.Text = " "
End Sub

Private Sub BtnHOME_Click()
FrmUpdate.Hide
Frmadminmodule.Show

TxtDonerName.Text = " "
    TxtDonerPhoneNumber.Text = " "
    TxtDonerAge.Text = " "
    TxtDonerAdharNumber.Text = " "
    TxtDonerAddress.Text = " "
    TxtDonerBloodGroup.Text = " "
    TxtDonerWeight.Text = " "
    TxtDonerGender.Text = " "
End Sub

Private Sub btnREPORT_Click()
FrmUpdate.Hide
FrmAdminLogin.Show

TxtDonerName.Text = " "
    TxtDonerPhoneNumber.Text = " "
    TxtDonerAge.Text = " "
    TxtDonerAdharNumber.Text = " "
    TxtDonerAddress.Text = " "
    TxtDonerBloodGroup.Text = " "
    TxtDonerWeight.Text = " "
    TxtDonerGender.Text = " "
End Sub

Private Sub btnSearch_Click()
FrmUpdate.Hide
FrmSearch.Show

TxtDonerName.Text = " "
    TxtDonerPhoneNumber.Text = " "
    TxtDonerAge.Text = " "
    TxtDonerAdharNumber.Text = " "
    TxtDonerAddress.Text = " "
    TxtDonerBloodGroup.Text = " "
    TxtDonerWeight.Text = " "
    TxtDonerGender.Text = " "
End Sub

Private Sub ContactUs_Click()
FrmUpdate.Hide
FrmContactUs.Show

TxtDonerName.Text = " "
    TxtDonerPhoneNumber.Text = " "
    TxtDonerAge.Text = " "
    TxtDonerAdharNumber.Text = " "
    TxtDonerAddress.Text = " "
    TxtDonerBloodGroup.Text = " "
    TxtDonerWeight.Text = " "
    TxtDonerGender.Text = " "
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


Private Sub TxtDonerAdharNumber_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
End If
End Sub

Private Sub TxtDonerAge_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
End If
End Sub

Private Sub txtDonerName_KeyUp(KeyCode As Integer, Shift As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8 Or KeyAscii = 32) Then
KeyAscii = 0
End If
End Sub


Private Sub TxtDonerPhoneNumber_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
End If
End Sub

Private Sub TxtDonerWeight_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
End If
End Sub
