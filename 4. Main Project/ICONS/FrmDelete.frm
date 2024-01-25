VERSION 5.00
Begin VB.Form FrmDelete 
   BorderStyle     =   0  'None
   Caption         =   "DELETE"
   ClientHeight    =   12960
   ClientLeft      =   210
   ClientTop       =   -405
   ClientWidth     =   22395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
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
      Left            =   20040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   11760
      Width           =   2655
   End
   Begin VB.ComboBox ComboDeleteRecord 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6360
      TabIndex        =   1
      Text            =   "Enter Adhar Number"
      Top             =   3120
      Width           =   5775
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
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
      TabIndex        =   6
      Top             =   8520
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   9360
      Width           =   2655
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
      TabIndex        =   5
      Top             =   7680
      UseMaskColor    =   -1  'True
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
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   6000
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
      TabIndex        =   9
      Top             =   11040
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "DELETE INFO"
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
      Left            =   8625
      TabIndex        =   12
      Top             =   1680
      Width           =   3255
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
      TabIndex        =   11
      Top             =   12360
      Width           =   8835
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
      TabIndex        =   10
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Image ImageMe 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3855
      Left            =   720
      Picture         =   "FrmDelete.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3330
   End
   Begin VB.Image BtnHOME 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   120
      Picture         =   "FrmDelete.frx":3DD5BD
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
      Left            =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   16695
   End
   Begin VB.Image Image1 
      Height          =   12975
      Left            =   0
      Picture         =   "FrmDelete.frx":3F6373
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22935
   End
End
Attribute VB_Name = "FrmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim ors As New ADODB.Recordset
Dim vAdharNumber As Long
Dim sqlQuery As String
Private Sub btnABTUs_Click()
FrmDelete.Hide
FrmAbout.Show
End Sub
Private Sub btnADD_Click()
FrmDelete.Hide
FrmAdminHome.Show
End Sub
Private Sub BtnDelete_Click()
'sqlQuery = "DELETE FROM tblname WHERE adharnumber=" & Combodelete.Text & ";"
sqlQuery = "DELETE FROM tblname WHERE AdharNumber='" & Replace(ComboDeleteRecord.Text, "'", "''") & "';"

conn.Open
conn.Execute sqlQuery
conn.Close
MsgBox "Data Deleted", , "Blood Bank Information System"
End Sub
Private Sub BtnHOME_Click()
FrmDelete.Hide
Frmadminmodule.Show
End Sub

Private Sub BtnPrint_Click()

FrmDelete.Hide
FrmPrint.Show
End Sub

Private Sub btnREPORT_Click()
FrmDelete.Hide
FrmAdminLogin.Show
End Sub

Private Sub BtnShowAll_Click()
FrmDelete.Hide
FrmShowAll.Show
End Sub

Private Sub BtnUpdate_Click()
FrmDelete.Hide
FrmUpdate.Show
End Sub
Private Sub CommandSearch_Click()
FrmDelete.Hide
FrmSearch.Show
End Sub
Private Sub ContactUs_Click()
FrmDelete.Hide
FrmContactUs.Show
End Sub
Private Sub Form_Load()
conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "\database\rd.mdb;Persist Security Info=False"

'sqlQuery = "Select * FROM tblname;"
'conn.Open
'    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
'    Do Until ors.EOF
'    vAdharNumber = ors.Fields("AdharNumber").Value'
'
 '   ComboDeleteRecord.AddItem vAdharNumber
 '   ors.MoveNext
 '  Loop

' conn.Close
End Sub
