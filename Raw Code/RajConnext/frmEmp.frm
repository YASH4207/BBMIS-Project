VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDataEnv 
   Caption         =   "Employee Form"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtid 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FEEADA&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmd_ClearPicture 
      BackColor       =   &H00FEEADA&
      Caption         =   "Clear Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmd_LoadPicture 
      BackColor       =   &H00FEEADA&
      Caption         =   "Load Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   405
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommDlg_Path 
      Left            =   5040
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "white.JPG"
   End
   Begin VB.Label Label2 
      Caption         =   "UID : "
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Image Img_Emp 
      Height          =   1215
      Left            =   3600
      Picture         =   "frmEmp.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Name"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ocon As New ADODB.Connection
Dim ors As New ADODB.Recordset
Dim checkValid As Integer
Dim ans


Private Sub cmd_ClearPicture_Click()
CommDlg_Path.FileName = App.Path & "\white.JPG"
'CommDlg_Path.FileName = ""
Img_Emp.Picture = LoadPicture()

End Sub

Private Sub cmd_LoadPicture_Click()
With CommDlg_Path
    .DialogTitle = "Search Employee picture"
    .Filter = "JPEG (Jpeg (*.jpg)|*.jpg|*.bmp)|*.bmp|Gif (*.gif)|*.gif|All Files (*.*)|*.*"
    .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
    .ShowOpen
    .FilterIndex = 1
    .CancelError = False
Img_Emp.Picture = LoadPicture(.FileName)
End With

End Sub

Private Sub cmdSave_Click()
checkValid = 0
'Validation
If Len(txtid.Text) = 0 Then
MsgBox "Enter ID"
txtid.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtName.Text) = 0 Then
MsgBox "Enter Name"
txtName.SetFocus
checkValid = checkValid + 1
End If


If checkValid = 0 Then
If MsgBox("Confirm add new Employee Record.", vbQuestion + vbYesNo) = vbYes Then
Dim ssql1 As String
ocon.Open
    ssql1 = ssql1 & "INSERT INTO tblname(id,name,imgrd)"
    ssql1 = ssql1 & " values("
    
    ssql1 = ssql1 & "" & UCase(Trim(txtid.Text)) & ","
    ssql1 = ssql1 & "'" & UCase(Trim(txtName.Text)) & "',"
    ssql1 = ssql1 & "'" & UCase(Trim(CommDlg_Path.FileName)) & "');"

MsgBox ssql1

ocon.Execute ssql1
ocon.Close
MsgBox "Recordset added successfully"
End If
End If
End Sub

Private Sub Form_Load()
ocon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "\db\rd.mdb;Persist Security Info=False"

CommDlg_Path.FileName = App.Path & "\white.JPG"

 

End Sub

