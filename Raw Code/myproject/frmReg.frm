VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Demo"
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Records"
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete Data"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdshowAll 
      Caption         =   "Find All Data"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Record"
      Height          =   1095
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
      Begin VB.TextBox txtCN 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblFN 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Label4"
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "City Name"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "First Name"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbID 
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "Find Data"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert Fix Values"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Search Record"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim ors As New ADODB.Recordset
 
Dim sqlQuery, vID, vName, vCity As String


Private Sub cmdDel_Click()
sqlQuery = "DELETE FROM tblinfo WHERE id=" & cmbID.Text & ";"
conn.Open
conn.Execute sqlQuery
conn.Close
MsgBox "Data Deleted"
End Sub

Private Sub cmdshowAll_Click()
sqlQuery = "SELECT * FROM tblinfo;"

conn.Open
    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
    Do Until ors.EOF
     vID = ors.Fields("id").Value
     vName = ors.Fields("name").Value
     vCity = ors.Fields("city").Value
     Print vID & "." & vName & " " & vCity
     ors.MoveNext
     
    Loop
  
conn.Close
End Sub

 

Private Sub cmdUpdate_Click()
Dim vTxtID As Integer
vTxtID = Val(cmbID.Text)
sqlQuery = "SELECT * FROM tblinfo WHERE id=" & vTxtID & ";"
conn.Open
    ors.Open sqlQuery, conn, adOpenDynamic, adLockOptimistic
    'ors.Update (2), txtCN.Text
    ors.Update "city", txtCN.Text
    MsgBox "Record Updated"
conn.Close
End Sub

Private Sub Command1_Click()
Print Time
Print Date
 
End Sub

Private Sub Form_Load()
conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "\database\xyz.mdb;Persist Security Info=False"


sqlQuery = "SELECT * FROM tblinfo;"

conn.Open
    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
    Do Until ors.EOF
     vID = ors.Fields("id").Value
      
     cmbID.AddItem vID
     ors.MoveNext
    Loop
  
conn.Close
End Sub

Private Sub cmdshow_Click()
Dim vTxtID As Integer
vTxtID = Val(cmbID.Text)
sqlQuery = "SELECT * FROM tblinfo WHERE id=" & vTxtID & ";"

conn.Open
    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
    If ors.EOF = True Then
     lblFN.Caption = "Record Not Found"
     txtCN.Text = "Record Not Found"
    Else
     vName = ors.Fields("name").Value
     vCity = ors.Fields("city").Value
     lblFN.Caption = vName
     txtCN.Text = vCity
     
    End If
  
conn.Close
End Sub
 

Private Sub cmdInsert_Click()
sqlQuery = "INSERT INTO tblinfo(name, city) VALUES('Virat','Yavatmal');"

conn.Open
conn.Execute sqlQuery
conn.Close
MsgBox "Data Inserted"

End Sub
 

