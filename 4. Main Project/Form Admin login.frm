VERSION 5.00
Begin VB.Form FrmAdminLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Log In"
   ClientHeight    =   12960
   ClientLeft      =   -210
   ClientTop       =   -210
   ClientWidth     =   23040
   ControlBox      =   0   'False
   LinkTopic       =   "Admin LogIn"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12960
   ScaleWidth      =   23040
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CommandEnd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "END"
      DisabledPicture =   "Form Admin login.frx":0000
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
      Left            =   7680
      TabIndex        =   4
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton BtnSignIn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "SIGN IN"
      DisabledPicture =   "Form Admin login.frx":48B0
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
      Left            =   7680
      TabIndex        =   3
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Caption         =   "Frame1"
      Height          =   13335
      Left            =   -360
      TabIndex        =   0
      Top             =   -240
      Width           =   23520
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   5880
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   5280
         Width           =   5895
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         TabIndex        =   1
         Top             =   3840
         Width           =   5895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT LOG IN"
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
         Left            =   6600
         TabIndex        =   10
         Top             =   2040
         Width           =   4245
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
         Left            =   480
         TabIndex        =   9
         Top             =   12600
         Width           =   8835
      End
      Begin VB.Image ImgLogo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   480
         Picture         =   "Form Admin login.frx":9160
         Stretch         =   -1  'True
         Top             =   360
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
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   16695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "USER ID:"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         TabIndex        =   5
         Top             =   3480
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD:"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   12990
         Left            =   240
         Picture         =   "Form Admin login.frx":21F16
         Stretch         =   -1  'True
         Top             =   240
         Width           =   23160
      End
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "DESIGNED AND DEVOLEPED BY YASH DARNE"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   8160
      Width           =   15465
   End
End
Attribute VB_Name = "FrmAdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim conn As New ADODB.Connection
Dim ors As New ADODB.Recordset
Dim sqlQuery As String

Public LoginSucceeded As Boolean
Public vUsername As String
Public vPassword As String

Private Sub BtnSignIn_Click()

vUsername = txtUserName.Text
vPassword = txtPassword.Text


sqlQuery = "SELECT * FROM tbllogin WHERE username='" & vUsername & "' and password='" & vPassword & "';"

conn.Open
    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
    If ors.EOF = True Then
     LoginSucceeded = False
    Else
     LoginSucceeded = True
    End If
  
conn.Close

    If LoginSucceeded = True Then
        FrmAdminLogin.Hide
        'Admin Form Show Here
        Frmadminmodule.Show
        
    ElseIf Len(txtUserName.Text) = 0 Then
    MsgBox "Enter User Name", , "User Name"
    txtUserName.SetFocus
    ElseIf Len(txtPassword.Text) = 0 Then
    MsgBox "Enter Password", , "Password"
    txtPassword.SetFocus
    
    ElseIf LoginSucceeded = False Then
        MsgBox "Invalid Username or Password, please try again!", , "Login"
        txtPassword.SetFocus
        
       
    End If
    
    
    If txtUserName = "" Then
    txtUserName.SetFocus
    ElseIf txtPassword = "" Then
    txtPassword.SetFocus
    End If
    
    If LoginSucceeded = True Then
    txtUserName = ""
    txtPassword = ""
    End If
        
End Sub

Private Sub CommandEnd_Click()
End
End Sub

Private Sub Form_Load()
conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "\database\xyz.mdb;Persist Security Info=False"
End Sub
