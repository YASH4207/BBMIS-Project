VERSION 5.00
Begin VB.Form FrmAdminHome 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   12960
   ClientLeft      =   180
   ClientTop       =   -180
   ClientWidth     =   22395
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   ScaleHeight     =   12960
   ScaleWidth      =   22395
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Caption         =   "Frame1"
      Height          =   13455
      Left            =   -600
      TabIndex        =   0
      Top             =   -120
      Width           =   23535
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
         Left            =   20760
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   11880
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
         TabIndex        =   25
         Top             =   8520
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
         TabIndex        =   24
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   28
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
         TabIndex        =   23
         Top             =   6840
         UseMaskColor    =   -1  'True
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
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   6000
         Width           =   2655
      End
      Begin VB.ComboBox ComboGender 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "Form Admin Home.frx":0000
         Left            =   13440
         List            =   "Form Admin Home.frx":000D
         TabIndex        =   8
         Top             =   6720
         Width           =   4815
      End
      Begin VB.ComboBox ComboBloodGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "Form Admin Home.frx":0026
         Left            =   13440
         List            =   "Form Admin Home.frx":0042
         TabIndex        =   6
         Top             =   5520
         Width           =   4695
      End
      Begin VB.TextBox txtPhoneNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   13440
         MaxLength       =   10
         TabIndex        =   2
         Top             =   3240
         Width           =   4695
      End
      Begin VB.CommandButton cmdSubmit 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "SUBMIT"
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
         Left            =   10800
         MaskColor       =   &H00000000&
         TabIndex        =   9
         Top             =   7920
         Width           =   2175
      End
      Begin VB.TextBox txtDonersWeight 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   4920
         MaxLength       =   3
         TabIndex        =   7
         Top             =   6720
         Width           =   5895
      End
      Begin VB.TextBox txtDonerAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   4920
         TabIndex        =   5
         Top             =   5520
         Width           =   5895
      End
      Begin VB.TextBox txtAdharNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   13440
         MaxLength       =   12
         TabIndex        =   4
         Top             =   4320
         Width           =   4695
      End
      Begin VB.TextBox txtDonerAge 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   3
         Top             =   4320
         Width           =   5895
      End
      Begin VB.TextBox txtDonerName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   4920
         TabIndex        =   1
         Top             =   3240
         Width           =   5895
      End
      Begin VB.Label Label2 
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
         TabIndex        =   21
         Top             =   5520
         Width           =   3375
      End
      Begin VB.Image ImageMe 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3855
         Left            =   720
         Picture         =   "Form Admin Home.frx":0068
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   3330
      End
      Begin VB.Label LabelBloodGroup 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BLOOD GROUP"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   13395
         TabIndex        =   20
         Top             =   5160
         Width           =   1785
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
         Left            =   720
         TabIndex        =   19
         Top             =   12480
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
         Left            =   3720
         TabIndex        =   18
         Top             =   120
         Width           =   16695
      End
      Begin VB.Image BtnHOME2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   720
         Picture         =   "Form Admin Home.frx":3DD625
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DONER'S WEIGHT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4950
         TabIndex        =   17
         Top             =   6360
         Width           =   2115
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DONER GENDER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   13440
         TabIndex        =   16
         Top             =   6360
         Width           =   1965
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DONER ADDRESS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4920
         TabIndex        =   15
         Top             =   5160
         Width           =   2085
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DONER AGE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4920
         TabIndex        =   14
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ADHAR NUMBER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   13440
         TabIndex        =   13
         Top             =   3960
         Width           =   2025
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "PHONE NUMBER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   13440
         TabIndex        =   12
         Top             =   2880
         Width           =   1995
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Doner Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5040
         TabIndex        =   11
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ADD DATA"
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
         Left            =   10860
         TabIndex        =   10
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   12990
         Left            =   0
         MousePointer    =   3  'I-Beam
         Picture         =   "Form Admin Home.frx":3F63DB
         Stretch         =   -1  'True
         Top             =   120
         Width           =   23520
      End
   End
End
Attribute VB_Name = "FrmAdminHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ocon As New ADODB.Connection
Dim ors As New ADODB.Recordset
Dim checkValid As Integer
Dim ans

Private Sub btnABTUs_Click()
FrmAdminHome.Hide
FrmAbout.Show
End Sub

Private Sub BtnDelete_Click()
FrmAdminHome.Hide
FrmDelete.Show
End Sub

Private Sub BtnHOME2_Click()
FrmAdminHome.Hide
Frmadminmodule.Show
End Sub

Private Sub btnLogOut_Click()
FrmAdminHome.Hide
FrmAdminLogin.Show
End Sub

Private Sub BtnPrint_Click()
FrmAdminHome.Hide
FrmPrint.Show

End Sub

Private Sub btnSearch_Click()
FrmAdminHome.Hide
FrmSearch.Show
End Sub

Private Sub BtnShowAll_Click()
FrmAdminHome.Hide
FrmShowAll.Show

End Sub

Private Sub BtnUpdate_Click()
FrmAdminHome.Hide
FrmUpdate.Show
End Sub

Private Sub cmdSubmit_Click()

'checkValid = 0
'Validation
If Len(TxtDonerName.Text) = 0 Then
MsgBox "Enter Doner's Name", , "Doner's Name"
TxtDonerName.SetFocus
Exit Sub
End If

If Len(txtPhoneNumber.Text) = 0 Then
MsgBox "Enter Phone Number", , "Doner's Phone Number"
txtPhoneNumber.SetFocus
Exit Sub
End If

If (Len(txtPhoneNumber) < 10) Then
MsgBox "enter 10 digits only....", , "Doner's Phone Number"
txtPhoneNumber.SetFocus
Exit Sub
End If

If Len(TxtDonerAge.Text) = 0 Then
MsgBox "Enter Doner's Age", , "Doner's Age"
TxtDonerAge.SetFocus
Exit Sub
End If

If Len(txtAdharNumber.Text) = 0 Then
MsgBox "Enter Adhaar Number", , "Doner's Adhar Number"
txtAdharNumber.SetFocus
Exit Sub
End If

If (Len(txtAdharNumber) < 12) Then
MsgBox "enter 12 digits only....", , "Doner's Adhar Number"
txtAdharNumber.SetFocus
Exit Sub
End If

If Len(TxtDonerAddress.Text) = 0 Then
MsgBox "Enter Doner's Address", , "Doner's Address"
TxtDonerAddress.SetFocus
Exit Sub
End If

If Len(ComboBloodGroup.Text) = 0 Then
MsgBox "Enter Doner's Bood Group", , "Doner's BLood Group"
ComboBloodGroup.SetFocus
Exit Sub
End If

If Len(txtDonersWeight.Text) = 0 Then
MsgBox "Enter Doner's Weight", , "Doner's Weight"
txtDonersWeight.SetFocus
Exit Sub
End If

If Len(ComboGender.Text) = 0 Then
MsgBox "Enter Doner's Gender", , "Doner's Gender"
ComboGender.SetFocus
Exit Sub
End If


If checkValid = 0 Then
If MsgBox("Confirm add new Doner Record.", vbQuestion + vbYesNo) = vbYes Then

Dim ssql1 As String
ocon.Open
    ssql1 = ssql1 & "INSERT INTO tblname(DonerName,PhoneNumber,DonerAge,AdharNumber,DonerAddress,DonerBloodGroup,DonerWeight,DonerGender)"
    ssql1 = ssql1 & " values("
    
    ssql1 = ssql1 & "'" & UCase(Trim(TxtDonerName.Text)) & "',"
    ssql1 = ssql1 & "'" & UCase(Trim(txtPhoneNumber.Text)) & "',"
    ssql1 = ssql1 & "'" & UCase(Trim(TxtDonerAge.Text)) & "',"
    ssql1 = ssql1 & "'" & UCase(Trim(txtAdharNumber.Text)) & "',"
    ssql1 = ssql1 & "'" & UCase(Trim(TxtDonerAddress.Text)) & "',"
    ssql1 = ssql1 & "'" & UCase(Trim(ComboBloodGroup.Text)) & "',"
    ssql1 = ssql1 & "'" & UCase(Trim(txtDonersWeight.Text)) & "',"
    ssql1 = ssql1 & "'" & UCase(Trim(ComboGender.Text)) & "');"

'MsgBox ssql1

ocon.Execute ssql1
ocon.Close
MsgBox "Recordset added successfully"

TxtDonerName.Text = " "
txtPhoneNumber.Text = " "
TxtDonerAge.Text = " "
txtAdharNumber.Text = " "
TxtDonerAddress.Text = " "
ComboBloodGroup.Text = " "
txtDonersWeight.Text = " "
ComboGender.Text = " "

End If
End If
End Sub

Private Sub ContactUs_Click()
FrmAdminHome.Hide
FrmContactUs.Show
End Sub

Private Sub Form_Load()
ocon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "\database\rd.mdb;Persist Security Info=False"
End Sub

Private Sub txtAdharNumber_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
End If
End Sub

Private Sub TxtDonerAddress_Change()

End Sub

Private Sub txtDonerName_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8 Or KeyAscii = 32) Then
KeyAscii = 0
End If
End Sub

Private Sub txtDonersWeight_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
End If
End Sub

Private Sub txtPhoneNumber_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
End If
End Sub

Private Sub TxtDonerAge_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
End If
End Sub
