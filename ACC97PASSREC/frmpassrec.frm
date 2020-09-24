VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmpassrec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Recovery By Vipul Parmar"
   ClientHeight    =   4050
   ClientLeft      =   2490
   ClientTop       =   3105
   ClientWidth     =   6735
   Icon            =   "frmpassrec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6735
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Password"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5520
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.mdb"
      DialogTitle     =   "Select the Database"
      FileName        =   "c:\my documents\db1.mdb"
      Filter          =   "*.mdb"
      InitDir         =   "c:\my documents"
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Mail me at: vipul_matrix@yahoo.com"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   $"frmpassrec.frx":0442
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   6495
   End
   Begin VB.Label Label3 
      Caption         =   $"frmpassrec.frx":04F9
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Access 97 Database Password Recovery Tool "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "frmpassrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Long, s1 As String * 1, s2 As String * 1
Dim dbname As String
Dim passw As String
Dim mask As String

Private Sub Command1_Click()
Dim ff
Dim strfilter, strlines, alltext As String
ff = FreeFile
strfilter = "Access97 Files (*.mdb)|*.mdb"
cd1.Filter = strfilter
cd1.ShowOpen
If cd1.FileName <> "" Then
Text1.Text = cd1.FileName
Command2.Visible = True
Command2.SetFocus
End If
cd1.CancelError = False
Close #ff
End Sub

Private Sub Command2_Click()
   mask = Chr(78) & Chr(134) & Chr(251) & Chr(236) & _
          Chr(55) & Chr(93) & Chr(68) & Chr(156) & _
          Chr(250) & Chr(198) & Chr(94) & Chr(40) & Chr(230) & Chr(19)
' set the masking characters
   dbname = Text1.Text
   Open dbname For Binary As #1     ' open the database
   Seek #1, &H42
   For n = 1 To 14
   ' actual password recovery module
      s1 = Mid(mask, n, 1)
      s2 = Input(1, 1)
      If (Asc(s1) Xor Asc(s2)) <> 0 Then
         passw = passw & Chr(Asc(s1) Xor Asc(s2))
      End If
   Next
   Close 1
   If passw = "" Then
      MsgBox "No Password Found"
   Else
      MsgBox "The Password Is: " & passw
   End If
End Sub

Private Sub Command3_Click()
MsgBox "Give your feedback at vipul_matrix@yahoo.com", vbOKOnly, "Thanx for using"
Unload Me
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Text1_LostFocus()
Command1.SetFocus
End Sub
