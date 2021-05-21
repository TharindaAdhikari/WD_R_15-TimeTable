VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Login"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "Login.frx":0000
   ScaleHeight     =   6360
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Login"
      Height          =   495
      Left            =   6120
      MaskColor       =   &H000000FF&
      TabIndex        =   8
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   4800
      PasswordChar    =   "#"
      TabIndex        =   5
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   4800
      TabIndex        =   4
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00343434&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dhammika Enterprices"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   6495
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Height          =   855
         Left            =   6720
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   27.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   7080
         TabIndex        =   1
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SQL = "Select * From Usert Where Uname='" + Text1.Text + "' and Password='" + Text2.Text + "' "
Set dataset1 = mddata(SQL)
With dataset1
If .RecordCount > 0 Then





SQL = "Select * From U_log Where LogID='-1'"
Set dataset = mddata(SQL)

With dataset
.Addnew
!LogID = dataset1!UIDNO
!Time = Now
'!Date = Now
.Update
End With

Home.Show
Me.Hide
Else

MsgBox "Invalide Login Id or UserName!"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus

End If
End With

End Sub

Private Sub Label55_Click()
Select Case MsgBox("Are you sure you want to Exit ?", vbYesNo)
Case vbYes
End
Case vbNo

End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
If Not Text1.Text = "" Then

Text2.SetFocus
End If
End If



End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

Command1_Click
End If

End Sub
