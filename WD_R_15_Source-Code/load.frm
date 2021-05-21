VERSION 5.00
Begin VB.Form Load 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Select Case MsgBox(""Are you sure you want to Exit ?"", vbYesNo)Select Case MsgBox(""Are you sure you want to Exit ?"", vbYesNo)"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   ScaleHeight     =   7215
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00343434&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
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
         TabIndex        =   3
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Height          =   855
         Left            =   6720
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time Table Managment"
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
         TabIndex        =   1
         Top             =   120
         Width           =   6495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   7695
      Begin VB.Frame Frame2 
         BackColor       =   &H008B8B00&
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   240
         TabIndex        =   8
         Top             =   0
         Width           =   7215
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
            Left            =   4680
            PasswordChar    =   "X"
            TabIndex        =   12
            Top             =   3840
            Width           =   2055
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
            Left            =   4680
            TabIndex        =   11
            Top             =   2880
            Width           =   2055
         End
         Begin VB.Label Label4 
            BackColor       =   &H008B8B00&
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2400
            TabIndex        =   10
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackColor       =   &H008B8B00&
            Caption         =   "User Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2400
            TabIndex        =   9
            Top             =   3000
            Width           =   1455
         End
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Caption         =   "Label2"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      Caption         =   "Label1"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
End
Attribute VB_Name = "Load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Text1.Text = "admin" And Text2.Text = "admin" Then
Me.Hide

Home.Show
Else
MsgBox "Login Fail!", vbCritical
End If

End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()

'
'sql = "select * from server_patch"
'Set dataset = mddata(sql)
'
'With dataset
' Text3.Text = !server_name
'End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
Select Case MsgBox("Are you sure you want to Exit ?", vbYesNo)
Case vbYes
End
Case vbNo

End Select
End Sub

Private Sub Label57_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.SetFocus
End If
End Sub
