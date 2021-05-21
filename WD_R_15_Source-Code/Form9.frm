VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form manage_tag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Tag"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14775
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   14775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   0
      TabIndex        =   19
      Top             =   1200
      Width           =   2535
      Begin VB.Label Label2 
         BackColor       =   &H008B8B00&
         Caption         =   "WD_R_15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   20
         Top             =   5040
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   15855
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   5280
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   14
         Top             =   4800
         WhatsThisHelpID =   1
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9000
         TabIndex        =   13
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10680
         TabIndex        =   12
         Top             =   5640
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "Form9.frx":0000
         Left            =   2400
         List            =   "Form9.frx":000D
         TabIndex        =   11
         Top             =   5400
         Width           =   3255
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Back"
         DisabledPicture =   "Form9.frx":0029
         Height          =   495
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Form9.frx":1A90C
         Height          =   495
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   11895
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   2535
            Left            =   120
            TabIndex        =   0
            Top             =   120
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   4471
            _Version        =   393216
            BackColor       =   16777215
            BackColorFixed  =   16776960
            BackColorSel    =   65280
            BackColorBkg    =   -2147483635
            BackColorUnpopulated=   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "Serch Here........"
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2400
         TabIndex        =   6
         Top             =   4200
         WhatsThisHelpID =   1
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4080
         TabIndex        =   4
         Text            =   "Text6"
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Related Tag"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   480
         TabIndex        =   18
         Top             =   5400
         Width           =   2325
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tag Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   600
         TabIndex        =   17
         Top             =   4800
         Width           =   1875
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tag Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   18
         Left            =   360
         TabIndex        =   16
         Top             =   4200
         Width           =   2175
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15855
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Tag"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   12255
      End
   End
   Begin VB.Label Label152 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   2775
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label152 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Group Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7800
      TabIndex        =   21
      Top             =   5760
      Width           =   1455
   End
End
Attribute VB_Name = "manage_tag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()


If Text8.Text = "" Then
MsgBox "Enter the Tag Name ", vbInformation
Text8.SetFocus
Exit Sub
End If



sql = "select * from Tag where ID='" + Text5.Text + "'"
Set dataset = mddata(sql)

With dataset

.Update
!Tag_Name = Text3.Text
!Tag_code = Text8.Text
!Relataed_Tag = Combo1.Text

.Update

MsgBox "Successfully!", vbInformation

Unload Me
Me.Show
End With




End Sub

Private Sub Command3_Click()

If Not Text5.Text = "" Then
sql = "delete Tag where ID='" + Text5.Text + "'"
Set dataset = mddata(sql)
MsgBox "Done!!", vbInformation
Unload Me
Me.Show

Else
MsgBox "Select Tag!", vbInformation
MSHFlexGrid3.SetFocus
Exit Sub

End If
End Sub

Private Sub Command4_Click()
Unload Me
Home.Show
End Sub

Private Sub Command5_Click()
Unload Me
Tag_Managment.Show
End Sub

Private Sub Form_Load()
sql = "select * from Tag"
Set dataset = mddata(sql)
With dataset

Set MSHFlexGrid3.DataSource = dataset


End With
End Sub

Private Sub MSHFlexGrid3_Click()
MSHFlexGrid3.Col = 1
Text5.Text = MSHFlexGrid3.Text
End Sub

Private Sub Text2_Change()

If Not Text2.Text = "" Then
sql = "select * from Tag where ID like '%" + Text2.Text + "%'"
Set dataset = mddata(sql)
With dataset

Set MSHFlexGrid3.DataSource = dataset


End With
End If
End Sub

Private Sub Text5_Change()
sql = "select * from Tag where ID = '" + Text5.Text + "'"
Set dataset = mddata(sql)
With dataset

If dataset.RecordCount > 0 Then
Text3.Text = !Tag_Name
Text8.Text = !Tag_code
Combo1.Text = !Relataed_Tag

End If


End With
End Sub
