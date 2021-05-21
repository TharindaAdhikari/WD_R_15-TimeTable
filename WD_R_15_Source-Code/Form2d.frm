VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Add_session 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add_session"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12045
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   9975
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2280
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Form2d.frx":0000
         Height          =   495
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4695
         Left            =   720
         TabIndex        =   6
         Top             =   960
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   8281
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Select Lectures And Tag"
         TabPicture(0)   =   "Form2d.frx":1A8E3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label152(18)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label152(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label152(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label152(9)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Combo1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Combo2"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Text3"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Command3"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Combo5"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Select Group And Subjects"
         TabPicture(1)   =   "Form2d.frx":1A8FF
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Text2"
         Tab(1).Control(1)=   "Command2"
         Tab(1).Control(2)=   "Command1"
         Tab(1).Control(3)=   "Text1"
         Tab(1).Control(4)=   "Text8"
         Tab(1).Control(5)=   "Combo4"
         Tab(1).Control(6)=   "Combo3"
         Tab(1).Control(7)=   "Label152(8)"
         Tab(1).Control(8)=   "Label152(7)"
         Tab(1).Control(9)=   "Label152(6)"
         Tab(1).Control(10)=   "Label152(5)"
         Tab(1).Control(11)=   "Label152(4)"
         Tab(1).Control(12)=   "Label152(3)"
         Tab(1).ControlCount=   13
         Begin VB.ComboBox Combo5 
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
            ItemData        =   "Form2d.frx":1A91B
            Left            =   2640
            List            =   "Form2d.frx":1A91D
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1800
            Width           =   2535
         End
         Begin VB.TextBox Text2 
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
            Height          =   390
            Left            =   -71040
            TabIndex        =   26
            Top             =   2640
            WhatsThisHelpID =   1
            Width           =   615
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Submit"
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
            Left            =   -69840
            TabIndex        =   25
            Top             =   3840
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Clear"
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
            Left            =   -68280
            TabIndex        =   24
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox Text1 
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
            Height          =   390
            Left            =   -72360
            TabIndex        =   22
            Top             =   2640
            WhatsThisHelpID =   1
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
            Height          =   390
            Left            =   -72360
            TabIndex        =   21
            Top             =   2160
            WhatsThisHelpID =   1
            Width           =   2535
         End
         Begin VB.ComboBox Combo4 
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
            ItemData        =   "Form2d.frx":1A91F
            Left            =   -72360
            List            =   "Form2d.frx":1A921
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1680
            Width           =   2535
         End
         Begin VB.ComboBox Combo3 
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
            ItemData        =   "Form2d.frx":1A923
            Left            =   -72360
            List            =   "Form2d.frx":1A925
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Clear"
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
            Left            =   6600
            TabIndex        =   14
            Top             =   3840
            Width           =   1335
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Height          =   450
            Left            =   2880
            TabIndex        =   13
            Top             =   3120
            WhatsThisHelpID =   1
            Width           =   4575
         End
         Begin VB.ComboBox Combo2 
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
            ItemData        =   "Form2d.frx":1A927
            Left            =   2760
            List            =   "Form2d.frx":1A929
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2520
            Width           =   2535
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
            ItemData        =   "Form2d.frx":1A92B
            Left            =   2640
            List            =   "Form2d.frx":1A92D
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Lecturer2"
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
            Index           =   9
            Left            =   -135
            TabIndex        =   32
            Top             =   1920
            Width           =   3675
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minutes"
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
            Index           =   8
            Left            =   -70335
            TabIndex        =   27
            Top             =   2760
            Width           =   885
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hrs"
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
            Index           =   7
            Left            =   -71760
            TabIndex        =   23
            Top             =   2760
            Width           =   645
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duration"
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
            Left            =   -75000
            TabIndex        =   18
            Top             =   2640
            Width           =   3315
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Of Student"
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
            Index           =   5
            Left            =   -75240
            TabIndex        =   17
            Top             =   2160
            Width           =   3645
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Subject"
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
            Index           =   4
            Left            =   -74520
            TabIndex        =   16
            Top             =   1680
            Width           =   2205
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Group"
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
            Index           =   3
            Left            =   -74625
            TabIndex        =   15
            Top             =   1200
            Width           =   2415
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Selected Lecturer(s)"
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
            Index           =   2
            Left            =   0
            TabIndex        =   12
            Top             =   3240
            Width           =   3435
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Tag"
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
            Index           =   1
            Left            =   -45
            TabIndex        =   10
            Top             =   2520
            Width           =   3375
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Lecturer1"
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
            Left            =   -135
            TabIndex        =   7
            Top             =   1320
            Width           =   3675
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
      Begin VB.Label Label4 
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
         Left            =   480
         TabIndex        =   28
         Top             =   6360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add Session"
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
         Left            =   0
         TabIndex        =   2
         Top             =   480
         Width           =   11775
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
         Left            =   12600
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label Label152 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Lecturer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4005
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Add_session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Text3.Text = Combo1.Text
End Sub

Private Sub Combo1_Click()
Text3.Text = Combo1.Text
End Sub

Private Sub Combo5_Change()
Combo5_LostFocus
End Sub

Private Sub Combo5_Click()
Combo5_LostFocus
End Sub


Private Sub Combo5_LostFocus()

Text3.Text = Combo1.Text + "," + Combo5.Text
End Sub


Private Sub Command1_Click()
Unload Me
Me.Show
End Sub

Private Sub Command2_Click()



If Text2.Text = "" Then
MsgBox "Error!", vbInformation
Text2.SetFocus
Exit Sub
End If

If Combo1.Text = "" Then
MsgBox "Error!", vbInformation
Combo1.SetFocus
Exit Sub
End If

If Combo2.Text = "" Then
MsgBox "Error!", vbInformation
Combo2.SetFocus
Exit Sub
End If

If Combo3.Text = "" Then
MsgBox "Error!", vbInformation
Combo3.SetFocus
Exit Sub
End If

If Combo4.Text = "" Then
MsgBox "Error!", vbInformation
Combo4.SetFocus
Exit Sub
End If

If Text8.Text = "" Then
MsgBox "Error!", vbInformation
Text8.SetFocus
Exit Sub
End If

If Text1.Text = "" Then
MsgBox "Error!", vbInformation
Text1.SetFocus
Exit Sub
End If

If Text2.Text = "" Then
MsgBox "Ener the proper value!", vbInformation
Text2.SetFocus
Exit Sub
End If

Text4.Text = val(Text1.Text)
Text5.Text = val(Text2.Text) / 30

sql = "select * from session"
Set dataset5 = mddata(sql)
With dataset5

.AddNew
!lecturer = Combo1.Text

If Not Combo5.Text = "" Then
!lecturer2 = Combo5.Text
Else
!lecturer2 = ""
End If
!Tag = Combo2.Text
!Group = Combo3.Text
!subject = Combo4.Text
!nostudent = val(Text8.Text)
!duration_hours = val(Text1.Text)
!duration_minuts = val(Text2.Text)
!val = val(Text4.Text) + val(Text5.Text)
.Update

End With
MsgBox "Done", vbInformation

Unload Me
Me.Show

End Sub

Private Sub Command3_Click()
Unload Me
Me.Show
End Sub

Private Sub Command4_Click()
Unload Me
Home.Show
End Sub

Private Sub Form_Activate()
Combo1.SetFocus
End Sub

Private Sub Form_Load()
sql = "select * from Lecturer"
Set dataset = mddata(sql)

With dataset
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo1.AddItem dataset!lec_name
Combo5.AddItem dataset!lec_name
.MoveNext
Wend

End If
End With


sql = "select * from Tag"
Set Dataset2 = mddata(sql)

With Dataset2
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo2.AddItem Dataset2!Tag_Name
.MoveNext
Wend

End If
End With


sql = "select * from Student_group"
Set dataset3 = mddata(sql)

With dataset3
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo3.AddItem dataset3!group_id
.MoveNext
Wend

End If
End With

sql = "select * from Subject"
Set dataset4 = mddata(sql)

With dataset4
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo4.AddItem dataset4!Sub_name
.MoveNext
Wend

End If
End With


End Sub

Private Sub Text1_LostFocus()
Text1.Text = Replace(Text1.Text, ".", "")
Text1.Text = Replace(Text1.Text, "'", "")
End Sub


Private Sub Text2_LostFocus()

If Not val(Text2.Text) = 30 And Not val(Text2.Text) = 0 Then
MsgBox "Error!", vbInformation
Text2.Text = 0
Text2.SetFocus
Exit Sub
End If

End Sub


