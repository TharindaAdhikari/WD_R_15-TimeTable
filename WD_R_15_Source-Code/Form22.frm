VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Session_And_Not_Time_Manage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Session And Not Availble Time"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14700
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   14700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   15855
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   5280
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Caption         =   "Select Session"
         DisabledPicture =   "Form22.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   4920
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   4200
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "View"
         DisabledPicture =   "Form22.frx":1A8E3
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4920
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Caption         =   "Add Session"
         DisabledPicture =   "Form22.frx":351C6
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4920
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   6240
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4440
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3375
         Left            =   480
         TabIndex        =   12
         Top             =   1320
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5953
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Consecutive"
         TabPicture(0)   =   "Form22.frx":4FAA9
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Parallel"
         TabPicture(1)   =   "Form22.frx":4FAC5
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSHFlexGrid1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "None Overlap"
         TabPicture(2)   =   "Form22.frx":4FAE1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "MSHFlexGrid2"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Not Available Time"
         TabPicture(3)   =   "Form22.frx":4FAFD
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Combo2"
         Tab(3).Control(1)=   "Combo4"
         Tab(3).Control(2)=   "Combo3"
         Tab(3).Control(3)=   "Combo1"
         Tab(3).Control(4)=   "Label6"
         Tab(3).Control(5)=   "Label5"
         Tab(3).Control(6)=   "Label4"
         Tab(3).Control(7)=   "Label3"
         Tab(3).ControlCount=   8
         Begin VB.ComboBox Combo2 
            Enabled         =   0   'False
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
            ItemData        =   "Form22.frx":4FB19
            Left            =   -72960
            List            =   "Form22.frx":4FB1B
            TabIndex        =   29
            Top             =   1560
            Width           =   3255
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
            ItemData        =   "Form22.frx":4FB1D
            Left            =   -68520
            List            =   "Form22.frx":4FB1F
            TabIndex        =   28
            Top             =   960
            Width           =   3255
         End
         Begin VB.ComboBox Combo3 
            Enabled         =   0   'False
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
            ItemData        =   "Form22.frx":4FB21
            Left            =   -72960
            List            =   "Form22.frx":4FB23
            TabIndex        =   27
            Top             =   2160
            Width           =   3255
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
            ItemData        =   "Form22.frx":4FB25
            Left            =   -72960
            List            =   "Form22.frx":4FB27
            TabIndex        =   22
            Top             =   960
            Width           =   3255
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   2535
            Left            =   120
            TabIndex        =   13
            Top             =   480
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   2535
            Left            =   -74880
            TabIndex        =   14
            Top             =   480
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   4471
            _Version        =   393216
            BackColor       =   14737632
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   2535
            Left            =   -74880
            TabIndex        =   15
            Top             =   480
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
         Begin VB.Label Label6 
            Caption         =   "Group"
            Height          =   495
            Left            =   -74400
            TabIndex        =   26
            Top             =   2160
            Width           =   2055
         End
         Begin VB.Label Label5 
            Caption         =   " Lecturers"
            Height          =   495
            Left            =   -74400
            TabIndex        =   25
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Time"
            Height          =   495
            Left            =   -69240
            TabIndex        =   24
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Select Session"
            Height          =   495
            Left            =   -74400
            TabIndex        =   23
            Top             =   960
            Width           =   2055
         End
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Text            =   "Text6"
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
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
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Form22.frx":4FB29
         Height          =   495
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   1440
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   19
         Top             =   5640
         Width           =   5415
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   0
      TabIndex        =   2
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
         Left            =   480
         TabIndex        =   3
         Top             =   5160
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
      Width           =   15855
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Session And Not Availble Time"
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
         TabIndex        =   1
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   5760
      Width           =   1455
   End
End
Attribute VB_Name = "Session_And_Not_Time_Manage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
sql = "select * from session where ID='" + Combo1.Text + "'"
Set Dataset2 = mddata(sql)
With Dataset2
Combo2.Text = !lecturer
Combo3.Text = !Group
End With
End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub

Private Sub Command1_Click()

If SSTab1.Caption = "Consecutive" Then

If Text5.Text = "" Then
MsgBox "Select the session", vbInformation
Exit Sub
End If

sql = "select * from Consecutive where session_id='" + Text5.Text + "'"
Set Dataset2 = mddata(sql)
If Dataset2.RecordCount > 0 Then
MsgBox "Session Already Added!", vbInformation
Exit Sub
End If


sql = "select * from Consecutive"
Set dataset = mddata(sql)
With dataset
.AddNew
!session_id = Text5.Text
.Update
Label1.Caption = "Consecutive Added Successfully!"
End With
Exit Sub
End If

If SSTab1.Caption = "Parallel" Then

If Text1.Text = "" Then
MsgBox "Select the session", vbInformation
Exit Sub
End If


sql = "select * from session where id='" + Text8.Text + "'"
Set Dataset2 = mddata(sql)
Text9.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room)

sql = "select * from session where id='" + Text1.Text + "'"
Set dataset = mddata(sql)
With dataset
.Update
!parallel = Text9.Text
.Update
Label1.Caption = "Parallel Added Successfully!"

End With
sql = "update session set [Flag]=0 where id='" + Text8.Text + "'"
Set dataset4 = mddata(sql)
Text1.Text = ""
Text8.Text = ""
Exit Sub
End If

If SSTab1.Caption = "None Overlap" Then

If Text3.Text = "" Then
MsgBox "Select the session", vbInformation
Exit Sub
End If

sql = "select * from None_Overlap where session_id='" + Text3.Text + "'"
Set Dataset2 = mddata(sql)
If Dataset2.RecordCount > 0 Then
MsgBox "Session Already Added!", vbInformation
Exit Sub
End If

sql = "select * from None_Overlap"
Set dataset = mddata(sql)
With dataset
.AddNew
!session_id = Text3.Text
.Update
Label1.Caption = "None Overlap Added Successfully!"
End With
Exit Sub
End If

If SSTab1.Caption = "Not Available Time" Then

If Combo4.Text = "" Or Combo1.Text = "" Then
MsgBox "Error!", vbInformation
Exit Sub
End If

sql = "select * from Non_availble_time where session_id='" + Combo1.Text + "'"
Set Dataset2 = mddata(sql)
If Dataset2.RecordCount > 0 Then
MsgBox "Session Already Added!", vbInformation
Exit Sub
End If

sql = "select * from session where id='" + Combo1.Text + "'"
Set dataset = mddata(sql)
With dataset
.Update
!not_ava_time = Combo4.Text
!val = !val + 1
.Update
Label1.Caption = "Not Available Time Added Successfully!"
End With
End If
Exit Sub
End Sub

Private Sub Command2_Click()
If SSTab1.Caption = "Not Available Time" Then
Unload Me
Form3.Show
End If
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Then
MSHFlexGrid1.Col = 1
Text1.Text = MSHFlexGrid1.Text
Label1.Caption = "One Session Added"
Exit Sub
End If

If Text8.Text = "" Then
MSHFlexGrid1.Col = 1
Text8.Text = MSHFlexGrid1.Text
Label1.Caption = "Second Session Added"
Command1.Enabled = True
Exit Sub
End If

If Not Text1.Text = "" And Not Text8.Text = "" Then
Label1.Caption = "Two Sessions Already Selected"
Exit Sub
End If
End Sub

Private Sub Command4_Click()
Unload Me
Home.Show
End Sub

Private Sub Form_Load()
sql = "select * from session where flag=1 "
Set dataset = mddata(sql)
With dataset

Set MSHFlexGrid1.DataSource = dataset
Set MSHFlexGrid2.DataSource = dataset
Set MSHFlexGrid3.DataSource = dataset

MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 500
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1600
MSHFlexGrid1.ColWidth(5) = 1800
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 0

MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(1) = 500
MSHFlexGrid2.ColWidth(2) = 2000
MSHFlexGrid2.ColWidth(3) = 1000
MSHFlexGrid2.ColWidth(4) = 1600
MSHFlexGrid2.ColWidth(5) = 1800
MSHFlexGrid2.ColWidth(6) = 1200
MSHFlexGrid2.ColWidth(7) = 1200
MSHFlexGrid2.ColWidth(8) = 1200
MSHFlexGrid2.ColWidth(9) = 1200
MSHFlexGrid2.ColWidth(10) = 0

MSHFlexGrid3.ColWidth(0) = 0
MSHFlexGrid3.ColWidth(1) = 500
MSHFlexGrid3.ColWidth(2) = 2000
MSHFlexGrid3.ColWidth(3) = 1000
MSHFlexGrid3.ColWidth(4) = 1600
MSHFlexGrid3.ColWidth(5) = 1800
MSHFlexGrid3.ColWidth(6) = 1200
MSHFlexGrid3.ColWidth(7) = 1200
MSHFlexGrid3.ColWidth(8) = 1200
MSHFlexGrid3.ColWidth(9) = 1200
MSHFlexGrid3.ColWidth(10) = 0


End With

sql = "select * from session"
Set dataset = mddata(sql)
Combo1.Clear
With dataset
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo1.AddItem dataset!id
.MoveNext
Wend

End If
End With

sql = "select * from Time_Table_Main"
Set dataset4 = mddata(sql)
Combo4.Clear
With dataset4
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo4.AddItem dataset4!Time
.MoveNext
Wend

End If
End With
End Sub

Private Sub MSHFlexGrid2_Click()
MSHFlexGrid3.Col = 1
Text3.Text = MSHFlexGrid3.Text
End Sub

Private Sub MSHFlexGrid3_Click()
MSHFlexGrid3.Col = 1
Text5.Text = MSHFlexGrid3.Text
End Sub

Private Sub MSHFlexGrid4_Click()
MSHFlexGrid3.Col = 1
Text7.Text = MSHFlexGrid3.Text
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Caption = "Parallel" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub

