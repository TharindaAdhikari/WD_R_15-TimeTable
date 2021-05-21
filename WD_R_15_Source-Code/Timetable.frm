VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Timetable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timetable"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14505
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   14505
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   8400
      TabIndex        =   46
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command14"
      Height          =   615
      Left            =   0
      TabIndex        =   42
      Top             =   9240
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   11400
      TabIndex        =   41
      Text            =   "0"
      Top             =   9600
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   10200
      TabIndex        =   40
      Text            =   "0"
      Top             =   9600
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   9000
      TabIndex        =   39
      Text            =   "0"
      Top             =   9600
      Width           =   855
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   7320
      TabIndex        =   38
      Text            =   "0"
      Top             =   9600
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   6000
      TabIndex        =   37
      Text            =   "0"
      Top             =   9600
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   4920
      TabIndex        =   36
      Text            =   "0"
      Top             =   9720
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4080
      TabIndex        =   35
      Text            =   "0"
      Top             =   9720
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3000
      TabIndex        =   34
      Text            =   "0"
      Top             =   9240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   8760
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   8880
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   9720
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   2520
      TabIndex        =   0
      Top             =   2760
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   10398
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lecuter"
      TabPicture(0)   =   "Timetable.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Combo1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "MSHFlexGrid1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Combo4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command19"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Student"
      TabPicture(1)   =   "Timetable.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Combo2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "MSHFlexGrid2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command10"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command20"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Locations"
      TabPicture(2)   =   "Timetable.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Combo3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "MSHFlexGrid3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command11"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command21"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CommandButton Command21 
         BackColor       =   &H0000FFFF&
         Caption         =   "Print"
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
         Left            =   -64920
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H0000FFFF&
         Caption         =   "Print"
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
         Left            =   -64920
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H0000FFFF&
         Caption         =   "Print"
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
         Left            =   10200
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "Timetable.frx":0054
         Left            =   4560
         List            =   "Timetable.frx":0056
         TabIndex        =   33
         Top             =   840
         Width           =   3135
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0000FF00&
         Caption         =   "Genarate"
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
         Left            =   -67080
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0000FF00&
         Caption         =   "Genarate"
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
         Left            =   -67200
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   720
         Width           =   1935
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   4215
         Left            =   -74760
         TabIndex        =   21
         Top             =   1560
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   7435
         _Version        =   393216
         BackColor       =   -2147483635
         BackColorBkg    =   -2147483635
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   4215
         Left            =   -74760
         TabIndex        =   20
         Top             =   1560
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   7435
         _Version        =   393216
         BackColor       =   -2147483635
         BackColorFixed  =   -2147483638
         BackColorSel    =   -2147483638
         BackColorBkg    =   -2147483635
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4215
         Left            =   240
         TabIndex        =   19
         Top             =   1560
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   7435
         _Version        =   393216
         BackColor       =   -2147483635
         BackColorFixed  =   -2147483646
         BackColorBkg    =   -2147483635
         ScrollTrack     =   -1  'True
         Enabled         =   0   'False
         AllowUserResizing=   2
         Appearance      =   0
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command2"
         Height          =   495
         Left            =   -65400
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0000FF00&
         Caption         =   "Genarate"
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
         Left            =   8040
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -73320
         TabIndex        =   15
         Top             =   840
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -73560
         TabIndex        =   13
         Top             =   840
         Width           =   3615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "Timetable.frx":0058
         Left            =   1560
         List            =   "Timetable.frx":005A
         TabIndex        =   10
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   14
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74520
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Lecturer"
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
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14535
      Begin VB.CommandButton Command17 
         Caption         =   "Command17"
         Height          =   255
         Left            =   2760
         TabIndex        =   45
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Command16"
         Height          =   495
         Left            =   7920
         TabIndex        =   44
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command15 
         Caption         =   "suspend"
         Height          =   375
         Left            =   9120
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Command13"
         Height          =   495
         Left            =   2640
         TabIndex        =   32
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Timetable.frx":005C
         Height          =   735
         Left            =   12960
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "insert main table"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time Table"
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
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   11775
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   2280
      TabIndex        =   7
      Top             =   1320
      Width           =   12495
      Begin VB.CommandButton Command18 
         BackColor       =   &H008B8B00&
         Caption         =   "Print Full Time Table"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4800
         TabIndex        =   31
         Text            =   "Text4"
         Top             =   7560
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Genarate "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
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
         Left            =   8280
         TabIndex        =   8
         Top             =   6240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   -240
      TabIndex        =   6
      Top             =   720
      Width           =   2655
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   600
         TabIndex        =   25
         Text            =   "X"
         Top             =   5280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "try fillter"
         Height          =   735
         Left            =   600
         TabIndex        =   24
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "teset insert copy table"
         Height          =   855
         Left            =   480
         TabIndex        =   23
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Main"
         Height          =   1335
         Left            =   480
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label5 
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
         Height          =   735
         Left            =   600
         TabIndex        =   30
         Top             =   6960
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Timetable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim monday As Double
Dim tuesday As Double
Dim wendsday As Double
Dim thursday As Double
Dim friday As Double
Dim saturday As Double
Dim sunday As Double

Dim i As Integer
Dim y As Integer
Dim id As Integer
Dim val As Integer
Dim x As Integer
Dim n As Integer
Dim e As Integer
Dim k As Integer

Dim weight As Integer

Dim mo As Integer
Dim tu As Integer
Dim we As Integer
Dim th As Integer
Dim fr As Integer
Dim st As Integer
Dim su As Integer


Private Sub Command1_Click()
Command14_Click
Command12_Click
Command17_Click

Form_Load
'Command3_Click

MsgBox "Done", vbInformation
End Sub

Private Sub Command10_Click()


Command5_Click

If Not Combo2.Text = "" Then

sql = "select * from Time_Tablecopy "
Set Dataset2 = mddata(sql)

If Dataset2.RecordCount > 0 Then
With Dataset2
Dataset2.MoveFirst
While Dataset2.EOF = False

sql = "truncate table filtertag "
Set dataset3 = mddata(sql)

sql = "select * from filtertag "
Set dataset4 = mddata(sql)
dataset4.AddNew
dataset4!combo = Combo1.Text
dataset4!monday = !monday
dataset4!tuesday = !tuesday
dataset4!Wednesday = !Wednesday
dataset4!thursday = !thursday
dataset4!friday = !friday
dataset4!saturday = !saturday
dataset4!sunday = !sunday

dataset4.Update



id = !id



sql = "select * from filtertag where Monday like'%" + Combo2.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then

sql = "update Time_Tablecopy SET[Monday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)

End If



sql = "select * from filtertag where Tuesday like'%" + Combo2.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then

sql = "update Time_Tablecopy SET[tuesday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)

End If

sql = "select * from filtertag where Wednesday like'%" + Combo2.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Wednesday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)

End If


sql = "select * from filtertag where Thursday like'%" + Combo2.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Thursday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)
End If


sql = "select * from filtertag where Friday like'%" + Combo2.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Friday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)
End If

sql = "select * from filtertag where Saturday like'%" + Combo2.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Saturday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)
Else
End If


sql = "select * from filtertag where Sunday like'%" + Combo2.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Sunday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)
Else
End If




Dataset2.MoveNext
Wend

'MsgBox "Done", vbInformation
End With

End If
sql = "select * from Time_Tablecopy "
Set Dataset2 = mddata(sql)
Set MSHFlexGrid2.DataSource = Dataset2




Else

MsgBox "Select Group", vbInformation
Combo2.SetFocus
Exit Sub

End If


End Sub

Private Sub Command11_Click()


Command5_Click

If Not Combo3.Text = "" Then

sql = "select * from Time_Tablecopy "
Set Dataset2 = mddata(sql)

If Dataset2.RecordCount > 0 Then
With Dataset2
Dataset2.MoveFirst
While Dataset2.EOF = False

sql = "truncate table filtertag "
Set dataset3 = mddata(sql)

sql = "select * from filtertag "
Set dataset4 = mddata(sql)
dataset4.AddNew
dataset4!combo = Combo1.Text
dataset4!monday = !monday
dataset4!tuesday = !tuesday
dataset4!Wednesday = !Wednesday
dataset4!thursday = !thursday
dataset4!friday = !friday
dataset4!saturday = !saturday
dataset4!sunday = !sunday

dataset4.Update



id = !id



sql = "select * from filtertag where Monday like'%" + Combo3.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then

sql = "update Time_Tablecopy SET[Monday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)

End If



sql = "select * from filtertag where Tuesday like'%" + Combo3.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then

sql = "update Time_Tablecopy SET[tuesday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)

End If

sql = "select * from filtertag where Wednesday like'%" + Combo3.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Wednesday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)

End If


sql = "select * from filtertag where Thursday like'%" + Combo3.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Thursday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)
End If


sql = "select * from filtertag where Friday like'%" + Combo3.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Friday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)
End If

sql = "select * from filtertag where Saturday like'%" + Combo3.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Saturday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)
Else
End If


sql = "select * from filtertag where Sunday like'%" + Combo3.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Sunday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)
Else
End If




Dataset2.MoveNext
Wend

'MsgBox "Done", vbInformation
End With

End If
sql = "select * from Time_Tablecopy "
Set Dataset2 = mddata(sql)
Set MSHFlexGrid3.DataSource = Dataset2
Else

MsgBox "Select Location", vbInformation
Combo3.SetFocus
Exit Sub

End If

End Sub


Private Sub Command12_Click()
sql = "truncate table Time_Table "
Set dataset = mddata(sql)


sql = "insert into Time_Table SELECT *FROM Time_Table_Main"
Set dataset = mddata(sql)
End Sub

Private Sub Command13_Click()
Form3.Show
End Sub

Private Sub Command14_Click()
e = 0
k = 0
n = 1
mo = 1
tu = 1
we = 1
th = 1
fr = 1
st = 1
su = 1

sql = "update session set[day]=1 where flag = 0 "
Set dataset4 = mddata(sql)

sql = "select * from session where flag = 1 "
Set Dataset2 = mddata(sql)

If Dataset2.RecordCount > 0 Then
Dataset2.MoveFirst
While Dataset2.EOF = False
k = Dataset2!val
Text6.Text = Dataset2!id




Dataset2.MoveNext
Wend

End If

End Sub


Private Sub Command15_Click()
If i < 9 Then

If y < 8 Then

For y = y To 7

If monday = 1 And y = 1 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)
dataset3.Update
dataset3!monday = Text1.Text
dataset3.Update
y = y + 1
Exit Sub
End If


If tuesday = 1 And y = 2 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)
dataset3.Update
dataset3!tuesday = Text1.Text
dataset3.Update
y = y + 1
Exit Sub
End If


If wendsday = 1 And y = 3 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)
dataset3.Update
dataset3!Wednesday = Text1.Text
dataset3.Update
y = y + 1
Exit Sub
End If

If thursday = 1 And y = 4 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)
dataset3.Update
dataset3!thursday = Text1.Text
dataset3.Update
y = y + 1
Exit Sub
End If



If friday = 1 And y = 5 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)
dataset3.Update
dataset3!friday = Text1.Text
dataset3.Update
y = y + 1
Exit Sub
End If



If saturday = 1 And y = 6 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)
dataset3.Update
dataset3!saturday = Text1.Text
dataset3.Update
y = y + 1
Exit Sub
End If

If sunday = 1 And y = 7 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)
dataset3.Update
dataset3!sunday = Text1.Text
dataset3.Update
y = y + 1
Exit Sub
End If



Next y
Text2_Change

Else
y = 1
End If
Else

MsgBox "All Days Are Full! ", vbInformation
Command3.SetFocus
Exit Sub
End If

End Sub

Private Sub Command16_Click()
Text1_Change
End Sub

Private Sub Command17_Click()



sql = "select * from session where day= 'monday' "
Set Dataset2 = mddata(sql)
With Dataset2
i = 1
If .RecordCount > 0 Then
While Dataset2.EOF = False
weight = Dataset2!val
If IsNull(Dataset2!not_ava_time) = False Then
Text4.Text = Dataset2!not_ava_time
Else
Text4.Text = ""
End If



If IsNull(Dataset2!parallel) = True Then
Text1.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room)
Else
Text1.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room) + " " + " " + "" + "" + RTrim(Dataset2!parallel)
End If
.MoveNext
Wend
End If
End With





sql = "select * from session where day= 'Tuesday' "
Set Dataset2 = mddata(sql)
With Dataset2
i = 1
If .RecordCount > 0 Then
While Dataset2.EOF = False
weight = Dataset2!val
If IsNull(Dataset2!not_ava_time) = False Then
Text4.Text = Dataset2!not_ava_time
Else
Text4.Text = ""
End If

If IsNull(Dataset2!parallel) = True Then
Text7.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room)
Else
Text7.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room) + " " + " " + "" + "" + RTrim(Dataset2!parallel)
End If
.MoveNext
Wend
End If
End With


sql = "select * from session where day= 'Wednesday' "
Set Dataset2 = mddata(sql)
With Dataset2
i = 1
If .RecordCount > 0 Then
While Dataset2.EOF = False
weight = Dataset2!val
If IsNull(Dataset2!not_ava_time) = False Then
Text4.Text = Dataset2!not_ava_time
Else
Text4.Text = ""
End If

If IsNull(Dataset2!parallel) = True Then
Text8.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room)
Else
Text8.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room) + " " + " " + "" + "" + RTrim(Dataset2!parallel)
End If
.MoveNext
Wend
End If
End With


sql = "select * from session where day= 'Thursday' "
Set Dataset2 = mddata(sql)
With Dataset2
i = 1
If .RecordCount > 0 Then
While Dataset2.EOF = False
weight = Dataset2!val
If IsNull(Dataset2!not_ava_time) = False Then
Text4.Text = Dataset2!not_ava_time
Else
Text4.Text = ""
End If

If IsNull(Dataset2!parallel) = True Then
Text9.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room)
Else
Text9.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room) + " " + " " + "" + "" + RTrim(Dataset2!parallel)
End If
.MoveNext
Wend
End If
End With


sql = "select * from session where day= 'Friday' "
Set Dataset2 = mddata(sql)
With Dataset2
i = 1
If .RecordCount > 0 Then
While Dataset2.EOF = False
weight = Dataset2!val
If IsNull(Dataset2!not_ava_time) = False Then
Text4.Text = Dataset2!not_ava_time
Else
Text4.Text = ""
End If

If IsNull(Dataset2!parallel) = True Then
Text10.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room)
Else
Text10.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room) + " " + " " + "" + "" + RTrim(Dataset2!parallel)
End If
.MoveNext
Wend
End If
End With


sql = "select * from session where day= 'Saturday' "
Set Dataset2 = mddata(sql)
With Dataset2
i = 1
If .RecordCount > 0 Then
While Dataset2.EOF = False
weight = Dataset2!val
If IsNull(Dataset2!not_ava_time) = False Then
Text4.Text = Dataset2!not_ava_time
Else
Text4.Text = ""
End If

If IsNull(Dataset2!parallel) = True Then
Text11.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room)
Else
Text11.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room) + " " + " " + "" + "" + RTrim(Dataset2!parallel)
End If
.MoveNext
Wend
End If
End With


sql = "select * from session where day= 'Sunday' "
Set Dataset2 = mddata(sql)
With Dataset2
i = 1
If .RecordCount > 0 Then
While Dataset2.EOF = False
weight = Dataset2!val
If IsNull(Dataset2!not_ava_time) = False Then
Text4.Text = Dataset2!not_ava_time
Else
Text4.Text = ""
End If

If IsNull(Dataset2!parallel) = True Then
Text12.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room)
Else
Text12.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!lecturer2) + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room) + " " + " " + "" + "" + RTrim(Dataset2!parallel)
End If
.MoveNext
Wend
End If
End With



End Sub

Private Sub Command18_Click()
Dim AppXls As Object
Dim ObjWb As Object
Dim objws As Object
Dim i As Integer
Dim range As Object
'Dim excel_app As Excel.Application

Set AppXls = CreateObject("Excel.Application")
Set ObjWb = AppXls.Workbooks.Add

Set objws = ObjWb.Worksheets.Add
objws.Name = "WD_R_15"

With objws.range("A1:E1").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 16
End With

With objws.range("A4:H4").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 14
End With

With objws.range("A2:B2").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 14
End With


With objws.range("A2:B2").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 14
End With
objws.cells(2, 1) = "Full Time Table"

'objws.cells(2, 1) = "Filter By :-"
objws.cells(1, 1) = "WD_R_15"
'objws.cells(1, 1) = "Your_Group_ID"
objws.cells(4, 1) = "Time"
objws.cells(4, 2) = "Monday"
objws.cells(4, 3) = "Tuesday"
objws.cells(4, 4) = "Wednesday"
objws.cells(4, 5) = "Thursday"
objws.cells(4, 6) = "Friday"
objws.cells(4, 7) = "Saturday"
objws.cells(4, 8) = "Sunday"



sql = "select [Time],[Monday] ,[Tuesday],[Wednesday],[Thursday],[Friday],[Saturday],[Sunday] from Time_Table  "
Set Dataset2 = mddata(sql)
With Dataset2
If .RecordCount > 0 Then
i = 5
Dataset2.MoveFirst
While Dataset2.EOF = False
objws.cells(i, 1) = Dataset2!Time
objws.cells(i, 2) = Dataset2!monday
objws.cells(i, 3) = Dataset2!tuesday
objws.cells(i, 4) = Dataset2!Wednesday
objws.cells(i, 5) = Dataset2!thursday
objws.cells(i, 6) = Dataset2!friday
objws.cells(i, 7) = Dataset2!saturday
objws.cells(i, 8) = Dataset2!sunday
i = i + 1

.MoveNext
Wend
End If
End With


With objws.range("A1:E1").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 16
End With

With objws.range("A4:H4").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 14
End With

With objws.range("B4:H4")
.columnwidth = 25
End With

With objws.range("A4")
.columnwidth = 15
End With

With objws.range("A9:H9")
.Interior.ColorIndex = 35

End With
With objws.range("B9:H9")
''MergeCells = True
'''.HorizontalAlignment = xlCenterAcrossSelection
'''.Alignment = xlCenter
'''.AutoFit
''.Merge
.cells = "Interval"
.WrapText = True
.Orientation = 0
.horizontalAlignment = 3
.VerticalAlignment = 1
MergeCells = True
End With

With objws.range("A5:H5")
.WrapText = True
'
'MergeCells = True
End With

With objws.range("A6:H6")
.WrapText = True
End With

With objws.range("A7:H7")
.WrapText = True
End With

With objws.range("A8:H8")
.WrapText = True
End With

With objws.range("A9:H89")
.WrapText = True
End With

With objws.range("A10:H10")
.WrapText = True
End With

With objws.range("A11:H11")
.WrapText = True
End With

With objws.range("A12:H12")
.WrapText = True
End With

With objws.range("A13:H13")
.WrapText = True
End With

With objws.range("A4:H4").Borders
.LineStyle = xlDouble
'.WrapText = True
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A5:H5").Borders
.LineStyle = xlDouble
'.WrapText = True
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A6:H6").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A7:H7").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A8:H8").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A9:H9").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A10:H10").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A11:H11").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A12:H12").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A13:H13").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With


Set objws = Nothing
Set ObjWb = Nothing
AppXls.Visible = True
Set AppXls = Nothing

End Sub

Private Sub Command19_Click()
Dim AppXls As Object
Dim ObjWb As Object
Dim objws As Object
Dim i As Integer
Dim range As Object
'Dim excel_app As Excel.Application

Set AppXls = CreateObject("Excel.Application")
Set ObjWb = AppXls.Workbooks.Add

Set objws = ObjWb.Worksheets.Add
objws.Name = "WD_R_15"

With objws.range("A1:E1").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 16
End With

With objws.range("A4:H4").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 14
End With

With objws.range("A2:B2").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 14
End With
objws.cells(2, 1) = "Filter By :-"
objws.cells(2, 2) = "" + Combo1.Text
objws.cells(1, 1) = "WD_R_15"
'objws.cells(1, 1) = "Your_Group_ID"
objws.cells(4, 1) = "Time"
objws.cells(4, 2) = "Monday"
objws.cells(4, 3) = "Tuesday"
objws.cells(4, 4) = "Wednesday"
objws.cells(4, 5) = "Thursday"
objws.cells(4, 6) = "Friday"
objws.cells(4, 7) = "Saturday"
objws.cells(4, 8) = "Sunday"



sql = "select [Time],[Monday] ,[Tuesday],[Wednesday],[Thursday],[Friday],[Saturday],[Sunday] from Time_Tablecopy  "
Set Dataset2 = mddata(sql)
With Dataset2
If .RecordCount > 0 Then
i = 5
Dataset2.MoveFirst
While Dataset2.EOF = False
objws.cells(i, 1) = Dataset2!Time
objws.cells(i, 2) = Dataset2!monday
objws.cells(i, 3) = Dataset2!tuesday
objws.cells(i, 4) = Dataset2!Wednesday
objws.cells(i, 5) = Dataset2!thursday
objws.cells(i, 6) = Dataset2!friday
objws.cells(i, 7) = Dataset2!saturday
objws.cells(i, 8) = Dataset2!sunday
i = i + 1

.MoveNext
Wend
End If
End With


With objws.range("A1:E1").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 16
End With

With objws.range("A4:H4").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 14
End With

With objws.range("B4:H4")
.columnwidth = 25
End With

With objws.range("A4")
.columnwidth = 15
End With

With objws.range("A9:H9")
.Interior.ColorIndex = 35

End With
With objws.range("B9:H9")
''MergeCells = True
'''.HorizontalAlignment = xlCenterAcrossSelection
'''.Alignment = xlCenter
'''.AutoFit
''.Merge
.cells = "Interval"
.WrapText = True
.Orientation = 0
.horizontalAlignment = 3
.VerticalAlignment = 1
MergeCells = True
End With

With objws.range("A5:H5")
.WrapText = True
'
'MergeCells = True
End With

With objws.range("A6:H6")
.WrapText = True
End With

With objws.range("A7:H7")
.WrapText = True
End With

With objws.range("A8:H8")
.WrapText = True
End With

With objws.range("A9:H89")
.WrapText = True
End With

With objws.range("A10:H10")
.WrapText = True
End With

With objws.range("A11:H11")
.WrapText = True
End With

With objws.range("A12:H12")
.WrapText = True
End With

With objws.range("A13:H13")
.WrapText = True
End With

With objws.range("A4:H4").Borders
.LineStyle = xlDouble
'.WrapText = True
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A5:H5").Borders
.LineStyle = xlDouble
'.WrapText = True
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A6:H6").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A7:H7").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A8:H8").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A9:H9").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A10:H10").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A11:H11").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A12:H12").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A13:H13").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With


Set objws = Nothing
Set ObjWb = Nothing
AppXls.Visible = True
Set AppXls = Nothing
End Sub


Private Sub Command2_Click()

'
'If i < 9 Then
'
'For y = 1 To 7
'
'If monday = 1 And y = 1 Then
'sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
'Set dataset3 = mddata(sql)
'dataset3.Update
'dataset3!monday = Text1.Text
'dataset3.Update
'Next y
'Exit Sub
'End If
'
'
'If tuesday = 1 And y = 2 Then
'sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
'Set dataset3 = mddata(sql)
'dataset3.Update
'dataset3!tuesday = Text1.Text
'dataset3.Update
'Next y
'Exit Sub
'End If
'
'
'If wendsday = 1 And y = 3 Then
'sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
'Set dataset3 = mddata(sql)
'dataset3.Update
'dataset3!Wednesday = Text1.Text
'dataset3.Update
'Next y
'Exit Sub
'End If
'
'If thursday = 1 And y = 4 Then
'sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
'Set dataset3 = mddata(sql)
'dataset3.Update
'dataset3!thursday = Text1.Text
'dataset3.Update
'Next y
'Exit Sub
'End If
'
'
'
'If friday = 1 And y = 5 Then
'sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
'Set dataset3 = mddata(sql)
'dataset3.Update
'dataset3!friday = Text1.Text
'dataset3.Update
'Next y
'Exit Sub
'End If
'
'
'
'If saturday = 1 And y = 6 Then
'sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
'Set dataset3 = mddata(sql)
'dataset3.Update
'dataset3!saturday = Text1.Text
'dataset3.Update
'Next y
'Exit Sub
'End If
'
'If sunday = 1 And y = 7 Then
'sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
'Set dataset3 = mddata(sql)
'dataset3.Update
'dataset3!sunday = Text1.Text
'dataset3.Update
'Next y
'Exit Sub
'End If
'
'
'
'Next y
'
'
'
'
'y = y + 1
'i = i + 1
'
'
'End If


End Sub

Private Sub Command20_Click()
Dim AppXls As Object
Dim ObjWb As Object
Dim objws As Object
Dim i As Integer
Dim range As Object
'Dim excel_app As Excel.Application

Set AppXls = CreateObject("Excel.Application")
Set ObjWb = AppXls.Workbooks.Add

Set objws = ObjWb.Worksheets.Add
'objws.Name = "WD_R_15"

With objws.range("A1:E1").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 16
End With

With objws.range("A4:H4").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 14
End With

With objws.range("A2:B2").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 14
End With
objws.cells(2, 1) = "Filter By :-"
objws.cells(2, 2) = "" + Combo2.Text
objws.cells(1, 1) = "WD_R_15"
'objws.cells(1, 1) = "Your_Group_ID"
objws.cells(4, 1) = "Time"
objws.cells(4, 2) = "Monday"
objws.cells(4, 3) = "Tuesday"
objws.cells(4, 4) = "Wednesday"
objws.cells(4, 5) = "Thursday"
objws.cells(4, 6) = "Friday"
objws.cells(4, 7) = "Saturday"
objws.cells(4, 8) = "Sunday"



sql = "select [Time],[Monday] ,[Tuesday],[Wednesday],[Thursday],[Friday],[Saturday],[Sunday] from Time_Tablecopy  "
Set Dataset2 = mddata(sql)
With Dataset2
If .RecordCount > 0 Then
i = 5
Dataset2.MoveFirst
While Dataset2.EOF = False
objws.cells(i, 1) = Dataset2!Time
objws.cells(i, 2) = Dataset2!monday
objws.cells(i, 3) = Dataset2!tuesday
objws.cells(i, 4) = Dataset2!Wednesday
objws.cells(i, 5) = Dataset2!thursday
objws.cells(i, 6) = Dataset2!friday
objws.cells(i, 7) = Dataset2!saturday
objws.cells(i, 8) = Dataset2!sunday
i = i + 1

.MoveNext
Wend
End If
End With


With objws.range("A1:E1").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 16
End With

With objws.range("A4:H4").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 14
End With

With objws.range("B4:H4")
.columnwidth = 25
End With

With objws.range("A4")
.columnwidth = 15
End With

With objws.range("A9:H9")
.Interior.ColorIndex = 35

End With
With objws.range("B9:H9")
''MergeCells = True
'''.HorizontalAlignment = xlCenterAcrossSelection
'''.Alignment = xlCenter
'''.AutoFit
''.Merge
.cells = "Interval"
.WrapText = True
.Orientation = 0
.horizontalAlignment = 3
.VerticalAlignment = 1
MergeCells = True
End With

With objws.range("A5:H5")
.WrapText = True
'
'MergeCells = True
End With

With objws.range("A6:H6")
.WrapText = True
End With

With objws.range("A7:H7")
.WrapText = True
End With

With objws.range("A8:H8")
.WrapText = True
End With

With objws.range("A9:H89")
.WrapText = True
End With

With objws.range("A10:H10")
.WrapText = True
End With

With objws.range("A11:H11")
.WrapText = True
End With

With objws.range("A12:H12")
.WrapText = True
End With

With objws.range("A13:H13")
.WrapText = True
End With

With objws.range("A4:H4").Borders
.LineStyle = xlDouble
'.WrapText = True
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A5:H5").Borders
.LineStyle = xlDouble
'.WrapText = True
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A6:H6").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A7:H7").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A8:H8").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A9:H9").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A10:H10").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A11:H11").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A12:H12").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A13:H13").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With


Set objws = Nothing
Set ObjWb = Nothing
AppXls.Visible = True
Set AppXls = Nothing
End Sub

Private Sub Command21_Click()
Dim AppXls As Object
Dim ObjWb As Object
Dim objws As Object
Dim i As Integer
Dim range As Object
'Dim excel_app As Excel.Application

Set AppXls = CreateObject("Excel.Application")
Set ObjWb = AppXls.Workbooks.Add

Set objws = ObjWb.Worksheets.Add
objws.Name = "WD_R_15"

With objws.range("A1:E1").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 16
End With

With objws.range("A4:H4").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 14
End With

With objws.range("A2:B2").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 14
End With
objws.cells(2, 1) = "Filter By :-"
objws.cells(2, 2) = "" + Combo3.Text
objws.cells(1, 1) = "WD_R_15"
'objws.cells(1, 1) = "Your_Group_ID"
objws.cells(4, 1) = "Time"
objws.cells(4, 2) = "Monday"
objws.cells(4, 3) = "Tuesday"
objws.cells(4, 4) = "Wednesday"
objws.cells(4, 5) = "Thursday"
objws.cells(4, 6) = "Friday"
objws.cells(4, 7) = "Saturday"
objws.cells(4, 8) = "Sunday"



sql = "select [Time],[Monday] ,[Tuesday],[Wednesday],[Thursday],[Friday],[Saturday],[Sunday] from Time_Tablecopy  "
Set Dataset2 = mddata(sql)
With Dataset2
If .RecordCount > 0 Then
i = 5
Dataset2.MoveFirst
While Dataset2.EOF = False
objws.cells(i, 1) = Dataset2!Time
objws.cells(i, 2) = Dataset2!monday
objws.cells(i, 3) = Dataset2!tuesday
objws.cells(i, 4) = Dataset2!Wednesday
objws.cells(i, 5) = Dataset2!thursday
objws.cells(i, 6) = Dataset2!friday
objws.cells(i, 7) = Dataset2!saturday
objws.cells(i, 8) = Dataset2!sunday
i = i + 1

.MoveNext
Wend
End If
End With


With objws.range("A1:E1").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 16
End With

With objws.range("A4:H4").Font
.Name = "Times New Roman"
.fontstyle = "Bold"
.Size = 14
End With

With objws.range("B4:H4")
.columnwidth = 25
End With

With objws.range("A4")
.columnwidth = 15
End With

With objws.range("A9:H9")
.Interior.ColorIndex = 35

End With
With objws.range("B9:H9")
''MergeCells = True
'''.HorizontalAlignment = xlCenterAcrossSelection
'''.Alignment = xlCenter
'''.AutoFit
''.Merge
.cells = "Interval"
.WrapText = True
.Orientation = 0
.horizontalAlignment = 3
.VerticalAlignment = 1
MergeCells = True
End With

With objws.range("A5:H5")
.WrapText = True
'
'MergeCells = True
End With

With objws.range("A6:H6")
.WrapText = True
End With

With objws.range("A7:H7")
.WrapText = True
End With

With objws.range("A8:H8")
.WrapText = True
End With

With objws.range("A9:H89")
.WrapText = True
End With

With objws.range("A10:H10")
.WrapText = True
End With

With objws.range("A11:H11")
.WrapText = True
End With

With objws.range("A12:H12")
.WrapText = True
End With

With objws.range("A13:H13")
.WrapText = True
End With

With objws.range("A4:H4").Borders
.LineStyle = xlDouble
'.WrapText = True
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A5:H5").Borders
.LineStyle = xlDouble
'.WrapText = True
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A6:H6").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A7:H7").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A8:H8").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A9:H9").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A10:H10").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A11:H11").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A12:H12").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With

With objws.range("A13:H13").Borders
.LineStyle = xlDouble
'.weight = xlThick
.Color = vbBalck
End With


Set objws = Nothing
Set ObjWb = Nothing
AppXls.Visible = True
Set AppXls = Nothing
End Sub


Private Sub Command3_Click()
x = 1
i = 1
y = 1

sql = "select * from session where flag = 1 "
Set Dataset2 = mddata(sql)

If Dataset2.RecordCount > 0 Then
Dataset2.MoveFirst
While Dataset2.EOF = False

val = Dataset2!val

If IsNull(Dataset2!not_ava_time) = False Then
Text4.Text = Dataset2!not_ava_time
Else
Text4.Text = ""
End If

If IsNull(Dataset2!parallel) = True Then
'Text1.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room)
Text5.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room)

Else
'Text1.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room) + " " + " " + "" + "" + RTrim(Dataset2!parallel)
Text5.Text = RTrim(Dataset2!lecturer) + "." + RTrim(Dataset2!subject) + "." + RTrim(Dataset2!Tag) + "." + RTrim(Dataset2!Group) + "." + RTrim(Dataset2!room) + " " + " " + "" + "" + RTrim(Dataset2!parallel)

End If




Dataset2.MoveNext

Wend
MsgBox "Done", vbInformation
Unload Me
Me.Show
End If


End Sub


Private Sub Command5_Click()
sql = "truncate table Time_Tablecopy "
Set dataset = mddata(sql)


sql = "insert into Time_Tablecopy SELECT *FROM Time_Table"
Set dataset = mddata(sql)

End Sub

Private Sub Command6_Click()
Unload Me
Home.Show
End Sub


Private Sub Command7_Click()
Command5_Click

If Not Combo1.Text = "" Then

sql = "select * from Time_Tablecopy "
Set Dataset2 = mddata(sql)

If Dataset2.RecordCount > 0 Then
With Dataset2
Dataset2.MoveFirst
While Dataset2.EOF = False

sql = "truncate table filtertag "
Set dataset3 = mddata(sql)

sql = "select * from filtertag "
Set dataset4 = mddata(sql)
dataset4.AddNew
dataset4!combo = Combo1.Text
dataset4!monday = !monday
dataset4!tuesday = !tuesday
dataset4!Wednesday = !Wednesday
dataset4!thursday = !thursday
dataset4!friday = !friday
dataset4!saturday = !saturday
dataset4!sunday = !sunday

dataset4.Update



id = !id



sql = "select * from filtertag where Monday like'%" + Combo1.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then

sql = "update Time_Tablecopy SET[Monday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)

End If



sql = "select * from filtertag where Tuesday like'%" + Combo1.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then

sql = "update Time_Tablecopy SET[tuesday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)

End If

sql = "select * from filtertag where Wednesday like'%" + Combo1.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Wednesday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)

End If


sql = "select * from filtertag where Thursday like'%" + Combo1.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Thursday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)
End If


sql = "select * from filtertag where Friday like'%" + Combo1.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Friday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)
End If

sql = "select * from filtertag where Saturday like'%" + Combo1.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Saturday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)
Else
End If


sql = "select * from filtertag where Sunday like'%" + Combo1.Text + "%' "
Set dataset5 = mddata(sql)
If dataset5.RecordCount = 0 Then
sql = "update Time_Tablecopy SET[Sunday]='" + Text3.Text + "' where id ='" + Str(id) + "'"
Set dataset7 = mddata(sql)
Else
End If




Dataset2.MoveNext
Wend

'MsgBox "Done", vbInformation
End With

End If
sql = "select * from Time_Tablecopy "
Set Dataset2 = mddata(sql)
Set MSHFlexGrid1.DataSource = Dataset2




Else

MsgBox "Select Lecturer", vbInformation
Combo1.SetFocus
Exit Sub

End If

End Sub

Private Sub Form_Load()
sql = "select * from working_days_hours"
Set dataset = mddata(sql)
With dataset
If .RecordCount > 0 Then



 monday = !monday
 tuesday = !tuesday
 wendsday = !wendesday
 thursday = !thursday
 friday = !friday
 saturday = !saturday
 sunday = !sunday
 
 End If
 End With
 
 
 sql = "select * from Time_Table"
Set Dataset2 = mddata(sql)

Set MSHFlexGrid1.DataSource = Dataset2

MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 0
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.ColWidth(4) = 1600
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1600
MSHFlexGrid1.ColWidth(8) = 1600
MSHFlexGrid1.ColWidth(9) = 1600
MSHFlexGrid1.ColWidth(10) = 0
'MSHFlexGrid1.RowHeight(1) = 600


Set MSHFlexGrid2.DataSource = Dataset2

MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(1) = 0
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(3) = 1600
MSHFlexGrid2.ColWidth(4) = 1600
MSHFlexGrid2.ColWidth(5) = 1600
MSHFlexGrid2.ColWidth(6) = 1600
MSHFlexGrid2.ColWidth(7) = 1600
MSHFlexGrid2.ColWidth(8) = 1600
MSHFlexGrid2.ColWidth(9) = 1600
MSHFlexGrid2.ColWidth(10) = 0
'MSHFlexGrid2.RowHeight(1) = 600

Set MSHFlexGrid3.DataSource = Dataset2

MSHFlexGrid3.ColWidth(0) = 0
MSHFlexGrid3.ColWidth(1) = 0
MSHFlexGrid3.ColWidth(2) = 1200
MSHFlexGrid3.ColWidth(3) = 1600
MSHFlexGrid3.ColWidth(4) = 1600
MSHFlexGrid3.ColWidth(5) = 1600
MSHFlexGrid3.ColWidth(6) = 1600
MSHFlexGrid3.ColWidth(7) = 1600
MSHFlexGrid3.ColWidth(8) = 1600
MSHFlexGrid3.ColWidth(9) = 1600
MSHFlexGrid3.ColWidth(10) = 0
'MSHFlexGrid3.RowHeight(1) = 600

Combo1.Clear
sql = "select * from Lecturer"
Set dataset3 = mddata(sql)

dataset3.MoveFirst
While dataset3.EOF = False
Combo1.AddItem dataset3!lec_name
dataset3.MoveNext
Wend

Combo4.Clear
sql = "select * from Lecturer"
Set dataset3 = mddata(sql)

dataset3.MoveFirst
While dataset3.EOF = False
Combo4.AddItem dataset3!lec_name
dataset3.MoveNext
Wend


Combo2.Clear
sql = "select * from Student_group"
Set dataset4 = mddata(sql)

dataset4.MoveFirst
While dataset4.EOF = False
Combo2.AddItem dataset4!group_id
dataset4.MoveNext
Wend

Combo3.Clear
sql = "select * from Loacation"
Set dataset5 = mddata(sql)

dataset5.MoveFirst
While dataset5.EOF = False
Combo3.AddItem dataset5!Room_Name
dataset5.MoveNext
Wend

End Sub


Private Sub Print_Click()



End Sub

Private Sub Text1_Change()
Dim k As Integer
If Not Text4.Text = "" Then

If weight > 1 Then
k = weight
Else
k = 2
End If
Else
k = 100
End If

For y = 1 To weight

If Not k = 1 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)

If Not Text4.Text = dataset3!Time Then
dataset3.Update
dataset3!monday = Text1.Text
dataset3.Update
i = i + 1
k = k - 1
Else
i = i + 1
End If
End If
Next y
Exit Sub
End Sub

Private Sub Text10_Change()
Dim k As Integer
If Not Text4.Text = "" Then

If weight > 1 Then
k = weight
Else
k = 2
End If
Else
k = 100
End If

For y = 1 To weight
If Not k = 1 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)

If Not Text4.Text = dataset3!Time Then
dataset3.Update
dataset3!friday = Text10.Text
dataset3.Update
i = i + 1
k = k - 1
Else
i = i + 1
End If
End If
Next y
Exit Sub

End Sub

Private Sub Text11_Change()
Dim k As Integer
If Not Text4.Text = "" Then

If weight > 1 Then
k = weight
Else
k = 2
End If
Else
k = 100
End If

For y = 1 To weight
If Not k = 1 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)

If Not Text4.Text = dataset3!Time Then
dataset3.Update
dataset3!saturday = Text11.Text
dataset3.Update
i = i + 1
k = k - 1
Else
i = i + 1
End If
End If
Next y
Exit Sub
End Sub

Private Sub Text12_Change()
Dim k As Integer
If Not Text4.Text = "" Then

If weight > 1 Then
k = weight
Else
k = 2
End If
Else
k = 100
End If

For y = 1 To weight
If Not k = 1 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)

If Not Text4.Text = dataset3!Time Then
dataset3.Update
dataset3!sunday = Text12.Text
dataset3.Update
i = i + 1
k = k - 1
Else
i = i + 1
End If
End If
Next y
Exit Sub

End Sub


Private Sub Text2_Change()
y = 1
i = i + 1
Text1_Change
End Sub


Private Sub Text3_Change()



If Text3.Text = Combo1.Text Then

sql = "update Time_Tablecopy SET[Monday]='" + Text3.Text + "' where id ='" + id + "'"
Set Dataset2 = mddata(sql)

Else

'sql = "update Time_Tablecopy SET[Monday]=X where id ='" + id + "'"
'Set Dataset2 = mddata(sql)

End If


End Sub

Private Sub Text5_Change()


For y = 1 To val

If monday = 1 And x > 0 And x < 9 Then
If i < 9 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)

If Not Text4.Text = dataset3!Time Then
dataset3.Update
dataset3!monday = Text1.Text
dataset3.Update
i = i + 1
x = x + 1

Else
i = i + 1
y = y - 1
End If
Else

End If
End If

Next y
Exit Sub


End Sub


Private Sub Text6_Change()

If monday = 1 Then
If mo + k < 10 Then
sql = "select * from session where id='" + Str(Text6.Text) + "'  "
Set dataset3 = mddata(sql)
dataset3.Update
dataset3!Day = "Monday"
dataset3.Update
mo = mo + k
Exit Sub
End If
End If




If tuesday = 1 Then
If tu + k < 10 Then
sql = "select * from session where id='" + Str(Text6.Text) + "'  "
Set dataset3 = mddata(sql)
dataset3.Update
dataset3!Day = "Tuesday"
dataset3.Update
tu = tu + k
Exit Sub
End If
End If





If wendsday = 1 Then
If we + k < 10 Then
sql = "select * from session where id='" + Str(Text6.Text) + "'  "
Set dataset3 = mddata(sql)
dataset3.Update
dataset3!Day = "Wednesday"
dataset3.Update
we = we + k
Exit Sub
End If
End If




If thursday = 1 Then
If th + k < 10 Then
sql = "select * from session where id='" + Str(Text6.Text) + "'  "
Set dataset3 = mddata(sql)
dataset3.Update
dataset3!Day = "Thursday"
dataset3.Update
th = th + k
Exit Sub
End If
End If



If friday = 1 Then
If fr + k < 10 Then
sql = "select * from session where id='" + Str(Text6.Text) + "'  "
Set dataset3 = mddata(sql)

dataset3.Update
dataset3!Day = "friday"

dataset3.Update
fr = fr + k
Exit Sub
End If
End If

If saturday = 1 Then

If sa + k < 10 Then
sql = "select * from session where id='" + Str(Text6.Text) + "'  "
Set dataset3 = mddata(sql)

dataset3.Update
dataset3!Day = "saturday"
dataset3.Update
sa = sa + k
Exit Sub
End If
End If



If sunday = 1 Then
If su + k < 10 Then
sql = "select * from session where id='" + Str(Text6.Text) + "'  "
Set dataset3 = mddata(sql)

dataset3.Update
dataset3!Day = "sunday"
dataset3.Update
su = su + k
Exit Sub
End If
End If


End Sub


Private Sub Text7_Change()

Dim k As Integer
If Not Text4.Text = "" Then

If weight > 1 Then
k = weight
Else
k = 2
End If
Else
k = 100
End If

For y = 1 To weight
If Not k = 1 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)

If Not Text4.Text = dataset3!Time Then
dataset3.Update
dataset3!tuesday = Text7.Text
dataset3.Update
i = i + 1
k = k - 1
Else
i = i + 1
End If
End If
Next y
Exit Sub
End Sub


Private Sub Text8_Change()
Dim k As Integer
If Not Text4.Text = "" Then

If weight > 1 Then
k = weight
Else
k = 2
End If
Else
k = 100
End If


For y = 1 To weight
If Not k = 1 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)

If Not Text4.Text = dataset3!Time Then
dataset3.Update
dataset3!Wednesday = Text8.Text
dataset3.Update
i = i + 1
k = k - 1
Else
i = i + 1
End If
End If
Next y
Exit Sub

End Sub


Private Sub Text9_Change()

Dim k As Integer
If Not Text4.Text = "" Then

If weight > 1 Then
k = weight
Else
k = 2
End If
Else
k = 100
End If

For y = 1 To weight
If Not k = 1 Then
sql = "select * from Time_Table where tagid='" + Str(i) + "'  "
Set dataset3 = mddata(sql)

If Not Text4.Text = dataset3!Time Then
dataset3.Update
dataset3!thursday = Text9.Text
dataset3.Update
i = i + 1
k = k - 1
Else
i = i + 1
End If
End If
Next y
Exit Sub
End Sub


