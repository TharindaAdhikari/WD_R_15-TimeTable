VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Locations"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14040
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   14040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14535
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add Locations"
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
         TabIndex        =   4
         Top             =   240
         Width           =   11775
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   11775
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Form4.frx":0000
         Height          =   495
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5415
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   9551
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Session"
         TabPicture(0)   =   "Form4.frx":1A8E3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Command1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Command5"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Consecutive"
         TabPicture(1)   =   "Form4.frx":1A8FF
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command6"
         Tab(1).Control(1)=   "Command2"
         Tab(1).Control(2)=   "MSHFlexGrid1"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Not Available Time"
         TabPicture(2)   =   "Form4.frx":1A91B
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Command7"
         Tab(2).Control(1)=   "Command3"
         Tab(2).Control(2)=   "Combo2"
         Tab(2).Control(3)=   "Combo1"
         Tab(2).Control(4)=   "Text2"
         Tab(2).Control(5)=   "Text1"
         Tab(2).Control(6)=   "Label152(2)"
         Tab(2).Control(7)=   "Label152(1)"
         Tab(2).Control(8)=   "Label152(0)"
         Tab(2).Control(9)=   "Label152(18)"
         Tab(2).ControlCount=   10
         Begin VB.CommandButton Command7 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Caption         =   "Clear"
            DisabledPicture =   "Form4.frx":1A937
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
            Left            =   -66000
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   4440
            Width           =   1695
         End
         Begin VB.CommandButton Command6 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Caption         =   "Refresh"
            DisabledPicture =   "Form4.frx":3521A
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
            Left            =   -65760
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   4440
            Width           =   1695
         End
         Begin VB.CommandButton Command5 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Caption         =   "Refresh"
            DisabledPicture =   "Form4.frx":4FAFD
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
            Left            =   9360
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   4440
            Width           =   1695
         End
         Begin VB.CommandButton Command3 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            Caption         =   "Add Session"
            DisabledPicture =   "Form4.frx":6A3E0
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
            Left            =   -68280
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   4440
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            Caption         =   "Add Session"
            DisabledPicture =   "Form4.frx":84CC3
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
            Left            =   -67800
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   4440
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            Caption         =   "Add Session"
            DisabledPicture =   "Form4.frx":9F5A6
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
            Left            =   7320
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   4440
            Width           =   1695
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
            ItemData        =   "Form4.frx":B9E89
            Left            =   -72000
            List            =   "Form4.frx":B9EA2
            TabIndex        =   16
            Top             =   2400
            Width           =   2775
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
            ItemData        =   "Form4.frx":B9EE8
            Left            =   -72000
            List            =   "Form4.frx":B9EF8
            TabIndex        =   15
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            Height          =   405
            Left            =   -67320
            TabIndex        =   14
            Top             =   2400
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Height          =   405
            Left            =   -67320
            TabIndex        =   13
            Top             =   1080
            Visible         =   0   'False
            Width           =   2415
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   3255
            Left            =   0
            TabIndex        =   7
            Top             =   840
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   5741
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
            Height          =   3255
            Left            =   -75000
            TabIndex        =   8
            Top             =   840
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   5741
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
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Time"
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
            Left            =   -69360
            TabIndex        =   12
            Top             =   2400
            Width           =   1875
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Time"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   -69000
            TabIndex        =   11
            Top             =   1080
            Width           =   1425
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Day"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   0
            Left            =   -74520
            TabIndex        =   10
            Top             =   2400
            Width           =   2265
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Room"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   18
            Left            =   -74640
            TabIndex        =   9
            Top             =   1080
            Width           =   2175
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
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   840
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
         Left            =   360
         TabIndex        =   5
         Top             =   6960
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
manage_session_room.Show
End Sub

Private Sub Command2_Click()
Unload Me
manage_session_room.Show

End Sub

Private Sub Command3_Click()
MsgBox "Done", vbInformation
End Sub

Private Sub Command4_Click()
Unload Me
Home.Show
End Sub

Private Sub Command7_Click()
Unload Me
Me.Show
End Sub

Private Sub Form_Load()
sql = "select * from session "
Set dataset = mddata(sql)
With dataset

Set MSHFlexGrid1.DataSource = dataset
Set MSHFlexGrid3.DataSource = dataset

MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 500
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1600
MSHFlexGrid1.ColWidth(5) = 800
MSHFlexGrid1.ColWidth(6) = 800
MSHFlexGrid1.ColWidth(7) = 800
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 0



MSHFlexGrid3.ColWidth(0) = 0
MSHFlexGrid3.ColWidth(1) = 500
MSHFlexGrid3.ColWidth(2) = 2000
MSHFlexGrid3.ColWidth(3) = 1000
MSHFlexGrid3.ColWidth(4) = 1600
MSHFlexGrid3.ColWidth(5) = 800
MSHFlexGrid3.ColWidth(6) = 800
MSHFlexGrid3.ColWidth(7) = 800
MSHFlexGrid3.ColWidth(8) = 1200
MSHFlexGrid3.ColWidth(9) = 1200
MSHFlexGrid3.ColWidth(10) = 0
End With


sql = "select * from Loacation"
Set dataset = mddata(sql)
Combo1.Clear
With dataset
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo1.AddItem dataset!Room_Name
.MoveNext
Wend

End If
End With


End Sub
