VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form Statictis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statistics"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12600
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
      Begin VB.Label Label1 
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
         TabIndex        =   6
         Top             =   7320
         Width           =   1455
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14535
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Statistics"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   480
         TabIndex        =   26
         Top             =   480
         Width           =   11775
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Statistics"
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
         Left            =   1080
         TabIndex        =   4
         Top             =   1440
         Width           =   12255
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   2400
      TabIndex        =   0
      Top             =   1320
      Width           =   15855
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Statictis.frx":0000
         Height          =   495
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
         Height          =   2055
         Left            =   5760
         TabIndex        =   18
         Top             =   5280
         Width           =   3495
         lastProp        =   500
         _cx             =   6165
         _cy             =   3625
         DisplayGroupTree=   -1  'True
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   -1  'True
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   -1  'True
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
         LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   5760
         TabIndex        =   17
         Top             =   960
         Width           =   4335
         Begin VB.TextBox Text3 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2280
            TabIndex        =   25
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2280
            TabIndex        =   24
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2280
            TabIndex        =   23
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Latest Subjects"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   270
            Index           =   8
            Left            =   -15
            TabIndex        =   22
            Top             =   2640
            Width           =   2355
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Latest Group"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   270
            Index           =   7
            Left            =   -150
            TabIndex        =   21
            Top             =   1800
            Width           =   2295
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Latest Lecturer"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   270
            Index           =   5
            Left            =   -15
            TabIndex        =   20
            Top             =   1080
            Width           =   2355
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "What is New?"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   270
            Index           =   4
            Left            =   -480
            TabIndex        =   19
            Top             =   120
            Width           =   3195
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Statictis.frx":1A8E3
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   -600
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   480
         TabIndex        =   16
         Top             =   6600
         Width           =   2535
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   480
         TabIndex        =   15
         Top             =   4560
         Width           =   2535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   480
         TabIndex        =   14
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registered Rooms"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   270
         TabIndex        =   13
         Top             =   6120
         Width           =   3015
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registered Subjects"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   225
         TabIndex        =   12
         Top             =   3960
         Width           =   3105
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registered Student"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   225
         TabIndex        =   11
         Top             =   2160
         Width           =   3105
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   480
         TabIndex        =   10
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registered Lecturer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   18
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2835
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
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
         Left            =   11640
         TabIndex        =   1
         Top             =   6480
         Width           =   1455
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
      Left            =   360
      TabIndex        =   8
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
      Left            =   4080
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "Statictis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Home.Show
End Sub

Private Sub Command4_Click()
Unload Me
Home.Show
End Sub

Private Sub Form_Load()


sql = "select * from lec"
Set dataset = mddata(sql)
With dataset
If .RecordCount > 0 Then
Label3.Caption = dataset!lec
End If
End With

sql = "select * from location"
Set dataset1 = mddata(sql)
With dataset1
If .RecordCount > 0 Then
Label6.Caption = dataset1!Location
End If
End With

sql = "select * from Student_group_count"
Set Dataset2 = mddata(sql)
With Dataset2
If .RecordCount > 0 Then
Label4.Caption = Dataset2!id
End If
End With

sql = "select * from subject_count"
Set dataset3 = mddata(sql)
With dataset3
If .RecordCount > 0 Then
Label5.Caption = dataset3!Sub
End If
End With

sql = "select * from lec_name"
Set dataset1 = mddata(sql)
With dataset1
If .RecordCount > 0 Then
Text1.Text = dataset1!lec_name
End If
End With

sql = "select * from sub_name"
Set Dataset2 = mddata(sql)
With Dataset2
If .RecordCount > 0 Then
Text3.Text = Dataset2!Sub_name
End If
End With

sql = "select * from group_id"
Set dataset3 = mddata(sql)
With dataset3
If .RecordCount > 0 Then
Text2.Text = dataset3!group_id
End If
End With

End Sub

Private Sub ItemEdit1_GotFocus()

End Sub
