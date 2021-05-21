VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form employee_edit 
   BackColor       =   &H00FEE2CF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Edit"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9765
   LinkTopic       =   "employee_edit"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check6 
      BackColor       =   &H00F4DBC8&
      Caption         =   "Steward"
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
      Left            =   -720
      TabIndex        =   93
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FAEDDF&
      Caption         =   "Per Hour"
      Height          =   375
      Left            =   13560
      TabIndex        =   83
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text18 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13080
      TabIndex        =   82
      Top             =   3840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00FAEDDF&
      Caption         =   "E.P.F On Transport"
      Height          =   375
      Left            =   12720
      TabIndex        =   81
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   3240
      Picture         =   "employee_edit.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   3600
      Width           =   1300
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FEE2CF&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   -120
      TabIndex        =   2
      Top             =   600
      Width           =   10215
      Begin TabDlg.SSTab SSTab1 
         Height          =   1095
         Left            =   6720
         TabIndex        =   94
         Top             =   2160
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   1931
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "employee_edit.frx":03A8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Command2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Tab 1"
         TabPicture(1)   =   "employee_edit.frx":03C4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "employee_edit.frx":03E0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   255
            Left            =   480
            TabIndex        =   95
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   87
         Top             =   7320
         Width           =   1815
      End
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   86
         Top             =   7800
         Width           =   1815
      End
      Begin VB.TextBox Text27 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   78
         Top             =   8280
         Width           =   2295
      End
      Begin VB.TextBox Text26 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7200
         TabIndex        =   0
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text25 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FDE4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   74
         Top             =   7680
         Width           =   2295
      End
      Begin VB.TextBox Text23 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   73
         Top             =   5160
         Width           =   975
      End
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   72
         Top             =   5640
         Width           =   975
      End
      Begin VB.TextBox Text24 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   71
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   65
         Top             =   3120
         Width           =   6735
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   64
         Top             =   5160
         Width           =   2295
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   61
         Top             =   5160
         Width           =   2655
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FAEDDF&
         Caption         =   "E.P.F On Basic Salary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   60
         Top             =   9480
         Width           =   2655
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   58
         Top             =   7080
         Width           =   2295
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   56
         Top             =   5160
         Width           =   2295
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Left            =   2760
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   840
         Width           =   3495
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFEEE3&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3735
         Left            =   11160
         TabIndex        =   4
         Top             =   4800
         Visible         =   0   'False
         Width           =   5895
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2520
            TabIndex        =   10
            Top             =   2880
            Width           =   2295
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2520
            TabIndex        =   9
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2520
            TabIndex        =   8
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2520
            TabIndex        =   7
            Top             =   2400
            Width           =   2295
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2520
            TabIndex        =   6
            Top             =   1920
            Width           =   2295
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            Picture         =   "employee_edit.frx":03FC
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Meal :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   16
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Driving L :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   15
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bike :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   14
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "COM :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   13
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cloth :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   12
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Shape Shape1 
            Height          =   3735
            Left            =   840
            Top             =   360
            Width           =   6015
         End
         Begin VB.Label Label17 
            BackColor       =   &H00F7DCC6&
            Caption         =   " Other Benefits"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   11415
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   29
         Top             =   1320
         Width           =   6735
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   28
         Top             =   4680
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FAEDDF&
         Caption         =   " E.P.F & E.T.F Calculation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   27
         Top             =   9480
         Width           =   2655
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   26
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FDE4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   25
         Top             =   4680
         Width           =   2295
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   24
         Top             =   4080
         Width           =   2655
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FAEDDF&
         Caption         =   " CONTRACT BASIS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   23
         Top             =   9480
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   22
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   21
         Top             =   6120
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   20
         Top             =   5640
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   19
         Top             =   3600
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Left            =   2760
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   17
         Top             =   6600
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   11040
         Picture         =   "employee_edit.frx":06C2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker D1 
         Height          =   375
         Left            =   3120
         TabIndex        =   30
         Top             =   9000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   123469826
         CurrentDate     =   40551.2916666667
      End
      Begin MSComCtl2.DTPicker D2 
         Height          =   375
         Left            =   6360
         TabIndex        =   31
         Top             =   9000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   123535362
         CurrentDate     =   40551.7083333333
      End
      Begin MSComCtl2.DTPicker DTShowStartDate 
         Height          =   375
         Left            =   3600
         TabIndex        =   32
         Top             =   4080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   123535363
         CurrentDate     =   39120
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SOFTLOGIC  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   89
         Top             =   7800
         Width           =   1815
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COMPANY  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   88
         Top             =   7320
         Width           =   1815
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LOANS INSTALLMENTS "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   80
         Top             =   6720
         Width           =   3255
      End
      Begin VB.Line Line4 
         X1              =   7200
         X2              =   11400
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ATT: BONUS /PER DAY :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   79
         Top             =   8280
         Width           =   3135
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   77
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   76
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL ALLOW :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   75
         Top             =   7680
         Width           =   2415
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LEAVES "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   70
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CASUAL :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   69
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SICK :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   68
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ANNUAL :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   67
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Line Line3 
         X1              =   7200
         X2              =   7200
         Y1              =   4560
         Y2              =   8760
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Other Names :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   66
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bank :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   63
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Branch :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   62
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   11400
         Y1              =   8760
         Y2              =   8760
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSPORT :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   59
         Top             =   7080
         Width           =   1815
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "B.R.A :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   57
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   11400
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label20 
         Caption         =   "Label20"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "Label19"
         Height          =   375
         Left            =   360
         TabIndex        =   54
         Top             =   2760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   52
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label10 
         Height          =   375
         Left            =   5760
         TabIndex        =   46
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Division :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   45
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Employee New Name :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   44
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N.I.C. NO :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   43
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TP No :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   42
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   41
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Out :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   40
         Top             =   9000
         Width           =   1215
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Shift  In :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   39
         Top             =   9000
         Width           =   1815
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "BASIC SALARY :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   38
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Designation :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   37
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Appointment Date :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   36
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SPECIAL :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   35
         Top             =   6120
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "INCENTIVE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   34
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FAMILY  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   33
         Top             =   6600
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   4800
      Picture         =   "employee_edit.frx":0CFA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1300
   End
   Begin VB.Label Label46 
      BackColor       =   &H00404040&
      Height          =   735
      Left            =   -120
      TabIndex        =   92
      Top             =   4080
      Width           =   10575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE EDIT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   90
      Top             =   50
      Width           =   3855
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSPORT :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -960
      TabIndex        =   85
      Top             =   5160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "OT RATE :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   84
      Top             =   4080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label48 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   -840
      TabIndex        =   50
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label49 
      Height          =   255
      Left            =   2040
      TabIndex        =   49
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9480
      TabIndex        =   48
      Top             =   9840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Account No :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   47
      Top             =   9840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00404040&
      Height          =   735
      Left            =   0
      TabIndex        =   91
      Top             =   -120
      Width           =   10575
   End
End
Attribute VB_Name = "employee_edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Label16.Caption = ""
My = "Division= '" + Combo1.Text + "'"

   Data3.Recordset.FindFirst My
  If Not Data3.Recordset.NoMatch Then
  Label16.Caption = Data3.Recordset!id
  End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

 Combo3.Clear
With Data4.Recordset
If Data4.Recordset.RecordCount > 0 Then
.MoveFirst
End If

While .EOF = False
If !Division = Label16.Caption Then
Combo3.additem !Name
End If
.MoveNext
Wend

End With
Combo3.SetFocus
SendKeys "{f4}"

End If
End Sub

Private Sub Combo2_Click()
Label10.Caption = ""

SQL = "Select * from  Staff_Division where Division= '" + Combo2.Text + "'"
Set dataset = mddata(SQL)

  
  If dataset.RecordCount > 0 Then
  Label10.Caption = dataset!id
  End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then


SQL = "Select * from  Employee where division =" + Label10.Caption
Set dataset = mddata(SQL)

 Combo3.Clear
With dataset
If .RecordCount > 0 Then
.MoveFirst
End If

While .EOF = False

Combo3.additem !Name

.MoveNext
Wend


End With
Combo3.SetFocus
SendKeys "{f4}"


End If
End Sub

Private Sub Combo3_Click()




  Text1.Text = ""
 Text2.Text = ""
 Label3.Caption = ""
Label10.Caption = ""
 Text3.Text = ""
 Text7.Text = ""
 Check1.Value = 0
 Check2.Value = 0
 Check3.Value = 0
 Check4.Value = 0
Check6.Value = 0

Text12.Text = ""
Text9.Text = ""

Text13.Text = ""
Text15.Text = ""
Text17.Text = ""


Text5.Text = ""
Text4.Text = ""
Text6.Text = ""
Text10.Text = ""
Text8.Text = ""
Text14.Text = ""
Text11.Text = ""
Text18.Text = ""
Text16.Text = ""
Text19.Text = ""
Text21.Text = ""
Text20.Text = ""
Check5.Value = 0

Text25.Text = ""

Text27.Text = ""
Text28.Text = ""
Text29.Text = ""






 On Error Resume Next

SQL = "Select * from  Employee where Name= '" + Combo3.Text + "'"
Set dataset = mddata(SQL)

With dataset
   
    

  If .RecordCount > 0 Then
  
  
  Text1.Text = !Name
 Text2.Text = !tp
 Label3.Caption = !acc_no
Label10.Caption = !Division
 Text3.Text = !NIC
 Text7.Text = !bank_acc
 Check1.Value = !extra
 Check2.Value = !cntrct
 Check3.Value = !Rate
 Check4.Value = !epf_basic

Text12.Text = !Salary
Text9.Text = !Link_Code
D1.Value = !in_t
D2.Value = !Out_t
Text13.Text = !Meal
Text15.Text = !Transport
Text17.Text = !desig
DTShowStartDate.Value = !appoinment

Text5.Text = !Living
Text4.Text = !Spec
Text6.Text = !Allow
Text10.Text = !Bike
Text8.Text = !Driving
Text14.Text = !Cloth
Text11.Text = !com
Text18.Text = !OT
Text16.Text = !BRA
Text19.Text = !bbranch
Check5.Value = !epf_trans
Check6.Value = Val(!othr_nm)
Text20.Text = !bank
Text23.Text = !annual
Text22.Text = !sick
Text24.Text = !casual
If IsNull(!Tot_allowance) = False Then
Text25.Text = !Tot_allowance
End If

If IsNull(!Tot_allowance) = False Then
Text25.Text = !Tot_allowance
End If

If IsNull(!installment) = False Then
Text29.Text = !installment
End If




If IsNull(!soft_install) = False Then
Text28.Text = !soft_install
End If


SQL = "Select * from  Staff_Division where id=" + Label10.Caption

Set dataset = mddata(SQL)

With dataset

If .RecordCount > 0 Then
Combo2.Text = !Division
End If
End With

End If

End With

Text1.SetFocus

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.SetFocus
End If
End Sub

Private Sub Command1_Click()
Frame1.Visible = True
Text15.SetFocus

End Sub

Private Sub Command2_Click()
'generate account code

If Not Val(Text9.Text) = 0 Then

If Not Val(Label10.Caption) = 0 Then
If Not Text1.Text = "" Then


With Data4.Recordset
.FindFirst "Link_Code =" + Text9.Text
If .NoMatch = False Then
MsgBox "Existing Link Code!", vbExclamation, "Error !"
Text9.SetFocus
Exit Sub
End If

End With

MyCriteria2 = "PName= '" + Text1.Text + "'"

 Data2.Recordset.FindFirst MyCriteria2
  If Data2.Recordset.NoMatch Then

With Data1.Recordset
.MoveFirst
.Edit
!PR_Code = Val(!PR_Code) + 1
Label3.Caption = "4-8-9-" + LTrim(str(!PR_Code))
.Update
End With

With Data2.Recordset
.AddNew
!P_Code = Val(Label3.Caption)
!PName = Text1.Text
!B_Code = 4
!T_Code = 8
!H_Code = 9
!L_Code = Label3.Caption
.Update
End With



With Data4.Recordset

.AddNew
!Name = Text1.Text
!tp = Text2.Text
!acc_no = Label3.Caption
!Division = Val(Label10.Caption)
!epf_no = Text3.Text
!bank_acc = Text7.Text
!extra = Check1.Value
!cntrct = Check2.Value
!epf_basic = Check4.Value
!Salary = Val(Text12.Text)
!Link_Code = Val(Text9.Text)
!in_t = TimeValue(D1.Value)
!Out_t = TimeValue(D2.Value)
!Meal = Val(Text13.Text)
!Transport = Val(Text15.Text)
!desig = Text17.Text
!appoinment = DTShowStartDate.Value

!Living = Val(Text5.Text)
!Spec = Val(Text4.Text)
!Allow = Val(Text6.Text)
!Bike = Val(Text10.Text)
!Driving = Val(Text8.Text)
!Cloth = Val(Text14.Text)
!com = Val(Text11.Text)
!BRA = Val(Text16.Text)
!annual = Text23.Text
!sick = Text22.Text
!casual = Text24.Text



.Update
End With
MsgBox "Update Successfully "
Command2.Enabled = False
Exit Sub

Else
MsgBox "Existing Name!", vbExclamation, "Error !"
Text1.SetFocus

End If


Else

MsgBox "Invalid Account Name", vbExclamation, "Error !"
End If

Else

MsgBox "Invalid Division", vbExclamation, "Error !"
Combo2.SetFocus
SendKeys "{f4}"
End If

Else

MsgBox "Invalid Link Code", vbExclamation, "Error !"
Text9.SetFocus
End If
End Sub

Private Sub Command3_Click()
If Not Label3.Caption = "" Then



If Val(Label20.Caption) = 1 Then

SQL = "Select * From Account_Posting where PName='" + Text1.Text + "'"
Set dataset = mddata(SQL)

  If dataset.RecordCount > 0 Then
  
  MsgBox "New Name Already in Use !", vbExclamation, "Error !"
  Text1.SetFocus
  
  Exit Sub
  End If
  End If


SQL = "Select * From Employee where Link_Code =" + Text9.Text
Set dataset = mddata(SQL)


If Val(Label19.Caption) = 1 Then


If dataset.RecordCount > 0 Then
MsgBox "Existing Link Code!", vbExclamation, "Error !"
Text9.SetFocus
Exit Sub
End If

End If


 
SQL = "UPDATE Employee set"

SQL = SQL + "[Name] = '" + Text1.Text
SQL = SQL + "',[tp] = '" + Text2.Text
SQL = SQL + "',[nic] = '" + Text3.Text
SQL = SQL + "',[bank_acc] = '" + Text7.Text
SQL = SQL + "',[extra] = '" + str(Check1.Value)
SQL = SQL + "',[cntrct] = '" + str(Check2.Value)
SQL = SQL + "',[Rate] = '" + str(Check3.Value)
SQL = SQL + "',[epf_basic] = '" + str(Check4.Value)
SQL = SQL + "',[Salary] = '" + Text12.Text
SQL = SQL + "',[Link_Code] = '" + Text9.Text
SQL = SQL + "',[in_t] = '" + Format(D1.Value, "yyyy/mmm/dd") 'Str(D1.Value)
SQL = SQL + "',[Out_t] = '" + Format(D2.Value, "yyyy/mmm/dd") 'Str(D2.Value)
SQL = SQL + "',[Meal] = '" + Text13.Text
SQL = SQL + "',[transport] = '" + Text15.Text
SQL = SQL + "',[desig] = '" + Text17.Text
SQL = SQL + "',[appoinment] = '" + Format(DTShowStartDate.Value, "yyyy/mmm/dd")  'Str(DTShowStartDate.Value)

SQL = SQL + "',[Living] = '" + Text5.Text
SQL = SQL + "',[Spec] = '" + Text4.Text
SQL = SQL + "',[Allow] = '" + Text6.Text
SQL = SQL + "',[Bike] = '" + Text10.Text
SQL = SQL + "',[Driving] = '" + Text8.Text
SQL = SQL + "',[Cloth] = '" + Text14.Text
SQL = SQL + "',[com] = '" + Text11.Text
SQL = SQL + "',[BRA] = '" + Text16.Text
SQL = SQL + "',[OT] = '" + Text18.Text
SQL = SQL + "',[bbranch] = '" + Text19.Text
SQL = SQL + "',[epf_trans] = '" + str(Check5.Value)
SQL = SQL + "',[othr_nm] = '" + str(Check6.Value)
SQL = SQL + "',[bank] = '" + Text20.Text
SQL = SQL + "',[annual] = '" + Text23.Text
SQL = SQL + "',[sick] = '" + Text22.Text
SQL = SQL + "',[casual] = '" + Text24.Text
 SQL = SQL + "',[Tot_allowance] = '" + Text25.Text
 SQL = SQL + "',[att_bonus] = '" + Text27.Text
 
SQL = SQL + "',[soft_install] = '" + Text28.Text
SQL = SQL + "',[installment] = '" + Text29.Text
 
SQL = SQL + "' where Acc_No= '" + Label3.Caption + "'"
Set Dataset2 = mddata(SQL)

  
MsgBox "Successfully Edit!", vbInformation
Command3.Enabled = False
Command4.SetFocus



SQL = "update Account_Posting set [PName]='" + Text1.Text + "' where L_Code= '" + Label3.Caption + "'"
Set dataset = mddata(SQL)


Else
MsgBox "Invalid Employee !", vbCritical
End If
End Sub

Private Sub Command4_Click()
Unload Me
Me.Show
End Sub

Private Sub Command5_Click()
Frame1.Visible = False
  
End Sub

Private Sub D1_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = 13 Then
D2.SetFocus
End If
End Sub

Private Sub D2_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = 13 Then
Check1.SetFocus
End If
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Check2.SetFocus
End If
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Check4.SetFocus
End If
End Sub

Private Sub Check4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3.SetFocus
End If
End Sub


Private Sub Check3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text23.SetFocus
End If
End Sub

Private Sub Text15_Change()
cal
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text22.SetFocus
End If
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text24.SetFocus
End If
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
D1.SetFocus
End If
End Sub

Private Sub Check5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text18.SetFocus
End If
End Sub

Private Sub DTShowStartDate_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = 13 Then
Text17.SetFocus
End If
End Sub

Private Sub Form_Activate()
'Me.Top = 2500
'Me.Left = 3000

DTShowStartDate.Value = Date

Combo2.Clear
'Combo1.Clear

SQL = "Select * from  Staff_Division"
Set dataset = mddata(SQL)

With dataset

If .RecordCount > 0 Then
.MoveFirst
End If

While .EOF = False
Combo2.additem !Division
'Combo1.AddItem !Division
.MoveNext
Wend
End With


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text9.SetFocus
Else
Label20.Caption = 1
End If
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus

End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text8.SetFocus

End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text13.SetFocus

End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Check1.SetFocus
Frame1.Visible = False
End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text11.SetFocus

End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text25.SetFocus
End If
End Sub


Private Sub cal()
Text25.Text = Val(Text5.Text) + Val(Text4.Text) + Val(Text6.Text) + Val(Text15.Text)
End Sub


Private Sub Text18_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Check3.SetFocus
End If
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text25_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text27.SetFocus
End If
End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If Text26.Text = "" Then
If Combo2.Text = "" Then
Combo2.SetFocus
SendKeys "{f4}"
End If

Else

SQL = "Select * from  Employee where Link_Code= " + Text26.Text
Set dataset = mddata(SQL)

With dataset
   
    


 
  If .RecordCount > 0 Then
  Combo3.Text = !Name
 Combo3_Click
  End If
  
End With
End If

End If
End Sub

Private Sub Text27_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text23.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text9.SetFocus
End If
End Sub

Private Sub Text4_Change()
cal
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text5_Change()
cal
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus

End If
End Sub

Private Sub Text6_Change()
cal
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text15.SetFocus
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text14.SetFocus
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
Else
Label19.Caption = 1
End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text16.SetFocus
End If
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3.SetFocus
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text20.SetFocus
End If
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text19.SetFocus
End If
End Sub

