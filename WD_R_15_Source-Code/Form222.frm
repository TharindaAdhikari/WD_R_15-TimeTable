VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7890
   LinkTopic       =   "Form2"
   ScaleHeight     =   5775
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   393216
      Format          =   114622465
      CurrentDate     =   44320
   End
   Begin VB.Timer Timer2 
      Left            =   2280
      Top             =   3960
   End
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   3840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1095
      Left            =   3480
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   1080
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


sql = "select * from session"
Set dataset = mddata(sql)
With dataset



Label1.Caption = "Mr." + RTrim(!lecturer) + "." + RTrim(!subject) + "." + RTrim(!Tag) + "." + RTrim(!Group)
'Label2.Caption = !subject
'Label3.Caption = RTrim(Label1.Caption) + "." + RTrim(Label2.Caption)
End With


End Sub
