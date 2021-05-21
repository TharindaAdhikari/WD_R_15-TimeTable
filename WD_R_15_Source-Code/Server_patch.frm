VERSION 5.00
Begin VB.Form Server_patch 
   BorderStyle     =   0  'None
   Caption         =   "Server_patch"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   LinkTopic       =   "Form4"
   ScaleHeight     =   6000
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7695
      Begin VB.Frame Frame2 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   7215
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
            Left            =   3240
            TabIndex        =   0
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C000&
            Caption         =   "Sql Server Name "
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
            Left            =   720
            TabIndex        =   3
            Top             =   840
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "Server_patch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
If Not Text1.Text = "" Then


'sql = "truncate table [Time_Table].[dbo].[server_patch]"
'Set dataset1 = mddata(sql)
'
'sql = "select * from server_patch"
'Set dataset = mddata(sql)
'
'With dataset
'
'.AddNew
'!server_name = Text1.Text
'
'.Update
'MsgBox "Done!", vbInformation


Me.Hide
Load.Show
'End With

Else
MsgBox "Enter the sql server name !"
Text1.SetFocus
End If
End If

End Sub
