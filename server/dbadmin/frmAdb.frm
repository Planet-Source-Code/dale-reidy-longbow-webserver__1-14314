VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAdb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Database"
   ClientHeight    =   2025
   ClientLeft      =   8520
   ClientTop       =   7665
   ClientWidth     =   4110
   Icon            =   "frmAdb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4110
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3045
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   1470
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update/Create"
      Height          =   330
      Left            =   2100
      TabIndex        =   7
      Top             =   1575
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   525
      TabIndex        =   6
      Top             =   1575
      Width           =   1380
   End
   Begin VB.CheckBox chkActive 
      Caption         =   "Active"
      Height          =   225
      Left            =   525
      TabIndex        =   5
      Top             =   1155
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse..."
      Height          =   330
      Left            =   2625
      TabIndex        =   4
      Top             =   1050
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   105
      Top             =   315
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.lbf"
      DialogTitle     =   "Choose Database File"
      Filter          =   "Longbow Database Files (*.lbf)|*.LBF|"
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1470
      TabIndex        =   2
      Top             =   735
      Width           =   2430
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1470
      TabIndex        =   1
      Top             =   210
      Width           =   2430
   End
   Begin VB.Label Label2 
      Caption         =   "Database File"
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   735
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Database Name"
      Height          =   225
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1275
   End
End
Attribute VB_Name = "frmAdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 cd.ShowOpen
 Text2.text = cd.filename
End Sub

Private Sub Command2_Click()
   Text3.text = "CANCEL"
   Me.Visible = False
End Sub

Private Sub Command3_Click()
   Text3.text = "CREATE"
   Me.Visible = False
End Sub
