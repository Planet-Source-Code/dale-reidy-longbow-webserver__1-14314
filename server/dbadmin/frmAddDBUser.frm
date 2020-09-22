VERSION 5.00
Begin VB.Form frmAddDBUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Database User"
   ClientHeight    =   2805
   ClientLeft      =   4560
   ClientTop       =   7665
   ClientWidth     =   3840
   Icon            =   "frmAddDBUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3840
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   105
      TabIndex        =   14
      Top             =   1890
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update/Create"
      Height          =   330
      Left            =   1995
      TabIndex        =   13
      Top             =   2310
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   315
      TabIndex        =   12
      Top             =   2310
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   1050
      TabIndex        =   7
      Top             =   1365
      Width           =   2745
      Begin VB.OptionButton optReadOnly 
         Caption         =   "Read Only"
         Height          =   225
         Left            =   1365
         TabIndex        =   11
         Top             =   525
         Width           =   1170
      End
      Begin VB.OptionButton optDisabled 
         Caption         =   "Disabled"
         Height          =   225
         Left            =   210
         TabIndex        =   10
         Top             =   525
         Width           =   1275
      End
      Begin VB.OptionButton optNorm 
         Caption         =   "Normal"
         Height          =   225
         Left            =   1365
         TabIndex        =   9
         Top             =   210
         Width           =   1170
      End
      Begin VB.OptionButton optAdmin 
         Caption         =   "Admin"
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   210
         Width           =   1170
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1155
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1050
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1155
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   630
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1155
      TabIndex        =   0
      Top             =   210
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "User Type"
      Height          =   225
      Left            =   210
      TabIndex        =   6
      Top             =   1575
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Database"
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   1050
      Width           =   750
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   630
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   225
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   855
   End
End
Attribute VB_Name = "frmAddDBUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Text3.text = "CANCEL"
   Me.Visible = False
End Sub

Private Sub Command2_Click()
   Text3.text = "CREATE"
   Me.Visible = False
End Sub
