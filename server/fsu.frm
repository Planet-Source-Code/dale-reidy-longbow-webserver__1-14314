VERSION 5.00
Begin VB.Form fsu 
   Caption         =   "Scripting Resources"
   ClientHeight    =   3195
   ClientLeft      =   6405
   ClientTop       =   3975
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.DirListBox di 
      Height          =   2565
      Index           =   0
      Left            =   1470
      TabIndex        =   2
      Top             =   315
      Width           =   1380
   End
   Begin VB.FileListBox fi 
      Height          =   2430
      Index           =   0
      Left            =   2835
      TabIndex        =   1
      Top             =   315
      Width           =   1590
   End
   Begin VB.PictureBox pic 
      Height          =   2220
      Index           =   0
      Left            =   210
      ScaleHeight     =   2160
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   105
      Width           =   3060
   End
End
Attribute VB_Name = "fsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
