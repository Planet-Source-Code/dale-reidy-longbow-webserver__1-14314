VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Longbow Database Administrator"
   ClientHeight    =   4815
   ClientLeft      =   4560
   ClientTop       =   2445
   ClientWidth     =   6150
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6150
   Begin MSComDlg.CommonDialog cd 
      Left            =   2205
      Top             =   2415
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Database Config File"
      Filter          =   "Database Config File|db_access.cfg"
   End
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   1575
      TabIndex        =   13
      Top             =   525
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Ref"
      Height          =   330
      Left            =   5460
      TabIndex        =   12
      Top             =   105
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database Access"
      Height          =   1170
      Left            =   105
      TabIndex        =   3
      Top             =   3570
      Width           =   6000
      Begin VB.CommandButton Command9 
         Caption         =   "&8  Exit"
         Height          =   330
         Left            =   4935
         TabIndex        =   11
         Top             =   735
         Width           =   960
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&4  Save"
         Height          =   330
         Left            =   4935
         TabIndex        =   10
         Top             =   315
         Width           =   960
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&7  Delete Database"
         Height          =   330
         Left            =   3150
         TabIndex        =   9
         Top             =   735
         Width           =   1800
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&6  Edit Database"
         Height          =   330
         Left            =   1680
         TabIndex        =   8
         Top             =   735
         Width           =   1485
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&3  Delete User"
         Height          =   330
         Left            =   3150
         TabIndex        =   6
         Top             =   315
         Width           =   1800
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&5  New Database"
         Height          =   330
         Left            =   105
         TabIndex        =   4
         Top             =   735
         Width           =   1590
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&2  Edit User"
         Height          =   330
         Left            =   1680
         TabIndex        =   7
         Top             =   315
         Width           =   1485
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&1  New User"
         Height          =   330
         Left            =   105
         TabIndex        =   5
         Top             =   315
         Width           =   1590
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   105
      TabIndex        =   2
      Top             =   525
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Load"
      Height          =   330
      Left            =   4725
      TabIndex        =   1
      Top             =   105
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Text            =   "d:\longbow\dbase\db_access.cfg"
      Top             =   105
      Width           =   4530
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If HASCHANGED = 1 Then
   dd = MsgBox("You have modified the database, it is recommended that you save before you reload. Reload?", vbCritical + vbYesNo, "Longbow Database Admin")
   If dd = vbNo Then Exit Sub
End If
On Error GoTo OpenError
   Open Text1.text For Input As #1
      Do Until EOF(1)
         Input #1, a$, b$, c$, d$, e$
         Select Case a$
            Case "[USER]"
               DBUSER(ax).username = m_is2e.Decrypt(b$, "WHATSTHEMATTER")
               DBUSER(ax).password = m_is2e.Decrypt(c$, "WHATSTHEMATTER")
               DBUSER(ax).usertype = m_is2e.Decrypt(d$, "WHATSTHEMATTER")
               DBUSER(ax).database = m_is2e.Decrypt(e$, "WHATSTHEMATTER")
               ax = ax + 1
            Case "[DBASE]"
               DB(bx).databasename = m_is2e.Decrypt(b$, "WHATSTHEMATTER")
               DB(bx).databasefile = m_is2e.Decrypt(c$, "WHATSTHEMATTER")
               DB(bx).status = m_is2e.Decrypt(d$, "WHATSTHEMATTER")
               bx = bx + 1
         End Select
      Loop
   Close 1
   HASCHANGED = 0
   RefreshList
   Exit Sub
OpenError:
   MsgBox Err.Description, vbExclamation, "Longbow Database Admin"
   Close 1
End Sub

Private Sub Command10_Click()
 RefreshList
End Sub

Private Sub Command2_Click()
   With frmAdb
      .Text1.text = ""
      .Text2.text = ""
      .Text3.text = "OPEN"
      .chkActive.Value = 1
   End With
   frmAdb.Visible = True
   Do Until frmAdb.Visible = False
      DoEvents
   Loop
   If frmAdb.Text3.text = "CANCEL" Then Exit Sub
   xx = AddDatabase(frmAdb.Text1.text, frmAdb.Text2.text, frmAdb.chkActive.Value)
   RefreshList
   If xx = 0 Then
      MsgBox "Unable to create database", vbCritical, "Longbow Database Admin"
      Exit Sub
   End If
End Sub

Private Sub Command3_Click()
   Dim OPT As Integer
   With frmAddDBUser
      .Text1.text = ""
      .Text2.text = ""
      .optNorm.Value = True
      .Combo1.Clear
      For t = 0 To 400
         If DB(t).databasename <> "" Then .Combo1.AddItem DB(t).databasename
      Next t
   End With
   frmAddDBUser.Text3.text = "OPEN"
   frmAddDBUser.Show
   Do Until frmAddDBUser.Text3.text <> "OPEN"
      If frmAddDBUser.Caption = "CLOSING..." Then Exit Sub
      DoEvents
   Loop
   If frmAddDBUser.Text3.text <> "CREATE" Then Exit Sub
   ' Add The User, and refresh the list
   With frmAddDBUser
      If .optAdmin.Value = True Then OPT = 1
      If .optNorm.Value = True Then OPT = 2
      If .optDisabled.Value = True Then OPT = 3
      If .optReadOnly.Value = True Then OPT = 4
      'Debug.Print frmAddDBUser.Text1.text
      Did_CREATE = AddUser(.Text1.text, .Text2.text, .Combo1.List(.Combo1.ListIndex), OPT)
      If Did_CREATE = 0 Then
         MsgBox "Unable To Create New User", vbExclamation, "Longbow Database Admin"
         Exit Sub
      End If
      main.RefreshList
   End With
End Sub


Private Sub Command4_Click()
   If List1.ListIndex = -1 Then Exit Sub
   
   d$ = List2.List(List1.ListIndex)
   
   If Left$(d$, 1) = "U" Then
      'It is a user
      un = Val(Right$(d$, Len(d$) - 1))
      dd = MsgBox("Are you sure you want to delete" & vbCrLf & DBUSER(un).username & " ?", vbCritical + vbYesNo, "Longbow Database Admin")
      If dd = vbNo Then Exit Sub
      DBUSER(un).database = ""
      DBUSER(un).password = ""
      DBUSER(un).username = ""
      DBUSER(un).usertype = ""
      MsgBox "User Deleted", vbInformation, "Longbow Database Admin"
      RefreshList
   End If
End Sub

Private Sub Command5_Click()
   If List1.ListIndex = -1 Then Exit Sub
   
   d$ = List2.List(List1.ListIndex)
   
   If Left$(d$, 1) = "U" Then
      'It is a user
      un = Val(Right$(d$, Len(d$) - 1))
      
      With frmAddDBUser
         .Text3.text = "OPEN"
         .Text1.text = DBUSER(un).username
         .Text2.text = DBUSER(un).password
         
         .Combo1.Clear
         For t = 0 To 400
            If DB(t).databasename <> "" Then .Combo1.AddItem DB(t).databasename
            If DB(t).databasename = DBUSER(un).database Then AGGA = t
         Next t
         
         .Combo1.ListIndex = AGGA
         dd$ = DBUSER(un).usertype
                        
                        Select Case dd$
                           Case "admin"
                              .optAdmin.Value = True
                           Case "disabled"
                              .optDisabled.Value = True
                           Case "readonly"
                              .optReadOnly.Value = True
                           Case "normal"
                              .optNorm.Value = True
                        End Select
                     End With
                     frmAddDBUser.Show
                     
                  Do Until frmAddDBUser.Text3.text <> "OPEN"
                        If frmAddDBUser.Caption = "CLOSING..." Then Exit Sub
                        DoEvents
                  Loop
      
      'Debug.Print "UPDATING"
      
      DBUSER(un).password = frmAddDBUser.Text2.text
      DBUSER(un).username = frmAddDBUser.Text1.text
                  
                  With frmAddDBUser
                     If .optAdmin.Value = True Then OPT = 1
                     If .optNorm.Value = True Then OPT = 2
                     If .optDisabled.Value = True Then OPT = 3
                     If .optReadOnly.Value = True Then OPT = 4
                  End With
      
                  Select Case OPT
                  Case 1
                     uts$ = "admin"
                  Case 2
                     uts$ = "normal"
                  Case 3
                     uts$ = "disabled"
                  Case 4
                     uts$ = "readonly"
                  Case Else
                     uts$ = "readonly"
               End Select
      DBUSER(un).usertype = uts$
   End If
   RefreshList
End Sub

Private Sub Command6_Click()
   If List1.ListIndex = -1 Then Exit Sub
   
   d$ = List2.List(List1.ListIndex)
   If Left$(d$, 1) = "D" Then
      'It is a user
      un = Val(Right$(d$, Len(d$) - 1))
      frmAdb.Text1.text = DB(un).databasename
      frmAdb.Text2.text = DB(un).databasefile
      If DB(un).status = "open" Then frmAdb.chkActive.Value = 1 Else frmAdb.chkActive.Value = 0
      frmAdb.Text3.text = "OPEN"
      frmAdb.Show
      Do Until frmAdb.Visible = False
         DoEvents
      Loop
      DB(un).databasefile = frmAdb.Text2.text
      DB(un).databasename = frmAdb.Text1.text
      If frmAdb.chkActive.Value = 0 Then DB(un).status = "locked" Else DB(un).status = "open"
      RefreshList
   End If
End Sub

Private Sub Command7_Click()
   If List1.ListIndex = -1 Then Exit Sub
   
   d$ = List2.List(List1.ListIndex)
   
   If Left$(d$, 1) = "D" Then
      'It is a user
      un = Val(Right$(d$, Len(d$) - 1))
      dd = MsgBox("Are you sure you want to delete" & vbCrLf & DB(un).databasename & " ?", vbCritical + vbYesNo, "Longbow Database Admin")
      If dd = vbNo Then Exit Sub
      DB(un).databasefile = ""
      DB(un).databasename = ""
      DB(un).status = ""
      MsgBox "Database Deleted", vbInformation, "Longbow Database Admin"
      RefreshList
   End If
End Sub

Private Sub Command8_Click()
   d$ = frmmain.Caption
   frmmain.Caption = "Saving " & Text1.text
   Open Text1.text For Output As #1
      For t = 0 To 400
         If DBUSER(t).username <> "" Then
            Print #1, "[USER]"
            Print #1, m_is2e.Encrypt(DBUSER(t).username, "WHATSTHEMATTER")
            Print #1, m_is2e.Encrypt(DBUSER(t).password, "WHATSTHEMATTER")
            Print #1, m_is2e.Encrypt(DBUSER(t).usertype, "WHATSTHEMATTER")
            Print #1, m_is2e.Encrypt(DBUSER(t).database, "WHATSTHEMATTER")
         End If
      Next t
      For t = 0 To 400
         If DB(t).databasename <> "" Then
            Print #1, "[DBASE]"
            Print #1, m_is2e.Encrypt(DB(t).databasename, "WHATSTHEMATTER")
            Print #1, m_is2e.Encrypt(DB(t).databasefile, "WHATSTHEMATTER")
            Print #1, m_is2e.Encrypt(DB(t).status, "WHATSTHEMATTER")
            Print #1, m_is2e.Encrypt("NULL", "WHATSTHEMATTER")
         End If
      Next t
   Close 1
   HASCHANGED = 0
   frmmain.Caption = d$
End Sub

Private Sub Command9_Click()
   dd = MsgBox("Sure you want to exit?", vbCritical + vbYesNo, "Longbow Database Admin")
   If dd = vbNo Then Exit Sub
   SaveSetting "Longbow", "DBaseAdmin", "ConfigFile", Text1.text
End
End Sub

Private Sub Form_Activate()
   Text1.text = ""
   Text1.text = GetSetting("LongBow", "DBaseAdmin", "ConfigFile")
   If Text1.text <> "" Then Command1_Click: Exit Sub
   List1.SetFocus
End Sub

Private Sub Form_UnLoad(cancel As Integer)
   Unload frmAddDBUser
   End
End Sub

Private Sub Text1_DblClick()
   cd.ShowOpen
   Text1.text = cd.filename
   Command1_Click
   SaveSetting "Longbow", "DBaseAdmin", "ConfigFile", Text1.text

End Sub
