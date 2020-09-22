Attribute VB_Name = "main"
Public Type TDBUser
   username As String
   password As String
   usertype As String
   database As String
End Type

Public Type TDBase
   databasename As String
   databasefile As String
   status As String
End Type

Public HASCHANGED As Integer
Public DBUSER(400) As TDBUser
Public DB(400) As TDBase

Public Sub RefreshList()
   frmmain.List1.Clear
   frmmain.List2.Clear
   For t = 0 To 400
      If DBUSER(t).username <> "" Then
         frmmain.List1.AddItem "User    :" & DBUSER(t).username & "   for   " & DBUSER(t).database
         frmmain.List2.AddItem "U" & Trim$(Str$(t))
      End If
   Next t
   For t = 0 To 400
      If DB(t).databasename <> "" Then
         frmmain.List1.AddItem "Database:" & DB(t).databasename & "   status:" & DB(t).status
         frmmain.List2.AddItem "D" & Trim$(Str$(t))
      End If
   Next t
End Sub

Public Function AddUser(username As String, password As String, database As String, utype As Integer) As Integer
   Select Case utype
      Case 1
         uts$ = "admin"
      Case 2
         uts$ = "normal"
      Case 3
         uts$ = "disabled"
      Case 4
         ut$ = "readonly"
      Case Else
         ut$ = "readonly"
   End Select
   For t = 0 To 400
      If DBUSER(t).username = "" Then
         DBUSER(t).username = username$
         DBUSER(t).password = password$
         DBUSER(t).usertype = uts$
         DBUSER(t).database = database$
         AddUser = 1
         HASCHANGED = 1
         Exit Function
      End If
   Next t
   AddUser = 0
End Function

Public Function AddDatabase(dbname As String, dbfile As String, activeval As Integer) As Integer
   For t = 0 To 400
      If DB(t).databasename = "" Then
         DB(t).databasename = dbname$
         DB(t).databasefile = dbfile$
         If activeval = 0 Then DB(t).status = "locked" Else DB(t).status = "open"
         AddDatabase = 1
         HASCHANGED = 1
         Exit Function
      End If
   Next t
   AddDatabase = 0
         
End Function
