Attribute VB_Name = "m_main"
Public Type SVMIME
   ext As String
   mtype As String
End Type

Public Type SVVDIR
   virt As String
   real As String
   acti As String
End Type

Public Type SVHOST
   svr As String
   root As String
   acti As String
End Type

Public Type SVDATA
   ServerName As String       ' The server name
   ServerAdmin As String      ' Server admins name
   ListenPort As Integer      ' Listen port for server
   MaxSocks As Integer        ' Maximum sockets
   DefaultRoot As String      ' Default server root
   DocLoc As String           ' Doc file directory
   LogLoc As String           ' Log file directory
   LogType As Integer         ' 0=NONE 1=FULL 2=ERRORS 3=REQUESTS
   IndexFile As String        ' Filename
   SecurityFile As String     ' Filename
   DirListing As Integer      ' 0=NONE 1=SIMPLE 2=GRAPHICAL
   TimerUpdate As Integer     ' Milliseconds For Timer Update
   TimeOut As Integer         ' Seconds To Timeout
   DatabaseCfg As String   ' Database Configuration File

End Type

Public Type SVUSERS
   username As String
   password As String
   Active As String           ' 0=NO 1=YES
   Directory As String        ' Users directory
End Type

Public Type SVSOCKET
   Buffer As String
   Header As String
   Reqok As Boolean
   TimeAlive As Long

End Type

Public Const LONGBOW_SERVER_DETAILS = "LongBow Server 1.0 By Dale Reidy"

Public ServerLogFile As String      ' Server Log File Name

Public ServerStartTime As Long      ' Timer Value from when server was started

Public Longbow As SVDATA

Public B64 As New Base64

Public ipban(200) As String
Public mimes(200) As SVMIME
Public users(2000) As SVUSERS
Public vdirz(60) As SVVDIR
Public vhost(60) As SVHOST

Public SERVER_SECURITY_TAG1$, SERVER_SECURITY_TAG2$ ' Server scripting security tags

Public NumReq As Long               ' Number of requests

Public sx() As SVSOCKET
Public cx() As New script

Public Sub InitServer()
   ServerStartTime = Timer
   
'   LoadMimes
'   LoadUsers
'   LoadVDirz
'   LoadHosts
   
   LoadServerDetails
   
   For t = 1 To Longbow.MaxSocks
      
      Load frmmain.ws(t)
      Load fsu.pic(t)
      Load fsu.fi(t)
      Load fsu.di(t)
   
   Next t
   
   ReDim sx(Longbow.MaxSocks)
   ReDim cx(Longbow.MaxSocks)
   

   frmmain.ws(0).LocalPort = Longbow.ListenPort
   frmmain.sxt.Interval = Longbow.TimerUpdate
   StartServer

End Sub

Public Sub StartServer()
If frmmain.ws(0).State <> sckClosed Then frmmain.ws(0).Close
   frmmain.ws(0).Listen
   frmmain.Command1.Enabled = True
   frmmain.Command2.Enabled = False

End Sub

Public Sub StopServer()
   frmmain.ws(0).Close
   frmmain.Command2.Enabled = True
   frmmain.Command1.Enabled = False
End Sub

Public Sub LoadServerDetails()
   
   ' Load the directory viewing color scheme and font face data
   LoadDirViewColorScheme
   
   'Public SERVER_SECURITY_TAG1$, SERVER_SECURITY_TAG2$ ' Server scripting security tags

   Open "..\conf\scriptsec.cfg" For Input As #1
      Line Input #1, SERVER_SECURITY_TAG1$
      Line Input #1, SERVER_SECURITY_TAG2$
   Close
   
   
   Open "..\conf\http.cfg" For Input As #1
   
   Do Until EOF(1)
      Line Input #1, x$
      d = InStr(x$, "=")
      a$ = Left$(x$, d - 1)
      b$ = Right$(x$, Len(x$) - d)
      
      Select Case a$
         Case "ServerName"
            Longbow.ServerName = b$
         
         Case "ServerAdmin"
            Longbow.ServerAdmin = b$
         
         Case "DatabaseCfg"
            Longbow.DatabaseCfg = b$
         
         Case "ListenPort"
            Longbow.ListenPort = Val(b$)
         
         Case "MaxSocks"
            Longbow.MaxSocks = Val(b$)
         
         Case "DefaultRoot"
            Longbow.DefaultRoot = b$
         
         Case "DocLoc"
            Longbow.DocLoc = b$
         
         Case "LogLoc"
            Longbow.LogLoc = b$
         
         Case "LogType"
            Longbow.LogType = Val(b$)
         
         Case "IndexFile"
            Longbow.IndexFile = b$
         
         Case "SecurityFile"
            Longbow.SecurityFile = b$
         
         Case "DirListing"
            Longbow.DirListing = Val(b$)
         
         Case "TimerUpdate"
            Longbow.TimerUpdate = Val(b$)
         
         Case "TimeOut"
            Longbow.TimeOut = Val(b$)
      
      End Select
   Loop
   
   Close 1
   
   xd = 0
   Open "..\conf\mime.cfg" For Input As #1
      Do Until EOF(1)
         Input #1, mimes(xd).ext, mimes(xd).mtype
         xd = xd + 1
      Loop
   Close 1
      
   xd = 0
   Open "..\conf\vdir.cfg" For Input As #1
      Do Until EOF(1)
         Input #1, vdirz(xd).virt, vdirz(xd).real, vdirz(xd).acti
         xd = xd + 1
      Loop
   Close 1
   
   xd = 0
   Open "..\conf\vhost.cfg" For Input As #1
      Do Until EOF(1)
         Input #1, vhost(xd).svr, vhost(xd).root, vhost(xd).acti
         xd = xd + 1
      Loop
   Close 1
   
   xd = 0
   Open "..\conf\users.cfg" For Input As #1
      Do Until EOF(1)
         Input #1, users(xd).username, users(xd).password, users(xd).Active, users(xd).Directory
         xd = xd + 1
      Loop
   Close 1

   'Debug.Print vhost(0).svr
   
   LoadBannedIPS

End Sub

Public Sub Cout(text As String)
   frmmain.Text1.text = frmmain.Text1.text & text$

End Sub

Public Sub ClrScr()
   frmmain.Text1.text = ""

End Sub

Public Sub CloseServer()
   For t = 0 To Longbow.MaxSocks
      frmmain.ws(t).Close
      If t <> 0 Then Unload frmmain.ws(t)
   Next t

End Sub

Public Sub WXB(socket As Integer, text As String)
   sx(socket).Buffer = sx(socket).Buffer & text$

End Sub

Public Sub WX_FILE(socket As Integer, filename As String)
   ' Writes a file to the socket, this file will be duplicated
   ' exactly, no SSI or modification
   ' USES:IMAGES,NON-ACTIVE HTML, TEXTFILES
   On Error GoTo NOFILE
   dd = FreeFile
   Open filename For Binary As #dd
      f1 = LOF(dd)
      f2$ = Space$(f1)
      Get #1, , f2$
   Close dd
   sx(socket).Buffer = f2$
   Exit Sub
NOFILE:
   Close dd

End Sub


Public Sub WX_PROPER(sck As Integer, filename As String)
   WX_FILE sck, filename
   ' TO BE CHANGED ;)

End Sub

Public Function SetStringLen(strtochange As String, setlen As Long) As String
   Select Case Len(strtochange$)
   
      Case Is > setlen
         If Len(strtochange$) > setlen Then
            SetStringLen = Left$(strtochange$, setlen)
         End If
      
      Case setlen
         If Len(strtochange$) = setlen Then
            SetStringLen = strtochange$
         End If
      
      Case Is < setlen
         If Len(strtochange$) < setlen Then
            SetStringLen = strtochange$ & Space$(setlen - Len(strtochange$))
         End If
   End Select

End Function

Public Sub WLog(TextToLog As String)
   If ServerLogFile = "" Then
      ServerLogFile = ReplaceStr(Date$ & "_" & Time$, "-", "_")
      ServerLogFile = ReplaceStr(ServerLogFile, ":", "_")
      ServerLogFile = Longbow.LogLoc & "\" & ServerLogFile & ".log"
   End If
   dx = FreeFile
   Open ServerLogFile For Append As #dx
   Print #dx, Time$ & ":" & Date$ & "   :" & TextToLog$
   Close dx
End Sub

Public Sub AppendIPBan(sck As Integer)
   For t = 0 To 200
      If ipban(t) = "" Then ipban(t) = frmmain.ws(t).RemoteHostIP: Exit Sub
   Next t
End Sub

Public Sub RemoveIPBan(ip As String)
   For t = 0 To 200
      If ipban(t) = ip$ Then ipban(t) = "": Exit Sub
   Next t
End Sub

Public Function IPBanned(sck As Integer) As Integer
   For t = 0 To 200
      If ipban(t) = frmmain.ws(sck).RemoteHostIP Then IPBanned = 1: Exit Function
   Next t
   IPBanned = 0
End Function

Public Sub SaveBannedIPS()
   Dim dxx As Long
   dxx = FreeFile
   Open "..\conf\banip.ini" For Output As #dxx
   For t = 0 To 200
      If ipban(t) <> "" Then Print #1, ipban(t)
   Next t
   Close dxx
End Sub

Public Sub LoadBannedIPS()
   Dim dxx, xy As Long
   dxx = FreeFile
   xy = 0
   Open "..\conf\banip.ini" For Input As #dxx
   Do Until EOF(dxx)
      Line Input #dxx, ipban(xy)
      xy = xy + 1
   Loop
   Close dxx
End Sub

Public Sub AppendIPBanIP(ip As String)
   For t = 0 To 200
      If ipban(t) = "" Then ipban(t) = ip$: Exit Sub
   Next t
End Sub
