Attribute VB_Name = "m_http"

' This is..evolution.. the monkey... the man... then gun

' This is.. your creation... the apple... of eden was a bomb

Public DIR_BACKCOLOR$, DIR_HEADCOLOR$, DIR_LISTCOLOR$, DIR_BARCOLOR$, DIR_LISTFACE$, DIR_HEADFACE$


Public Sub LoadDirViewColorScheme()
   On Error GoTo DIRVIEWLOADDEFAULTS
   dx = FreeFile
   Open "..\conf\dircols.cfg" For Input As #dx
      Do Until EOF(dx)
         Input #dx, x$, y$
         x$ = LCase$(x$)
         Select Case x$
            Case "dir_backcolor"
               DIR_BACKCOLOR$ = y$
            Case "dir_headcolor"
               DIR_HEADCOLOR$ = y$
            Case "dir_listcolor"
               DIR_LISTCOLOR$ = y$
            Case "dir_barcolor"
               DIR_BARCOLOR$ = y$
            Case "dir_listface"
               DIR_LISTFACE$ = y$
            Case "dir_headface"
               DIR_HEADFACE$ = y$
         End Select
      Loop
   Close dx
   Exit Sub
DIRVIEWLOADDEFAULTS:
   DIR_BACKCOLOR$ = "white"
   DIR_HEADCOLOR$ = "navy"
   DIR_LISTCOLOR$ = "blue"
   DIR_BARCOLOR$ = "red"
   DIR_LISTFACE$ = "fixedsys"
   DIR_HEADFACE$ = "verdana"
   Close dx
   Exit Sub
End Sub



Public Sub WriteHTTP(sck As Integer, rettype As Integer, lparam As String)
   zz$ = vbCrLf & "Server: LongBow" & vbCrLf & "Server-Version: 1.0b" & vbCrLf & "Server-Programmer: Dale Reidy(dreidy@btinternet.com)"
   
   WLog "HTTP_ERROR," & Trim$(Str$(rettype)) & "," & frmmain.ws(sck).RemoteHostIP
   
   Select Case rettype
      Case 400
         d$ = "HTTP/1.0 400 Bad Request" & zz$ & vbCrLf & vbCrLf & vbCrLf
      Case 200
         d$ = "HTTP/1.0 200 OK" & zz$ & vbCrLf & vbCrLf & vbCrLf
      Case 204
         d$ = "HTTP/1.0 200 OK But Empty File" & zz$ & vbCrLf & vbCrLf & vbCrLf
      Case 401
         d$ = "HTTP/1.0 401 Unauthorized" & zz$ & vbCrLf & "WWW-Authenticate: Basic realm=" & Chr$(34) & lparam$ & Chr$(34) & vbCrLf & vbCrLf & vbCrLf
      Case 403
         d$ = "HTTP/1.0 403 Unauthorized" & zz$ & vbCrLf & vbCrLf & vbCrLf
      Case 404
         d$ = "HTTP/1.0 404 File Not Found" & zz$ & vbCrLf & vbCrLf & vbCrLf
      Case 500
         d$ = "HTTP/1.0 500 Internal Server Error" & zz$ & vbCrLf & vbCrLf & vbCrLf
   End Select
   'Debug.print d$
   sx(sck).Buffer = d$ & sx(sck).Buffer
End Sub

Public Function GetWWWRoot(host_name As String) As String

   For t = 0 To 60
      If vhost(t).acti = "YES" And host_name$ = vhost(t).svr Then GetWWWRoot$ = vhost(t).root: Exit Function
   Next t
   GetWWWRoot = Longbow.DefaultRoot
      
End Function

Public Function ValidateUser(username As String, password As String, userlist As String) As Integer
   On Error GoTo BADBADERROR

   For t = 0 To 2000
      If (username$ = users(t).username) And (password$ = users(t).password) And InStr(userlist$, users(t).username) > 0 And (users(t).Active = "yes") Then

         ValidateUser = 1
         Exit Function
      End If
      If t = 1000 Then DoEvents ': Debug.Print "HALFWAY"
   Next t
   ValidateUser = 0
   Exit Function:
BADBADERROR:
  ValidateUser = 0
  WLog "VALIDATE USER ERROR " & username$ & "," & userlist$
End Function

Public Function GetFile(req As String) As String
   ' Get The File From A Path
   'c:\windows\desktop\hello.bmp
   On Error GoTo GETFILEERROR
   a = Len(req$)
   For b = a To 1 Step -1
    If Mid$(req$, b, 1) = "\" Then
      GetFile = Right$(req$, Len(req$) - b)
      Exit Function
    End If
   Next b
   GetFile = req$
   Exit Function
GETFILEERROR:
End Function

Public Sub Write_HTML(sck As Integer, freq As String, method As String, host As String)
   On Error GoTo WRITEHTMLERROR
   If Exists(freq) = 0 Then GoTo WRITEHTMLERROR
   dx = FreeFile
   Open freq For Binary As #dx
   d = LOF(dx)
   If d = 0 Then Close dx: WXB sck, "File Is Empty": sx(sck).Reqok = True: Exit Sub
   f$ = Space$(d)
   Get #dx, , f$
   Close dx
   f$ = ReplaceStr(f$, "<%METHOD%>", method$)
   f$ = ReplaceStr(f$, "<%HOST%>", host$)
   f$ = ReplaceStr(f$, "<%TIME%>", Time$)
   f$ = ReplaceStr(f$, "<%DATE%>", Date$)
   f$ = ReplaceStr(f$, "<%TIMER%>", Timer)
   f$ = ReplaceStr(f$, "<%SERVER%>", LONGBOW_SERVER_DETAILS)
   sx(sck).Buffer = f$
   Exit Sub
WRITEHTMLERROR:
   Close dx
   WriteHTTP sck, 500, "-"
End Sub

Public Sub Write_BINARY(sck As Integer, filename As String)
   On Error GoTo WRITEBINARYERROR

   If Exists(filename) = 0 Then GoTo WRITEBINARYERROR
   dx = FreeFile
   Open filename For Binary As #dx
   d = LOF(dx)
   If d = 0 Then Close dx: WXB sck, "File Is Empty": sx(sck).Reqok = True: Exit Sub
   f$ = Space$(d)
   Get #dx, , f$
   Close dx
   sx(sck).Buffer = f$
   Exit Sub
WRITEBINARYERROR:
   Close dx
   WriteHTTP sck, 500, "-"
End Sub

Public Sub Write_TEXT(sck As Integer, filename As String)
   On Error GoTo WRITETEXTERROR
   dx = FreeFile
   Open filename For Binary As #dx
   d = LOF(dx)
   If d = 0 Then Close dx: WXB sck, "File Is Empty": sx(sck).Reqok = True: Exit Sub
   f$ = Space$(d)
   Get #dx, , f$
   Close dx
   sx(sck).Buffer = f$
   Exit Sub
WRITETEXTERROR:
   Close dx
   WriteHTTP sck, 500, "-"
End Sub


Public Function GetMimeType(filename As String) As String
   For t = 0 To 200
      If InStr(filename, mimes(t).ext) Then GetMimeType = mimes(t).mtype: Exit Function
   Next t
End Function

Public Sub ProcessHeader(sck As Integer)
   'On Error GoTo PROCESSERROR
   Dim Errloc$
   
   NumReq = NumReq + 1
   
   ' *** GATHER ALL THE INFORMATION FROM THE HTTP HEADER ***
   Errloc$ = "HEADER"
   
   cheader$ = sx(sck).Header
   
   
   ' GET THE HTTP METHOD FOR THE REQUEST
      a$ = Left$(cheader$, 3)
   
      If a$ = "POS" Then method$ = "post" Else method$ = "get"
   
      a1 = InStr(cheader$, " ")
      a2 = InStr(a1 + 1, cheader$, " ")
   
   ' GET THE REQUEST
      request$ = Mid$(cheader, a1 + 1, a2 - a1 - 1)
   
      uprequest$ = request$
   
   ' PROCESS DATA PASSED AS PARAMETERS TO THE PAGE
      If InStr(request$, "?") Then
         a1 = InStr(request$, "?")
         postdata$ = "&" & Right$(request$, Len(request$) - a1)
         request$ = Left$(request$, a1 - 1)
      End If
   
      a1 = InStr(cheader$, vbCrLf & vbCrLf)
   
      a1 = a1 + Len(vbCrLf & vbCrLf)

      a1 = a1 - 1
      
      If a1 < Len(cheader$) Then
         temp1$ = Trim$(Right$(cheader$, Len(cheader$) - a1))
         'temp1$ = Left$(temp1$, Len(temp1$) - 2)
      End If
   
      If temp1$ <> "" Then
         If postdata$ = "" Then
            postdata$ = postdata$ & temp1$
         Else
            postdata$ = postdata$ & "&" & temp1$
         End If
      End If
   
   
   ' GET THE HOST
      a1 = InStr(cheader$, "Host:") + Len("Host: ")
   
      a2 = InStr(a1 + 1, cheader$, vbCrLf)
   
      host$ = Mid$(cheader$, a1, a2 - a1)
   
   ' GET AUTHORIZATION DATA
      If InStr(cheader$, "Authorization: Basic ") Then
         a1 = InStr(cheader$, "Authorization: Basic ") + Len("Authorization: Basic ")
         a2 = InStr(a1, cheader$, vbCrLf)
         auth_data$ = B64.Decode(Mid$(cheader, a1, a2 - a1))
         a1 = InStr(auth_data$, ":")
         auth_name$ = Left$(auth_data$, a1 - 1)
         auth_pword$ = Right$(auth_data$, Len(auth_data$) - a1)
      End If
   
   ' *** Replace virtual directory entries with actual server directory entries
   
   Errloc$ = "PROCESS": VDIR = 0: REQTYPE = 0: request$ = Trim$(request$)

         For t = 0 To 60
            old_request$ = request$
            request$ = Trim$(ReplaceStr(" " & request$, " " & vdirz(t).virt, vdirz(t).real))
            If old_request$ <> request$ Then VDIR = 1: Exit For
         Next t
   
      request$ = Trim$(ReplBack(request$))
   
      If VDIR = 0 Then request$ = GetWWWRoot(host$) & request$

         If request$ = "\" Then request$ = GetWWWRoot(host$)
   
               If IsDir(request$) = 1 Then REQTYPE = 1
               If Exists(request$) = 1 Then REQTYPE = 2
   
         If REQTYPE = 1 And Right$(request$, 1) <> "\" Then
            WXB sck, "<META HTTP-EQUIV=" & Chr$(34) & "Refresh" & Chr$(34) & " CONTENT=" & Chr$(34) & "0; URL=" & "http://" & host$ & ":" & Trim$(Str$(Longbow.ListenPort)) & uprequest$ & "/" & Chr$(34) & ">"
            sx(sck).Reqok = True
               Exit Sub
         End If
   
      If REQTYPE = 0 Then
         WriteHTTP sck, 404, "_"
         sx(sck).Reqok = True
         Exit Sub
      End If
   
      If REQTYPE = 1 Then req_dir$ = request$
      
      If REQTYPE = 2 Then req_dir$ = GetDirectory(request$): req_file$ = GetFilename(request$)

      If Exists(req_dir$ & Longbow.SecurityFile) = 0 Then
         WriteHTTP sck, 403, "_"
         sx(sck).Reqok = True
         Exit Sub
      End If
      
      ' Get The Security Settings
         
            DIR_READ = 1
            DIR_WRITE = 1
            DIR_EXECUTE = 1
            DIR_SECURITY = 1
            DIR_VIEW = 1
            DIR_USERS$ = ""
         
         
         dx = FreeFile
         Open req_dir$ & Longbow.SecurityFile For Input As #dx
         Do Until EOF(dx)
            Line Input #dx, f$
            f$ = LCase$(f$)
            Select Case f$
               Case "read=no"
                  DIR_READ = 0
               Case "write=no"
                  DIR_WRITE = 0
               Case "dirview=no"
                  DIR_VIEW = 0
               Case "execute=no"
                  DIR_EXECUTE = 0
               Case "secure=no"
                  DIR_SECURITY = 0
            End Select
            If Left$(f$, 6) = "domain" Then DIR_DOMAIN$ = Right$(f$, Len(f$) - 7)
            If Left$(f$, 5) = "users" Then DIR_USERS$ = Right$(f$, Len(f$) - 6)
         Loop
         Close dx
         
         If DIR_SECURITY = 1 Then
         
            If auth_name$ = "" Or auth_pword$ = "" Then
               WriteHTTP sck, 401, DIR_DOMAIN$
               sx(sck).Reqok = True
               Exit Sub
            End If
            
            CAN_ACCESS = ValidateUser(auth_name$, auth_pword$, DIR_USERS$)
         
            If CAN_ACCESS = 0 Then
               WriteHTTP sck, 403, "-"
               sx(sck).Reqok = True
               Exit Sub
            End If
            
         End If
         
         If LCase$(req_file$) = LCase$(Longbow.SecurityFile) Then
            WLog "Security File Attempted Access " & frmmain.ws(sck).RemoteHostIP
            AppendIPBan sck
            WriteHTTP sck, 403, "-"
            sx(sck).Reqok = True
            Exit Sub
         End If
            
         
         If REQTYPE = 1 And Exists(req_dir$ & Longbow.IndexFile) = 1 Then
            req_file$ = Longbow.IndexFile
            REQTYPE = 2
         End If
         
         If REQTYPE = 1 And DIR_READ = 0 Then
            If Exists(req_dir$ & Longbow.IndexFile) = 1 And DIR_READ = 1 Then
               req_file$ = Longbow.IndexFile
               REQTYPE = 2
            Else
               WriteHTTP sck, 403, "-"
               sx(sck).Reqok = True
               Exit Sub
            End If
         End If
         
         If REQTYPE = 2 And DIR_READ = 0 Then
            WriteHTTP sck, 403, "-"
            sx(sck).Reqok = True
            Exit Sub
         End If

         ADDIT$ = username$
         If ADDIT$ = "" Then ADDIT$ = "NoAuth"

         WLog frmmain.ws(sck).RemoteHostIP & "," & req_dir$ & req_file$

         Select Case REQTYPE
            Case 1
' DIR_BACKCOLOR$, DIR_HEADCOLOR$, DIR_LISTCOLOR$, DIR_BARCOLOR$, DIR_LISTFACE$, DIR_HEADFACE$
            WXB sck, "<!-- AUTO GENERATED DIRECTORY LISTING -->" & vbCrLf
            WXB sck, "<html><body bgcolor=" & DIR_BACKCOLOR$ & ">" & vbCrLf
            WXB sck, "<font face=" & Chr$(34) & DIR_HEADFACE$ & Chr$(34) & " color=" & DIR_HEADCOLOR$ & ">" & vbCrLf
            WXB sck, "<h1 align=center>Directory Listing For " & uprequest$ & "</h1></font>" & vbCrLf
            WXB sck, "<hr noshade color=" & DIR_BARCOLOR$ & ">" & vbCrLf
            WXB sck, "<font face=" & DIR_LISTFACE$ & " color=" & DIR_LISTCOLOR$ & ">" & vbCrLf
            
            fsu.di(sck).Path = req_dir$
            fsu.fi(sck).Path = req_dir$
            
            HREF$ = "http://" & host$ & ":" & Trim$(Str$(Longbow.ListenPort))
            
               If uprequest$ = "/" Then
                  For t = 0 To 60
                     If vdirz(t).virt <> "" Then
                        WXB sck, "<IMG SRC=" & Chr$(34) & HREF$ & "/docs/FILE_DIRECTORY.GIF" & Chr$(34) & ">"
                        WXB sck, "<A HREF=" & Chr$(34) & HREF$ & vdirz(t).virt & "/" & Chr$(34) & ">" & vdirz(t).virt & "</a><br>" & vbCrLf
                     End If
                  Next t
               Else
                  WXB sck, "<IMG SRC=" & Chr$(34) & HREF$ & "/docs/FILE_DIRECTORY.GIF" & Chr$(34) & ">"
                  WXB sck, "<A HREF=" & Chr$(34) & ".." & Chr$(34) & ">..</a><br>"
               End If
               
               For t = 0 To fsu.di(sck).ListCount - 1
                  d$ = LCase$(GetFilename(fsu.di(sck).List(t)))
                  WXB sck, "<IMG SRC=" & Chr$(34) & HREF$ & "/docs/FILE_DIRECTORY.GIF" & Chr$(34) & ">"
                  WXB sck, "<A HREF=" & Chr$(34) & HREF$ & "/" & d$ & "/" & Chr$(34) & ">/" & d$ & "</a><br>" & vbCrLf
               Next t
               
               For t = 0 To fsu.fi(sck).ListCount - 1
                  
                  d$ = fsu.fi(sck).List(t)
                  If LCase$(d$) = LCase$(Longbow.SecurityFile) Then GoTo NOTHIS
                  x$ = GetMimeType(d$)
                  If x$ = "" Then x$ = "UNKNOWN"
                  WXB sck, "<IMG SRC=" & Chr$(34) & HREF$ & "/docs/FILE_" & x$ & ".GIF" & Chr$(34) & ">"
                  WXB sck, "<A HREF=" & Chr$(34) & HREF$ & uprequest$ & d$ & Chr$(34) & ">" & d$ & "</a><br>" & vbCrLf
NOTHIS:
                  
               Next t
               
               WXB sck, "</font></body></html>"
               sx(sck).Reqok = True
               Exit Sub
            Case 2
               e$ = GetMimeType(req_file$)
                  Select Case e$
                     Case "SCRIPT"
                        cx(sck).Execute sck, req_dir$ & req_file$, postdata$
                        sx(sck).Reqok = True
                        Exit Sub
                     Case "HTML"
                        Write_HTML sck, req_dir$ & req_file$, method$, host$
                        sx(sck).Reqok = True
                        Exit Sub
                     Case Else
                        Write_BINARY sck, req_dir$ & req_file$
                        sx(sck).Reqok = True
                        Exit Sub
                  End Select
               WriteHTTP sck, 500, "-"
               sx(sck).Reqok = True
               Exit Sub
         End Select
   
   WXB sck, "<html><body><font face=courier size=3>"
   WXB sck, "Request:" & request$ & ".<br>" & vbCrLf
   WXB sck, "Method:" & method$ & ".<br>" & vbCrLf
   WXB sck, "PostData:" & postdata$ & ".<br>" & vbCrLf
   WXB sck, "Host:" & host$ & ".<br>" & vbCrLf
   WXB sck, "</font></body></html>"
   sx(sck).Reqok = True
   
   Exit Sub
PROCESSERROR:
   WLog "Error Processing Request"
   Select Case Errloc
      Case "HEADER"
         WriteHTTP sck, 400, "-"
      Case Else
         WriteHTTP sck, 500, "_"
   End Select
   sx(sck).Reqok = True
End Sub
