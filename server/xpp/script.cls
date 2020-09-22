VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "script"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private code$(3000)
Private std$(1000), stn$(1000), ind(1000), inn$(1000)
Private gs(10) ' Gosub Stack

Private file(10) As Integer ' File access numbers
Private filea(10) As Integer ' File access type

Public Sub Execute(sck As Integer, filename As String, postdata As String)
   'On Error GoTo Execute_Error
   Dim p$(200), w$(20)
   Dim scr, iq As Boolean
   Dim ww As Integer
   Dim x As Integer
   
   Erase code$(), std$(), stn$(), ind(), inn$()
   For t = 0 To 10
    If file(t) <> 0 Then Close file(t): filea(t) = 0
   Next t
   Erase file(), filea()
   
   ' Load The Script File
   
   dx = FreeFile
   lc = 0
   Open filename For Input As #dx
   Do Until EOF(dx)
      Line Input #1, code$(lc)
      code$(lc) = ReplaceStr(code$(lc), vbTab, "   ")
      lc = lc + 1
   Loop
   Close dx
   

   'Debug.Print postdata$
   
   postdata$ = ReplaceStr(postdata$, vbCrLf, " ")
   postdata$ = Trim$(postdata$)
   
   postdata$ = ReplaceStr(postdata$, "&&", "&")
   If Right$(postdata$, 1) <> "&" Then postdata$ = postdata$ & "&"
   a = Len(postdata$)
   For b = 1 To a
      d$ = Mid$(postdata$, b, 1)
      If d$ = "&" And tmp$ <> "" Then
         c = InStr(tmp$, "=")
         If InStr(tmp$, "=") Then
            CreateString RidFormatting(Left$(tmp$, c - 1)), RidFormatting(Right$(tmp$, Len(tmp$) - c))
         End If
         tmp$ = ""
      End If
      If d$ <> "&" Then tmp$ = tmp$ & d$
   Next b
   
   ' Begin Parsing
   
   pc = 0
   scr = True
   
   If code$(0) = "//" & SERVER_SECURITY_TAG1$ And code$(1) = "//" & SERVER_SECURITY_TAG2$ Then
      SFA = 1
      'Debug.Print "Secure"
   Else
      SFA = 0
   End If
   
   For tt = 0 To lc
      If LTrim$(LCase$(code$(tt))) = "proc main" Then pc = tt + 1: Exit For
   Next tt
   
   Do Until scr = False
      
      cl$ = LTrim$(code$(pc))
      
      If Left$(cl$, 2) <> "//" And cl$ <> "" Then
        
         ' Only parse if the line is valid
         
         iq = False
         
         Erase p$()
         
         v$ = LTrim$(code$(pc))
         
         c2 = Len(v$)
         
         xc = 0
         
         dd = 0
         
         For c1 = 1 To c2
            
            f$ = Mid$(v$, c1, 1)
            
            If f$ = "(" And iq = False Then dd = dd + 1
            If f$ = ")" And iq = False Then dd = dd - 1
            
            If f$ = " " And iq = False And dd = 0 Then xc = xc + 1
            
            If f$ = "'" Then iq = Not iq
            
            If iq = False And dd = 0 Then p$(xc) = p$(xc) & Trim$(f$)
            If iq = False And dd <> 0 Then p$(xc) = p$(xc) & f$
            
            If iq = True Then p$(xc) = p$(xc) & f$
            
         Next c1
            
         Erase w$()
         
         cmd$ = LCase$(p$(0))
         
         xc = xc + 1
         
         For t = 1 To xc
            w$(t) = p$(t)
            'Debug.Print w$(t)
         Next t
         
               ' Execute the commands
               
               
               
               'Debug.Print pc, cmd$
               'Debug.Print ind(0), inn(0)
               'Debug.Print GetString("stcount$"), GetInteger("incount%"), cmd$
               
               If IsInteger(cmd$) = 1 Then
                  ww = 0
                  For t = 1 To xc
                     tturn = 1
                     If w$(t) = "+" Or w$(t) = "-" Or w$(t) = "*" Or w$(t) = "/" Then
                        xx$ = w$(t): tturn = 0
                     End If
                     If tturn = 1 Then
                        If xx$ = "+" Then ww = ww + Val(ParseParam(w$(t)))
                        If xx$ = "-" Then ww = ww - Val(ParseParam(w$(t)))
                        If xx$ = "/" Then ww = ww / Val(ParseParam(w$(t)))
                        If xx$ = "*" Then ww = ww * Val(ParseParam(w$(t)))
                     End If
                     If xx$ = "" Then ww = Val(ParseParam(w$(t)))
                  Next t
                  SetInteger cmd$, ww
               End If
               
                        
                     
               
               
               If cmd$ = "dec" Then
                  If IsInteger(w$(1)) = 1 Then
                     f = Val(w$(2))
                     If f = 0 Then f = 1
                     SetInteger w$(1), GetInteger(w$(1)) - f
                  End If
               End If
               
               If cmd$ = "inc" Then
                  If IsInteger(w$(1)) = 1 Then
                     f = Val(w$(2))
                     If f = 0 Then f = 1
                     'Debug.Print f
                     'Debug.Print GetInteger(w$(1))
                     'Debug.Print GetInteger(w$(1)) + f
                     SetInteger w$(1), GetInteger(w$(1)) + f
                  End If
               End If
               
               If cmd$ = "strtoint" Then
                  If IsString(w$(1)) = 1 Then
                     If IsInteger(w$(2)) = 1 Then
                        SetInteger w$(2), Val(GetString(w$(1)))
                     End If
                  End If
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "inttostr" Then
                  If IsInteger(w$(1)) = 1 Then
                     If IsString(w$(2)) = 1 Then
                        SetString w$(2), Trim$(Str$(GetInteger(w$(1))))
                     End If
                  End If
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "fwrite" Then
                  If file(Val(w$(1))) <> 0 Then
                     Print #file(Val(w$(1))), ParseParam(w$(2))
                  End If
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "fmread" Then
                  ' Multi Param Read
                  If file(Val(w$(1))) <> 0 Then
                     Line Input #file(Val(w$(1))), xtm$
                     Erase p$()
                     df = 0: cv = 1
                     dj = Len(xtm$)
                     For h = 1 To dj
                        hi$ = Mid$(xtm$, h, 1)
                        If hi$ = "," Then
                           hi$ = "": cv = cv + 1
                        End If
                           p$(cv) = p$(cv) + hi$

                     Next h
                    ' Debug.Print cv
                     For h = 1 To cv
                        'If p$(h) <> "" Then Debug.Print p$(h)
                        If w$(h + 1) <> "" Then
                           SetString w$(h + 1), p$(h)
                        End If
                     Next h
                  End If
               End If
               
               If cmd$ = "fread" Then
                  If file(Val(w$(1))) <> 0 Then
                     Input #file(Val(w$(1))), g$
                     'Debug.Print g$
                     'Debug.Print w$(2)
                     Debug.Print IsString(w$(2))
                     Debug.Print GetString(w$(2))
                     If IsString(w$(2)) = 1 Then SetString w$(2), g$
                     If IsInteger(w$(2)) = 1 Then SetInteger w$(2), Val(g$)
                  End If
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "file" Then
                  Select Case w$(1)
                     
                     Case "openappend"
                        Debug.Print "OPENING FOR APPEND"
                        w$(2) = ParseParam(w$(2))
                        xx = Val(w$(3))
                        If file(xx) = 0 Then
                            file(xx) = FreeFile
                            If InStr(w$(2), "..") Or InStr(w$(2), Longbow.SecurityFile) Then
                                WXB sck, "Script Security Prevented This Script From Being Run"
                                sx(sck).Reqok = True
                                Exit Sub
                            End If
                            Open GetDirectory(filename$) & w$(2) For Append As file(xx)
                        End If
                     
                     Case "openin"
                        w$(2) = ParseParam(w$(2))
                        xx = Val(w$(3))
                        If file(xx) = 0 Then
                           file(xx) = FreeFile
                           If InStr(w$(2), "..") Or InStr(w$(2), Longbow.SecurityFile) Then
                              WXB sck, "Script Security Prevented This Script From Being Run"
                              sx(sck).Reqok = True
                              Exit Sub
                           End If
                           If Exists(GetDirectory(filename$) & w$(2)) = 1 Then
                              Open GetDirectory(filename$) & w$(2) For Input As file(xx)
                           Else
                              file(xx) = 0: filea(xx) = 0
                           End If
                        End If
                     
                     Case "openout"
                        w$(2) = ParseParam(w$(2))
                        xx = Val(w$(3))
                        If file(xx) = 0 Then
                           file(xx) = FreeFile
                           If InStr(w$(2), "..") Or InStr(w$(2), Longbow.SecurityFile) Then
                              WXB sck, "Script Security Prevented This Script From Being Run"
                              sx(sck).Reqok = True
                              Exit Sub
                           End If
                           Open GetDirectory(filename$) & w$(2) For Output As file(xx)
                        End If
                     
                     Case "close"
                        If file(Val(w$(2))) <> 0 Then
                           Close file(Val(w$(2)))
                           file(Val(w$(2))) = 0
                           filea(Val(w$(2))) = 0
                        End If
                     
                     Case "kill"
                        If Exists(GetDirectory(filename$) & w$(2)) = 1 Then
                           Kill GetDirectory(filename$) & w$(2)
                        End If
                     
                     Case "copy"
                        If Exists(GetDirectory(filename$) & w$(2)) = 1 Then
                           SecureCopy GetDirectory(filename$) & w$(2), GetDirectory(filename$) & w$(3)
                        End If
                  End Select
                  GoTo WIPEOUT
               End If
               
               
               If cmd$ = "debug" Then
                  Debug.Print w$(1)
               End If
               
               If IsString(cmd$) = 1 Then
                  we$ = ""
                  For t = 1 To xc
                     we$ = we$ & ParseParam(w$(t))
                  Next t
                  SetString cmd$, we$
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "proc" Then
                  For t = pc To lc
                     If LTrim$(Left$(code$(t), 7)) = "endproc" Then pc = t + 1: Exit For
                  Next t
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "call" Or cmd$ = "gosub" Then
                  For t = 0 To 10
                     If gs(t) = 0 Then
                        For gg = 0 To lc
                           'Debug.Print ">" & w$(1) & "<"
                           If LTrim$(code$(gg)) = "proc " & w$(1) Then
                              'Debug.Print "FOUND PROC"
                              gs(t) = pc: pc = gg
                              GoTo DONE_CALL_STACK
                           End If
                        Next gg
                     End If
                  Next t
DONE_CALL_STACK:
               GoTo WIPEOUT
               End If
               
               If cmd$ = "endproc" Then
                  For t = 10 To 0 Step -1
                     If gs(t) <> 0 Then
                        pc = gs(t)
                        gs(t) = 0
                        Exit For
                     End If
                  Next t
               GoTo WIPEOUT
               End If
               
               
               If cmd$ = "newint" Then
                  w$(2) = ParseParam(w$(2))
                  x = Val(w$(2))
                  CreateInteger w$(1), x
               End If
               
               If cmd$ = "redefint" Then
                  If IsInteger(w$(1)) = 0 Then
                     CreateInteger w$(1), Val(ParseParam(w$(2)))
                  Else
                     If w$(2) <> "" Then
                        SetInteger w$(1), Val(ParseParam(w$(2)))
                     End If
                  End If
                  GoTo WIPEOUT
               End If
               

               If cmd$ = "redefstr" Then
                  If IsString(w$(1)) = 0 Then
                     CreateString w$(1), ParseParam(w$(2))
                  Else
                     If w$(2) <> "" Then
                      SetString w$(1), ParseParam(w$(2))
                     End If
                  End If
               GoTo WIPEOUT
               End If
               
               If cmd$ = "strcopy" Then
                  ' strcopy dest src
                  SetString w$(1), ParseParam(w$(2))
               GoTo WIPEOUT
               End If
                  
               If cmd$ = "newstr" Then
                  w$(2) = ParseParam(w$(2))
                  CreateString w$(1), w$(2)
               GoTo WIPEOUT
               End If
                  
               If cmd$ = "includefile" Then
                  w$(1) = ParseParam(w$(1))
                  If Exists(GetDirectory(filename$) & w$(1)) Then
                     w$(1) = GetDirectory(filename$) & w$(1)
                  End If
                  
                  If Exists(w$(1)) = 0 Then
                     WXB sck, "SCRIPT ERROR: File Not Found On Line " & Trim$(Str$(pc))
                     sx(sck).Reqok = True
                     Exit Sub
                  End If
                  g$ = GetDirectory(w$(1))
                  If Exists(g$ & Longbow.SecurityFile) = 0 Then
                     WXB sck, "SCRIPT ERROR: No Directory Access " & Trim$(Str$(pc))
                     sx(sck).Reqok = True
                     Exit Sub
                  End If
                  DIR_READ = 1
                  Open g$ & Longbow.SecurityFile For Input As #4
                  Do Until EOF(4)
                     Line Input #4, f$
                     If LCase$(f$) = "read=no" Then DIR_READ = 0
                     If LCase$(f$) = "secure=yes" Then DIR_READ = 0
                  Loop
                  Close 4
                  If SFA = 1 Then DIR_READ = 1
                  If DIR_READ = 0 Then
                     WXB sck, "SCRIPT ERROR: Access Denied On Line " & Trim$(Str$(pc))
                     sx(sck).Reqok = True
                     Exit Sub
                  End If
                  Open w$(1) For Binary As #33
                     ff = LOF(33)
                     jh$ = Space$(ff)
                     Get #33, , jh$
                  Close 33
                  WXB sck, jh$
                  jh$ = ""
                  GoTo WIPEOUT
               End If
                  

               
               
               If cmd$ = "goto" Then
                  For t = 0 To lc
                     If LTrim$(code$(t)) = w$(1) & ">" Then pc = t: Exit For
                  Next t
               GoTo WIPEOUT
               End If
               

               'Debug.Print ">" & cmd$ & "<"
               
               If cmd$ = "delint" Then
                  DeleteInteger w$(1)
                  GoTo WIPEOUT
               End If
               
               
               If cmd$ = "delstr" Then
                  DeleteString w$(1)
                  GoTo WIPEOUT
               End If
               
               If cmd$ = "else" Then
                  nebo = 0
                  For t = pc To lc
                     If LTrim$(Left$(code$(t), 2)) = "if" Then nebo = nebo + 1
                     If LTrim$(Left$(code$(t), 5)) = "endif" Then nebo = nebo - 1
                     'Debug.Print nebo
                     If LTrim$(Left$(code$(t), 5)) = "endif" And nebo = -1 Then pc = t: Exit For
                  Next t
               GoTo WIPEOUT
               End If
               
               If cmd$ = "end" Then scr = False
               
               If cmd$ = "print" Then
                  ptemp$ = ""
                  For t = 1 To xc
                     ptemp$ = ptemp$ & ParseParam(w$(t))
                  Next t
                  'Debug.Print ptemp$
                  WXB sck, ptemp$
                  ptemp$ = ""
               GoTo WIPEOUT
               End If

               If cmd$ = "if" Then
                    'Debug.Print "IF FUNCTION CALLED"
                  
                  cvalue = 0
                  pre_op$ = ""
                  total_if_params = xc
                    'Debug.Print xc
                  For if_loop = 1 To xc
                     c_check$ = w$(if_loop)
                     'Debug.Print c_check$
                     If c_check$ = "&&" Then pre_op$ = "&&"
                     If c_check$ = "||" Then pre_op$ = "||"
                     If Left$(c_check$, 1) = "(" Then
                        ' If Segment
                        '(name$(hello) == "PASS$")
                        c_check$ = Mid$(c_check$, 2, Len(c_check$) - 2)
                        'name$(hello) == "PASS$"
                        s1 = Len(c_check$)
                        Erase p$()
                        iq = False: cc = 0
                        For tt = 1 To s1
                           hl$ = Mid$(c_check$, tt, 1)
                           
                           If hl$ = "(" And iq = False Then ik = ik + 1
                           If hl$ = ")" And iq = False Then ik = ik - 1
                                                      
                           If hl$ = " " And iq = False And ik = 0 Then cc = cc + 1
                           
                           If hl$ = "'" Then iq = Not iq
                                                      
                           If iq = False And ik = 0 Then p$(cc) = p$(cc) & Trim$(hl$)
                           If iq = False And ik <> 0 Then p$(cc) = p$(cc) & hl$
                           
                           If iq = True Then p$(cc) = p$(cc) & hl$
                        Next tt
                         'Debug.Print "x1"
                         'WXB sck, Str$(cc)
                         'Debug.Print "x2"
                         'WXB sck, "--" & vbCrLf
                         'Debug.Print "x3"
                        'For t = 0 To cc
                           'WXB sck, p$(t) & vbCrLf
                        'Next t
                        'Debug.Print "1", p$(0)
                        'Debug.Print "2", p$(2)
                        c0x$ = p$(0)
                        c2x$ = p$(2)
                        p$(0) = ParseParam(p$(0))
                        p$(2) = ParseParam(p$(2))
                        'Debug.Print "1", p$(0)
                        'Debug.Print "2", p$(2)
                        kk = 0
                        If (p$(0) <> "" And p$(2) <> "") And Val(p$(0)) = Val(Str$(Val(p$(0)))) Then
                            ISTHISINT = 1
                        Else
                            ISTHISINT = 0
                        End If
                        
                        If IsString(c0x$) = 1 Then ISTHISINT = 0
                        If IsString(c2x$) = 1 Then ISTHISINT = 0
                        If IsInteger(c0x$) = 1 Then ISTHISINT = 1
                        If IsInteger(c2x$) = 1 Then ISTHISINT = 1
                        
                        'Debug.Print ISTHISINT
                        
                        Select Case ISTHISINT
                                
                                Case 0
                                    Select Case p$(1)
                                       Case ">"
                                          If p$(0) > p$(2) Then kk = 1
                                       Case "<"
                                          If p$(0) < p$(2) Then kk = 1
                                       Case "=="
                                          If p$(0) = p$(2) Then kk = 1
                                       Case "!="
                                          If p$(0) <> p$(2) Then kk = 1
                                       Case ">="
                                          If p$(0) >= p$(2) Then kk = 1
                                       Case "<="
                                          If p$(0) <= p$(2) Then kk = 1
                                    End Select
                                
                                Case 1
                                        p1 = Val(p$(0))
                                        p2 = Val(p$(2))
                                    Select Case p$(1)
                                       Case ">"
                                          If p1 > p2 Then kk = 1
                                       Case "<"
                                          If p1 < p2 Then kk = 1
                                       Case "=="
                                          If p1 = p2 Then kk = 1
                                       Case "!="
                                          If p1 <> p2 Then kk = 1
                                       Case ">="
                                          If p1 >= p2 Then kk = 1
                                       Case "<="
                                          If p1 <= p2 Then kk = 1
                                    End Select
                        End Select

                       ' Debug.Print p$(0)
                       ' Debug.Print p$(2)
                       ' Debug.Print kk
                        
                        If kk = 0 Then
                           If pre_op$ = "||" Or pre_op$ = "OR" Then
                              If cvalue = 1 Then cvalue = 1
                              If cvalue = 0 Then cvalue = 0
                           End If
                           If pre_op$ = "&&" Or pre_op$ = "AND" Then
                              If cvalue = 0 Then cvalue = 0
                              If cvalue = 1 Then cvalue = 0
                           End If
                        End If
                        If kk = 1 Then
                           If pre_op$ = "||" Or pre_op$ = "OR" Then
                              If cvalue = 0 Then cvalue = 1
                              If cvalue = 1 Then cvalue = 1
                           End If
                           If pre_op$ = "&&" Or pre_op$ = "AND" Then
                              If cvalue = 1 Then cvalue = 1
                              If cvalue = 0 Then cvalue = 0
                           End If
                        End If
                        
                       ' Debug.Print kk
                        
                        If pre_op$ <> "&&" And pre_op$ <> "||" And pre_op$ <> "AND" And pre_op$ <> "OR" Then
                              cvalue = kk
                        End If
                              
                     End If
                  Next if_loop
                  
                  it = cvalue
                  If it = 0 Then
                     cando = 0
                     For t = pc + 1 To lc
                        If Left$(LTrim$(code$(t)), 2) = "if" Then cando = cando + 1
                        If Left$(LTrim$(code$(t)), 5) = "endif" Then cando = cando - 1
                        'Debug.Print Left$(LTrim$(code$(t)), 5); cando
                        If cando < 1 And Left$(LTrim$(code$(t)), 4) = "else" Then pc = t: Exit For
                        If cando < 1 And Left$(LTrim$(code$(t)), 5) = "endif" Then pc = t:  Exit For
                     Next t
                  End If
                  'WXB sck, "I HAD BEEN CHECKED"
               GoTo WIPEOUT
               
               End If
WIPEOUT:
      End If
                  
SkipMe:
      
      pc = pc + 1
      If pc > lc Then scr = False
   
   Loop

      If sx(sck).Buffer = "" Then sx(sck).Buffer = vbCrLf & vbCrLf


   Exit Sub
   
Execute_Error:
   WXB sck, "Error Occured In Script Processing On Line " & Trim$(Str$(pc))
   sx(sck).Reqok = True
End Sub

Public Function ParseParam(para As String) As String
   'Debug.Print "PARA=" & para$
   If para$ = "&" Or para$ = "" Or para$ = " &" Then Exit Function
   Dim d$(200)
   Dim iq As Boolean
   If para$ = "lbCrLf" Then ParseParam = vbCrLf: Exit Function
   
   ' Is String?
   
   If IsString(para$) = 1 Then
      ParseParam = GetString(para$)
      Exit Function
   End If
   
   ' Is A Number?
   If IsNumeric(para$) Then
      ParseParam = para$
      Exit Function
   End If
   
   ' Is Bracketed?
   '(3 * 10)
   If Left$(para$, 1) = "(" And Right$(para$, 1) = ")" Then
      ' This parameter is bracketed, we need to resolve whatevers inside it
      f1$ = Mid$(para$, 2, Len(para$) - 2)
      '3 * 10
      
      ' Now we parse the whole thing to get the sub params
      r = Len(f1$)
      jk = 0: wc = 0
      Erase d$()
      For g = 1 To r
         k$ = Mid$(f1$, g, 1)
         If k$ = "'" Then iq = Not iq
         If k$ = "(" And iq = False Then jk = jk + 1
         If k$ = ")" And iq = False Then jk = jk - 1
         If k$ = " " And iq = False And jk = 0 Then wc = wc + 1
            d$(wc) = d$(wc) & k$
      Next g
      
      For g = 0 To wc
         welly = 1
         d$(g) = Trim$(d$(g))
         If d$(g) = "+" Or d$(g) = "-" Or d$(g) = "/" Or d$(g) = "*" Then
            xc$ = d$(g): welly = 0
         End If
         If welly = 1 Then
            If xc$ = "+" Then tempo = tempo + Val(ParseParam(d$(g)))
            If xc$ = "-" Then tempo = tempo - Val(ParseParam(d$(g)))
            If xc$ = "/" Then tempo = tempo / Val(ParseParam(d$(g)))
            If xc$ = "*" Then tempo = tempo * Val(ParseParam(d$(g)))
         End If
         If xc$ = "" Then tempo = Val(ParseParam(d$(g)))
      Next g
      ParseParam = Trim$(Str$(tempo))
      Exit Function
   End If
   
   
   ' Is Integer?
   If IsInteger(para$) = 1 Then
      ParseParam = Trim$(Str$(GetInteger(para$)))
      Exit Function
   End If
      
   
   ' Is Server Const?
   If GetConst(para$) <> "" Then
      ParseParam = GetConst(para$)
      Exit Function
   End If
   
   ' Is Normal Text?
   dx = Len(para$)
   iq = False: jk = 0
   For t = 1 To dx
      If Mid$(para$, t, 1) = "(" And iq = False Then jk = jk + 1
      If Mid$(para$, t, 1) = ")" And iq = False Then jk = jk - 1
      If Mid$(para$, t, 1) = "'" Then iq = Not iq
      If Mid$(para$, t, 1) = "&" And iq = False And jk = 0 Then MULTI_PARAM = 1: GoTo MULTI
   Next t
   GoTo NOMULTI
MULTI:
   'Debug.Print "this para$=" & para$ & "="
   Erase d$()
   wc = 0
   iq = False: jk = 0
   If MULTI_PARAM = 1 Then
      'Debug.Print "MULTIPARAM"
      ' MultiParameter
      r = Len(para$)
      For g = 1 To r
         k$ = Mid$(para$, g, 1)
         If k$ = "'" Then iq = Not iq
         If k$ = "(" And iq = False Then jk = jk + 1
         If k$ = ")" And iq = False Then jk = jk - 1
         If k$ = " " And iq = False And jk = 0 Then wc = wc + 1
          d$(wc) = d$(wc) & k$
      Next g
   
      For g = 0 To wc
         
         d$(g) = Trim$(d$(g))
         'Debug.Print "tempo$=" & d$(g) & "="
         tempoo$ = tempoo$ & ParseParam(d$(g))
      Next g
      ParseParam = tempoo$
      Exit Function
   End If
   
   
NOMULTI:
   
   
   If Left$(para$, 1) = "'" Then
      ParseParam = ReplaceStr(Mid$(para$, 2, Len(para$) - 2), "\n", vbCrLf)
      Exit Function
   End If
   
   dd = Len(para$)
   For cc = 1 To dd
      f$ = Mid$(para$, cc, 1)
      If f$ = "'" Then Exit Function
      If f$ <> "(" Then GoTo BAILMEOUT
   Next cc
   
   Exit Function
   
BAILMEOUT:
   ' ucase(lcase(forname) & ' ' & lcase(surname))
   '      a                                    cc
   
   a = InStr(para$, "(")
   b = Len(para$)
   For c = b To 1 Step -1
      If Mid$(para$, c, 1) = ")" Then x = x + 1
      If Mid$(para$, c, 1) <> ")" Then GoTo ALLOUT
      
   Next c
   
ALLOUT:
   
   x = x - 1
   
   z$ = Mid$(para$, a + 1, Len(para$) - a - x)
   
   'Debug.Print "z$=" & z$
   'ucase(lcase(ucase('name')))
   'ucase(lcase('NAME'))
   'ucase('name')
   
   If Left$(para$, 5) = "lcase" Then
      ParseParam = LCase(ParseParam(Mid$(para$, 7, Len(para$) - 7)))
      Exit Function
   End If
   
   If Left$(para$, 5) = "ucase" Then
      ParseParam = UCase(ParseParam(Mid$(para$, 7, Len(para$) - 7)))
      Exit Function
   End If
   
    If Left$(para$, 5) = "ltrim" Then
        ParseParam = LTrim(ParseParam(Mid$(para$, 7, Len(para$) - 7)))
        Exit Function
    End If
    
    If Left$(para$, 5) = "rtrim" Then
        ParseParam = RTrim(ParseParam(Mid$(para$, 7, Len(para$) - 7)))
        Exit Function
    End If
    
    If Left$(para$, 4) = "trim" Then
        ParseParam = Trim(ParseParam(Mid$(para$, 6, Len(para$) - 6)))
    End If
   
   If Left$(para$, 3) = "rnd" Then
      d1 = Val(ParseParam(GetParam(para$, 0)))
      d2 = Val(ParseParam(GetParam(para$, 1)))
      j = (Rnd(1) * (d2 - d1)) + d1
      ParseParam = Trim$(Str$(j))
      Exit Function
   End If
   
   If Left$(para$, 3) = "mid" Then
      rl$ = ParseParam(GetParam(para$, 0))
      d1 = Val(ParseParam(GetParam(para$, 1)))
      d2 = Val(ParseParam(GetParam(para$, 2)))
      'Debug.Print , rl$, d1, d2
      ParseParam = Mid$(ParseParam(rl$), d1, d2)
      Exit Function
   End If
   
    If Left$(para$, 3) = "chr" Then
        d1 = Val(ParseParam(GetParam(para$, 0)))
        ParseParam = Chr$(d1)
    End If
   
    If Left$(para$, 3) = "asc" Then
        ds2$ = ParseParam(GetParam(para$, 0))
        'Debug.Print ParseParam(GetParam(para$, 0))
        'Debug.Print GetParam(para$, 0)
        'Debug.Print ">"; ds2$
        'Debug.Print Str$(Asc(ds2$))
        ParseParam = Trim$(Str$(Asc(ds2$)))
    End If
    
    If Left$(para$, 3) = "eof" Then
        
        ds2$ = ParseParam(GetParam(para$, 0))
        'Debug.Print "EOF:" & ds2$
        ff$ = Trim$(Str$(EOF(file(Val(ds2$)))))
        'Debug.Print "EOF2:" & ff$
        If ff$ = "True" Then ff$ = "-1"
        If ff$ = "False" Then ff$ = "0"
        ParseParam = ff$
    End If
End Function


Public Function IsInteger(pint_name As String) As Integer
   For t = 0 To 1000
      If pint_name$ = inn$(t) & "%" Then IsInteger = 1: Exit Function
   Next t
   IsInteger = 0
End Function

Public Sub CreateInteger(int_name As String, Optional int_data As Integer)
   For t = 0 To 1000
      If inn$(t) = "" Then inn$(t) = int_name$: ind(t) = int_data: Exit Sub
      If t = 500 Then DoEvents
   Next t
End Sub

Public Function GetInteger(int_name As String) As Integer
   For t = 0 To 1000
      If inn$(t) & "%" = int_name$ Then GetInteger = ind(t): Exit Function
      If t = 500 Then DoEvents
   Next t
End Function

Public Sub DeleteInteger(int_name As String)
   For t = 0 To 1000
      If inn$(t) & "%" = int_name$ Then inn$(t) = "": ind(t) = 0: Exit Sub
      If t = 500 Then DoEvents
   Next t
End Sub

Public Sub SetInteger(int_name As String, int_data As Integer)
   For t = 0 To 1000
      If inn$(t) & "%" = int_name$ Then ind(t) = int_data: Exit Sub
      If t = 500 Then DoEvents
   Next t
End Sub

Public Function IsString(pstring_name As String) As Integer
   For t = 0 To 1000
      If pstring_name$ = stn$(t) & "$" Then IsString = 1: Exit Function
   Next t
   IsString = 0
End Function



Public Sub CreateString(string_name As String, Optional string_data As String)
   If Right$(string_name$, 1) = "$" Then string_name$ = Left$(string_name$, Len(string_name$) - 1)
   For t = 0 To 1000
      If stn$(t) = "" Then stn$(t) = string_name: std$(t) = string_data: Exit Sub
      If t = 500 Then DoEvents
   Next t
End Sub

Public Function GetString(string_name As String) As String
   For t = 0 To 1000
      If stn$(t) & "$" = string_name$ Then GetString = std$(t): Exit Function
      If t = 500 Then DoEvents
   Next t
End Function

Public Sub DeleteString(string_name As String)
   For t = 0 To 1000
      If stn$(t) & "$" = string_name$ Then stn$(t) = "": std$(t) = "": Exit Sub
      If t = 500 Then DoEvents
   Next t
End Sub

Public Sub SetString(string_name As String, string_data As String)
   'Debug.Print "SetString(" & string_name$ & "," & string_data$ & ")"
   For t = 0 To 1000
      If stn$(t) & "$" = string_name$ Then std$(t) = string_data$: Exit Sub
      If t = 500 Then DoEvents
   Next t
End Sub

Public Function GetConst(poss_const As String) As String
   Select Case poss_const
      Case "#time"
         r$ = Time$
      Case "#date"
         r$ = Date$
      Case "#timer"
         r$ = Trim$(Str$(Timer))
      Case "#svname"
         r$ = Longbow.ServerName
      Case "#svadmin"
         r$ = Longbow.ServerAdmin
   End Select
   GetConst = r$
End Function

Public Function GetParam(ByVal comd As String, paramnum As Integer) As String
Dim d$(200)
Dim iq As Boolean
dd = InStr(comd$, "(")
comd$ = Right$(comd$, Len(comd$) - dd)
comd$ = Left$(comd$, Len(comd$) - 1)
a = Len(comd$)
iq = False: ij = 0
For b = 1 To a
   c$ = Mid$(comd$, b, 1)
   If c$ = "(" And iq = False Then ij = ij + 1
   If c$ = ")" And iq = False Then ij = ij - 1
   If c$ = "'" Then iq = Not iq
   If c$ = "," And iq = False And ij = 0 Then wc = wc + 1: c$ = ""
   d$(wc) = d$(wc) & c$
Next b
'Debug.Print d$(0)
'Debug.Print d$(1)
'Debug.Print d$(2)
GetParam = d$(paramnum)

End Function

Public Sub SecureCopy(src As String, dest As String)
   On Error GoTo SECCOPYERROR
   FileCopy src, dest
   Exit Sub
SECCOPYERROR:
End Sub
