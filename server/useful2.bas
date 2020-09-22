Attribute VB_Name = "m_misc"

Public Function GetTime(TimerVal As Long) As String
 ' Returns time in HH:MM:SS format from Timer value(use 0 for current time :))
 tx = Timer - TimerVal
 
 x = Int((tx / 60) / 60)
 tx = tx - (x * 60 * 60)
 y = Int(tx / 60)
 tx = tx - (y * 60)
 z = Int(tx)
 sx1$ = Trim$(Str$(x))
 sy1$ = Trim$(Str$(y))
 sz1$ = Trim$(Str$(z))
 If Len(sx1$) = 1 Then sx1$ = "0" & sx1$
 If Len(sy1$) = 1 Then sy1$ = "0" & sy1$
 If Len(sz1$) = 1 Then sz1$ = "0" & sz1$
 GetTime$ = sx1$ & ":" & sy1$ & ":" & sz1$
End Function

Public Function ReplBack(txt As String) As String
 Dim t As Integer
 Dim u As Integer
 Dim temp As String
 Dim rv As String * 1
 t = Len(txt$)
 For u = 1 To t
  rv$ = Mid$(txt$, u, 1)
  If rv$ = "/" Then rv$ = "\"
  temp$ = temp$ & rv$
 Next u
 ReplBack = temp$
End Function

Public Function ReplaceStr(ByVal strMain As String, strFind As String, strReplace As String) As String

    Dim lngSpot As Long, lngNewSpot As Long, strLeft As String
    Dim strRight As String, strNew As String
    lngSpot& = InStr(LCase(strMain$), LCase(strFind$))
    lngNewSpot& = lngSpot&
    Do
        If lngNewSpot& > 0& Then
            strLeft$ = Left(strMain$, lngNewSpot& - 1)
            If lngSpot& + Len(strFind$) <= Len(strMain$) Then
                strRight$ = Right(strMain$, Len(strMain$) - lngNewSpot& - Len(strFind$) + 1)
            Else
                strRight = ""
            End If
            strNew$ = strLeft$ & strReplace$ & strRight$
            strMain$ = strNew$
        Else
            strNew$ = strMain$
        End If
        lngSpot& = lngNewSpot& + Len(strReplace$)
        If lngSpot& > 0 Then
            lngNewSpot& = InStr(lngSpot&, LCase(strMain$), LCase(strFind$))
        End If
    Loop Until lngNewSpot& < 1
    ReplaceStr$ = strNew$
End Function

Public Function Exists(txt As String) As Integer
 On Error GoTo noexist
  Open txt For Input As #12
  Exists = 1
  Close 12
  Exit Function
noexist:
  Close 12
  Exists = 0
End Function

Public Function RidFormatting(para As String) As String
      para = ReplaceStr(para, "+", " ")
      para = ReplaceStr(para, "%0D%0A", "<br>")
      para = ReplaceStr(para, "%21", "!")
      para = ReplaceStr(para, "%22", "&quot;")
      para = ReplaceStr(para, "%20", " ")
      para = ReplaceStr(para, "%A7", "§")
      para = ReplaceStr(para, "%24", "$")
      para = ReplaceStr(para, "%25", "%")
      para = ReplaceStr(para, "%26", "&")
      para = ReplaceStr(para, "%2F", "/")
      para = ReplaceStr(para, "%28", "(")
      para = ReplaceStr(para, "%29", ")")
      para = ReplaceStr(para, "%3D", "=")
      para = ReplaceStr(para, "%3F", "?")
      para = ReplaceStr(para, "%B2", "²")
      para = ReplaceStr(para, "%B3", "³")
      para = ReplaceStr(para, "%7B", "{")
      para = ReplaceStr(para, "%5B", "[")
      para = ReplaceStr(para, "%5D", "]")
      para = ReplaceStr(para, "%7D", "}")
      para = ReplaceStr(para, "%5C", "\")
      para = ReplaceStr(para, "%DF", "ß")
      para = ReplaceStr(para, "%23", "#")
      para = ReplaceStr(para, "%27", "'")
      para = ReplaceStr(para, "%3A", ":")
      para = ReplaceStr(para, "%2C", ",")
      para = ReplaceStr(para, "%3B", ";")
      para = ReplaceStr(para, "%60", "`")
      para = ReplaceStr(para, "%7E", "~")
      para = ReplaceStr(para, "%2B", "+")
      para = ReplaceStr(para, "%B4", "´")
      RidFormatting = para
End Function

Public Function RevStr(txt As String) As String
 ' Reverses the contents of a string
 Dim a, b, c$
 a = Len(txt)
 For b = a To 1 Step -1
  c$ = c$ & Mid$(txt, b, 1)
 Next b
 RevStr = c$
End Function

Public Function GetDirectory(filename As String) As String
 If IsDir(filename) Then GetDirectory = filename: Exit Function
 If Right$(filename, 1) = "\" Then GetDirectory = filename: Exit Function
 If Right$(filename, 1) = "/" Then GetDirectory = filename: Exit Function
 
 t = Len(filename)
 For a = t To 1 Step -1
 If Mid$(filename, a, 1) = "\" Or Mid$(filename, a, 1) = "/" Then SE = a: GoTo okok
 Next a
 GetDirectory = filename
 Exit Function
okok:
 GetDirectory = Left$(filename, SE)
End Function

Public Function IsDir(dire As String) As Integer
 On Error GoTo BADERROR
 c$ = CurDir$
 ChDir dire
 ChDir c$
 IsDir = 1
 Exit Function
BADERROR:
 ChDir c$
 IsDir = 0
End Function

Public Function GetFilename(fullpath As String) As String
   r = Len(fullpath$)
   For a = r To 1 Step -1
      If Mid$(fullpath$, a, 1) = "\" Then
         GetFilename = Right$(fullpath$, r - a)
         Exit Function
      End If
   Next a
   GetFilename = fullpath$
End Function
