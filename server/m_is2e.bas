Attribute VB_Name = "m_is2e"
' IS23 Encryption

Public Function Encrypt(text As String, password As String) As String
   Dim r$
   r$ = password$
   
   If Len(r$) = 0 Then Exit Function
   a = Len(text$)
   b = 0
   xx = 1
   dd$ = ""
   
   For c = 1 To a
      m$ = Mid$(text$, c, 1)
      h = Asc(m$)
      If xx > Len(r$) Then xx = 1
      h = h + Asc(Mid$(r$, xx, 1)) + c
      h = h + xx
      h = h + Len(r$)
      Do Until h < 255
         h = h - 255
      Loop
      xx = xx + 1
      dd$ = dd$ & Chr$(h)
   Next c
   
   Encrypt = dd$

End Function

Public Function Decrypt(text As String, password As String) As String
   Dim r$
   r$ = password$
   
   If Len(r$) = 0 Then Exit Function
   
   a = Len(text$)
   b = 0
   xx = 1
   dd$ = ""
   
   For c = 1 To a
      m$ = Mid$(text$, c, 1)
      h = Asc(m$)
      If xx > Len(r$) Then xx = 1
      h = h - Asc(Mid$(r$, xx, 1)) - c
      h = h - xx
      h = h - Len(r$)
      Do Until h > 0
         h = h + 255
      Loop
      xx = xx + 1
      dd$ = dd$ & Chr$(h)
   Next c
   
   Decrypt = dd$
   

End Function
