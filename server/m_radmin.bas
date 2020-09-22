Attribute VB_Name = "m_radmin"
Public Sub ExecRadminCommand(rac As String, lparam As String, hparam As String)
   ' Execute Remote Admin Commands
   Select Case rac$
      Case "clearlogs"
         ' Clear all the logfiles
         Kill Longbow.LogLoc & "*.log"
      Case "banip"
         AppendIPBanIP lparam$
      Case "unbanip"
         m_main.RemoveIPBan lparam$
   End Select
End Sub

Public Function NumberOfRequests() As String
   NumberOfRequests = Trim$(Str$(ReqNum))
End Function
