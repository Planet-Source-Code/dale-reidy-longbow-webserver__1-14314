VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LongBow Server"
   ClientHeight    =   3315
   ClientLeft      =   2145
   ClientTop       =   3780
   ClientWidth     =   5580
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5580
   Begin VB.CommandButton Command3 
      Caption         =   "Read Log"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2205
      TabIndex        =   5
      Top             =   2835
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1155
      TabIndex        =   1
      Top             =   2835
      Width           =   1065
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   120
      Top             =   105
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   2745
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5580
   End
   Begin VB.Timer sxu 
      Interval        =   10
      Left            =   105
      Top             =   105
   End
   Begin VB.Timer sxt 
      Interval        =   10
      Left            =   105
      Top             =   105
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   2835
      Width           =   1065
   End
   Begin VB.Label lblReq 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   225
      Left            =   4560
      TabIndex        =   4
      Top             =   2940
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Requests"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   225
      Left            =   3570
      TabIndex        =   3
      Top             =   2940
      Width           =   960
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 StopServer
End Sub

Private Sub Command2_Click()
   StartServer
End Sub

Private Sub Command3_Click()
 Shell "notepad.exe " & ServerLogFile, vbNormalFocus
End Sub

Private Sub Form_Load()
   InitServer
   WLog "Server Started"
   
   Cout "LongBow Server 1.0b" & vbCrLf
   Cout "-------------------" & vbCrLf
   Cout "Listening On Port " & Trim$(Str$(Longbow.ListenPort)) & vbCrLf
   Cout "Server Config Loaded" & vbCrLf
   Cout "LogFile:" & ServerLogFile & vbCrLf
   Cout "DBConFile:" & Longbow.DatabaseCfg & vbCrLf
End Sub

Private Sub Form_Unload(Cancel As Integer)
   CloseServer
   End
End Sub

Private Sub sxt_Timer()
   lblReq.Caption = Trim$(Str$(NumReq))
   Dim t As Integer
   For t = 1 To Longbow.MaxSocks Step 2
      
      DoEvents

      If sx(t).Header <> "" And sx(t).Reqok = False And sx(t).Buffer = "" And ws(t).State = sckConnected Then

         ProcessHeader t
      End If

         sx(t).TimeAlive = sx(t).TimeAlive + 1


         If sx(t).TimeAlive > Longbow.TimeOut Then
            ws(t).Close
            sx(t).Buffer = ""
            sx(t).Header = ""
            sx(t).Reqok = False
            sx(t).TimeAlive = 0
         End If

      If sx(t).Reqok = True And ws(t).State <> sckConnected Then
            ws(t).Close
            sx(t).Buffer = ""
            sx(t).Header = ""
            sx(t).Reqok = False
            sx(t).TimeAlive = 0
            ws(t).Tag = ""
      End If

      If sx(t).Reqok = True And ws(t).State = sckConnected Then

         a = Len(sx(t).Buffer)

         'Debug.Print a

         If a = 0 And frmmain.ws(t).Tag = "LASTPACKET" Then
            ws(t).Close

            sx(t).Buffer = ""
            sx(t).Header = ""
            sx(t).Reqok = False
            sx(t).TimeAlive = 0
            ws(t).Tag = ""
            GoTo RABIDO
         End If
         'If a = 0 Then GoTo RABIDO
         If a > 3000 Then g = 3000 Else g = a: ws(t).Tag = "LASTSEND"
         r$ = Left$(sx(t).Buffer, g)
         sx(t).Buffer = Right$(sx(t).Buffer, Len(sx(t).Buffer) - g)

         ws(t).SendData r$

         sx(t).TimeAlive = 0
      End If
RABIDO:
'         ws(t).SendData sx(t).Buffer
'         sx(t).Reqok = False
'         sx(t).Buffer = ""


      If sx(t).Reqok = True And ws(t).State <> sckConnected Then
         ws(t).Close
         sx(t).Buffer = ""
         sx(t).Header = ""
         sx(t).Reqok = False
         sx(t).TimeAlive = 0
         ws(t).Tag = ""
      End If

      
   Next t
      
End Sub

Private Sub sxu_Timer()
   Dim t As Integer
   For t = 0 To Longbow.MaxSocks Step 2
      
      If t <> 0 Then
      
      
      
      DoEvents

      If sx(t).Header <> "" And sx(t).Reqok = False And sx(t).Buffer = "" And ws(t).State = sckConnected Then

         ProcessHeader t
      End If

         sx(t).TimeAlive = sx(t).TimeAlive + 1
      If ws(t).State = sckConnected Then Debug.Print sx(t).TimeAlive

         If sx(t).TimeAlive > Longbow.TimeOut Then
            ws(t).Close
            sx(t).Buffer = ""
            sx(t).Header = ""
            sx(t).Reqok = False
            sx(t).TimeAlive = 0
         End If

      If sx(t).Reqok = True And ws(t).State <> sckConnected Then
            ws(t).Close
            sx(t).Buffer = ""
            sx(t).Header = ""
            sx(t).Reqok = False
            sx(t).TimeAlive = 0
            ws(t).Tag = ""
      End If

      If sx(t).Reqok = True And ws(t).State = sckConnected Then

         a = Len(sx(t).Buffer)

         'Debug.Print a

         If a = 0 And frmmain.ws(t).Tag = "LASTPACKET" Then
            ws(t).Close

            sx(t).Buffer = ""
            sx(t).Header = ""
            sx(t).Reqok = False
            sx(t).TimeAlive = 0
            ws(t).Tag = ""
            GoTo RABIDO
         End If
         'If a = 0 Then GoTo RABIDO
         If a > 3000 Then g = 3000 Else g = a: ws(t).Tag = "LASTSEND"
         r$ = Left$(sx(t).Buffer, g)
         sx(t).Buffer = Right$(sx(t).Buffer, Len(sx(t).Buffer) - g)

         ws(t).SendData r$

         sx(t).TimeAlive = 0
      End If
RABIDO:
'         ws(t).SendData sx(t).Buffer
'         sx(t).Reqok = False
'         sx(t).Buffer = ""


      If sx(t).Reqok = True And ws(t).State <> sckConnected Then
         ws(t).Close
         sx(t).Buffer = ""
         sx(t).Header = ""
         sx(t).Reqok = False
         sx(t).TimeAlive = 0
         ws(t).Tag = ""
      End If


      End If
      
   Next t
      
End Sub

Private Sub ws_Close(Index As Integer)
   sx(Index).Buffer = ""
   sx(Index).Header = ""
   sx(Index).Reqok = False
   sx(Index).TimeAlive = 0
   ws(t).Tag = ""
End Sub

Private Sub ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
   Dim t As Integer
   For t = 1 To Longbow.MaxSocks
      If ws(t).State = sckClosing Then ws(t).Close
      If ws(t).State = sckClosed Then
         ws(t).Accept requestID
         sx(t).Buffer = ""
         sx(t).Header = ""
         sx(t).Reqok = False
         sx(t).TimeAlive = 0
         ws(t).Tag = ""
         If IPBanned(t) = 1 Then
            WriteHTTP t, 403, "-"
            sx(t).Reqok = True
            sx(t).Header = "HAS BEEN BANNED"
            Exit Sub
         End If
         Exit Sub
      End If
   Next t
End Sub

Private Sub ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
   On Error GoTo WSDAO
   ws(Index).GetData sx(Index).Header
   'ProcessHeader Index
   Exit Sub
WSDAO:
   ws(Index).Close
         sx(t).Buffer = ""
         sx(t).Header = ""
         sx(t).Reqok = False
         sx(t).TimeAlive = 0
         ws(t).Tag = ""
End Sub

Private Sub ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   ws(Index).Close
   ' If the listening socket closes, we're buggered, so open it up again, probably causing another error
   ' and the program to crash, but hey, windows crashes, so why can't this?
   If Index = 0 Then ws(0).Listen
End Sub

Private Sub ws_SendComplete(Index As Integer)
If ws(Index).Tag = "LASTSEND" Then
   ws(Index).Tag = "LASTPACKET"
End If
End Sub
