VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C32049A2-E8D3-11D4-9D82-00D0B73B61A4}#2.1#0"; "cwSvc.ocx"
Begin VB.Form frmServer 
   Caption         =   "Form1"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin Service.Service SERVICE 
      Left            =   1800
      Top             =   480
      _Version        =   131073
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      ServiceName     =   "WinName"
      StartMode       =   2
   End
   Begin VB.Timer tmrSelfCheck 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1320
      Top             =   480
   End
   Begin MSWinsockLib.Winsock sender 
      Index           =   0
      Left            =   360
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock listener 
      Left            =   840
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin VB.Frame Frame1 
      Caption         =   "Most Recent Request"
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text2 
         Height          =   2295
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const HomePage As String = "index.html"
Dim sending() As Boolean, stopSend() As Boolean, freeConn() As Boolean
Private Sub Form_Load()
    Me.Hide
    ' Prepare Service Control
    With Service
        .ControlsAccepted = svcCtrlPauseContinue
        .DisplayName = "ServiceWebDemo"
        .Interactive = False
        .ServiceName = "ServiceWebDemo"
        .StartMode = svcStartAutomatic
    End With
    
    Select Case Trim(LCase(Command$))
        Case "-i": Call InstallService: Exit Sub
        Case "-u": Call UninstallService: Exit Sub
    End Select
    
    Service.StartService
    tmrSelfCheck.Enabled = True
End Sub
Private Sub listener_Close()
    ' Reset connection
    listener.Close: DoEvents: listener.Listen
End Sub
Private Sub listener_ConnectionRequest(ByVal requestID As Long)
    Dim i As Integer
    For i = 0 To 255
        If freeConn(i) Then
            freeConn(i) = False
            sender(i).Close
            sender(i).Accept requestID
            List1.AddItem sender(i).RemoteHostIP
            Me.Caption = List1.ListCount & " Clients Connected."
            Exit For
        End If
    Next
End Sub
Private Sub sender_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim dat As String, file As String, start As Integer, start2 As Integer, fileNum As Integer
' get request
    sender(Index).GetData dat
' take log
    fileNum = FreeFile
    Open App.Path & "\log.txt" For Append As #fileNum
    Print #fileNum, "Client " & sender(Index).RemoteHostIP & " Request:"
    Print #fileNum, dat
    Close (fileNum)
    Text2.Text = dat
' check request, if the first 3 are "GET" then it is a request for getting a file
    If Mid(dat, 1, 3) = "GET" Then
' check position for "GET "
        start = InStr(dat, "GET ")
' check position for the end of the file name
        start2 = InStr(start + 5, dat, " ")
' get the file name
        file = Mid(dat, start + 5, start2 - (start + 4))
' trim the file name for ending space
        file = RTrim(file)
' if name is empty, it means it is something like ".../" so it will call the default file
        If file = "" Or Right(file, 1) = "/" Then
            ' We need to find the DEFAULT file
            file = file & HomePage
        End If
' send file
        sendfile App.Path & "\wwwroot\" & file, Index
    End If
End Sub
Private Sub sender_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
' stop send if experience error
    stopSend(Index) = True
End Sub
Private Sub sender_SendComplete(Index As Integer)
' set the sending variable for the sock to false when sending is done
    sending(Index) = False
End Sub

Private Sub sendfile(file As String, sock As Integer)
' substitute for invalid characters
    file = Replace(file, "/", "\")
    file = Replace(file, "%20", " ")
' set up error handler in case there's error with the file
    On Error GoTo handler
    Dim fileNum As Integer
    Dim fileBin As String
    Dim fileSize As Long
    Dim sentSize As Long
    Dim i As Integer

' prepare opening the file requested
    fileSize = FileLen(file)
    fileNum = FreeFile
    Open file For Binary As #fileNum
        
' have it send 1024 bit a time
        fileBin = Space(1024)
        
        Do
' set sending for that sock to true so that we can check its progress
            sending(sock) = True
' get the packet
            Get #fileNum, , fileBin
' calculate amount sent
            sentSize = sentSize + Len(fileBin)
' check if it will be done or not in the next packet
            If sentSize > fileSize Then
                sender(sock).SendData Mid(fileBin, 1, Len(fileBin) - (sentSize - fileSize))
            Else
                sender(sock).SendData fileBin
            End If
            Do
' wait until sending is done -- the sending variable is changed by the sock's sendcomplete event
                DoEvents
' if it is to be stopped in the middle, send it to the error handler
                If stopSend(sock) Then GoTo handler
            Loop Until sending(sock) = False
            
        DoEvents
' keep sending until the file is sent
        Loop Until EOF(fileNum)
    
' close file and free sock
    Close (fileNum)
' remove ip log
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = sender(sock).RemoteHostIP Then
            List1.RemoveItem i
            Me.Caption = List1.ListCount & " Clients Connected."
            Exit For
        End If
    Next
    sender(sock).Close
    freeConn(sock) = True
    
    Exit Sub
handler:

' tell client about the error
    If sender(sock).State = 7 Then
        sending(sock) = True
        sender(sock).SendData "Internal Error" & vbNewLine
        Do
            DoEvents
        Loop Until sending(sock) = False
    End If
    
' close and free sock
' remove ip log
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = sender(sock).RemoteHostIP Then
            List1.RemoveItem i
            Me.Caption = List1.ListCount & " Clients Connected."
            Exit For
        End If
    Next
    sender(sock).Close
    freeConn(sock) = True
    stopSend(sock) = False
End Sub





'================================================================================'
'= Service Processes:                                                           ='
'================================================================================'
Private Sub SERVICE_Continue(Success As Boolean)
    Success = PrepareServer
End Sub
Private Sub SERVICE_Control(ByVal lEvent As Long)
    '
End Sub
Private Sub SERVICE_Pause(Success As Boolean)
    Success = CloseServer
End Sub
Private Sub SERVICE_Start(Success As Boolean)
    tmrSelfCheck.Enabled = False        ' Turn off timer
    Success = PrepareServer
End Sub
Private Sub SERVICE_Stop()
    CloseServer
    Unload Me
End Sub
Private Sub tmrSelfCheck_Timer()
    ' Timer fires if Service Manager DID NOT contact us!
    tmrSelfCheck.Enabled = False
    Unload Me
End Sub
'================================================================================'
Private Sub InstallService()
    If Service.Install Then
        MsgBox "Service Installed Successfully!", vbInformation
    Else
        MsgBox "Service Failed To Install!", vbCritical
    End If
    Unload Me
End Sub
Private Sub UninstallService()
    If Service.Uninstall Then
        MsgBox "Service Uninstalled Successfully!", vbInformation
    Else
        MsgBox "Service Failed To Uninstall!", vbCritical
    End If
    Unload Me
End Sub
'================================================================================'
Private Function CloseServer() As Boolean
    listener.Close
    Dim CTL As Control
    For Each CTL In Controls
        If CTL.Name = "sender" Then
            CTL.Close
            If CTL.Index > 0 Then
                Unload CTL
            End If
        End If
    Next
    Set CTL = Nothing
End Function
Private Function PrepareServer() As Boolean
    ReDim freeConn(255): ReDim sending(255): ReDim stopSend(255)
    Dim i As Integer
    For i = 0 To 255
        If i <> 0 Then Load sender(i)
        freeConn(i) = True
        sending(i) = False
        stopSend(255) = False
    Next
    
ConnectTo80:
    ' Connect to PORT 80
    On Error GoTo BadPort80
    listener.LocalPort = 80
    listener.Listen
    PrepareServer = True
    Exit Function
    
ConnectTo8080:
    ' Connect to PORT 8080
    On Error GoTo BadPort
    listener.LocalPort = 8080
    listener.Listen
    PrepareServer = True
    Exit Function
    
BadPort80:
    Resume ConnectTo8080
BadPort:
    Beep
End Function

