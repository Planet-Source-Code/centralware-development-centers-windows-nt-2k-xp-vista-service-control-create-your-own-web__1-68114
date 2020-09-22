VERSION 5.00
Object = "{C32049A2-E8D3-11D4-9D82-00D0B73B61A4}#2.1#0"; "cwSvc.ocx"
Begin VB.Form frmService 
   BorderStyle     =   0  'None
   Caption         =   "cwService VB Demonstration Project"
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Service.Service SERVICE 
      Left            =   4800
      Top             =   0
      _Version        =   131073
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      ServiceName     =   "WinName"
      StartMode       =   2
   End
   Begin VB.CommandButton cmdLink 
      Caption         =   "CLICK HERE TO RATE THIS PROJECT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   970
      Width           =   4575
   End
   Begin TestService.CPU CPU 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1720
   End
   Begin VB.Timer tmrSelfCheck 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5280
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form1.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   7935
   End
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Form_Load()
    Me.Hide                                     ' Hide just in case!
    
    With SERVICE
    .ControlsAccepted = svcCtrlPauseContinue    ' Supports Pause/Continue
    .DisplayName = App.Title                    ' Descriptive Name
    .Interactive = True                         ' Can interact with user
    .ServiceName = App.EXEName                  ' Actual Service Name
    .StartMode = svcStartAutomatic              ' Do we start with Windows?
    End With
    
    ' Check to see if the command-line has an INSTALL or
    ' UNINSTALL prompt.  If so, do that instead...
    If InStr(Command$, "-i") Then
        If SERVICE.Install Then
            MsgBox "Service Installed Successfully!", vbInformation + vbOKOnly, App.Title
        Else
            MsgBox "Service Failed To Install!", vbCritical + vbOKOnly, App.Title
        End If
        Unload Me
        Exit Sub
    End If
    If InStr(Command$, "-u") Then
        If SERVICE.Uninstall Then
            MsgBox "Service Uninstalled Successfully!", vbInformation + vbOKOnly, App.Title
        Else
            MsgBox "Service Failed To Uninstall!", vbCritical + vbOKOnly, App.Title
        End If
        Unload Me
        Exit Sub
    End If
    
    ' No command-line items are being processed, so let's start
    ' the application.  JUST IN CASE someone runs the service EXE
    ' manually, let's create a timer to self-check to ensure we
    ' are actually being called from the Service Manager...
    SERVICE.StartService
    tmrSelfCheck.Enabled = True
End Sub

Private Sub SERVICE_Continue(Success As Boolean)
    ' We've received a CONTINUE command from the service
    ' manager.  This should only be received if we were
    ' paused previously so there's no sense double-checking
    ' to see if that's true.  Treat a CONTINUE in the same
    ' fashion as you would a START command (assuming your
    ' application doesn't do additional tasks while paused.)
    Success = StartApplication
End Sub
Private Sub SERVICE_Control(ByVal EventID As Long)
    ' The CONTROL event is fired for different reasons by
    ' the service manager.  Under most cases, you'll likely
    ' never directly interact with this routine, but it's here
    ' just in case you need it.  For applications which sub-class,
    ' it is recommended that this routine is LISTED in your VB
    ' application, even if it contains nothing more than comments
    ' such as this.
End Sub
Private Sub SERVICE_Pause(Success As Boolean)
    ' We're being told to hold our horses!  STOP the application
    ' (Assuming you do not run tasks while paused) and sit idly
    ' by until we're told to go again.
    Success = StopApplication
End Sub
Private Sub SERVICE_Start(Success As Boolean)
    ' The service manager is telling us that we're being called
    ' directly (instead of being executed by double-clicking on
    ' the EXE file.)  For this reason, we'll need to shut off
    ' our self-check timer so that the application continues to
    ' run...
    tmrSelfCheck.Enabled = False
    ' From here, let's start doing what ever the application is
    ' intended to do.
    Success = StartApplication
    
    Height = CPU.Height + cmdLink.Height
    Width = CPU.Width
    Top = 0: Left = Screen.Width - Me.Width
    Me.Show
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
Private Sub SERVICE_Stop()
    ' We're being told that it's time to go!  The service manager
    ' doesn't care one way or the other if our STOP fails - it's
    ' going to close connections with us irregardless, so it's best
    ' to shut down everything weather there's an error or not.
    Call StopApplication
    Call ShutDown
End Sub

Private Sub tmrSelfCheck_Timer()
    ' The service manager didn't send us a START command.
    ' Let's exit here since we're not running when we're
    ' supposed to be!
    tmrSelfCheck.Enabled = False
    Unload Me
End Sub

Private Function StartApplication() As Boolean
    ' It's time to get things moving!
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' We're being asked to get the application started.  For
    ' socket based apps, this is a good place to BIND/LISTEN
    ' your main socket and create any additional run-time
    ' controls that may be needed.  If everything works out
    ' as planned, be sure to send a TRUE value back so that
    ' the service manager knows we're running properly.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    StartApplication = True
    CPU.Enabled = True
End Function
Private Function StopApplication() As Boolean
    ' We've been told to stop!
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' For applications that handle WinSock connections, this is
    ' a good time to simply close the socket connections and
    ' unload any CREATED run-time controls if applicable.
    ' DO NOT unload the form at this point - we could merely be
    ' paused so the ShutDown routine will handle closure for us.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CPU.Enabled = False
    StopApplication = True
End Function
Private Sub ShutDown()
    ' You may have additional steps to take before shutting down...
    Unload Me
End Sub





Private Sub cmdLink_Click()
Dim URL As String: URL = "http://www.centralware.com/psc/rate.php?project=cwservice"
ShellExecute hWnd, "Open", URL, vbNullString, vbNullString, 0
End Sub

