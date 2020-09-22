VERSION 5.00
Begin VB.UserControl CPU 
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   HasDC           =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4905
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   960
      Top             =   240
   End
   Begin VB.PictureBox picUsage 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   975
      Left            =   0
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   1
      Top             =   0
      Width           =   990
      Begin VB.Label lblCpuUsage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00008000&
      Height          =   975
      Left            =   960
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "CPU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private QueryObject As Object
Private Const m_def_Enabled = 0
Dim m_Enabled As Boolean
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
End Sub
Private Sub UserControl_Resize()
    Width = picGraph.Left + picGraph.Width
    Height = picGraph.Height
End Sub
Private Sub UserControl_Terminate()
    Call DeinitControl
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
End Sub

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled: PropertyChanged "Enabled"
    Select Case New_Enabled
        Case True: Call InitControl
        Case False: Call DeinitControl
    End Select
End Property
Private Sub InitControl()
    On Error Resume Next
    SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    If IsWinNT Then
        Set QueryObject = New clsCPUUsageNT
    Else
        Set QueryObject = New clsCPUUsage
    End If
    QueryObject.Initialize
    tmrRefresh.Enabled = True
    tmrRefresh_Timer
End Sub
Private Sub DeinitControl()
    On Error Resume Next
    tmrRefresh.Enabled = False
    QueryObject.Terminate
    Set QueryObject = Nothing
End Sub
Private Sub tmrRefresh_Timer()
    Dim Ret As Long
    On Error GoTo NoConnection
    Ret = QueryObject.Query
    If Ret = -1 Then
        tmrRefresh.Enabled = False
        lblCpuUsage.Caption = ":("
    Else
        DrawUsage Ret, picUsage, picGraph
        lblCpuUsage.Caption = CStr(Ret) + "%"
    End If
ExitSub:
    Exit Sub
NoConnection:
    Resume ExitSub
End Sub

