VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "cwService User Control"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Click here to install the User Control and its sources"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdStart_Click()
    Dim X As Integer, TMP As String, FF As Integer
    Dim fName As String, fData() As Byte
    
    Call SourceDir          ' Creates Source Directory
    TMP = " ": X = 100
    On Error Resume Next
    Do While TMP <> ""
        X = X + 1
        TMP = ""
        TMP = LoadResString(X)
        If Trim(TMP) <> "" Then
            Erase fData
            fData = LoadResData(X, "CUSTOM")
            FF = FreeFile
            If InStr(LCase(TMP), "ocx") Then
                Open sPath & "\" & TMP For Binary As FF
                Put #FF, , fData
                Close FF
                Call RegisterControl(sPath & "\" & TMP)
            Else
                Open sPath & "\Source\" & TMP For Binary As FF
                Put #FF, , fData
                Close FF
            End If
        End If
    Loop
    MsgBox "Files extracted to " & vbCrLf & sPath
    Unload Me
End Sub
Private Sub SourceDir()
    Dim PATH As String: PATH = Trim(Environ("systemdrive"))
    If PATH = "" Then PATH = Left(Environ("windir"), 2)
    PATH = Left(PATH, 2) & "\cwService"
    On Error Resume Next
    MkDir PATH
    MkDir PATH & "\Source"
End Sub
Private Function sPath() As String
    Dim PATH As String: PATH = Trim(Environ("systemdrive"))
    If PATH = "" Then PATH = Left(Environ("windir"), 2)
    sPath = Left(PATH, 2) & "\cwService"
End Function
Private Sub RegisterControl(OCXName As String)
    ShellExecute Me.hwnd, "Open", "RegSvr32.exe", OCXName, vbNullString, 0
End Sub
