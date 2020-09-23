VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ShowScreen 
   Caption         =   "Show Screen Saver Preview"
   ClientHeight    =   8415
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9780
   Icon            =   "ShowScreen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Screen Saver"
      Filter          =   "scr"
      InitDir         =   "C:\WinNT\System32a"
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8415
      Left            =   0
      ScaleHeight     =   8355
      ScaleWidth      =   9375
      TabIndex        =   0
      Top             =   0
      Width           =   9435
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Click Here to open a Screen Saver for Viewing"
   End
End
Attribute VB_Name = "ShowScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SysDir As String

Private Sub Form_Load()
Dim verinfo As OSVERSIONINFO
Dim build As String, ver_major As String, ver_minor As String
Dim ret As Long

    Picture1.Move 0, 0, Me.Width, Me.Height
    
    WinDir = GetWindowsDir
    verinfo.dwOSVersionInfoSize = Len(verinfo)
    ret = GetVersionEx(verinfo)

    Select Case verinfo.dwPlatformId
        Case 0
            SysDir = ""
        Case 1
            SysDir = "system"
        Case 2
            SysDir = "System32"
    End Select
    

End Sub

Private Sub Form_Resize()
    Picture1.Move 0, 0, Me.Width, Me.Height
    
End Sub

Private Sub mnuOpen_Click()
    
    Me.CommonDialog1.InitDir = GetWindowsDir & SysDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "Screen Savers (*.scr)|*.scr"
    ' Specify default filter
    CommonDialog1.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Show selected Screen Saver
    FileCopy Me.CommonDialog1.filename, WinDir & "ZZZ.exe"
    Shell WinDir & "ZZZ.exe /p " & Picture1.hWnd, vbHide
    mnuOpen.Enabled = False
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
    
End Sub


Private Sub mnuView_Click()

    Me.CommonDialog1.InitDir = GetWindowsDir & SysDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "Screen Savers (*.scr)|*.scr"
    ' Specify default filter
    CommonDialog1.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file
    'MsgBox CommonDialog1.filename
    Debug.Print Me.CommonDialog1.FileTitle
    FileCopy Me.CommonDialog1.filename, WinDir & "ZZZ.exe"
    Shell WinDir & "ZZZ.exe /p " & Picture1.hWnd, vbHide
    mnuView.Enabled = False
    
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub
