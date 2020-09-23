Attribute VB_Name = "Module1"
Option Explicit
Global WinDir As String

'Used to get Windows directory
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Declare Function GetWindowsDirectory Lib "Kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Global Const gintMAX_SIZE% = 255
Global Const gstrSEP_DIR$ = "\"
Global Const gstrNULL$ = ""

Type SYSTEM_INFO
       dwOemID As Long
       dwPageSize As Long
       lpMinimumApplicationAddress As Long
       lpMaximumApplicationAddress As Long
       dwActiveProcessorMask As Long
       dwNumberOrfProcessors As Long
       dwProcessorType As Long
       dwAllocationGranularity As Long
       dwReserved As Long
 End Type
 Type OSVERSIONINFO
       dwOSVersionInfoSize As Long
       dwMajorVersion As Long
       dwMinorVersion As Long
       dwBuildNumber As Long
       dwPlatformId As Long
       szCSDVersion As String * 128
 End Type
 Type MEMORYSTATUS
       dwLength As Long
       dwMemoryLoad As Long
       dwTotalPhys As Long
       dwAvailPhys As Long
       dwTotalPageFile As Long
       dwAvailPageFile As Long
       dwTotalVirtual As Long
       dwAvailVirtual As Long
 End Type

 Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (LpVersionInformation As OSVERSIONINFO) As Long
 Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As _
    MEMORYSTATUS)
 Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As _
    SYSTEM_INFO)

 Public Const PROCESSOR_INTEL_386 = 386
 Public Const PROCESSOR_INTEL_486 = 486
 Public Const PROCESSOR_INTEL_PENTIUM = 586
 Public Const PROCESSOR_MIPS_R4000 = 4000
 Public Const PROCESSOR_ALPHA_21064 = 21064

Function StripNull(S As String) As String
    Dim n%
    
    n% = InStr(S, Chr(0))
    If n% > 0 Then
        StripNull = Mid$(S, 1, n% - 1)
    Else
        StripNull = S
    End If
End Function

Public Function ReturnSystemInfo() As String
    Dim msg As String         ' Status information.

    Screen.MousePointer = 11    ' Hourglass.

        ' Get operating system and version.
        Dim verinfo As OSVERSIONINFO
        Dim build As String, ver_major As String, ver_minor As String
        Dim ret As Long
        verinfo.dwOSVersionInfoSize = Len(verinfo)
        ret = GetVersionEx(verinfo)
        If ret = 0 Then
            MsgBox "Error Getting Version Information"
            End
        End If
        Select Case verinfo.dwPlatformId
            Case 0
                msg = msg & "Windows 32s "
            Case 1
                If verinfo.dwMinorVersion = 0 Then
                    msg = msg & "Windows 95 "
                Else
                    msg = msg & "Windows 98 "
                End If
            Case 2
                msg = msg & "Windows NT "
        End Select

        ver_major = verinfo.dwMajorVersion
        ver_minor = verinfo.dwMinorVersion
        build = verinfo.dwBuildNumber
        msg = msg & ver_major & "." & ver_minor
        msg = msg & " (Build " & build & ")"
        msg = msg & " " & StripNull(verinfo.szCSDVersion) & vbCrLf

        ' Get CPU type and operating mode.
        Dim sysinfo As SYSTEM_INFO
        GetSystemInfo sysinfo
        msg = msg & "CPU: "
        Select Case sysinfo.dwProcessorType
            Case PROCESSOR_INTEL_386
                msg = msg & "Intel 386" & vbCrLf
            Case PROCESSOR_INTEL_486
                msg = msg & "Intel 486" & vbCrLf
            Case PROCESSOR_INTEL_PENTIUM
                msg = msg & "Intel Pentium" & vbCrLf
            Case PROCESSOR_MIPS_R4000
                msg = msg & "MIPS R4000" & vbCrLf
            Case PROCESSOR_ALPHA_21064
                msg = msg & "DEC Alpha 21064" & vbCrLf
            Case Else
                msg = msg & "(unknown)" & vbCrLf

        End Select
        ' Get free memory.
        Dim memsts As MEMORYSTATUS
        Dim memory As Long
        GlobalMemoryStatus memsts
        memory = memsts.dwTotalPhys
        msg = msg & "Total Physical Memory: "
        msg = msg & Format$(memory \ 1024, "###,###,###") & "K" _
                  & vbCrLf
        memory& = memsts.dwAvailPhys
        msg = msg & "Available Physical Memory: "
        msg = msg & Format$(memory \ 1024, "###,###,###") & "K" _
                  & vbCrLf
        memory& = memsts.dwTotalVirtual
        msg = msg & "Total Virtual Memory: "
        msg = msg & Format$(memory \ 1024, "###,###,###") & "K" _
                  & vbCrLf
        memory& = memsts.dwAvailVirtual
        msg = msg & "Available Virtual Memory: "
        msg = msg & Format$(memory \ 1024, "###,###,###") & "K" _
                  & vbCrLf
        ReturnSystemInfo = msg
    Screen.MousePointer = 0
End Function

Function GetWindowsDir() As String
    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator$(strBuf)
        AddDirSep strBuf
        GetWindowsDir = UCase16(strBuf)
    Else
        GetWindowsDir = gstrNULL
    End If
End Function

Function FileExists(filename) As Boolean
    FileExists = (Dir(filename) <> "")
End Function

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Sub AddDirSep(strPathName As String)
    If Right$(RTrim$(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub

Function UCase16(ByVal str As String)
#If Win16 Then
    UCase16 = UCase$(str)
#Else
    UCase16 = str
#End If
End Function

