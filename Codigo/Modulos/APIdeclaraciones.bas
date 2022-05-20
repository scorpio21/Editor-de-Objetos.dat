Attribute VB_Name = "Api"
Option Explicit

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Function EnumDisplaySettings _
               Lib "user32" _
               Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, _
                                             ByVal iModeNum As Long, _
                                             lptypDevMode As Any) As Boolean

Public Declare Function ChangeDisplaySettings _
               Lib "user32" _
               Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, _
                                               ByVal dwFlags As Long) As Long

Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal Hwnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long

Public Declare Function GetActiveWindow Lib "user32" () As Long

Public ForbidenNames(500) As String

Public NumInsultos        As Integer

Public NumSeparadores     As Integer

Public Separadores(255)   As String

Public Declare Function SetWindowLong _
               Lib "user32" _
               Alias "SetWindowLongA" (ByVal Hwnd As Long, _
                                       ByVal nIndex As Long, _
                                       ByVal dwNewLong As Long) As Long

'Public Const LWA_ALPHA = &H2
Public Const GWL_EXSTYLE = -20

'Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT As Long = &H20&

Public Declare Function Process32First _
               Lib "kernel32" (ByVal hSnapshot As Long, _
                               lppe As PROCESSENTRY32) As Long

Public Declare Function Process32Next _
               Lib "kernel32" (ByVal hSnapshot As Long, _
                               lppe As PROCESSENTRY32) As Long

Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal handle As Long) As Long

Public Declare Function OpenProcess _
               Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, _
                                   ByVal bInheritHandle As Long, _
                                   ByVal dwProcId As Long) As Long

Public Declare Function EnumProcesses _
               Lib "psapi.dll" (ByRef lpidProcess As Long, _
                                ByVal cb As Long, _
                                ByRef cbNeeded As Long) As Long

Public Declare Function GetModuleFileNameExA _
               Lib "psapi.dll" (ByVal hProcess As Long, _
                                ByVal hModule As Long, _
                                ByVal ModuleName As String, _
                                ByVal nSize As Long) As Long

Public Declare Function EnumProcessModules _
               Lib "psapi.dll" (ByVal hProcess As Long, _
                                ByRef lphModule As Long, _
                                ByVal cb As Long, _
                                ByRef cbNeeded As Long) As Long

Public Declare Function CreateToolhelp32Snapshot _
               Lib "kernel32" (ByVal dwFlags As Long, _
                               ByVal th32ProcessID As Long) As Long

Public Declare Function GetVersionExA _
               Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Public Type PROCESSENTRY32

    dwSize     As Long
    cntUsage   As Long
    th32ProcessID As Long    ' This process
    th32DefaultHeapID As Long
    th32ModuleID As Long    ' Associated exe
    cntThreads As Long
    th32ParentProcessID As Long    ' This process's parent process
    pcPriClassBase As Long    ' Base priority of process threads
    dwFlags    As Long
    szexeFile  As String * 260    ' MAX_PATH

End Type

Public Type OSVERSIONINFO

    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long    '1 = Windows 95.
    '2 = Windows NT
    szCSDVersion As String * 128

End Type

Public Const PROCESS_QUERY_INFORMATION = 1024

Public Const PROCESS_VM_READ = 16

Public Const MAX_PATH = 260

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

Public Const SYNCHRONIZE = &H100000

'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Public Const PROCESS_ALL_ACCESS = &H1F0FFF

Public Const TH32CS_SNAPPROCESS = &H2&

Public Const hNull = 0

Public Type STARTUPINFO

    cb         As Long
    lpReserved As String
    lpDesktop  As String
    lpTitle    As String
    dwX        As Long
    dwY        As Long
    dwXSize    As Long
    dwYSize    As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags    As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput  As Long
    hStdOutput As Long
    hStdError  As Long

End Type

Public Type PROCESS_INFORMATION

    hProcess   As Long
    hThread    As Long
    dwProcessId As Long
    dwThreadID As Long

End Type

Public Const NORMAL_PRIORITY_CLASS = &H20&

Public Const INFINITE = -1&

Public Const WS_MINIMIZE = &H20000000

Public Declare Function mciSendString _
               Lib "winmm.dll" _
               Alias "mciSendStringA" (ByVal lpstrCommand As String, _
                                       ByVal lpstrReturnString As String, _
                                       ByVal uReturnLength As Long, _
                                       ByVal hwndCallback As Long) As Long

Public Const OFN_FILEMUSTEXIST = &H1000&

Public Const OFN_READONLY = &H4&

Public DialogCaption As String

Public FileName      As String

Public Declare Function GetShortPathName _
               Lib "kernel32" _
               Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
                                          ByVal lpszShortPath As String, _
                                          ByVal cchBuffer As Long) As Long

Public Type OPENFILENAME

    lStructSize As Long
    hwndOwner  As Long
    hInstance  As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile  As String
    nMaxFile   As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags      As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData  As Long
    lpfnHook   As Long
    lpTemplateName As String

End Type

Public Declare Function GetOpenFileName _
               Lib "comdlg32.dll" _
               Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

'////////CAMBIAR MODO VIDEO//////////////////////
Public Const CCDEVICENAME = 32

Public Const CCFORMNAME = 32

Public Const DM_BITSPERPEL = &H40000

Public Const DM_PELSWIDTH = &H80000

Public Const DM_PELSHEIGHT = &H100000

Public Const DM_DISPLAYFLAGS = 2097152    '; 0x200000

Public Const DM_DISPLAYFREQUENCY = 4194304    '; 0x400000

Public Const CDS_UPDATEREGISTRY = &H1

Public Const CDS_TEST = &H4

Public Const DISP_CHANGE_SUCCESSFUL = 0

Public Const DISP_CHANGE_RESTART = 1

Type typDevMODE

    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize     As Integer
    dmDriverExtra As Integer
    dmFields   As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale    As Integer
    dmCopies   As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor    As Integer
    dmDuplex   As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate  As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long

End Type

Public oldResHeight As Long, oldResWidth As Long

Public Sub ActiveProcess()

    Dim F                 As Long, SName As String, ModuleName As String

    Dim PROC              As PROCESSENTRY32

    Dim cb                As Long, cbNeeded As Long, NumElements As Long, ProcessIDs() As Long, CbNeeded2 As Long, NumElements2 As Long, HSnap As Long

    Dim LRet              As Long, nSize As Long, hProcess As Long, i As Long

    Dim Modules(1 To 300) As Long

    Select Case getVersion()

        Case 1    'Windows 95/98
            HSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)

            If HSnap = hNull Then End
            PROC.dwSize = Len(PROC)
            ' Iterate through the processes
            F = Process32First(HSnap, PROC)

            Do While F
                SName = StrZToStr(PROC.szexeFile)

                TaskList(itl) = SName
                itl = itl + 1

                Debug.Print SName

                F = Process32Next(HSnap, PROC)
            Loop

        Case 2    'Windows NT
            'Get the array containing the process id's for each process object
            cb = 8
            cbNeeded = 96

            Do While cb <= cbNeeded
                cb = cb * 2
                ReDim ProcessIDs(cb / 4) As Long
                LRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
            Loop

            NumElements = cbNeeded / 4

            For i = 1 To NumElements
                'Get a handle to the Process
                hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
                'Got a Process handle

                'If HProcess <> 0 Then

                'Get an array of the module handles for the specified
                'process
                LRet = EnumProcessModules(hProcess, Modules(1), 300, CbNeeded2)

                'If the Module Array is retrieved, Get the ModuleFileName
                If LRet <> 0 Then
                    ModuleName = Space(MAX_PATH)
                    nSize = 500
                    LRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)

                    Debug.Print ModuleName

                    TaskList(itl) = ModuleName
                    itl = itl + 1

                    'Open "c:\procesos_activos.log" For Append As #100
                    'Print #100, ModuleName
                    'Close #100
                End If

                'End If
                'Close the handle to the process
                LRet = CloseHandle(hProcess)
            Next

    End Select

End Sub

Public Function getVersion() As Long

    Dim osinfo   As OSVERSIONINFO

    Dim retvalue As Integer

    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    getVersion = osinfo.dwPlatformId

End Function

Function StrZToStr(S As String) As String
    StrZToStr = Left$(S, Len(S) - 1)

End Function

