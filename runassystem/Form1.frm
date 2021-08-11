VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "[VB6] Run As System 0.1"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   8355
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox Text2 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      OLEDropMode     =   1  '수동
      TabIndex        =   2
      Text            =   "명령줄"
      Top             =   480
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "실행"
      Height          =   660
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      OLEDropMode     =   1  '수동
      TabIndex        =   0
      Text            =   "파일경로 (파일 드래그)"
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MAXIMUM_ALLOWED As Long = &H2000000


'## 프로세스 루프
Private Declare Function CreateToolhelp32Snapshot Lib "KERNEL32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "KERNEL32.dll" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "KERNEL32.dll" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Const TH32CS_SNAPPROCESS As Long = &H2
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

'## 프로세스 생성
Private Declare Function CreateProcessWithTokenW Lib "Advapi32.dll" ( _
    ByVal hToken As Long, _
    ByVal dwLogonFlags As Long, _
    ByVal lpApplicationName As Long, _
    ByVal lpCommandLine As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As Long, _
    ByRef lpStartupInfo As STARTUPINFO, _
    ByRef lpProcessInfo As PROCESS_INFORMATION) As Long


Private Declare Function CreateEnvironmentBlock Lib "userenv.dll" (ByRef lpEnvironment As Any, ByVal hToken As Long, ByVal bInherit As Long) As Long


Private Declare Function DuplicateTokenEx Lib "Advapi32.dll" ( _
    ByVal hExistingToken As Long, _
    ByVal dwDesiredAccess As Long, _
    ByVal lpTokenAttributes As Long, _
    ByVal ImpersonationLevel As Long, _
    ByVal TokenType As Long, _
    ByRef phNewToken As Long) As Long

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type




'## Vista
Private Declare Function WTSGetActiveConsoleSessionId Lib "KERNEL32.dll" () As Long
Private Declare Function ProcessIdToSessionId Lib "KERNEL32.dll" (ByVal dwProcessId As Long, ByRef pSessionId As Long) As Long



Private Function GetEnvironmentBlock(ByVal sProcess As String) As Long

    Dim hSnapshot As Long
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = -1& Then
        MsgBox "CreateToolhelp32Snapshot Error"
        Exit Function
    End If
    
    Dim PE As PROCESSENTRY32
    PE.dwSize = Len(PE)
    
    Call Process32First(hSnapshot, PE)
    Do
        If StrComp(PE.szExeFile, sProcess, vbTextCompare) = 0& Then
            CloseHandle hSnapshot
            GoTo Go:
        End If
    Loop While Process32Next(hSnapshot, PE)
    
    CloseHandle hSnapshot
    Exit Function
Go:
    
    
    Dim hProcess As Long
    hProcess = OpenProcess(MAXIMUM_ALLOWED, False, PE.th32ProcessID)
    If hProcess = 0& Then
        MsgBox "OpenProcess"
        End
    End If
    
    Dim hToken As Long
    If OpenProcessToken(hProcess, TOKEN_DUPLICATE Or TOKEN_QUERY, hToken) = 0& Then
        MsgBox "OpenProcessToken"
        CloseHandle hProcess
        End
    End If

    Dim pEnv As Long
    Call CreateEnvironmentBlock(pEnv, hToken, 1)
    GetEnvironmentBlock = pEnv
    
    CloseHandle hProcess
    CloseHandle hToken
End Function

'## 프로세스를 루프하여 그 프로세스가 인수랑 맞으면 PID 를 반환하는 함수.
Private Function GetProcessSessionPID(ByVal SessionId As Long) As Long
    Dim hSnapshot As Long
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = -1& Then
        MsgBox "CreateToolhelp32Snapshot Error"
        Exit Function
    End If
    
    Dim PE As PROCESSENTRY32
    PE.dwSize = Len(PE)
    
    Call Process32First(hSnapshot, PE)
    Do
        Dim ProcessSessionID As Long
        ProcessIdToSessionId PE.th32ProcessID, ProcessSessionID
    
        If ProcessSessionID = SessionId Then
            GetProcessSessionPID = PE.th32ProcessID
            GoTo End1:
        End If
        
    Loop While Process32Next(hSnapshot, PE)
    
End1:
    CloseHandle hSnapshot
End Function
Private Sub Command1_Click()

    '## 현재 활성화된 세션 아이디를 얻음
    Dim dwSessionId As Long
    dwSessionId = WTSGetActiveConsoleSessionId()
    If (Err.LastDllError Or dwSessionId) = &HFFFFFFFF Then
        MsgBox "WTSGetActiveConsoleSessionId Error"
        Exit Sub
    End If
    
    '## 프로세스 루프해서 해당 세션ID 와 일치하는 PID 를 얻음
    Dim ProcessPID As Long
    ProcessPID = GetProcessSessionPID(dwSessionId)
    
    Dim hProcess As Long
    hProcess = OpenProcess(MAXIMUM_ALLOWED, False, ProcessPID)
    If hProcess = 0& Then
        MsgBox "OpenProcess"
        Exit Sub
    End If
    
    Dim hToken As Long
    If OpenProcessToken(hProcess, TOKEN_DUPLICATE, hToken) = 0& Then
        MsgBox "OpenProcessToken"
        CloseHandle hProcess
        Exit Sub
    End If
    
    Dim hDupToken As Long, NewToken As Long
    hDupToken = DuplicateTokenEx(hToken, MAXIMUM_ALLOWED, ByVal 0&, ByVal 1&, ByVal 1&, NewToken)
    If hDupToken = 0& Then
        MsgBox "DuplicateTokenEx"
        CloseHandle hProcess
        CloseHandle hToken
        Exit Sub
    End If
    
    Dim pEnvBlock As Long
    pEnvBlock = GetEnvironmentBlock("winlogon.exe")
    
    
    Dim PI As PROCESS_INFORMATION
    Dim SI As STARTUPINFO
    SI.cb = Len(SI)
    SI.lpDesktop = StrPtr("winsta0\default" & vbNullChar & vbNullChar)

    Dim FilePath As String, CmdLine As String, CurrentDir As String
    FilePath = Text1.Text
    CmdLine = Text2.Text
    CurrentDir = Left(Text1.Text, InStrRev(Text1.Text, "\", , vbBinaryCompare))
    
    CreateProcessWithTokenW NewToken, 1, ByVal 0&, ByVal StrPtr(FilePath & " " & CmdLine), &H20& Or &H400&, pEnvBlock, ByVal StrPtr(CurrentDir), SI, PI
    
    CloseHandle hProcess
    CloseHandle hToken
End Sub

Private Sub Form_Load()
    '## 권한상승
    SetPrivilege "SeDebugPrivilege"
    SetPrivilege "SeAssignPrimaryTokenPrivilege"
    SetPrivilege "SeIncreaseQuotaPrivilege"
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.Text = Data.Files(1)
    Text1.SelStart = Len(Text1.Text)
    Text2.Text = vbNullString
End Sub
Private Sub Text2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text2.Text = Data.Files(1)
    Text2.SelStart = Len(Text2.Text)
End Sub
