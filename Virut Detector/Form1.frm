VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Process Win32/Virut Detector v0.1 (x86)"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11505
   BeginProperty Font 
      Name            =   "맑은 고딕"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   11505
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   4440
      Width           =   9495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "종료"
      Height          =   855
      Left            =   9720
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "전체치료"
      Height          =   735
      Left            =   9720
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "새로고침"
      Height          =   735
      Left            =   9720
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7435
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Label1"
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   4800
      Width           =   525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function VirtualQueryEx Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpAddress As Any, ByRef lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Private Type MEMORY_BASIC_INFORMATION
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    lType As Long
End Type

Private Declare Function ZwQueryInformationProcess Lib "ntdll.dll" ( _
     ByVal ProcessHandle As Long, _
     ByVal ProcessInformationClass As Long, _
     ProcessInformation As Any, _
     ByVal ProcessInformationLength As Long, _
     ReturnLength As Long _
) As Long

Private Declare Function GetModuleFileNameExW Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As Long, ByVal nSize As Long) As Long
Private Declare Function VirtualProtectEx Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long


Private Is32Bit  As Boolean
Private Declare Function RtlAdjustPrivilege Lib "ntdll" (ByVal Privilege As Long, ByVal bEnablePrivilege As Long, ByVal bCurrentThread As Long, ByRef OldState As Long) As Long
Private Declare Function IsWow64Process Lib "kernel32" _
    (ByVal hProc As Long, ByRef bWow64Process As Boolean) As Long

Private Type FunctionAddress
    FunName As String
    FunOriginal(5) As Byte
End Type

Dim FunByte() As FunctionAddress
Dim numFunByte As Long
Private Declare Function ReadProcessMemory Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long

Public Function RVA2RAW(ByVal RVA As Long, ByRef INH As IMAGE_NT_HEADERS, ByRef ISH() As IMAGE_SECTION_HEADER) As Long
    Dim NumSection As Long: NumSection = INH.FileHeader.NumberOfSections
    Dim dwSizeOfImage As Long: dwSizeOfImage = INH.OptionalHeader.SizeOfImage

    If RVA Then
    
        Dim i As Long
        For i = 0 To NumSection - 1
        
            If i = NumSection - 1 Then
                If ISH(i).VirtualAddress <= RVA Then
                    If dwSizeOfImage >= RVA Then
                        RVA2RAW = RVA - ISH(i).VirtualAddress + ISH(i).PointerToRawData
                        Exit Function
                    End If
                End If
            End If
            
            
            If ISH(i).VirtualAddress <= RVA Then
                If RVA < ISH(i + 1).VirtualAddress Then
                    RVA2RAW = RVA - ISH(i).VirtualAddress + ISH(i).PointerToRawData
                    Exit Function
                End If
            End If
            
        Next i
        
    End If
End Function
Private Sub Command1_Click()    '## 새로고침
    ListView1.ListItems.Clear
    
   '## 프로세스 루프.
    Dim hSnapshot As Long
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = INVALID_HANDLE_VALUE Then
        MsgBox "프로세스를 불러오지 못하였습니다."
        End
    End If
    
    
    Dim Process As PROCESSENTRY32W
    Process.dwSize = LenB(Process)
    
    If Process32FirstW(hSnapshot, Process) = 0& Then
        MsgBox "모든 프로세스를 조회하지 못하였습니다."
        End
    End If
    
    Do
        If Process.th32ProcessID = 1688 Then
            'Stop
        End If
        
        Dim hProcess As Long, b32Process As Long
        hProcess = OpenProcess(MAXIMUM_ALLOWED, False, Process.th32ProcessID)
  
        
        If hProcess Then
        
            '## 64비트 환경이면 -> 64비트 프로세스인지 확인 -> 64비트 프로세스면 안함
            If Is32Bit = False Then
                'ProcessWow64Information = 26d
                
                Call ZwQueryInformationProcess(hProcess, 26, b32Process, 4, 0)
                If b32Process = False Then
                    CloseHandle hProcess
                    GoTo Continue1
                End If
            End If

            Dim APIAddress As Long
            
            Dim i As Long
            For i = 0 To numFunByte - 1
             
                APIAddress = GetProcAddressByPID(Process.th32ProcessID, "NTDLL", FunByte(i).FunName)
                If APIAddress Then
                    Dim ProcessByte(5) As Byte, lpNumberOfBytesRead As Long
                    
                    Call ReadProcessMemory(hProcess, ByVal APIAddress, ProcessByte(0), 6, lpNumberOfBytesRead)
                    If MyMemcmp(ProcessByte, FunByte(i).FunOriginal) = False Then
                    
                        Dim CALLBytes(4) As Byte, CALLAddress As Long
                        RtlMoveMemory CALLBytes(0), ProcessByte(0), 5
                        
                        If CALLBytes(0) = &HE8 Then
                            RtlMoveMemory CALLAddress, CALLBytes(1), 4
                            
                            Dim MBI As MEMORY_BASIC_INFORMATION
                            Call VirtualQueryEx(hProcess, ByVal CALLAddress + APIAddress + 5&, MBI, Len(MBI))
                            
                            Const MEM_MAPPED = &H40000
                            '## DLL 방지
                            If MBI.lType = MEM_MAPPED Then
                                Dim lLen As Long, Buffer As String
                                Buffer = String$(260, 0)
                                lLen = GetModuleFileNameExW(hProcess, 0&, ByVal StrPtr(Buffer), 260)
                                
                                ListView1.ListItems.Add , , Process.szExeFile
                                ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(1) = Process.th32ProcessID
                                ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(2) = Buffer
                                
                            End If
                        End If
                        
                        GoTo Continue1:
                    End If
                    
                End If
            Next i

            
            CloseHandle hProcess
        End If
        
Continue1:
    Loop While Process32NextW(hSnapshot, Process)
    
    CloseHandle hSnapshot
End Sub

Private Sub Command2_Click()    '## 전체복구

    If ListView1.ListItems.Count = 0& Then Exit Sub

Dim i As Long, j As Long
For i = 1 To ListView1.ListItems.Count

    Dim PID As Long
    PID = ListView1.ListItems(i).SubItems(1)
    
    
        Dim hProcess As Long, b32Process As Long
        hProcess = OpenProcess(MAXIMUM_ALLOWED, False, PID)
  
        
        If hProcess Then
            
            For j = 0 To numFunByte - 1
                Dim APIAddress As Long
                APIAddress = GetProcAddressByPID(PID, "NTDLL", FunByte(j).FunName)
                If APIAddress Then
                    Dim ProcessByte(5) As Byte, lpNumberOfBytesRead As Long
                    
                    Call ReadProcessMemory(hProcess, ByVal APIAddress, ProcessByte(0), 6, lpNumberOfBytesRead)
                    If MyMemcmp(ProcessByte, FunByte(j).FunOriginal) = False Then
                    
                        Dim OldProProtect As Long
                        VirtualProtectEx hProcess, ByVal APIAddress, ByVal 6, PAGE_EXECUTE_READWRITE, OldProProtect
                        Call WriteProcessMemory(hProcess, ByVal APIAddress, FunByte(j).FunOriginal(0), 6, lpNumberOfBytesRead)
                        VirtualProtectEx hProcess, ByVal APIAddress, ByVal 6, OldProProtect, ByVal 0&
                        
                    End If
                    
                End If
            Next j
        End If
            
            CloseHandle hProcess
Next i

    '## 새로고침
    Command1_Click
End Sub


Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Activate()
    RtlAdjustPrivilege 20, 1, 0, 0
    
    Label1.Caption = "리스트에 뜨는 프로세스는 바이러스에 감염되었거나 어떤것 의하여 조작된 것 입니다."
    Text1.Text = vbNullString
    
    With ListView1
        .Font.Size = 9
        .View = lvwReport
        .ColumnHeaders.Add , , "프로세스이름", 2000
        .ColumnHeaders.Add , , "PID", 700, lvwColumnCenter
        .ColumnHeaders.Add , , "프로세스경로", 6500, lvwColumnCenter
    End With
    
    SendMessage ListView1.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT + LVS_EX_GRIDLINES + LVS_EX_HEADERDRAGDROP, ByVal -1


    '## 32비트 판단  '[ 64비트라면 32비트 프로세스 판단으로 IsWow64Process API 를 사용하기 때문 ]
    Dim lpWow64Process As Long
    lpWow64Process = GetProcAddress(GetModuleHandle("Kernel32"), "IsWow64Process")
    If lpWow64Process = 0& Then Is32Bit = True
    
    '## NTDLL.DLL 의 바이너리를 불러옴  ( GetProcAddress 로 주소를 구해서 바이트를 구하면 , 이 프로세스도 감염되어 있을수도 있으니 안됨 + 호환성 )
    Dim FileBinary() As Byte, FileSize As Long
    Open "C:\WIndows\system32\NTDLL.DLL" For Binary Access Read As #1
        FileSize = LOF(1)
        ReDim FileBinary(FileSize - 1) As Byte
        Get #1, , FileBinary
    Close #1
    
    Dim IDH As IMAGE_DOS_HEADER
    Dim INH As IMAGE_NT_HEADERS
    Dim ISH() As IMAGE_SECTION_HEADER
    
    RtlMoveMemory IDH, FileBinary(0), Len(IDH)
    RtlMoveMemory INH, FileBinary(IDH.e_lfanew), Len(INH)

    Dim pSection As Long: pSection = IDH.e_lfanew + 4& + Len(INH.FileHeader) + INH.FileHeader.SizeOfOptionalHeader
    ReDim ISH(INH.FileHeader.NumberOfSections - 1) As IMAGE_SECTION_HEADER
    RtlMoveMemory ISH(0), FileBinary(pSection), INH.FileHeader.NumberOfSections * SIZE_OF_SECTION32
    
    If INH.FileHeader.Machine <> IMAGE_FILE_MACHINE_I386 Then
        MsgBox "오류) 해당 프로그램은 32비트 DLL 파일이 아닙니다."
        Exit Sub
    End If
    
    Dim pExport As Long
    pExport = INH.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT).VirtualAddress
    pExport = RVA2RAW(pExport, INH, ISH)
    
    Dim Export As IMAGE_EXPORT_DIRECTORY
    RtlMoveMemory Export, FileBinary(pExport), Len(Export)
    
    '## 정의
    Dim AddressTable() As Long
    Dim Ordinal() As Integer
    Dim pNameTable() As Long
    Dim NameTable() As String
    
    '## 재정의
    ReDim AddressTable(Export.NumberOfFunctions - 1) As Long
    ReDim Ordinal(Export.NumberOfFunctions - 1) As Integer
    ReDim NameTable(Export.NumberOfFunctions - 1) As String
    ReDim pNameTable(Export.NumberOfNames - 1) As Long
    
    RtlMoveMemory AddressTable(0), FileBinary(RVA2RAW(Export.AddressOfFunctions, INH, ISH)), Export.NumberOfFunctions * 4
    RtlMoveMemory Ordinal(0), FileBinary(RVA2RAW(Export.AddressOfNameOrdinals, INH, ISH)), Export.NumberOfFunctions * 2
    RtlMoveMemory pNameTable(0), FileBinary(RVA2RAW(Export.AddressOfNames, INH, ISH)), Export.NumberOfNames * 4
    
    Dim i As Long
    For i = 0 To UBound(pNameTable)
        If FileBinary(RVA2RAW(pNameTable(i), INH, ISH)) Then
            NameTable(Ordinal(i)) = GetPointerToString(VarPtr(FileBinary(RVA2RAW(pNameTable(i), INH, ISH))))
        End If
    Next i
    
'## 구조체 정리
    
    Dim FunName As String
    Dim Address As Long
    Dim tmpByte(5) As Byte
    
    
    FunName = "ZwCreateProcess"
    Address = RVA2RAW(AddressTable(FindArrString(NameTable, FunName)), INH, ISH)
    RtlMoveMemory tmpByte(0), FileBinary(Address), 6
    AddFun FunByte, numFunByte, FunName, tmpByte

    FunName = "ZwQueryInformationProcess"
    Address = RVA2RAW(AddressTable(FindArrString(NameTable, FunName)), INH, ISH)
    RtlMoveMemory tmpByte(0), FileBinary(Address), 6
    AddFun FunByte, numFunByte, FunName, tmpByte
    
    FunName = "ZwOpenFile"
    Address = RVA2RAW(AddressTable(FindArrString(NameTable, FunName)), INH, ISH)
    RtlMoveMemory tmpByte(0), FileBinary(Address), 6
    AddFun FunByte, numFunByte, FunName, tmpByte
    
    FunName = "ZwCreateProcessEx"
    Address = RVA2RAW(AddressTable(FindArrString(NameTable, FunName)), INH, ISH)
    RtlMoveMemory tmpByte(0), FileBinary(Address), 6
    AddFun FunByte, numFunByte, FunName, tmpByte
    
    FunName = "ZwCreateProcess"
    Address = RVA2RAW(AddressTable(FindArrString(NameTable, FunName)), INH, ISH)
    RtlMoveMemory tmpByte(0), FileBinary(Address), 6
    AddFun FunByte, numFunByte, FunName, tmpByte
    
    FunName = "ZwCreateFile"
    Address = RVA2RAW(AddressTable(FindArrString(NameTable, FunName)), INH, ISH)
    RtlMoveMemory tmpByte(0), FileBinary(Address), 6
    AddFun FunByte, numFunByte, FunName, tmpByte '
    
    FunName = "ZwCreateUserProcess"
    Address = RVA2RAW(AddressTable(FindArrString(NameTable, FunName)), INH, ISH)
    RtlMoveMemory tmpByte(0), FileBinary(Address), 6
    AddFun FunByte, numFunByte, FunName, tmpByte
    
    
    '## 프로세스 로딩
    Command1_Click
 
End Sub
Private Function MyMemcmp(ByRef Arr() As Byte, ByRef Arr2() As Byte) As Boolean
    Dim i As Long
    For i = 0 To UBound(Arr)
        If Arr(i) <> Arr2(i) Then
            Exit Function
        End If
    Next i
    MyMemcmp = True
End Function

Private Function GetProcAddressByPID(ByVal PID As Long, szModule As String, FunctionName As String) As Long
    Dim BaseAddress As Long
    BaseAddress = GetModuleHandle(szModule)
    BaseAddress = GetProcAddress(BaseAddress, FunctionName) - BaseAddress
    
    Dim hModSnapshot As Long
    hModSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, PID)

    
    If hModSnapshot = INVALID_HANDLE_VALUE Then Exit Function
    
    
    Dim Module As MODULEENTRY32W  'MODULEENTRY32 '
    Module.dwSize = Len(Module)
    
    Do
        If StrComp(Module.szModule, szModule, vbTextCompare) = 0& Then
            GetProcAddressByPID = Module.modBaseAddr + BaseAddress
            GoTo End1:
        End If
        
        If StrComp(Module.szModule, szModule & ".DLL", vbTextCompare) = 0& Then
            GetProcAddressByPID = Module.modBaseAddr + BaseAddress
            GoTo End1:
        End If
        
    Loop While Module32NextW(hModSnapshot, Module)
    
End1:
    CloseHandle hModSnapshot
End Function
Private Sub AddFun(ByRef FunByte() As FunctionAddress, ByRef numFunByte As Long, ByVal FunName As String, ByRef Original() As Byte)
    If (Not (FunByte)) = -1& Then
        ReDim FunByte(0) As FunctionAddress
        numFunByte = 1&
    Else
        ReDim Preserve FunByte(numFunByte) As FunctionAddress
        numFunByte = numFunByte + 1
    End If
    
    With FunByte(numFunByte - 1)
        .FunName = FunName
        RtlMoveMemory .FunOriginal(0), Original(0), 6
    End With
End Sub
Private Function FindArrString(ByRef Arr() As String, ByRef FindString As String) As Long
    For FindArrString = 0 To UBound(Arr)
        If StrComp(Arr(FindArrString), FindString, vbTextCompare) = 0& Then Exit Function
    Next FindArrString
End Function
Private Function GetPointerToString(ByVal mPointer As Long) As String
    Dim Buffer As String
    Buffer = String$(260, 0)
    
    lstrcpyA_p2 Buffer, ByVal mPointer
    Buffer = Left(Buffer, InStr(1, Buffer, vbNullChar, vbBinaryCompare) - 1)
    
    GetPointerToString = Buffer
End Function

Private Sub ListView1_Click()
If ListView1.ListItems.Count Then
    Text1.Text = ListView1.SelectedItem.SubItems(2)
End If
End Sub

Private Sub Text1_Click()
    Text1.SelStart = 0&
    Text1.SelLength = Len(Text1.Text)
End Sub
