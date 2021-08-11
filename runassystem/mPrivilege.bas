Attribute VB_Name = "mPrivilege"
Option Explicit
Public Declare Function OpenProcess Lib "KERNEL32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function OpenProcessToken Lib "Advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Public Const SE_PRIVILEGE_ENABLED = &H2
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const TOKEN_ASSIGN_PRIMARY As Long = &H1
Public Const TOKEN_DUPLICATE As Long = &H2
Public Const TOKEN_IMPERSONATE As Long = &H4
Public Const TOKEN_ADJUST_DEFAULT As Long = &H80
Public Const TOKEN_ADJUST_SESSIONID As Long = &H100
Public Const TOKEN_ADJUST_GROUPS As Long = &H40
Public Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Public Const TOKEN_QUERY_SOURCE As Long = &H10
Public Const TOKEN_QUERY As Long = &H8
Public Const TOKEN_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_SESSIONID Or TOKEN_ADJUST_DEFAULT)

Private Declare Function LookupPrivilegeValue Lib "Advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "Advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, BufferLength As Any, PreviousState As Any, ReturnLength As Any) As Long

Private Const ANYSIZE_ARRAY = 1
Private Type LUID
   lowpart As Long
   highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
   pLuid As LUID
   Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Public Function SetPrivilege(ByVal Privilege As String) As Boolean
   Dim tpPrev As TOKEN_PRIVILEGES
   Dim lid As LUID
   Dim tpSize As Long
   Dim lRet As Long
   
   Dim hCurProc As Long: hCurProc = -1&
   
   Dim hToken As Long
   lRet = OpenProcessToken(hCurProc, TOKEN_ALL_ACCESS, hToken)
   If lRet = 0& Then Exit Function

   lRet = LookupPrivilegeValue(vbNullString, Privilege, lid)
   If lRet = 0& Then Exit Function
   
   Dim tp As TOKEN_PRIVILEGES
   tpSize = Len(tp)
   tp.PrivilegeCount = 1
   tp.Privileges(0).pLuid = lid
   tp.Privileges(0).Attributes = 0
   lRet = AdjustTokenPrivileges(hToken, 0, tp, tpSize, tpPrev, tpSize)

   tpPrev.PrivilegeCount = 1
   tpPrev.Privileges(0).pLuid = lid
   tpPrev.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
   lRet = AdjustTokenPrivileges(hToken, 0, tpPrev, tpSize, ByVal 0&, ByVal 0&)
   If lRet = 0& Then Exit Function
   
   CloseHandle hToken
   SetPrivilege = True
End Function
