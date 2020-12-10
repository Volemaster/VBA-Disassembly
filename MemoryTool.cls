VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MemoryTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private Type MEMORY_BASIC_INFORMATION
  BaseAddress As LongPtr
  AllocationBase As LongPtr
  AllocationProtect As Long
  PartitionId As Integer
  RegionSize As LongPtr
  State As Long
  Protect As Long
  Type As Long
End Type

Private Declare PtrSafe Function GetModuleInformation Lib "psapi" (ByVal hProcess As LongPtr, ByVal hModule As LongPtr, ByRef lpModInfo As MODULEINFO, cb As Long) As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef destination As Any, ByVal lpSource As LongPtr, ByVal size As Long)
Private Declare PtrSafe Function VirtualQueryEx Lib "kernel32" (ByVal hProcess As LongPtr, ByVal lpAddress As LongPtr, ByRef MemoryInformation As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long

Private mMemoryRange() As MEMORY_BASIC_INFORMATION
Private mMemoryRangeCount As Long
Private MemoryRangeSize As Long
Private mSessionId As String

Private mhProcess As LongPtr
Private mPageSize As Long

Private Sub Class_Initialize()
  ReDim mMemoryRange(0 To &HF)
  MemoryRangeSize = LenB(mMemoryRange(0))
  mhProcess = GetCurrentProcess()
  mSessionId = vbNullChar
  mPageSize = &H1000
End Sub

Public Sub StartSession(SessionId As String)
  If SessionId <> mSessionId Then
    PurgeSessionData
  End If
  mSessionId = SessionId
End Sub

Public Sub EndSession()
  PurgeSessionData
  mSessionId = vbNullChar
End Sub

Private Sub PurgeSessionData()
  mMemoryRangeCount = 0
  ReDim mMemoryRange(0 To &HF)
End Sub

Private Function IsValidPointer(ByVal lpAddress As LongPtr) As Boolean
  Dim i As Long, rangeCount As Long
  'Null pointers are the most common exception...
  If lpAddress = 0 Then Exit Function
  'Inherently invalid addresses don't need to be checked
  #If Win64 Then
    If lpAddress And &H8000000000000000^ Then Exit Function
  #Else
    If lpAddress And &H80000000 Then Exit Function
  #End If
  
  Dim thisMemoryRange As MEMORY_BASIC_INFORMATION
  
  If mSessionId = vbNullChar Then
    If VirtualQueryEx(mhProcess, lpAddress, thisMemoryRange, MemoryRangeSize) Then
      IsValidPointer = True
    Else
      IsValidPointer = False
    End If
    Exit Function
  End If
  
  'Of course this doesn't take memory deallocations into account....
  For i = 0 To mMemoryRangeCount - 1
    'In C this wouldn't work because of how the compiler simplifies pointer arithmatic... but we don't have to worry about it here.
    If lpAddress >= mMemoryRange(i).AllocationBase And lpAddress < (mMemoryRange(i).AllocationBase + mMemoryRange(i).RegionSize) Then
      IsValidPointer = True
      Exit Function
    End If
  Next
  rangeCount = UBound(mMemoryRange)
  If mMemoryRangeCount >= rangeCount Then
    rangeCount = rangeCount + &H10
    ReDim Preserve mMemoryRange(0 To rangeCount)
  End If
  If VirtualQueryEx(mhProcess, lpAddress, mMemoryRange(mMemoryRangeCount), MemoryRangeSize) Then
    mMemoryRangeCount = mMemoryRangeCount + 1
    IsValidPointer = True
    Exit Function
  Else
    IsValidPointer = False
  End If
End Function

Public Sub FollowPointer(ByVal lpTarget As LongPtr, ByVal lpSource As LongPtr, ByVal size As Long, Optional ThrowOnSourceExceedsPageError As Boolean = True)
  If Not (IsValidPointer(lpSource)) Then
    Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, "MemoryTool.FollowPointer", "The source pointer 0x" & Hex(lpSource) & " is not a valid pointer."
  ElseIf Not (IsValidPointer(lpTarget)) Then
    Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, "MemoryTool.FollowPointer", "The target pointer 0x" & Hex(lpTarget) & " is not a valid pointer."
  Else
    If ((lpTarget Mod mPageSize) + size) > mPageSize Then
      If Not (IsValidPointer(lpTarget + size)) Then
        Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, "MemoryTool.FollowPointer", "The memory size 0x" & Hex(size) & " is invalid at address 0x" & Hex(lpTarget) & "."
      End If
    ElseIf ((lpSource Mod mPageSize) + size) > mPageSize Then
      If Not (IsValidPointer(lpSource + size)) Then
        If ThrowOnSourceExceedsPageError Then
          Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, "MemoryTool.FollowPointer", "The memory size 0x" & Hex(size) & " is invalid at address 0x" & Hex(lpSource) & "."
        Else
          size = mPageSize - CLng(lpSource Mod mPageSize)
        End If
      End If
    End If
  End If
  RtlMoveMemory ByVal lpTarget, lpSource, size
End Sub


