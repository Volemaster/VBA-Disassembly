Attribute VB_Name = "Supplemental"
Option Explicit


Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Const MAX_VBA_OBJECT_NAME_LENGTH As Long = 1024
Public Const INVALID_PROCEDURE_CALL_OR_ARGUMENT As Long = 5

Public Declare PtrSafe Function GetModuleInformation Lib "psapi" (ByVal hProcess As LongPtr, ByVal hModule As LongPtr, ByRef lpModInfo As MODULEINFO, cb As Long) As Long
Public Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Public Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef destination As Any, ByVal lpSource As LongPtr, ByVal size As Long)
Public Declare PtrSafe Function VirtualQueryEx Lib "kernel32" (ByVal hProcess As LongPtr, ByVal lpAddress As LongPtr, ByRef MemoryInformation As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long


Private mMemory As MemoryTool

Public Property Get Memory() As MemoryTool
  If mMemory Is Nothing Then
    Set mMemory = New MemoryTool
  End If
  Set Memory = mMemory
End Property


Public Function RTrimNull(str As String) As String
  Dim l As Long
  l = VBA.InStr(1, str, vbNullChar)
  If l > 0 Then
    RTrimNull = VBA.Left$(str, l - 1)
  Else
    RTrimNull = str
  End If
End Function

Public Function TryGetVbaProjects(ByVal lpDataSegmentProjectPointer As LongPtr, ByRef outVbaProjects As VbaProjects) As Boolean
  Dim iLoader As ILoadedFromAddress
  On Error GoTo HandleError
  Set iLoader = New VbaProjects
  Call iLoader.LoadFromAddress(lpDataSegmentProjectPointer)
  Set outVbaProjects = iLoader
  TryGetVbaProjects = True
  Exit Function
HandleError:
  Set outVbaProjects = Nothing
  TryGetVbaProjects = False
End Function


