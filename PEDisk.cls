VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PEDisk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text


'Private Declare PtrSafe Function GetModuleInformation Lib "psapi.dll" (ByVal hProcess As LongPtr, ByVal hModule As LongPtr, ByRef lpModInfo As MODULEINFO, cb As Long) As Long
'Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32.dll" () As LongPtr
'Private Declare PtrSafe Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
'Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32.dll" (ByRef destination As Any, ByVal lpSource As LongPtr, ByVal size As Long)

'Private mhProcess As LongPtr, mhModule As LongPtr
Private mFileHandle As Long
Private mModuleInfo As MODULEINFO
Private mDosHeader As IMAGE_DOS_HEADER
Private mModuleName As String
Private mNtHeader As IMAGE_NT_HEADERS
Private mSections() As IMAGE_SECTION_HEADER_Disk
Private mModuleInfoSize As Long, mDosHeaderSize As Long, mNtHeaderSize As Long, mSectionSize As Long, mNtOptionalHeaderStructureSize

Implements iPortableExecutable

Private Sub Class_Initialize()
  Dim tempSection As IMAGE_SECTION_HEADER_Memory, tempOptionalHeader As IMAGE_OPTIONAL_HEADER
  mModuleInfoSize = LenB(mModuleInfo)
  mDosHeaderSize = LenB(mDosHeader)
  mNtHeaderSize = LenB(mNtHeader)
  mSectionSize = LenB(tempSection)
  mNtOptionalHeaderStructureSize = LenB(tempOptionalHeader)
End Sub

Private Property Get ModuleInfoSize() As Long
  ModuleInfoSize = mModuleInfoSize
End Property

Private Property Get DosHeaderSize() As Long
  DosHeaderSize = mDosHeaderSize
End Property

Private Property Get NtHeaderSize() As Long
  NtHeaderSize = mNtHeaderSize
End Property

Private Property Get SectionSize() As Long
  SectionSize = mSectionSize
End Property

Private Sub LoadDosHeader()
  Get #mFileHandle, 1, mDosHeader
End Sub

Private Sub LoadNtHeader()
  If mDosHeader.e_lfanew = 0 Then
    Exit Sub
  End If
  Get #mFileHandle, CLng(NtHeaderStart) + 1, mNtHeader
  If (mNtHeader.FileHeader.SizeOfOptionalHeader > 0) Then
    mModuleInfo.EntryPoint = mNtHeader.OptionalHeader.AddressOfEntryPoint
    mModuleInfo.lpBaseOfDll = mNtHeader.OptionalHeader.ImageBase
    mModuleInfo.SizeOfImage = mNtHeader.OptionalHeader.SizeOfImage
  End If
End Sub

Private Property Get NumberOfSections() As Long
  NumberOfSections = mNtHeader.FileHeader.NumberOfSections
End Property

Private Property Get NtOptionalHeaderSize() As Long
  NtOptionalHeaderSize = CLng(mNtHeader.FileHeader.SizeOfOptionalHeader)
End Property

Private Property Get NtOptionalHeaderStructureSize() As Long
  NtOptionalHeaderStructureSize = mNtOptionalHeaderStructureSize
End Property

Private Property Get SectionStart() As LongPtr
  SectionStart = NtHeaderStart + NtHeaderSize + NtOptionalHeaderSize - NtOptionalHeaderStructureSize
End Property

Private Sub LoadSectionHeaders()
If NumberOfSections < 1 Then Exit Sub
ReDim mSections(0 To NumberOfSections - 1)
Dim i As Long
Dim thisSection As IMAGE_SECTION_HEADER_Disk
  Get #mFileHandle, CLng(SectionStart) + 1, thisSection
  mSections(0) = thisSection
For i = 1 To NumberOfSections - 1
  Get #mFileHandle, , thisSection
  mSections(i) = thisSection
Next
End Sub

Private Property Get NtHeaderStart() As LongPtr
  NtHeaderStart = mDosHeader.e_lfanew
End Property

Public Property Get DllBase() As LongPtr
  DllBase = mModuleInfo.lpBaseOfDll
End Property

Public Property Get ImageSize() As LongPtr
  ImageSize = mModuleInfo.SizeOfImage
End Property

Public Function FindInSection(SectionName As String, BufferToFind() As Byte, Optional StartingLocation As LongPtr = 0) As LongPtr
On Error GoTo HandleError
  Dim SectionNumber As Long
  Dim basePointer As LongPtr, bytesToCompare As Long
  Dim BufferToSearch() As Byte
  SectionNumber = GetSectionId(SectionName)
  With mSections(SectionNumber)
    basePointer = DllBase + .VirtualAddress
    bytesToCompare = CLng(.SizeOfRawData - StartingLocation)
    If bytesToCompare <= 0 Then
      Err.Raise VBA.vbObjectError, "FindInSection", "Invalid starting location (" & Hex(StartingLocation) & ")."
      Exit Function
    End If
    ReDim BufferToSearch(0 To bytesToCompare - 1)
    Get #mFileHandle, CLng(basePointer + StartingLocation), BufferToSearch
    FindInSection = VBA.InStrB(1, BufferToSearch, BufferToFind, vbBinaryCompare) - 1
  End With

Exit Function
HandleError:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetSectionId(SectionName As String) As Long
  Dim i As Long, myName As String
  For i = LBound(mSections) To UBound(mSections)
    myName = StrConv(mSections(i).SectionName, vbUnicode)
    If InStr(1, myName, vbNullChar) > 0 Then
      myName = Left$(myName, InStr(1, myName, vbNullChar))
    End If
    If SectionName = myName Then
      GetSectionId = i
      Exit Function
    End If
  Next
  GetSectionId = -1
  Err.Raise VBA.vbObjectError, "GetSectionId", "The section named """ & SectionName & """ was not found."
End Function

Private Sub Class_Terminate()
  On Error Resume Next
  Close #mFileHandle
End Sub

#If Win64 Then
  Private Function iPortableExecutable_FindInSection(SectionName As String, BufferToFind() As Byte, Optional StartingLocation As LongLong = 0) As LongLong
    iPortableExecutable_FindInSection = FindInSection(SectionName, BufferToFind, StartingLocation)
  End Function
  
  Private Property Get iPortableExecutable_ImageSize() As LongPtr
    iPortableExecutable_ImageSize = ImageSize
  End Property

#Else
  Private Function iPortableExecutable_FindInSection(SectionName As String, BufferToFind() As Byte, Optional StartingLocation As Long = 0) As Long
    iPortableExecutable_FindInSection = FindInSection(SectionName, BufferToFind, StartingLocation)
  End Function
  
  Private Property Get iPortableExecutable_ImageSize() As Long
    iPortableExecutable_ImageSize = ImageSize
  End Property

#End If

Private Sub iPortableExecutable_Load(moduleName As String)
  Dim result As Long
  mFileHandle = FreeFile
  On Error GoTo HandleError
  Open moduleName For Binary Shared As #mFileHandle
  LoadDosHeader
  LoadNtHeader
  LoadSectionHeaders
Exit Sub
HandleError:
  Err.Raise Err.Number, "iPortableExecutable_Load", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

