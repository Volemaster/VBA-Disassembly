Attribute VB_Name = "PE_Helper"

Public Type IMAGE_DOS_HEADER ' DOS .EXE header
  e_magic As Integer         ' Magic number
  e_cblp As Integer          ' Bytes on last page of file
  e_cp As Integer            ' Pages in file
  e_crlc As Integer          ' Relocations
  e_cparhdr As Integer       ' Size of header in paragraphs
  e_minalloc As Integer      ' Minimum extra paragraphs needed
  e_maxalloc As Integer      ' Maximum extra paragraphs needed
  e_ss As Integer            ' Initial (relative) SS value
  e_sp As Integer            ' Initial SP value
  e_csum As Integer          ' Checksum
  e_ip As Integer            ' Initial IP value
  e_cs As Integer            ' Initial (relative) CS value
  e_lfarlc As Integer        ' File address of relocation table
  e_ovno As Integer          ' Overlay number
  e_res(0 To 3) As Integer        ' Reserved words
  e_oemid As Integer         ' OEM identifier (for e_oeminfo)
  e_oeminfo As Integer       ' OEM information; e_oemid specific
  e_res2(0 To 9) As Integer      ' Reserved words
  e_lfanew As Long           ' File address of new exe header
End Type

Public Type IMAGE_FILE_HEADER
  Machine As Integer
  NumberOfSections As Integer
  TimeDateStamp As Long
  PointerToSymbolTable As Long
  NumberOfSymbols As Long
  SizeOfOptionalHeader As Integer
  Characteristics As Integer
End Type

Public Type IMAGE_DATA_DIRECTORY
  VirtualAddress As Long
  size As Long
End Type

#If Win64 Then
Public Type IMAGE_OPTIONAL_HEADER
  Magic As Integer
  MajorLinkerVersion As Byte
  MinorLinkerVersion As Byte
  SizeOfCode As Long
  SizeOfInitializedData As Long
  SizeOfUninitializedData As Long
  AddressOfEntryPoint As Long
  BaseOfCode As Long
  ImageBase As LongPtr
  SectionAlignment As Long
  FileAlignment As Long
  MajorOperatingSystemVersion As Integer
  MinorOperatingSystemVersion As Integer
  MajorImageVersion As Integer
  MinorImageVersion As Integer
  MajorSubsystemVersion As Integer
  MinorSubsystemVersion As Integer
  Win32VersionValue As Long
  SizeOfImage As Long
  SizeOfHeaders As Long
  CheckSum As Long
  Subsystem As Integer
  DllCharacteristics As Integer
  SizeOfStackReserve As LongPtr
  SizeOfStackCommit As LongPtr
  SizeOfHeapReserve As LongPtr
  SizeOfHeapCommit As LongPtr
  LoaderFlags As Long
  NumberOfRvaAndSizes As Long   ' Number of Image_Data_Directory entries
  ImageDataDirectories(0 To &HF) As IMAGE_DATA_DIRECTORY    ' This is the maximum number of entries.
                                                            ' Use the number to determine the actual size.
End Type

#Else
Public Type IMAGE_OPTIONAL_HEADER
  Magic As Integer
  MajorLinkerVersion As Byte
  MinorLinkerVersion As Byte
  SizeOfCode As Long
  SizeOfInitializedData As Long
  SizeOfUninitializedData As Long
  AddressOfEntryPoint As Long
  BaseOfCode As Long
  BaseOfData As Long
  ImageBase As LongPtr
  SectionAlignment As Long
  FileAlignment As Long
  MajorOperatingSystemVersion As Integer
  MinorOperatingSystemVersion As Integer
  MajorImageVersion As Integer
  MinorImageVersion As Integer
  MajorSubsystemVersion As Integer
  MinorSubsystemVersion As Integer
  Win32VersionValue As Long
  SizeOfImage As Long
  SizeOfHeaders As Long
  CheckSum As Long
  Subsystem As Integer
  DllCharacteristics As Integer
  SizeOfStackReserve As LongPtr
  SizeOfStackCommit As LongPtr
  SizeOfHeapReserve As LongPtr
  SizeOfHeapCommit As LongPtr
  LoaderFlags As Long
  NumberOfRvaAndSizes As Long   ' Number of Image_Data_Directory entries
  ImageDataDirectories(0 To &HF) As IMAGE_DATA_DIRECTORY    ' This is the maximum number of entries.
                                                            ' Use the number to determine the actual size.
End Type
#End If

Public Enum IMAGE_SECTION_CHARACTERISTICS
  RESERVED_1 = &H0           ' Reserved.
  RESERVED_2 = &H1           ' Reserved.
  RESERVED_3 = &H2           ' Reserved.
  RESERVED_4 = &H4           ' Reserved.
  IMAGE_SCN_TYPE_NO_PAD = &H8           ' The section should not be padded to the next boundary. This flag is obsolete and is replaced by IMAGE_SCN_ALIGN_1BYTES.
  RESERVERD_5 = &H10         ' Reserved.
  IMAGE_SCN_CNT_CODE = &H20          ' The section contains executable code.
  IMAGE_SCN_CNT_INITIALIZED_DATA = &H40          ' The section contains initialized data.
  IMAGE_SCN_CNT_UNINITIALIZED_DATA = &H80          ' The section contains uninitialized data.
  IMAGE_SCN_LNK_OTHER = &H100         ' Reserved.
  IMAGE_SCN_LNK_INFO = &H200         ' The section contains comments or other information. This is valid only for object files.
  RESERVED_6 = &H400          ' Reserved.
  IMAGE_SCN_LNK_REMOVE = &H800         ' The section will not become part of the image. This is valid only for object files.
  IMAGE_SCN_LNK_COMDAT = &H1000        ' The section contains COMDAT data. This is valid only for object files.
  RESERVED_7 = &H2000         ' Reserved.
  IMAGE_SCN_NO_DEFER_SPEC_EXC = &H4000        ' Reset speculative exceptions handling bits in the TLB entries for this section.
  IMAGE_SCN_GPREL = &H8000        ' The section contains data referenced through the global pointer.
  RESERVED_8 = &H10000        ' Reserved.
  IMAGE_SCN_MEM_PURGEABLE = &H20000       ' Reserved.
  IMAGE_SCN_MEM_LOCKED = &H40000       ' Reserved.
  IMAGE_SCN_MEM_PRELOAD = &H80000       ' Reserved.
  IMAGE_SCN_ALIGN_1BYTES = &H100000      ' Align data on a 1-byte boundary. This is valid only for object files.
  IMAGE_SCN_ALIGN_2BYTES = &H200000      ' Align data on a 2-byte boundary. This is valid only for object files.
  IMAGE_SCN_ALIGN_4BYTES = &H300000      ' Align data on a 4-byte boundary. This is valid only for object files.
  IMAGE_SCN_ALIGN_8BYTES = &H400000      ' Align data on a 8-byte boundary. This is valid only for object files.
  IMAGE_SCN_ALIGN_16BYTES = &H500000      ' Align data on a 16-byte boundary. This is valid only for object files.
  IMAGE_SCN_ALIGN_32BYTES = &H600000      ' Align data on a 32-byte boundary. This is valid only for object files.
  IMAGE_SCN_ALIGN_64BYTES = &H700000      ' Align data on a 64-byte boundary. This is valid only for object files.
  IMAGE_SCN_ALIGN_128BYTES = &H800000      ' Align data on a 128-byte boundary. This is valid only for object files.
  IMAGE_SCN_ALIGN_256BYTES = &H900000      ' Align data on a 256-byte boundary. This is valid only for object files.
  IMAGE_SCN_ALIGN_512BYTES = &HA00000      ' Align data on a 512-byte boundary. This is valid only for object files.
  IMAGE_SCN_ALIGN_1024BYTES = &HB00000      ' Align data on a 1024-byte boundary. This is valid only for object files.
  IMAGE_SCN_ALIGN_2048BYTES = &HC00000      ' Align data on a 2048-byte boundary. This is valid only for object files.
  IMAGE_SCN_ALIGN_4096BYTES = &HD00000      ' Align data on a 4096-byte boundary. This is valid only for object files.
  IMAGE_SCN_ALIGN_8192BYTES = &HE00000      ' Align data on a 8192-byte boundary. This is valid only for object files.
  IMAGE_SCN_LNK_NRELOC_OVFL = &H1000000     ' The section contains extended relocations. The count of relocations for the section exceeds the 16 bits that is reserved for it in the section header. If the NumberOfRelocations field in the section header is 0xffff, the actual relocation count is stored in the VirtualAddress field of the first relocation. It is an error if IMAGE_SCN_LNK_NRELOC_OVFL is set and there are fewer than 0xffff relocations in the section.
  IMAGE_SCN_MEM_DISCARDABLE = &H2000000     ' The section can be discarded as needed.
  IMAGE_SCN_MEM_NOT_CACHED = &H4000000     ' The section cannot be cached.
  IMAGE_SCN_MEM_NOT_PAGED = &H8000000     ' The section cannot be paged.
  IMAGE_SCN_MEM_SHARED = &H10000000    ' The section can be shared in memory.
  IMAGE_SCN_MEM_EXECUTE = &H20000000    ' The section can be executed as code.
  IMAGE_SCN_MEM_READ = &H40000000    ' The section can be read.
  IMAGE_SCN_MEM_WRITE = &H80000000    ' The section can be written to.
End Enum

Public Type IMAGE_SECTION_HEADER_Memory
  SectionName(0 To 7) As Byte
  VirtualSize As Long             ' For files: PhysicalAddress
  VirtualAddress As Long
  SizeOfRawData As Long
  PointerToRawData As Long
  PointerToRelocations As Long
  PointerToLinenumbers As Long
  NumberOfRelocations As Integer
  NumberOfLinenumbers As Integer
  Characteristics As IMAGE_SECTION_CHARACTERISTICS
End Type

Public Type IMAGE_SECTION_HEADER_Disk
  SectionName(0 To 7) As Byte
  PhysicalAddress As Long             ' For memory: VirtualSize
  VirtualAddress As Long
  SizeOfRawData As Long
  PointerToRawData As Long
  PointerToRelocations As Long
  PointerToLinenumbers As Long
  NumberOfRelocations As Integer
  NumberOfLinenumbers As Integer
  Characteristics As Long 'IMAGE_SECTION_CHARACTERISTICS
End Type

Public Enum IMAGE_DIRECTORY_ENTRY_INDEX
  Exports = &H0            ' Export table address and size
  Imports = &H1            ' Import table address and size
  Resource = &H2          ' Resource table address and size
  Exception = &H3         ' Exception table address and size
  Certificate = &H4       ' Certificate table address and size
  Base = &H5              ' Base relocation table address and size
  Debugging = &H6         ' Debugging information starting address and size
  Architecture = &H7      ' Architecture-specific data address and size
  GlobalPtr = &H8         ' Global pointer register relative virtual address
  Thread = &H9            ' Thread local storage (TLS) table address and size
  LoadConfiguration = &HA ' Load configuration table address and size
  Bound = &HB             ' Bound import table address and size
  Import = &HC            ' Import address table address and size
  Delay = &HD             ' Delay import descriptor address and size
  CLR = &HE               ' The CLR header address and size
  Reserved = &HF          ' Reserved
End Enum

Public Type IMAGE_NT_HEADERS
  Signature As Long
  FileHeader As IMAGE_FILE_HEADER
  OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type

Public Type MODULEINFO
  lpBaseOfDll As LongPtr
  SizeOfImage As Long
  EntryPoint As LongPtr
End Type

Public Function LoadPE(ModuleName As String) As iPortableExecutable
  'Distinguish between files in memory and those on disk.
  Dim PEName As String
  PEName = VBA.Mid$(ModuleName, VBA.InStrRev(ModuleName, "\") + 1)
  If PEName <> ModuleName Then
    Set LoadPE = New PEDisk
    LoadPE.Load ModuleName
    'Try to pull up the file on disk
  Else
    Set LoadPE = New PEMemory
    On Error GoTo LoadFromDisk
    LoadPE.Load PEName
  End If
Exit Function
LoadFromDisk:
  On Error GoTo HandleError
  Exit Function
HandleError:
  Set LoadPE = Nothing
End Function
