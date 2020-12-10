VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 0
Option Explicit


Private Type DataSegmentProjectPointer
    Null1 As LongLong                          '   0x00
    lpProjectEntry As LongPtr                  '   0x08
    Ptr1 As LongPtr                            '   0x10
    Null2 As LongLong                          '   0x18
    Ptr2 As LongPtr                            '   0x20
    Null3 As LongPtr                           '   0x28
End Type

Private Type ProjectEntry
    lpNextEntry As LongPtr                     '   0x00
    lpPreviousEntry As LongPtr                 '   0x08
    lpProjectInfo As LongPtr                   '   0x10
    QWord1 As LongLong                         '   0x18
    QWord2 As LongLong                         '   0x20
End Type

Private Type ProjectInfo
    ObjectInterfaces As LongPtr                '   0x00
    QWord1 As LongLong                         '   0x08
    lpD270 As LongPtr                          '   0x10    'Unknown
    lpObjectTable As LongPtr                   '   0x18
    FileNameDescriptor(3) As Byte              '   0x20
    FileName(2045) As Byte                     '   0x23
    QWord2 As LongLong                         ' 0x0820
    lpE288 As LongPtr                          ' 0x0828    'Unknown
    lpPublicObjectDescriptors As LongPtr       ' 0x0830    'Pointer to array of PODs themselves, not an array of pointers.
    ObjectCount As LongPtr                     ' 0x0838
    ObjectInfo1Array As LongPtr                ' 0x0840
    QWord3 As LongLong                         ' 0x0848
End Type

Private mDSPP As DataSegmentProjectPointer
Private mProjectCount As Long
Private mProjectEntry() As ProjectEntry
Private mProjectInfo() As ProjectInfo
Private mProjectEntrySize As Long
Private mProjectInfoSize As Long
Private mProjectNames As Scripting.Dictionary
Private mOriginalAddress As LongPtr

Private mVbaProject() As IVbaProject

Private Const CLASS_NAME As String = "VbaProjects"

Implements IVbaProjects

Private Sub Class_Initialize()
  Const METHOD_NAME As String = CLASS_NAME & ".Class_Initialize"
  On Error GoTo HandleError
  mProjectCount = 0
  ReDim mProjectEntry(0 To &HF) 'Very unlikely to see more than 2 or 3, but this memory is cheap.
  ReDim mProjectInfo(0 To &H1)  'We can safely assume at least 2 since this is a supporting library. We won't load the immediate one.
  ReDim mVbaProject(0 To &HF)  'This is reasonable to do since the painful memory allocations take place within the class itself, not in this class.
  mProjectEntrySize = LenB(mProjectEntry(0))
  mProjectInfoSize = LenB(mProjectInfo(0))
Exit Sub

HandleError:
  Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub Setup(lpAddress As LongPtr)
Const METHOD_NAME As String = CLASS_NAME & ".Setup"
' Do not expose this to the main class interface
  On Error GoTo HandleError
  mOriginalAddress = lpAddress
  Call LoadProjectEntries(lpAddress)
  Call LoadProjectInfos
Exit Sub
HandleError:
  Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Sub LoadProjectEntries(ByVal lpAddress As LongPtr)
  Const METHOD_NAME As String = "LoadProjectEntries"
  Dim nextProject As LongPtr, lastProject As LongPtr
  On Error GoTo HandleError
  Call Memory.FollowPointer(VarPtr(mDSPP), lpAddress, LenB(mDSPP))
  nextProject = mDSPP.lpProjectEntry
  lastProject = 0
  #If Win64 Then
    mProjectCount = 0
  #Else
    mProjectCount = 1
  #End If
  Do Until nextProject = 0
    Call Memory.FollowPointer(VarPtr(mProjectEntry(mProjectCount)), nextProject, LenB(mProjectEntry(mProjectCount)))
    If mProjectEntry(mProjectCount).lpProjectInfo = 0 Then
      #If Win64 Then
        Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "ProjectEntry(" & mProjectCount & ") structure invalid [bad lpProjectInfo]."
      #Else
        'This is valid and indicates that there are no more projects
        nextProject = 0
      #End If
    ElseIf mProjectCount > 0 Then
      'Does not exist in Win32
      #If Win64 Then
        If (mProjectEntry(mProjectCount).lpPreviousEntry) <> lastProject Then
          Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "ProjectEntry(" & mProjectCount & ") structure invalid [bad lpPreviousEntry]."
        End If
      #End If
    Else
      'Does not exist in Win32
      #If Win64 Then
        If mProjectEntry(0).lpPreviousEntry <> 0 Then
          Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "Initial ProjectEntry structure invalid [bad lpPreviousEntry]."
        End If
      #End If
      
    End If
    lastProject = nextProject
    nextProject = mProjectEntry(mProjectCount).lpNextEntry
    If nextProject <> 0 Then
      mProjectCount = mProjectCount + 1
      If mProjectCount > UBound(mProjectEntry) Then
        ReDim Preserve mProjectEntry(mProjectCount + &H10)  'Total overkill.
      End If
    End If
  Loop
  
Exit Sub
HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Property Get IVbaProjects_Count() As Long
  Const METHOD_NAME As String = CLASS_NAME & ".get_Count"
  On Error GoTo HandleError
  IVbaProjects_Count = mProjectCount
Exit Property
HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IVbaProjects_Item(Index As Variant) As IVbaProject
  Const METHOD_NAME = CLASS_NAME & ".get_Item"
On Error GoTo HandleError
  If IsNull(Index) Then
    Err.Raise 94, METHOD_NAME
  End If
  If IsNumeric(Index) Then
    If CLng(Index) < mProjectCount And Index >= 0 Then
      Set IVbaProjects_Item = mVbaProject(CLng(Index + 1))  ' This is a 1-based array.
      Exit Property
    End If
  End If
  If mProjectNames.Exists(Index) Then
    Set IVbaProjects_Item = mVbaProject(CLng(mProjectNames(Index)))
    Exit Property
  End If
  If IsNumeric(Index) Then
    Err.Raise 381, METHOD_NAME, "Invalid project specified: the zero-based identifier you specified (" & Index & ") does not exist."
  Else
    Err.Raise 381, METHOD_NAME, "Invalid project specified: the project name """ & Index & """ could not be found."
  End If
Exit Property
HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get ProjectNames() As Scripting.Dictionary
  Const METHOD_NAME As String = CLASS_NAME & ".get_ProjectNames"
  On Error GoTo HandleError
  If mProjectNames Is Nothing Then
    Set mProjectNames = New Scripting.Dictionary
    mProjectNames.CompareMode = TextCompare
  End If
  Set ProjectNames = mProjectNames
Exit Property
HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Sub LoadProjectInfos()
  Const METHOD_NAME As String = CLASS_NAME & ".LoadProjectInfos"
  Dim i As Long, lpProjectInfo As LongPtr
  Dim lpObjectTable As LongPtr, lpOTProjectInfo As LongPtr
  Const OTPROJECTINFOOFFSET As Long = &H8
  
  On Error GoTo HandleError
  ReDim mVbaProject(1 To mProjectCount)
  ProjectNames.RemoveAll
  For i = 1 To mProjectCount
    Set mVbaProject(i) = PrivateFactory.CreateVbaProject(mProjectEntry(i).lpProjectInfo)  ' New VbaProject
    ProjectNames.Add mVbaProject(i).Name, i
  Next
Exit Sub

HandleError:
  If (i <= UBound(mVbaProject)) And (i >= LBound(mVbaProject)) Then
    Set mVbaProject(i) = Nothing
  End If
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Sub Class_Terminate()
  Dim i As Long
  Dim DC As IChildOfDisposable
  For i = 1 To IVbaProjects_Count
    Set DC = IVbaProjects_Item(i - 1)
    If Not DC Is Nothing Then
      DC.DisposeOfParent
    End If
  Next
End Sub
