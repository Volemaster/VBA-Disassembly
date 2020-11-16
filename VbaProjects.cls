VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private mVbaProject() As VbaProject

Private Const CLASS_NAME As String = "VbaProjects"

Implements ILoadedFromAddress
Implements IMethodical

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

#If Win64 Then
  Private Sub ILoadedFromAddress_LoadFromAddress(ByVal lpAddress As LongLong, Optional Index As Long = 0)
#Else
  Private Sub ILoadedFromAddress_LoadFromAddress(ByVal lpAddress As Long, Optional Index As Long = 0)
#End If
Const METHOD_NAME As String = "ILoadedFromAddress_LoadFromAddress[" & CLASS_NAME & "]"
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
  Do Until nextProject = 0
    Call Memory.FollowPointer(VarPtr(mProjectEntry(mProjectCount)), nextProject, LenB(mProjectEntry(mProjectCount)))
    If mProjectEntry(mProjectCount).lpProjectInfo = 0 Then
        Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "ProjectEntry(" & mProjectCount & ") structure invalid [bad lpProjectInfo]."
    ElseIf mProjectCount > 0 Then
      If (mProjectEntry(mProjectCount).lpPreviousEntry) <> lastProject Then
        Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "ProjectEntry(" & mProjectCount & ") structure invalid [bad lpPreviousEntry]."
      End If
    Else
      If mProjectEntry(0).lpPreviousEntry <> 0 Then
        Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "Initial ProjectEntry structure invalid [bad lpPreviousEntry]."
      End If
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

Public Property Get Count() As Long
  Const METHOD_NAME As String = CLASS_NAME & ".get_Count"
  On Error GoTo HandleError
  Count = mProjectCount
Exit Property
HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Public Property Get Item(Index As Variant) As VbaProject
  Const METHOD_NAME = CLASS_NAME & ".get_Item"
On Error GoTo HandleError
  If IsNull(Index) Then
    Err.Raise 94, METHOD_NAME
  End If
  If IsNumeric(Index) Then
    If CLng(Index) < mProjectCount And Index > 0 Then
      Set Item = mVbaProject(CLng(Index))
      Exit Property
    End If
  End If
  If mProjectNames.Exists(Index) Then
    Set Item = mVbaProject(CLng(mProjectNames(Index)))
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
  Dim iLoader As ILoadedFromAddress
  Const OTPROJECTINFOOFFSET As Long = &H8
  
  On Error GoTo HandleError
  ReDim mVbaProject(1 To mProjectCount)
  ProjectNames.RemoveAll
  For i = 1 To mProjectCount
    Set mVbaProject(i) = New VbaProject
    Set iLoader = mVbaProject(i)
    iLoader.LoadFromAddress mProjectEntry(i).lpProjectInfo
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

#If Win64 Then
Private Property Get ILoadedFromAddress_OriginalAddress() As LongLong
#Else
Private Property Get ILoadedFromAddress_OriginalAddress() As Long
#End If
Const METHOD_NAME As String = "ILoadedFromAddress_OriginalAddress[" & CLASS_NAME & "]"
  ILoadedFromAddress_OriginalAddress = mOriginalAddress
End Property


#If Win64 Then
Private Function IMethodical_GetAddressByName(CallName As String) As LongLong
#Else
Private Function IMethodical_GetAddressByName(CallName As String) As Long
#End If
Const METHOD_NAME As String = "IMethodical_GetAddressByName[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  IMethodical_GetAddressByName = Me.GetAddressByName(CallName)
Exit Function
HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Function

Public Function GetAddressByName(CallName As String) As LongPtr
  ' Parse out the first part and see if any such project exists. _
    In order to avoid side effects, we're going to require the project _
    name to be specified.
  Const METHOD_NAME As String = CLASS_NAME & ".GetAddressByName"
  Dim thisProjectName As String, firstPeriod As Long, lastPeriod As Long, thisProject As VbaProject
  On Error GoTo HandleError
  firstPeriod = VBA.InStr(1, CallName, ".", vbBinaryCompare)
  If firstPeriod <= 0 Then
    'Invalid argument - raise error.
      Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "The project name must be specified in calls to GetAddressByName"
  End If
  lastPeriod = VBA.InStrRev(CallName, ".", -1, vbBinaryCompare)
  If firstPeriod = lastPeriod Then
    'Invalid argument - raise error.
      Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "The project, module, and method/procedure/function names must be specified in calls to GetAddressByName."
  End If
  thisProjectName = VBA.Left$(CallName, firstPeriod - 1)
  With ProjectNames
    If .Exists(thisProjectName) Then
      On Error GoTo HandleError
      Set thisProject = Me.Item(thisProjectName)
      GetAddressByName = thisProject.GetAddressByName(VBA.Mid$(CallName, firstPeriod + 1))
    Else
      Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "Project """ & thisProjectName & """ not found.", Err.HelpFile, Err.HelpContext
    End If
  End With
Exit Function
HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Function

