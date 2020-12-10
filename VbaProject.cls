VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 0
Option Explicit


Private Type ProjectInfo
    ObjectInterfaces As LongPtr                '   0x00
    QWord1 As LongLong                         '   0x08
    lpD270 As LongPtr                          '   0x10    'Unknown
    lpObjectTable As LongPtr                   '   0x18
    FileNameDescriptor(0 To 5) As Byte         '   0x20
    FileName(0 To 4090) As Byte                '   0x23
    lpE288 As LongPtr                          ' 0x0828    'Unknown
    lpPublicObjectDescriptorArray As LongPtr   ' 0x0830    'Pointer to array of pointers to PODs themselves, not an array of PODs.
    ObjectCount As LongPtr                     ' 0x0838
    ObjectInfo1Array As LongPtr                ' 0x0840
    QWord2 As LongLong                         ' 0x0848
End Type

Private Type ObjectTable
    lpX1 As LongPtr                            '   0x00
    lpProjectInfo As LongPtr                   '   0x08
    lpSecondaryProjectInfo As LongPtr          '   0x10
    Reserved As LongLong                       '   0x18
    Null1 As LongLong                          '   0x20
    lp6CC0 As LongPtr                          '   0x28
    GUID As GUID                               '   0x30
    CompiledState As Integer                   '   0x40
    ObjectCount As Integer                     '   0x42
    CompiledObjects As Integer                 '   0x44
    ObjectsInUse As Integer                    '   0x46
    lpPublicObjectDescriptors As LongPtr       '   0x48
    ReferenceCount As LongLong                 '   0x50
    lpReferenceTable As LongPtr                '   0x58
    QWord1 As LongLong                         '   0x60
    lpProjectName As LongPtr                   '   0x68    'Pointer to char[]
    PrimaryLanguageCodeID As Long              '   0x70
    SecondaryLangagueCodeID As Long            '   0x74
    Ptr1 As LongPtr                            '   0x78
    QWord2 As LongLong                         '   0x80
    QWord3 As LongLong                         '   0x88
    QWord4 As LongLong                         '   0x90
    QWord5 As LongLong                         '   0x98
    QWord6 As LongLong                         '   0xA0
    QWord7 As LongLong                         '   0xA8
    QWord8 As LongLong                         '   0xB0
    QWord9 As LongLong                         '   0xB8
End Type

Private mModuleNames As Scripting.Dictionary
Private mModules() As IVbaObject
Private mObjectsLoaded As Boolean

Private mProjectInfo As ProjectInfo
Private mObjectTable As ObjectTable
Private mProjectInfoSize As Long
Private mObjectTableSize As Long
Private mProjectName As String
Private mFileName As String
Private mObjects As IVbaObjects
Private mOriginalAddress As LongPtr

Private mWeakPointerToParent As LongPtr

Private Const CLASS_NAME As String = "VbaProject"

Implements IVbaProject
Implements IVbaObjects
Implements IChildOfDisposable

Private Sub Class_Initialize()
Const METHOD_NAME As String = CLASS_NAME & ".Class_Initialize"
On Error GoTo HandleError
  mProjectInfoSize = LenB(mProjectInfo)
  mObjectTableSize = LenB(mObjectTable)
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Public Sub Setup(lpProjectInfo As LongPtr, Optional Parent As IVbaProjects = Nothing)
Const METHOD_NAME As String = CLASS_NAME & ".Setup"
On Error GoTo HandleError
  ' Do not expose this to the main class interface
  mOriginalAddress = lpProjectInfo
  Set Me.ChildOfDisposable.Parent = Parent
  Call LoadProjectInfo(lpProjectInfo)
  Call LoadObjectTable(mProjectInfo.lpObjectTable)
  Call LoadObjects
Exit Sub
HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Sub LoadObjectTable(lpObjectTable As LongPtr)
  Const METHOD_NAME = CLASS_NAME & ".LoadObjectTable"
  On Error GoTo HandleError
  Memory.FollowPointer VarPtr(mObjectTable), lpObjectTable, mObjectTableSize
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Sub LoadProjectInfo(ByVal lpProjectInfo As LongPtr)
  Dim i As Long
  Dim errorMessage As String
  Const METHOD_NAME As String = CLASS_NAME & ".LoadProjectInfo"
  On Error GoTo HandleError
  errorMessage = "Invalid lpProjectInfo."
  Memory.FollowPointer VarPtr(mProjectInfo), lpProjectInfo, mProjectInfoSize
  errorMessage = "Invalid lpProjectInfo (bad lpObjectTable)."
  LoadObjectTable mProjectInfo.lpObjectTable
  errorMessage = "Invalid ObjectTable (lpProjectInfo mismatch)."
  If lpProjectInfo <> mObjectTable.lpProjectInfo Then
    Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT
  End If
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Sub LoadObjects()
  Const METHOD_NAME As String = CLASS_NAME & ".LoadObjects"
  On Error GoTo HandleError
  Dim i As Long
  Dim Count As Long
  Count = Me.Objects.Count
  If Count > 0 Then
    ReDim mModules(0 To Count - 1)
    If mModuleNames Is Nothing Then
      Set mModuleNames = New Scripting.Dictionary
      mModuleNames.CompareMode = TextCompare
    Else
      mModuleNames.RemoveAll
    End If
  Else
    Erase mModules
    Exit Sub
  End If
  For i = 0 To Objects.Count - 1
    Set mModules(i) = PrivateFactory.CreateVbaObject(mObjectTable.lpPublicObjectDescriptors, i)
    mModuleNames.Add mModules(i).Name, i
  Next
  mObjectsLoaded = True
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Public Property Get Objects() As IVbaObjects
  Const METHOD_NAME As String = CLASS_NAME & ".get_Objects"
  On Error GoTo HandleError
  Set Objects = IVbaProject_Objects
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Sub IChildOfDisposable_DisposeOfParent()
  Set IChildOfDisposable_Parent = Nothing
End Sub

Private Property Set IChildOfDisposable_Parent(Value As Object)
  mWeakPointerToParent = ObjPtr(Value)
End Property

Private Property Get IVbaObjects_Count() As Long
  Const METHOD_NAME = "IVbaObjects_Count[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  IVbaObjects_Count = mObjectTable.ObjectCount
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property


Private Property Get IVbaObjects_Item(Index As Variant) As IVbaObject
  Dim moduleName As String
  Const METHOD_NAME = "IVbaObjects_Item[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  Dim dwIndex As Long
  If IsNull(Index) Then
    Err.Raise 94, METHOD_NAME
  End If
  If Not mObjectsLoaded Then Call LoadObjects
  
  If IsNumeric(Index) Then
    dwIndex = CLng(Index)
'    Set IVbaObjects_Item = PrivateFactory.CreateVbaObject(mObjectTable.lpPublicObjectDescriptors, dwIndex)
    If CLng(Index) < Me.Objects.Count And Index >= 0 Then
      Set IVbaObjects_Item = mModules(Index)
      Exit Property
      'moduleName = mModuleNames(Index)
    End If
  Else
    moduleName = CStr(Index)
  End If
  If mModuleNames.Exists(moduleName) Then
    Set IVbaObjects_Item = mModules(mModuleNames(moduleName))
    Exit Property
  End If
  If IsNumeric(Index) Then
    Err.Raise 381, METHOD_NAME, "Invalid module specified: the zero-based identifier you specified (" & Index & ") does not exist."
  Else
    Err.Raise 381, METHOD_NAME, "Invalid module specified: the module name """ & Index & """ could not be found."
  End If
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property


Private Property Get IVbaObjects_Parent() As IVbaProject
  Set IVbaObjects_Parent = Me
End Property

Private Property Get IVbaProject_FileName() As String
  Const METHOD_NAME = "IVbaProject_FileName[" & CLASS_NAME & "]"
On Error GoTo HandleError
  If mFileName = vbNullString Then
    mFileName = RTrimNull(CStr(mProjectInfo.FileName))
  End If
  IVbaProject_FileName = mFileName
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IVbaProject_Name() As String
  Const METHOD_NAME = "IVbaProject_Name[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  If mProjectName = vbNullString Then
    mProjectName = GetProjectName()
  End If
  IVbaProject_Name = mProjectName
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Function GetProjectName() As String
  Dim n(0 To 1023) As Byte, l As Long
  Dim result As String
  Const METHOD_NAME As String = CLASS_NAME & ".GetProjectName"
  On Error GoTo HandleError
  Memory.FollowPointer VarPtr(n(0)), mObjectTable.lpProjectName, MAX_VBA_OBJECT_NAME_LENGTH
  GetProjectName = RTrimNull(VBA.StrConv(n, vbUnicode))
  Exit Function

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Function

Private Property Get IVbaProject_Objects() As IVbaObjects
  Const METHOD_NAME = "IVbaProject_Objects[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  Set IVbaProject_Objects = Me
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IVbaProject_Parent() As IVbaProjects
  If mWeakPointerToParent = 0 Then Exit Property
  Set IVbaProject_Parent = PrivateFactory.GetObjectFromWeakPointer(mWeakPointerToParent)
End Property

Private Sub Class_Terminate()
  Dim i As Long
  Dim DC As IChildOfDisposable
  For i = 1 To IVbaObjects_Count
    Set DC = mModules(i - 1)
    If Not DC Is Nothing Then
      DC.DisposeOfParent
    End If
  Next
End Sub

Public Property Get ChildOfDisposable() As IChildOfDisposable
  Set ChildOfDisposable = Me
End Property
