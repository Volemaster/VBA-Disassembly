VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 0
Option Explicit


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

Private Const CLASS_NAME As String = "VbaObjects"

Private mObjectTable As ObjectTable
Private mModules As Scripting.Dictionary
Private mModuleNames() As String
Private mObjectTableSize As Long
Private mProjectName As String
Private mObjectsLoaded As Boolean
Private mOriginalAddress As LongPtr

Implements ILoadedFromAddress
Implements IObjectTableData
Implements IMethodical

Private Sub Class_Initialize()
Const METHOD_NAME = CLASS_NAME & ".Class_Initialize"
On Error GoTo HandleError
  mObjectTableSize = LenB(mObjectTable)
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

#If Win64 Then
  Private Sub ILoadedFromAddress_LoadFromAddress(ByVal lpAddress As LongLong, Optional Index As Long = 0)
#Else
  Private Sub ILoadedFromAddress_LoadFromAddress(ByVal lpAddress As Long, Optional Index As Long = 0)
#End If

Const METHOD_NAME = "ILoadedFromAddress_LoadFromAddress[" & CLASS_NAME & "]"
On Error GoTo HandleError
  mOriginalAddress = lpAddress
  LoadObjectTable lpAddress
Exit Sub
HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

#If Win64 Then
  Private Property Get IObjectTableData_lpProjectInfo() As LongLong
#Else
  Private Property Get IObjectTableData_lpProjectInfo() As Long
#End If
  Const METHOD_NAME = "IObjectTableData_lpProjectInfo [" & CLASS_NAME & "]"
  On Error GoTo HandleError
  IObjectTableData_lpProjectInfo = mObjectTable.lpProjectInfo
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Sub LoadObjectTable(lpAddress As LongPtr)
  Const METHOD_NAME = CLASS_NAME & ".LoadObjectTable"
  On Error GoTo HandleError
  Memory.FollowPointer VarPtr(mObjectTable), lpAddress, mObjectTableSize
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Property Get Modules() As Scripting.Dictionary
  Const METHOD_NAME = CLASS_NAME & ".get_Modules"
  On Error GoTo HandleError
  If mModules Is Nothing Then
    Set mModules = New Scripting.Dictionary
    mModules.CompareMode = TextCompare
  End If
  Set Modules = mModules
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Public Property Get Count() As Long
  Const METHOD_NAME = CLASS_NAME & ".get_Count"
  On Error GoTo HandleError
  Count = mObjectTable.ObjectCount
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
  Const METHOD_NAME = CLASS_NAME & ".GetProjectName"
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

Private Property Get IObjectTableData_ProjectName() As String
  Const METHOD_NAME = "get_IObjectTableData_ProjectName[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  If mProjectName = vbNullString Then
    mProjectName = GetProjectName()
  End If
  IObjectTableData_ProjectName = mProjectName
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Public Property Get Item(Index As Variant) As VbaModule
  Dim moduleName As String
  Const METHOD_NAME = CLASS_NAME & ".get_Item"
  On Error GoTo HandleError
  If IsNull(Index) Then
    Err.Raise 94, METHOD_NAME
  End If
  If Not mObjectsLoaded Then Call LoadObjects
  
  If IsNumeric(Index) Then
    If CLng(Index) < Count And Index > 0 Then
      moduleName = mModuleNames(Index)
    End If
  Else
    moduleName = Index
  End If
  If Modules.Exists(moduleName) Then
    Set Item = Modules(moduleName)
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

Private Sub LoadObjects()
  Dim i As Long
  Dim thisModule As VbaModule, loader As ILoadedFromAddress
  Const METHOD_NAME = CLASS_NAME & ".LoadObjects"
  On Error GoTo HandleError
  If Count <= 0 Then
    mObjectsLoaded = True
    Exit Sub
  End If
  Modules.RemoveAll
  ReDim mModuleNames(0 To Count - 1)
  For i = 1 To Count
    Set thisModule = New VbaModule
    Set loader = thisModule
    Call loader.LoadFromAddress(mObjectTable.lpPublicObjectDescriptors, i - 1)
    mModuleNames(i - 1) = thisModule.Name
    Modules.Add mModuleNames(i - 1), thisModule
  Next
  
Exit Sub

HandleError:
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
  Const METHOD_NAME = "get_ILoadedFromAddress_OriginalAddress[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  ILoadedFromAddress_OriginalAddress = mOriginalAddress
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

#If Win64 Then
Private Function IMethodical_GetAddressByName(CallName As String) As LongLong
#Else
Private Function IMethodical_GetAddressByName(CallName As String) As Long
#End If
  Const METHOD_NAME = "get_IMethodical_GetAddressByName[" & CLASS_NAME & "]"
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
  Const METHOD_NAME As String = CLASS_NAME & ".GetAddressByName"
  On Error GoTo HandleError
  
  Dim firstPeriod As Long, callToPass As String, targetObjectName As String
  Dim thisModule As VbaModule
  firstPeriod = VBA.InStr(1, CallName, ".", vbBinaryCompare)
  If firstPeriod <= 0 Then
    Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "The module name must be specified in calls to GetAddressByName."
    Exit Function
  End If
  targetObjectName = VBA.Left$(CallName, firstPeriod - 1)
  On Error GoTo HandleError
  Set thisModule = Me.Item(targetObjectName)
  callToPass = VBA.Mid$(CallName, firstPeriod + 1)
  GetAddressByName = thisModule.GetAddressByName(callToPass)
  Exit Function

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Function

