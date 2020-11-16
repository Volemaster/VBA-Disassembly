VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit


Private Type ProjectInfo
    ObjectInterfaces As LongPtr                '   0x00
    QWord1 As LongLong                         '   0x08
    lpD270 As LongPtr                          '   0x10    'Unknown
    lpObjectTable As LongPtr                   '   0x18
    FileNameDescriptor(0 To 5) As Byte         '   0x20
    FileName(0 To 2044) As Byte                '   0x23
    QWord2 As LongLong                         ' 0x0820
    lpE288 As LongPtr                          ' 0x0828    'Unknown
    lpPublicObjectDescriptors As LongPtr       ' 0x0830    'Pointer to array of PODs themselves, not an array of pointers.
    ObjectCount As LongPtr                     ' 0x0838
    ObjectInfo1Array As LongPtr                ' 0x0840
    QWord3 As LongLong                         ' 0x0848
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

Private mProjectInfo As ProjectInfo
Private mObjectTable As ObjectTable
Private mProjectInfoSize As Long
Private mObjectTableSize As Long
Private mProjectName As String
Private mFileName As String
Private mObjects As VbaObjects
Private mOriginalAddress As LongPtr

Private Const CLASS_NAME As String = "VbaProject"

Implements ILoadedFromAddress
Implements IMethodical

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


#If Win64 Then
  Private Sub ILoadedFromAddress_LoadFromAddress(ByVal lpAddress As LongLong, Optional Index As Long = 0)
#Else
  Private Sub ILoadedFromAddress_LoadFromAddress(ByVal lpAddress As Long, Optional Index As Long = 0)
#End If
Const METHOD_NAME As String = "ILoadedFromAddress_LoadFromAddress[" & CLASS_NAME & "]"
On Error GoTo HandleError
  ' Do not expose this to the main class interface
  mOriginalAddress = lpAddress
  Call LoadProjectInfo(lpAddress)
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
  Set mObjects = New VbaObjects
  errorMessage = "Invalid lpProjectInfo (bad lpObjectTable)."
  Call ObjectLoader.LoadFromAddress(mProjectInfo.lpObjectTable)
  errorMessage = "Invalid ObjectTable (lpProjectInfo mismatch)."
  If lpProjectInfo <> ObjectData.lpProjectInfo Then
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

Private Property Get ObjectLoader() As ILoadedFromAddress
Const METHOD_NAME As String = CLASS_NAME & ".get_ObjectLoader"
On Error GoTo HandleError
  Set ObjectLoader = mObjects
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get ObjectData() As IObjectTableData
Const METHOD_NAME As String = CLASS_NAME & ".get_ObjectData"
On Error GoTo HandleError
  Set ObjectData = mObjects
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Public Property Get FileName() As String
Const METHOD_NAME As String = CLASS_NAME & ".get_FileName"
On Error GoTo HandleError
  If mFileName = vbNullString Then
    mFileName = RTrimNull(CStr(mProjectInfo.FileName))
  End If
  FileName = mFileName
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Public Property Get Name() As String
  Const METHOD_NAME As String = CLASS_NAME & ".get_Name"
  On Error GoTo HandleError
  Name = ObjectData.ProjectName
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


#If Win64 Then
Private Property Get ILoadedFromAddress_OriginalAddress() As LongLong
#Else
Private Property Get ILoadedFromAddress_OriginalAddress() As Long
#End If
  Const METHOD_NAME As String = "ILoadedFromAddress_OriginalAddress[" & CLASS_NAME & "]"
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
  ' If a three-part name is specified, verify that the first part _
    matches this project name, then pass the next two parts down. _
    If not, pass the arguments to VbaObjects as is.
  Const METHOD_NAME As String = CLASS_NAME & ".GetAddressByName"
  Dim firstPeriod As Long, lastPeriod As Long, callToPass As String
  firstPeriod = VBA.InStr(1, CallName, ".", vbBinaryCompare)
  If firstPeriod <= 0 Then
    Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "The module name must be specified in calls to GetAddressByName"
    Exit Function
  End If
  lastPeriod = VBA.InStrRev(CallName, ".", -1, vbBinaryCompare)
  If firstPeriod <> lastPeriod Then
    If VBA.Left$(CallName, firstPeriod - 1) <> Me.Name Then
      Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "The specified call name to the object " & Me.Name & " was invalid (" & CallName & ").", Err.HelpFile, Err.HelpContext
      Exit Function
    End If
    callToPass = VBA.Mid$(CallName, firstPeriod + 1)
  Else
    callToPass = CallName
  End If
  On Error GoTo HandleError
  GetAddressByName = mObjects.GetAddressByName(callToPass)
  Exit Function

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Function

