VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type PublicObjectDescriptor
    lpObjectInfo As LongPtr                    '   0x00
    Reserved1 As LongLong                      '   0x08
    lpPublicVariableIntegers As LongPtr        '   0x10
    lpStaticVariableIntegers As LongPtr        '   0x18
    lpDataPublicVariables As LongPtr           '   0x20
    lpDataStaticVariables As LongPtr           '   0x28
    lpModuleName As LongPtr                    '   0x30    'Pointer to char[]
    MethodCount As LongLong                    '   0x38
    lpMethodNamePtrArray As LongPtr            '   0x40    'Array of pointers to char[]
    OffsetToStaticVariables As Long            '   0x48
    Flags As Long                              '   0x4C
    Null1 As Long                              '   0x50
End Type

Private Type ObjectInfo
    ReferenceCount As Integer                  '   0x00
    ObjectIndex As Integer                     '   0x02
    DWord1 As Long                             '   0x04
    lpObjectTable As LongPtr                   '   0x08
    Ptr1 As LongPtr                            '   0x10
    Ptr2 As LongPtr                            '   0x18
    QWord1 As LongLong                         '   0x20
    QWord2 As LongLong                         '   0x28
    lpPublicObjectDescriptor As LongPtr        '   0x30    'Points directly to the object descriptor itself
    Ptr3 As LongPtr                            '   0x38
    MethodCount  As Integer                    '   0x40
    MethodCount2 As Integer                    '   0x42    'Potentially unreliable?
    DWord2 As Long                             '   0x44
    lpMethodInfoPointers As LongPtr            '   0x48    'Array of pointers to MethodInfo structures
    Word1 As Integer                           '   0x50    'Constants in constant pool?
    Word2 As Integer                           '   0x52    'Constants to allocate in constant pool?
    DWord3 As Long                             '   0x54
    Ptr4 As LongPtr                            '   0x58
    Ptr5 As LongPtr                            '   0x60
    QWord3 As LongLong                         '   0x68
    QWord4 As LongLong                         '   0x70
    QWord5 As LongLong                         '   0x78
    QWord6 As LongLong                         '   0x80
    Ptr6 As LongPtr                            '   0x88
End Type

Private mOriginalAddress As LongPtr
Private mOriginalObjectInfoAddress As LongPtr
Private mObjectInfo As ObjectInfo
Private mPublicObjectDescriptor As PublicObjectDescriptor
Private mObjectInfoSize As Long
Private mPublicObjectDescriptorSize As Long
Private mModuleName As String

Private mMethodData As Scripting.Dictionary

Private Const CLASS_NAME As String = "VbaModule"

Implements ILoadedFromAddress
Implements IMethodical

Private Sub Class_Initialize()
  'Alas that we can't use a static class variable.
  Const METHOD_NAME As String = CLASS_NAME & ".Class_Initialize"
  On Error GoTo HandleError
  mObjectInfoSize = LenB(mObjectInfo)
  mPublicObjectDescriptorSize = LenB(mPublicObjectDescriptor)
  Set mMethodData = New Scripting.Dictionary
  mMethodData.CompareMode = TextCompare
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

Const METHOD_NAME As String = "get_ILoadedFromAddress_LoadFromAddress[" & CLASS_NAME & "]"
On Error GoTo HandleError
  Call LoadPublicObjectDescriptor(lpAddress + (Index * mPublicObjectDescriptorSize))
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Property Get MethodData() As Scripting.Dictionary
Const METHOD_NAME As String = CLASS_NAME & ".get_MethodData"
On Error GoTo HandleError
  If MethodCount <> mMethodData.Count Then
    'Lazy loading again!
    LoadMethods
  End If
  Set MethodData = mMethodData
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Sub LoadMethods()
Const METHOD_NAME As String = CLASS_NAME & ".LoadMethods"
On Error GoTo HandleError
  Dim i As Long
  Dim methodName As String
  Dim methodAddress As LongPtr
  Dim namePointers() As LongPtr
  Dim addressPointers() As LongPtr
  Dim n(0 To MAX_VBA_OBJECT_NAME_LENGTH - 1) As Byte
  
  If MethodCount <= 0 Then Exit Sub
  ReDim namePointers(0 To MethodCount - 1)
  ReDim addressPointers(0 To MethodCount - 1)
  Memory.FollowPointer VarPtr(namePointers(0)), mPublicObjectDescriptor.lpMethodNamePtrArray, POINTER_SIZE * MethodCount
  Memory.FollowPointer VarPtr(addressPointers(0)), mObjectInfo.lpMethodInfoPointers, POINTER_SIZE * MethodCount
  With mMethodData
    .RemoveAll
    For i = 1 To MethodCount
      Memory.FollowPointer VarPtr(n(0)), namePointers(i - 1), MAX_VBA_OBJECT_NAME_LENGTH
      mMethodData.Add (RTrimNull(VBA.StrConv(n, vbUnicode))), addressPointers(i - 1)
    Next
  End With
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Sub LoadPublicObjectDescriptor(lpAddress As LongPtr)
Const METHOD_NAME As String = CLASS_NAME & ".LoadPublicObjectDescriptor"
On Error GoTo HandleError
  Dim n(0 To 1023) As Byte
  mOriginalAddress = lpAddress
  Memory.FollowPointer VarPtr(mPublicObjectDescriptor), lpAddress, mPublicObjectDescriptorSize
  With mPublicObjectDescriptor
    Memory.FollowPointer VarPtr(n(0)), .lpModuleName, MAX_VBA_OBJECT_NAME_LENGTH
    mModuleName = RTrimNull(VBA.StrConv(n, vbUnicode))
    If .Reserved1 <> -1 Then '&HFFFFFFFF in x86, &HFFFFFFFFFFFFFFFF in x64
      'Invalid POD. - Always -1
      GoTo InvalidObjectInfo
    ElseIf .MethodCount < 0 Or .MethodCount > &H100 Then
      'Invalid POD - Max methods allowed is 255 (&HFF)
      GoTo InvalidObjectInfo
    End If
    Call LoadObjectInfo(.lpObjectInfo)
      ' Error handling in LoadObjectInfo ensures that this won't be loaded if it's invalid.
  
  End With
Exit Sub
InvalidObjectInfo:
  
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Public Property Get MethodCount() As Long
Const METHOD_NAME As String = CLASS_NAME & ".get_MethodCount"
On Error GoTo HandleError
  MethodCount = mObjectInfo.MethodCount: Exit Property

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
  Name = mModuleName: Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Sub LoadObjectInfo(lpAddress As LongPtr)
Const METHOD_NAME As String = CLASS_NAME & ".LoadObjectInfo"
On Error GoTo HandleError
  Memory.FollowPointer VarPtr(mObjectInfo), lpAddress, mObjectInfoSize
  'Validate the objects
  With mObjectInfo
    Debug.Assert .lpPublicObjectDescriptor = mOriginalAddress
    If .lpPublicObjectDescriptor <> mOriginalAddress Then
      GoTo InvalidObject
      Exit Sub
    End If
    Debug.Assert .MethodCount = mPublicObjectDescriptor.MethodCount
  End With
  mOriginalObjectInfoAddress = lpAddress
  Exit Sub
HandleError:
  mOriginalObjectInfoAddress = 0
  mPublicObjectDescriptor.lpObjectInfo = 0
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
  Exit Sub
InvalidObject:
  mOriginalObjectInfoAddress = 0
  mPublicObjectDescriptor.lpObjectInfo = 0
End Sub

#If Win64 Then
Private Property Get ILoadedFromAddress_OriginalAddress() As LongLong
#Else
Private Property Get ILoadedFromAddress_OriginalAddress() As Long
#End If
Const METHOD_NAME As String = "get_ILoadedFromAddress_OriginalAddress[" & CLASS_NAME & "]"
On Error GoTo HandleError
  ILoadedFromAddress_OriginalAddress = mOriginalAddress: Exit Property

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
Const METHOD_NAME As String = "get_IMethodical_GetAddressByName[" & CLASS_NAME & "]"
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
  ' If a two-part name is specified, verify that the first part _
    matches this module name, then look for the final piece. _
    If not, search as is.
  Dim firstPeriod As Long, callToPass As String
  firstPeriod = VBA.InStr(1, CallName, ".", vbBinaryCompare)
  If firstPeriod > 0 Then
    If VBA.Left$(CallName, firstPeriod - 1) <> Me.Name Then
      Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "The specified call name to the object " & Me.Name & " was invalid (" & CallName & ").", Err.HelpFile, Err.HelpContext
      Exit Function
    End If
    callToPass = VBA.Mid$(CallName, firstPeriod + 1)
  Else
    callToPass = CallName
  End If
  On Error GoTo HandleError
  If MethodData.Exists(callToPass) Then
    GetAddressByName = MethodData.Item(callToPass)
    Exit Function
  End If
  On Error GoTo 0
  Err.Raise INVALID_PROCEDURE_CALL_OR_ARGUMENT, METHOD_NAME, "Method, function, or procedure not found.", Err.HelpFile, Err.HelpContext
  Exit Function

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Function
