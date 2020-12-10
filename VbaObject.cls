VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    lpObjectMethodData As LongPtr              '   0x18    'see ObjectMethodData
    QWord1 As LongLong                         '   0x20
    QWord2 As LongLong                         '   0x28
    lpPublicObjectDescriptor As LongPtr        '   0x30    'Points directly to the object descriptor itself
    Ptr3 As LongPtr                            '   0x38    'Points to something that is generally null
    MethodCount  As Integer                    '   0x40
    MethodCount2 As Integer                    '   0x42    'Count of compiled methods?
    DWord2 As Long                             '   0x44
    lpMethodInfoPointers As LongPtr            '   0x48    'Array of pointers to MethodInfo structures
    Word1 As Integer                           '   0x50    'Constants in constant pool?
    Word2 As Integer                           '   0x52    'Constants to allocate in constant pool?
    DWord3 As Long                             '   0x54
    Ptr4 As LongPtr                            '   0x58
    lpStringTable As LongPtr                   '   0x60    'Array of pointers to literal strings used in module?
    QWord3 As LongLong                         '   0x68    'Max size of ObjectInfo = 0x70
End Type

Private Type ObjectMethodData
  lpPtr1 As LongPtr
  lpObjectInfo As LongPtr
  Reserved1 As LongLong
  Null1 As LongLong
  Null2 As LongLong
  lpMethodDataArray As LongPtr               '   0x28    'Array of pointers to MethodData
  lpMethodData2Array As LongPtr              '   0x30    'Array of pointers to MethodData2
End Type

Private mOriginalAddress As LongPtr
Private mOriginalObjectInfoAddress As LongPtr
Private mObjectInfo As ObjectInfo
Private mPublicObjectDescriptor As PublicObjectDescriptor
Private mObjectInfoSize As Long
Private mPublicObjectDescriptorSize As Long
Private mModuleName As String
Private mMethods() As IVbaMethod

Private mWeakPointerToParent As LongPtr

Private mlpMethodDataArray As LongPtr, mlpMethodData2Array As LongPtr, mlpMethodNamePtrArray As LongPtr
Private mMethodCount As Long
Private mObjectMethodData As ObjectMethodData
Private mObjectMethodDataSize As Long

Private Const CLASS_NAME As String = "VbaObject"

Implements IVbaObject
Implements IVbaMethods
Implements IChildOfDisposable

Public Property Get PublicObjectDescriptorSize() As Long
  PublicObjectDescriptorSize = mPublicObjectDescriptorSize
End Property

Private Sub IChildOfDisposable_DisposeOfParent()
  Set Me.ChildOfDisposable.Parent = Nothing
End Sub

Private Property Set IChildOfDisposable_Parent(RHS As Object)
  mWeakPointerToParent = ObjPtr(RHS)
End Property

Private Property Get IVbaObject_Methods() As IVbaMethods
  Set IVbaObject_Methods = Me
End Property

Private Property Get IVbaMethods_Parent() As IVbaObject
  Set IVbaMethods_Parent = Me
End Property

Private Sub Class_Initialize()
  'Alas that we can't use a static class variable.
  Const METHOD_NAME As String = CLASS_NAME & ".Class_Initialize"
  On Error GoTo HandleError
  mObjectInfoSize = LenB(mObjectInfo)
  mPublicObjectDescriptorSize = LenB(mPublicObjectDescriptor)
  mObjectMethodDataSize = LenB(mObjectMethodData)
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Public Sub Setup(lpPublicObjectDescriptor As LongPtr, Optional Parent As IVbaObjects = Nothing)
Const METHOD_NAME = CLASS_NAME & ".Setup"
On Error GoTo HandleError
  Set Me.ChildOfDisposable.Parent = Parent
  Call LoadPublicObjectDescriptor(lpPublicObjectDescriptor)
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

Private Property Get IVbaObject_Name() As String
Const METHOD_NAME As String = "IVbaObject_Name[" & CLASS_NAME & "]"
On Error GoTo HandleError
  IVbaObject_Name = mModuleName: Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Sub LoadObjectInfo(lpAddress As LongPtr)
Const METHOD_NAME As String = CLASS_NAME & ".LoadObjectInfo"
On Error GoTo InvalidObject
  Memory.FollowPointer VarPtr(mObjectInfo), lpAddress, mObjectInfoSize
  'Validate the objects
  With mObjectInfo
    Debug.Assert .lpPublicObjectDescriptor = mOriginalAddress
    If .lpPublicObjectDescriptor <> mOriginalAddress Then
      GoTo InvalidObject
      Exit Sub
    End If
  End With
  mOriginalObjectInfoAddress = lpAddress
  LoadObjectMethodData
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

Private Sub LoadObjectMethodData()
  Memory.FollowPointer VarPtr(mObjectMethodData), mObjectInfo.lpObjectMethodData, mObjectMethodDataSize
  Call SetupMethods(mObjectMethodData.lpMethodDataArray, mObjectMethodData.lpMethodData2Array, mPublicObjectDescriptor.lpMethodNamePtrArray, CLng(mObjectInfo.MethodCount))
End Sub

Public Sub SetupMethods(ByVal lpMethodDataArray As LongPtr, ByVal lpMethodData2Array As LongPtr, ByVal lpMethodNamePtrArray As LongPtr, Optional MethodCount As Long = 0)
    Const METHOD_NAME = CLASS_NAME & ".SetupMethods"
On Error GoTo HandleError
  mlpMethodDataArray = lpMethodDataArray
  mlpMethodData2Array = lpMethodData2Array
  mlpMethodNamePtrArray = lpMethodNamePtrArray
  mMethodCount = MethodCount
  If MethodCount > 0 Then
    ReDim mMethods(0 To MethodCount - 1)
  Else
    Erase mMethods
  End If
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Property Get IVbaMethods_Count() As Long
Const METHOD_NAME As String = "IVbaMethods_Count[" & CLASS_NAME & "]"
On Error GoTo HandleError
  IVbaMethods_Count = mMethodCount: Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IVbaMethods_Item(Index As Long) As IVbaMethod
Const METHOD_NAME As String = "IVbaMethods_Item[" & CLASS_NAME & "]"
On Error GoTo HandleError
  If Index < Me.Count Then
    Dim lpMyMethodData As LongPtr, lpMyMethodData2 As LongPtr, lpMyMethodName As LongPtr
    Memory.FollowPointer VarPtr(lpMyMethodData), mlpMethodDataArray + (POINTER_SIZE * Index), POINTER_SIZE
    Memory.FollowPointer VarPtr(lpMyMethodData2), mlpMethodData2Array + (POINTER_SIZE * Index), POINTER_SIZE
    Memory.FollowPointer VarPtr(lpMyMethodName), mlpMethodNamePtrArray + (POINTER_SIZE * Index), POINTER_SIZE
    Set IVbaMethods_Item = PrivateFactory.CreateVbaMethod(lpMyMethodData, lpMyMethodData2, lpMyMethodName)
    'This is to maintain the ability to notify the method that it's parent has been disposed
    Set mMethods(Index) = IVbaMethods_Item
  End If
Exit Property
HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Function IVbaMethods_Find(Identifier As IMethodIdentifier, Optional Criteria As MethodComparisonResult = MethodComparisonResult.Matches) As IVbaMethod
Const METHOD_NAME As String = "IVbaMethods_Find[" & CLASS_NAME & "]"
On Error GoTo HandleError
  Dim i As Long
  Dim theMethod As IVbaMethod
  Dim compareResult As MethodComparisonResult
  With Me
    For i = 0 To .Count - 1
      Set theMethod = .Item(i)
      compareResult = theMethod.Compare(Identifier)
      If (compareResult And Criteria) = Criteria Then
        Set IVbaMethods_Find = theMethod
        Exit Function
      End If
    Next
  End With
Exit Function
HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Function

Public Property Get Methods() As IVbaMethods
  Const METHOD_NAME As String = CLASS_NAME & ".get_Methods"
  On Error GoTo HandleError
  Set Methods = IVbaObject_Methods
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
  Name = Me.Methods.Parent.Name
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Public Property Get Count() As Long
  Const METHOD_NAME As String = CLASS_NAME & ".get_Count"
  On Error GoTo HandleError
  Count = Me.Methods.Count
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Public Property Get Item(Index As Long) As IVbaMethod
  Const METHOD_NAME As String = CLASS_NAME & ".get_Item"
  On Error GoTo HandleError
  Set Item = Me.Methods.Item(Index)
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IVbaObject_Parent() As IVbaObjects
  If mWeakPointerToParent = 0 Then Exit Property
  Set IVbaObject_Parent = PrivateFactory.GetObjectFromWeakPointer(mWeakPointerToParent)
End Property

Private Sub Class_Terminate()
  Dim i As Long
  Dim DC As IChildOfDisposable
  For i = 1 To IVbaMethods_Count
    Set DC = mMethods(i - 1)
    If Not DC Is Nothing Then
      DC.DisposeOfParent
    End If
  Next
End Sub

Public Property Get ChildOfDisposable() As IChildOfDisposable
  Set ChildOfDisposable = Me
End Property
