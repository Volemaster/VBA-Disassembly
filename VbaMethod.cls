VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mAddresses([_atMin] To [_atMax]) As LongPtr

Private Type ObjectMethodData
  lpPtr1 As LongPtr
  lpObjectInfo As LongPtr
  Reserved1 As LongLong
  Null1 As LongLong
  Null2 As LongLong
  lpMethodDataArray As LongPtr               '   0x28    'Array of pointers to MethodData
  lpMethodData2Array As LongPtr              '   0x30    'Array of pointers to MethodData2
End Type

Private Enum MethodDataFlag2
  ProcTypeMask = &HFF00
  IsMethod = &H6000
  IsProperty = &H6800
  IsDeclare = &H7000
End Enum

Private Enum ArgIndex
  ReturnArg = -1
  Argument01 = 0
  Argument02 = 1
  Argument03 = 2
  Argument04 = 3
  Argument05 = 4
  Argument06 = 5
  Argument07 = 6
  Argument08 = 7
  Argument09 = 8
  Argument10 = 9
  Argument11 = 10
  [_MaxArgs] = 10
  [_MaxTheoreticalSize] = &H9C - 1  ' Each UDT pads out the remainder of the &H10 byte line with 0s.
  [_AlignmentOffset] = &HC
End Enum
' Note that VBA only permits 10 arguments *however* _
  since class calls always return HResult, an additional _
  ByRef argument will be allocated at the end of the _
  argument array when _
  (MethodData.dwMethodFlags AND MethodFlags.ReturnLastArg = MethodFlags.ReturnLastArg)

Private Type MethodData
  dwMethodFlag As MethodFlags                '   0x00    'Very complicated. See MethodFlags enum
  wReserved As Integer                       '   0x04    'Always 0xFFFF
  wNull As Integer                           '   0x06    'Always 0x0
  qwNull1 As LongLong                        '   0x08
  wSequence As Integer                       '   0x10    'Some sort of unique sequence identifier for the method. VTable placement?
  wFlag2 As Integer                          '   0x12    'MethodDataFlag2
  dwNull2 As Long                            '   0x14    '
  lpArgumentNameArray As LongPtr             '   0x18    'Array of pointers to char[] names
  qwNull3 As LongLong                        '   0x20
  qwNull4 As LongLong                        '   0x28
  dwNull5 As Long                            '   0x30
  ArgTypes(ArgIndex.ReturnArg To _
    ArgIndex.[_MaxTheoreticalSize]) As Integer          '   0x34    'Only valid up to the actual number of args/returns
End Type                                                 'Using an array for all due to byte alignment issues in VBA

Private Type MethodData2
  qwMethodIndex As LongLong                  '   0x00
  lpStackInfo As LongPtr                     '   0x08
  qwNull1 As LongLong                        '   0x10
  lpMethodInfo As LongPtr                    '   0x18
End Type


Private Type MethodInfo2Header
  SizeInBytes As Integer                      '   0x00
  Unknown1(12) As Byte                        '   0x02
  'The structure is extended by unknown content up to the total size given by SizeInBytes
  'MethodInfo3Header follows MethodInfo2 immediately. It is then byte aligned on 4-byte boundaries.
End Type

Private Type MethodInfo3Header
  SizeInBytes As Integer                      '   0x00
  Null1 As LongLong                           '   0x02
  VariableCount As Integer                    '   0x0A
  dwUnknown1 As Long                          '   0x0C
  wUnknown2 As Integer                        '   0x0E
  wUnknown3 As Integer                        '   0x10
  'an array of MethodVariableInfo(VariableCount) immediately follows this header.
End Type

Private Type MethodInfo
  ' MethodInfo is prefixed by the actual PCode and, immediately before that, _
    a linked list that points to the previous/next entries in the linked list _
    (which is immediately prior to the PCode).
  '---- lpPreviousMethodInfoPrefix           '  -0x(dwCodeSize + 0x10)
  '---- lpNextMethodInfoPrefix               '  -0x(dwCodeSize + 0x08)
  '---- PCode(dwCodeSize) as byte            '  -0xdwCodeSize
    lpObjectInfo As LongPtr                    '   0x00
  Flag1 As Integer                           '   0x08
  Flag2 As Integer                           '   0x0A
  dwCodeSize As Long               '   0x0C   'Negative offset to P-Code
  Flag4 As Long                              '   0x10
  Flag5 As Long                              '   0x14
  Null1 As LongLong                          '   0x18
  Flag6 As Long                              '   0x20
  MI2Header As MethodInfo2Header             '   0x24
  'next integer is the length of MethodInfo2 -> a variable byte structure
  'immediately after MethodInfo2 is MethodInfo3.
  'first integer of MethodInfo3 is the length -> a variable byte structure
  'Minimum size of each appears to be 0x14.
End Type

Private Type MethodVariableInfo
  Offset As Long
  VarType As Integer
End Type



Private mVariables() As IVbaVariable

Private mMethodData As MethodData
Private mMethodData2 As MethodData2
Private mMethodInfo As MethodInfo
Private mMethodName As String
Private mReturnType As IVbaVariable
Private mlpMethodName As LongPtr
Private mArguments As Long
Private mIsClassMethod As Boolean
Private mWeakPointerToParent As LongPtr
Private mDeclaration As String
Private mReturnTypeSetupComplete As Boolean
Private mArgumentTypeSetupComplete As Boolean

Implements IMethodIdentifier
Implements IVbaMethod
Implements IVbaVariableList
Implements IChildOfDisposable

Private Const CLASS_NAME As String = "VbaMethod"

Private Property Get ReturnType() As IVbaVariable
  If Not mArgumentTypeSetupComplete Then SetupArgumentList
  Set ReturnType = mReturnType
End Property

Public Property Get Arguments() As IVbaVariableList
  If Not mArgumentTypeSetupComplete Then SetupArgumentList
  Set Arguments = Me
End Property

Public Property Get Method() As IVbaMethod
  Set Method = Me
End Property

Public Property Get IsClassMethod() As Boolean
  IsClassMethod = IVbaMethod_IsClassMethod
End Property

Public Property Get callType() As MethodTypeEnum
  callType = IVbaMethod_CallType
End Property

Public Property Get ArgumentCount() As Long
  ArgumentCount = Me.Arguments.Count
End Property

Public Property Get visibility() As VisibilityEnum
  visibility = IVbaMethod_Visibility
End Property

Public Property Get Name() As String
  Name = mMethodName
End Property

Private Sub Class_Initialize()
  mArguments = -1
End Sub

Private Sub IChildOfDisposable_DisposeOfParent()
  Set Me.ChildOfDisposable.Parent = Nothing
End Sub

Private Property Set IChildOfDisposable_Parent(RHS As Object)
  mWeakPointerToParent = ObjPtr(RHS)
End Property

Private Property Get IVbaMethod_Arguments() As IVbaVariableList
  Const METHOD_NAME As String = "IVbaMethod_Arguments[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  Set IVbaMethod_Arguments = Me.Arguments
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IVbaMethod_Declaration() As String
  Dim visibility As String, callType As String, args As String, returns As String
  If mDeclaration = vbNullString Then
    'Ease of readability over speed here:
    Select Case Me.visibility
      Case VisibilityEnum.PublicCall
        visibility = "Public"
      Case Else
        visibility = "Private"
    End Select
    Select Case Me.callType
      Case MethodTypeEnum.IsFunction
        callType = "Function"
      Case MethodTypeEnum.IsProcedure
        callType = "Sub"
      Case MethodTypeEnum.IsPropertyGet
        callType = "Property Get"
      Case MethodTypeEnum.IsPropertyLet
        callType = "Property Let"
      Case MethodTypeEnum.IsPropertySet
        callType = "Property Set"
      Case Else
        callType = "Unknown"
    End Select
    args = Me.Arguments.ToString()
    If Not Me.Method.ReturnType Is Nothing Then
      With Me.Method.ReturnType
          returns = "as " & .Description & _
          Switch(.IsArray, "()", True, vbNullString)
      End With
    End If
    mDeclaration = visibility & " " & callType & " " & Me.Name & "(" & args & ") " & returns
    'mDeclaration = Switch(Me.visibility = PublicCall, "Public ", True, "Private ") & _
      Switch(Me.callType = IsFunction, "Function ", Me.callType = IsProcedure, "Sub ", Me.callType = IsPropertyGet, "Property Get ", Me.callType = IsPropertyLet, "Property Let ", Me.callType = IsPropertySet, "Property Set", True, "Unknown ") & _
      Me.Name & "(" & Me.Arguments.ToString() & ")"
  End If
  IVbaMethod_Declaration = mDeclaration
End Property

Private Property Get IVbaMethod_IsClassMethod() As Boolean
  IVbaMethod_IsClassMethod = mIsClassMethod
End Property

Private Property Get IVbaMethod_Parent() As IVbaMethods
  Const METHOD_NAME As String = "IVbaMethod_Parent[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  If mWeakPointerToParent = 0 Then Exit Property
  Set IVbaMethod_Parent = PrivateFactory.GetObjectFromWeakPointer(mWeakPointerToParent)
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IVbaMethod_ReturnType() As IVbaVariable
  Const METHOD_NAME As String = "IVbaMethod_ReturnType[" & CLASS_NAME & "]"
  On Error GoTo HandleError

  Set IVbaMethod_ReturnType = ReturnType
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IVbaMethod_Visibility() As VisibilityEnum
  Const METHOD_NAME As String = "IVbaMethod_Visibility[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  IVbaMethod_Visibility = (mMethodData.dwMethodFlag And MethodFlags.VisibilityMask)
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IVbaMethod_Name() As String
  Const METHOD_NAME As String = "IVbaMethod_Name[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  IVbaMethod_Name = mMethodName
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Function IVbaMethod_Compare(Target As IMethodIdentifier) As MethodComparisonResult
  Const METHOD_NAME As String = "IVbaMethod_Compare[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  If Target Is Nothing Then Exit Function
  With Target
    IVbaMethod_Compare = IVbaMethod_Compare Or ((.Name = IVbaMethod_Name) And NameMatch)
    IVbaMethod_Compare = IVbaMethod_Compare Or ((.callType = IVbaMethod_CallType) And CallTypeMatch)
    IVbaMethod_Compare = IVbaMethod_Compare Or ((.visibility = IVbaMethod_Visibility) And VisibilityMatch)
  End With
Exit Function

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Function

Private Property Get IVbaMethod_CallType() As MethodTypeEnum
  Const METHOD_NAME As String = "IVbaMethod_CallType[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  IVbaMethod_CallType = (mMethodData.dwMethodFlag And MethodFlags.PropertyMask)
  If IVbaMethod_CallType Then Exit Property
  Debug.Assert ((mMethodData.wFlag2 And MethodDataFlag2.ProcTypeMask) = MethodDataFlag2.IsMethod) Or ((mMethodData.wFlag2 And MethodDataFlag2.ProcTypeMask) = MethodDataFlag2.IsDeclare)
  
  Select Case Me.IsClassMethod
    Case True
      If (mMethodData.dwMethodFlag And ReturnLastArg) Then
        IVbaMethod_CallType = IsFunction
      Else
        IVbaMethod_CallType = IsProcedure
      End If
    Case False
      If (mMethodData.ArgTypes(ArgIndex.ReturnArg) And ArgumentIdEnum.ArgumentTypeMask) Then
        IVbaMethod_CallType = IsFunction
      Else
        IVbaMethod_CallType = IsProcedure
      End If
  End Select
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IVbaMethod_Address() As LongPtr
  Const METHOD_NAME As String = "IVbaMethod_Address[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  IVbaMethod_Address = mMethodData2.lpMethodInfo
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Public Sub Setup(ByVal lpMethodData As LongPtr, ByVal lpMethodData2 As LongPtr, ByVal lpMethodName As LongPtr, Optional Parent As IVbaMethods = Nothing)
  Const METHOD_NAME = CLASS_NAME & ".Setup"
On Error GoTo HandleError

  Call LoadName(lpMethodName)
  Call LoadMethodData(lpMethodData)
  Call LoadMethodData2(lpMethodData2)
  'Lazy load these instead:
    'Call SetupReturnType
    'Call SetupArgumentList
  Set Me.ChildOfDisposable.Parent = Parent
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Sub LoadName(lpAddress As LongPtr)
  Dim n(0 To MAX_VBA_OBJECT_NAME_LENGTH) As Byte, l As Long
  Const METHOD_NAME = CLASS_NAME & ".LoadName"
  On Error GoTo HandleError
  mlpMethodName = lpAddress
  Memory.FollowPointer VarPtr(n(0)), lpAddress, MAX_VBA_OBJECT_NAME_LENGTH
  mMethodName = RTrimNull(VBA.StrConv(n, vbUnicode))
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Sub LoadMethodData(lpAddress As LongPtr)
  Const METHOD_NAME = CLASS_NAME & ".LoadMethodData"
  On Error GoTo HandleError
  Memory.FollowPointer VarPtr(mMethodData), lpAddress, LenB(mMethodData), False
  With mMethodData
    mIsClassMethod = (.ArgTypes(ArgIndex.ReturnArg) And ArgumentIdEnum.HResult_AID) = ArgumentIdEnum.HResult_AID
  End With
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Sub LoadMethodData2(lpAddress As LongPtr)
  Const METHOD_NAME = CLASS_NAME & ".LoadMethodData2"
  On Error GoTo HandleError
  Memory.FollowPointer VarPtr(mMethodData2), lpAddress, LenB(mMethodData2)
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Property Get IMethodIdentifier_CallType() As MethodTypeEnum
  Const METHOD_NAME As String = "IMethodIdentifier_CallType[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  IMethodIdentifier_CallType = IVbaMethod_CallType
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IMethodIdentifier_Name() As String
  Const METHOD_NAME As String = "IMethodIdentifier_Name[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  IMethodIdentifier_Name = IVbaMethod_Name
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IMethodIdentifier_Visibility() As VisibilityEnum
  Const METHOD_NAME As String = "IMethodIdentifier_Visibility[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  IMethodIdentifier_Visibility = IVbaMethod_Visibility

Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IVbaVariableList_Count() As Long
  Const METHOD_NAME As String = "IVbaVariableList_Count[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  IVbaVariableList_Count = mArguments
  Exit Property
HandleError:
  IVbaVariableList_Count = -1
End Property



Private Property Get IVbaVariableList_Item(Index As Long) As IVbaVariable
  Const METHOD_NAME As String = "IVbaVariableList_Item[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  
  Set IVbaVariableList_Item = mVariables(Index)
Exit Function

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If

End Property

Private Property Get SetupReturnAsArgument() As Boolean
  SetupReturnAsArgument = mMethodData.dwMethodFlag And ReturnLastArg
End Property

Private Property Get HasReturn() As Boolean
  Select Case mMethodData.ArgTypes(ArgIndex.ReturnArg) And ArgumentIdEnum.ArgumentTypeMask
    Case 0
      HasReturn = False
    Case ArgumentIdEnum.HResult_AID
      HasReturn = SetupReturnAsArgument
    Case Else
      HasReturn = True
  End Select
End Property

Private Sub SetupArgumentList()
  Const METHOD_NAME As String = CLASS_NAME & ".SetupArgumentList"
  On Error GoTo HandleError
  
  Dim i As Long, Count As Long, j As Long
  Dim typeAddress As LongPtr, nameAddress As LongPtr
  Dim lpTypeArrayAddress As LongPtr, lpNameArrayPosition As LongPtr
  Dim ContainsParamArray As Boolean
  Dim offsetLoc As LongPtr
  Dim varSizeUsed As Long
  
  With mMethodData
    mArguments = ((.dwMethodFlag And ArgumentCountMask) / 4)
    If .dwMethodFlag And ReturnLastArg Then
      mArguments = mArguments - 1
    End If
    If mMethodName = "UserClassTypeTest" Or mMethodName = "CreateVbaProject" Then Stop
  
    lpTypeArrayAddress = VarPtr(.ArgTypes(ArgIndex.Argument01))   'Location of the first argument
    Count = mArguments
    If Count > 0 Then
      ReDim mVariables(0 To Count - 1)
      ContainsParamArray = .dwMethodFlag And MethodFlags.ContainsParamArray
      lpNameArrayPosition = .lpArgumentNameArray                    'AddressOf ArgumentArray(0)
      For i = 0 To Count - 1
        typeAddress = lpTypeArrayAddress + offsetLoc
        nameAddress = lpNameArrayPosition + (i * POINTER_SIZE)
        Memory.FollowPointer VarPtr(nameAddress), nameAddress, POINTER_SIZE
        Set mVariables(i) = PrivateFactory.CreateVbaVariable(typeAddress, nameAddress, Switch(i <> Count - 1, False, True, ContainsParamArray), varSizeUsed)
        offsetLoc = offsetLoc + varSizeUsed
'        Select Case mVariables(i).VarType
'          Case VbVarType.vbObject
'            If mVariables(i).Description <> "Object" Then
'              'User Class?
'                offsetLoc = offsetLoc + (4 - (typeAddress Mod 4)) + POINTER_SIZE
'                If (lpTypeArrayAddress + offsetLoc) Mod 8 <> 0 Then offsetLoc = offsetLoc + 8 - ((lpTypeArrayAddress + offsetLoc) Mod 8)
'            End If
'          Case Else
'            If mVariables(i).VarType = vbUserDefinedType Or ContainsParamArray Then
'              'UDTs shift the location of the next variable to a relative address aligned on 0x10 boundaries.
'              If (i < Count - 1) Or SetupReturnAsArgument Then
'                Do
'                  offsetLoc = offsetLoc + PrivateFactory.vbVarTypeValueSize
'                Loop Until .ArgTypes(offsetLoc / 2) <> 0
'              End If
'            Else
'              offsetLoc = offsetLoc + PrivateFactory.vbVarTypeValueSize
'            End If
'        End Select
      Next
    End If
    If HasReturn Then
      nameAddress = mlpMethodName
      If Not SetupReturnAsArgument Then
        typeAddress = VarPtr(.ArgTypes(ArgIndex.ReturnArg))
      Else
        typeAddress = lpTypeArrayAddress + offsetLoc
      End If
      Set mReturnType = PrivateFactory.CreateVbaVariable(typeAddress, nameAddress, False, varSizeUsed)
    Else
      Set mReturnType = Nothing
    End If
  End With
  mArgumentTypeSetupComplete = True
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Property Get IVbaVariableList_Parent() As IVbaMethod
  Set IVbaVariableList_Parent = Me
End Property

Public Property Set Parent(Parent As IVbaMethod)
  mWeakPointerToParent = ObjPtr(Parent)
End Property

Public Property Get ChildOfDisposable() As IChildOfDisposable
  Set ChildOfDisposable = Me
End Property

Private Function IVbaVariableList_ToString() As String
  Dim i As Long
  Dim args As IVbaVariableList
  Set args = Me.Arguments
  For i = 1 To args.Count
    IVbaVariableList_ToString = IVbaVariableList_ToString & Switch(IVbaVariableList_ToString <> vbNullString, ", ", True, vbNullString) & args.Item(i - 1).Declaration
  Next
End Function
