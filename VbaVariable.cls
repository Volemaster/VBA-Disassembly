VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IVbaVariable

Private mDescription As String
Private mName As String
Private mVarType As VbVarType
Private mIsArray As Boolean
Private mIsParamArray As Boolean
Private mIsOptional As Boolean
Private mPassBy As PassByEnum
Private mSizeUsed As Long

Private mTypeValue As Integer

Private Const CLASS_NAME As String = "VbaVariable"

Public Property Get SizeUsed() As Long
  SizeUsed = mSizeUsed
End Property

Private Property Get IVbaVariable_Declaration() As String
  '[ByRef|ByVal] [Optional|ParamArray] <VarName> [as <VarType>[()]]
  Const METHOD_NAME As String = "IVbaVariable_Declaration[" & CLASS_NAME & "]"
  On Error GoTo HandleError
  IVbaVariable_Declaration = _
    Switch(mPassBy = PassByReference, "ByRef ", mPassBy = PassByValue, "ByVal ", True, vbNullString) & _
    Switch(mIsParamArray, "ParamArray ", mIsOptional, "Optional ", True, vbNullString) & _
    mName & " as " & mDescription & _
    Switch(mIsArray, "()", True, vbNullString)
Exit Property

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
End Property

Private Property Get IVbaVariable_Description() As String
  IVbaVariable_Description = mDescription
End Property

Private Property Get IVbaVariable_IsArray() As Boolean
  IVbaVariable_IsArray = mIsArray
End Property

Private Property Get IVbaVariable_IsOptional() As Boolean
  IVbaVariable_IsOptional = mIsOptional
End Property

Private Property Get IVbaVariable_Name() As String
  IVbaVariable_Name = mName
End Property

Private Property Get IVbaVariable_PassBy() As PassByEnum
  IVbaVariable_PassBy = mPassBy
End Property

Private Property Get IVbaVariable_VarType() As VbVarType
  IVbaVariable_VarType = mVarType
End Property

Public Sub Setup(lpTypeAddress As LongPtr, lpName As LongPtr, IsParamArray As Boolean)
  Const METHOD_NAME As String = CLASS_NAME & ".get_Objects"
  On Error GoTo HandleError
  Dim bName(0 To MAX_VBA_OBJECT_NAME_LENGTH) As Byte
    mIsParamArray = IsParamArray
    Memory.FollowPointer VarPtr(mTypeValue), lpTypeAddress, PrivateFactory.vbVarTypeValueSize
    mSizeUsed = PrivateFactory.vbVarTypeValueSize
    Select Case mTypeValue And ArgumentIdEnum.ArgumentTypeMask
      Case 0
        'No return
        Exit Sub
      Case Integer_AID
        mVarType = vbInteger
        mDescription = "Integer"
      Case Long_AID
        mVarType = vbLong
        mDescription = "Long"
      Case Single_AID
        mVarType = vbSingle
        mDescription = "Single"
      Case Double_AID
        mVarType = vbDouble
        mDescription = "Double"
      Case Currency_AID
        mVarType = vbCurrency
        mDescription = "Currency"
      Case Date_AID
        mVarType = vbDate
        mDescription = "Date"
      Case String_AID
        mVarType = vbString
        mDescription = "String"
      Case Object_AID
        mVarType = vbObject
        mDescription = "Object"
      Case Boolean_AID
        mVarType = vbBoolean
        mDescription = "Boolean"
      Case Variant_AID
        mVarType = vbVariant
        mDescription = "Variant"
      Case Byte_AID
        mVarType = vbByte
        mDescription = "Byte"
      Case LongLong_AID
        mVarType = vbLongLong
        mDescription = "LongLong"
      Case ArgumentIdEnum.UserDefined_AID
        mVarType = vbUserDefinedType
        mDescription = "UserDefinedType"
      Case ArgumentIdEnum.Any_AID
        mVarType = vbVariant
        mDescription = "Any"
      Case ArgumentIdEnum.HResult_AID
        mVarType = VbVarType.vbNull
        mDescription = "HRESULT"
      Case ArgumentIdEnum.VbaClass_AID
        mVarType = vbObject
        mDescription = GetTargetClassName(lpTypeAddress)
      Case ArgumentIdEnum.ReferenceClass_AID
        mVarType = vbObject
        mDescription = GetInterfaceName(lpTypeAddress)
      Case Else
        mVarType = VbVarType.vbNull
        mDescription = "Unknown"
    End Select
    Select Case mTypeValue And ArgumentIdEnum.ArgumentTypeMask
      Case ArgumentIdEnum.VbaClass_AID, _
           ArgumentIdEnum.ReferenceClass_AID, _
           ArgumentIdEnum.UserDefined_AID
        'Pointer to class is aligned on 4 byte boundaries
        If (mSizeUsed + lpTypeAddress) Mod 4 <> 0 Then
          mSizeUsed = mSizeUsed + (4 - CLng((mSizeUsed + lpTypeAddress) Mod 4))
        End If
        mSizeUsed = mSizeUsed + POINTER_SIZE
        'Next argument is aligned on 8 byte boundaries
        If (lpTypeAddress + mSizeUsed) Mod 8 <> 0 Then
          mSizeUsed = mSizeUsed + 8 - CLng((lpTypeAddress + mSizeUsed) Mod 8)
        End If
      Case Else
      
    End Select
    
    mIsArray = mTypeValue And ArgumentIdEnum.IsArray_AID
    If mTypeValue And ArgumentIdEnum.IsByRef_AID Then
      mPassBy = PassByReference
    Else
      mPassBy = PassByValue
    End If
    mIsOptional = mTypeValue And ArgumentIdEnum.IsOptional_AID
    
    Memory.FollowPointer VarPtr(bName(0)), lpName, MAX_VBA_OBJECT_NAME_LENGTH
    mName = RTrimNull(StrConv(bName, vbUnicode))
Exit Sub

HandleError:
  If Err.Source <> METHOD_NAME Then
    Err.Raise Err.Number, METHOD_NAME & "." & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End If
  Resume
End Sub

Private Function GetTargetClassName(lpTypeAddress As LongPtr) As String
Dim lpTypeObjectPointer As LongPtr
Dim lpPublicObjectDescriptor As LongPtr
Dim lpObjectName As LongPtr
Dim lpObjectInfo As LongPtr
Dim b(0 To MAX_VBA_OBJECT_NAME_LENGTH) As Byte
  Select Case lpTypeAddress Mod 4
    Case 0
      lpTypeObjectPointer = lpTypeAddress + 4
    Case Else
      lpTypeObjectPointer = lpTypeAddress + (lpTypeAddress Mod 4)
  End Select
  Memory.FollowPointer VarPtr(lpObjectInfo), lpTypeObjectPointer, POINTER_SIZE
  Memory.FollowPointer VarPtr(lpPublicObjectDescriptor), lpObjectInfo + &H30, POINTER_SIZE
  Memory.FollowPointer VarPtr(lpObjectName), lpPublicObjectDescriptor + &H30, POINTER_SIZE
  Memory.FollowPointer VarPtr(b(0)), lpObjectName, MAX_VBA_OBJECT_NAME_LENGTH, False
  GetTargetClassName = RTrimNull(StrConv(b, vbUnicode))
  
End Function

Private Function GetInterfaceName(lpTypeAddress As LongPtr) As String
' It should be easy enough to extract the type name by looking up the _
  IID from ExternalObjectRef in the registry or by calling the library's _
  ITypeLib methods. The latter is probably preferable, but annoying to _
  write here.
' So, for now, we'll just return IUnknown'
  GetInterfaceName = "IUnknown"
End Function

