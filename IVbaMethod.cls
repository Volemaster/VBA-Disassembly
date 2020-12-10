VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVbaMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum AddressTypeEnum
  [_atMin] = 0
  atany = [_atMin]
  atPropGet = atany + 1
  atPropSet = atPropGet + 1
  atPropLet = atPropSet
  atMethod = atPropSet + 1
  atFunction = atMethod
  atProcedure = atMethod
  [_atMax] = atMethod
End Enum

Public Enum MethodFlags
  PropertyMask = &H3&
    IsPropertyGet_ = &H1&
    IsPropertyLet_ = &H2&
    IsPropertySet_ = &H3&
    IsMethod = &H0&          ' is sub or function. If ((ReturnValue AND NOT &H10) OR ReturnLastArg) then Function
  VisibilityMask = &H10000
    IsPublic = &H10000
    IsPrivate = &H0&
  ReturnLastArg = &H100&         ' Only seen in classes (because they always return HResults in the calls)
  ContainsParamArray = &HFC00&
  ArgumentCountMask = &HFC&          ' Count = (dwMethodFlags AND ArgumentCountMask)>>2
  VTableOffsetMask = &HFFFC0000     ' starts at Pointer_Size * 8 -> first seven pointers are IUnknown / IDispatch.
    InStandardModule = &HFFFC0000   ' Otherwise in class module.
End Enum

Public Enum MethodComparisonResult
  NoMatch = &H0
  NameMatch = &H1
  ReturnTypeMatch = &H2
  CallTypeMatch = &H4
  VisibilityMatch = &H8
  Matches = NameMatch Or CallTypeMatch Or VisibilityMatch ' Or ReturnTypeMatch
End Enum

Public Enum VisibilityEnum
  PublicCall = MethodFlags.IsPublic
  PrivateCall = MethodFlags.IsPrivate
End Enum

Public Enum MethodTypeEnum
  IsPropertyGet = MethodFlags.IsPropertyGet_
  IsPropertyLet = MethodFlags.IsPropertyLet_
  IsPropertySet = MethodFlags.IsPropertySet_
  [_mtUnset] = &HFFFFFFFF
  IsProcedure = [_mtUnset] And (Not MethodFlags.PropertyMask)
  IsFunction = [_mtUnset]
End Enum

Public Property Get visibility() As VisibilityEnum
End Property

Public Property Get Name() As String
End Property

Public Property Get callType() As MethodTypeEnum
End Property

Public Property Get Address() As LongPtr
End Property

Public Property Get Arguments() As IVbaVariableList
End Property

Public Function Compare(Target As IMethodIdentifier) As MethodComparisonResult
End Function

Public Property Get ReturnType() As IVbaVariable
End Property

Public Property Get Parent() As IVbaMethods
End Property

Public Property Get IsClassMethod() As Boolean
End Property

Public Property Get Declaration() As String
End Property
