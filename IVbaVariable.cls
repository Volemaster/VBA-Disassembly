VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVbaVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ArgumentIdEnum
  'Just go with it, I guess.
  ArgumentTypeMask = &H1F
  'Types:
    Integer_AID = &H6
    Long_AID = &H8
    Single_AID = &HA
    Double_AID = &HB
    Currency_AID = &HD
    Date_AID = &HC
    String_AID = &H10
    Object_AID = &H1C
    Boolean_AID = &H3
    Variant_AID = &HF
    Byte_AID = &H5
    LongLong_AID = &H11
    HResult_AID = &H1F  'Returned from all class calls
    Any_AID = &H2
    UserDefined_AID = &H15
    VbaClass_AID = &H14
    ReferenceClass_AID = &H1E
  'Flags:
  IsByRef_AID = &H40 'Flag
  IsArray_AID = &H80 'Flag
  IsOptional_AID = &H100 ' Flag
End Enum

Public Enum PassByEnum
  PassByReference
  PassByValue
End Enum

Public Property Get Name() As String

End Property

Public Property Get VarType() As VbVarType
  
End Property

Public Property Get Description() As String

End Property

Public Property Get IsArray() As Boolean

End Property

Public Property Get PassBy() As PassByEnum

End Property

Public Property Get Declaration() As String

End Property

Public Property Get IsOptional() As Boolean

End Property
