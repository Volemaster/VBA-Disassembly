VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MethodIdentifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IMethodIdentifier

Private mVisibility As VisibilityEnum
Private mName As String
Private mCallType As MethodTypeEnum

Public Property Let visibility(Value As VisibilityEnum)
  mVisibility = Value
End Property

Public Property Let Name(Value As String)
  mName = Value
End Property

Public Property Let callType(Value As MethodTypeEnum)
  mCallType = Value
End Property

Public Property Get visibility() As VisibilityEnum
  visibility = mVisibility
End Property

Public Property Get Name() As String
  Name = mName
End Property

Public Property Get callType() As MethodTypeEnum
  callType = mCallType
End Property

Private Property Get IMethodIdentifier_CallType() As MethodTypeEnum
  IMethodIdentifier_CallType = Me.callType
End Property

Private Property Get IMethodIdentifier_Name() As String
  IMethodIdentifier_Name = Me.Name
End Property

Private Property Get IMethodIdentifier_Visibility() As VisibilityEnum
  IMethodIdentifier_Visibility = Me.visibility
End Property
