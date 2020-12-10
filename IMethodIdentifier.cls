VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMethodIdentifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get visibility() As VisibilityEnum
End Property

Public Property Get Name() As String
End Property

Public Property Get callType() As MethodTypeEnum
End Property
