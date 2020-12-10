VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PrivateFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mTypeValueSize As Long = 2
Private mPublicObjectDescriptorSize As Long

Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32.dll" (ByRef destination As Any, ByVal lpSource As LongPtr, ByVal size As Long)

' Concrete classes are deliberately private so that the implementation isn't exposed outside of this project. _
  Interfaces are public so that they can be referenced. _
  This class is a private singleton so that its members are public within the project but cannot be referenced _
  outside of the project.
' Use Supplemental.TryGetVbaProjects() to grab an IVbaProjects instance, not GetVbaProjects.

Public Property Get vbVarTypeValueSize() As Long
  vbVarTypeValueSize = mTypeValueSize
End Property

Public Function GetVbaProjects(lpAddress As LongPtr) As IVbaProjects
  Dim concrete As VbaProjects
  Set concrete = New VbaProjects
  concrete.Setup lpAddress
  Set GetVbaProjects = concrete
End Function

Public Function CreateVbaProject(lpProjectInfoPointer As LongPtr, Optional Parent As IVbaProjects = Nothing) As IVbaProject
  Dim concrete As VbaProject
  Set concrete = New VbaProject
  concrete.Setup lpProjectInfoPointer, Parent
  Set CreateVbaProject = concrete
End Function

Public Function CreateVbaObject(lpPublicObjectDescriptorBase As LongPtr, Index As Long, Optional Parent As IVbaObjects = Nothing) As IVbaObject
  Dim concrete As VbaObject
  Set concrete = New VbaObject
  If mPublicObjectDescriptorSize = 0 Then
    mPublicObjectDescriptorSize = concrete.PublicObjectDescriptorSize
  End If
  concrete.Setup lpPublicObjectDescriptorBase + (Index * mPublicObjectDescriptorSize)
  Set CreateVbaObject = concrete
End Function

Public Function CreateVbaMethod(ByVal lpMethodData As LongPtr, ByVal lpMethodData2 As LongPtr, ByVal lpMethodName As LongPtr, Optional Parent As IVbaMethods = Nothing)
  Dim concrete As VbaMethod
  Set concrete = New VbaMethod
  Call concrete.Setup(lpMethodData, lpMethodData2, lpMethodName)
  Set CreateVbaMethod = concrete
End Function

Public Function CreateVbaVariable(lpTypeAddress As LongPtr, lpName As LongPtr, IsParamArray As Boolean, ByRef SizeUsedOut As Long) As IVbaVariable
  Dim concrete As VbaVariable
  Set concrete = New VbaVariable
  With concrete
    .Setup lpTypeAddress, lpName, IsParamArray
    SizeUsedOut = .SizeUsed
  End With
  Set CreateVbaVariable = concrete
End Function

Public Function GetObjectFromWeakPointer(lpObjectPointer As LongPtr) As Object
  Dim tmp As Object
  ' Get a reference without calling AddReference/QueryInterface (illegal under COM rules):
  RtlMoveMemory ByVal VarPtr(tmp), lpObjectPointer, POINTER_SIZE
  ' Use the memory address to call QueryInterface(IDispatch) and assign to our return object:
  Set GetObjectFromWeakPointer = tmp
  ' Clear the illegal reference so the garbage collector doesn't try to call Release() and trash the object count:
  RtlMoveMemory ByVal VarPtr(tmp), 0, POINTER_SIZE
End Function
