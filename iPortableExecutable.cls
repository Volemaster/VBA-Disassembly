VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iPortableExecutable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Sub Load(ModuleName As String)
End Sub

Public Function FindInSection(SectionName As String, BufferToFind() As Byte, Optional StartingLocation As LongPtr = 0) As LongPtr
End Function


Public Property Get ImageSize() As LongPtr
End Property
