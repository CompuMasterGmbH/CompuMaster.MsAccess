Imports CompuMaster.ComInterop

Public Class AccessModule
    Inherits ComChildObject(Of Object, AccessModuleCollection)

    Friend Sub New(parent As AccessModuleCollection, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public ReadOnly Property Name As String
        Get
            Return InvokePropertyGet(Of String)("Name")
        End Get
    End Property

    Public ReadOnly Property CodeName As String
        Get
            Return InvokePropertyGet(Of String)("CodeName")
        End Get
    End Property

    Public Function GetText(textFormat As Enumerations.acTextFormat) As String
        Return InvokePropertyGet(Of String)("GetText", textFormat)
    End Function

    Protected Overrides Sub OnDisposeChildren()
    End Sub

    Protected Overrides Sub OnClosing()
    End Sub

    Protected Overrides Sub OnClosed()
    End Sub

End Class
