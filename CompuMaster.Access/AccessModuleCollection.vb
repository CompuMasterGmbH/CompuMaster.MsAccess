Imports CompuMaster.ComInterop

Public Class AccessModuleCollection
    Inherits ObjectReadOnlyCollectionBase(Of Object, AccessModuleCollection, AccessApplication, Object, AccessModule)

    Friend Sub New(parent As AccessApplication, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public Overrides ReadOnly Property Count As Integer
        Get
            Return InvokePropertyGet(Of Integer)("Count")
        End Get
    End Property

    Public Overrides ReadOnly Property Item(index As Integer) As AccessModule
        Get
            Return New AccessModule(Me, InvokePropertyGet(Of Object)("Item", index + 1))
        End Get
    End Property

    Protected Overrides Sub OnDisposeChildren()
    End Sub

    Protected Overrides Sub OnClosing()
    End Sub

    Protected Overrides Sub OnClosed()
    End Sub
End Class
