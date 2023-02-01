Imports CompuMaster.ComInterop

Public Class CodeProjectCollection
    Inherits ObjectReadOnlyCollectionBase(Of Object, CodeProjectCollection, AccessApplication, Object, CodeProject)

    Friend Sub New(parent As AccessApplication, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public Overrides ReadOnly Property Count As Integer
        Get
            Return InvokePropertyGet(Of Integer)("Count")
        End Get
    End Property

    Public Overrides ReadOnly Property Item(index As Integer) As CodeProject
        Get
            Return New CodeProject(Me, InvokePropertyGet(Of Object)("Item", index + 1))
        End Get
    End Property

End Class
