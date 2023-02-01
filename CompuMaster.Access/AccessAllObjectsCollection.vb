Imports CompuMaster.ComInterop

Public Class AccessAllObjectsCollection
    Inherits ObjectReadOnlyCollectionBase(Of Object, AccessAllObjectsCollection, ComObjectBase, Object, AccessObject)

    Friend Sub New(parent As ComObjectBase, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public Overrides ReadOnly Property Count As Integer
        Get
            Return InvokePropertyGet(Of Integer)("Count")
        End Get
    End Property

    Public Overrides ReadOnly Property Item(index As Integer) As AccessObject
        Get
            Return New AccessObject(Me, InvokePropertyGet(Of Object)("Item", index + 1))
        End Get
    End Property

End Class
