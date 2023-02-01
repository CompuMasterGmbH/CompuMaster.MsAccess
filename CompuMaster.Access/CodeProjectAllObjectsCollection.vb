Imports CompuMaster.ComInterop

Public Class CodeProjectAllObjectsCollection
    Inherits ObjectReadOnlyCollectionBase(Of Object, CodeProjectAllObjectsCollection, CodeData, Object, CodeProjectAccessObject)

    Friend Sub New(parent As CodeData, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public Overrides ReadOnly Property Count As Integer
        Get
            Return InvokePropertyGet(Of Integer)("Count")
        End Get
    End Property

    Public Overrides ReadOnly Property Item(index As Integer) As CodeProjectAccessObject
        Get
            Return New CodeProjectAccessObject(Me, InvokePropertyGet(Of Object)("Item", index + 1))
        End Get
    End Property

End Class
