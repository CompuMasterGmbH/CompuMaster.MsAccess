Imports CompuMaster.ComInterop

Public Class CodeDataAllObjectsCollection
    Inherits ObjectReadOnlyCollectionBase(Of Object, CodeDataAllObjectsCollection, CodeData, Object, CodeDataAccessObject)

    Friend Sub New(parent As CodeData, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public Overrides ReadOnly Property Count As Integer
        Get
            Return InvokePropertyGet(Of Integer)("Count")
        End Get
    End Property

    Public Overrides ReadOnly Property Item(index As Integer) As CodeDataAccessObject
        Get
            Return New CodeDataAccessObject(Me, InvokePropertyGet(Of Object)("Item", index + 1))
        End Get
    End Property

End Class
