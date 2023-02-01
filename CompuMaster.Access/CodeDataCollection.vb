Imports CompuMaster.ComInterop

Public Class CodeDataCollection
    Inherits ObjectReadOnlyCollectionBase(Of Object, CodeDataCollection, AccessApplication, Object, CodeData)

    Friend Sub New(parent As AccessApplication, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public Overrides ReadOnly Property Count As Integer
        Get
            Return InvokePropertyGet(Of Integer)("Count")
        End Get
    End Property

    Public Overrides ReadOnly Property Item(index As Integer) As CodeData
        Get
            Return New CodeData(Me, InvokePropertyGet(Of Object)("Item", index + 1))
        End Get
    End Property

End Class
