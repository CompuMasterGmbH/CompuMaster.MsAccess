Imports CompuMaster.ComInterop

Public Class VBComponentCollection
    Inherits ObjectReadOnlyCollectionBase(Of Object, VBComponentCollection, AccessApplication, Object, VBComponent)

    Friend Sub New(parent As AccessApplication, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public Overrides ReadOnly Property Count As Integer
        Get
            Return InvokePropertyGet(Of Integer)("Count")
        End Get
    End Property

    Public Overrides ReadOnly Property Item(index As Integer) As VBComponent
        Get
            Return New VBComponent(Me, InvokePropertyGet(Of Object)("Item", index + 1))
        End Get
    End Property

End Class
