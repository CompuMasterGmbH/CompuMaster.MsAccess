Imports CompuMaster.ComInterop

Public Class VBProjectCollection
    Inherits ObjectReadOnlyCollectionBase(Of Object, VBProjectCollection, AccessApplication, Object, VBProject)

    Friend Sub New(parent As AccessApplication, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public Overrides ReadOnly Property Count As Integer
        Get
            Return InvokePropertyGet(Of Integer)("Count")
        End Get
    End Property

    Public Overrides ReadOnly Property Item(index As Integer) As VBProject
        Get
            Return New VBProject(Me, InvokePropertyGet(Of Object)("Item", index + 1))
        End Get
    End Property

End Class
