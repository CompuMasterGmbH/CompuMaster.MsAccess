Imports CompuMaster.ComInterop
Public Class VBECollection
    Inherits ObjectReadOnlyCollectionBase(Of Object, VBECollection, AccessApplication, Object, VBE)

    Friend Sub New(parent As AccessApplication, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public Property ActiveProject As ComInterop.ComObjectBase

    Public Overrides ReadOnly Property Count As Integer
        Get
            Return InvokePropertyGet(Of Integer)("Count")
        End Get
    End Property

    Public Overrides ReadOnly Property Item(index As Integer) As VBE
        Get
            Return New VBE(Me, InvokePropertyGet(Of Object)("Item", index + 1))
        End Get
    End Property

End Class
