Imports CompuMaster.ComInterop

Public Class AccessObject
    Inherits ComChildObject(Of AccessAllObjectsCollection, Object)

    Friend Sub New(parent As AccessAllObjectsCollection, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public ReadOnly Property Name As String
        Get
            Return InvokePropertyGet(Of String)("Name")
        End Get
    End Property

End Class
