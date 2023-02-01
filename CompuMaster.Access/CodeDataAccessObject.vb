Imports CompuMaster.ComInterop

Public Class CodeDataAccessObject
    Inherits ComChildObject(Of Object, CodeDataAllObjectsCollection)

    Friend Sub New(parent As CodeDataAllObjectsCollection, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public ReadOnly Property Name As String
        Get
            Return InvokePropertyGet(Of String)("Name")
        End Get
    End Property

End Class
