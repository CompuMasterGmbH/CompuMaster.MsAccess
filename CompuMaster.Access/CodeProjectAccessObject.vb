Imports CompuMaster.ComInterop

Public Class CodeProjectAccessObject
    Inherits ComChildObject(Of Object, CodeprojectAllObjectsCollection)

    Friend Sub New(parent As CodeprojectAllObjectsCollection, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public ReadOnly Property Name As String
        Get
            Return InvokePropertyGet(Of String)("Name")
        End Get
    End Property

End Class
