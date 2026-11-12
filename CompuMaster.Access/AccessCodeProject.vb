Imports CompuMaster.ComInterop

Public Class AccessCodeProject
    Inherits ComChildObject(Of AccessApplication, Object)

    Friend Sub New(parent As AccessApplication, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Private _AllModules As AccessAllObjectsCollection
    Public ReadOnly Property AllModules() As AccessAllObjectsCollection
        Get
            If _AllModules Is Nothing Then
                _AllModules = New AccessAllObjectsCollection(Me)
            End If
            Return _AllModules
        End Get
    End Property

End Class
