Imports CompuMaster.ComInterop

Public Class AccessCodeProject
    Inherits ComChildObject(Of Object, AccessApplication)

    Friend Sub New(parent As AccessApplication, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Private _AllModules As AccessAllObjectsCollection
    Public ReadOnly Property AllModules() As AccessAllObjectsCollection
        Get
            If _AllModules Is Nothing Then
                _AllModules = New AccessAllObjectsCollection(Me, Me.InvokeFunction(Of Object)("AllModules"))
            End If
            Return _AllModules
        End Get
    End Property

End Class
