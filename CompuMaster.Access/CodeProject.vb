Imports CompuMaster.ComInterop

Public Class CodeProject
    Inherits ComChildObject(Of Object, CodeProjectCollection)

    Friend Sub New(parent As CodeProjectCollection, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Private _AllModules As AccessAllObjectsCollection
    Public ReadOnly Property AllModules() As AccessAllObjectsCollection
        Get
            If _AllModules Is Nothing Then
                _AllModules = New AccessAllObjectsCollection(Me, Me.InvokePropertyGet(Of Object)("AllModules"))
            End If
            Return _AllModules
        End Get
    End Property

End Class
