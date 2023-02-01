Imports CompuMaster.ComInterop

Public Class CodeData
    Inherits ComChildObject(Of Object, CodeDataCollection)

    Friend Sub New(parent As CodeDataCollection, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Private _AllQueries As CodeDataAllObjectsCollection
    Public ReadOnly Property AllQueries() As CodeDataAllObjectsCollection
        Get
            If _AllQueries Is Nothing Then
                _AllQueries = New CodeDataAllObjectsCollection(Me, Me.InvokePropertyGet(Of Object)("AllQueries"))
            End If
            Return _AllQueries
        End Get
    End Property

End Class
