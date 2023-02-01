Imports CompuMaster.ComInterop

Public Class AccessDBEngine
    Inherits ComObjectBase

    Friend Sub New(parentItemResponsibleForDisposal As ComObjectBase, app As AccessApplication)
        MyBase.New(parentItemResponsibleForDisposal, app.InvokePropertyGet(Of Object)("DBEngine"))
        Me.Parent = app
    End Sub

    Friend ReadOnly Parent As AccessApplication

    Public Function OpenDatabase(path As String) As AccessDatabase
        If Parent.CurrentDb IsNot Nothing Then Throw New InvalidOperationException("Close the current database before opening next database")
        Return New AccessDatabase(Me.Parent, Me, path)
    End Function

    Friend ReadOnly Property IsClosedComObject As Boolean
        Get
            Return MyBase.IsDisposedComObject
        End Get
    End Property

    Protected Overrides Sub OnDisposeChildren()
    End Sub

    Protected Overrides Sub OnClosing()
    End Sub

    Protected Overrides Sub OnClosed()
    End Sub

End Class
