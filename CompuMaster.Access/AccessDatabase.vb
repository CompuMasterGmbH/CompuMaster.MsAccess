Imports CompuMaster.ComInterop
Imports System.IO

Public Class AccessDatabase
    Inherits ComChildObject(Of AccessApplication, Object)

    Friend Sub New(parentItemResponsibleForDisposal As AccessApplication, c As AccessDBEngine, path As String)
        MyBase.New(parentItemResponsibleForDisposal, c.InvokeFunction(Of Object)("OpenDatabase", New Object() {path}))
        FilePath = path
        Me.Parent.SetCurrentDb(Me)
    End Sub

    Public ReadOnly FilePath As String

    Public Shadows ReadOnly Property IsClosed As Boolean
        Get
            Return MyBase.IsDisposedComObject
        End Get
    End Property

    Public Overrides Sub Close()
        MyBase.CloseAndDisposeChildrenAndComObject()
    End Sub

    Protected Overrides Sub OnDisposeChildren()
    End Sub

    Protected Overrides Sub OnClosing()
        InvokeMethod("Close")
    End Sub

    Protected Overrides Sub OnClosed()
        Parent.SetCurrentDb(Nothing)
    End Sub

End Class
