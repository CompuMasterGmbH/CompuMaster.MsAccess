Imports CompuMaster.ComInterop

Public Class ComAppObject
    Inherits ComRootObject(Of Microsoft.Office.Interop.Access.Application)

    Public Sub New(parentItemResponsibleForDisposal As Global.CompuMaster.ComInterop.ComObjectBase, obj As Microsoft.Office.Interop.Access.Application)
        MyBase.New(parentItemResponsibleForDisposal, obj)
    End Sub

    Protected Overrides Sub OnDisposeChildren()
    End Sub

    Protected Overrides Sub OnClosing()
        Me.ComObjectStronglyTyped.Quit()
    End Sub

    Protected Overrides Sub OnClosed()
    End Sub

End Class
