Imports CompuMaster.ComInterop

Public Class VBProject
    Inherits ComChildObject(Of VBProjectCollection, Object)

    Friend Sub New(parent As VBProjectCollection, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub


End Class
