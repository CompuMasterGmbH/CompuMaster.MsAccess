Imports CompuMaster.ComInterop

Public Class VBComponent
    Inherits ComChildObject(Of VBComponentCollection, Object)

    Friend Sub New(parent As VBComponentCollection, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

End Class
