Imports CompuMaster.ComInterop

Public Class VBE
    Inherits ComChildObject(Of VBECollection, Object)

    Friend Sub New(parent As VBECollection, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public Property ActiveProject As ComInterop.ComObjectBase

End Class
