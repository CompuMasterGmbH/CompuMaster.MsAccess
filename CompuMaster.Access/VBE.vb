Imports CompuMaster.ComInterop

Public Class VBE
    Inherits ComChildObject(Of Object, VBECollection)

    Friend Sub New(parent As VBECollection, comObject As Object)
        MyBase.New(parent, comObject)
    End Sub

    Public Property ActiveProject As ComInterop.ComObjectBase

End Class
