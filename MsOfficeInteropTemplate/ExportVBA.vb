Imports Microsoft.Office.Interop.Access
Imports Microsoft.Office.Interop.Access.Dao

Public Class ExportVBA
    Public Shared Sub ExportModules(fileName As String)
        Dim acc As Application = New Application()
        Dim accDB As Database = acc.DBEngine.OpenDatabase(fileName)

        For Each accModule As [Module] In acc.Modules
            Dim moduleName As String = accModule.Name
            'Dim moduleText As String = accModule.GetText(AcTextFormat.acDefault)

            ' Hier können Sie den exportierten Code speichern, z. B. in eine Textdatei.
            'System.IO.File.WriteAllText(moduleName & ".txt", moduleText)
        Next

        accDB.Close()
        acc.Quit()
    End Sub

End Class
