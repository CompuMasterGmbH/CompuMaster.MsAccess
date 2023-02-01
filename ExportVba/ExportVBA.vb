Imports CompuMaster.MsAccess

Public Class ExportVBA

    Public Delegate Sub ExportPostAction(fullPath As String, filename As String)

    Public Shared Sub ExportModules(databaseName As String, targetPath As String, codeFileExtension As String, actionOnWrittenExportFile As ExportPostAction)
        Dim acc As AccessApplication = New AccessApplication()
        Dim accDB = acc.DBEngine.OpenDatabase(databaseName)

        For MyCounter As Integer = 0 To acc.Modules.Count - 1
            Dim accModule = acc.Modules.Item(MyCounter)
            Dim moduleName As String = accModule.Name
            Dim moduleText As String = accModule.GetText(Enumerations.acTextFormat.acDefault)
            System.IO.File.WriteAllText(System.IO.Path.Combine(targetPath, moduleName & codeFileExtension), moduleText)
            actionOnWrittenExportFile(targetPath, moduleName & codeFileExtension)
        Next

        accDB.Close()
        acc.Quit()
    End Sub

End Class
