Imports CompuMaster.ComInterop
Imports Microsoft.Office.Interop.Access.Dao
Imports NUnit.Framework

Namespace CompuMaster.Test.MsAccess

    <NonParallelizable>
    Public Class MsAccessInteropTests

        Public Sub CompileTests_MsAccessInteropApp()

            Dim App As Microsoft.Office.Interop.Access.Application = Nothing
            Dim DbEngine As Microsoft.Office.Interop.Access.Dao._DBEngine = App.DBEngine
            Dim Db = DbEngine.OpenDatabase("")

            App.CurrentDb()
            'ComObject.VBE
            'ComObject.DBEngine
            'ComObject.CodeData.AllFunctions.Item
            'ComObject.CodeProject.AllForms
            Dim dummy As Integer = App.Modules.Count

        End Sub

        <SetUp>
        Public Sub Setup()
        End Sub

        <OneTimeSetUp>
        Public Sub OneTimeSetUp()
            OpenAccessAppAndDatabase(TestEnvironment.TestFiles.TestFileNorthwindDatabase.FullName)
        End Sub

        Private _AppWithNordwind As ComRootObject(Of Microsoft.Office.Interop.Access.Application)
        Private ReadOnly Property AppWithNordwind As Microsoft.Office.Interop.Access.Application
            Get
                Return _AppWithNordwind.ComObjectStronglyTyped
            End Get
        End Property

        Private _DbEngine As ComChildObject(Of Microsoft.Office.Interop.Access.Dao._DBEngine, ComRootObject(Of Microsoft.Office.Interop.Access.Application))
        Private ReadOnly Property DbEngine As Microsoft.Office.Interop.Access.Dao._DBEngine
            Get
                Return _DbEngine.ComObjectStronglyTyped
            End Get
        End Property

        Private _NordwindDb As ComChildObject(Of Microsoft.Office.Interop.Access.Dao.Database, ComRootObject(Of Microsoft.Office.Interop.Access.Application))
        Private ReadOnly Property NordwindDb As Microsoft.Office.Interop.Access.Dao.Database
            Get
                Return _NordwindDb.ComObjectStronglyTyped
            End Get
        End Property

        Protected Sub OpenAccessAppAndDatabase(databasePath As String)
            _AppWithNordwind = New ComRootObject(Of Microsoft.Office.Interop.Access.Application)(
                New Microsoft.Office.Interop.Access.Application(),
                Nothing,
                Sub(instance)
                    instance.ComObjectStronglyTyped.Quit()
                End Sub,
                Nothing)

            _DbEngine = New ComChildObject(Of Microsoft.Office.Interop.Access.Dao._DBEngine, ComRootObject(Of Microsoft.Office.Interop.Access.Application))(_AppWithNordwind, AppWithNordwind.DBEngine)
            _NordwindDb = New ComChildObject(Of Microsoft.Office.Interop.Access.Dao.Database, ComRootObject(Of Microsoft.Office.Interop.Access.Application))(_AppWithNordwind, DbEngine.OpenDatabase(databasePath))
            Assert.NotNull(_NordwindDb.ComObjectStronglyTyped.Name)
        End Sub

        <Test>
        Public Sub AccessBasicObjectAccess()
            Assert.IsNotNull(AppWithNordwind)
        End Sub

        <Test>
        Public Sub CurrentDb()
            Assert.IsNotNull(_NordwindDb.ComObjectStronglyTyped)
            Assert.IsNull(AppWithNordwind.CurrentDb)
        End Sub

        <Test>
        Public Sub CurrentProject()
            Assert.IsNotNull(_NordwindDb.ComObjectStronglyTyped)
            Assert.IsNotNull(AppWithNordwind.CurrentProject)
        End Sub

        <Test>
        Public Sub Modules()
            Assert.IsNotNull(AppWithNordwind.Modules)
            For Each Item As Microsoft.Office.Interop.Access.Module In AppWithNordwind.Modules
                Assert.IsNotNull(Item)
                Console.WriteLine("Found Module: " & Item.Name)
            Next
            Assert.IsNotNull(AppWithNordwind.Modules.Count)
        End Sub

        <Test>
        Public Sub CodeData()
            Assert.IsNotNull(AppWithNordwind.CodeData)
            Assert.NotZero(AppWithNordwind.CodeData.AllTables.Count)
            Assert.NotZero(AppWithNordwind.CodeData.AllQueries.Count)
            Assert.NotZero(AppWithNordwind.CodeData.AllViews.Count)
            Assert.NotZero(AppWithNordwind.CodeData.AllStoredProcedures.Count)
            Assert.NotZero(AppWithNordwind.CodeData.AllFunctions.Count)
            Assert.NotZero(AppWithNordwind.CodeData.AllDatabaseDiagrams.Count)
        End Sub

        <Test>
        Public Sub CodeProject()
            Assert.IsNotNull(AppWithNordwind.CodeProject)
            Dim M = AppWithNordwind.CodeProject.AllModules
            Assert.NotZero(M.Count)
            Dim M0 = M(0)
            Assert.NotNull(M0.Name)

            Assert.NotZero(AppWithNordwind.CodeProject.AllModules.Count)
            Assert.NotZero(AppWithNordwind.CodeProject.AllMacros.Count)
            Assert.NotZero(AppWithNordwind.CodeProject.AllForms.Count)
            Assert.NotZero(AppWithNordwind.CodeProject.AllReports.Count)
            Assert.NotZero(AppWithNordwind.CodeProject.AllDataAccessPages.Count)
        End Sub

        <Test>
        Public Sub VBE()
            Assert.IsNotNull(AppWithNordwind.VBE)
            Assert.Zero(AppWithNordwind.VBE.VBProjects.Count)
            Assert.IsNotNull(AppWithNordwind.VBE.VBProjects.Item(0))
            Assert.IsNotNull(AppWithNordwind.VBE.VBProjects.Item(0).Name)
            Assert.NotZero(AppWithNordwind.VBE.VBProjects.Item(0).VBComponents.Count)
            Assert.IsNotNull(AppWithNordwind.VBE.VBProjects.Item(0).VBComponents.Item(0))
            Assert.IsNotNull(AppWithNordwind.VBE.VBProjects.Item(0).VBComponents.Item(0).Type)
        End Sub

        'Public Sub CodeData()
        '    Dim App As Global.CompuMaster.MsAccess.AccessApplication = OpenAccessAppAndDatabase(TestEnvironment.TestFiles.TestFileNorthwindDatabase.FullName)
        '    Assert.IsNotNull(App.Modules)
        '    Assert.IsNotNull(App.Modules.Count)
        '    ComObject.VBE
        '    ComObject.DBEngine
        '    ComObject.CodeData.AllFunctions.Item
        '    ComObject.CodeProject.AllForms
        '    ComObject.Modules
        'End Sub

    End Class

End Namespace
